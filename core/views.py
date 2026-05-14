import json
import os
import re
import shutil
import sqlite3
import tempfile
from datetime import datetime
from pathlib import Path
from collections import Counter
from math import ceil
from urllib.parse import quote

from django.conf import settings
from django.shortcuts import render, redirect, get_object_or_404
from django.urls import reverse
from django.contrib.auth import get_user_model, login, logout
from django.contrib.auth.models import User
from django.contrib.auth.decorators import login_required, user_passes_test
from django.contrib import messages
from django.db import connections
from django.db.models import Q
from django.http import HttpResponse, HttpResponseBadRequest, JsonResponse
from django.views.decorators.http import require_POST

from .models import (
    Announcement,
    AnnouncementRead,
    ActivityLog,
    Indicator,
    Office,
    PerformanceRecord,
    StrategicLevel,
)
from .services import ingest_excel_monitor, ingest_word_monitor, ingest_word_monitor_with_stats


def is_admin(user):
    return user.is_superuser


def _full_database_clear_allowed():
    if settings.DEBUG:
        return True
    return getattr(settings, 'SOPM_ENABLE_FULL_DATABASE_CLEAR', False)


def _database_restore_allowed():
    if settings.DEBUG:
        return True
    return getattr(settings, 'SOPM_ENABLE_DATABASE_RESTORE', False)


_SQLITE_MAGIC = b'SQLite format 3\x00'


def _default_sqlite_db_path():
    name = settings.DATABASES['default']['NAME']
    return Path(os.fspath(name))


def _validate_sqlite_backup_path(path: Path):
    """Return (ok, error_message)."""
    if not path.is_file() or path.stat().st_size < 100:
        return False, 'File is missing or too small to be a SQLite database.'
    with open(path, 'rb') as f:
        if f.read(16) != _SQLITE_MAGIC:
            return False, 'File is not a SQLite 3 database.'
    try:
        cx = sqlite3.connect(os.fspath(path))
        try:
            rows = cx.execute(
                "SELECT name FROM sqlite_master WHERE type='table' "
                "AND name IN ('django_migrations', 'auth_user')"
            ).fetchall()
            names = {r[0] for r in rows}
            if 'django_migrations' not in names or 'auth_user' not in names:
                return False, (
                    'This file does not look like a Django/OPMM database '
                    '(expected tables django_migrations and auth_user).'
                )
        finally:
            cx.close()
    except sqlite3.Error as exc:
        return False, f'Could not read SQLite file: {exc}'
    return True, ''


def _office_for_user(user):
    if user.is_superuser:
        return None
    return Office.objects.filter(name=user.first_name).first()


def _activity_log(request, action, detail='', level='info', office=None):
    """Write a persistent activity log row (best-effort)."""
    try:
        ActivityLog.objects.create(
            level=level[:16] or 'info',
            action=(action or '')[:120] or 'Activity',
            detail=(detail or '')[:4000],
            user=request.user if getattr(request, 'user', None) and request.user.is_authenticated else None,
            office=office,
        )
    except Exception:
        # Logging should never block user actions.
        pass


def _plan_labels_for_indicator(indicator):
    """Outcome / strategy / PAP names for an Indicator (same tree walk as the performance matrix)."""
    pap = indicator.pap
    strat = pap.parent if pap else None
    out = strat.parent if strat else None
    outcome = (out.name if out else '').strip() or 'Outcome'
    strategy = (strat.name if strat else '').strip() or 'Strategy'
    pap_name = (pap.name if pap else '').strip() or 'Program / PAP'
    return outcome, strategy, pap_name


def _plan_labels_for_record(r):
    """Outcome / strategy / PAP names from the ingested strategic tree (one row per performance record)."""
    return _plan_labels_for_indicator(r.indicator)


_MATRIX_REMARKS_SPLIT = '\n\n— Remarks —\n'


def _annotate_record_matrix_columns(r):
    """Attach Word-style matrix columns for the dashboard table (accomplishment vs remarks split)."""
    o, s, p = _plan_labels_for_record(r)
    r.matrix_outcome = o
    r.matrix_strategy = s
    r.matrix_pap = p
    raw = (r.raw_actual_text or '').strip()
    if _MATRIX_REMARKS_SPLIT in raw:
        acmp, _, rem = raw.partition(_MATRIX_REMARKS_SPLIT)
        r.matrix_accomplishment = acmp.strip()
        r.matrix_remarks = rem.strip()
    else:
        r.matrix_accomplishment = raw
        r.matrix_remarks = ''
    r.matrix_status = r.status
    r.matrix_indicator = (r.indicator.description or '').strip()


def _viewer_strip_dev_area_prefix_from_outcome(outcome_label, area_display_name):
    """Show the matrix outcome text as in the source when ingest prefixed ``Area — Outcome``."""
    if not outcome_label or not (area_display_name or '').strip():
        return outcome_label
    prefix = area_display_name.strip() + ' — '
    if outcome_label.startswith(prefix):
        return outcome_label[len(prefix) :].strip()
    return outcome_label


def _viewer_variance_cell_class(variance_raw):
    """Table cell class for variance column (MET / UNMET / +N / −N), aligned with OPMM matrix."""
    if not (variance_raw or '').strip():
        return ''
    s = str(variance_raw).strip()
    su = s.upper()
    if su == 'MET':
        return 'viewer-var-cell-met'
    if su == 'UNMET':
        return 'viewer-var-cell-unmet'
    if re.match(r'^\s*\+\s*\d+\s*$', s):
        return 'viewer-var-cell-plus'
    if re.match(r'^\s*-\s*\d+\s*$', s):
        return 'viewer-var-cell-minus'
    return ''


def _viewer_hierarchy_triplet(r):
    """Stable (outcome_pk, strategy_pk, pap_pk) for rowspan grouping on the current page."""
    pap = r.indicator.pap
    if not pap:
        return (None, None, None)
    strat = pap.parent
    out = strat.parent if strat else None
    return (
        out.pk if out else None,
        strat.pk if strat else None,
        pap.pk if pap else None,
    )


def _viewer_augment_matrix_hierarchy(indicator_row_items, area_display_name=''):
    """Add Outcome/Strategy/PAP rowspans (within the paginated slice) and variance cell classes."""
    if not indicator_row_items:
        return indicator_row_items
    rows = [item['record'] for item in indicator_row_items]
    n = len(rows)
    triples = [_viewer_hierarchy_triplet(r) for r in rows]
    out_text = []
    strat_text = []
    pap_text = []
    for r in rows:
        o, s, p = _plan_labels_for_record(r)
        out_text.append(_viewer_strip_dev_area_prefix_from_outcome(o, area_display_name))
        strat_text.append(s)
        pap_text.append(p)

    out = []
    for i, item in enumerate(indicator_row_items):
        t0, t1, t2 = triples[i]

        if i == 0 or triples[i - 1][0] != t0:
            j = i
            while j < n and triples[j][0] == t0:
                j += 1
            o_rs = j - i
        else:
            o_rs = 0

        if i == 0 or triples[i - 1][:2] != (t0, t1):
            j = i
            while j < n and triples[j][:2] == (t0, t1):
                j += 1
            s_rs = j - i
        else:
            s_rs = 0

        if i == 0 or triples[i - 1] != triples[i]:
            j = i
            while j < n and triples[j] == triples[i]:
                j += 1
            p_rs = j - i
        else:
            p_rs = 0

        r = rows[i]
        var_cls = _viewer_variance_cell_class(r.variance_text or '')
        tone = (r.status or 'PENDING').strip().lower()
        if tone not in ('met', 'unmet', 'pending'):
            tone = 'pending'

        # Stable striping by outcome (same hue for every row under this outcome on the page).
        outcome_pk = t0 or 0
        outcome_band_idx = outcome_pk % 8

        out.append(
            {
                **item,
                'outcome_rowspan': o_rs,
                'strategy_rowspan': s_rs,
                'pap_rowspan': p_rs,
                'outcome_text': out_text[i],
                'strategy_text': strat_text[i],
                'pap_text': pap_text[i],
                'variance_cell_class': var_cls,
                'matrix_row_tone': tone,
                'outcome_band_idx': outcome_band_idx,
            }
        )
    return out


def _office_deletion_snapshot(office):
    """Counts and sample PAP titles for selective office data removal (same scope as clear_office_data)."""
    paps = StrategicLevel.objects.filter(office=office, level_type='PAP').order_by('name')
    pap_count = paps.count()
    pap_samples = list(paps.values_list('name', flat=True)[:15])
    indicator_count = Indicator.objects.filter(pap__office=office).count()
    record_count = PerformanceRecord.objects.filter(indicator__pap__office=office).count()
    return {
        'pap_count': pap_count,
        'indicator_count': indicator_count,
        'record_count': record_count,
        'pap_samples': pap_samples,
    }


@login_required
def dashboard_home(request):
    """Entry for sidebar Dashboard: superusers go to OPMM viewer with a full-FY overview query string."""
    if request.user.is_superuser and not request.GET:
        years = sorted(
            set(PerformanceRecord.objects.values_list('year', flat=True)),
            reverse=True,
        )
        default_year = years[0] if years else 2026
        q = f'year={default_year}&quarter=all&area=all'
        return redirect(f'{reverse("performance_viewer")}?{q}')

    params = request.GET.copy()
    target = '/performance-viewer/'
    if not request.user.is_superuser:
        office = _office_for_user(request.user)
        if office:
            params['office'] = str(office.id)
            if not request.GET:
                years = sorted(
                    set(
                        PerformanceRecord.objects.filter(
                            indicator__pap__office_id=office.id
                        ).values_list('year', flat=True)
                    ),
                    reverse=True,
                )
                if years:
                    params['year'] = str(years[0])
                params['quarter'] = 'all'
                params['area'] = 'all'
    qs = params.urlencode()
    if qs:
        target = f'{target}?{qs}'
    return redirect(target)


def _search_suggestion_focus(q, description, target_text, outcome, strategy, pap):
    """Pick which matrix field to emphasize in search dropdown (hierarchy before KPI text)."""
    needle = (q or '').casefold()
    if not needle:
        return 'indicator'

    def has(text):
        return needle in (text or '').casefold()

    if has(strategy):
        return 'strategy'
    if has(pap):
        return 'pap'
    if has(outcome):
        return 'outcome'
    if has(description) or has(target_text):
        return 'indicator'
    return 'indicator'


@login_required
def search_suggestions(request):
    q = (request.GET.get('q') or '').strip()
    # Allow single-character searches (e.g. short office abbreviations),
    # but keep empty queries silent to avoid noisy dropdown flashes.
    if len(q) < 1:
        return HttpResponse('')

    office = _office_for_user(request.user)
    indicators = Indicator.objects.select_related(
        'pap',
        'pap__parent',
        'pap__parent__parent',
        'pap__parent__parent__parent',
    ).order_by('pap__name', 'id')
    if not request.user.is_superuser:
        if office:
            indicators = indicators.filter(pap__office=office)
        else:
            indicators = indicators.none()
    # Search across the whole strategic chain visible in the matrix:
    # Indicator -> PAP -> Strategy -> Outcome (and optional deeper roots).
    # Name icontains on FK joins covers typical 3-level trees; ID-chain matching
    # covers varied depths and cases where joins would miss (same label logic as _plan_labels_for_record).
    #
    # IMPORTANT: ingest stores office only on PAP-level StrategicLevel rows; OUTCOME and STRATEGY
    # rows typically have office=NULL (shared across offices). Never filter StrategicLevel by
    # office here — that would exclude every outcome/strategy from matching. Scope is enforced
    # by Indicator.objects.filter(pap__office=office) above.
    base_filter = (
        Q(description__icontains=q)
        | Q(target_text__icontains=q)
        | Q(pap__name__icontains=q)
        | Q(pap__parent__name__icontains=q)
        | Q(pap__parent__parent__name__icontains=q)
        | Q(pap__parent__parent__parent__name__icontains=q)
    )
    sl_qs = StrategicLevel.objects.filter(
        name__icontains=q,
        level_type__in=('OUTCOME', 'STRATEGY', 'PAP'),
    )
    sl_ids = list(sl_qs.values_list('pk', flat=True))
    if sl_ids:
        ancestor_q = Q()
        path = 'pap'
        for _ in range(16):
            ancestor_q |= Q(**{f'{path}_id__in': sl_ids})
            path = f'{path}__parent'
        base_filter = base_filter | ancestor_q
    # If the user types a number, also match exact Indicator primary key.
    if q.isdigit():
        base_filter = base_filter | Q(pk=int(q))
    indicators = list(indicators.filter(base_filter)[:12])
    tag_for = {'outcome': 'Outcome', 'strategy': 'Strategy', 'pap': 'PAP', 'indicator': 'Indicator'}
    rows = []
    for ind in indicators:
        o, s, p = _plan_labels_for_indicator(ind)
        desc = ind.description or ''
        tgt = ind.target_text or ''
        focus = _search_suggestion_focus(q, desc, tgt, o, s, p)
        fields = {'outcome': o, 'strategy': s, 'pap': p, 'indicator': desc}
        primary_text = fields[focus]
        secondary = []
        for key in ('outcome', 'strategy', 'pap', 'indicator'):
            if key == focus:
                continue
            text = fields[key]
            if key == 'indicator' and not (text or '').strip():
                continue
            secondary.append({'tag': tag_for[key], 'text': text})
        rows.append(
            {
                'id': ind.pk,
                'focus': focus,
                'primary_tag': tag_for[focus],
                'primary_text': primary_text,
                'secondary': secondary,
            }
        )
    return render(request, 'core/partials/search_results.html', {'results': rows})


def _ingest_period_from_post(post):
    quarter = None
    year = None
    raw_q = (post.get('ingest_quarter') or '').strip()
    if raw_q in ('1', '2', '3', '4'):
        quarter = int(raw_q)
    raw_y = (post.get('ingest_year') or '').strip()
    if raw_y.isdigit() and len(raw_y) == 4:
        y = int(raw_y)
        if 2000 <= y <= 2099:
            year = y
    return quarter, year


def _ingest_validation_snapshot(office_name, year=None, quarter=None):
    """Post-ingest quality snapshot for one office and optional FY/quarter."""
    qs = PerformanceRecord.objects.select_related('indicator__pap__office').filter(
        indicator__pap__office__name=office_name
    )
    if year is not None:
        qs = qs.filter(year=year)
    if quarter is not None:
        qs = qs.filter(quarter=quarter)
    rows = list(qs)
    total = len(rows)
    met = sum(1 for r in rows if r.status == 'MET')
    unmet = sum(1 for r in rows if r.status == 'UNMET')
    pending = total - met - unmet
    indicators = len({r.indicator_id for r in rows})
    return {
        'indicators': indicators,
        'total': total,
        'met': met,
        'unmet': unmet,
        'pending': pending,
    }


@login_required
def upload_blueprint(request):
    if request.method == 'POST':
        files = [f for f in request.FILES.getlist('blueprint') if getattr(f, 'name', None)]
        if not files:
            return render(request, 'core/upload.html', {'error': 'No file selected.'})
        for uploaded in files:
            lower = uploaded.name.lower()
            if not (lower.endswith('.docx') or lower.endswith('.xlsx')):
                return render(
                    request,
                    'core/upload.html',
                    {
                        'error': (
                            f'"{uploaded.name}" is not supported. '
                            'Please upload .docx (Word) or .xlsx (Excel) monitor files.'
                        ),
                    },
                )

        post_q, post_y = _ingest_period_from_post(request.POST)
        office_name = 'OVCRDES'
        if not request.user.is_superuser:
            office = _office_for_user(request.user)
            if not office:
                return render(
                    request,
                    'core/upload.html',
                    {
                        'error': (
                            'Your account has no office on file. '
                            'Ask an administrator to register you with your full office name.'
                        ),
                    },
                )
            office_name = office.name

        failures = []
        successes = 0
        rows_ingested_total = 0
        ingest_breakdown_lines = []
        for uploaded in files:
            tmp_path = None
            try:
                suffix = '.xlsx' if uploaded.name.lower().endswith('.xlsx') else '.docx'
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
                tmp_path = tmp.name
                for chunk in uploaded.chunks():
                    tmp.write(chunk)
                tmp.close()
                hint = uploaded.name or ''
                if suffix == '.xlsx':
                    added = ingest_excel_monitor(
                        tmp_path,
                        office_name=office_name,
                        quarter=post_q,
                        year=post_y,
                        extra_hint=hint,
                    )
                    rows_ingested_total += added
                    _activity_log(
                        request,
                        'Upload ingested',
                        f'File={uploaded.name}; type=xlsx; rows={added}; office={office_name}; year={post_y or "auto"}; quarter={post_q or "auto"}',
                        office=Office.objects.filter(name=office_name).first(),
                    )
                else:
                    wstats = ingest_word_monitor_with_stats(
                        tmp_path,
                        office_name=office_name,
                        quarter=post_q,
                        year=post_y,
                        extra_hint=hint,
                    )
                    rows_ingested_total += wstats['rows_written']
                    bits = [f'{label}: {count}' for label, count in wstats.get('breakdown', [])]
                    if bits:
                        ingest_breakdown_lines.append(f'{uploaded.name} -> ' + ', '.join(bits))
                    _activity_log(
                        request,
                        'Upload ingested',
                        f'File={uploaded.name}; type=docx; rows={wstats["rows_written"]}; breakdown={"; ".join(bits)}; office={office_name}; year={post_y or "auto"}; quarter={post_q or "auto"}',
                        office=Office.objects.filter(name=office_name).first(),
                    )
                successes += 1
            except Exception as exc:
                failures.append((uploaded.name, str(exc)))
                _activity_log(
                    request,
                    'Upload failed',
                    f'File={uploaded.name}; error={exc}',
                    level='error',
                    office=Office.objects.filter(name=office_name).first(),
                )
            finally:
                if tmp_path and os.path.isfile(tmp_path):
                    os.unlink(tmp_path)

        if successes and not failures:
            snap = _ingest_validation_snapshot(office_name, year=post_y, quarter=post_q)
            scope_bits = []
            if post_y is not None:
                scope_bits.append(f'FY{post_y}')
            if post_q is not None:
                scope_bits.append(f'Q{post_q}')
            scope = '-'.join(scope_bits) if scope_bits else 'office'
            breakdown_q = ''
            if ingest_breakdown_lines:
                breakdown_q = '&bd=' + quote('\n'.join(ingest_breakdown_lines[:12]))
            return redirect(
                f'{reverse("success_page")}?n={successes}&rows={rows_ingested_total}'
                f'&scope={scope}&ind={snap["indicators"]}&tot={snap["total"]}'
                f'&met={snap["met"]}&unmet={snap["unmet"]}&pending={snap["pending"]}'
                f'{breakdown_q}'
            )

        if successes and failures:
            return render(
                request,
                'core/upload.html',
                {
                    'success_count': successes,
                    'failure_details': failures,
                },
            )

        err_parts = [f'{name}: {err}' for name, err in failures]
        return render(
            request,
            'core/upload.html',
            {'error': 'Could not ingest any file. ' + ' | '.join(err_parts)},
        )
    return render(request, 'core/upload.html')


@login_required
def success_page(request):
    raw_n = (request.GET.get('n') or '').strip()
    if raw_n.isdigit() and int(raw_n) >= 1:
        ingested_count = int(raw_n)
    else:
        ingested_count = 1
    raw_rows = (request.GET.get('rows') or '').strip()
    rows_ingested = int(raw_rows) if raw_rows.isdigit() else None
    raw_ind = (request.GET.get('ind') or '').strip()
    raw_tot = (request.GET.get('tot') or '').strip()
    raw_met = (request.GET.get('met') or '').strip()
    raw_unmet = (request.GET.get('unmet') or '').strip()
    raw_pending = (request.GET.get('pending') or '').strip()
    raw_breakdown = (request.GET.get('bd') or '').strip()
    val_indicators = int(raw_ind) if raw_ind.isdigit() else None
    val_total = int(raw_tot) if raw_tot.isdigit() else None
    val_met = int(raw_met) if raw_met.isdigit() else None
    val_unmet = int(raw_unmet) if raw_unmet.isdigit() else None
    val_pending = int(raw_pending) if raw_pending.isdigit() else None
    scope_raw = (request.GET.get('scope') or '').strip()
    if scope_raw == 'office':
        validation_scope = 'Office-wide (all periods)'
    elif scope_raw.startswith('FY'):
        validation_scope = scope_raw.replace('-', ' · ')
    else:
        validation_scope = ''
    ingest_breakdown_lines = [line.strip() for line in raw_breakdown.splitlines() if line.strip()]
    return render(
        request,
        'core/success.html',
        {
            'ingested_count': ingested_count,
            'rows_ingested': rows_ingested,
            'validation_scope': validation_scope,
            'val_indicators': val_indicators,
            'val_total': val_total,
            'val_met': val_met,
            'val_unmet': val_unmet,
            'val_pending': val_pending,
            'ingest_breakdown_lines': ingest_breakdown_lines,
        },
    )


def _viewer_outcome_for_record(record):
    """Walk PAP -> Strategy -> Outcome (StrategicLevel parents) for a record."""
    pap = record.indicator.pap
    strategy = pap.parent if pap else None
    return strategy.parent if strategy else None


def _viewer_monitor_sort_parts(record):
    """Order rows like the uploaded monitor: Outcome → Strategy → PAP → row order (ingest id)."""
    pap = record.indicator.pap
    strategy = pap.parent if pap else None
    outcome = strategy.parent if strategy else None
    out_name = (outcome.name or '').strip() if outcome else ''
    strat_name = (strategy.name or '').strip() if strategy else ''
    pap_name = (pap.name or '').strip() if pap else ''
    # ``record.id`` preserves upload order within the same PAP better than ``indicator_id``.
    return (out_name.lower(), strat_name.lower(), pap_name.lower(), record.id)


# OPMM matrix: six development areas only (fixed order, like spreadsheet tabs).
# Needles are matched against a compact lowercase alphanumeric blob (see
# ``_viewer_record_dev_area_key``). Use distinctive phrases; order within a
# spec is longest-first so partial phrases win before shorter overlaps.
VIEWER_DEVELOPMENT_AREAS = (
    {
        'key': 'academic_leadership',
        'display': 'Academic Leadership',
        'needles': (
            'academicleadership',
            'academicexcellence',
            'academicquality',
            'academicdevelopment',
            'studentaffairs',
            'studentdevelopment',
            'teachingexcellence',
            'learningoutcomes',
        ),
    },
    {
        'key': 'social_responsibility',
        'display': 'Social Responsibility',
        'needles': (
            'socialresponsibility',
            # RDE / OVCRES-style mandates: match before pure "research" buckets.
            'researchdevelopmentandextension',
            'researchdevelopmentextension',
            'extensionservices',
            'extensionservice',
            'extensionprogram',
            'extensionoffice',
            'extensionfunction',
            'extensionactivities',
            'extension',
            'communityengagement',
            'communityextension',
            'communityoutreach',
            'communitydevelopment',
            'communityservice',
            'publicservice',
            'servicelearning',
            'outreachprogram',
            'lifelonglearning',
            'socialimpact',
            # Common Philippine OP / HEI program tags (often extension + research mix).
            'ibar',
            'aanr',
            'ieetr',
            'drrm',
            'disasterrisk',
            'healthprograms',
            'technologymobilization',
            'adoptaschool',
            'livelihood',
        ),
    },
    {
        'key': 'research_and_innovation',
        'display': 'Research and Innovation',
        'needles': (
            'researchandinnovation',
            'researchdevelopment',
            'researchmodernization',
            'researchagenda',
            'strategicresearch',
            'externallyfundedresearch',
            'researchcollaboration',
            'researchpartnership',
            'researchprogram',
            'researchproject',
            'researchproposal',
            'researchinfrastructure',
            'scienceandtechnology',
            'technologytransfer',
            'innovationecosystem',
            'knowledgeproduction',
            'randd',
            'submittedproposal',
            'submittedproposals',
            'fundedresearch',
            'highimpact',
            'publicationimpact',
            'sustainabledevelopment',
            'innovativesolutions',
        ),
    },
    {
        'key': 'internationalization',
        'display': 'Internationalization',
        'needles': (
            'internationalization',
            'internationalisation',
            'studentmobility',
            'outboundexchange',
            'inboundexchange',
            'globalpartnership',
            'foreignpartnership',
            'internationalranking',
        ),
    },
    {
        'key': 'advancing_interdisciplinarity',
        'display': 'Advancing Interdisciplinarity',
        'needles': (
            'advancinginterdisciplinarity',
            'interdisciplinarity',
            'interdisciplinary',
            'crossdisciplinary',
            'multidisciplinary',
            'transdisciplinary',
            'jointdegree',
            'jointprogram',
            'crosslisted',
        ),
    },
    {
        'key': 'sustainability',
        'display': 'Sustainability',
        'needles': (
            'sustainability',
            'environmentalstewardship',
            'climateaction',
            'carbonfootprint',
            'carbonreduction',
            'greening',
            'wastemanagement',
            'energymanagement',
            'campussustainability',
        ),
    },
)

_VIEWER_DEV_AREA_KEYS = frozenset(s['key'] for s in VIEWER_DEVELOPMENT_AREAS)
VIEWER_AREA_ALL = 'all'

# Official OPMM display order (six pillars). Used for dashboard / viewer lists only;
# ``VIEWER_DEVELOPMENT_AREAS`` tuple order stays optimized for text matching.
OPMM_DASHBOARD_AREA_ORDER = (
    'sustainability',
    'academic_leadership',
    'research_and_innovation',
    'internationalization',
    'social_responsibility',
    'advancing_interdisciplinarity',
)


def _viewer_compact_outcome_name(name):
    if not name:
        return ''
    return re.sub(r'[^a-z0-9]', '', name.lower())


def _viewer_dev_area_key_from_compact_text(compact):
    """First matching needle wins (``VIEWER_DEVELOPMENT_AREAS`` order)."""
    if not compact:
        return 'social_responsibility'
    for spec in VIEWER_DEVELOPMENT_AREAS:
        for needle in spec['needles']:
            if needle in compact:
                return spec['key']
    # Raw OPMM for OVCRDES rarely uses the Sustainability pillar; unlabeled rows skew Social/R&I.
    return 'social_responsibility'


def _viewer_dev_area_key_from_outcome(outcome):
    """Map an OUTCOME StrategicLevel to one of the six canonical development areas.

    Used when only the outcome node is available (e.g. legacy URL mapping).
    """
    if outcome is None:
        return 'social_responsibility'
    return _viewer_dev_area_key_from_compact_text(_viewer_compact_outcome_name(outcome.name))


def _viewer_record_dev_area_key(record):
    """Bucket a performance row using outcome, strategy, PAP, and indicator text.

    LPC / FYDP uploads often carry the true development area in a merged
    outcome title, strategy line, PAP label, or indicator wording. Matching
    only the outcome name sent many legitimate R&I / Social rows to the
    Sustainability fallback.
    """
    parts = []
    out = _viewer_outcome_for_record(record)
    if out and out.name:
        parts.append(out.name)
    pap = record.indicator.pap
    strat = pap.parent if pap else None
    if strat and strat.name:
        parts.append(strat.name)
    if pap and pap.name:
        parts.append(pap.name)
    ind = record.indicator
    if ind and ind.description:
        parts.append(ind.description)
    blob = _viewer_compact_outcome_name(' '.join(parts))
    return _viewer_dev_area_key_from_compact_text(blob)


def _build_office_chart_payload(records_list, max_offices=14):
    """Aggregate met/unmet/pending per office for charts (stacked counts + met %).

    Only offices with at least one performance row in ``records_list`` appear (no
    zero-only placeholders for fiscal years where an office has no ingest).
    """
    from collections import defaultdict

    agg = defaultdict(lambda: {'name': '', 'met': 0, 'unmet': 0, 'pending': 0})
    for r in records_list:
        office = r.indicator.pap.office
        if office is None:
            continue
        oid = office.id
        if not agg[oid]['name']:
            agg[oid]['name'] = office.name
        st = r.status
        if st == 'MET':
            agg[oid]['met'] += 1
        elif st == 'UNMET':
            agg[oid]['unmet'] += 1
        else:
            agg[oid]['pending'] += 1

    rows = []
    for oid, d in agg.items():
        total = d['met'] + d['unmet'] + d['pending']
        if total < 1:
            continue
        rows.append((oid, d['name'], d['met'], d['unmet'], d['pending'], total))

    rows.sort(key=lambda x: (-x[2], x[3], -x[5], (x[1] or '').lower()))
    rows = rows[:max_offices]

    labels = []
    for _oid, name, _met, _unmet, _pen, _total in rows:
        short = name if len(name) <= 34 else name[:31] + '…'
        labels.append(short)

    met_pcts = []
    scored_pcts = []
    for r in rows:
        met, unmet, pending, total = r[2], r[3], r[4], r[5]
        met_pcts.append(round(100 * met / total, 1) if total else 0.0)
        scored = met + unmet
        scored_pcts.append(round(100 * met / scored, 1) if scored else 0.0)

    return {
        'labels': labels,
        'met': [r[2] for r in rows],
        'unmet': [r[3] for r in rows],
        'pending': [r[4] for r in rows],
        'totals': [r[5] for r in rows],
        'met_pct': met_pcts,
        'met_pct_scored_only': scored_pcts,
    }


def _build_quarterly_trend_payload(records_list):
    """Q1–Q4 met/unmet/pending and % met (scored) for Chart.js (analytics)."""
    q_rows = []
    running = 0
    for q in (1, 2, 3, 4):
        q_subset = [r for r in records_list if r.quarter == q]
        q_met = sum(1 for r in q_subset if r.status == 'MET')
        q_unmet = sum(1 for r in q_subset if r.status == 'UNMET')
        q_pend = len(q_subset) - q_met - q_unmet
        scored = q_met + q_unmet
        pct = round(100 * q_met / scored, 1) if scored else None
        running += q_met
        q_rows.append(
            {
                'q': f'Q{q}',
                'met': q_met,
                'unmet': q_unmet,
                'pending': max(0, q_pend),
                'met_pct_scored': pct,
                'cum_met': running,
            }
        )
    return {
        'labels': [x['q'] for x in q_rows],
        'met': [x['met'] for x in q_rows],
        'unmet': [x['unmet'] for x in q_rows],
        'pending': [x['pending'] for x in q_rows],
        'met_pct_scored': [x['met_pct_scored'] for x in q_rows],
        'cum_met': [x['cum_met'] for x in q_rows],
    }


def _analytics_extract_partner_counts(records, top_n=12):
    """Best-effort partner mentions from accomplishment narratives."""
    ignore = {
        'FY', 'Q1', 'Q2', 'Q3', 'Q4', 'MET', 'UNMET', 'IPR', 'PAP', 'OVCRDES',
        'SOPM', 'LPC', 'OPMM', 'CICS', 'CABE', 'CAS', 'CENG', 'COE',
    }
    counts = Counter()
    for r in records:
        txt = (r.raw_actual_text or '').strip()
        if not txt:
            continue
        for token in re.findall(r'\b[A-Z][A-Za-z0-9&\-/]{2,}\b', txt):
            t = token.strip('.,;:()[]{}')
            if len(t) < 3 or t.upper() in ignore or t.isdigit():
                continue
            counts[t] += 1
    top = counts.most_common(top_n)
    return {'labels': [x[0] for x in top], 'values': [x[1] for x in top]}


def _analytics_extract_training_heat(records):
    """Approximate training impact by college/unit keywords in narratives."""
    units = ['CICS', 'CABE', 'CAS', 'CENG', 'COE', 'CTHM', 'CIT', 'CCS', 'CBA', 'Nursing']
    counts = Counter({u: 0 for u in units})
    for r in records:
        txt = (r.raw_actual_text or '')
        low = txt.lower()
        val = r.actual_value if r.actual_value is not None else 1
        weight = int(val) if isinstance(val, (int, float)) and val > 0 else 1
        for u in units:
            if u.lower() in low:
                counts[u] += weight
    pairs = [(k, v) for k, v in counts.items() if v > 0]
    pairs.sort(key=lambda x: x[1], reverse=True)
    return {'labels': [x[0] for x in pairs], 'values': [x[1] for x in pairs]}


def _performance_viewer_scope(request, *, export_mode=False):
    """Resolve year / quarter / office / area / status filters (shared by matrix and LPC export).

    When ``export_mode`` is True (LPC Excel download), ignore matrix filters for quarter,
    development area, and status: export every ``PerformanceRecord`` for the selected
    fiscal year and office scope so the file matches full institutional data.
    """
    base = PerformanceRecord.objects.select_related(
        'indicator__pap__parent__parent',
        'indicator__pap__office',
    )

    if not request.user.is_superuser:
        user_office = _office_for_user(request.user)
        if user_office:
            base = base.filter(indicator__pap__office=user_office)
        else:
            base = base.none()

    office_ids = [
        oid
        for oid in base.values_list('indicator__pap__office_id', flat=True).distinct()
        if oid is not None
    ]
    available_offices = list(
        Office.objects.filter(id__in=office_ids).order_by('name').values('id', 'name')
    )

    office_id = None
    if request.user.is_superuser:
        raw_office = (request.GET.get('office') or '').strip()
        if raw_office.isdigit():
            candidate = int(raw_office)
            if candidate in office_ids:
                office_id = candidate
    else:
        user_office = _office_for_user(request.user)
        if user_office:
            office_id = user_office.id

    scope_for_years = base
    if office_id is not None:
        scope_for_years = scope_for_years.filter(indicator__pap__office_id=office_id)

    available_years = sorted(set(scope_for_years.values_list('year', flat=True)), reverse=True)
    if not available_years:
        available_years = [2025]
    try:
        year = int(request.GET.get('year') or available_years[0])
    except ValueError:
        year = available_years[0]
    if year not in available_years:
        year = available_years[0]

    year_records = scope_for_years.filter(year=year)
    year_office_ids = set(
        oid
        for oid in year_records.values_list('indicator__pap__office_id', flat=True).distinct()
        if oid is not None
    )
    combined_offices = office_id is None and len(year_office_ids) > 1
    selected_office_name = ''
    if office_id is not None:
        selected_office_name = next(
            (o['name'] for o in available_offices if o['id'] == office_id),
            '',
        )
    available_quarters = sorted(set(year_records.values_list('quarter', flat=True)))
    quarter_options = [1, 2, 3, 4]

    if export_mode:
        view_all_quarters = True
        quarter = None
        period_records = list(year_records)
    else:
        raw_quarter = (request.GET.get('quarter') or '').strip().lower()
        view_all_quarters = False
        quarter = None
        if raw_quarter == 'all':
            view_all_quarters = True
        elif not raw_quarter:
            # First load: full fiscal year across quarters (office users and superusers).
            view_all_quarters = True
        else:
            try:
                qn = int(raw_quarter) if raw_quarter else 0
            except ValueError:
                qn = 0
            if qn in quarter_options:
                quarter = qn
            elif available_quarters:
                quarter = available_quarters[-1]
            else:
                view_all_quarters = True

        if view_all_quarters:
            period_records = list(year_records)
        else:
            period_records = list(year_records.filter(quarter=quarter))

    counts = {spec['key']: {'met': 0, 'total': 0} for spec in VIEWER_DEVELOPMENT_AREAS}
    for r in period_records:
        key = _viewer_record_dev_area_key(r)
        counts[key]['total'] += 1
        if r.status == 'MET':
            counts[key]['met'] += 1

    development_areas = []
    for aid in OPMM_DASHBOARD_AREA_ORDER:
        spec = next(s for s in VIEWER_DEVELOPMENT_AREAS if s['key'] == aid)
        key = spec['key']
        c = counts[key]
        pct = round((c['met'] / c['total']) * 100) if c['total'] else 0
        development_areas.append(
            {
                'id': key,
                'name': spec['display'],
                'met': c['met'],
                'total': c['total'],
                'met_pct': pct,
            }
        )
    all_dev_total = sum(a['total'] for a in development_areas)
    all_dev_met = sum(a['met'] for a in development_areas)
    all_dev_met_pct = (
        round((all_dev_met / all_dev_total) * 100) if all_dev_total else 0
    )
    if export_mode:
        selected_area_id = VIEWER_AREA_ALL
        area_records = list(period_records)
        selected_area_name = 'All development areas'
    else:
        raw_area = (request.GET.get('area') or '').strip()
        area_slug = re.sub(r'[^a-z0-9_]', '', raw_area.lower().replace(' ', '_').replace('-', '_'))

        selected_area_id = None
        available_area_ids = {a['id'] for a in development_areas}
        if area_slug == VIEWER_AREA_ALL or raw_area.lower() == VIEWER_AREA_ALL:
            selected_area_id = VIEWER_AREA_ALL
        elif area_slug in available_area_ids:
            selected_area_id = area_slug
        elif raw_area.isdigit():
            legacy = StrategicLevel.objects.filter(
                pk=int(raw_area), level_type='OUTCOME'
            ).first()
            if legacy:
                mapped_key = _viewer_dev_area_key_from_outcome(legacy)
                if mapped_key and mapped_key in available_area_ids:
                    selected_area_id = mapped_key
        if selected_area_id is None:
            if not raw_area:
                # First load / overview: all pillars (explicit ?area= still selects one bucket).
                selected_area_id = VIEWER_AREA_ALL
            else:
                selected_area = max(development_areas, key=lambda a: a['total'], default=None)
                selected_area_id = (
                    selected_area['id'] if selected_area else VIEWER_DEVELOPMENT_AREAS[0]['key']
                )

        if selected_area_id == VIEWER_AREA_ALL:
            area_records = list(period_records)
        else:
            area_records = [
                r for r in period_records if _viewer_record_dev_area_key(r) == selected_area_id
            ]

        if selected_area_id == VIEWER_AREA_ALL:
            selected_area_name = 'All development areas'
        else:
            selected_area_name = next(
                (a['name'] for a in development_areas if a['id'] == selected_area_id),
                '',
            )

    total = len(area_records)
    met = sum(1 for r in area_records if r.status == 'MET')
    unmet = sum(1 for r in area_records if r.status == 'UNMET')
    pending = total - met - unmet
    completion_pct = round((met / total) * 100) if total else 0
    met_pct = round((met / total) * 100) if total else 0
    unmet_pct = round((unmet / total) * 100) if total else 0

    if export_mode:
        status_filter = 'all'
        scoped_records = area_records
    else:
        raw_status = (request.GET.get('status') or '').strip().lower()
        if raw_status in ('met', 'unmet', 'pending'):
            status_filter = raw_status
        else:
            status_filter = 'all'

        scoped_records = area_records
        if status_filter != 'all':
            scoped_records = [
                r for r in area_records if r.status.lower() == status_filter
            ]

    indicator_rows_full = list(scoped_records)
    if view_all_quarters:
        indicator_rows_full.sort(
            key=lambda r: (r.quarter,) + _viewer_monitor_sort_parts(r),
        )
    else:
        indicator_rows_full.sort(key=_viewer_monitor_sort_parts)

    return {
        'available_offices': available_offices,
        'office_id': office_id,
        'year': year,
        'quarter': quarter,
        'view_all_quarters': view_all_quarters,
        'available_years': available_years,
        'quarter_options': quarter_options,
        'available_quarters': available_quarters,
        'combined_offices': combined_offices,
        'selected_office_name': selected_office_name,
        'development_areas': development_areas,
        'all_dev_total': all_dev_total,
        'all_dev_met_pct': all_dev_met_pct,
        'selected_area_id': selected_area_id,
        'selected_area_name': selected_area_name,
        'total': total,
        'met': met,
        'unmet': unmet,
        'pending': pending,
        'completion_pct': completion_pct,
        'met_pct': met_pct,
        'unmet_pct': unmet_pct,
        'status_filter': status_filter,
        'indicator_rows_full': indicator_rows_full,
    }


@login_required
def performance_viewer(request):
    """OPMM Viewer: admin-only performance summary with quarter tabs and 3-pane layout."""
    PAGE_SIZE = 25
    scope = _performance_viewer_scope(request)
    available_offices = scope['available_offices']
    office_id = scope['office_id']
    year = scope['year']
    quarter = scope['quarter']
    view_all_quarters = scope['view_all_quarters']
    available_years = scope['available_years']
    quarter_options = scope['quarter_options']
    available_quarters = scope['available_quarters']
    combined_offices = scope['combined_offices']
    selected_office_name = scope['selected_office_name']
    development_areas = scope['development_areas']
    all_dev_total = scope['all_dev_total']
    all_dev_met_pct = scope['all_dev_met_pct']
    selected_area_id = scope['selected_area_id']
    selected_area_name = scope['selected_area_name']
    status_filter = scope['status_filter']
    indicator_rows_full = list(scope['indicator_rows_full'])

    total = scope['total']
    met = scope['met']
    unmet = scope['unmet']
    pending = scope['pending']
    completion_pct = scope['completion_pct']
    met_pct = scope['met_pct']
    unmet_pct = scope['unmet_pct']

    selected_record = None
    page = 1
    raw_record = (request.GET.get('record') or '').strip()
    raw_indicator = (request.GET.get('indicator') or request.GET.get('focus') or '').strip()

    if raw_record.isdigit():
        rid = int(raw_record)
        for idx, r in enumerate(indicator_rows_full):
            if r.pk == rid:
                selected_record = r
                page = idx // PAGE_SIZE + 1
                break

    if selected_record is None and raw_indicator.isdigit():
        sid = int(raw_indicator)
        for idx, r in enumerate(indicator_rows_full):
            if r.indicator_id == sid:
                selected_record = r
                page = idx // PAGE_SIZE + 1
                break

    num_rows_pre = len(indicator_rows_full)
    num_pages_pre = max(1, ceil(num_rows_pre / PAGE_SIZE)) if num_rows_pre else 1

    if selected_record is None:
        try:
            page = int(request.GET.get('page') or 1)
        except ValueError:
            page = 1
        page = max(1, min(num_pages_pre, page))

    indicator_focus_active = raw_indicator.isdigit()
    if indicator_focus_active:
        sid = int(raw_indicator)
        indicator_rows_full = [r for r in indicator_rows_full if r.indicator_id == sid]
        if indicator_rows_full:
            if selected_record is None or getattr(selected_record, 'indicator_id', None) != sid:
                selected_record = indicator_rows_full[0]
            page = 1
            nlen = len(indicator_rows_full)
            total = nlen
            met = sum(1 for r in indicator_rows_full if r.status == 'MET')
            unmet = sum(1 for r in indicator_rows_full if r.status == 'UNMET')
            pending = nlen - met - unmet
            completion_pct = round((met / nlen) * 100) if nlen else 0
            met_pct = round((met / nlen) * 100) if nlen else 0
            unmet_pct = round((unmet / nlen) * 100) if nlen else 0
        else:
            selected_record = None
            total = met = unmet = pending = 0
            completion_pct = met_pct = unmet_pct = 0

    pap_counts = Counter(r.indicator.pap_id for r in indicator_rows_full)
    pap_running = {}
    pap_kpi_meta_by_pk = {}
    for r in indicator_rows_full:
        pid = r.indicator.pap_id
        pap_running[pid] = pap_running.get(pid, 0) + 1
        pap_kpi_meta_by_pk[r.pk] = {
            'pap_kpi_index': pap_running[pid],
            'pap_kpi_total': pap_counts[pid],
        }

    num_rows = len(indicator_rows_full)
    num_pages = max(1, ceil(num_rows / PAGE_SIZE)) if num_rows else 1
    if not indicator_focus_active and selected_record is None:
        page = max(1, min(num_pages, page))

    start = (page - 1) * PAGE_SIZE
    indicator_rows = indicator_rows_full[start : start + PAGE_SIZE]
    page_end_index = start + len(indicator_rows)

    for r in indicator_rows:
        _annotate_record_matrix_columns(r)

    pap_group_idx = -1
    prev_pid = None
    indicator_row_items = []
    for r in indicator_rows:
        pid = r.indicator.pap_id
        if pid != prev_pid:
            pap_group_idx += 1
            pap_group_start = True
            prev_pid = pid
        else:
            pap_group_start = False
        meta = pap_kpi_meta_by_pk.get(
            r.pk, {'pap_kpi_index': 1, 'pap_kpi_total': 1}
        )
        indicator_row_items.append(
            {
                'record': r,
                'pap_group_start': pap_group_start,
                'pap_band': pap_group_idx % 2,
                'pap_kpi_index': meta['pap_kpi_index'],
                'pap_kpi_total': meta['pap_kpi_total'],
            }
        )

    area_strip = selected_area_name if selected_area_id != VIEWER_AREA_ALL else ''
    indicator_row_items = _viewer_augment_matrix_hierarchy(indicator_row_items, area_strip)

    table_colspan = (
        1  # row #
        + (1 if view_all_quarters else 0)
        + 4  # outcome, strategy, pap, indicator
        + (1 if combined_offices else 0)
        + 5  # target, accomplishment, variance, status, remarks
    )

    ctx = {
        'year': year,
        'quarter': quarter,
        'view_all_quarters': view_all_quarters,
        'period_label': 'All quarters' if view_all_quarters else f'Q{quarter}',
        'available_years': available_years,
        'quarter_options': quarter_options,
        'available_quarters': available_quarters,
        'total': total,
        'met': met,
        'unmet': unmet,
        'pending': pending,
        'completion_pct': completion_pct,
        'met_pct': met_pct,
        'unmet_pct': unmet_pct,
        'development_areas': development_areas,
        'all_dev_total': all_dev_total,
        'all_dev_met_pct': all_dev_met_pct,
        'selected_area_id': selected_area_id,
        'selected_area_name': selected_area_name,
        'indicator_rows': indicator_rows,
        'indicator_row_items': indicator_row_items,
        'total_indicator_rows': num_rows,
        'selected_record': selected_record,
        'page': page,
        'num_pages': num_pages,
        'page_size': PAGE_SIZE,
        'page_start_index': start,
        'page_end_index': page_end_index,
        'prev_page': page - 1 if page > 1 else None,
        'next_page': page + 1 if page < num_pages else None,
        'viewer_table_colspan': table_colspan,
        'available_offices': available_offices,
        'office_id': office_id,
        'combined_offices': combined_offices,
        'selected_office_name': selected_office_name,
        'status_filter': status_filter,
        'indicator_focus_active': indicator_focus_active,
    }
    pop_qs = list(Announcement.objects.visible_for(request.user)[:1])
    ctx['viewer_popup_announcement'] = pop_qs[0] if pop_qs else None
    if request.headers.get('HX-Request'):
        return render(request, 'core/partials/performance_viewer_content.html', ctx)
    return render(request, 'core/performance_viewer.html', ctx)


@login_required
def export_lpc_wide_excel(request):
    """Download FYDP / LPC wide .xlsx matching institutional column headers (re-importable layout).

    Exports every performance row for the selected **fiscal year** and **office** filter (same as
    the viewer), including **all quarters**, **all development-area buckets**, and **all statuses**
    — independent of the matrix tabs (quarter / area / Met filter) on screen.
    """
    from .lpc_wide_export import build_lpc_wide_workbook_bytes

    scope = _performance_viewer_scope(request, export_mode=True)
    records = scope['indicator_rows_full']
    year = scope['year']
    view_all = scope['view_all_quarters']
    office_name = (scope.get('selected_office_name') or '').strip()
    if office_name:
        banner = f'{office_name} · FY {year} — all quarters — all development areas & statuses'
    else:
        banner = f'FY {year} — all quarters — all offices (combined) — all development areas & statuses'

    payload = build_lpc_wide_workbook_bytes(
        records,
        dev_area_banner=banner,
        report_year=year,
        view_all_quarters=view_all,
    )
    off_slug = 'ALL_OFFICES' if not scope.get('office_id') else f"OFFICE_{scope['office_id']}"
    fname = f'LPC_wide_FY{year}_ALL_QTR_{off_slug}.xlsx'
    resp = HttpResponse(
        payload,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    resp['Content-Disposition'] = f'attachment; filename="{fname}"'
    resp['X-Export-Row-Count'] = str(len(records))
    return resp


def _analytics_row_tooltip_label_for_pillar(record, pillar_key):
    """Text to show for one row in a pillar's analytics tooltip.

    Rows are bucketed using outcome+strategy+PAP+indicator combined text, but the
    outcome node alone may belong to a different pillar (merged cells / carry-down).
    Prefer the first single field whose needle match equals ``pillar_key`` so Social
    Responsibility (etc.) lists the line that actually triggered the bucket.
    """
    out = _viewer_outcome_for_record(record)
    o_text = (out.name if out else '').strip()
    pap = record.indicator.pap
    strat = pap.parent if pap else None
    s_text = (strat.name if strat else '').strip()
    p_text = (pap.name if pap else '').strip()
    ind = record.indicator
    i_text = (ind.description if ind else '').strip() if ind else ''

    def area_key_for_single_field(text):
        t = (text or '').strip()
        if not t:
            return None
        return _viewer_dev_area_key_from_compact_text(_viewer_compact_outcome_name(t))

    for label in (o_text, s_text, p_text, i_text):
        if not label:
            continue
        if area_key_for_single_field(label) == pillar_key:
            return label
    if o_text:
        return o_text
    if s_text:
        return s_text
    if p_text:
        return p_text
    if i_text:
        return i_text
    return ''


def _analytics_distinct_pillar_tooltip_lines(subset, pillar_key, max_lines=22, max_chars=220):
    """Distinct per-row labels for a pillar's stacked-bar tooltip (aligned with bucketing)."""
    seen = set()
    lines = []
    for r in subset:
        raw = _analytics_row_tooltip_label_for_pillar(r, pillar_key).strip()
        if not raw:
            continue
        norm = re.sub(r'\s+', ' ', raw).casefold()
        if norm in seen:
            continue
        seen.add(norm)
        display = raw if len(raw) <= max_chars else raw[: max_chars - 1] + '…'
        lines.append(display)
    lines.sort(key=lambda s: s.casefold())
    if len(lines) > max_lines:
        extra = len(lines) - max_lines
        lines = lines[:max_lines] + [f'… and {extra} more distinct label(s)']
    return lines


@login_required
def analytics_dashboard(request):
    """Analytics view for variance, hierarchy contribution, trend, partners, and training impact."""
    base = PerformanceRecord.objects.select_related(
        'indicator__pap__parent__parent',
        'indicator__pap__office',
    )
    available_years = sorted(set(base.values_list('year', flat=True)), reverse=True) or [2025]
    try:
        year = int(request.GET.get('year') or available_years[0])
    except ValueError:
        year = available_years[0]
    if year not in available_years:
        year = available_years[0]

    office_ids = [x for x in base.values_list('indicator__pap__office_id', flat=True).distinct() if x]
    available_offices = list(Office.objects.filter(id__in=office_ids).order_by('name').values('id', 'name'))
    can_choose_office = request.user.is_superuser
    office_id = None
    scoped_office_name = None
    if request.user.is_superuser:
        raw_off = (request.GET.get('office') or '').strip()
        if raw_off.isdigit():
            cand = int(raw_off)
            if cand in office_ids:
                office_id = cand
        if office_id:
            match = next((o for o in available_offices if o['id'] == office_id), None)
            if match:
                scoped_office_name = match['name']
    else:
        office = _office_for_user(request.user)
        if office:
            office_id = office.id
            scoped_office_name = office.name
        else:
            office_id = -1

    qs = base.filter(year=year)
    if office_id is not None:
        qs = qs.filter(indicator__pap__office_id=office_id)
    records = list(qs)

    variance_rows = []
    for r in records:
        if r.target_value is None or r.actual_value is None:
            continue
        gap = float(r.actual_value) - float(r.target_value)
        variance_rows.append((r.indicator.description[:68], round(gap, 2)))
    variance_rows.sort(key=lambda x: abs(x[1]), reverse=True)
    variance_rows = variance_rows[:16]
    variance_payload = {'labels': [x[0] for x in variance_rows], 'values': [x[1] for x in variance_rows]}

    area_payload = {'labels': [], 'met': [], 'unmet': [], 'pending': [], 'dev_area_lines': []}
    for aid in OPMM_DASHBOARD_AREA_ORDER:
        spec = next(s for s in VIEWER_DEVELOPMENT_AREAS if s['key'] == aid)
        subset = [r for r in records if _viewer_record_dev_area_key(r) == aid]
        met_n = sum(1 for r in subset if r.status == 'MET')
        unmet_n = sum(1 for r in subset if r.status == 'UNMET')
        pend_n = len(subset) - met_n - unmet_n
        area_payload['labels'].append(spec['display'])
        area_payload['met'].append(met_n)
        area_payload['unmet'].append(unmet_n)
        area_payload['pending'].append(max(0, pend_n))
        area_payload['dev_area_lines'].append(
            _analytics_distinct_pillar_tooltip_lines(subset, aid)
        )

    trend_payload = _build_quarterly_trend_payload(records)

    partner_payload = _analytics_extract_partner_counts(records)
    training_payload = _analytics_extract_training_heat(records)

    if can_choose_office:
        office_chart_payload = _build_office_chart_payload(records)
        labels = office_chart_payload.get('labels') or []
        if not labels:
            office_chart_payload = None
    else:
        office_chart_payload = None

    ctx = {
        'year': year,
        'available_years': available_years,
        'office_id': office_id,
        'available_offices': available_offices,
        'can_choose_office': can_choose_office,
        'scoped_office_name': scoped_office_name,
        'total_records': len(records),
        'met': sum(1 for r in records if r.status == 'MET'),
        'unmet': sum(1 for r in records if r.status == 'UNMET'),
        'variance_payload': variance_payload,
        'area_payload': area_payload,
        'trend_payload': trend_payload,
        'partner_payload': partner_payload,
        'training_payload': training_payload,
        'office_chart_payload': office_chart_payload,
    }
    return render(request, 'core/analytics.html', ctx)


@login_required
def activity_log(request):
    """System activity feed. Superusers see all events; users see own + office events."""
    qs = ActivityLog.objects.select_related('user', 'office').all()
    if not request.user.is_superuser:
        office = _office_for_user(request.user)
        if office:
            qs = qs.filter(Q(user=request.user) | Q(office=office))
        else:
            qs = qs.filter(user=request.user)
    events = list(qs[:200])
    return render(request, 'core/activity_log.html', {'events': events})


@login_required
def performance_viewer_indicator_detail(request, record_id):
    """HTMX fragment: indicator detail card for the performance viewer modal."""
    try:
        selected_qs = PerformanceRecord.objects.select_related(
            'indicator__pap__parent__parent',
            'indicator__pap__office',
        ).filter(pk=record_id)
        if not request.user.is_superuser:
            office = _office_for_user(request.user)
            if office:
                selected_qs = selected_qs.filter(indicator__pap__office=office)
            else:
                selected_qs = selected_qs.none()
        selected_record = selected_qs.get()
    except PerformanceRecord.DoesNotExist:
        selected_record = None
    if selected_record:
        _annotate_record_matrix_columns(selected_record)
    return render(
        request,
        'core/partials/performance_viewer_indicator_detail.html',
        {'selected_record': selected_record},
    )


@user_passes_test(is_admin)
def selective_office_reset(request):
    """Superuser: pick offices by checkbox, preview impact, confirm with DELETE phrase, then remove PAP tree."""
    offices_rows = []
    for office in Office.objects.all().order_by('name'):
        snap = _office_deletion_snapshot(office)
        offices_rows.append({'office': office, **snap})

    offices_meta = {
        str(r['office'].id): {
            'name': r['office'].name,
            'pap_count': r['pap_count'],
            'indicator_count': r['indicator_count'],
            'record_count': r['record_count'],
            'samples': r['pap_samples'][:12],
        }
        for r in offices_rows
    }

    form_error = ''
    if request.method == 'POST':
        raw_ids = request.POST.getlist('office_ids')
        confirm = (request.POST.get('confirm_phrase') or '').strip()
        valid_ids = set(Office.objects.values_list('id', flat=True))
        selected_ids = []
        for x in raw_ids:
            if str(x).isdigit():
                i = int(x)
                if i in valid_ids:
                    selected_ids.append(i)
        selected_ids = list(dict.fromkeys(selected_ids))

        if not selected_ids:
            form_error = 'Select at least one office that still has ingested data.'
            return render(
                request,
                'core/selective_office_reset.html',
                {
                    'offices_rows': offices_rows,
                    'selected_ids': selected_ids,
                    'offices_meta': offices_meta,
                    'form_error': form_error,
                },
            )
        if confirm != 'DELETE':
            form_error = 'Confirmation failed: type the word DELETE (all caps) in the confirmation box.'
            return render(
                request,
                'core/selective_office_reset.html',
                {
                    'offices_rows': offices_rows,
                    'selected_ids': selected_ids,
                    'offices_meta': offices_meta,
                    'form_error': form_error,
                },
            )

        selected_names = []
        for oid in selected_ids:
            office = Office.objects.filter(pk=oid).first()
            if office:
                selected_names.append(office.name)
                StrategicLevel.objects.filter(office=office, level_type='PAP').delete()
        _activity_log(
            request,
            'Selective office reset',
            f'Offices={", ".join(selected_names)}; count={len(selected_ids)}',
            level='warning',
        )
        return redirect('selective_office_reset')

    return render(
        request,
        'core/selective_office_reset.html',
        {
            'offices_rows': offices_rows,
            'selected_ids': [],
            'offices_meta': offices_meta,
            'form_error': form_error,
        },
    )


@user_passes_test(is_admin)
def download_database_backup(request):
    """Let a superuser download a point-in-time SQLite snapshot using the hot-backup API."""
    db = settings.DATABASES['default']
    if db.get('ENGINE') != 'django.db.backends.sqlite3':
        return HttpResponseBadRequest(
            'Full database backup download is only available when using SQLite.',
            content_type='text/plain',
        )

    db_name = db['NAME']
    if hasattr(db_name, '__fspath__'):
        db_name = os.fspath(db_name)

    fd, tmp_path = tempfile.mkstemp(suffix='.sqlite3')
    os.close(fd)
    try:
        source_conn = sqlite3.connect(db_name)
        try:
            dest_conn = sqlite3.connect(tmp_path)
            try:
                source_conn.backup(dest_conn)
            finally:
                dest_conn.close()
        finally:
            source_conn.close()
        with open(tmp_path, 'rb') as bf:
            payload = bf.read()
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass

    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'sopm_backup_{stamp}.sqlite3'
    response = HttpResponse(payload, content_type='application/x-sqlite3')
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    _activity_log(request, 'Database backup downloaded', f'file={filename}', level='info')
    return response


@user_passes_test(is_admin)
def restore_database_backup(request):
    """Superuser: replace the live SQLite database from an uploaded OPMM backup (.sqlite3)."""
    db = settings.DATABASES['default']
    uses_sqlite = db.get('ENGINE') == 'django.db.backends.sqlite3'
    restore_ok = _database_restore_allowed()

    if request.method == 'GET':
        return render(
            request,
            'core/restore_database_backup.html',
            {
                'uses_sqlite': uses_sqlite,
                'restore_allowed': restore_ok,
            },
        )

    if not uses_sqlite:
        messages.error(request, 'Restore is only supported when using SQLite.')
        return redirect('dashboard_home')

    if not restore_ok:
        messages.error(
            request,
            'Database restore is disabled. Set DEBUG or SOPM_ENABLE_DATABASE_RESTORE=1 for this server.',
        )
        return redirect('dashboard_home')

    if request.POST.get('confirm') != '1':
        messages.error(request, 'You must confirm to replace the database.')
        return redirect('restore_database_backup')

    uploaded = request.FILES.get('backup_file')
    if not uploaded:
        messages.error(request, 'Choose a backup file (.sqlite3).')
        return redirect('restore_database_backup')

    resume_username = request.user.get_username()

    db_path = _default_sqlite_db_path()
    db_path.parent.mkdir(parents=True, exist_ok=True)

    tmp_created = None
    pre_saved_name = None
    try:
        fd, tmp_path = tempfile.mkstemp(suffix='.sqlite3', dir=os.fspath(db_path.parent))
        os.close(fd)
        tmp_created = Path(tmp_path)
        with open(tmp_created, 'wb') as out:
            for chunk in uploaded.chunks():
                out.write(chunk)

        ok, err = _validate_sqlite_backup_path(tmp_created)
        if not ok:
            messages.error(request, err)
            return redirect('restore_database_backup')

        backup_dir = settings.BASE_DIR / 'backups'
        backup_dir.mkdir(exist_ok=True)
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        pre_path = backup_dir / f'before_restore_{stamp}.sqlite3'

        connections.close_all()

        if db_path.exists():
            shutil.copy2(db_path, pre_path)
            pre_saved_name = pre_path.name

        os.replace(os.fspath(tmp_created), os.fspath(db_path))
        tmp_created = None

    except OSError as exc:
        messages.error(request, f'Could not replace the database file: {exc}')
        return redirect('restore_database_backup')
    finally:
        if tmp_created is not None and tmp_created.exists():
            try:
                tmp_created.unlink()
            except OSError:
                pass

    connections.close_all()
    UserModel = get_user_model()
    try:
        restored_user = UserModel.objects.get(**{UserModel.USERNAME_FIELD: resume_username})
    except UserModel.DoesNotExist:
        logout(request)
        messages.warning(
            request,
            'Database restored, but this backup has no account with your current username. Please sign in.',
        )
        if pre_saved_name:
            messages.info(request, f'Previous database saved as backups/{pre_saved_name}.')
        return redirect('login')

    if not restored_user.is_active:
        logout(request)
        messages.warning(
            request,
            'Database restored, but that username is inactive in the backup. Please sign in as another user.',
        )
        if pre_saved_name:
            messages.info(request, f'Previous database saved as backups/{pre_saved_name}.')
        return redirect('login')

    login(request, restored_user, backend='django.contrib.auth.backends.ModelBackend')

    if pre_saved_name:
        messages.success(
            request,
            'Database restored. Previous database saved as backups/'
            f'{pre_saved_name}. If you run multiple app processes, restart them so every worker loads the new file.',
        )
    else:
        messages.success(
            request,
            'Database restored. If you run multiple app processes, restart them so every worker loads the new file.',
        )
    return redirect('dashboard_home')


@user_passes_test(is_admin)
@require_POST
def clear_all_data(request):
    if not _full_database_clear_allowed():
        _activity_log(
            request,
            'Clear all data blocked',
            'Attempt blocked because SOPM_ENABLE_FULL_DATABASE_CLEAR is disabled.',
            level='error',
        )
        return redirect('dashboard_home')
    PerformanceRecord.objects.all().delete()
    Indicator.objects.all().delete()
    StrategicLevel.objects.all().delete()
    _activity_log(request, 'Clear all data', 'All strategic performance data removed.', level='warning')
    return redirect('dashboard_home')


@login_required
@require_POST
def clear_office_data(request):
    """Remove PAPs, indicators, and performance rows for the signed-in office only."""
    if request.user.is_superuser:
        _activity_log(
            request,
            'Clear office data skipped',
            'Superuser attempted office-only clear via staff endpoint.',
            level='info',
        )
        return redirect('dashboard_home')
    office = _office_for_user(request.user)
    if not office:
        _activity_log(request, 'Clear office data failed', 'No office linked to account.', level='error')
        return redirect('dashboard_home')
    StrategicLevel.objects.filter(office=office, level_type='PAP').delete()
    _activity_log(
        request,
        'Clear office data',
        f'Cleared ingested data for {office.name}.',
        level='warning',
        office=office,
    )
    return redirect('dashboard_home')


@user_passes_test(is_admin)
def user_management(request):
    if request.method == 'POST':
        full_name = request.POST.get('full_name')
        abbreviation = request.POST.get('abbreviation')
        password = request.POST.get('password')

        if full_name and abbreviation and password:
            if not User.objects.filter(username=abbreviation).exists():
                User.objects.create_user(
                    username=abbreviation,
                    password=password,
                    first_name=full_name,
                )
                office_obj, _ = Office.objects.get_or_create(name=full_name)
                _activity_log(
                    request,
                    'User created',
                    f'Created office user {abbreviation} ({full_name}).',
                    level='info',
                    office=office_obj,
                )
                messages.success(request, f"Office {abbreviation} registered successfully.")
                return redirect('user_management')
            else:
                _activity_log(
                    request,
                    'User create failed',
                    f'Abbreviation already taken: {abbreviation}.',
                    level='error',
                )
                messages.error(request, "Abbreviation already taken.")

    staff_users = User.objects.filter(is_superuser=False)
    return render(request, 'core/user_management.html', {'staff_users': staff_users})


@user_passes_test(is_admin)
def edit_user(request, user_id):
    staff = get_object_or_404(User, pk=user_id)
    if staff.is_superuser:
        messages.error(request, 'Superuser accounts cannot be edited from this page.')
        return redirect('user_management')

    if request.method == 'POST':
        full_name = (request.POST.get('full_name') or '').strip()
        abbreviation = (request.POST.get('abbreviation') or '').strip()
        password = (request.POST.get('password') or '').strip()

        if not full_name or not abbreviation:
            messages.error(request, 'Full name and abbreviation are required.')
            return render(request, 'core/user_edit.html', {'staff': staff})

        if User.objects.filter(username=abbreviation).exclude(pk=staff.pk).exists():
            messages.error(request, 'That abbreviation is already in use.')
            return render(request, 'core/user_edit.html', {'staff': staff})

        old_office_name = (staff.first_name or '').strip()
        linked_office = Office.objects.filter(name=old_office_name).first()
        name_taken = Office.objects.filter(name=full_name)
        if linked_office:
            name_taken = name_taken.exclude(pk=linked_office.pk)
        if name_taken.exists():
            messages.error(
                request,
                'Another office already uses that full name. Choose a different name.',
            )
            return render(request, 'core/user_edit.html', {'staff': staff})

        staff.first_name = full_name
        staff.username = abbreviation
        if password:
            staff.set_password(password)
        staff.save()

        if linked_office:
            linked_office.name = full_name
            linked_office.code = abbreviation[:50] if abbreviation else linked_office.code
            linked_office.save(update_fields=['name', 'code'])
            office_obj = linked_office
        else:
            office_obj, _ = Office.objects.get_or_create(
                name=full_name,
                defaults={'code': (abbreviation[:50] if abbreviation else None)},
            )
        _activity_log(
            request,
            'User updated',
            f'Updated account to {abbreviation} ({full_name}).',
            level='info',
            office=office_obj,
        )

        messages.success(request, f'Updated account for {abbreviation}.')
        return redirect('user_management')

    return render(request, 'core/user_edit.html', {'staff': staff})


@user_passes_test(is_admin)
def delete_user(request, user_id):
    """Handle office deletion."""
    staff = get_object_or_404(User, id=user_id)
    username = staff.username
    office_obj = Office.objects.filter(name=(staff.first_name or '').strip()).first()
    _activity_log(
        request,
        'User deleted',
        f'Deleted account for {username}.',
        level='warning',
        office=office_obj,
    )
    staff.delete()
    messages.warning(request, f"Account for {username} has been deleted.")
    return redirect('user_management')


@login_required
@require_POST
def announcement_mark_read(request):
    """Record that this user has seen one or more announcements (nav bell unread count)."""
    raw_ids = (request.POST.get('ids') or '').strip()
    visible = Announcement.objects.visible_for(request.user)
    if raw_ids:
        pks = []
        for part in raw_ids.split(','):
            part = part.strip()
            if part.isdigit():
                pks.append(int(part))
        if not pks:
            return JsonResponse({'ok': True, 'marked': 0})
        visible = visible.filter(pk__in=pks)
    pks = list(visible.values_list('pk', flat=True))
    if not pks:
        return JsonResponse({'ok': True, 'marked': 0})
    existing = set(
        AnnouncementRead.objects.filter(user=request.user, announcement_id__in=pks).values_list(
            'announcement_id', flat=True
        )
    )
    new_rows = [
        AnnouncementRead(user_id=request.user.pk, announcement_id=pk) for pk in pks if pk not in existing
    ]
    AnnouncementRead.objects.bulk_create(new_rows, batch_size=200, ignore_conflicts=True)
    return JsonResponse({'ok': True, 'marked': len(new_rows)})


@user_passes_test(is_admin)
def announcement_manage(request):
    """Create and toggle announcements (global or office-targeted)."""
    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        if action == 'create':
            title = (request.POST.get('title') or '').strip()
            body = (request.POST.get('body') or '').strip()
            scope = (request.POST.get('scope') or Announcement.SCOPE_GLOBAL).strip()
            if scope not in (Announcement.SCOPE_GLOBAL, Announcement.SCOPE_OFFICES):
                scope = Announcement.SCOPE_GLOBAL
            if not title or not body:
                messages.error(request, 'Title and message are required.')
            else:
                ann = Announcement.objects.create(
                    title=title[:200],
                    body=body,
                    scope=scope,
                    is_active=True,
                )
                if scope == Announcement.SCOPE_OFFICES:
                    raw_ids = request.POST.getlist('office_ids')
                    pks = []
                    for x in raw_ids:
                        if str(x).isdigit():
                            pks.append(int(x))
                    ann.offices.set(Office.objects.filter(pk__in=pks))
                    if not ann.offices.exists():
                        messages.warning(
                            request,
                            'Posted as office-specific but no offices were selected — nobody will see it until you edit it in admin or post again with offices checked.',
                        )
                _activity_log(
                    request,
                    'Announcement posted',
                    f'{title[:80]} ({scope})',
                    level='info',
                )
                messages.success(request, 'Announcement is live. Users will see it in Help center → Announcements and in the bell drawer.')
            return redirect('announcement_manage')
        if action == 'toggle':
            raw_id = (request.POST.get('id') or '').strip()
            if raw_id.isdigit():
                ann = get_object_or_404(Announcement, pk=int(raw_id))
                ann.is_active = not ann.is_active
                ann.save()
                state = 'activated' if ann.is_active else 'archived'
                _activity_log(request, f'Announcement {state}', ann.title[:120], level='info')
                messages.success(request, f'Announcement {state}.')
            return redirect('announcement_manage')
        return redirect('announcement_manage')

    items = Announcement.objects.all().order_by('-created_at')[:80]
    offices = Office.objects.order_by('name')
    return render(
        request,
        'core/announcement_manage.html',
        {'items': items, 'offices': offices},
    )
