"""
Load sample offices, PAPs, indicators, and performance rows for local testing.

Usage:
  python manage.py seed_sample_data --reset
  python manage.py seed_sample_data --reset --with-users
  python manage.py seed_sample_data --reset --stress
  python manage.py seed_sample_data --reset --bulk-indicators 100 --bulk-quarters 4
  python manage.py seed_sample_data --reset --balanced-only --balanced-per-area 100
      # exactly six areas x 100 indicators, ~80% met (last area one fewer met)
  python manage.py seed_sample_data --reset --multi-office-years-demo
      # 3 Demo offices, 50 KPIs across six development areas, Q1–Q4 for 2024 and 2025

Demo staff password (with --with-users): demo12345

Bulk options add synthetic KPIs (split across the three demo offices) for pagination / query stress tests.
"""

from django.contrib.auth.models import User
from django.core.management.base import BaseCommand, CommandError
from django.db import transaction

from core.models import Indicator, Office, PerformanceRecord, StrategicLevel

DEMO_PREFIX = 'Demo — '
DEMO_PASSWORD = 'demo12345'

DEMO_USERNAMES = (
    ('demo.central', f'{DEMO_PREFIX}Central Programs'),
    ('demo.east', f'{DEMO_PREFIX}East Regional Office'),
    ('demo.west', f'{DEMO_PREFIX}West Research Lab'),
)


def _clear_strategic_data():
    PerformanceRecord.objects.all().delete()
    Indicator.objects.all().delete()
    StrategicLevel.objects.all().delete()


def _remove_demo_offices_and_users():
    Office.objects.filter(name__startswith=DEMO_PREFIX).delete()
    User.objects.filter(username__in=[u for u, _ in DEMO_USERNAMES]).delete()


def _ensure_pap(office, outcome_name, strategy_name, pap_name):
    out, _ = StrategicLevel.objects.get_or_create(
        name=outcome_name,
        level_type='OUTCOME',
        defaults={'parent': None, 'office': None},
    )
    st, _ = StrategicLevel.objects.get_or_create(
        name=strategy_name,
        level_type='STRATEGY',
        defaults={'parent': out, 'office': None},
    )
    if st.parent_id != out.id:
        st.parent = out
        st.save(update_fields=['parent'])
    pap, _ = StrategicLevel.objects.get_or_create(
        name=pap_name,
        level_type='PAP',
        parent=st,
        office=office,
    )
    return pap


def _add_indicator_records(pap, description, target_value, rows):
    """rows: list of (year, quarter, actual_value|None, raw_text)."""
    ind, _ = Indicator.objects.get_or_create(
        pap=pap,
        description=description,
        defaults={'target_text': ''},
    )
    for year, quarter, actual, raw in rows:
        PerformanceRecord.objects.update_or_create(
            indicator=ind,
            quarter=quarter,
            year=year,
            defaults={
                'target_value': target_value,
                'actual_value': actual,
                'raw_actual_text': raw or '',
            },
        )


def _split_counts(total, n_parts):
    base = total // n_parts
    rem = total % n_parts
    return [base + (1 if i < rem else 0) for i in range(n_parts)]


def _bulk_actual_for_cell(global_idx, quarter, target):
    """Vary MET / UNMET / PENDING across rows and quarters."""
    r = (global_idx + quarter) % 11
    if r == 0:
        return None
    if r in (1, 2, 3, 4, 5):
        return float(target) + 2.0 + (quarter % 2)
    return max(0.0, float(target) - 5.0 - quarter)


def bulk_append_synthetic(offices, bulk_indicators, bulk_quarters, command):
    """
    offices: list of 3 Office instances (Central, East, West).
    Creates one bulk PAP per office and bulk_indicators total indicators split evenly.
    """
    if bulk_indicators < 1 or bulk_quarters < 1:
        return
    bulk_quarters = min(4, max(1, bulk_quarters))
    counts = _split_counts(bulk_indicators, len(offices))
    year = 2026
    global_idx = 0

    for office, cnt in zip(offices, counts):
        if cnt < 1:
            continue
        pap = _ensure_pap(
            office,
            f'Bulk stress outcome ({office.name})',
            'Bulk synthetic KPI stream 2026',
            f'Bulk synthetic PAP ({office.name})',
        )
        inds = []
        seq_keys = []
        for _ in range(cnt):
            seq_keys.append(global_idx)
            inds.append(
                Indicator(
                    pap=pap,
                    description=f'BULK Synthetic KPI #{global_idx} (target 100)',
                    target_text='',
                )
            )
            global_idx += 1
        Indicator.objects.bulk_create(inds)

        recs = []
        for ind, g in zip(inds, seq_keys):
            for q in range(1, bulk_quarters + 1):
                tgt = 100.0
                act = _bulk_actual_for_cell(g, q, tgt)
                recs.append(
                    PerformanceRecord(
                        indicator=ind,
                        quarter=q,
                        year=year,
                        target_value=tgt,
                        actual_value=act,
                        raw_actual_text='' if act is None else f'Q{q} bulk',
                    )
                )
        PerformanceRecord.objects.bulk_create(recs, batch_size=2000)

    command.stdout.write(
        command.style.SUCCESS(
            f'Bulk block: {bulk_indicators} synthetic indicators x {bulk_quarters} quarter(s) ({year}).'
        )
    )


# Outcome names must compact-match viewer needles (see core.views.VIEWER_DEVELOPMENT_AREAS).
_EVEN_PORTFOLIO_AREAS = (
    (
        'Academic Leadership — even portfolio',
        'Even strategy AL',
        'Even PAP AL',
    ),
    (
        'Research and Innovation — even portfolio',
        'Even strategy RNI',
        'Even PAP RNI',
    ),
    (
        'Social Responsibility — even portfolio',
        'Even strategy SR',
        'Even PAP SR',
    ),
    (
        'Internationalization — even portfolio',
        'Even strategy INT',
        'Even PAP INT',
    ),
    (
        'Advancing Interdisciplinarity — even portfolio',
        'Even strategy ID',
        'Even PAP ID',
    ),
    (
        'Sustainability — even portfolio',
        'Even strategy SUS',
        'Even PAP SUS',
    ),
)


def append_balanced_dev_area_rows(
    offices, per_area, met_pct, year, quarter, command
):
    """Create per_area indicators (and one performance row each) per development area.

    Rows map to the six OPMM development areas via OUTCOME names. Met share is
    approximately met_pct; the last area is one met row short so bars are not
    perfectly identical.
    """
    if per_area < 1 or not offices:
        return
    if not (1 <= met_pct <= 99):
        raise CommandError('--balanced-met-pct must be between 1 and 99.')
    met_base = (per_area * met_pct + 50) // 100
    target = 100.0

    for ai, (oname, sname, pname) in enumerate(_EVEN_PORTFOLIO_AREAS):
        met_n = met_base if ai < len(_EVEN_PORTFOLIO_AREAS) - 1 else max(0, met_base - 1)
        office = offices[ai % len(offices)]
        pap = _ensure_pap(office, oname, sname, pname)
        inds = []
        for i in range(per_area):
            inds.append(
                Indicator(
                    pap=pap,
                    description=f'Even demo area {ai + 1} KPI #{i + 1}',
                    target_text='',
                )
            )
        Indicator.objects.bulk_create(inds, batch_size=1000)

        recs = []
        for i, ind in enumerate(inds):
            if i < met_n:
                actual = target
            else:
                actual = max(0.0, target - 25.0)
            recs.append(
                PerformanceRecord(
                    indicator=ind,
                    quarter=quarter,
                    year=year,
                    target_value=target,
                    actual_value=actual,
                    raw_actual_text='even-demo',
                )
            )
        PerformanceRecord.objects.bulk_create(recs, batch_size=2000)

    command.stdout.write(
        command.style.SUCCESS(
            f'Even portfolio: {per_area} indicators x {len(_EVEN_PORTFOLIO_AREAS)} areas '
            f'(~{met_pct}% met; last area one fewer met), FY {year} Q{quarter}.'
        )
    )


def _multi_office_year_demo_pairs(offices):
    """18 (office, area_tuple, unique pap suffix) cells: 6 areas × 3 offices."""
    pairs = []
    for oname, sname, pname in _EVEN_PORTFOLIO_AREAS:
        for office in offices:
            code = (office.code or str(office.pk))[:12]
            pap_label = f'{pname} [{code}]'
            pairs.append((office, oname, sname, pap_label))
    return pairs


def append_multi_office_years_demo(offices, command):
    """50 indicators across six development areas and three offices; 8 quarters each (2024–2025).

    Rows map to viewer development-area buckets via OUTCOME names in ``_EVEN_PORTFOLIO_AREAS``.
    Indicators are assigned round-robin across the 18 (area × office) PAP slots so every office
    and every pillar appears in the combined Performance viewer.
    """
    years = (2024, 2025)
    total_indicators = 50
    target = 100.0
    pairs = _multi_office_year_demo_pairs(offices)

    for k in range(total_indicators):
        pair_idx = k % len(pairs)
        office, oname, sname, pap_label = pairs[pair_idx]
        pap = _ensure_pap(office, oname, sname, pap_label)
        short_off = office.name.replace(DEMO_PREFIX, '').strip() or office.name
        desc = (
            f'Multi-year demo KPI #{k + 1}: {short_off} - '
            f'FY {years[0]}-{years[1]} (viewer bucket from outcome title)'
        )
        ind, _ = Indicator.objects.get_or_create(
            pap=pap,
            description=desc,
            defaults={'target_text': str(int(target))},
        )
        for year in years:
            for quarter in (1, 2, 3, 4):
                mix = (k + year * 3 + quarter * 5) % 7
                if mix == 0:
                    actual = None
                    raw = f'FY{year} Q{quarter}: pending narrative'
                elif mix in (1, 2, 3, 4):
                    actual = target + (quarter % 2)
                    raw = f'FY{year} Q{quarter}: target met at {int(actual)}%'
                else:
                    actual = max(0.0, target - 20.0 - quarter)
                    raw = f'FY{year} Q{quarter}: below target ({int(actual)}%)'
                PerformanceRecord.objects.update_or_create(
                    indicator=ind,
                    quarter=quarter,
                    year=year,
                    defaults={
                        'target_value': target,
                        'actual_value': actual,
                        'raw_actual_text': raw,
                    },
                )

    n_ind = Indicator.objects.filter(description__startswith='Multi-year demo KPI #').count()
    n_rec = PerformanceRecord.objects.filter(
        indicator__description__startswith='Multi-year demo KPI #',
        year__in=years,
    ).count()
    command.stdout.write(
        command.style.SUCCESS(
            f'Multi-office years demo: {n_ind} indicators, {n_rec} performance rows '
            f'({years[0]}-{years[1]}, Q1-Q4), across {len(pairs)} area x office PAPs.'
        )
    )


class Command(BaseCommand):
    help = 'Load sample Demo offices and performance data for testing the SOPM app.'

    def add_arguments(self, parser):
        parser.add_argument(
            '--reset',
            action='store_true',
            help='Delete all strategic data (and demo offices/users) before loading.',
        )
        parser.add_argument(
            '--with-users',
            action='store_true',
            help=f'Create demo staff users (password: {DEMO_PASSWORD}) linked to Demo offices.',
        )
        parser.add_argument(
            '--bulk-indicators',
            type=int,
            default=0,
            metavar='N',
            help='After core seed, add N extra synthetic indicators (split across demo offices) for load testing.',
        )
        parser.add_argument(
            '--bulk-quarters',
            type=int,
            default=2,
            metavar='Q',
            help='With --bulk-indicators, create Q performance rows per indicator (quarters 1..Q, year 2026). Default 2, max 4.',
        )
        parser.add_argument(
            '--stress',
            action='store_true',
            help='Shorthand for ~100+ indicators load test: same as --bulk-indicators 100 --bulk-quarters 4.',
        )
        parser.add_argument(
            '--balanced-per-area',
            type=int,
            default=0,
            metavar='N',
            help='Add N indicators per development area (six areas), each with one FY row '
            '(year 2026, quarter from --balanced-quarter), roughly --balanced-met-pct met. '
            'With default seed, this is appended after curated data; with --balanced-only, '
            'this is the only strategic load (exactly 6×N indicators).',
        )
        parser.add_argument(
            '--balanced-met-pct',
            type=int,
            default=80,
            metavar='P',
            help='With --balanced-per-area, target met %% (default 80). Last area uses one fewer met row.',
        )
        parser.add_argument(
            '--balanced-quarter',
            type=int,
            default=1,
            choices=(1, 2, 3, 4),
            help='Quarter for balanced demo rows (default 1).',
        )
        parser.add_argument(
            '--balanced-only',
            action='store_true',
            help='With --reset, skip curated scenarios and bulk; load only three demo offices '
            'and --balanced-per-area rows (exactly six areas × N indicators).',
        )
        parser.add_argument(
            '--multi-office-years-demo',
            action='store_true',
            help='With --reset, skip curated/bulk/balanced loads; create only three Demo offices '
            'with 50 indicators across six development areas and all three offices, '
            'each with Q1–Q4 rows for 2024 and 2025 (multi-office / multi-year viewer test).',
        )

    def handle(self, *args, **options):
        reset = options['reset']
        with_users = options['with_users']
        bulk_n = max(0, options['bulk_indicators'] or 0)
        bulk_q = max(1, min(4, int(options['bulk_quarters'] or 2)))
        if options.get('stress'):
            bulk_n = 100
            bulk_q = 4

        balanced_n = max(0, int(options.get('balanced_per_area') or 0))
        balanced_met = int(options.get('balanced_met_pct') or 80)
        balanced_q = int(options.get('balanced_quarter') or 1)
        balanced_only = bool(options.get('balanced_only'))
        multi_years_demo = bool(options.get('multi_office_years_demo'))
        if balanced_n > 10_000:
            raise CommandError('--balanced-per-area is capped at 10000 for safety.')
        if balanced_only and not reset:
            raise CommandError('--balanced-only requires --reset.')
        if balanced_only and balanced_n < 1:
            raise CommandError('--balanced-only requires --balanced-per-area >= 1.')
        if balanced_only and bulk_n > 0:
            raise CommandError(
                '--balanced-only cannot be combined with --bulk-indicators or --stress.'
            )
        if multi_years_demo and not reset:
            raise CommandError('--multi-office-years-demo requires --reset.')
        if multi_years_demo and (balanced_only or bulk_n > 0 or balanced_n > 0):
            raise CommandError(
                '--multi-office-years-demo cannot be combined with --balanced-only, '
                '--bulk-indicators, --stress, or --balanced-per-area.'
            )

        if bulk_n > 50_000:
            raise CommandError('--bulk-indicators is capped at 50000 for safety.')

        if StrategicLevel.objects.exists() and not reset:
            raise CommandError(
                'This database already has strategic data. Re-run with --reset to replace '
                'everything with sample data (destructive), or use an empty database.'
            )

        with transaction.atomic():
            if reset:
                self.stdout.write('Clearing strategic tables...')
                _clear_strategic_data()
                self.stdout.write('Removing prior Demo offices and demo logins...')
                _remove_demo_offices_and_users()

            o1, _ = Office.objects.get_or_create(
                name=f'{DEMO_PREFIX}Central Programs',
                defaults={'code': 'DCP'},
            )
            o2, _ = Office.objects.get_or_create(
                name=f'{DEMO_PREFIX}East Regional Office',
                defaults={'code': 'DRE'},
            )
            o3, _ = Office.objects.get_or_create(
                name=f'{DEMO_PREFIX}West Research Lab',
                defaults={'code': 'DWL'},
            )

            if multi_years_demo:
                append_multi_office_years_demo([o1, o2, o3], self)
            elif balanced_only:
                append_balanced_dev_area_rows(
                    [o1, o2, o3],
                    balanced_n,
                    balanced_met,
                    2026,
                    balanced_q,
                    self,
                )
            else:
                # --- Demo — Central Programs
                pap = _ensure_pap(
                    o1,
                    'Academic Leadership and student success',
                    'Teaching excellence 2026',
                    'PAP Graduate outcomes dashboard',
                )
                _add_indicator_records(
                    pap,
                    'Graduation rate target (85)',
                    85.0,
                    [
                        (2026, 1, 88.0, '88%'),
                        (2026, 2, 82.0, 'Q2 dip'),
                        (2026, 3, None, ''),
                        (2026, 4, 90.0, 'Recovery'),
                        (2025, 4, 80.0, 'Prior year'),
                    ],
                )
                _add_indicator_records(
                    pap,
                    'First-year retention (90)',
                    90.0,
                    [(2026, 1, 92.0, ''), (2026, 2, 89.0, ''), (2026, 3, 91.0, '')],
                )

                pap = _ensure_pap(
                    o1,
                    'Research and Innovation acceleration',
                    'Grants and partnerships',
                    'PAP External research funding',
                )
                _add_indicator_records(
                    pap,
                    'New awards count (12)',
                    12.0,
                    [(2026, 1, 15.0, '15 awards'), (2026, 2, 10.0, '')],
                )
                _add_indicator_records(
                    pap,
                    'Industry partnerships (8)',
                    8.0,
                    [(2026, 1, 5.0, 'Behind'), (2026, 2, 8.0, 'On track')],
                )

                pap = _ensure_pap(
                    o1,
                    'Social responsibility community impact',
                    'Outreach programs',
                    'PAP Service learning placements',
                )
                _add_indicator_records(
                    pap,
                    'Placements completed (200)',
                    200.0,
                    [(2026, 1, 195.0, ''), (2026, 2, 205.0, '')],
                )

                # --- Demo — East Regional Office
                pap = _ensure_pap(
                    o2,
                    'Internationalization student mobility',
                    'Exchange and mobility',
                    'PAP Outbound exchange participants',
                )
                _add_indicator_records(
                    pap,
                    'Participants target (60)',
                    60.0,
                    [(2026, 1, 72.0, '72 students'), (2026, 2, 58.0, '')],
                )
                _add_indicator_records(
                    pap,
                    'Partner institutions (10)',
                    10.0,
                    [(2026, 1, 10.0, ''), (2026, 2, 9.0, '')],
                )

                pap = _ensure_pap(
                    o2,
                    'Advancing interdisciplinarity cross-college',
                    'Joint programs',
                    'PAP Joint degree enrollments',
                )
                _add_indicator_records(
                    pap,
                    'Joint enrollments (40)',
                    40.0,
                    [(2026, 1, None, 'Pending report'), (2026, 2, 38.0, '')],
                )
                _add_indicator_records(
                    pap,
                    'Cross-listed courses (25)',
                    25.0,
                    [(2026, 1, 30.0, '30 courses')],
                )

                pap = _ensure_pap(
                    o2,
                    'Sustainability campus operations',
                    'Carbon reduction plan',
                    'PAP Energy use reduction',
                )
                _add_indicator_records(
                    pap,
                    'Reduction vs baseline (15)',
                    15.0,
                    [(2026, 1, 12.0, '12%'), (2026, 2, 16.0, '')],
                )

                # --- Demo — West Research Lab
                pap = _ensure_pap(
                    o3,
                    'Research and Innovation infrastructure',
                    'Core facilities',
                    'PAP Instrument uptime',
                )
                _add_indicator_records(
                    pap,
                    'Uptime target (95)',
                    95.0,
                    [(2026, 1, 97.0, ''), (2026, 2, 96.0, '')],
                )
                _add_indicator_records(
                    pap,
                    'User sessions (3000)',
                    3000.0,
                    [(2026, 1, 2800.0, ''), (2026, 2, 3100.0, '')],
                )

                pap = _ensure_pap(
                    o3,
                    'Academic Leadership research mentoring',
                    'Graduate mentoring',
                    'PAP Doctoral completion timeline',
                )
                _add_indicator_records(
                    pap,
                    'On-time completion rate (70)',
                    70.0,
                    [(2026, 1, 75.0, '75%'), (2026, 2, 68.0, '')],
                )

                if bulk_n > 0:
                    bulk_append_synthetic([o1, o2, o3], bulk_n, bulk_q, self)

                if balanced_n > 0:
                    append_balanced_dev_area_rows(
                        [o1, o2, o3],
                        balanced_n,
                        balanced_met,
                        2026,
                        balanced_q,
                        self,
                    )

            if with_users:
                for username, full_name in DEMO_USERNAMES:
                    if User.objects.filter(username=username).exists():
                        u = User.objects.get(username=username)
                        created = False
                    else:
                        u = User.objects.create_user(
                            username=username,
                            password=DEMO_PASSWORD,
                            first_name=full_name,
                        )
                        created = True
                    u.first_name = full_name
                    u.is_superuser = False
                    u.is_staff = False
                    u.set_password(DEMO_PASSWORD)
                    u.save()
                    action = 'Created' if created else 'Updated'
                    self.stdout.write(self.style.SUCCESS(f'{action} user {username}'))

        n_pap = StrategicLevel.objects.filter(level_type='PAP').count()
        n_ind = Indicator.objects.count()
        n_rec = PerformanceRecord.objects.count()
        self.stdout.write(self.style.SUCCESS(
            f'Sample data ready: 3 demo offices, {n_pap} PAPs, {n_ind} indicators, {n_rec} performance rows.'
        ))
        self.stdout.write(
            'Use a superuser for Performance viewer / charts. '
            f'With --with-users, sign in as demo.central / demo.east / demo.west (password {DEMO_PASSWORD}).'
        )
