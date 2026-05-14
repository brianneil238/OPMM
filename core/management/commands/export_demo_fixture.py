"""Write core/fixtures/sample_strategic_data.json (Demo — offices and related rows only)."""

from pathlib import Path

from django.core import serializers
from django.core.management.base import BaseCommand

from core.management.commands.seed_sample_data import DEMO_PREFIX
from core.models import Indicator, Office, PerformanceRecord, StrategicLevel


class Command(BaseCommand):
    help = (
        'Export Demo-prefix offices and related strategic rows to '
        'core/fixtures/sample_strategic_data.json'
    )

    def handle(self, *args, **options):
        offices = list(Office.objects.filter(name__startswith=DEMO_PREFIX).order_by('pk'))
        if not offices:
            self.stderr.write(
                'No Demo — offices found. Run: python manage.py seed_sample_data --reset'
            )
            return

        office_ids = {o.id for o in offices}
        paps = list(
            StrategicLevel.objects.filter(
                level_type='PAP', office_id__in=office_ids
            ).order_by('pk')
        )
        sl_ids = set()
        for pap in paps:
            sl_ids.add(pap.pk)
            st = pap.parent
            if st:
                sl_ids.add(st.pk)
                out = st.parent
                if out:
                    sl_ids.add(out.pk)

        levels = list(StrategicLevel.objects.filter(pk__in=sl_ids).order_by('pk'))
        # loaddata: parents before children
        outcomes = [x for x in levels if x.level_type == 'OUTCOME']
        strategies = [x for x in levels if x.level_type == 'STRATEGY']
        pap_levels = [x for x in levels if x.level_type == 'PAP']

        pap_ids = [p.pk for p in paps]
        indicators = list(
            Indicator.objects.filter(pap_id__in=pap_ids).order_by('pk')
        )
        ind_ids = [i.pk for i in indicators]
        records = list(
            PerformanceRecord.objects.filter(indicator_id__in=ind_ids).order_by('pk')
        )

        combined = offices + outcomes + strategies + pap_levels + indicators + records
        data = serializers.serialize('json', combined, indent=2)

        out_path = Path(__file__).resolve().parents[2] / 'fixtures' / 'sample_strategic_data.json'
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(data + '\n', encoding='utf-8')

        self.stdout.write(
            self.style.SUCCESS(
                f'Wrote {out_path} ({len(offices)} offices, {len(levels)} strategic levels, '
                f'{len(indicators)} indicators, {len(records)} performance rows).'
            )
        )
