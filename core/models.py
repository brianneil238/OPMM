from django.db import models
import re
from django.contrib.auth.models import User

class Office(models.Model):
    name = models.CharField(max_length=255, unique=True)
    code = models.CharField(max_length=50, blank=True, null=True)

    def __str__(self):
        return self.name

class StrategicLevel(models.Model):
    name = models.TextField()
    level_type = models.CharField(max_length=50) # OUTCOME, STRATEGY, PAP
    parent = models.ForeignKey('self', on_delete=models.CASCADE, null=True, blank=True, related_name='children')
    office = models.ForeignKey(Office, on_delete=models.SET_NULL, null=True)

    def __str__(self):
        return f"{self.level_type}: {self.name[:30]}"

class Indicator(models.Model):
    pap = models.ForeignKey(StrategicLevel, on_delete=models.CASCADE, related_name='indicators')
    description = models.TextField()
    target_text = models.CharField(max_length=255, blank=True)

    def __str__(self):
        return self.description[:50]

class PerformanceRecord(models.Model):
    indicator = models.ForeignKey(Indicator, on_delete=models.CASCADE, related_name='records')
    quarter = models.IntegerField()
    year = models.IntegerField(default=2026) 
    target_value = models.FloatField(blank=True, null=True)
    actual_value = models.FloatField(null=True, blank=True)
    raw_actual_text = models.TextField(blank=True)
    # Raw matrix "Variance" column (e.g. MET, UNMET, +1, -1); used for status when explicit.
    variance_text = models.TextField(blank=True)
    # When set (e.g. from LPC "Status of Accomplishment" column), overrides numeric MET/UNMET.
    explicit_status = models.CharField(max_length=8, blank=True, null=True)

    def _infer_float_from_text(self, text):
        """Best-effort parse of a numeric value from free-form text.

        We intentionally avoid treating years (e.g. 2026) as values.
        """
        if not text:
            return None
        s = str(text)
        # Prefer percentage-like tokens first (e.g. "83%", "83.5 %").
        m = re.search(r'(?<!\d)(\d{1,3}(?:\.\d+)?)(?:\s*%)(?!\w)', s)
        if m:
            return float(m.group(1))
        # Fallback: first reasonable number token that is not a year.
        m = re.search(r'(?<!\d)(\d+(?:\.\d+)?)(?!\d)', s)
        if not m:
            return None
        token = m.group(1)
        if re.fullmatch(r'20\d{2}', token):
            return None
        return float(token)

    def _infer_status_from_text(self, text):
        """Infer MET/UNMET from common accomplishment phrases."""
        if not text:
            return None
        s = str(text).strip().lower()
        # Pure placeholders mean "no meaningful signal"; caller may fall back.
        if re.fullmatch(r'(n\s*/\s*a|na|none|nil)\.?\s*', s):
            return None
        # Strong UNMET signals first (avoid "met" matching inside "unmet").
        if re.search(r'\bunmet\b', s):
            return 'UNMET'
        if re.search(r'\bnot\s+met\b', s) or re.search(r'\bnot\s+achiev', s):
            return 'UNMET'
        # Common "nothing happened" phrasing.
        if re.search(r'\bno\s+(meeting|meetings)\b', s):
            return 'UNMET'
        if re.search(r'\bno\b', s) and re.search(r'\b(conducted|submitted|accomplished|achieved|completed|done)\b', s):
            return 'UNMET'
        if re.search(r'\bno\b', s) and re.search(r'\b(accomplish|achiev|complete|submit)\w*\b', s):
            return 'UNMET'
        # Strong MET signals.
        if re.search(r'\bmet\b', s) and re.search(r'\b(target|goal|requirement|kpi)\b', s):
            return 'MET'
        if re.search(
            r'\b(?:fully\s+accomplish(?:ed)?|fully\s+met|substantially\s+met|'
            r'target\s+exceeded|beyond\s+target|exceeded\s+target|'
            r'100\s*%\s+accomplish|completed\s+ahead|ahead\s+of\s+target)\b',
            s,
        ):
            return 'MET'
        if re.search(r'\b(achiev|accomplish|complete|completed|submit|submitted|done)\w*\b', s):
            return 'MET'
        if re.search(r'\b(attend|attended|conduct|conducted|issued|coordinated|monitored|evaluated|held)\w*\b', s):
            return 'MET'
        return None

    @property
    def status(self):
        if self.explicit_status in ('MET', 'UNMET'):
            return self.explicit_status
        vt = (self.variance_text or '').strip()
        if vt:
            from .services import _parse_explicit_met_unmet_cell

            parsed = _parse_explicit_met_unmet_cell(vt)
            if parsed in ('MET', 'UNMET'):
                return parsed
        # Primary: numeric scoring.
        actual = self.actual_value
        target = self.target_value

        # If ingestion didn't populate numeric values, infer them from stored text.
        if actual is None:
            actual = self._infer_float_from_text(self.raw_actual_text)
        if target is None:
            target = self._infer_float_from_text(getattr(self.indicator, 'target_text', '') or '') or self._infer_float_from_text(
                getattr(self.indicator, 'description', '') or ''
            )

        if actual is not None and target is not None:
            if actual >= target:
                return "MET"
            raw = self.raw_actual_text or ''
            # Narrative can carry an explicit MET/UNMET label while the scraped number is wrong.
            # Do not use broad "achieved" heuristics here (e.g. "88% achieved" can still be UNMET vs 100).
            if re.search(r'(?i)\bunmet\b', raw):
                return "UNMET"
            if re.search(r'(?i)\bmet\b', raw):
                return "MET"
            if re.search(
                r'(?i)\b(?:fully\s+accomplish(?:ed)?|fully\s+met|substantially\s+met|'
                r'target\s+exceeded|beyond\s+target|exceeded\s+target)\b',
                raw,
            ):
                return "MET"
            return "UNMET"

        # Secondary: textual status inference (e.g. "MET/UNMET", "achieved", "not met").
        inferred = self._infer_status_from_text(self.raw_actual_text)
        if inferred in ('MET', 'UNMET'):
            return inferred
        # Last resort: if there is any non-empty text, treat it as progress (MET),
        # unless it was an explicit placeholder (handled above) or an UNMET signal.
        if (self.raw_actual_text or '').strip():
            return "MET"
        return "PENDING"

    class Meta:
        # Crucial: Allows multiple quarters for the same indicator
        unique_together = ('indicator', 'quarter', 'year')


class AnnouncementQuerySet(models.QuerySet):
    def visible_for(self, user):
        """Active rows: everyone sees ``global``; office users also see targeted office posts."""
        from django.db.models import Q

        if not user or not getattr(user, 'is_authenticated', False):
            return self.none()
        qs = self.filter(is_active=True)
        if getattr(user, 'is_superuser', False):
            return qs.order_by('sort_order', '-created_at', '-id')
        office = Office.objects.filter(name=user.first_name).first()
        if office:
            return (
                qs.filter(Q(scope=Announcement.SCOPE_GLOBAL) | Q(scope=Announcement.SCOPE_OFFICES, offices__pk=office.pk))
                .distinct()
                .order_by('sort_order', '-created_at', '-id')
            )
        return qs.filter(scope=Announcement.SCOPE_GLOBAL).order_by('sort_order', '-created_at', '-id')


class AnnouncementManager(models.Manager):
    def get_queryset(self):
        return AnnouncementQuerySet(self.model, using=self._db)

    def visible_for(self, user):
        return self.get_queryset().visible_for(user)


class Announcement(models.Model):
    """Short-lived notices shown in Help center / bell drawer (global or selected offices)."""

    SCOPE_GLOBAL = 'global'
    SCOPE_OFFICES = 'offices'
    SCOPE_CHOICES = [
        (SCOPE_GLOBAL, 'All offices'),
        (SCOPE_OFFICES, 'Selected offices only'),
    ]

    title = models.CharField(max_length=200)
    body = models.TextField()
    is_active = models.BooleanField(default=True)
    scope = models.CharField(max_length=20, choices=SCOPE_CHOICES, default=SCOPE_GLOBAL)
    offices = models.ManyToManyField(Office, blank=True, related_name='announcements')
    sort_order = models.PositiveSmallIntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    objects = AnnouncementManager()

    class Meta:
        ordering = ['sort_order', '-created_at', '-id']

    def __str__(self):
        return self.title[:80]


class AnnouncementRead(models.Model):
    """Per-user acknowledgement so the nav bell badge reflects unread notices only."""

    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='announcement_reads')
    announcement = models.ForeignKey(Announcement, on_delete=models.CASCADE, related_name='reads')
    read_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(fields=('user', 'announcement'), name='uniq_announcementread_user_announcement'),
        ]
        ordering = ['-read_at', '-id']

    def __str__(self):
        return f'{self.user_id}:{self.announcement_id}'


class ActivityLog(models.Model):
    """System activity trail for uploads, resets, and admin operations."""
    created_at = models.DateTimeField(auto_now_add=True)
    level = models.CharField(max_length=16, default='info')
    action = models.CharField(max_length=120)
    detail = models.TextField(blank=True)
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    office = models.ForeignKey(Office, on_delete=models.SET_NULL, null=True, blank=True)

    class Meta:
        ordering = ['-created_at', '-id']