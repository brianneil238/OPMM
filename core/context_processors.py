"""Template context shared across authenticated pages."""

from core.models import Announcement, AnnouncementRead


def nav_announcements(request):
    """Active announcements visible to this user (for bell + slide-in panel)."""
    user = getattr(request, 'user', None)
    if not user or not user.is_authenticated:
        return {'nav_announcement_list': [], 'nav_announcement_count': 0}
    qs = list(Announcement.objects.visible_for(user).prefetch_related('offices')[:25])
    if not qs:
        return {'nav_announcement_list': [], 'nav_announcement_count': 0}
    ids = [a.pk for a in qs]
    read_ids = set(
        AnnouncementRead.objects.filter(user=user, announcement_id__in=ids).values_list(
            'announcement_id', flat=True
        )
    )
    unread = sum(1 for a in qs if a.pk not in read_ids)
    return {'nav_announcement_list': qs, 'nav_announcement_count': unread}
