from django import template

from core.models import Announcement

register = template.Library()


@register.inclusion_tag('core/partials/announcements_panel.html', takes_context=True)
def announcements_panel(context, limit=8):
    request = context.get('request')
    user = getattr(request, 'user', None) if request else None
    if not user or not user.is_authenticated:
        return {'announcements': []}
    lim = int(limit) if limit else 8
    qs = Announcement.objects.visible_for(user).prefetch_related('offices')[:lim]
    return {'announcements': list(qs)}
