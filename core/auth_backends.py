"""Authentication helpers for OPMM."""

from django.contrib.auth import get_user_model
from django.contrib.auth.backends import ModelBackend


class CaseInsensitiveUsernameBackend(ModelBackend):
    """Allow office logins regardless of username letter case (e.g. OVCRDES vs ovcrdes)."""

    def authenticate(self, request, username=None, password=None, **kwargs):
        UserModel = get_user_model()
        if username is None:
            username = kwargs.get(UserModel.USERNAME_FIELD)
        if username is None or password is None:
            return None
        username = str(username).strip()
        if not username:
            return None
        try:
            user = UserModel._default_manager.get(**{f'{UserModel.USERNAME_FIELD}__iexact': username})
        except UserModel.DoesNotExist:
            return None
        if user.check_password(password) and self.user_can_authenticate(user):
            return user
        return None
