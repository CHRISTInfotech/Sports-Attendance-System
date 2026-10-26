import logging

from allauth.core.exceptions import ImmediateHttpResponse
from allauth.socialaccount.adapter import DefaultSocialAccountAdapter
from django.contrib import messages
from django.contrib.auth import get_user_model
from django.shortcuts import redirect

from .email_validation import INVALID_CHRIST_EMAIL_MESSAGE, is_allowed_christ_email
from .models import LoginAccessRequest, UserProfile

logger = logging.getLogger(__name__)


def is_allowed_google_email(email, allowed_domains):
    return is_allowed_christ_email(email)


class ChristGoogleAccountAdapter(DefaultSocialAccountAdapter):
    def _email_details(self, sociallogin):
        email = (getattr(sociallogin.user, "email", "") or "").strip().lower()
        if not email and getattr(sociallogin, "account", None):
            email = (sociallogin.account.extra_data.get("email", "") or "").strip().lower()
        if not email:
            for address in getattr(sociallogin, "email_addresses", []) or []:
                candidate = (getattr(address, "email", "") or "").strip().lower()
                if candidate:
                    email = candidate
                    break
        first_name = getattr(sociallogin.user, "first_name", "") or ""
        last_name = getattr(sociallogin.user, "last_name", "") or ""
        full_name = " ".join(part for part in [first_name, last_name] if part).strip()
        if not full_name and getattr(sociallogin, "account", None):
            full_name = (sociallogin.account.extra_data.get("name", "") or "").strip()
        return email, full_name

    def _validate_university_email(self, request, email):
        if not is_allowed_christ_email(email):
            messages.error(request, INVALID_CHRIST_EMAIL_MESSAGE)
            raise ImmediateHttpResponse(redirect("login"))

    def pre_social_login(self, request, sociallogin):
        email, _ = self._email_details(sociallogin)
        self._validate_university_email(request, email)
        user = get_user_model().objects.filter(email__iexact=email).first()
        if user and not user.is_active:
            messages.error(request, "Your account is inactive. Please contact the Sports Department admin.")
            raise ImmediateHttpResponse(redirect("login"))
        if user and not sociallogin.is_existing:
            sociallogin.connect(request, user)
            sociallogin.user = user

    def on_authentication_error(self, request, provider, error=None, exception=None, extra_context=None):
        logger.warning(
            "Google OAuth authentication error. provider=%s error=%s exception=%s extra_context=%s",
            getattr(provider, "id", provider),
            error,
            exception,
            extra_context,
        )
        messages.error(request, "Google login could not be completed. Please try again.")
        raise ImmediateHttpResponse(redirect("login"))

    def is_open_for_signup(self, request, sociallogin):
        email, full_name = self._email_details(sociallogin)
        self._validate_university_email(request, email)
        if not get_user_model().objects.exists():
            return True
        if get_user_model().objects.filter(email__iexact=email).exists():
            return True
        access_request, created = LoginAccessRequest.objects.get_or_create(
            email=email,
            defaults={"full_name": full_name, "status": LoginAccessRequest.Status.PENDING},
        )
        if access_request.status == LoginAccessRequest.Status.APPROVED:
            return True
        if access_request.status == LoginAccessRequest.Status.REJECTED:
            messages.error(request, "Your Sportix: sports attendance login request was rejected. Please contact the Sports Department.")
            raise ImmediateHttpResponse(redirect("login"))
        if not created and full_name and access_request.full_name != full_name:
            access_request.full_name = full_name
            access_request.save(update_fields=["full_name", "updated_at"])
        messages.warning(
            request,
            "First login? Your account requires admin approval before access is granted.",
        )
        raise ImmediateHttpResponse(redirect("login"))

    def save_user(self, request, sociallogin, form=None):
        user = super().save_user(request, sociallogin, form)
        if not get_user_model().objects.exclude(pk=user.pk).exists():
            user.is_staff = True
            user.is_superuser = True
            user.save(update_fields=["is_staff", "is_superuser"])
            user.profile.role = UserProfile.Role.SUPER_ADMIN
            user.profile.save(update_fields=["role", "updated_at"])
        return user
