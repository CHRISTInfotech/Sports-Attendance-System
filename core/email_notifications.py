import logging
from email.mime.image import MIMEImage
from pathlib import Path

from django.conf import settings
from django.core.mail import EmailMultiAlternatives, get_connection
from django.template.loader import render_to_string
from django.utils.html import strip_tags


logger = logging.getLogger(__name__)


SUBJECTS = {
    "student_signup_request": "New Student Signup Request",
    "practice_scheduled": "Practice Session Scheduled",
    "practice_updated": "Practice Session Updated",
    "practice_rescheduled": "Practice Session Rescheduled",
    "practice_cancelled": "Practice Session Cancelled",
    "meeting_scheduled": "Meeting Scheduled",
    "meeting_updated": "Meeting Updated",
    "meeting_rescheduled": "Meeting Rescheduled",
    "meeting_completed": "Meeting Completed",
    "meeting_cancelled": "Meeting Cancelled",
    "feedback_received": "New Feedback Received",
    "captain_assigned": "Captain Assignment",
    "captain_removed": "Captain Role Update",
    "vice_captain_assigned": "Vice Captain Assignment",
    "vice_captain_removed": "Vice Captain Role Update",
    "trainer_assigned": "Trainer Assignment",
    "trainer_removed": "Trainer Assignment Update",
    "team_created": "New Team Created",
    "sport_created": "New Sport Added",
    "team_membership": "Team Membership Update",
    "session_incharge": "Session Incharge Assignment",
}


def queue_notification_emails(users, notification_type, context):
    """Persist a durable email job without performing network I/O in the request."""
    if not getattr(settings, "EMAIL_NOTIFICATIONS_ENABLED", False):
        return None

    user_ids = []
    seen = set()
    for user in users:
        if not user or user.pk in seen or not user.is_active or not (user.email or "").strip():
            continue
        seen.add(user.pk)
        user_ids.append(user.pk)
    if not user_ids:
        return None

    from .models import EmailNotificationJob

    return EmailNotificationJob.objects.create(
        notification_type=notification_type,
        context=context,
        recipient_user_ids=user_ids,
    )


def send_notification_emails(users, notification_type, context):
    """Send one notification email per unique, active user with an email address.

    Mail delivery is deliberately best-effort: configuration and provider failures are
    logged and never escape into the application action that created the notification.
    """
    if not getattr(settings, "EMAIL_NOTIFICATIONS_ENABLED", False):
        return {"sent_user_ids": [], "failed_user_ids": []}

    seen = set()
    recipients = []
    for user in users:
        email = (getattr(user, "email", "") or "").strip().lower()
        if not email or email in seen or not getattr(user, "is_active", True):
            continue
        seen.add(email)
        recipients.append((user, email))

    subject = SUBJECTS.get(notification_type, context.get("title", "Sportix: Sports Attendance Notification"))
    messages = []
    preparation_failed_ids = []
    for user, email in recipients:
        try:
            email_context = {**context, "recipient": user, "notification_type": notification_type}
            html = render_to_string("emails/notification.html", email_context)
            message = EmailMultiAlternatives(
                subject=f"[Sportix: Sports Attendance System] {subject}",
                body=strip_tags(html),
                from_email=settings.DEFAULT_FROM_EMAIL,
                to=[email],
            )
            message.attach_alternative(html, "text/html")
            logo_path = Path(settings.BASE_DIR) / "static" / "img" / "christ-logo.jpeg"
            if logo_path.exists():
                logo = MIMEImage(logo_path.read_bytes(), _subtype="jpeg")
                logo.add_header("Content-ID", "<christ-logo>")
                logo.add_header("Content-Disposition", "inline", filename="christ-logo.jpeg")
                message.attach(logo)
            messages.append((user, message))
        except Exception as exc:
            preparation_failed_ids.append(user.pk)
            logger.exception(
                "Email preparation failed (notification=%s, user_id=%s, error_type=%s, error=%s).",
                notification_type,
                user.pk,
                type(exc).__name__,
                exc,
            )

    if not messages:
        return {"sent_user_ids": [], "failed_user_ids": [user.pk for user, _ in recipients]}

    connection = get_connection(fail_silently=False)
    try:
        # Authenticate once for the whole notification batch. In particular, a 535
        # response now delays the application only once rather than once per user.
        connection.open()
    except Exception as exc:
        logger.exception(
            "SMTP connection/authentication failed; skipped %s email(s) "
            "(notification=%s, error_type=%s, error=%s). Check SMTP AUTH policy, "
            "username/app password, host/port, and TLS settings.",
            len(messages),
            notification_type,
            type(exc).__name__,
            exc,
        )
        return {"sent_user_ids": [], "failed_user_ids": preparation_failed_ids + [user.pk for user, _ in messages]}

    sent_user_ids = []
    failed_user_ids = list(preparation_failed_ids)
    try:
        for user, message in messages:
            try:
                message.connection = connection
                message.send(fail_silently=False)
                sent_user_ids.append(user.pk)
            except Exception as exc:
                failed_user_ids.append(user.pk)
                logger.exception(
                    "SMTP message delivery failed (notification=%s, user_id=%s, error_type=%s, error=%s).",
                    notification_type,
                    user.pk,
                    type(exc).__name__,
                    exc,
                )
    finally:
        try:
            connection.close()
        except Exception as exc:
            logger.warning("SMTP connection close failed (error_type=%s, error=%s).", type(exc).__name__, exc)
    return {"sent_user_ids": sent_user_ids, "failed_user_ids": failed_user_ids}
