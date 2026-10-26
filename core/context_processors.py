from .models import Notification, Team, UserProfile


def breadcrumbs(request):
    match = getattr(request, "resolver_match", None)
    url_name = getattr(match, "url_name", "") if match else ""
    labels = {
        "dashboard": "Dashboard",
        "analytics": "Analytics",
        "sports": "Sports",
        "teams": "Teams",
        "members": "Add Students",
        "member_detail": "Student Details",
        "trainers": "Add Trainer",
        "schools": "Schools",
        "venues": "Venues",
        "settings": "Settings",
        "deactivated_users": "Deleted / Inactive Users",
        "deleted_sessions": "Deleted Sessions",
        "deleted_meetings": "Deleted Meetings",
        "sessions": "Schedule Practice",
        "meetings": "Schedule Meeting",
        "attendance_detail": "Attendance Details",
        "take_attendance": "Take Attendance",
        "delegate_attendance": "Assign Session Incharge",
        "my_attendance": "My Attendance",
        "my_profile": "My Profile",
        "feedback": "Feedback",
        "send_feedback": "Send Feedback",
        "reports": "Reports",
        "notifications": "Notifications",
    }
    if not request.user.is_authenticated:
        return {}
    if url_name == "dashboard":
        return {"breadcrumb_items": []}
    if url_name in {"notifications", "deactivated_users", "deleted_sessions", "deleted_meetings"}:
        return {"breadcrumb_items": [
            {"label": "Dashboard", "url_name": "dashboard"},
            {"label": "Settings", "url_name": "settings"},
            {"label": labels.get(url_name, "Page")},
        ]}
    if url_name == "send_feedback":
        return {"breadcrumb_items": [
            {"label": "Dashboard", "url_name": "dashboard"},
            {"label": "Feedback", "url_name": "feedback"},
            {"label": labels.get(url_name, "Page")},
        ]}
    if url_name in {"take_attendance", "attendance_detail"}:
        return {"breadcrumb_items": [
            {"label": "Dashboard", "url_name": "dashboard"},
            {"label": "Practice Sessions", "url_name": "sessions"},
            {"label": labels.get(url_name, "Page")},
        ]}
    return {"breadcrumb_items": [{"label": "Dashboard", "url_name": "dashboard"}, {"label": labels.get(url_name, "Page")}]}


def role_display(request):
    user = getattr(request, "user", None)
    if not user or not user.is_authenticated:
        return {}
    if user.is_superuser:
        return {"topbar_role_label": "Super Admin"}
    profile = getattr(user, "profile", None)
    if not profile:
        return {"topbar_role_label": "Student"}
    if profile.role == UserProfile.Role.TRAINER:
        return {"topbar_role_label": "Trainer"}
    if profile.role in {UserProfile.Role.SUPER_ADMIN, UserProfile.Role.SUB_ADMIN}:
        return {"topbar_role_label": profile.get_role_display()}
    labels = ["Student"]
    if profile.role == UserProfile.Role.CAPTAIN or Team.objects.filter(captain=user).exists():
        labels.append("Captain")
    if profile.role == UserProfile.Role.VICE_CAPTAIN or Team.objects.filter(vice_captain=user).exists():
        labels.append("Vice Captain")
    if profile.role == UserProfile.Role.COORDINATOR or Team.objects.filter(coordinator=user).exists():
        labels.append("Coordinator")
    return {"topbar_role_label": " | ".join(labels)}


def notifications(request):
    user = getattr(request, "user", None)
    if not user or not user.is_authenticated:
        return {}
    qs = Notification.objects.filter(user=user).select_related("actor", "team", "sport", "session")
    return {
        "topbar_notifications": qs[:5],
        "topbar_unread_notifications": qs.filter(is_read=False).count(),
    }
