from django.conf import settings
from django.db import models
from django.db.models import Q
from django.utils import timezone


class TimeStampedModel(models.Model):
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        abstract = True


class UserProfile(TimeStampedModel):
    class AccountStatus(models.TextChoices):
        ACTIVE = "ACTIVE", "Active"
        DEACTIVATED = "DEACTIVATED", "Deactivated"
        DELETED = "DELETED", "Deleted"

    class Role(models.TextChoices):
        SUPER_ADMIN = "SUPER_ADMIN", "Super Admin"
        SUB_ADMIN = "SUB_ADMIN", "Coordinator"
        TRAINER = "TRAINER", "Trainer"
        CAPTAIN = "CAPTAIN", "Captain"
        VICE_CAPTAIN = "VICE_CAPTAIN", "Vice Captain"
        COORDINATOR = "COORDINATOR", "Coordinator"
        MEMBER = "MEMBER", "Member"

    user = models.OneToOneField(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="profile")
    role = models.CharField(max_length=20, choices=Role.choices, default=Role.MEMBER)
    school = models.ForeignKey("School", on_delete=models.SET_NULL, null=True, blank=True, related_name="profiles")
    department = models.CharField(max_length=120, blank=True)
    class_name = models.CharField(max_length=120, blank=True)
    phone = models.CharField(max_length=20, blank=True)
    dob = models.DateField(null=True, blank=True)
    address = models.TextField(blank=True)
    register_no = models.CharField(max_length=60, blank=True)
    gender = models.CharField(max_length=30, blank=True)
    profile_image = models.ImageField(upload_to="profile_images/", blank=True, null=True)
    account_status = models.CharField(max_length=20, choices=AccountStatus.choices, default=AccountStatus.ACTIVE)
    account_status_note = models.TextField(blank=True)

    def __str__(self):
        return f"{self.user.get_full_name() or self.user.username} ({self.get_role_display()})"


class LoginAccessRequest(TimeStampedModel):
    class Status(models.TextChoices):
        PENDING = "PENDING", "Pending"
        APPROVED = "APPROVED", "Approved"
        REJECTED = "REJECTED", "Rejected"

    email = models.EmailField(unique=True)
    full_name = models.CharField(max_length=180, blank=True)
    status = models.CharField(max_length=20, choices=Status.choices, default=Status.PENDING)
    requested_at = models.DateTimeField(auto_now_add=True)
    reviewed_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="login_requests_reviewed")
    reviewed_at = models.DateTimeField(null=True, blank=True)
    note = models.CharField(max_length=255, blank=True)

    class Meta:
        ordering = ["-requested_at"]

    def __str__(self):
        return f"{self.email} - {self.get_status_display()}"


class Sport(TimeStampedModel):
    name = models.CharField(max_length=120, unique=True)
    description = models.TextField(blank=True)
    is_active = models.BooleanField(default=True)

    def __str__(self):
        return self.name


class School(TimeStampedModel):
    name = models.CharField(max_length=160, unique=True)
    description = models.TextField(blank=True)
    is_active = models.BooleanField(default=True)

    class Meta:
        ordering = ["name"]

    def __str__(self):
        return self.name


class Venue(TimeStampedModel):
    name = models.CharField(max_length=160, unique=True)
    location = models.CharField(max_length=180, blank=True)
    is_active = models.BooleanField(default=True)

    class Meta:
        ordering = ["name"]

    def __str__(self):
        return self.name


class Team(TimeStampedModel):
    class TeamType(models.TextChoices):
        UNIVERSITY = "UNIVERSITY", "University"
        RESERVE = "RESERVE", "Reserve"
        RECREATIONAL = "RECREATIONAL", "Recreational"

    class TeamGender(models.TextChoices):
        MALE = "MALE", "Male"
        FEMALE = "FEMALE", "Female"

    sport = models.ForeignKey(Sport, on_delete=models.CASCADE, related_name="teams")
    name = models.CharField(max_length=120)
    team_type = models.CharField(max_length=20, choices=TeamType.choices)
    gender = models.CharField(max_length=10, choices=TeamGender.choices, default=TeamGender.MALE)
    captain = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="captained_teams")
    vice_captain = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="vice_captained_teams")
    coordinator = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="coordinated_teams")
    is_active = models.BooleanField(default=True)

    class Meta:
        unique_together = ("sport", "name", "team_type", "gender")

    def __str__(self):
        return f"{self.sport} - {self.get_gender_display()} {self.name}"


class Membership(TimeStampedModel):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="memberships")
    team = models.ForeignKey(Team, on_delete=models.CASCADE, related_name="memberships")
    jersey_number = models.CharField(max_length=10, blank=True)
    is_active = models.BooleanField(default=True)
    joined_on = models.DateField(default=timezone.localdate)

    class Meta:
        unique_together = ("user", "team")

    def __str__(self):
        return f"{self.user.get_full_name() or self.user.username} - {self.team}"


class ActiveRecordManager(models.Manager):
    def get_queryset(self):
        return super().get_queryset().filter(is_deleted=False)


class Session(TimeStampedModel):
    class Status(models.TextChoices):
        SCHEDULED = "SCHEDULED", "Scheduled"
        CANCELLED = "CANCELLED", "Cancelled"

    class ScheduleSlot(models.TextChoices):
        MORNING = "MORNING", "Morning"
        EVENING = "EVENING", "Evening"

    team = models.ForeignKey(Team, on_delete=models.CASCADE, related_name="sessions")
    title = models.CharField(max_length=160, blank=True)
    start_at = models.DateTimeField()
    end_at = models.DateTimeField()
    venue = models.CharField(max_length=160)
    schedule_slot = models.CharField(max_length=20, choices=ScheduleSlot.choices, default=ScheduleSlot.MORNING)
    notes = models.TextField(blank=True)
    full_day = models.BooleanField(default=False)
    scheduled_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, related_name="scheduled_sessions")
    attendance_submitted = models.BooleanField(default=False)
    submitted_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="submitted_sessions")
    submitted_by_role = models.CharField(max_length=40, blank=True)
    submitted_at = models.DateTimeField(null=True, blank=True)
    attendance_started_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="attendance_sessions_started")
    attendance_started_by_role = models.CharField(max_length=40, blank=True)
    attendance_started_at = models.DateTimeField(null=True, blank=True)
    attendance_lock_expires_at = models.DateTimeField(null=True, blank=True)
    status = models.CharField(max_length=20, choices=Status.choices, default=Status.SCHEDULED)
    is_deleted = models.BooleanField(default=False)
    deleted_at = models.DateTimeField(null=True, blank=True)
    deleted_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="deleted_sessions",
    )

    objects = ActiveRecordManager()
    all_objects = models.Manager()

    class Meta:
        ordering = ["-start_at"]

    def __str__(self):
        return f"{self.title} - {self.team} ({self.start_at:%d %b %Y})"

    @property
    def date_range_display(self):
        start = timezone.localtime(self.start_at)
        end = timezone.localtime(self.end_at)
        if start.date() == end.date():
            return start.strftime("%d %b %Y")
        return f"{start:%d %b %Y} - {end:%d %b %Y}"

    @property
    def time_range_display(self):
        start = timezone.localtime(self.start_at)
        end = timezone.localtime(self.end_at)
        label = f"{start:%I:%M %p} - {end:%I:%M %p}"
        return f"{label} (Full Day)" if self.full_day else label


class AttendanceDelegate(TimeStampedModel):
    session = models.ForeignKey(Session, on_delete=models.CASCADE, related_name="delegates")
    assigned_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, related_name="delegations_made")
    assigned_to = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="delegations_received")
    assigned_at = models.DateTimeField(auto_now_add=True)
    reason = models.CharField(max_length=255, blank=True)

    class Meta:
        unique_together = (("session",),)

    def __str__(self):
        return f"{self.assigned_to} delegated for {self.session}"


class AttendanceDelegateLog(TimeStampedModel):
    session = models.ForeignKey(Session, on_delete=models.CASCADE, related_name="delegate_logs")
    previous_delegate = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="previous_delegate_logs")
    new_delegate = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="new_delegate_logs")
    changed_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, related_name="delegate_changes_made")
    changed_at = models.DateTimeField(auto_now_add=True)
    reason = models.CharField(max_length=255, blank=True)

    class Meta:
        ordering = ["-changed_at"]

    def __str__(self):
        return f"{self.session}: {self.previous_delegate} -> {self.new_delegate}"


class AttendanceRecord(TimeStampedModel):
    class Status(models.TextChoices):
        PRESENT = "PRESENT", "Present"
        ABSENT = "ABSENT", "Absent"
        LATE = "LATE", "Late"
        EARLY_EXIT = "EARLY_EXIT", "Early Exit"

    session = models.ForeignKey(Session, on_delete=models.CASCADE, related_name="attendance_records")
    member = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="attendance_records")
    status = models.CharField(max_length=20, choices=Status.choices)
    marked_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, related_name="attendance_marked")
    marked_at = models.DateTimeField(auto_now_add=True)
    remarks = models.CharField(max_length=255, blank=True)

    class Meta:
        unique_together = ("session", "member")
        ordering = ["member__first_name", "member__last_name", "member__username"]

    def __str__(self):
        return f"{self.member} - {self.session}: {self.status}"


class AttendanceEditLog(TimeStampedModel):
    attendance_record = models.ForeignKey(AttendanceRecord, on_delete=models.CASCADE, related_name="edit_logs")
    edited_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True)
    old_status = models.CharField(max_length=20, choices=AttendanceRecord.Status.choices)
    new_status = models.CharField(max_length=20, choices=AttendanceRecord.Status.choices)
    reason = models.CharField(max_length=255)
    edited_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.attendance_record} {self.old_status}->{self.new_status}"


class Meeting(TimeStampedModel):
    class Status(models.TextChoices):
        SCHEDULED = "SCHEDULED", "Scheduled"
        COMPLETED = "COMPLETED", "Completed"
        CANCELLED = "CANCELLED", "Cancelled"

    title = models.CharField(max_length=160)
    meeting_date = models.DateField()
    start_time = models.TimeField()
    end_time = models.TimeField()
    venue = models.CharField(max_length=160)
    agenda = models.TextField(blank=True)
    sports = models.ManyToManyField(Sport, related_name="meetings", blank=True)
    teams = models.ManyToManyField(Team, related_name="meetings", blank=True)
    trainers = models.ManyToManyField(settings.AUTH_USER_MODEL, related_name="trainer_meetings", blank=True)
    participants = models.ManyToManyField(settings.AUTH_USER_MODEL, related_name="meetings", blank=True)
    scheduled_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, related_name="meetings_scheduled")
    status = models.CharField(max_length=20, choices=Status.choices, default=Status.SCHEDULED)
    ending_note = models.TextField(blank=True)
    is_deleted = models.BooleanField(default=False)
    deleted_at = models.DateTimeField(null=True, blank=True)
    deleted_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="deleted_meetings",
    )

    objects = ActiveRecordManager()
    all_objects = models.Manager()

    class Meta:
        ordering = ["-meeting_date", "-start_time"]

    def __str__(self):
        return f"{self.title} ({self.meeting_date:%d %b %Y})"


class Feedback(TimeStampedModel):
    class FeedbackType(models.TextChoices):
        ADMIN_TO_STUDENT = "ADMIN_TO_STUDENT", "Admin to Student"
        STUDENT_TO_ADMIN = "STUDENT_TO_ADMIN", "Student to Admin"
        COORDINATOR_TO_STUDENT = "COORDINATOR_TO_STUDENT", "Coordinator to Student"

    sender = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, related_name="feedback_sent")
    receiver = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="feedback_received")
    session = models.ForeignKey(Session, on_delete=models.SET_NULL, null=True, blank=True, related_name="feedback")
    feedback_type = models.CharField(max_length=30, choices=FeedbackType.choices)
    message = models.TextField()
    is_read = models.BooleanField(default=False)

    class Meta:
        ordering = ["-created_at"]
        constraints = [
            models.CheckConstraint(
                condition=~Q(message=""),
                name="feedback_message_not_empty",
            )
        ]

    def __str__(self):
        return f"{self.get_feedback_type_display()} from {self.sender} to {self.receiver}"


class Notification(TimeStampedModel):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="notifications")
    actor = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="notifications_created")
    title = models.CharField(max_length=160)
    message = models.CharField(max_length=255)
    target_url = models.CharField(max_length=255, blank=True)
    sport = models.ForeignKey(Sport, on_delete=models.SET_NULL, null=True, blank=True, related_name="notifications")
    team = models.ForeignKey(Team, on_delete=models.SET_NULL, null=True, blank=True, related_name="notifications")
    session = models.ForeignKey(Session, on_delete=models.SET_NULL, null=True, blank=True, related_name="notifications")
    is_read = models.BooleanField(default=False)
    read_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        ordering = ["-created_at"]
        indexes = [
            models.Index(fields=["user", "is_read", "-created_at"]),
        ]

    def __str__(self):
        return f"{self.title} for {self.user}"


class EmailNotificationJob(TimeStampedModel):
    class Status(models.TextChoices):
        PENDING = "PENDING", "Pending"
        PROCESSING = "PROCESSING", "Processing"
        SENT = "SENT", "Sent"
        FAILED = "FAILED", "Failed"

    notification_type = models.CharField(max_length=80)
    context = models.JSONField(default=dict)
    recipient_user_ids = models.JSONField(default=list)
    status = models.CharField(max_length=12, choices=Status.choices, default=Status.PENDING, db_index=True)
    attempts = models.PositiveSmallIntegerField(default=0)
    next_attempt_at = models.DateTimeField(default=timezone.now, db_index=True)
    last_error = models.TextField(blank=True)
    processed_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        ordering = ["created_at"]

    def __str__(self):
        return f"{self.notification_type} ({self.status})"
