from django import forms
from django.contrib.auth import get_user_model
from django.db.models import Q
from django.utils import timezone
from datetime import datetime, time

from .models import AttendanceRecord, Feedback, LoginAccessRequest, Membership, School, Session, Sport, Team, UserProfile, Venue
from .email_validation import INVALID_CHRIST_EMAIL_MESSAGE, validate_christ_email

User = get_user_model()

REGISTRATION_NUMBER_EXISTS_MESSAGE = "This registration number already exists"


def validate_profile_register_no(register_no, user=None):
    register_no = str(register_no or "").strip()
    if not register_no:
        return register_no
    existing = UserProfile.objects.filter(register_no__iexact=register_no)
    if user and user.pk:
        existing = existing.exclude(user=user)
    if existing.exists():
        raise forms.ValidationError(REGISTRATION_NUMBER_EXISTS_MESSAGE)
    return register_no


class BootstrapFormMixin:
    select_placeholders = {
        "sport": "Select Sport",
        "team": "Select Team",
        "student": "Select Student",
        "user": "Select Student",
        "captain": "Select Captain",
        "vice_captain": "Select Vice Captain",
        "coordinator": "Select Trainer",
        "assigned_to": "Select Student",
        "receiver": "Select Student",
        "session": "Select Session",
        "role": "Select Role",
        "school": "Select School",
        "gender": "Select Gender",
        "team_type": "Select Team Type",
        "schedule_slot": "Select Schedule",
        "venue_choice": "Select Venue",
        "status": "Select Status",
    }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        for field_name, field in self.fields.items():
            css = "form-select" if isinstance(field.widget, forms.Select) else "form-control"
            if isinstance(field.widget, forms.CheckboxInput):
                css = "form-check-input"
            field.widget.attrs.setdefault("class", css)
            if isinstance(field, forms.ModelChoiceField):
                field.empty_label = self.select_placeholders.get(field_name, f"Select {field.label}")
            elif isinstance(field.widget, forms.Select) and not isinstance(field.widget, forms.CheckboxSelectMultiple):
                choices = list(field.choices)
                if choices and choices[0][0] not in {"", None}:
                    placeholder = self.select_placeholders.get(field_name, f"Select {field.label}")
                    field.choices = [("", placeholder)] + choices


class StudentSignupRequestForm(BootstrapFormMixin, forms.Form):
    student_name = forms.CharField(max_length=180, label="Student Name")
    email = forms.EmailField(label="Email")
    mobile_number = forms.CharField(max_length=10, label="Mobile Number")
    register_no = forms.CharField(max_length=60, label="Registration Number")
    school = forms.ModelChoiceField(queryset=School.objects.none())
    department = forms.CharField(max_length=120)
    gender = forms.ChoiceField(choices=[("", "Select Gender"), ("Male", "Male"), ("Female", "Female")])
    sports = forms.ModelMultipleChoiceField(queryset=Sport.objects.none())

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["school"].queryset = School.objects.filter(is_active=True).order_by("name")
        self.fields["sports"].queryset = Sport.objects.filter(is_active=True).order_by("name")
        self.fields["email"].widget.attrs.update({
            "pattern": r"^[^@\s]+@([A-Za-z0-9-]+\.)*christuniversity\.in$",
            "title": "Use a Christ University email ending with christuniversity.in",
        })
        self.fields["mobile_number"].widget.attrs.update({"inputmode": "numeric", "pattern": r"[0-9]{10}", "maxlength": "10"})

    def clean_email(self):
        try:
            email = validate_christ_email(self.cleaned_data["email"])
        except ValueError:
            raise forms.ValidationError(INVALID_CHRIST_EMAIL_MESSAGE)
        if User.objects.filter(email__iexact=email).exists():
            raise forms.ValidationError("An account already exists for this email. Please continue with Google login.")
        existing_request = LoginAccessRequest.objects.filter(email__iexact=email).first()
        if existing_request:
            if existing_request.status == LoginAccessRequest.Status.APPROVED:
                raise forms.ValidationError("This request is already approved. Please continue with Google login.")
            raise forms.ValidationError("Your details/request have already been submitted.")
        return email

    def clean_mobile_number(self):
        value = str(self.cleaned_data.get("mobile_number") or "").strip()
        if not value.isdigit() or len(value) != 10:
            raise forms.ValidationError("Mobile number must be exactly 10 digits.")
        return value

    def clean_register_no(self):
        value = str(self.cleaned_data.get("register_no") or "").strip()
        if UserProfile.objects.filter(register_no__iexact=value).exists() or LoginAccessRequest.objects.filter(register_no__iexact=value).exists():
            raise forms.ValidationError(REGISTRATION_NUMBER_EXISTS_MESSAGE)
        return value


class SportForm(BootstrapFormMixin, forms.ModelForm):
    class Meta:
        model = Sport
        fields = ["name", "description", "is_active"]


class VenueForm(BootstrapFormMixin, forms.ModelForm):
    class Meta:
        model = Venue
        fields = ["name", "location", "is_active"]


class SchoolForm(BootstrapFormMixin, forms.ModelForm):
    class Meta:
        model = School
        fields = ["name", "description", "is_active"]


class TeamForm(BootstrapFormMixin, forms.ModelForm):
    class Meta:
        model = Team
        fields = ["sport", "name", "team_type", "gender", "captain", "vice_captain", "coordinator", "is_active"]

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        student_roles = [UserProfile.Role.MEMBER, UserProfile.Role.CAPTAIN, UserProfile.Role.VICE_CAPTAIN]
        trainer_roles = [UserProfile.Role.TRAINER, UserProfile.Role.COORDINATOR]
        student_qs = User.objects.filter(is_active=True, profile__role__in=student_roles).order_by("first_name", "last_name", "email")
        trainer_qs = User.objects.filter(is_active=True, profile__role__in=trainer_roles).order_by("first_name", "last_name", "email")
        self.fields["captain"].queryset = student_qs
        self.fields["vice_captain"].queryset = student_qs
        self.fields["coordinator"].queryset = trainer_qs

    def clean(self):
        cleaned = super().clean()
        sport = cleaned.get("sport")
        gender = cleaned.get("gender")
        captain = cleaned.get("captain")
        vice_captain = cleaned.get("vice_captain")
        if captain and vice_captain and captain == vice_captain:
            self.add_error("vice_captain", "Captain and Vice Captain cannot be the same student.")
        for field_name, leader in (("captain", captain), ("vice_captain", vice_captain)):
            if not leader:
                continue
            leader_name = leader.get_full_name() or leader.email or leader.username
            if sport and not Membership.objects.filter(user=leader, team__sport=sport, is_active=True).exists():
                label = "Captain" if field_name == "captain" else "Vice Captain"
                self.add_error(field_name, f"{label} must be associated with the selected sport.")
            if sport and gender:
                same_group_memberships = Membership.objects.filter(
                    user=leader,
                    is_active=True,
                    team__sport=sport,
                    team__gender=gender,
                ).select_related("team", "team__sport")
                if self.instance and self.instance.pk:
                    same_group_memberships = same_group_memberships.exclude(team=self.instance)
                existing_membership = same_group_memberships.first()
                if existing_membership:
                    self.add_error(
                        field_name,
                        f"{leader_name} is already a member of {existing_membership.team} and cannot be assigned to another team category in the same Sport and Gender.",
                    )
            existing_captain = Team.objects.filter(captain=leader)
            existing_vice_captain = Team.objects.filter(vice_captain=leader)
            if self.instance and self.instance.pk:
                existing_captain = existing_captain.exclude(pk=self.instance.pk)
                existing_vice_captain = existing_vice_captain.exclude(pk=self.instance.pk)
            if existing_captain.exists():
                self.add_error(field_name, f"{leader_name} is already a captain of other team and cannot be captain again.")
            if existing_vice_captain.exists():
                self.add_error(field_name, f"{leader_name} is already a vice captain of other team and cannot be assigned again.")
        return cleaned


class MembershipForm(BootstrapFormMixin, forms.ModelForm):
    class Meta:
        model = Membership
        fields = ["user", "team", "jersey_number", "is_active", "joined_on"]
        widgets = {"joined_on": forms.DateInput(attrs={"type": "date"})}


class UserRoleForm(BootstrapFormMixin, forms.ModelForm):
    first_name = forms.CharField(max_length=150, required=False)
    last_name = forms.CharField(max_length=150, required=False)
    email = forms.EmailField()

    class Meta:
        model = UserProfile
        fields = ["role", "department", "phone"]

    def __init__(self, *args, **kwargs):
        self.user_instance = kwargs.pop("user_instance", None)
        super().__init__(*args, **kwargs)
        if self.user_instance:
            self.fields["first_name"].initial = self.user_instance.first_name
            self.fields["last_name"].initial = self.user_instance.last_name
            self.fields["email"].initial = self.user_instance.email

    def save(self, commit=True):
        profile = super().save(commit=False)
        user = profile.user
        user.first_name = self.cleaned_data["first_name"]
        user.last_name = self.cleaned_data["last_name"]
        user.email = self.cleaned_data["email"]
        user.username = self.cleaned_data["email"]
        if commit:
            user.save()
            profile.save()
        return profile

    def clean_email(self):
        try:
            return validate_christ_email(self.cleaned_data["email"])
        except ValueError:
            raise forms.ValidationError(INVALID_CHRIST_EMAIL_MESSAGE)


class ProfileForm(BootstrapFormMixin, forms.ModelForm):
    first_name = forms.CharField(max_length=150, required=False)
    last_name = forms.CharField(max_length=150, required=False)
    email = forms.EmailField(disabled=True, required=False)
    gender = forms.ChoiceField(choices=[("", "Select Gender"), ("Male", "Male"), ("Female", "Female")], required=False)

    class Meta:
        model = UserProfile
        fields = ["profile_image", "school", "department", "phone", "dob", "address", "register_no", "gender"]
        widgets = {
            "dob": forms.DateInput(attrs={"type": "date"}),
            "address": forms.Textarea(attrs={"rows": 3}),
            "profile_image": forms.ClearableFileInput(attrs={"accept": "image/*"}),
        }

    def __init__(self, *args, **kwargs):
        self.user_instance = kwargs.pop("user_instance")
        super().__init__(*args, **kwargs)
        self.fields["school"].queryset = School.objects.filter(is_active=True).order_by("name")
        self.fields["first_name"].initial = self.user_instance.first_name
        self.fields["last_name"].initial = self.user_instance.last_name
        self.fields["email"].initial = self.user_instance.email

    def clean_profile_image(self):
        image = self.cleaned_data.get("profile_image")
        if image and hasattr(image, "content_type") and not image.content_type.startswith("image/"):
            raise forms.ValidationError("Only image files can be uploaded.")
        return image

    def clean_register_no(self):
        return validate_profile_register_no(self.cleaned_data.get("register_no"), self.user_instance)

    def save(self, commit=True):
        profile = super().save(commit=False)
        self.user_instance.first_name = self.cleaned_data["first_name"]
        self.user_instance.last_name = self.cleaned_data["last_name"]
        if commit:
            self.user_instance.save(update_fields=["first_name", "last_name"])
            profile.save()
        return profile


class SessionForm(BootstrapFormMixin, forms.ModelForm):
    start_date = forms.DateField(widget=forms.DateInput(attrs={"type": "date"}))
    end_date = forms.DateField(widget=forms.DateInput(attrs={"type": "date"}))
    from_time = forms.TimeField(widget=forms.TimeInput(attrs={"type": "time"}), required=True)
    to_time = forms.TimeField(widget=forms.TimeInput(attrs={"type": "time"}), required=True)
    venue_choice = forms.ChoiceField(choices=(), required=True)
    other_venue = forms.CharField(max_length=160, required=False)

    class Meta:
        model = Session
        fields = ["team", "title", "schedule_slot", "full_day", "notes"]
        widgets = {"notes": forms.Textarea(attrs={"rows": 3})}

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["title"].required = False
        venue_choices = [(venue.name, venue.name) for venue in Venue.objects.filter(is_active=True)]
        self.fields["venue_choice"].choices = [("", "Select Venue")] + venue_choices + [("OTHER", "Other")]
        if self.instance and self.instance.pk:
            self.fields["start_date"].initial = timezone.localtime(self.instance.start_at).date()
            self.fields["end_date"].initial = timezone.localtime(self.instance.end_at).date()
            self.fields["from_time"].initial = timezone.localtime(self.instance.start_at).time().replace(second=0, microsecond=0)
            self.fields["to_time"].initial = timezone.localtime(self.instance.end_at).time().replace(second=0, microsecond=0)
            self.fields["venue_choice"].initial = self.instance.venue
            if self.instance.venue.startswith("Other - "):
                self.fields["venue_choice"].initial = "OTHER"
                self.fields["other_venue"].initial = self.instance.venue.replace("Other - ", "", 1)

    def clean(self):
        cleaned = super().clean()
        start_date = cleaned.get("start_date")
        end_date = cleaned.get("end_date")
        from_time = cleaned.get("from_time")
        to_time = cleaned.get("to_time")
        if start_date and end_date and end_date < start_date:
            self.add_error("end_date", "End date cannot be before start date.")
        if start_date and end_date and from_time and to_time and start_date == end_date and to_time <= from_time:
            self.add_error("to_time", "To Time must be after From Time.")
        return cleaned

    def save(self, commit=True):
        session = super().save(commit=False)
        session.start_at = timezone.make_aware(datetime.combine(self.cleaned_data["start_date"], self.cleaned_data["from_time"]))
        session.end_at = timezone.make_aware(datetime.combine(self.cleaned_data["end_date"], self.cleaned_data["to_time"]))
        session.title = self.cleaned_data.get("title") or "Practice session"
        if self.cleaned_data["venue_choice"] == "OTHER":
            session.venue = f"Other - {self.cleaned_data['other_venue']}" if self.cleaned_data.get("other_venue") else "Other"
        else:
            session.venue = self.cleaned_data["venue_choice"]
        if commit:
            session.save()
        return session


class DelegateForm(BootstrapFormMixin, forms.Form):
    assigned_to = forms.ModelChoiceField(queryset=User.objects.none(), required=False)
    remove_delegate = forms.BooleanField(required=False, label="Remove current Session Incharge")
    reason = forms.CharField(max_length=255, required=False)

    def __init__(self, *args, **kwargs):
        session = kwargs.pop("session")
        super().__init__(*args, **kwargs)
        self.fields["assigned_to"].queryset = User.objects.filter(
            memberships__team__sport=session.team.sport,
            memberships__is_active=True,
        ).distinct()
        current_delegate = session.delegates.select_related("assigned_to").first()
        if current_delegate:
            self.fields["assigned_to"].initial = current_delegate.assigned_to

    def clean(self):
        cleaned = super().clean()
        assigned_to = cleaned.get("assigned_to")
        remove_delegate = cleaned.get("remove_delegate")
        if assigned_to and remove_delegate:
            self.add_error("remove_delegate", "Choose either a new Session Incharge or remove the current one.")
        if not assigned_to and not remove_delegate:
            self.add_error("assigned_to", "Select a Session Incharge or choose remove current Session Incharge.")
        return cleaned


class AttendanceEditForm(BootstrapFormMixin, forms.ModelForm):
    reason = forms.CharField(max_length=255, required=False)

    class Meta:
        model = AttendanceRecord
        fields = ["status", "remarks"]


class FeedbackForm(BootstrapFormMixin, forms.ModelForm):
    class Meta:
        model = Feedback
        fields = ["receiver", "session", "message"]
        widgets = {"message": forms.Textarea(attrs={"rows": 4})}


class SessionFeedbackForm(BootstrapFormMixin, forms.ModelForm):
    class Meta:
        model = Feedback
        fields = ["session", "message"]
        widgets = {"message": forms.Textarea(attrs={"rows": 4})}

    def __init__(self, *args, **kwargs):
        user = kwargs.pop("user", None)
        super().__init__(*args, **kwargs)
        if user is None:
            self.fields["session"].queryset = Session.objects.none()
            return
        assigned_teams = Team.objects.filter(
            Q(memberships__user=user, memberships__is_active=True)
            | Q(captain=user)
            | Q(vice_captain=user)
        ).distinct()
        self.fields["session"].queryset = Session.objects.filter(team__in=assigned_teams).select_related("team", "team__sport").order_by("-start_at")


class ReportFilterForm(BootstrapFormMixin, forms.Form):
    REPORT_TYPES = [
        ("attendance", "Attendance Report"),
        ("sessions", "Practice Session Report"),
        ("gender", "Gender-wise Report"),
        ("sports_teams", "Sports & Team Report"),
        ("trainer", "Trainer Activity Report"),
        ("feedback", "Feedback Report"),
    ]
    SESSION_STATUS_CHOICES = [
        ("", "Select Session Status"),
        ("SCHEDULED", "Scheduled"),
        ("UPCOMING", "Upcoming"),
        ("COMPLETED", "Completed"),
    ]

    report_type = forms.ChoiceField(choices=REPORT_TYPES, initial="attendance")
    sport = forms.ModelChoiceField(queryset=Sport.objects.all(), required=False)
    gender = forms.ChoiceField(choices=Team.TeamGender.choices, required=False)
    team_type = forms.ChoiceField(choices=Team.TeamType.choices, required=False, label="Team Category")
    team = forms.ModelChoiceField(queryset=Team.objects.all(), required=False)
    session = forms.ModelChoiceField(queryset=Session.objects.select_related("team", "team__sport").all(), required=False, label="Practice Session")
    trainer = forms.ModelChoiceField(
        queryset=User.objects.filter(profile__role__in=[UserProfile.Role.TRAINER, UserProfile.Role.COORDINATOR]).distinct(),
        required=False,
    )
    student = forms.ModelChoiceField(queryset=User.objects.all(), required=False)
    attendance_status = forms.ChoiceField(choices=AttendanceRecord.Status.choices, required=False, label="Attendance Status")
    venue = forms.ChoiceField(choices=[], required=False)
    session_status = forms.ChoiceField(choices=SESSION_STATUS_CHOICES, required=False)
    start_date = forms.DateField(required=False, widget=forms.DateInput(attrs={"type": "date"}))
    end_date = forms.DateField(required=False, widget=forms.DateInput(attrs={"type": "date"}))

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        venue_values = Session.objects.exclude(venue="").order_by("venue").values_list("venue", flat=True).distinct()
        self.fields["venue"].choices = [("", "Select Venue")] + [(venue, venue) for venue in venue_values]
        self.fields["student"].queryset = User.objects.filter(
            Q(profile__role=UserProfile.Role.MEMBER)
            | Q(profile__role=UserProfile.Role.CAPTAIN)
            | Q(profile__role=UserProfile.Role.VICE_CAPTAIN)
        ).distinct().order_by("first_name", "last_name", "email")
        self.fields["team"].queryset = Team.objects.select_related("sport").order_by("sport__name", "gender", "team_type", "name")
        self.fields["session"].queryset = Session.objects.select_related("team", "team__sport").order_by("-start_at")
