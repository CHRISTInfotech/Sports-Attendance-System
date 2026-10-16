from io import BytesIO

from django.contrib.auth import get_user_model
from django.test import TestCase
from django.test import override_settings
from django.core import mail
from django.core.management import call_command
from unittest.mock import patch
from django.urls import reverse
from django.utils import timezone
from openpyxl import Workbook

from .models import AttendanceRecord, EmailNotificationJob, Membership, Session, Sport, Team, UserProfile
from .views import import_sessions_from_file, import_students_from_file
from .email_notifications import queue_notification_emails, send_notification_emails


User = get_user_model()


@override_settings(
    EMAIL_NOTIFICATIONS_ENABLED=True,
    EMAIL_BACKEND="django.core.mail.backends.locmem.EmailBackend",
    DEFAULT_FROM_EMAIL="sports.lavasa@christuniversity.in",
)
class EmailNotificationTests(TestCase):
    def setUp(self):
        self.user = User.objects.create_user(
            username="mail-recipient@christuniversity.in",
            email="mail-recipient@christuniversity.in",
            first_name="Mail",
            last_name="Recipient",
        )

    def test_sends_one_html_email_for_duplicate_recipient(self):
        send_notification_emails(
            [self.user, self.user],
            "meeting_scheduled",
            {"title": "New meeting scheduled", "message": "A meeting was scheduled.", "details": [("Venue", "Seminar Hall")]},
        )

        self.assertEqual(len(mail.outbox), 1)
        self.assertEqual(mail.outbox[0].to, [self.user.email])
        self.assertIn("Meeting Scheduled", mail.outbox[0].subject)
        self.assertTrue(mail.outbox[0].alternatives)
        html = mail.outbox[0].alternatives[0].content
        self.assertIn("Dear Mail Recipient", html)
        self.assertIn("Seminar Hall", html)
        self.assertNotIn("{{", html)
        self.assertNotIn("}}", html)

    @patch("core.email_notifications.get_connection")
    def test_queue_does_not_open_smtp_connection(self, mocked_get_connection):
        queue_notification_emails(
            [self.user, self.user],
            "practice_scheduled",
            {"title": "Practice", "message": "Practice scheduled."},
        )

        mocked_get_connection.assert_not_called()
        job = EmailNotificationJob.objects.get()
        self.assertEqual(job.recipient_user_ids, [self.user.pk])

    def test_worker_delivers_queued_email(self):
        job = queue_notification_emails(
            [self.user],
            "practice_scheduled",
            {"title": "Practice", "message": "Practice scheduled."},
        )

        call_command("run_email_worker", once=True, verbosity=0)

        job.refresh_from_db()
        self.assertEqual(job.status, EmailNotificationJob.Status.SENT)
        self.assertEqual(len(mail.outbox), 1)

    @patch("core.email_notifications.EmailMultiAlternatives.send", side_effect=RuntimeError("SMTP unavailable"))
    def test_delivery_failure_does_not_escape(self, mocked_send):
        send_notification_emails(
            [self.user],
            "feedback_received",
            {"title": "New feedback", "message": "Feedback received."},
        )
        mocked_send.assert_called_once()

    @patch("core.email_notifications.get_connection")
    def test_batch_authentication_failure_is_attempted_only_once(self, mocked_get_connection):
        second_user = User.objects.create_user(
            username="second-recipient@christuniversity.in",
            email="second-recipient@christuniversity.in",
        )
        connection = mocked_get_connection.return_value
        connection.open.side_effect = RuntimeError("Authentication rejected")

        send_notification_emails(
            [self.user, second_user],
            "practice_scheduled",
            {"title": "Practice", "message": "Practice scheduled."},
        )

        connection.open.assert_called_once_with()


class TrainerStudentPermissionTests(TestCase):
    def setUp(self):
        self.trainer = self.make_user("trainer@christuniversity.in", UserProfile.Role.TRAINER)
        self.other_trainer = self.make_user("other-trainer@christuniversity.in", UserProfile.Role.TRAINER)
        self.assigned_student = self.make_user("assigned@christuniversity.in", UserProfile.Role.MEMBER)
        self.other_student = self.make_user("other@christuniversity.in", UserProfile.Role.MEMBER)
        self.sport = Sport.objects.create(name="Cricket")
        self.other_sport = Sport.objects.create(name="Football")
        self.assigned_team = Team.objects.create(
            sport=self.sport,
            name="University Team",
            team_type=Team.TeamType.UNIVERSITY,
            gender=Team.TeamGender.MALE,
            coordinator=self.trainer,
        )
        self.other_team = Team.objects.create(
            sport=self.other_sport,
            name="Reserve Team",
            team_type=Team.TeamType.RESERVE,
            gender=Team.TeamGender.FEMALE,
            coordinator=self.other_trainer,
        )
        self.assigned_membership = Membership.objects.create(user=self.assigned_student, team=self.assigned_team)
        self.other_membership = Membership.objects.create(user=self.other_student, team=self.other_team)

    def make_user(self, email, role):
        user = User.objects.create_user(username=email, email=email, password="pass")
        profile, _ = UserProfile.objects.get_or_create(user=user)
        profile.role = role
        profile.save(update_fields=["role", "updated_at"])
        return user

    def test_trainer_sees_only_students_from_assigned_teams(self):
        self.client.force_login(self.trainer)

        response = self.client.get(reverse("members"))

        self.assertContains(response, self.assigned_student.email)
        self.assertNotContains(response, self.other_student.email)

    def test_trainer_can_add_student_only_to_assigned_team(self):
        self.client.force_login(self.trainer)

        response = self.client.post(reverse("members"), {
            "action": "create",
            "student_name": "New Player",
            "student_email": "new.player@christuniversity.in",
            "gender": "Female",
            "sport": self.assigned_team.sport_id,
            "team_type": self.assigned_team.team_type,
            "team": self.assigned_team.pk,
            "is_active": "on",
        })

        self.assertRedirects(response, reverse("members"))
        self.assertTrue(Membership.objects.filter(user__email="new.player@christuniversity.in", team=self.assigned_team).exists())
        self.assertEqual(User.objects.get(email="new.player@christuniversity.in").profile.gender, "Female")

    def test_trainer_cannot_add_student_to_other_team(self):
        self.client.force_login(self.trainer)

        response = self.client.post(reverse("members"), {
            "action": "create",
            "student_name": "Blocked Player",
            "student_email": "blocked.player@christuniversity.in",
            "team": self.other_team.pk,
            "is_active": "on",
        })

        self.assertRedirects(response, reverse("members"))
        self.assertFalse(Membership.objects.filter(user__email="blocked.player@christuniversity.in", team=self.other_team).exists())

    def test_trainer_can_deactivate_and_delete_assigned_membership(self):
        self.client.force_login(self.trainer)

        self.client.post(reverse("members"), {
            "action": "deactivate",
            "membership_id": self.assigned_membership.pk,
        })
        self.assigned_membership.refresh_from_db()
        self.assertFalse(self.assigned_membership.is_active)

        self.client.post(reverse("members"), {
            "action": "delete",
            "membership_id": self.assigned_membership.pk,
        })
        self.assertFalse(Membership.objects.filter(pk=self.assigned_membership.pk).exists())

    def test_trainer_cannot_modify_other_team_membership(self):
        self.client.force_login(self.trainer)

        response = self.client.post(reverse("members"), {
            "action": "deactivate",
            "membership_id": self.other_membership.pk,
        })

        self.assertRedirects(response, reverse("members"))
        self.other_membership.refresh_from_db()
        self.assertTrue(self.other_membership.is_active)

    def test_trainer_student_detail_is_restricted_to_assigned_team_records(self):
        session = Session.objects.create(
            team=self.assigned_team,
            title="Practice session",
            start_at=timezone.now(),
            end_at=timezone.now(),
            venue="Ground A",
            scheduled_by=self.trainer,
        )
        other_session = Session.objects.create(
            team=self.other_team,
            title="Other practice",
            start_at=timezone.now(),
            end_at=timezone.now(),
            venue="Ground B",
            scheduled_by=self.other_trainer,
        )
        AttendanceRecord.objects.create(session=session, member=self.assigned_student, status=AttendanceRecord.Status.PRESENT, marked_by=self.trainer)
        AttendanceRecord.objects.create(session=other_session, member=self.assigned_student, status=AttendanceRecord.Status.ABSENT, marked_by=self.other_trainer)
        Membership.objects.create(user=self.assigned_student, team=self.other_team)
        self.client.force_login(self.trainer)

        response = self.client.get(reverse("member_detail", args=[self.assigned_student.pk]))

        self.assertContains(response, "Practice session")
        self.assertContains(response, "100.0%")
        self.assertNotContains(response, "Other practice")

    def test_trainer_can_submit_attendance_for_assigned_team_session(self):
        session = Session.objects.create(
            team=self.assigned_team,
            title="Attendance Practice",
            start_at=timezone.now(),
            end_at=timezone.now(),
            venue="Ground A",
            scheduled_by=self.trainer,
        )
        self.client.force_login(self.trainer)

        response = self.client.post(reverse("take_attendance", args=[session.pk]), {
            f"status_{self.assigned_student.pk}": AttendanceRecord.Status.PRESENT,
            f"remarks_{self.assigned_student.pk}": "On time",
        })

        self.assertRedirects(response, reverse("attendance_detail", args=[session.pk]))
        session.refresh_from_db()
        self.assertTrue(session.attendance_submitted)
        self.assertEqual(session.submitted_by, self.trainer)
        record = AttendanceRecord.objects.get(session=session, member=self.assigned_student)
        self.assertEqual(record.status, AttendanceRecord.Status.PRESENT)
        self.assertEqual(record.remarks, "On time")
        self.assertEqual(record.session.team, self.assigned_team)
        self.assertEqual(record.session.team.sport, self.sport)

    def test_student_bulk_upload_supports_gender_sport_and_team_type(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Student Name", "Student Email", "Mobile Number", "Reg No", "Department", "Class", "Gender", "Sport", "Team Type", "Team", "Status"])
        sheet.append(["Bulk Player", "bulk.player@christuniversity.in", "9876543210", "B001", "MSDS", "I MSc", "Male", "Cricket", "University", "University Team", "Active"])
        uploaded = BytesIO()
        workbook.save(uploaded)
        uploaded.seek(0)

        created, errors = import_students_from_file(uploaded, self.trainer)

        self.assertEqual(created, 1)
        self.assertEqual(errors, [])
        student = User.objects.get(email="bulk.player@christuniversity.in")
        self.assertEqual(student.profile.gender, "Male")
        self.assertTrue(Membership.objects.filter(user=student, team=self.assigned_team, is_active=True).exists())

    @patch("core.views.create_notifications")
    def test_session_bulk_upload_matches_all_four_team_fields(self, mocked_notifications):
        target_team = Team.objects.create(
            sport=self.sport,
            name="Shared Team",
            team_type=Team.TeamType.RESERVE,
            gender=Team.TeamGender.FEMALE,
            coordinator=self.trainer,
        )
        Team.objects.create(
            sport=self.sport,
            name="Shared Team",
            team_type=Team.TeamType.UNIVERSITY,
            gender=Team.TeamGender.FEMALE,
            coordinator=self.trainer,
        )
        Team.objects.create(
            sport=self.sport,
            name="Shared Team",
            team_type=Team.TeamType.RESERVE,
            gender=Team.TeamGender.MALE,
            coordinator=self.trainer,
        )
        Team.objects.create(
            sport=self.other_sport,
            name="Shared Team",
            team_type=Team.TeamType.RESERVE,
            gender=Team.TeamGender.FEMALE,
            coordinator=self.trainer,
        )
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Sport", "Team Name", "Team Type", "Gender", "Title", "Start Date", "End Date", "From Time", "To Time", "Schedule", "Full Day", "Venue", "Other Venue", "Notes"])
        sheet.append(["Cricket", "Shared Team", "Reserve", "Female", "Exact Team Practice", "2026-11-01", "2026-11-01", "06:30", "08:30", "Morning", "No", "Ground A", "", "Drills"])
        uploaded = BytesIO()
        workbook.save(uploaded)
        uploaded.seek(0)

        created, errors = import_sessions_from_file(uploaded, self.trainer)

        self.assertEqual(created, 1)
        self.assertEqual(errors, [])
        self.assertEqual(Session.objects.get(title="Exact Team Practice").team, target_team)
        mocked_notifications.assert_called_once()

    def test_session_bulk_upload_reports_missing_and_invalid_team_details(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Sport", "Team Name", "Team Type", "Gender", "Title", "Start Date", "End Date", "From Time", "To Time", "Schedule", "Full Day", "Venue", "Other Venue", "Notes"])
        sheet.append(["Cricket", "", "University", "Male", "Missing Team", "2026-11-01", "2026-11-01", "06:30", "08:30", "Morning", "No", "Ground A", "", ""])
        sheet.append(["Cricket", "University Team", "Invalid Type", "Male", "Invalid Type", "2026-11-01", "2026-11-01", "06:30", "08:30", "Morning", "No", "Ground A", "", ""])
        sheet.append(["Cricket", "University Team", "University", "Unknown", "Invalid Gender", "2026-11-01", "2026-11-01", "06:30", "08:30", "Morning", "No", "Ground A", "", ""])
        uploaded = BytesIO()
        workbook.save(uploaded)
        uploaded.seek(0)

        created, errors = import_sessions_from_file(uploaded, self.trainer)

        self.assertEqual(created, 0)
        self.assertEqual(len(errors), 3)
        self.assertIn("Team Name required", errors[0])
        self.assertIn("Team Type 'Invalid Type' is invalid", errors[1])
        self.assertIn("Gender 'Unknown' is invalid", errors[2])
