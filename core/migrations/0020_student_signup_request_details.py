from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ("core", "0019_session_cancellation_note"),
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.AddField(model_name="loginaccessrequest", name="department", field=models.CharField(blank=True, max_length=120)),
        migrations.AddField(model_name="loginaccessrequest", name="gender", field=models.CharField(blank=True, max_length=30)),
        migrations.AddField(model_name="loginaccessrequest", name="phone", field=models.CharField(blank=True, max_length=20)),
        migrations.AddField(model_name="loginaccessrequest", name="register_no", field=models.CharField(blank=True, max_length=60)),
        migrations.AddField(
            model_name="loginaccessrequest",
            name="school",
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name="login_requests", to="core.school"),
        ),
        migrations.AddField(
            model_name="loginaccessrequest",
            name="sports",
            field=models.ManyToManyField(blank=True, related_name="student_signup_requests", to="core.sport"),
        ),
        migrations.CreateModel(
            name="StudentSport",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("created_at", models.DateTimeField(auto_now_add=True)),
                ("updated_at", models.DateTimeField(auto_now=True)),
                ("is_active", models.BooleanField(default=True)),
                ("sport", models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name="student_associations", to="core.sport")),
                ("user", models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name="sport_associations", to=settings.AUTH_USER_MODEL)),
            ],
            options={"ordering": ["sport__name"], "unique_together": {("user", "sport")}},
        ),
    ]
