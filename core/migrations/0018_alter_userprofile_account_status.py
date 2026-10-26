from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("core", "0017_meeting_soft_delete"),
    ]

    operations = [
        migrations.AlterField(
            model_name="userprofile",
            name="account_status",
            field=models.CharField(
                choices=[("ACTIVE", "Active"), ("DEACTIVATED", "Inactive"), ("DELETED", "Deleted")],
                default="ACTIVE",
                max_length=20,
            ),
        ),
    ]
