from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("core", "0018_alter_userprofile_account_status"),
    ]

    operations = [
        migrations.AddField(
            model_name="session",
            name="cancellation_note",
            field=models.TextField(blank=True),
        ),
    ]
