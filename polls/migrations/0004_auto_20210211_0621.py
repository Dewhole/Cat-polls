# Generated by Django 2.2.12 on 2021-02-11 06:21

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("polls", "0003_auto_20210211_0617"),
    ]

    operations = [
        migrations.AlterField(
            model_name="users",
            name="partyID",
            field=models.PositiveIntegerField(
                blank=True, null=True, verbose_name="Код каталога"
            ),
        ),
    ]
