# Generated by Django 4.0.5 on 2023-02-01 18:21

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('contas', '0008_remove_contas_entrada_extra_and_more'),
    ]

    operations = [
        migrations.AddField(
            model_name='contas',
            name='saldo_final',
            field=models.DecimalField(blank=True, decimal_places=2, max_digits=10, null=True),
        ),
    ]
