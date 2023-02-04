# Generated by Django 4.0.5 on 2022-10-27 18:59

from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='Contas',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('dia', models.IntegerField(blank=True, null=True)),
                ('descrição', models.CharField(blank=True, max_length=60, null=True)),
                ('símbolo', models.CharField(blank=True, max_length=2, null=True)),
                ('donativos_Entrada', models.DecimalField(blank=True, decimal_places=2, max_digits=10, null=True)),
                ('donativos_Saída', models.DecimalField(blank=True, decimal_places=2, max_digits=10, null=True)),
                ('conta_Entrada', models.DecimalField(blank=True, decimal_places=2, max_digits=10, null=True)),
                ('conta_Saída', models.DecimalField(blank=True, decimal_places=2, max_digits=10, null=True)),
                ('outra_Entrada', models.DecimalField(blank=True, decimal_places=2, max_digits=10, null=True)),
                ('outra_Saída', models.DecimalField(blank=True, decimal_places=2, max_digits=10, null=True)),
            ],
        ),
        migrations.CreateModel(
            name='Gerais',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('congregação', models.CharField(blank=True, max_length=60, null=True)),
                ('cidade', models.CharField(blank=True, max_length=60, null=True)),
                ('estado', models.CharField(blank=True, max_length=60, null=True)),
                ('mês', models.CharField(blank=True, max_length=60, null=True)),
                ('ano', models.CharField(blank=True, max_length=60, null=True)),
                ('data_do_Fechamento', models.DateField(blank=True, null=True)),
                ('saldo_Final_do_Extrato_Mensal', models.DecimalField(blank=True, decimal_places=2, max_digits=10, null=True)),
                ('remessa_Enviada_para_Betel_Resolução', models.DecimalField(blank=True, decimal_places=2, max_digits=10, null=True)),
                ('saldo_Final_do_Extrato_de_Betel', models.DecimalField(blank=True, decimal_places=2, max_digits=10, null=True)),
                ('saldo_Final_dos_Donativos_Mês_Anterior', models.DecimalField(blank=True, decimal_places=2, max_digits=10, null=True)),
                ('saldo_Final_da_Conta_Bancária_Mês_Anterior', models.DecimalField(blank=True, decimal_places=2, max_digits=10, null=True)),
                ('saldo_Final_da_Conta_em_Betel_Mês_Anterior', models.DecimalField(blank=True, decimal_places=2, max_digits=10, null=True)),
            ],
        ),
    ]
