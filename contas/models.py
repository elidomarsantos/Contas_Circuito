from django.db import models


class Contas(models.Model):
   
    dia = models.IntegerField(blank=True, null=True)
    descrição = models.CharField(max_length=60, blank=True, null=True)
    símbolo = models.CharField(max_length=2, blank=True, null=True)
    donativos_Entrada = models.DecimalField(blank=True, null=True, max_digits = 10, decimal_places = 2)
    donativos_Saída = models.DecimalField(blank=True, null=True, max_digits = 10, decimal_places = 2)
    conta_Entrada = models.DecimalField(blank=True, null=True, max_digits = 10, decimal_places = 2)
    conta_Saída = models.DecimalField(blank=True, null=True, max_digits = 10, decimal_places = 2)
    outra_Entrada = models.DecimalField(blank=True, null=True, max_digits = 10, decimal_places = 2)
    outra_Saída = models.DecimalField(blank=True, null=True, max_digits = 10, decimal_places = 2)
    saldo_final = models.DecimalField(blank=True, null=True, max_digits = 10, decimal_places = 2)
    

class Meta:
    ordering = ['dia']

class Gerais(models.Model):   
    congregação = models.CharField(max_length=60, blank=True, null=True)
    cidade = models.CharField(max_length=60, blank=True, null=True)
    estado = models.CharField(max_length=60, blank=True, null=True)
    mês = models.CharField(max_length=60, blank=True, null=True)
    ano = models.CharField(max_length=60, blank=True, null=True)
    data_do_Fechamento = models.DateField(blank=True, null=True,)
    servo_de_contas = models.CharField(max_length=60, blank=True, null=True)
    secretário = models.CharField(max_length=60, blank=True, null=True)
    saldo_Final_do_Extrato_Mensal = models.DecimalField(blank=True, null=True, max_digits = 10, decimal_places = 2, verbose_name='Saldo da Conta Bancária no Fim do Mês')
    remessa_Enviada_para_Betel_Resolução = models.DecimalField(blank=True, null=True, max_digits = 10, decimal_places = 2, verbose_name='Envio Fixo de Remessa para Betel ')
    saldo_Final_do_Extrato_de_Betel = models.DecimalField(blank=True, null=True, max_digits = 10, decimal_places = 2, verbose_name='Saldo da Conta em Betel no Fim do Mês')
    saldo_Final_dos_Donativos_Mês_Anterior = models.DecimalField(blank=True, null=True, max_digits = 10, decimal_places = 2, verbose_name='Saldo Donativos Mês Anterior' )
    saldo_Final_da_Conta_Bancária_Mês_Anterior = models.DecimalField(blank=True, null=True, max_digits = 10, decimal_places = 2, verbose_name='Saldo da Conta Bancária no Mês Anterior	')
    saldo_Final_da_Conta_em_Betel_Mês_Anterior = models.DecimalField(blank=True, null=True, max_digits = 10, decimal_places = 2, verbose_name='Saldo da Conta em Betel no Mês Anterior	')

