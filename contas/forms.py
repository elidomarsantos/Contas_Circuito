from django import forms

from .models import Contas, Gerais

class DateInput(forms.DateInput):
    input_type = 'date'
    



class Form_FContas(forms.ModelForm):

    class Meta:
        model = Contas
        fields = '__all__'
       


class Form_Gerais(forms.ModelForm):

    class Meta:
        model = Gerais
        fields = '__all__'  
        widgets = {
            'data_do_Fechamento': forms.DateInput(format=('%Y-%m-%d'), attrs={'class':'form-control', 'placeholder':'Select Date','type': 'date'})
            
        }        
          

