from django.urls import path

from . import views



urlpatterns = [
    path('contas/', views.contas),
    path('editar_fc/<int:id>', views.editar_fc),
    path('deletar_fc/<int:id>', views.deletar_fc),
    path('view_gerais_contas/', views.view_gerais_contas),
    path('add_gerais_contas/', views.add_gerais_contas),
    path('editar_gerais_contas/<int:id>', views.editar_gerais_contas),
    path('deletar_gerais_contas/', views.deletar_gerais_contas),
    path('deletar_fc/', views.deletar_folha_de_contas),
    path('relatório_mensal/', views.relatório_mensal),
    path('registro/', views.registro),
    path('recibo/<int:id>', views.recibo),
    path('imprimir_FC', views.imprimir_FC),
    path('', views.home),
    path('resultado/', views.resultado),
    
    

]
    