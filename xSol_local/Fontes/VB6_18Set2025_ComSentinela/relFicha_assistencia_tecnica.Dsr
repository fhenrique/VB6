VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} relFicha_Assistencia_Tecnica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project1 - relFicha_Assistencia_Tecnica (ActiveReport)"
   ClientHeight    =   10575
   ClientLeft      =   0
   ClientTop       =   345
   ClientWidth     =   20490
   Icon            =   "relFicha_assistencia_tecnica.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   WindowState     =   2  'Maximized
   _ExtentX        =   36142
   _ExtentY        =   18653
   SectionData     =   "relFicha_assistencia_tecnica.dsx":08CA
End
Attribute VB_Name = "relFicha_Assistencia_Tecnica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()

    
    '0 equipamento e 1 serviço
    If (Form_Service.qtdEquipamentos = 0) And (Form_Service.qtdServicos = 1) Then
        lblViewEquipamento1.Visible = False
        lblViewEquipamento2.Visible = False
        lblViewEquipamento3.Visible = False
        
        lblEquipamento1.Visible = False
        lblEquipamento2.Visible = False
        lblEquipamento3.Visible = False
        
        lblViewTipoServico1.Visible = True
        lblTipoServico1.Visible = True
        lblViewTipoServico2.Visible = False
        lblTipoServico2.Visible = False
        
        lblViewEquipamento1.Top = 0
        Exit Sub
    End If
    
    '0 equipamento e 2 serviços
    If (Form_Service.qtdEquipamentos = 0) And (Form_Service.qtdServicos = 2) Then
        lblViewEquipamento1.Visible = False
        lblViewEquipamento2.Visible = False
        lblViewEquipamento3.Visible = False
        
        lblEquipamento1.Visible = False
        lblEquipamento2.Visible = False
        lblEquipamento3.Visible = False
        
        
        lblViewTipoServico1.Visible = True
        lblViewTipoServico2.Visible = True
        
        lblTipoServico1.Visible = True
        lblTipoServico2.Visible = True
        
        lblViewTipoServico1.Top = 2500
        lblTipoServico1.Top = 2500
        
        lblViewTipoServico2.Top = 2800
        lblTipoServico2.Top = 2800
        
        Exit Sub
    End If
        
        
    '1 equipamento e 1 serviço
    If (Form_Service.qtdEquipamentos = 1) And (Form_Service.qtdServicos = 1) Then
        lblViewEquipamento1.Visible = True
        lblViewEquipamento2.Visible = False
        lblViewEquipamento3.Visible = False
        
        lblEquipamento1.Visible = True
        lblEquipamento2.Visible = False
        lblEquipamento3.Visible = False
        
        lblViewTipoServico1.Visible = True
        lblViewTipoServico2.Visible = False
        
        lblTipoServico1.Visible = True
        lblTipoServico2.Visible = False
        Exit Sub
    End If
    
    '1 equipamento e 2 serviços
    If (Form_Service.qtdEquipamentos = 1) And (Form_Service.qtdServicos = 2) Then
        lblViewEquipamento1.Visible = True
        lblViewEquipamento2.Visible = False
        lblViewEquipamento3.Visible = False
        
        lblEquipamento1.Visible = True
        lblEquipamento2.Visible = False
        lblEquipamento3.Visible = False
        
        lblViewTipoServico1.Visible = True
        lblViewTipoServico2.Visible = True
        
        lblTipoServico1.Visible = True
        lblTipoServico2.Visible = True
        
        lblViewTipoServico1.Top = 2750
        lblTipoServico1.Top = 2750
        
        lblViewTipoServico2.Top = 3000
        lblTipoServico2.Top = 3000
        
        Exit Sub
    End If
        
    '2 equipamentos e 1 serviço
    If (Form_Service.qtdEquipamentos = 2) And (Form_Service.qtdServicos = 1) Then
        lblViewEquipamento1.Visible = True
        lblViewEquipamento2.Visible = True
        lblViewEquipamento3.Visible = False
        
        lblEquipamento1.Visible = True
        lblEquipamento2.Visible = True
        lblEquipamento3.Visible = False
        
        lblViewTipoServico1.Visible = True
        lblViewTipoServico2.Visible = False
        
        lblTipoServico1.Visible = True
        lblTipoServico2.Visible = False
        
        lblViewEquipamento1.Top = 2400
        lblEquipamento1.Top = 2400
        
        lblViewTipoServico1.Top = 2850
        lblTipoServico1.Top = 2850
        
        Exit Sub
    End If
    
    '2 equipamentos e 2 serviços
    If (Form_Service.qtdEquipamentos = 2) And (Form_Service.qtdServicos = 2) Then
        lblViewEquipamento1.Visible = True
        lblViewEquipamento2.Visible = True
        lblViewEquipamento3.Visible = False
        
        lblEquipamento1.Visible = True
        lblEquipamento2.Visible = True
        lblEquipamento3.Visible = False
        
        lblViewTipoServico1.Visible = True
        lblViewTipoServico2.Visible = True
        
        lblTipoServico1.Visible = True
        lblTipoServico2.Visible = True
        
        lblViewEquipamento1.Top = 2400
        lblEquipamento1.Top = 2400
        lblViewTipoServico1.Top = 2850
        lblTipoServico1.Top = 2850
        
        lblViewEquipamento2.Top = 2600
        lblEquipamento2.Top = 2600
        
        lblViewTipoServico2.Top = 3050
        lblTipoServico2.Top = 3050
        
        lblViewValorInstalacao.Top = 3600
        lblValorInstalacao.Top = 3600
        
        lblObs.Top = 3350
        
        lblViewRelatorio.Top = 4050
        
        
        Exit Sub
    End If
        
    '3 equipamentos e 1 serviço
    If (Form_Service.qtdEquipamentos = 3) And (Form_Service.qtdServicos = 1) Then
        lblViewEquipamento1.Visible = True
        lblViewEquipamento2.Visible = True
        lblViewEquipamento3.Visible = True
        
        lblEquipamento1.Visible = True
        lblEquipamento2.Visible = True
        lblEquipamento3.Visible = True
        
        lblViewTipoServico1.Visible = True
        lblViewTipoServico2.Visible = False
        
        lblTipoServico1.Visible = True
        lblTipoServico2.Visible = False
        
        lblViewEquipamento1.Top = 2400
        lblEquipamento1.Top = 2400
        
        lblViewTipoServico1.Top = 3100
        lblTipoServico1.Top = 3100
        
        lblViewEquipamento2.Top = 2600
        lblEquipamento2.Top = 2600
        
        lblViewTipoServico2.Top = 3050
        lblTipoServico2.Top = 3050
        
        lblViewValorInstalacao.Top = 3600
        lblValorInstalacao.Top = 3600
        
        lblObs.Top = 3350
        
        lblViewRelatorio.Top = 4050
        Exit Sub
    End If
        
    '3 equipamentos e 2 serviços
    If (Form_Service.qtdEquipamentos = 3) And (Form_Service.qtdServicos = 2) Then
        lblViewEquipamento1.Visible = True
        lblViewEquipamento2.Visible = True
        lblViewEquipamento3.Visible = True
        
        lblEquipamento1.Visible = True
        lblEquipamento2.Visible = True
        lblEquipamento3.Visible = True
        
        lblViewTipoServico1.Visible = True
        lblViewTipoServico2.Visible = True
        
        lblTipoServico1.Visible = True
        lblTipoServico2.Visible = True
        
        lblViewEquipamento1.Top = 2400
        lblEquipamento1.Top = 2400
        
        lblViewTipoServico1.Top = 3100
        lblTipoServico1.Top = 3100
        
        lblViewEquipamento2.Top = 2600
        lblEquipamento2.Top = 2600
        
        lblViewTipoServico2.Top = 3350
        lblTipoServico2.Top = 3350
        
        lblViewValorInstalacao.Top = 3950
        lblValorInstalacao.Top = 3950
        
        lblObs.Top = 3650
        
        lblViewRelatorio.Top = 4250
        Exit Sub
    End If
        
        
End Sub

