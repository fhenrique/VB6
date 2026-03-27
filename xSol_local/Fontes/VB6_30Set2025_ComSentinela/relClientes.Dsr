VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} relClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project1 - relClientes (ActiveReport)"
   ClientHeight    =   10575
   ClientLeft      =   -20490
   ClientTop       =   345
   ClientWidth     =   20490
   MaxButton       =   0   'False
   MinButton       =   0   'False
   WindowState     =   2  'Maximized
   _ExtentX        =   36142
   _ExtentY        =   18653
   SectionData     =   "relClientes.dsx":0000
End
Attribute VB_Name = "relClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim totalQtd As Double

Private Sub ActiveReport_ReportStart()
    
    If lblTipoRelatorio.Caption = "Arquivo Morto" Then
        'lblJustificativa.Visible = True
        'Field14.Visible = True
    Else
        'lblJustificativa.Visible = False
        'Field14.Visible = False
    End If

End Sub

Private Sub Detail_BeforePrint()
    
    totalQtd = totalQtd + 1
'    txtTotal.Text = Format(Str(total), "R$###,###,##0.00")
'    totalQtd = totalQtd + CDbl(Field3.Text)
'    txtTotalQtd.Text = Format(Str(totalQtd), "###,###,##0.00")
'    totalVal = totalVal + CDbl(Field4.Text)
    txtQtdClientes.Text = Str(totalQtd)
    
End Sub

