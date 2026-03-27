VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} relSaidaEquipamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project1 - relSaidaEquipamento (ActiveReport)"
   ClientHeight    =   8235
   ClientLeft      =   2205
   ClientTop       =   1740
   ClientWidth     =   13320
   MaxButton       =   0   'False
   MinButton       =   0   'False
   WindowState     =   2  'Maximized
   _ExtentX        =   23495
   _ExtentY        =   14526
   SectionData     =   "relSaidaEquipamento.dsx":0000
End
Attribute VB_Name = "relSaidaEquipamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim total As Double
Dim totalQtd As Double
Dim totalVal As Double
Dim totalGeral As Double
Dim ProdutoAtual As String

Private Sub Detail_Format()
    If Field8.Text = "" Then
        MsgBox "Sem dados no período informado", vbOKOnly, "xSol Aquecedores"
        Unload Me
        Exit Sub
    End If
    
    
    ProdutoAtual = Field8.Text
    
    total = total + CDbl(Field9.Text)
    txtTotal.Text = Format(Str(total), "R$###,###,##0.00")
    
    totalQtd = totalQtd + CDbl(Field3.Text)
    txtTotalQtd.Text = Format(Str(totalQtd), "###,###,##0.00")
    
    totalVal = totalVal + CDbl(Field4.Text)
    txtTotalValor.Text = Format(Str(totalVal), "R$###,###,##0.00")
End Sub

Private Sub GroupFooter1_Format()
    If txtTotal.Text = "txtTotal" Then
        MsgBox "Sem dados no período informado", vbOKOnly, "xSol Aquecedores"
        Unload Me
        Exit Sub
    End If
    
    total = 0
    totalQtd = 0
    totalVal = 0
    
    totalGeral = totalGeral + CDbl(txtTotal.Text)
    txtTotalGeral.Text = Format(Str(totalGeral), "R$###,###,##0.00")

End Sub
