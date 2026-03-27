VERSION 5.00
Begin VB.Form frmOrcamentoSolucoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XSol Aquecedores - Soluções do orcamento"
   ClientHeight    =   2340
   ClientLeft      =   4365
   ClientTop       =   3480
   ClientWidth     =   7800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2250
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   7725
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   6615
         TabIndex        =   10
         Top             =   1755
         Width           =   960
      End
      Begin VB.TextBox txtValTotal 
         Height          =   315
         Left            =   1110
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   9
         Top             =   1695
         Width           =   720
      End
      Begin VB.TextBox txtValUnitario 
         Height          =   315
         Left            =   1110
         MaxLength       =   200
         TabIndex        =   7
         Top             =   1215
         Width           =   720
      End
      Begin VB.TextBox txtQtd 
         Height          =   315
         Left            =   1095
         MaxLength       =   200
         TabIndex        =   5
         Top             =   735
         Width           =   720
      End
      Begin VB.CommandButton btnAdicionar 
         Caption         =   "Adicionar solução"
         Height          =   300
         Left            =   4950
         TabIndex        =   3
         Top             =   1755
         Width           =   1470
      End
      Begin VB.TextBox txtDescricao 
         Height          =   315
         Left            =   1110
         MaxLength       =   200
         TabIndex        =   2
         Top             =   255
         Width           =   6465
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor Total:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1785
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Val. unitário:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   1320
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Qtde:"
         Height          =   195
         Left            =   660
         TabIndex        =   4
         Top             =   840
         Width           =   390
      End
      Begin VB.Label Lbl_Registro 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   285
         TabIndex        =   1
         Top             =   360
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmOrcamentoSolucoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public blnEditando As Boolean
Public strNumeroOrcamento
Public strDescricaoSolucao As String

Private Sub btnAdicionar_Click()
    
    If Trim(txtDescricao.Text) = "" Then
        MsgBox "Informe uma descirção para a solução."
        txtDescricao.SetFocus
        Exit Sub
    End If
    

    If Trim(txtQtd.Text) = "" Then txtQtd.Text = "0"
    If Trim(txtValUnitario.Text) = "" Then txtValUnitario.Text = "0"
    If Trim(txtValTotal.Text) = "" Then txtValTotal.Text = "0"
    
    frmOrcamento.OrcamentoSolucaoDescricao = txtDescricao.Text
    frmOrcamento.OrcamentoSolucaoQtd = txtQtd.Text
    frmOrcamento.OrcamentoSolucaoValUnitario = txtValUnitario.Text
    frmOrcamento.OrcamentoSolucaoValTotal = txtValTotal.Text
    
    Unload Me
    
End Sub

Private Sub btnCancelar_Click()
    frmOrcamento.bln_cancelado = True
    Unload Me
End Sub

Private Sub Form_Activate()
    
    
    txtDescricao.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    txtQtd.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    txtValUnitario.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    txtValTotal.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 180 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtQtd_Change()

    If (IsNumeric(txtQtd.Text)) And (IsNumeric(txtValUnitario.Text)) Then
        If frmOrcamentoSolucoes.Visible Then
            If Trim(txtValUnitario.Text) <> "" And Trim(txtQtd.Text) <> "" And Trim(txtQtd.Text) <> "0" Then
                txtValTotal.Text = Format(CDbl(txtQtd.Text) * CDbl(txtValUnitario.Text), "###,###,##0.00")
            Else
                txtValTotal.Text = ""
            End If
        End If
    End If
    
End Sub

Private Sub txtQtd_KeyPress(KeyAscii As Integer)
    
 On Error GoTo erro
    
    If KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 46 And KeyAscii <> 48 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 56 And KeyAscii <> 57 Then
        KeyAscii = 0
        Exit Sub
    End If
    

    Exit Sub
erro:
    MsgBox "Favor verificar o valor digitado!"
    KeyAscii = 0
    
'txtQtd.Text = Format(txtQtd, "###,###,##0.00")

End Sub

Private Sub txtValUnitario_Change()
    
    If IsNumeric(txtQtd.Text) And (IsNumeric(txtValUnitario.Text)) Then
        If frmOrcamentoSolucoes.Visible Then
            If Trim(txtQtd.Text) <> "" And Trim(txtQtd.Text) <> "0" And Trim(txtValUnitario.Text) <> "" And Trim(txtValUnitario.Text) <> "-" Then
                txtValTotal.Text = Format(CDbl(txtQtd.Text) * CDbl(txtValUnitario.Text), "###,###,##0.00")
            Else
                txtValTotal.Text = ""
            End If
        End If
    End If
End Sub

Private Sub txtValUnitario_KeyPress(KeyAscii As Integer)
 On Error GoTo erro
    
    If KeyAscii = 13 Then
        btnAdicionar_Click
    End If
    
    
    If KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 48 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 56 And KeyAscii <> 57 Then
        KeyAscii = 0
        Exit Sub
    End If
    

    Exit Sub
erro:
    MsgBox "Favor verificar o valor digitado!"
    KeyAscii = 0
    
End Sub

Private Function RetornaSolucao() As Recordset
    Set RetornaSolucao = Return_Recordset("Select * from orcamento_solucao where cod_orcamento = " & strNumeroOrcamento & " and descricao = '" & strDescricaoSolucao & "'")
End Function

