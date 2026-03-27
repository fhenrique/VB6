VERSION 5.00
Begin VB.Form frmOrcamentoObservacoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XSol Aquecedores - Observações do orçamento"
   ClientHeight    =   1410
   ClientLeft      =   5325
   ClientTop       =   3585
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7725
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   6570
         TabIndex        =   4
         Top             =   840
         Width           =   960
      End
      Begin VB.TextBox txtDescricao 
         Height          =   315
         Left            =   1050
         MaxLength       =   200
         TabIndex        =   2
         Top             =   345
         Width           =   6465
      End
      Begin VB.CommandButton btnAdicionar 
         Caption         =   "Adicionar observação"
         Height          =   300
         Left            =   4620
         TabIndex        =   1
         Top             =   840
         Width           =   1770
      End
      Begin VB.Label Lbl_Registro 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Top             =   450
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmOrcamentoObservacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public blnEditando As Boolean
Public strNumeroOrcamento
Public strDescricaoObservacao As String


Private Sub btnAdicionar_Click()
    frmOrcamento.OrcamentoObservacaoDescricao = txtDescricao.Text
    
    Unload Me

End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    txtDescricao.SetFocus
    txtDescricao.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 180 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        btnAdicionar_Click
    End If
End Sub
