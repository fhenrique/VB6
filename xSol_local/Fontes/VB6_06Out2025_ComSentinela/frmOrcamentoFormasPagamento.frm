VERSION 5.00
Begin VB.Form frmOrcamentoFormasPagamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XSol Aquecedores - Formas de pagamento do orçamento"
   ClientHeight    =   1365
   ClientLeft      =   11100
   ClientTop       =   4845
   ClientWidth     =   7800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   7725
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   6555
         TabIndex        =   4
         Top             =   870
         Width           =   960
      End
      Begin VB.CommandButton btnAdicionar 
         Caption         =   "Adicionar garantia"
         Height          =   300
         Left            =   4770
         TabIndex        =   2
         Top             =   870
         Width           =   1605
      End
      Begin VB.TextBox txtDescricao 
         Height          =   315
         Left            =   1050
         MaxLength       =   90
         TabIndex        =   1
         Top             =   345
         Width           =   6465
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
Attribute VB_Name = "frmOrcamentoFormasPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public blnEditando As Boolean
Public strNumeroOrcamento
Public strDescricaoFormaDePagamento As String


Private Sub btnAdicionar_Click()
    frmOrcamento.OrcamentoFormaPagamentoCancelado = False
    frmOrcamento.OrcamentoFormaPagamentoDescricao = txtDescricao.Text
    
    Unload Me

End Sub

Private Sub btnCancelar_Click()
    frmOrcamento.OrcamentoFormaPagamentoCancelado = True
    Unload Me
End Sub

Private Sub Form_Activate()
    txtDescricao.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    If frmOrcamento.bln_editando Then
        txtDescricao.Text = strDescricaoFormaDePagamento
    End If
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 180 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        btnAdicionar_Click
    End If
End Sub
