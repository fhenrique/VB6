VERSION 5.00
Begin VB.Form frmOrcamentoGarantias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XSol Aquecedores - Garantias do orçamento"
   ClientHeight    =   1365
   ClientLeft      =   4380
   ClientTop       =   4005
   ClientWidth     =   7785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   0
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
         Top             =   330
         Width           =   6465
      End
      Begin VB.CommandButton btnAdicionar 
         Caption         =   "Adicionar garantia"
         Height          =   300
         Left            =   4875
         TabIndex        =   1
         Top             =   840
         Width           =   1515
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
Attribute VB_Name = "frmOrcamentoGarantias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public blnEditando As Boolean
Public strNumeroOrcamento
Public strDescricaoGarantia As String
Public cancelou As Boolean


Private Sub btnAdicionar_Click()

    If Trim(txtDescricao.Text) = "" Then
        MsgBox "Informe uma descrição."
        txtDescricao.SetFocus
        Exit Sub
    End If
    
    frmOrcamento.OrcamentoGarantiaDescricao = txtDescricao.Text

    Unload Me
End Sub

Private Sub btnAdicionar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btnAdicionar_Click
    End If
End Sub

Private Sub btnCancelar_Click()
    cancelou = True
    Unload Me
End Sub

Private Sub Form_Activate()
    txtDescricao.SetFocus
    txtDescricao.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    
    If blnEditando Then
        Dim rs_garantia As Recordset
        Set rs_garantia = Return_Recordset("Select * from orcamento_garantias where cod_orcamento = " & strNumeroOrcamento & " and garantia_descricao = '" & strDescricaoGarantia & "'")
        
        txtDescricao.Text = strDescricaoGarantia
    End If
    
End Sub

Private Sub Form_Load()
    cancelou = False
    
    txtDescricao.Text = strDescricaoGarantia
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 180 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        btnAdicionar_Click
    End If
End Sub
