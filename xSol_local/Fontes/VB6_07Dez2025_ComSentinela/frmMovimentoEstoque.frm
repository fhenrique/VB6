VERSION 5.00
Begin VB.Form frmMovimentoEstoque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimentação"
   ClientHeight    =   690
   ClientLeft      =   9135
   ClientTop       =   5490
   ClientWidth     =   3000
   ControlBox      =   0   'False
   Icon            =   "frmMovimentoEstoque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   3000
   Begin VB.Frame Frame1 
      Height          =   750
      Left            =   15
      TabIndex        =   0
      Top             =   -60
      Width           =   2970
      Begin VB.CommandButton btnFechar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2295
         Picture         =   "frmMovimentoEstoque.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancelar"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Saída"
         Height          =   405
         Left            =   1200
         TabIndex        =   2
         Top             =   195
         Width           =   900
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Entrada"
         Height          =   405
         Left            =   180
         TabIndex        =   1
         Top             =   210
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmMovimentoEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnFechar_Click()
        
        Form_Estoque.cancelaMovimento = True
        Unload Me
        
End Sub

Private Sub Command1_Click()

    
    Dim continua As Boolean
    Form_Estoque.cancelaMovimento = False
    Form_Estoque.cancelaMovimento = False
    Form_Estoque.continuaMovimento = True
    
    Form_Estoque.movimentoEntrada = True
    Form_Estoque.movimentoSaida = False
    Unload Me
End Sub

Private Sub Command2_Click()
    If Trim(Form_Estoque.cboVisaoEstoqueMaterial.Text) = "" Then
        MsgBox "Informe um Material", vbOKOnly, "xSOL Aquecedores"
        Unload Me
        Exit Sub
    End If
    
    Dim continua As Boolean
    Form_Estoque.cancelaMovimento = False
    Form_Estoque.continuaMovimento = True
    
    If Trim(Form_Estoque.txtVisaoEstoque_Estoque.Text) = "0" Then
        MsgBox "Sem estoque para o produto informado.", vbOKOnly, "xSol Aquecedores."
        Form_Estoque.continuaMovimento = False
    End If
    
    If Form_Estoque.continuaMovimento Then
        Form_Estoque.movimentoEntrada = False
        Form_Estoque.movimentoSaida = True
        Unload Me
    Else
        Unload Me
    End If
End Sub

