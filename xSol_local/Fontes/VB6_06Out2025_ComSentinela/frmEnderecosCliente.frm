VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Begin VB.Form frmEnderecosCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cliente com mais de um endereço"
   ClientHeight    =   2475
   ClientLeft      =   7080
   ClientTop       =   4245
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6510
   Begin VB.CommandButton But_Ok 
      BackColor       =   &H0080FF80&
      Caption         =   "&OK"
      Height          =   495
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1890
      Width           =   690
   End
   Begin MSFlexGridLib.MSFlexGrid gridEnderecos 
      Height          =   1275
      Left            =   90
      TabIndex        =   1
      Top             =   600
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   2249
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "O cliente possui mais de um endereço, favor selecionar abaixo o endereço onde ocorrerá o atendimento:"
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   6240
   End
End
Attribute VB_Name = "frmEnderecosCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strEnderecoAtendimento As String

Private Sub But_Ok_Click()
    strEnderecoAtendimento = gridEnderecos.TextMatrix(gridEnderecos.RowSel, 0)
    Unload Me
End Sub

Private Sub Form_Activate()
gridEnderecos.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
End Sub

