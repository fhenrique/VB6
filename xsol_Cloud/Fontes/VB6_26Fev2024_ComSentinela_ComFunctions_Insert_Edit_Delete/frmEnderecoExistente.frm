VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Begin VB.Form frmEnderecoExistente 
   Caption         =   "Endereço existente"
   ClientHeight    =   2640
   ClientLeft      =   8925
   ClientTop       =   2175
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSair 
      BackColor       =   &H80000000&
      Caption         =   "&Sair"
      Height          =   495
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2055
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid_Address 
      Height          =   1890
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   3334
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmEnderecoExistente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Grid_Address.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
End Sub

