VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_Xsol_03 
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   3300
   ClientTop       =   3300
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4995
   Begin VB.CommandButton Cmd_Cancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Cmd_Acessar 
      Caption         =   "&Acessar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text_Senha 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   960
      MaxLength       =   8
      PasswordChar    =   "§"
      TabIndex        =   1
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox Text_Usuario 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   960
      MaxLength       =   15
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3030
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            TextSave        =   "10:59"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            TextSave        =   "22/3/2004"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   120
      Picture         =   "Form_Xsol_03.frx":0000
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   240
      Picture         =   "Form_Xsol_03.frx":1CCA
      Top             =   480
      Width           =   510
   End
End
Attribute VB_Name = "Form_Xsol_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public W_DB As Connection
Public Tabela_Usuarios As Recordset
Public W_DBname As String
Private Sub Cmd_Acessar_Click()

Dim Aux_Usuario As String
Dim Aux_Senha As String

Dim X_Usuario As String
Dim X_Senha As String


Aux_Usuario = Trim(UCase(Text_Usuario.Text))
Aux_Senha = Trim(UCase(Text_Senha.Text))

X_Usuario = ""
X_Senha = ""

Tabela_Usuarios.MoveFirst

X_Usuario = Trim(UCase(Tabela_Usuarios.Fields("Usuario")))
X_Senha = Trim(UCase(Tabela_Usuarios.Fields("Senha")))

If Aux_Usuario = X_Usuario And Aux_Senha = X_Senha Then
   Inicializar
Else
   Do While (Not Tabela_Usuarios.EOF) And _
   (Aux_Usuario <> X_Usuario And Aux_Senha <> X_Senha)
      DoEvents
      X_Usuario = Trim(UCase(Tabela_Usuarios.Fields("Usuario")))
      X_Senha = Trim(UCase(Tabela_Usuarios.Fields("Senha")))
      Tabela_Usuarios.MoveNext
   Loop
   If Aux_Usuario = X_Usuario And Aux_Senha = X_Senha Then
      Inicializar
   End If
End If

Text_Usuario.Text = ""
Text_Senha.Text = ""

If Aux_Usuario <> X_Usuario And Aux_Senha <> X_Senha Then
   Text_Usuario.SetFocus
   MDIForm_Xsol.nlogon.Enabled = True
   Exit Sub
End If

Unload Me

End Sub

Private Sub Cmd_Cancelar_Click()

Unload Me

End Sub

Private Sub Form_Load()

Dim hMenu As Long
Dim lItemCount As Long

hMenu = GetSystemMenu(Me.hWnd, 0)
If hMenu Then
   lItemCount = GetMenuItemCount(hMenu)
   Call RemoveMenu(hMenu, lItemCount - 1, MF_REMOVE Or _
   MF_BYPOSITION)
   Call RemoveMenu(hMenu, lItemCount - 2, MF_REMOVE Or _
   MF_BYPOSITION)
   Call DrawMenuBar(Me.hWnd)
End If

Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

If Right(App.Path, 1) = "\" Then
  W_DBname = App.Path & "dados.mdb"
 Else
  W_DBname = App.Path & "\" & "dados.mdb"
End If

Set W_DB = New Connection

Set Tabela_Usuarios = New Recordset

W_DB.Mode = adModeReadWrite

W_DB.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"

Tabela_Usuarios.Open "select * from Usuarios Order By Usuario ", W_DB, adOpenStatic, adLockOptimistic

MDIForm_Xsol.nclientes.Enabled = False
MDIForm_Xsol.ntecnico.Enabled = False

End Sub

Private Sub Form_Terminate()

W_DB.Close
Tabela_Usuarios.Close

End Sub

Sub Inicializar()

MDIForm_Xsol.Caption = "Usuário logado: " & Text_Usuario.Text

MDIForm_Xsol.nclientes.Enabled = True
MDIForm_Xsol.nlogon.Enabled = False
MDIForm_Xsol.ntecnico.Enabled = True
MDIForm_Xsol.Arrange vbTileHorizontal

End Sub
