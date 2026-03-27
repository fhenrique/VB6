VERSION 5.00
Begin VB.Form Form_Login 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Identificação de usuário"
   ClientHeight    =   3870
   ClientLeft      =   6090
   ClientTop       =   3375
   ClientWidth     =   5400
   Icon            =   "Form_Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   15
      TabIndex        =   5
      Top             =   -60
      Width           =   5385
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "XSOL Aquecedores - Login"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   6
         Top             =   135
         Width           =   5115
      End
   End
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
      Left            =   2880
      TabIndex        =   3
      Top             =   2700
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
      Left            =   495
      TabIndex        =   2
      Top             =   2700
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
      Left            =   1200
      MaxLength       =   8
      PasswordChar    =   "§"
      TabIndex        =   1
      Top             =   1875
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
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   0
      Top             =   915
      Width           =   3735
   End
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   5340
      TabIndex        =   4
      Top             =   3495
      Width           =   5400
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   465
      Picture         =   "Form_Login.frx":000C
      Stretch         =   -1  'True
      Top             =   1770
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   555
      Picture         =   "Form_Login.frx":114B
      Stretch         =   -1  'True
      Top             =   810
      Width           =   555
   End
End
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public W_DB As Connection
Public Tabela_Usuarios As Recordset

Private Sub Cmd_Acessar_Click()

Dim Aux_Usuario As String
Dim Aux_Senha As String

Dim X_Usuario As String
Dim X_Senha As String
Dim X_Perfil As String


If MDIForm_Xsol.AtaqueSentinela = True Then
   MsgBox "Unexpected exception ocurred during operation ", vbCritical, " Critical error"
   MDIForm_Xsol.nlogon.enabled = True
   Exit Sub
End If

Aux_Usuario = Trim(UCase(Text_Usuario.Text))
Aux_Senha = Trim(UCase(Text_Senha.Text))

X_Usuario = ""
X_Senha = ""

Tabela_Usuarios.MoveFirst

X_Usuario = Trim(UCase(Tabela_Usuarios.Fields("Usuario")))
X_Senha = Trim(UCase(Tabela_Usuarios.Fields("Senha")))
X_Perfil = Trim(UCase(Tabela_Usuarios.Fields("Perfil")))

If Aux_Usuario = X_Usuario And Aux_Senha = X_Senha Then
   Call Inicializar(X_Usuario, X_Perfil)
Else
   Do While (Not Tabela_Usuarios.EOF) And _
   (Aux_Usuario <> X_Usuario And Aux_Senha <> X_Senha)
      DoEvents
      X_Usuario = Trim(UCase(Tabela_Usuarios.Fields("Usuario")))
      X_Senha = Trim(UCase(Tabela_Usuarios.Fields("Senha")))
      X_Perfil = Trim(UCase(Tabela_Usuarios.Fields("Perfil")))
      Tabela_Usuarios.MoveNext
   Loop
   If Aux_Usuario = X_Usuario And Aux_Senha = X_Senha Then
      Call Inicializar(X_Usuario, X_Perfil)
   End If
End If

Text_Usuario.Text = ""
Text_Senha.Text = ""
Tabela_Usuarios.Close

If Aux_Usuario <> X_Usuario And Aux_Senha <> X_Senha Then
   Text_Usuario.SetFocus
   MDIForm_Xsol.nlogon.enabled = True
   Exit Sub
End If

If Aux_Usuario = X_Usuario And Aux_Senha = X_Senha Then
    MDIForm_Xsol.usuarioLogado = Aux_Usuario
End If



Unload Me

End Sub

Private Sub Cmd_Cancelar_Click()

Unload Me

End Sub

Private Sub Form_Activate()
Text_Usuario.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Senha.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
End Sub

Private Sub Form_Load()

    Dim hMenu As Long
    Dim lItemCount As Long
    
    'Text_Usuario.Text = "xsol"
    'Text_Senha.Text = "xsol"
    
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
    
    Call Arquivos_Dados
    Call Set_Connection
    
    If W_DBname = "" Then
       MsgBox "Não encontrado o arquivo CONFIG.INI ou a configuração no arquivo CONFIG.INI", vbOKOnly, "Xsol Aquecedores"
       Exit Sub
    End If
    
    Set W_DB = New Connection
    
    Set Tabela_Usuarios = New Recordset
    
    W_DB.Mode = adModeReadWrite
    
    W_DB.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
    
    Tabela_Usuarios.Open "select * from Usuarios Order By Usuario ", W_DB, adOpenStatic, adLockOptimistic
    
    MDIForm_Xsol.nclientes.enabled = False
    MDIForm_Xsol.ncadastros.enabled = False
    MDIForm_Xsol.nuser.enabled = False
    MDIForm_Xsol.nuser.Visible = False
      MDIForm_Xsol.nservtype.enabled = False
    MDIForm_Xsol.ntecnico.enabled = False
    MDIForm_Xsol.nzone.enabled = False
    MDIForm_Xsol.nAgenda.enabled = False
    MDIForm_Xsol.nVisitas.enabled = False
    MDIForm_Xsol.nconfig.enabled = False
    MDIForm_Xsol.nconfig.Visible = False

End Sub

Private Sub Form_Terminate()

W_DB.Close
Tabela_Usuarios.Close

End Sub

Sub Inicializar(Login As String, Profile As String)
    
    MDIForm_Xsol.Caption = "Usuário logado: " & Text_Usuario.Text & " | " & versaoAtual
    
    MDIForm_Xsol.nclientes.enabled = True
    MDIForm_Xsol.nlogon.enabled = False
    MDIForm_Xsol.ncadastros.enabled = True
    MDIForm_Xsol.nuser.enabled = (Profile = ADMIN_PROFILE)
    MDIForm_Xsol.nuser.Visible = (Profile = ADMIN_PROFILE)
    MDIForm_Xsol.nservtype.enabled = True
    MDIForm_Xsol.ntecnico.enabled = True
    MDIForm_Xsol.nzone.enabled = True
    MDIForm_Xsol.nAgenda.enabled = True
    MDIForm_Xsol.nVisitas.enabled = True
    MDIForm_Xsol.nconfig.enabled = (Profile = ADMIN_PROFILE)
    MDIForm_Xsol.nconfig.Visible = (Profile = ADMIN_PROFILE)
    MDIForm_Xsol.Arrange vbTileHorizontal
    
    Default_Printer = Printer.DeviceName
    
    usuario.Login = Login
    usuario.Profile = Profile
End Sub

Private Sub Text_Senha_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Cmd_Acessar_Click
    End If
    
End Sub


