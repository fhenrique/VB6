VERSION 5.00
Begin VB.Form Form_Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XSol Aquecedores"
   ClientHeight    =   3405
   ClientLeft      =   6495
   ClientTop       =   4500
   ClientWidth     =   4995
   Icon            =   "Form_Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
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
      Left            =   255
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
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   4935
      TabIndex        =   4
      Top             =   3030
      Width           =   4995
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   120
      Picture         =   "Form_Login.frx":000C
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   240
      Picture         =   "Form_Login.frx":1CD6
      Top             =   480
      Width           =   510
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

Sub cloud_ConfiguracaoArquivos_Dados()
    
    cloud_ConfigFile = App.Path & "\CONFIG.INI"
    
    cloud_Servidor = RTrim(GetINISetting("DATABASE", "server", cloud_ConfigFile))
    cloud_Database = RTrim(GetINISetting("DATABASE", "database", cloud_ConfigFile))
    cloud_usuario = RTrim(GetINISetting("DATABASE", "user", cloud_ConfigFile))
    cloud_senha = RTrim(GetINISetting("DATABASE", "senha", cloud_ConfigFile))

    
    cloud_FileName = App.Path & "\" & GetINISetting("IMAGE", "FOLDER", cloud_ConfigFile)
    
End Sub

Private Sub Cmd_Acessar_Click()
On Error GoTo fim
    Dim local_usuario As String
    Dim local_senha As String
    
    If (RTrim(LTrim(Text_Usuario.Text)) = "") Or (RTrim(LTrim(Text_Senha.Text)) = "") Then
        MsgBox "Os campos usuário e senha devem estar preenchidos!"
        Exit Sub
    End If
    
    local_usuario = Text_Usuario.Text
    local_senha = Text_Senha.Text
    
    cloud_select = "SELECT perfil FROM Usuarios where usuario = '" & local_usuario & "' and senha = '" & local_senha & "'"

retenta:
    Set cloud_registro = cloud_conexao.Execute(cloud_select)
    
    If cloud_registro.EOF Then
        MsgBox "Usuário ou Senha errado(s)", vbCritical, "XSOL Aquecedores"
        Exit Sub
    Else
        Unload Me
        Exit Sub
    End If
    
fim:
    
    If cloud_conexao.State = 1 Then
        cloud_conexao.Close
        If ConectaMySQL Then
            GoTo retenta
        Else
            MsgBox "Cloud indisponível, tente novamente em alguns minutos!", vbCritical, "XSOL Aquecedores"
            End
        End If
        
    End If
    
End Sub


Private Sub Cmd_Cancelar_Click()

Unload Me

End Sub

Private Sub Form_Activate()
    Text_Usuario.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    Text_Senha.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito

'Obtem os dados do arquivo de configuração
cloud_ConfiguracaoArquivos_Dados

'FAZ A CONEXÃO GLOBAL com o MySQL
If ConectaMySQL Then

Else
    MsgBox "Não foi possível conectar ao banco de dados", vbCritical, "XSOL Aquecedores"
    Unload Me
    Unload MDIForm_Xsol
End If
    
    
    
End Sub

Private Sub Form_Load()

    
'    MDIForm_Xsol.nclientes.enabled = False
'    MDIForm_Xsol.ncadastros.enabled = False
'    MDIForm_Xsol.nuser.enabled = False
'    MDIForm_Xsol.nuser.Visible = False
'      MDIForm_Xsol.nservtype.enabled = False
'    MDIForm_Xsol.ntecnico.enabled = False
'    MDIForm_Xsol.nzone.enabled = False
'    MDIForm_Xsol.nAgenda.enabled = False
'    MDIForm_Xsol.nVisitas.enabled = False
'    MDIForm_Xsol.nconfig.enabled = False
'    MDIForm_Xsol.nconfig.Visible = False
    

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


