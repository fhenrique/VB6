VERSION 5.00
Begin VB.MDIForm FrmPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Gestão do Seminário São José  / Garanhuns-PE"
   ClientHeight    =   7635
   ClientLeft      =   3315
   ClientTop       =   2235
   ClientWidth     =   11490
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   23040
      Left            =   0
      Picture         =   "FrmPrincipal.frx":0000
      ScaleHeight     =   23040
      ScaleWidth      =   11490
      TabIndex        =   0
      Top             =   -15405
      Width           =   11490
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "Cadastros"
      Index           =   0
      Begin VB.Menu mnuCadastroHospedes 
         Caption         =   "Hóspedes"
         Index           =   1
      End
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "Movimentos"
      Index           =   4
      Begin VB.Menu mnuQuartos 
         Caption         =   "Quartos"
         Index           =   5
      End
   End
   Begin VB.Menu mnuOutros 
      Caption         =   "Outros"
      Index           =   9
      Begin VB.Menu mnuConfiguracoes 
         Caption         =   "Configurações"
         Index           =   7
      End
      Begin VB.Menu mnuFecharSistema 
         Caption         =   "Fechar sistema"
         Index           =   3
      End
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ConfigFile As String
Public Path As String
Public FileName As String

Private Sub mnuCadastroHospedes_Click(Index As Integer)
    FrmCadHospedes.Show (1)
End Sub

Private Sub mnuConfiguracoes_Click(Index As Integer)
    frmConfiguracoes.Show (1)
End Sub

Private Sub mnuFecharSistema_Click(Index As Integer)
    If MsgBox("Confirma o fechamento do sistema?", vbYesNo, "Confirmação!") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub mnuQuartos_Click(Index As Integer)
    frmQuartos.Show (1)
End Sub

Private Sub ArquivoDados()
    ConfigFile = App.Path & "\CONFIG.INI"
    
    Path = RTrim(GetINISetting("DATABASE", "PATH", ConfigFile))
    If Right(Path, 1) <> "\" Then
       Path = Path + "\"
    End If
    FileName = GetINISetting("DATABASE", "FILENAME", ConfigFile)
End Sub
