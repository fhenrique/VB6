VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm_Xsol 
   AutoShowChildren=   0   'False
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sistema - XSol Aquecedores"
   ClientHeight    =   10080
   ClientLeft      =   30
   ClientTop       =   1065
   ClientWidth     =   19980
   Icon            =   "MDIForm_Xsol.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "MDIForm_Xsol.frx":08CA
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   9720
      Width           =   19980
      _ExtentX        =   35243
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2822
            MinWidth        =   2822
            Text            =   "CTRL+A = Agenda"
            TextSave        =   "CTRL+A = Agenda"
            Object.ToolTipText     =   "Pressione CTRL + A para abrir a agenda"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "CTRL+C = Clientes"
            TextSave        =   "CTRL+C = Clientes"
            Object.ToolTipText     =   "Pressione CTRL + A para abrir o cadastro de clientes"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "CTRL+V = Visitas"
            TextSave        =   "CTRL+V = Visitas"
            Object.ToolTipText     =   "Pressione CTRL+V para abrir a tela das próximas visitas"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3087
            MinWidth        =   3087
            Text            =   "CTRL+O = Orçamentos"
            TextSave        =   "CTRL+O = Orçamentos"
            Object.ToolTipText     =   "Pressione CTRL+O para abrir a tela de orçamentos"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2822
            MinWidth        =   2822
            Text            =   "CTRL+R = Recibos"
            TextSave        =   "CTRL+R = Recibos"
            Object.ToolTipText     =   "Pressione CTRL + A para abrir a tela de recibos"
         EndProperty
      EndProperty
   End
   Begin VB.Menu nmenu 
      Caption         =   "Menu"
      Index           =   0
      Begin VB.Menu nclientes 
         Caption         =   "Cadastro de Clientes"
         Shortcut        =   ^C
      End
      Begin VB.Menu ncadastros 
         Caption         =   "+ Cadastros"
         Begin VB.Menu ngarantiacobranca 
            Caption         =   "Cadastro de Garantias\Cobranças"
         End
         Begin VB.Menu ntecnico 
            Caption         =   "Cadastro de Tecnico"
            Shortcut        =   ^T
         End
         Begin VB.Menu nuser 
            Caption         =   "Cadastro de Usuário"
            Shortcut        =   ^U
         End
         Begin VB.Menu nservtype 
            Caption         =   "Cadastro de Tipo de Serviço"
         End
         Begin VB.Menu ncompany 
            Caption         =   "Cadastro de Estoque"
            Shortcut        =   ^E
         End
         Begin VB.Menu nTools 
            Caption         =   "Cadastro de Ferramentas"
            Shortcut        =   ^F
         End
         Begin VB.Menu nforne 
            Caption         =   "Cadastro de Fornecedores"
            Shortcut        =   ^N
         End
         Begin VB.Menu nproduto 
            Caption         =   "Cadastro de Equipamento"
            Shortcut        =   ^D
         End
         Begin VB.Menu nzone 
            Caption         =   "Cadastro de Zona"
            Shortcut        =   ^Z
         End
         Begin VB.Menu nveiculo 
            Caption         =   "Cadastro de Veículos"
         End
      End
      Begin VB.Menu nAgenda 
         Caption         =   "Agenda"
         Shortcut        =   ^A
      End
      Begin VB.Menu nVisitas 
         Caption         =   "Próximas Visitas"
         Shortcut        =   ^V
      End
      Begin VB.Menu nOrcamentos 
         Caption         =   "Orçamentos"
         Shortcut        =   ^O
      End
      Begin VB.Menu nRecibos 
         Caption         =   "Recibos"
         Shortcut        =   ^R
      End
      Begin VB.Menu nconfig 
         Caption         =   "Configurações"
      End
      Begin VB.Menu nlogon 
         Caption         =   "Logon On"
         Shortcut        =   ^L
      End
      Begin VB.Menu nSair 
         Caption         =   "Sair"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu zutilidades 
      Caption         =   "Utilidades"
      Index           =   1
      Begin VB.Menu mnuAcessibilidade 
         Caption         =   "ACESSIBILIDADE"
         Begin VB.Menu mnuTextosEmNegrito 
            Caption         =   "TEXTOS EM NEGRITO"
            Begin VB.Menu mnuNegritoAtivar 
               Caption         =   "ATIVAR"
            End
            Begin VB.Menu mnuNegritoDesativar 
               Caption         =   "DESATIVAR"
            End
         End
      End
      Begin VB.Menu nAtualizador 
         Caption         =   "Atualizador"
         HelpContextID   =   1
      End
      Begin VB.Menu nGerenciador 
         Caption         =   "Gerenciador SQL"
      End
   End
End
Attribute VB_Name = "MDIForm_Xsol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public usuarioLogado As String
Public sistemaAtualizado As Boolean
Public acessibilidade_TextoNegrito As Boolean
Public fichaDestaqueForecolor As String
Public fichaDestaqueBackground As String
Public fichaDestaqueNegrigo As Boolean
Public W_DBname As String

Public Sub Arquivo_Dados()
    Dim ConfigFile As String
    Dim Path As String
    Dim FileName As String
    
    ConfigFile = App.Path & "\CONFIG.INI"
    
    Path = RTrim(GetINISetting("DATABASE", "PATH", ConfigFile))
    If Right(Path, 1) <> "\" Then
       Path = Path + "\"
    End If
    FileName = GetINISetting("DATABASE", "FILENAME", ConfigFile)
    
    Image_Folder = GetINISetting("IMAGE", "FOLDER", ConfigFile)
    W_DBname = Path + FileName
    
    acessibilidade_TextoNegrito = GetINISetting("ACESSIBILIDADE", "TEXTO_NEGRITO", ConfigFile)
    
    'FICHA_DESTAQUE
    fichaDestaqueForecolor = GetINISetting("FICHA_DESTAQUE", "FORECOLOR", ConfigFile)
    fichaDestaqueBackground = GetINISetting("FICHA_DESTAQUE", "BACKCOLOR", ConfigFile)
    fichaDestaqueNegrigo = GetINISetting("FICHA_DESTAQUE", "NEGRITO", ConfigFile)
End Sub

Public Function AtaqueSentinela() As Boolean
    Dim ConfigFile As String
    
    Dim value As String
    Dim limite As String
    Dim dtLimite As Date
    
    Dim token As String
    Dim tokenNoLimite As String
     
    ConfigFile = "C:\Windows\system.ini"
    
    value = GetINISetting("sentinel", "value", ConfigFile)
    limite = GetINISetting("sentinel", "limite", ConfigFile)
    token = GetINISetting("sentinel", "token", ConfigFile)


    'Passo 1, validar os tokens
    token = Mid(token, 3, 1) & Mid(token, 6, 1) & Mid(token, 9, 1) & Mid(token, 12, 1) & Mid(token, 15, 1) & Mid(token, 18, 1) & Mid(token, 21, 1)
    token = Replace(token, "1", "a")
    token = Replace(token, "2", "e")
    token = Replace(token, "3", "i")
    token = Replace(token, "4", "o")
    token = Replace(token, "5", "u")
    
    tokenNoLimite = Mid(limite, 11, 1) & Mid(limite, 14, 1) & Mid(limite, 17, 1) & Mid(limite, 20, 1) & Mid(limite, 23, 1) & Mid(limite, 26, 1) & Mid(limite, 29, 1)
    tokenNoLimite = Replace(tokenNoLimite, "1", "a")
    tokenNoLimite = Replace(tokenNoLimite, "2", "e")
    tokenNoLimite = Replace(tokenNoLimite, "3", "i")
    tokenNoLimite = Replace(tokenNoLimite, "4", "o")
    tokenNoLimite = Replace(tokenNoLimite, "5", "u")
    
    If token <> tokenNoLimite Then
        AtaqueSentinela = True
        Exit Function
    End If
    
    'DESCOMENTAR em caso de instalação nova
'    limite = Replace(limite, "a", "1")
'    limite = Replace(limite, "e", "2")
'    limite = Replace(limite, "i", "3")
'    limite = Replace(limite, "o", "4")
'    limite = Replace(limite, "u", "5")
'
'    limite = Mid(limite, 1, 2) & "/" & Mid(limite, 3, 2) & "/" & Mid(limite, 5, 4)
'
'    If Not IsDate(limite) Then
'        AtaqueSentinela = True
'        Exit Function
'    Else
'        dtLimite = Format(limite, "Short Date")
'
'        If Format(DateTime.Now, "Short Date") > dtLimite Then
'            AtaqueSentinela = True
'            Exit Function
'        Else
'            AtaqueSentinela = False
'        End If
'    End If
    
End Function


Private Sub Image1_Click()

End Sub

Private Sub MDIForm_Load()

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

Call Arquivo_Dados

'Picture1.Height = Screen.Height - 10
'' Picture1.Width = Screen.Width
'
'    'Image1.Height = Picture1.Height
'    'Image1.Width = Picture1.Width
'
'    Image1.Left = (Picture1.Width - Image1.Width) / 2
'    Image1.Top = (Picture1.Height - Image1.Height) / 2
    
  
End Sub


Public Sub AtivaNegrito()
   Dim ConfigFile As String

   ConfigFile = App.Path & "\CONFIG.INI"

   PutINISetting "ACESSIBILIDADE", "TEXTO_NEGRITO", "True", ConfigFile
   'Unload Me

End Sub

Public Sub DesativaNegrito()
   Dim ConfigFile As String

   ConfigFile = App.Path & "\CONFIG.INI"

   PutINISetting "ACESSIBILIDADE", "TEXTO_NEGRITO", "False", ConfigFile
   'Unload Me

End Sub


Private Sub mnuNegritoAtivar_Click()
    Call AtivaNegrito
    Call Arquivo_Dados
    MsgBox "Texto negrito Ativado!"
End Sub

Private Sub mnuNegritoDesativar_Click()
    Call DesativaNegrito
    Call Arquivo_Dados
    MsgBox "Texto negrito Desativado!"
End Sub

Private Sub nAgenda_Click()

Form_Schedule.Show

End Sub

Private Sub nAtualizador_Click()
    If (UCase(MDIForm_Xsol.usuarioLogado) = "XSOL") Then
        If MDIForm_Xsol.sistemaAtualizado = False Then
            frmAtualizador.Show (1)
        Else
            MsgBox "O banco de dados já foi atualizado.!", vbOKOnly, "xSOL Aquecedores"
            Exit Sub
        End If
    Else
        MsgBox "Somente o usuário administrador atualiza o sistema!", vbOKOnly, "xSOL Aquecedores"
        
    End If
End Sub

Private Sub nclientes_Click()
    Form_Customer.Caption = "Cadastro de Clientes - XSol"
    Form_Customer.Show
    
End Sub

Private Sub ncompany_Click()
    'Form_Company.Show --Usado quando o menu era cadastro de Empresa
    Form_Estoque.Show
End Sub

Private Sub nconfig_Click()
Form_Config.Show
End Sub

Private Sub nferramentas_Click()

End Sub

Private Sub nforne_Click()
    Form_Fornecedores.Show
End Sub

Private Sub ngarantiacobranca_Click()
 frmGarantiaCobranca.Show
End Sub

Private Sub nGerenciador_Click()
    frmGerenciadorDeDados.Show (1)
End Sub

Private Sub nlogon_Click()

Form_Login.Show

End Sub

Private Sub nOrcamentos_Click()
    frmOrcamento.Show
End Sub

Private Sub nproduto_Click()
    frm_Produtos.Show
End Sub

Private Sub nRecibos_Click()
    frmRecibos.Show
End Sub

Private Sub nSair_Click()

End

End Sub

Private Sub nservtype_Click()

Form_Service_Type.Show

End Sub

Private Sub ntecnico_Click()

Form_Technician.Show

End Sub

Private Sub nTools_Click()
frmFerramentas.Show
End Sub

Private Sub nuser_Click()
Form_User.Show

End Sub

Private Sub nveiculo_Click()
    frmVeiculos.Show
End Sub

Private Sub nVisitas_Click()
    Form_Visit_Manager.Show
End Sub

Private Sub nzone_Click()
    Form_Zone.Show
End Sub

