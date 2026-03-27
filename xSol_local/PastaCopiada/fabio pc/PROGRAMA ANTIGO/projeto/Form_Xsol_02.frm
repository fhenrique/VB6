VERSION 5.00
Begin VB.Form Form_Xsol_02 
   BorderStyle     =   0  'None
   Caption         =   "Cadastro de Clientes - XSol"
   ClientHeight    =   7575
   ClientLeft      =   1095
   ClientTop       =   0
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAtende 
      Caption         =   "A&tende"
      Height          =   495
      Left            =   7800
      TabIndex        =   47
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Fechar 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   7800
      TabIndex        =   41
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Edita 
      Caption         =   "&Editar"
      Height          =   495
      Left            =   4200
      TabIndex        =   40
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Exclui 
      Caption         =   "E&xclui"
      Height          =   495
      Left            =   6240
      TabIndex        =   39
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Salva 
      Caption         =   "&Salvar"
      Height          =   495
      Left            =   2160
      TabIndex        =   38
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Inclui 
      Caption         =   "&Incluir"
      Height          =   495
      Left            =   240
      TabIndex        =   37
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Ultimo 
      Caption         =   "&Último  >>"
      Height          =   495
      Left            =   6240
      TabIndex        =   36
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Proximo 
      Caption         =   "Pr&oximo  >"
      Height          =   495
      Left            =   4200
      TabIndex        =   35
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Anterior 
      Caption         =   "<  &Anterior"
      Height          =   495
      Left            =   2160
      TabIndex        =   34
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Primeiro 
      Caption         =   "<<   &Primeiro"
      Height          =   495
      Left            =   240
      TabIndex        =   33
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Frame Frame_Cliente 
      Caption         =   "Cadastro de Clientes"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin VB.TextBox Text_Dt_Contrato 
         Height          =   285
         Left            =   5400
         TabIndex        =   54
         Top             =   5760
         Width           =   1095
      End
      Begin VB.TextBox Text_Contrato 
         Height          =   285
         Left            =   2880
         TabIndex        =   52
         Top             =   5760
         Width           =   1215
      End
      Begin VB.ComboBox Combo_Solicitante 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   5400
         Width           =   3975
      End
      Begin VB.TextBox Text_Solicitante 
         Height          =   285
         Left            =   2880
         TabIndex        =   48
         Top             =   5400
         Width           =   3735
      End
      Begin VB.Frame FrameSelecao 
         Caption         =   "Selecione"
         Height          =   585
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   1665
         Begin VB.OptionButton Option_Cnpj 
            Caption         =   "CNPJ"
            Height          =   195
            Left            =   840
            TabIndex        =   46
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option_Cpf 
            Caption         =   "CPF"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.ComboBox Combo_Zona 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   3240
         Width           =   2775
      End
      Begin VB.ComboBox Combo_Uf 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text_NomeCli 
         Height          =   285
         Left            =   2880
         TabIndex        =   16
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox Text_End 
         Height          =   285
         Left            =   2880
         TabIndex        =   15
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox Text_Dt_Emis 
         Height          =   285
         Left            =   2880
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text_Cpf_Cnpj 
         Height          =   285
         Left            =   2880
         TabIndex        =   13
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text_Bairro 
         Height          =   285
         Left            =   2880
         TabIndex        =   12
         Top             =   1800
         Width           =   4575
      End
      Begin VB.TextBox Text_Cidade 
         Height          =   285
         Left            =   2880
         TabIndex        =   11
         Top             =   2160
         Width           =   4575
      End
      Begin VB.TextBox Text_Uf 
         Height          =   285
         Left            =   2880
         TabIndex        =   10
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox Text_Cep 
         Height          =   285
         Left            =   4080
         TabIndex        =   9
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox Text_Ref 
         Height          =   285
         Left            =   2880
         TabIndex        =   8
         Top             =   2880
         Width           =   4455
      End
      Begin VB.TextBox Text_Zona 
         Height          =   285
         Left            =   2880
         TabIndex        =   7
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox Text_Fone_Resid 
         Height          =   285
         Left            =   2880
         TabIndex        =   6
         Top             =   3600
         Width           =   2535
      End
      Begin VB.TextBox Text_Fone_Com 
         Height          =   285
         Left            =   2880
         TabIndex        =   5
         Top             =   3960
         Width           =   2535
      End
      Begin VB.TextBox Text_Celular 
         Height          =   285
         Left            =   2880
         TabIndex        =   4
         Top             =   4320
         Width           =   2535
      End
      Begin VB.TextBox Text_E_Mail 
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Top             =   4680
         Width           =   2535
      End
      Begin VB.TextBox Text_Tipo_Equip 
         Height          =   285
         Left            =   2880
         TabIndex        =   2
         Top             =   5040
         Width           =   2775
      End
      Begin VB.TextBox Text_Atendente 
         Height          =   285
         Left            =   6240
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Lbl_Dt_Contrato 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt. Contrato:"
         Height          =   255
         Left            =   4320
         TabIndex        =   53
         Top             =   5760
         Width           =   975
      End
      Begin VB.Label Lbl_Contrato 
         Alignment       =   1  'Right Justify
         Caption         =   "Contrato:"
         Height          =   255
         Left            =   2040
         TabIndex        =   51
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label Llb_Solitante 
         Alignment       =   1  'Right Justify
         Caption         =   "Empresa solicitante:"
         Height          =   255
         Left            =   1200
         TabIndex        =   49
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Label Lbl_Nome 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome do Cliente / Razão Social:"
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Lbl_End 
         Alignment       =   1  'Right Justify
         Caption         =   "Endereço:"
         Height          =   255
         Left            =   2040
         TabIndex        =   31
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Lbl_Dt_Emis 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Emissão:"
         Height          =   255
         Left            =   1680
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Lbl_Cpf_Cnpj 
         Alignment       =   1  'Right Justify
         Caption         =   "Cpf / Cnpj:"
         Height          =   255
         Left            =   1920
         TabIndex        =   29
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Lbl_Bairro 
         Alignment       =   1  'Right Justify
         Caption         =   "Bairro:"
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Lbl_Cidade 
         Alignment       =   1  'Right Justify
         Caption         =   "Cidade:"
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Lbl_Uf 
         Alignment       =   1  'Right Justify
         Caption         =   "Uf:"
         Height          =   255
         Left            =   2520
         TabIndex        =   26
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Lbl_Cep 
         Alignment       =   1  'Right Justify
         Caption         =   "Cep:"
         Height          =   255
         Left            =   3600
         TabIndex        =   25
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Lbl_Ref 
         Alignment       =   1  'Right Justify
         Caption         =   "Referência:"
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Lbl_Zona 
         Alignment       =   1  'Right Justify
         Caption         =   "Zona:"
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Lbl_Fone_Resid 
         Alignment       =   1  'Right Justify
         Caption         =   "Fone Resid:"
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Lbl_Fone_Com 
         Alignment       =   1  'Right Justify
         Caption         =   "Fone Com.:"
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Lbl_Celular 
         Alignment       =   1  'Right Justify
         Caption         =   "Celular:"
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label Lbl_E_Mail 
         Alignment       =   1  'Right Justify
         Caption         =   "E-Mail:"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   4680
         Width           =   495
      End
      Begin VB.Label Lbl_Tipo_Eq 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo de Equip / Modelo:"
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label Lbl_Atendente 
         Alignment       =   1  'Right Justify
         Caption         =   "Atendente:"
         Height          =   255
         Left            =   5280
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form_Xsol_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public W_DB As Connection
Public Tabela_Clientes As Recordset
Public Tabela_Empresas As Recordset
Public W_DBname As String

Private Sub Cmd_Anterior_Click()

If Tabela_Clientes.BOF = False Then
   Tabela_Clientes.MovePrevious
   Call Mostra_Registro
   If Tabela_Clientes.BOF = True Then
      Tabela_Clientes.MoveFirst
      Call Mostra_Registro
   End If
   Else
   If Tabela_Clientes.EOF Then
      MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
   Else
      Tabela_Clientes.MoveFirst
      Call Mostra_Registro
   End If
End If

End Sub

Private Sub Cmd_Edita_Click()

    If Cmd_Edita.Caption = "&Editar" Then
            
        Cmd_Edita.Caption = "&Cancela"
    
        Cmd_Salva.Enabled = True
        
        Call DesabilitaNavegacao
        Cmd_Inclui.Enabled = False
        Cmd_Exclui.Enabled = False
        
        Call Destrava_Registro
        Call Cad_Ed
               
        Text_Cpf_Cnpj.BackColor = vbYellow
        Text_NomeCli.SetFocus
        
        
    ElseIf Cmd_Edita.Caption = "&Cancela" Then
        Cmd_Edita.Caption = "&Editar"
        
        Cmd_Salva.Enabled = False
        
        Call HabilitaNavegacao
        
        Cmd_Inclui.Enabled = True
        Cmd_Exclui.Enabled = True
        
        Call Cad_Ed_Fim
        Call Trava_Registro
        
      
        
    End If

End Sub

Private Sub Cmd_Exclui_Click()

If MsgBox("Confirma Exclusão deste Registro ? ", vbYesNo, "Excluir Registros") = vbYes Then
        
   On Error Resume Next
        
   Tabela_Clientes.Delete
        
   Tabela_Clientes.MoveNext
   If Tabela_Clientes.EOF = True Then
      Tabela_Clientes.MoveLast
      If Tabela_Clientes.BOF = True Then
         Call LimpaControles
              MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
      End If
   End If
   ElseIf Tabela_Clientes.EOF = True And Tabela_Clientes.BOF = True Then
                MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
End If
Call Mostra_Registro

End Sub

Private Sub Cmd_Fechar_Click()

Unload Me

End Sub

Private Sub Cmd_Inclui_Click()

Call Cad_Ed

If Cmd_Inclui.Caption = "&Incluir" Then
   
   Cmd_Inclui.Caption = "&Cancela"
   Cmd_Salva.Enabled = True
   Cmd_Edita.Enabled = False
   Cmd_Exclui.Enabled = False
   
   Call DesabilitaNavegacao
   Call LimpaControles
   Call Destrava_Registro
   
   Text_NomeCli.SetFocus
   
ElseIf Cmd_Inclui.Caption = "&Cancela" Then
   
   Cmd_Inclui.Caption = "&Incluir"
   Cmd_Salva.Enabled = False
   Cmd_Edita.Enabled = True
   Cmd_Exclui.Enabled = True
   
   Call Trava_Registro
   Call HabilitaNavegacao
   Call Mostra_Registro
   
   Call Cad_Ed_Fim

End If

End Sub

Private Sub Cmd_Primeiro_Click()

If Tabela_Clientes.BOF = False Then
   Tabela_Clientes.MoveFirst
   Call Mostra_Registro
ElseIf Tabela_Clientes.BOF = True _
   And Tabela_Clientes.EOF = True Then
   MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
End If

End Sub

Private Sub Cmd_Proximo_Click()

If Tabela_Clientes.EOF = False Then
   Tabela_Clientes.MoveNext
   Call Mostra_Registro
   If Tabela_Clientes.EOF Then
      Tabela_Clientes.MoveLast
      Call Mostra_Registro
   End If
Else
   If Tabela_Clientes.BOF Then
      MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
   Else
      Tabela_Clientes.MoveLast
      Call Mostra_Registro
   End If
End If

End Sub

Private Sub Cmd_Salva_Click()

If Trim(Text_NomeCli.Text) = "" Then
   MsgBox "É necessário digitar o Nome do Cliente / Razão Social! ", vbOKOnly, "X.Sol Arte em Aquecedores"
   Text_NomeCli.SetFocus
   Exit Sub
End If
    
If Trim(Text_End.Text) = "" Then
   MsgBox "É necessário digitar o Endereço! ", vbOKOnly, "X.Sol Arte em Aquecedores"
   Text_End.SetFocus
   Exit Sub
End If

If Trim(Text_Bairro.Text) = "" Then
   MsgBox "É necessário digitar o Bairro! ", vbOKOnly, "X.Sol Arte em Aquecedores"
   Text_Bairro.SetFocus
   Exit Sub
End If

If Trim(Text_Cidade.Text) = "" Then
   MsgBox "É necessário digitar Cidade! ", vbOKOnly, "X.Sol Arte em Aquecedores"
   Text_Cidade.SetFocus
   Exit Sub
End If

If Trim(Text_Cep.Text) = "" Then
   MsgBox "É necessário digitar CEP! ", vbOKOnly, "X.Sol Arte em Aquecedores"
   Text_Cep.SetFocus
   Exit Sub
End If

If Combo_Uf.Text = "" And Text_Zona.Text = "" Then
   MsgBox "É necessário selecionar a Unidade Federativa! ", vbOKOnly, "X.Sol Arte em Aquecedores"
   Combo_Uf.SetFocus
   Exit Sub
End If

If Combo_Zona.Text = "" And Text_Zona.Text = "" Then
   MsgBox "É necessário selecionar a ZONA! ", vbOKOnly, "X.Sol Arte em Aquecedores"
   Combo_Zona.SetFocus
   Exit Sub
End If

If Trim(Text_E_Mail.Text) = "" Then
   MsgBox "É necessário digitar E-mail! ", vbOKOnly, "X.Sol Arte em Aquecedores"
   Text_E_Mail.SetFocus
   Exit Sub
End If

If Combo_Solicitante.Text = "" And Text_Solicitante.Text = "" Then
   MsgBox "É necessário selecionar o Solicitante! ", vbOKOnly, "X.Sol Arte em Aquecedores"
   Combo_Zona.SetFocus
   Exit Sub
End If


If Cmd_Inclui.Caption = "&Cancela" Then
   Tabela_Clientes.AddNew
End If

Call GravaRegistros

Tabela_Clientes.Update

Call Trava_Registro
Call HabilitaNavegacao
Call Mostra_Registro


If Cmd_Inclui.Caption = "&Cancela" Then
   Cmd_Inclui.Caption = "&Incluir"
End If
    
If Cmd_Edita.Caption = "&Cancela" Then
   Cmd_Edita.Caption = "&Editar"
End If
     
Call Cad_Ed_Fim

Cmd_Edita.Enabled = True
Cmd_Inclui.Enabled = True
Cmd_Exclui.Enabled = True
    
End Sub

Private Sub Cmd_Ultimo_Click()

If Tabela_Clientes.EOF = False Then
   Tabela_Clientes.MoveLast
   Call Mostra_Registro
ElseIf Tabela_Clientes.BOF = True _
  And Tabela_Clientes.EOF = True Then
  MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
End If

End Sub

Private Sub CmdAtende_Click()

Cod_Cli = 0
Cod_Cli = Tabela_Clientes.Fields("Codigo")

Nome_Cli_Atend = ""
Nome_Cli_Atend = Tabela_Clientes.Fields("Nome_Cli")

Form_Xsol_05.Show

End Sub




Private Sub Combo_Solicitante_Click()

Text_Solicitante.Text = Combo_Solicitante.Text

End Sub
Private Sub Combo_Solicitante_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{TAB}"
   Text_Solicitante.Text = Combo_Solicitante.Text
End If

End Sub

Private Sub Combo_Uf_Click()

Text_Uf.Text = Combo_Uf.Text

End Sub

Private Sub Combo_Uf_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{TAB}"
   Text_Uf.Text = Combo_Uf.Text
End If

End Sub
Private Sub Combo_Zona_Click()

Text_Zona.Text = Combo_Zona.Text

End Sub

Private Sub Combo_Zona_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{TAB}"
   Text_Zona.Text = Combo_Zona.Text
End If

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

Call Ajusta_Form(Me)

Text_NomeCli.MaxLength = 80
Text_Cpf_Cnpj.MaxLength = 15
Text_Atendente.MaxLength = 20
Text_End.MaxLength = 70
Text_Bairro.MaxLength = 50
Text_Cidade.MaxLength = 50
Text_Uf.MaxLength = 2
Text_Cep.MaxLength = 11
Text_Ref.MaxLength = 70
Text_Zona.MaxLength = 20
Text_Fone_Resid.MaxLength = 25
Text_Fone_Com.MaxLength = 25
Text_Celular.MaxLength = 25
Text_E_Mail.MaxLength = 50
Text_Solicitante.MaxLength = 50
Text_Contrato.MaxLength = 25

If Right(App.Path, 1) = "\" Then
  W_DBname = App.Path & "dados.mdb"
 Else
  W_DBname = App.Path & "\" & "dados.mdb"
End If

Set W_DB = New Connection

Set Tabela_Clientes = New Recordset

W_DB.Mode = adModeReadWrite

W_DB.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"

Tabela_Clientes.Open "select * from Clientes Order By Codigo ", W_DB, adOpenStatic, adLockOptimistic

Set Tabela_Empresas = New Recordset
Tabela_Empresas.Open "select Tipo_Cliente from Empresa Order By Tipo_Cliente ", W_DB, adOpenStatic, adLockOptimistic

Form_Xsol_02.Caption = "Cadastro de Clientes Xsol"
Text_Dt_Emis.Enabled = False

Call Trava_Registro
Call Mostra_Registro

End Sub

Sub Trava_Registro()


Combo_Zona.Visible = False
Combo_Uf.Visible = False
FrameSelecao.Visible = False
Combo_Solicitante.Visible = False

Text_NomeCli.Locked = True
Text_Cpf_Cnpj.Locked = True
Text_Atendente.Locked = True
Text_End.Locked = True
Text_Bairro.Locked = True
Text_Cidade.Locked = True
Text_Uf.Locked = True
Text_Cep.Locked = True
Text_Ref.Locked = True
Text_Zona.Locked = True
Text_Fone_Resid.Locked = True
Text_Fone_Com.Locked = True
Text_Celular.Locked = True
Text_E_Mail.Locked = True
Text_Tipo_Equip.Locked = True
Text_Solicitante.Locked = True
Text_Contrato.Locked = True
Text_Dt_Contrato.Locked = True

End Sub

Sub Destrava_Registro()

Combo_Zona.Visible = True
Combo_Uf.Visible = True
FrameSelecao.Visible = True
Combo_Solicitante.Visible = True

Text_NomeCli.Locked = False
Text_Cpf_Cnpj.Locked = False
Text_Atendente.Locked = False
Text_End.Locked = False
Text_Bairro.Locked = False
Text_Cidade.Locked = False
Text_Uf.Locked = False
Text_Cep.Locked = False
Text_Ref.Locked = False
Text_Zona.Locked = False
Text_Fone_Resid.Locked = False
Text_Fone_Com.Locked = False
Text_Celular.Locked = False
Text_E_Mail.Locked = False
Text_Tipo_Equip.Locked = False
Text_Solicitante.Locked = False
Text_Contrato.Locked = False
Text_Dt_Contrato.Locked = False

End Sub

Sub LimpaControles()

Text_NomeCli.Text = ""
Text_Cpf_Cnpj.Text = ""
Text_Atendente.Text = ""
Text_End.Text = ""
Text_Bairro.Text = ""
Text_Cidade.Text = ""
Text_Uf.Text = ""
Text_Cep.Text = ""
Text_Ref.Text = ""
Text_Zona.Text = ""
Text_Fone_Resid.Text = ""
Text_Fone_Com.Text = ""
Text_Celular.Text = ""
Text_E_Mail.Text = ""
Text_Tipo_Equip.Text = ""
Text_Contrato.Text = ""
Text_Dt_Contrato.Text = ""

End Sub

Private Sub GravaRegistros()
            
   
   If Trim(Text_NomeCli.Text) = "" Then
      Tabela_Clientes.Fields("Nome_Cli").Value = Null
   Else
      Tabela_Clientes.Fields("Nome_Cli").Value = Text_NomeCli.Text
   End If
   
   If Trim(Text_Cpf_Cnpj.Text) = "" Then
      Tabela_Clientes.Fields("CPF_CNJP").Value = Null
   Else
      Tabela_Clientes.Fields("CPF_CNJP").Value = Text_Cpf_Cnpj.Text
   End If
   
   If Trim(Text_Atendente.Text) = "" Then
      Tabela_Clientes.Fields("Atendente").Value = Null
   Else
      Tabela_Clientes.Fields("Atendente").Value = Text_Atendente.Text
   End If
      
   If Trim(Text_End.Text) = "" Then
      Tabela_Clientes.Fields("End").Value = Null
   Else
      Tabela_Clientes.Fields("End").Value = Text_End.Text
   End If
   
   If Trim(Text_Bairro.Text) = "" Then
      Tabela_Clientes.Fields("Bairro").Value = Null
   Else
      Tabela_Clientes.Fields("Bairro").Value = Text_Bairro.Text
   End If
   
   If Trim(Text_Cidade.Text) = "" Then
      Tabela_Clientes.Fields("Cidade").Value = Null
   Else
      Tabela_Clientes.Fields("Cidade").Value = Text_Cidade.Text
   End If
   
   If Trim(Text_Uf.Text) = "" Then
      Tabela_Clientes.Fields("Uf").Value = Null
   Else
      Tabela_Clientes.Fields("Uf").Value = Text_Uf.Text
   End If
   
   If Trim(Text_Cep.Text) = "" Then
      Tabela_Clientes.Fields("Cep").Value = Null
   Else
      Tabela_Clientes.Fields("Cep").Value = Text_Cep.Text
   End If
   
   If Trim(Text_Ref.Text) = "" Then
      Tabela_Clientes.Fields("Ref").Value = Null
   Else
      Tabela_Clientes.Fields("Ref").Value = Text_Ref.Text
   End If
   
   If Trim(Text_Zona.Text) = "" Then
      Tabela_Clientes.Fields("Zona").Value = Null
   Else
      Tabela_Clientes.Fields("Zona").Value = Text_Zona.Text
   End If
   
   If Trim(Text_Fone_Resid.Text) = "" Then
      Tabela_Clientes.Fields("Fone_Res").Value = Null
   Else
      Tabela_Clientes.Fields("Fone_Res").Value = Text_Fone_Resid.Text
   End If
   
   If Trim(Text_Fone_Com.Text) = "" Then
      Tabela_Clientes.Fields("Fone_Com").Value = Null
   Else
      Tabela_Clientes.Fields("Fone_Com").Value = Text_Fone_Com.Text
   End If
   
   If Trim(Text_Celular.Text) = "" Then
      Tabela_Clientes.Fields("Fone_Celular").Value = Null
   Else
      Tabela_Clientes.Fields("Fone_Celular").Value = Text_Celular.Text
   End If
   
   If Trim(Text_E_Mail.Text) = "" Then
      Tabela_Clientes.Fields("E_Mail").Value = Null
   Else
      Tabela_Clientes.Fields("E_Mail").Value = Text_E_Mail.Text
   End If
   
   If Trim(Text_Tipo_Equip.Text) = "" Then
      Tabela_Clientes.Fields("Tipo_Eq_Mod").Value = Null
   Else
      Tabela_Clientes.Fields("Tipo_Eq_Mod").Value = Text_Tipo_Equip.Text
   End If
   
   If Trim(Text_Tipo_Equip.Text) = "" Then
      Tabela_Clientes.Fields("Tipo_Eq_Mod").Value = Null
   Else
      Tabela_Clientes.Fields("Tipo_Eq_Mod").Value = Text_Tipo_Equip.Text
   End If
   
   If Trim(Text_Solicitante.Text) = "" Then
      Tabela_Clientes.Fields("Solicitante").Value = Null
   Else
      Tabela_Clientes.Fields("Solicitante").Value = Text_Solicitante.Text
   End If
   
  If Trim(Text_Contrato.Text) = "" Then
      Tabela_Clientes.Fields("Contrato").Value = Null
   Else
      Tabela_Clientes.Fields("Contrato").Value = Text_Contrato.Text
   End If
   
   If Trim(Text_Dt_Contrato.Text) = "" Then
      Tabela_Clientes.Fields("Dt_Contrato").Value = Null
   Else
      Tabela_Clientes.Fields("Dt_Contrato").Value = Text_Dt_Contrato.Text
   End If
   
    
    
End Sub



Sub Mostra_Registro()

If (Tabela_Clientes.EOF = True) Or (Tabela_Clientes.BOF = True) Then
   Exit Sub
End If

Text_NomeCli.Text = IIf(IsNull(Tabela_Clientes.Fields("Nome_Cli")), "", Tabela_Clientes.Fields("Nome_Cli"))

Text_Cpf_Cnpj.Text = IIf(IsNull(Tabela_Clientes.Fields("CPF_CNJP")), "", Tabela_Clientes.Fields("CPF_CNJP"))

Text_Atendente.Text = IIf(IsNull(Tabela_Clientes.Fields("Atendente")), "", Tabela_Clientes.Fields("Atendente"))

Text_End.Text = IIf(IsNull(Tabela_Clientes.Fields("End")), "", Tabela_Clientes.Fields("End"))

Text_Bairro.Text = IIf(IsNull(Tabela_Clientes.Fields("Bairro")), "", Tabela_Clientes.Fields("Bairro"))

Text_Cidade.Text = IIf(IsNull(Tabela_Clientes.Fields("Cidade")), "", Tabela_Clientes.Fields("Cidade"))

Text_Uf.Text = IIf(IsNull(Tabela_Clientes.Fields("Uf")), "", Tabela_Clientes.Fields("Uf"))

Text_Cep.Text = IIf(IsNull(Tabela_Clientes.Fields("Cep")), "", Tabela_Clientes.Fields("Cep"))

Text_Ref.Text = IIf(IsNull(Tabela_Clientes.Fields("Ref")), "", Tabela_Clientes.Fields("Ref"))

Text_Zona.Text = IIf(IsNull(Tabela_Clientes.Fields("Zona")), "", Tabela_Clientes.Fields("Zona"))

Text_Fone_Resid.Text = IIf(IsNull(Tabela_Clientes.Fields("Fone_Res")), "", Tabela_Clientes.Fields("Fone_Res"))

Text_Fone_Com.Text = IIf(IsNull(Tabela_Clientes.Fields("Fone_Com")), "", Tabela_Clientes.Fields("Fone_Com"))

Text_Celular.Text = IIf(IsNull(Tabela_Clientes.Fields("Fone_Celular")), "", Tabela_Clientes.Fields("Fone_Celular"))

Text_E_Mail.Text = IIf(IsNull(Tabela_Clientes.Fields("E_Mail")), "", Tabela_Clientes.Fields("E_Mail"))

Text_Tipo_Equip.Text = IIf(IsNull(Tabela_Clientes.Fields("Tipo_Eq_Mod")), "", Tabela_Clientes.Fields("Tipo_Eq_Mod"))

Text_Solicitante.Text = IIf(IsNull(Tabela_Clientes.Fields("Solicitante")), "", Tabela_Clientes.Fields("Solicitante"))

Text_Contrato.Text = IIf(IsNull(Tabela_Clientes.Fields("Contrato")), "", Tabela_Clientes.Fields("Contrato"))

Text_Dt_Contrato.Text = IIf(IsNull(Tabela_Clientes.Fields("Dt_Contrato")), "", Tabela_Clientes.Fields("Dt_Contrato"))

End Sub

Sub DesabilitaNavegacao()

   Cmd_Fechar.Enabled = False
   Cmd_Primeiro.Enabled = False
   Cmd_Anterior.Enabled = False
   Cmd_Proximo.Enabled = False
   Cmd_Ultimo.Enabled = False

End Sub

Sub HabilitaNavegacao()

   Cmd_Fechar.Enabled = True
   Cmd_Primeiro.Enabled = True
   Cmd_Anterior.Enabled = True
   Cmd_Proximo.Enabled = True
   Cmd_Ultimo.Enabled = True

End Sub

Private Sub Form_Terminate()

W_DB.Close
Tabela_Clientes.Close

End Sub

Private Sub Text_Bairro_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If

End Sub

Private Sub Text_Celular_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If

End Sub

Private Sub Text_Cep_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If

End Sub

Private Sub Text_Cidade_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If

End Sub

Private Sub Text_Cpf_Cnpj_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If CPFValido(Text_Cpf_Cnpj.Text) = False Then
      MsgBox ("CPF Invalido!!!")
   Else
      SendKeys "{TAB}"
   End If
End If

End Sub

Private Sub Text_E_Mail_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If

End Sub

Private Sub Text_End_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If

End Sub

Private Sub Text_Fone_Com_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If

End Sub

Private Sub Text_Fone_Resid_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If

End Sub

Private Sub Text_NomeCli_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If

End Sub

Private Sub Text_Uf_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If

End Sub

Sub Cad_Ed()

Combo_Uf.Clear
Combo_Uf.AddItem "AC"
Combo_Uf.AddItem "AL"
Combo_Uf.AddItem "AM"
Combo_Uf.AddItem "AP"
Combo_Uf.AddItem "BA"
Combo_Uf.AddItem "CE"
Combo_Uf.AddItem "DF"
Combo_Uf.AddItem "ES"
Combo_Uf.AddItem "GO"
Combo_Uf.AddItem "MA"
Combo_Uf.AddItem "MG"
Combo_Uf.AddItem "MS"
Combo_Uf.AddItem "MT"
Combo_Uf.AddItem "PA"
Combo_Uf.AddItem "PB"
Combo_Uf.AddItem "PE"
Combo_Uf.AddItem "PI"
Combo_Uf.AddItem "PR"
Combo_Uf.AddItem "RJ"
Combo_Uf.AddItem "RN"
Combo_Uf.AddItem "RS"
Combo_Uf.AddItem "RO"
Combo_Uf.AddItem "RR"
Combo_Uf.AddItem "SC"
Combo_Uf.AddItem "SE"
Combo_Uf.AddItem "SP"
Combo_Uf.AddItem "TO"

Combo_Zona.Clear
Combo_Zona.AddItem "INTERIOR"
Combo_Zona.AddItem "ABC"
Combo_Zona.AddItem "LITORAL"
Combo_Zona.AddItem "ZONA NORTE"
Combo_Zona.AddItem "ZONA LESTE"
Combo_Zona.AddItem "ZONA SUL"
Combo_Zona.AddItem "ZONA OESTE"

Call Lista_Empresa

Text_NomeCli.TabIndex = 1
Text_NomeCli.BackColor = vbYellow

Text_Cpf_Cnpj.TabIndex = 2
Text_Cpf_Cnpj.BackColor = vbYellow

Text_End.TabIndex = 3
Text_End.BackColor = vbYellow

Text_Bairro.TabIndex = 4
Text_Bairro.BackColor = vbYellow

Text_Cidade.TabIndex = 5
Text_Cidade.BackColor = vbYellow

Combo_Uf.TabIndex = 6
Combo_Uf.BackColor = vbYellow

Text_Cep.TabIndex = 7
Text_Cep.BackColor = vbYellow

Text_Ref.TabIndex = 8
Text_Ref.BackColor = vbYellow

Combo_Zona.TabIndex = 9
Combo_Zona.BackColor = vbYellow

Text_Fone_Resid.TabIndex = 10
Text_Fone_Resid.BackColor = vbYellow

Text_Fone_Com.TabIndex = 11
Text_Fone_Com.BackColor = vbYellow

Text_Celular.TabIndex = 12
Text_Celular.BackColor = vbYellow

Text_E_Mail.TabIndex = 13
Text_E_Mail.BackColor = vbYellow


Cmd_Salva.TabIndex = 15


End Sub
Sub Cad_Ed_Fim()

   Text_NomeCli.BackColor = vbWhite
   Text_Cpf_Cnpj.BackColor = vbWhite
   Text_End.BackColor = vbWhite
   Text_Bairro.BackColor = vbWhite
   Text_Cidade.BackColor = vbWhite
   Combo_Uf.BackColor = vbWhite
   Text_Cep.BackColor = vbWhite
   Text_Ref.BackColor = vbWhite
   Combo_Zona.BackColor = vbWhite
   Text_Fone_Resid.BackColor = vbWhite
   Text_Fone_Com.BackColor = vbWhite
   Text_Celular.BackColor = vbWhite
   Text_E_Mail.BackColor = vbWhite

End Sub

Sub Lista_Empresa()

If Tabela_Empresas.BOF = True And Tabela_Empresas.EOF = True Then
Else
    Tabela_Empresas.MoveFirst
    Combo_Solicitante.Clear
    Do While Not Tabela_Empresas.EOF
       DoEvents
       Combo_Solicitante.AddItem Tabela_Empresas.Fields("Tipo_Cliente")
       Tabela_Empresas.MoveNext
    Loop
End If

End Sub
