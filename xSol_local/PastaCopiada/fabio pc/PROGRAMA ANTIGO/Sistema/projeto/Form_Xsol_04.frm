VERSION 5.00
Begin VB.Form Form_Xsol_04 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   2475
   ClientTop       =   2820
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   7485
   Begin VB.Frame Frame 
      Caption         =   "Cadastramento de Técnico"
      Height          =   5415
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   7215
      Begin VB.TextBox Text_Cod_Tec 
         Height          =   285
         Left            =   1560
         TabIndex        =   33
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text_Nome_Tec 
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox Text_Dt_Nasc_Tec 
         Height          =   285
         Left            =   3480
         TabIndex        =   19
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text_Rg_Tec 
         Height          =   285
         Left            =   1560
         TabIndex        =   18
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox Text_Tec_CPF 
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox Text_End_Tec 
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         Top             =   3120
         Width           =   3975
      End
      Begin VB.TextBox Text_Cidade_Tec 
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Top             =   3600
         Width           =   2535
      End
      Begin VB.TextBox Text_Cep_Tec 
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox Text_Fone_Tec 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   4560
         Width           =   2295
      End
      Begin VB.TextBox Text_Celular_Tec 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   5040
         Width           =   2295
      End
      Begin VB.TextBox Text_Razao_Tec 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox Text_CNPJ_Tec 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Lbl_Cod_Tec 
         Alignment       =   1  'Right Justify
         Caption         =   "Cod. Técnico:"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label_Nome_Tec 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome do Técnico:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Lbl_Dt_Nasc_Tec 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt. de Nasc.:"
         Height          =   255
         Left            =   2400
         TabIndex        =   30
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Lbl_Rg_Tec 
         Alignment       =   1  'Right Justify
         Caption         =   "RG.:"
         Height          =   255
         Left            =   1080
         TabIndex        =   29
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Lbl_Tec_CPF 
         Alignment       =   1  'Right Justify
         Caption         =   "CPF.:"
         Height          =   255
         Left            =   960
         TabIndex        =   28
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Lbl_End_Tec 
         Alignment       =   1  'Right Justify
         Caption         =   "Endereço:"
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Lbl_Cid_Tec 
         Alignment       =   1  'Right Justify
         Caption         =   "Cidade:"
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Lbl_Cep_Tec 
         Alignment       =   1  'Right Justify
         Caption         =   "Cep.:"
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   4080
         Width           =   495
      End
      Begin VB.Label Lbl_Tec_Fone 
         Alignment       =   1  'Right Justify
         Caption         =   "Fone:"
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Lbl_Celular_Tec 
         Alignment       =   1  'Right Justify
         Caption         =   "Celular:"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   5040
         Width           =   615
      End
      Begin VB.Label Lbl_Razao_Tec 
         Alignment       =   1  'Right Justify
         Caption         =   "Razão Social:"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Lbl_CNPJ_Tec 
         Alignment       =   1  'Right Justify
         Caption         =   "CNPJ:"
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.CommandButton Cmd_Inclui 
      Caption         =   "&Incluir"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Salva 
      Caption         =   "&Salvar"
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Exclui 
      Caption         =   "E&xclui"
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Edita 
      Caption         =   "&Editar"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Primeiro 
      Caption         =   "<<   &Primeiro"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Anterior 
      Caption         =   "<  &Anterior"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Proximo 
      Caption         =   "Pr&oximo  >"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Ultimo 
      Caption         =   "&Último  >>"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Fechar 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   6720
      Width           =   1215
   End
End
Attribute VB_Name = "Form_Xsol_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public W_DB As Connection
Public Tabela_Tecnicos As Recordset
Public W_DBname As String

Private Sub Cmd_Anterior_Click()

If Tabela_Tecnicos.BOF = False Then
   Tabela_Tecnicos.MovePrevious
   Call Mostra_Registro
   If Tabela_Tecnicos.BOF = True Then
      Tabela_Tecnicos.MoveFirst
      Call Mostra_Registro
   End If
   Else
   If Tabela_Tecnicos.EOF Then
      MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
   Else
      Tabela_Tecnicos.MoveFirst
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
               
        'Text_Cpf_Cnpj.BackColor = vbYellow
        'Text_NomeCli.SetFocus
        
    ElseIf Cmd_Edita.Caption = "&Cancela" Then
        Cmd_Edita.Caption = "&Editar"
        
        Cmd_Salva.Enabled = False
        
        Call HabilitaNavegacao
        
        Cmd_Inclui.Enabled = True
        Cmd_Exclui.Enabled = True
        
        Call Trava_Registro
        
    End If

End Sub

Private Sub Cmd_Exclui_Click()

If MsgBox("Confirma Exclusão deste Registro ? ", vbYesNo, "Excluir Registros") = vbYes Then
        
   On Error Resume Next
        
   Tabela_Tecnicos.Delete
        
   Tabela_Tecnicos.MoveNext
   If Tabela_Tecnicos.EOF = True Then
      Tabela_Tecnicos.MoveLast
      If Tabela_Tecnicos.BOF = True Then
         Call LimpaControles
              MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
      End If
   End If
   ElseIf Tabela_Tecnicos.EOF = True And Tabela_Tecnicos.BOF = True Then
                MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
End If
Call Mostra_Registro

End Sub

Private Sub Cmd_Fechar_Click()

Unload Me

End Sub

Private Sub Cmd_Inclui_Click()

If Cmd_Inclui.Caption = "&Incluir" Then
   
   Cmd_Inclui.Caption = "&Cancela"
   Cmd_Salva.Enabled = True
   Cmd_Edita.Enabled = False
   Cmd_Exclui.Enabled = False
   
   Call DesabilitaNavegacao
   Call LimpaControles
   Call Destrava_Registro
   
   Text_Nome_Tec.SetFocus
   
ElseIf Cmd_Inclui.Caption = "&Cancela" Then
   
   Cmd_Inclui.Caption = "&Incluir"
   Cmd_Salva.Enabled = False
   Cmd_Edita.Enabled = True
   Cmd_Exclui.Enabled = True
   
   Call Trava_Registro
   Call HabilitaNavegacao
   Call Mostra_Registro

End If

End Sub

Private Sub Cmd_Primeiro_Click()

If Tabela_Tecnicos.BOF = False Then
   Tabela_Tecnicos.MoveFirst
   Call Mostra_Registro
ElseIf Tabela_Tecnicos.BOF = True _
   And Tabela_Tecnicos.EOF = True Then
   MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
End If

End Sub

Private Sub Cmd_Proximo_Click()

If Tabela_Tecnicos.EOF = False Then
   Tabela_Tecnicos.MoveNext
   Call Mostra_Registro
   If Tabela_Tecnicos.EOF Then
      Tabela_Tecnicos.MoveLast
      Call Mostra_Registro
   End If
Else
   If Tabela_Tecnicos.BOF Then
      MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
   Else
      Tabela_Tecnicos.MoveLast
      Call Mostra_Registro
   End If
End If

End Sub

Private Sub Cmd_Salva_Click()

If Trim(Text_Razao_Tec.Text) = "" Then
   MsgBox "É necessário digitar Razão Social! ", vbOKOnly, "X.Sol Arte em Aquecedores"
   Text_Razao_Tec.SetFocus
   Exit Sub
End If
    
If Trim(Text_CNPJ_Tec.Text) = "" Then
   MsgBox "É necessário digitar o CNPJ! ", vbOKOnly, "X.Sol Arte em Aquecedores"
   Text_CNPJ_Tec.SetFocus
   Exit Sub
End If

If Trim(Text_Cod_Tec.Text) = "" Then
   MsgBox "É necessário digitar Codigo! ", vbOKOnly, "X.Sol Arte em Aquecedores"
   Text_Cod_Tec.SetFocus
   Exit Sub
End If

If Trim(Text_Dt_Nasc_Tec.Text) = "" Then
   MsgBox "É necessário digitar a data de nascimento! ", vbOKOnly, "X.Sol Arte em Aquecedores"
   Text_Dt_Nasc_Tec.SetFocus
   Exit Sub
End If

If Trim(Text_Rg_Tec.Text) = "" Then
   MsgBox "É necessário digitar RG! ", vbOKOnly, "X.Sol Arte em Aquecedores"
   Text_Rg_Tec.SetFocus
   Exit Sub
End If

If Trim(Text_Tec_CPF.Text) = "" Then
   MsgBox "É necessário digitar CPF! ", vbOKOnly, "X.Sol Arte em Aquecedores"
   Text_Tec_CPF.SetFocus
   Exit Sub
End If

If Cmd_Inclui.Caption = "&Cancela" Then
   Tabela_Tecnicos.AddNew
End If

Call GravaRegistros

Tabela_Tecnicos.Update

Call Trava_Registro
Call HabilitaNavegacao
Call Mostra_Registro

If Cmd_Inclui.Caption = "&Cancela" Then
   Cmd_Inclui.Caption = "&Incluir"
End If
    
If Cmd_Edita.Caption = "&Cancela" Then
   Cmd_Edita.Caption = "&Editar"
End If

Cmd_Edita.Enabled = True
Cmd_Inclui.Enabled = True
Cmd_Exclui.Enabled = True

End Sub

Private Sub Cmd_Ultimo_Click()

If Tabela_Tecnicos.EOF = False Then
   Tabela_Tecnicos.MoveLast
   Call Mostra_Registro
ElseIf Tabela_Tecnicos.BOF = True _
  And Tabela_Tecnicos.EOF = True Then
  MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
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

Form_Xsol_04.Top = (Screen.Height - Form_Xsol_04.Height) / 2
Form_Xsol_04.Left = (Screen.Width - Form_Xsol_04.Width) / 2

Text_Razao_Tec.MaxLength = 50
Text_CNPJ_Tec.MaxLength = 25
Text_Nome_Tec.MaxLength = 70
Text_Cod_Tec.MaxLength = 10
Text_Dt_Nasc_Tec.MaxLength = 10
Text_Rg_Tec.MaxLength = 20
Text_Tec_CPF.MaxLength = 20
Text_End_Tec.MaxLength = 50
Text_Cidade_Tec.MaxLength = 50
Text_Cep_Tec.MaxLength = 11
Text_Fone_Tec.MaxLength = 25
Text_Celular_Tec.MaxLength = 25

If Right(App.Path, 1) = "\" Then
  W_DBname = App.Path & "dados.mdb"
 Else
  W_DBname = App.Path & "\" & "dados.mdb"
End If

Set W_DB = New Connection

Set Tabela_Tecnicos = New Recordset

W_DB.Mode = adModeReadWrite

W_DB.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"

Tabela_Tecnicos.Open "select * from Tecnico Order By Nome ", W_DB, adOpenStatic, adLockOptimistic

Form_Xsol_02.Caption = "Cadastro de Técnico Xsol"

Call Trava_Registro
Call Mostra_Registro

End Sub

Sub Trava_Registro()

Text_Razao_Tec.Locked = True
Text_CNPJ_Tec.Locked = True
Text_Nome_Tec.Locked = True
Text_Cod_Tec.Locked = True
Text_Dt_Nasc_Tec.Locked = True
Text_Rg_Tec.Locked = True
Text_Tec_CPF.Locked = True
Text_End_Tec.Locked = True
Text_Cidade_Tec.Locked = True
Text_Cep_Tec.Locked = True
Text_Fone_Tec.Locked = True
Text_Celular_Tec.Locked = True

End Sub

Sub Destrava_Registro()

Text_Razao_Tec.Locked = False
Text_CNPJ_Tec.Locked = False
Text_Nome_Tec.Locked = False
Text_Cod_Tec.Locked = False
Text_Dt_Nasc_Tec.Locked = False
Text_Rg_Tec.Locked = False
Text_Tec_CPF.Locked = False
Text_End_Tec.Locked = False
Text_Cidade_Tec.Locked = False
Text_Cep_Tec.Locked = False
Text_Fone_Tec.Locked = False
Text_Celular_Tec.Locked = False

End Sub

Private Sub Form_Terminate()

W_DB.Close
Tabela_Tecnicos.Close

End Sub

Sub Mostra_Registro()

If (Tabela_Tecnicos.EOF = True) Or (Tabela_Tecnicos.BOF = True) Then
   Exit Sub
End If

Text_Razao_Tec.Text = IIf(IsNull(Tabela_Tecnicos.Fields("Razao")), "", Tabela_Tecnicos.Fields("Razao"))
Text_CNPJ_Tec.Text = IIf(IsNull(Tabela_Tecnicos.Fields("CNPJ")), "", Tabela_Tecnicos.Fields("CNPJ"))
Text_Nome_Tec.Text = IIf(IsNull(Tabela_Tecnicos.Fields("Nome")), "", Tabela_Tecnicos.Fields("Nome"))
Text_Cod_Tec.Text = IIf(IsNull(Tabela_Tecnicos.Fields("Cod_Tec")), "", Tabela_Tecnicos.Fields("Cod_Tec"))
Text_Dt_Nasc_Tec.Text = IIf(IsNull(Tabela_Tecnicos.Fields("Dt_Nasc")), "", Tabela_Tecnicos.Fields("Dt_Nasc"))
Text_Rg_Tec.Text = IIf(IsNull(Tabela_Tecnicos.Fields("RG")), "", Tabela_Tecnicos.Fields("RG"))
Text_Tec_CPF.Text = IIf(IsNull(Tabela_Tecnicos.Fields("CPF")), "", Tabela_Tecnicos.Fields("CPF"))
Text_End_Tec.Text = IIf(IsNull(Tabela_Tecnicos.Fields("End")), "", Tabela_Tecnicos.Fields("End"))
Text_Cidade_Tec.Text = IIf(IsNull(Tabela_Tecnicos.Fields("Cidade")), "", Tabela_Tecnicos.Fields("Cidade"))
Text_Cep_Tec.Text = IIf(IsNull(Tabela_Tecnicos.Fields("Cep")), "", Tabela_Tecnicos.Fields("Cep"))
Text_Fone_Tec.Text = IIf(IsNull(Tabela_Tecnicos.Fields("Fone")), "", Tabela_Tecnicos.Fields("Fone"))
Text_Celular_Tec.Text = IIf(IsNull(Tabela_Tecnicos.Fields("Celular")), "", Tabela_Tecnicos.Fields("Celular"))

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

Private Sub GravaRegistros()
            
If Trim(Text_Razao_Tec.Text) = "" Then
   Tabela_Tecnicos.Fields("Razao").Value = Null
Else
   Tabela_Tecnicos.Fields("Razao").Value = Text_Razao_Tec.Text
End If
   
If Trim(Text_CNPJ_Tec.Text) = "" Then
   Tabela_Tecnicos.Fields("CNPJ").Value = Null
Else
   Tabela_Tecnicos.Fields("CNPJ").Value = Text_CNPJ_Tec.Text
End If
      
If Trim(Text_Nome_Tec.Text) = "" Then
   Tabela_Tecnicos.Fields("Nome").Value = Null
Else
   Tabela_Tecnicos.Fields("Nome").Value = Text_Nome_Tec.Text
End If
   
If Trim(Text_Dt_Nasc_Tec.Text) = "" Then
   Tabela_Tecnicos.Fields("Dt_Nasc").Value = Null
Else
   Tabela_Tecnicos.Fields("Dt_Nasc").Value = Text_Dt_Nasc_Tec.Text
End If
      
If Trim(Text_Rg_Tec.Text) = "" Then
   Tabela_Tecnicos.Fields("RG").Value = Null
Else
   Tabela_Tecnicos.Fields("RG").Value = Text_Rg_Tec.Text
End If
   
If Trim(Text_Tec_CPF.Text) = "" Then
   Tabela_Tecnicos.Fields("CPF").Value = Null
Else
   Tabela_Tecnicos.Fields("CPF").Value = Text_Tec_CPF.Text
End If
   
If Trim(Text_End_Tec.Text) = "" Then
   Tabela_Tecnicos.Fields("End").Value = Null
Else
   Tabela_Tecnicos.Fields("End").Value = Text_End_Tec.Text
End If

If Trim(Text_Cidade_Tec.Text) = "" Then
   Tabela_Tecnicos.Fields("Cidade").Value = Null
Else
   Tabela_Tecnicos.Fields("Cidade").Value = Text_Cidade_Tec.Text
End If

If Trim(Text_Cep_Tec.Text) = "" Then
   Tabela_Tecnicos.Fields("Cep").Value = Null
Else
   Tabela_Tecnicos.Fields("Cep").Value = Text_Cep_Tec.Text
End If

If Trim(Text_Fone_Tec.Text) = "" Then
   Tabela_Tecnicos.Fields("Cep").Value = Null
Else
   Tabela_Tecnicos.Fields("Cep").Value = Text_Fone_Tec.Text
End If

If Trim(Text_Fone_Tec.Text) = "" Then
   Tabela_Tecnicos.Fields("Fone").Value = Null
Else
   Tabela_Tecnicos.Fields("Fone").Value = Text_Fone_Tec.Text
End If
   
If Trim(Text_Celular_Tec.Text) = "" Then
   Tabela_Tecnicos.Fields("Celular").Value = Null
Else
   Tabela_Tecnicos.Fields("Celular").Value = Text_Celular_Tec.Text
End If
      
End Sub

Sub LimpaControles()

Text_Razao_Tec.Text = ""
Text_CNPJ_Tec.Text = ""
Text_Nome_Tec.Text = ""
Text_Dt_Nasc_Tec.Text = ""
Text_Rg_Tec.Text = ""
Text_Tec_CPF.Text = ""
Text_End_Tec.Text = ""
Text_Cidade_Tec.Text = ""
Text_Cep_Tec.Text = ""
Text_Fone_Tec.Text = ""
Text_Celular_Tec.Text = ""

End Sub
