VERSION 5.00
Begin VB.Form Form_Xsol_06 
   Caption         =   "Retorno de Atendimento Xsol Arte em Aquecedores"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd_Primeiro 
      Caption         =   "<<   &Primeiro"
      Height          =   495
      Left            =   120
      TabIndex        =   24
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Anterior 
      Caption         =   "<  &Anterior"
      Height          =   495
      Left            =   2040
      TabIndex        =   23
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Proximo 
      Caption         =   "Pr&oximo  >"
      Height          =   495
      Left            =   4080
      TabIndex        =   22
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Ultimo 
      Caption         =   "&Último  >>"
      Height          =   495
      Left            =   6120
      TabIndex        =   21
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Retorno de Atendimento"
      Height          =   4935
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   7335
      Begin VB.TextBox Text_Tecnico 
         Height          =   285
         Left            =   3120
         TabIndex        =   27
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox Text_Dt_Serv 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2640
         TabIndex        =   26
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text_Status 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   11
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox Text_NomeCli 
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox Text_Dt_Retorno 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2640
         TabIndex        =   9
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text_Serv_Solic 
         Height          =   285
         Left            =   2640
         TabIndex        =   8
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox Text_Valor 
         Height          =   285
         Left            =   2640
         TabIndex        =   7
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Text_Tecnico_Cod 
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox Combo_G_C 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2040
         Width           =   2295
      End
      Begin VB.ComboBox Combo_C 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox Text_Obs 
         Height          =   1095
         Left            =   2640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   3480
         Width           =   4455
      End
      Begin VB.Label Lbl_Dt_Serv 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt. Serviço:"
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label LblStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "Posição Ficha:"
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Lbl_Nome 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome do Cliente / Razão Social:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Lbl_Dt_Retorno 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt. Retorno:"
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Lbl_Cob_Garant 
         Alignment       =   1  'Right Justify
         Caption         =   "Garantia / Cobrança:"
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Lbl_Serv_Solic 
         Alignment       =   1  'Right Justify
         Caption         =   "Serviço Solicitado:"
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Lbl_Valor 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor:"
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Cod_Tec 
         Alignment       =   1  'Right Justify
         Caption         =   "Técnico:"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Cobrança:"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Lbl_Obs 
         Alignment       =   1  'Right Justify
         Caption         =   "Obs.:"
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   3480
         Width           =   375
      End
   End
   Begin VB.CommandButton Cmd_Finalizar 
      Caption         =   "&Finalizar"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Fechar 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
End
Attribute VB_Name = "Form_Xsol_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public W_DB As Connection
Public Tabela_Tecnicos As Recordset
Public Tabela_Servicos As Recordset
Public Tabela_Atendimento As Recordset
Public Tabela_Clientes As Recordset

Private Sub Cmd_Anterior_Click()

If Tabela_Atendimento.BOF = False Then
   Tabela_Atendimento.MovePrevious
   Call Mostra_Registro
   If Tabela_Atendimento.BOF = True Then
      Tabela_Atendimento.MoveFirst
      Call Mostra_Registro
   End If
   Else
   If Tabela_Atendimento.EOF Then
      MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
   Else
      Tabela_Atendimento.MoveFirst
      Call Mostra_Registro
   End If
End If

End Sub

Private Sub Cmd_Fechar_Click()

Unload Me

End Sub

Private Sub Cmd_Finalizar_Click()
If Tabela_Tecnicos.BOF = True And Tabela_Tecnicos.EOF = True Then
    MsgBox ("Não existe registro"), vbOKOnly, "X.Sol Arte em Aquecedores"
Else
   If "BAIXADO" = Tabela_Atendimento.Fields("Status") Then
       Text_Obs.Locked = True
       MsgBox ("Atendimento realizado"), vbOKOnly, "X.Sol Arte em Aquecedores"
   Else
      If Trim(Combo_G_C.Text) = "" Then
         MsgBox "É necessário selecionar!", vbOKOnly, "X.Sol Arte em Aquecedores"
         Combo_G_C.SetFocus
         Exit Sub
      End If
      If Trim(Combo_G_C.Text) = "COBRANÇA" And Trim(Combo_C.Text) = "" Then
         MsgBox "É necessário selecionar tipo de cobrança!", vbOKOnly, "X.Sol Arte em Aquecedores"
         Combo_C.SetFocus
         Exit Sub
      End If
      Call GravaRegistros
      Tabela_Atendimento.Update
      Call Mostra_Registro
   End If
End If

End Sub

Private Sub Cmd_Primeiro_Click()

If Tabela_Atendimento.BOF = False Then
   Tabela_Atendimento.MoveFirst
   Call Mostra_Registro
ElseIf Tabela_Atendimento.BOF = True _
   And Tabela_Atendimento.EOF = True Then
   MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
End If

End Sub

Private Sub Cmd_Proximo_Click()

If Tabela_Atendimento.EOF = False Then
   Tabela_Atendimento.MoveNext
   Call Mostra_Registro
   If Tabela_Atendimento.EOF Then
      Tabela_Atendimento.MoveLast
      Call Mostra_Registro
   End If
Else
   If Tabela_Atendimento.BOF Then
      MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
   Else
      Tabela_Atendimento.MoveLast
      Call Mostra_Registro
   End If
End If

End Sub

Private Sub Cmd_Ultimo_Click()

If Tabela_Atendimento.EOF = False Then
   Tabela_Atendimento.MoveLast
   Call Mostra_Registro
ElseIf Tabela_Atendimento.BOF = True _
  And Tabela_Atendimento.EOF = True Then
  MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
End If

End Sub

Private Sub Combo_G_C_Click()

If Combo_G_C.Text = "COBRANÇA" Then
    Call Listagem_C
    Combo_C.Enabled = True
Else
    Combo_C.Enabled = False
    Combo_C.Clear
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

Form_Xsol_06.Top = (Screen.Height - Form_Xsol_05.Height) / 2
Form_Xsol_06.Left = (Screen.Width - Form_Xsol_05.Width) / 2


Form_Xsol_05.Top = (Screen.Height - Form_Xsol_05.Height) / 2
Form_Xsol_05.Left = (Screen.Width - Form_Xsol_05.Width) / 2

Call Arquivos_Dados

Set W_DB = New Connection

Set Tabela_Tecnicos = New Recordset

W_DB.Mode = adModeReadWrite

W_DB.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"

Tabela_Tecnicos.Open "select * from Tecnico Order By Nome ", W_DB, adOpenStatic, adLockOptimistic

Set Tabela_Clientes = New Recordset
Tabela_Clientes.Open "select * from Clientes where Codigo=" & Cod_Cli, W_DB, adOpenStatic, adLockOptimistic

Set Tabela_Servicos = New Recordset
Tabela_Servicos.Open "select Tipo_Serv from Servico Order By Tipo_Serv ", W_DB, adOpenStatic, adLockOptimistic

Set Tabela_Atendimento = New Recordset

Tabela_Atendimento.Open "select * from Atendimento where Cod_Cliente=" & Cod_Cli & "and Status='ABERTO' or Status='PENDENTE'", W_DB, adOpenStatic, adLockOptimistic

Call Listagem_G_C
Combo_C.Enabled = False
Text_Dt_Retorno.Enabled = False
Text_Dt_Serv.Enabled = False
Text_Dt_Retorno.Text = Date
Call Mostra_Registro

End Sub

Sub Listagem_G_C()

Combo_G_C.Clear
Combo_G_C.AddItem "GARANTIA SOLETROL"
Combo_G_C.AddItem "GARANTIA LUCSOL"
Combo_G_C.AddItem "GARANTIA X - SOL"
Combo_G_C.AddItem "GARANTIA PRESTADORA"
Combo_G_C.AddItem "COBRANÇA"

End Sub

Sub Listagem_C()

Combo_C.Clear
Combo_C.AddItem "DINHEIRO"
Combo_C.AddItem "BOLETO"
Combo_C.AddItem "CHEQUE A VISTA"
Combo_C.AddItem "CHEQUE PARCELADO 2 VEZES"
Combo_C.AddItem "CHEQUE PARCELADO 3 VEZES"

End Sub

Sub Mostra_Registro()

If (Tabela_Atendimento.EOF = True) Or (Tabela_Atendimento.BOF = True) Then
   Exit Sub
End If

Text_NomeCli.Text = Nome_Cli_Atend
Text_Tecnico_Cod.Text = IIf(IsNull(Tabela_Atendimento.Fields("Cod_Tecnico")), "", Tabela_Atendimento.Fields("Cod_Tecnico"))
Call Mostra_Tecnico
Text_Serv_Solic = IIf(IsNull(Tabela_Atendimento.Fields("Tipo_Serv")), "", Tabela_Atendimento.Fields("Tipo_Serv"))
Text_Valor.Text = FormatCurrency(IIf(IsNull(Tabela_Atendimento.Fields("Valor_Serv")), "", Tabela_Atendimento.Fields("Valor_Serv")))
Text_Status.Text = IIf(IsNull(Tabela_Atendimento.Fields("Tipo_Serv")), "", Tabela_Atendimento.Fields("Status"))
Text_Dt_Serv = IIf(IsNull(Tabela_Atendimento.Fields("Dt_Atend_Serv")), "", Tabela_Atendimento.Fields("Dt_Atend_Serv"))

If "ABERTO" = Tabela_Atendimento.Fields("Status") Then
   Text_Obs.Locked = False
End If


End Sub

Sub Mostra_Tecnico()

Dim Cod_Tec As Integer
Dim X_Cod_Tec As Integer

X_Cod_Tec = IIf(IsNull(Tabela_Atendimento.Fields("Cod_Tecnico")), "", Tabela_Atendimento.Fields("Cod_Tecnico"))

If Tabela_Tecnicos.BOF = True And Tabela_Tecnicos.EOF = True Then
Else
    Tabela_Tecnicos.MoveFirst
    Do While Not Tabela_Tecnicos.EOF And (Cod_Tec <> X_Cod_Tec)
       DoEvents
       Text_Tecnico.Text = ""
       Cod_Tec = Tabela_Tecnicos.Fields("Codigo")
       Text_Tecnico = UCase(Tabela_Tecnicos.Fields("Nome"))
       Tabela_Tecnicos.MoveNext
    Loop
    
End If

End Sub

Sub GravaRegistros()

Tabela_Atendimento.Fields("Dt_Retorno").Value = Date
Tabela_Atendimento.Fields("Status").Value = "BAIXADO"

If Trim(Text_Obs.Text) = "" Then
   Tabela_Atendimento.Fields("Obs").Value = Null
Else
   Tabela_Atendimento.Fields("Obs").Value = Text_Obs.Text
End If

If Trim(Combo_G_C.Text) = "" Then
   Tabela_Atendimento.Fields("G_C").Value = Null
Else
   Tabela_Atendimento.Fields("G_C").Value = Combo_G_C.Text
End If

If Trim(Combo_C.Text) = "" Then
   Tabela_Atendimento.Fields("Tipo_Cobranca").Value = Null
Else
   Tabela_Atendimento.Fields("Tipo_Cobranca").Value = Combo_C.Text
End If

End Sub
