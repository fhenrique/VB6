VERSION 5.00
Begin VB.Form Form_Xsol_05 
   Caption         =   "Atendimento Xsol Arte em Aquecedores"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option_Inst_Res 
      Caption         =   "Instalação de Residência"
      Height          =   255
      Left            =   4560
      TabIndex        =   28
      Top             =   1920
      Width           =   2295
   End
   Begin VB.OptionButton Option_Inst_Pisc 
      Caption         =   "Instalação de Piscina"
      Height          =   195
      Left            =   4560
      TabIndex        =   27
      Top             =   2160
      Width           =   2175
   End
   Begin VB.OptionButton Option_Assistencia 
      Caption         =   "Assistência Técnica"
      Height          =   195
      Left            =   4560
      TabIndex        =   26
      Top             =   2400
      Width           =   2295
   End
   Begin VB.OptionButton Option_Pisc 
      Caption         =   "Vistoria Técnica - Piscina"
      Height          =   255
      Left            =   4560
      TabIndex        =   25
      Top             =   2880
      Width           =   2415
   End
   Begin VB.OptionButton Option_Res_Ind 
      Caption         =   "Vistoria Técnica - Resd. Ind"
      Height          =   195
      Left            =   4560
      TabIndex        =   24
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton Cmd_Print 
      Caption         =   "&Imprimir"
      Height          =   495
      Left            =   2040
      TabIndex        =   23
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text_NomeCli 
      Height          =   285
      Left            =   2640
      TabIndex        =   21
      Top             =   720
      Width           =   4575
   End
   Begin VB.CommandButton Cmd_Primeiro 
      Caption         =   "<<   &Primeiro"
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Anterior 
      Caption         =   "<  &Anterior"
      Height          =   495
      Left            =   2040
      TabIndex        =   19
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Proximo 
      Caption         =   "Pr&oximo  >"
      Height          =   495
      Left            =   4080
      TabIndex        =   18
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Ultimo 
      Caption         =   "&Último  >>"
      Height          =   495
      Left            =   6120
      TabIndex        =   17
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Inclui 
      Caption         =   "&Incluir"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Salva 
      Caption         =   "&Salvar"
      Height          =   495
      Left            =   2040
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Exclui 
      Caption         =   "E&xclui"
      Height          =   495
      Left            =   6120
      TabIndex        =   14
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Edita 
      Caption         =   "&Editar"
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text_Dt_Serv 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      TabIndex        =   12
      Top             =   1440
      Width           =   975
   End
   Begin VB.ComboBox Combo_Servico 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox Text_Cob_Garant 
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text_Serv_Solic 
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox Text_Valor 
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Cmd_Fechar 
      Caption         =   "Fechar"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ComboBox Combo_Tecnico 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   4695
   End
   Begin VB.TextBox Text_Tecnico_Cod 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   6600
      Picture         =   "Form_Xsol_05.frx":0000
      Top             =   1080
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Label Lbl_Nome 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome do Cliente / Razão Social:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Data Serv.:"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Lbl_Cob_Garant 
      Alignment       =   1  'Right Justify
      Caption         =   "Garantia / Cobrança:"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Lbl_Serv_Solic 
      Alignment       =   1  'Right Justify
      Caption         =   "Serviço Solicitado:"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Lbl_Valor 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor:"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Cod_Tec 
      Alignment       =   1  'Right Justify
      Caption         =   "Técnico:"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form_Xsol_05"
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
Public W_DBname As String

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

Private Sub Cmd_Edita_Click()

    If Cmd_Edita.Caption = "&Editar" Then
            
        Cmd_Edita.Caption = "&Cancela"
    
        Cmd_Salva.Enabled = True
        
        Call Lista_Servico
        Call Lista_Tecnico
        Call DesabilitaNavegacao
        
        Cmd_Inclui.Enabled = False
        Cmd_Exclui.Enabled = False
        
        Call Destrava_Registro
               
        
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
        
   Tabela_Atendimento.Delete
        
   Tabela_Atendimento.MoveNext
   If Tabela_Atendimento.EOF = True Then
      Tabela_Atendimento.MoveLast
      If Tabela_Atendimento.BOF = True Then
         Call LimpaControles
              MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
      End If
   End If
   ElseIf Tabela_Atendimento.EOF = True And Tabela_Atendimento.BOF = True Then
                MsgBox "Não há dados no Arquivo!", , "Sem Dados !"
End If
Call Mostra_Registro

End Sub

Private Sub Cmd_Fechar_Click()

Unload Me

End Sub

Private Sub Cmd_Inclui_Click()

If Cmd_Inclui.Caption = "&Incluir" Then
   Text_NomeCli.Text = Nome_Cli_Atend
   Cmd_Inclui.Caption = "&Cancela"
   Cmd_Salva.Enabled = True
   Cmd_Edita.Enabled = False
   Cmd_Exclui.Enabled = False
   
   Call DesabilitaNavegacao
   Call LimpaControles
   Call Destrava_Registro
   Call Lista_Tecnico
   Call Lista_Servico

   Combo_Tecnico.SetFocus
   
ElseIf Cmd_Inclui.Caption = "&Cancela" Then
   Text_NomeCli.Text = ""
   Cmd_Inclui.Caption = "&Incluir"
   Cmd_Salva.Enabled = False
   Cmd_Edita.Enabled = True
   Cmd_Exclui.Enabled = True
   
   Call Trava_Registro
   Call HabilitaNavegacao
   Call Mostra_Registro
   
   'Call Cad_Ed_Fim

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

Private Sub Cmd_Print_Click()

If Option_Res_Ind = True Then
   Call Ficha_Vistoria_Resid_Ind
   Call Orcamento
   Option_Res_Ind = False
   Option_Pisc = False
   Option_Assistencia = False
   Option_Inst_Pisc = False
   Option_Inst_Res = False
End If

If Option_Pisc = True Then
   Call Ficha_Vistoria_Pisc
   Call Orcamento
   Option_Res_Ind = False
   Option_Pisc = False
   Option_Assistencia = False
   Option_Inst_Pisc = False
   Option_Inst_Res = False
End If

If Option_Assistencia = True Then
   Call Ficha_Assistencia
   Option_Res_Ind = False
   Option_Pisc = False
   Option_Assistencia = False
   Option_Inst_Pisc = False
   Option_Inst_Res = False
End If

If Option_Inst_Pisc = True Then
   Call Ficha_Inst_Piscina
   Option_Res_Ind = False
   Option_Pisc = False
   Option_Assistencia = False
   Option_Inst_Pisc = False
   Option_Inst_Res = False
End If

If Option_Inst_Res = True Then
   Call Ficha_Inst_Res
   Option_Res_Ind = False
   Option_Pisc = False
   Option_Assistencia = False
   Option_Inst_Pisc = False
   Option_Inst_Res = False
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

Private Sub Cmd_Salva_Click()

Dim Dias As Integer

If Trim(Combo_Tecnico.Text) = "" Then
   MsgBox "É necessário digitar selecionar o técnico!", vbOKOnly, "X.Sol Arte em Aquecedores"
   Combo_Tecnico.SetFocus
   Exit Sub
End If

If Trim(Combo_Servico.Text) = "" Then
   MsgBox "É necessário digitar selecionar o Serviço!", vbOKOnly, "X.Sol Arte em Aquecedores"
   Combo_Tecnico.SetFocus
   Exit Sub
End If

Dias = DateDiff("d", Date, vData(Text_Dt_Serv))
If Dias < 0 Or Trim(Text_Dt_Serv.Text) = "" Then
  MsgBox "É necessário digitar um dia valido!", vbOKOnly, "X.Sol Arte em Aquecedores"
  Text_Dt_Serv.SetFocus
  Exit Sub
End If

If Cmd_Inclui.Caption = "&Cancela" Then
   Tabela_Atendimento.AddNew
End If

Call GravaRegistros
Tabela_Atendimento.Update

Call Trava_Registro
Call HabilitaNavegacao
Call Mostra_Registro

If Cmd_Inclui.Caption = "&Cancela" Then
   Cmd_Inclui.Caption = "&Incluir"
End If
    
If Cmd_Edita.Caption = "&Cancela" Then
   Cmd_Edita.Caption = "&Editar"
End If
     
'Call Cad_Ed_Fim

Cmd_Edita.Enabled = True
Cmd_Inclui.Enabled = True
Cmd_Exclui.Enabled = True
Cmd_Salva.Enabled = False

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

Form_Xsol_05.Top = (Screen.Height - Form_Xsol_05.Height) / 2
Form_Xsol_05.Left = (Screen.Width - Form_Xsol_05.Width) / 2

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

Set Tabela_Clientes = New Recordset
Tabela_Clientes.Open "select * from Clientes where Codigo=" & Cod_Cli, W_DB, adOpenStatic, adLockOptimistic

Set Tabela_Servicos = New Recordset
Tabela_Servicos.Open "select Tipo_Serv from Servico Order By Tipo_Serv ", W_DB, adOpenStatic, adLockOptimistic

Set Tabela_Atendimento = New Recordset

Tabela_Atendimento.Open "select * from Atendimento where Cod_Cliente=" & Cod_Cli, W_DB, adOpenStatic, adLockOptimistic



Call Trava_Registro
Call Mostra_Registro

Text_NomeCli.Locked = True

End Sub

Sub Lista_Tecnico()

Dim Cod_Tec As Integer
Dim Nome_Tec As String

If Tabela_Tecnicos.BOF = True And Tabela_Tecnicos.EOF = True Then
Else
    Tabela_Tecnicos.MoveFirst
    Combo_Tecnico.Clear
    Do While Not Tabela_Tecnicos.EOF
       DoEvents
       Cod_Tec = Tabela_Tecnicos.Fields("Codigo")
       Nome_Tec = Tabela_Tecnicos.Fields("Nome")
       Combo_Tecnico.AddItem (Cod_Tec) & ("   -   ") & (Nome_Tec)
       Tabela_Tecnicos.MoveNext
    Loop
End If

End Sub

Sub Lista_Servico()

If Tabela_Servicos.BOF = True And Tabela_Servicos.EOF = True Then
Else
    Tabela_Servicos.MoveFirst
    Combo_Servico.Clear
    Do While Not Tabela_Servicos.EOF
       DoEvents
       Combo_Servico.AddItem Tabela_Servicos.Fields("Tipo_Serv")
       Tabela_Servicos.MoveNext
    Loop
End If

End Sub

Private Sub GravaRegistros()

Dim Cod_Tec As Integer

Cod_Tec = CInt(Mid$(Combo_Tecnico.Text, 1, 3))
Tabela_Atendimento.Fields("Cod_Tecnico").Value = Cod_Tec

Tabela_Atendimento.Fields("Cod_Cliente").Value = Cod_Cli

If Trim(Combo_Servico.Text) = "" Then
   Tabela_Atendimento.Fields("Tipo_Serv").Value = Null
Else
   Tabela_Atendimento.Fields("Tipo_Serv").Value = Combo_Servico.Text
End If

If Trim(Val(Text_Valor.Text)) = "" Then
   Tabela_Atendimento.Fields("Valor_Serv").Value = Null
Else
   Tabela_Atendimento.Fields("Valor_Serv").Value = cvInt(Text_Valor.Text)
End If

If Trim(Val(Text_Dt_Serv.Text)) = "" Then
   Tabela_Atendimento.Fields("Dt_Atend_Serv").Value = Null
Else
   Tabela_Atendimento.Fields("Dt_Atend_Serv").Value = vData(Text_Dt_Serv.Text)
End If

Tabela_Atendimento.Fields("Dt_Atend").Value = Date

End Sub

Sub DesabilitaNavegacao()

   Cmd_Fechar.Enabled = False
   Cmd_Primeiro.Enabled = False
   Cmd_Anterior.Enabled = False
   Cmd_Proximo.Enabled = False
   Cmd_Ultimo.Enabled = False

End Sub


Sub Destrava_Registro()

Combo_Tecnico.Visible = True
Combo_Servico.Visible = True

Text_Tecnico_Cod.Locked = False
Text_Serv_Solic.Locked = False
Text_Dt_Serv.Locked = False
Text_Cob_Garant.Locked = False
Text_Valor.Locked = False

End Sub


Sub Trava_Registro()

Combo_Tecnico.Visible = False
Combo_Servico.Visible = False

Text_Tecnico_Cod.Locked = True
Text_Serv_Solic.Locked = True
Text_Dt_Serv.Locked = True
Text_Cob_Garant.Locked = True
Text_Valor.Locked = True

End Sub

Sub Mostra_Registro()

If (Tabela_Atendimento.EOF = True) Or (Tabela_Atendimento.BOF = True) Then
   Exit Sub
End If

Text_NomeCli.Text = Nome_Cli_Atend
Text_Tecnico_Cod.Text = IIf(IsNull(Tabela_Atendimento.Fields("Cod_Tecnico")), "", Tabela_Atendimento.Fields("Cod_Tecnico"))
Text_Serv_Solic = IIf(IsNull(Tabela_Atendimento.Fields("Tipo_Serv")), "", Tabela_Atendimento.Fields("Tipo_Serv"))
Text_Valor.Text = FormatCurrency(IIf(IsNull(Tabela_Atendimento.Fields("Valor_Serv")), "", Tabela_Atendimento.Fields("Valor_Serv")))
Text_Dt_Serv = IIf(IsNull(Tabela_Atendimento.Fields("Dt_Atend_Serv")), "", Tabela_Atendimento.Fields("Dt_Atend_Serv"))

End Sub

Sub LimpaControles()

Text_Tecnico_Cod.Text = ""
Text_Serv_Solic.Text = ""
Text_Dt_Serv.Text = ""
Text_Cob_Garant.Text = ""
Text_Valor.Text = ""

End Sub

Sub HabilitaNavegacao()

   Cmd_Fechar.Enabled = True
   Cmd_Primeiro.Enabled = True
   Cmd_Anterior.Enabled = True
   Cmd_Proximo.Enabled = True
   Cmd_Ultimo.Enabled = True

End Sub

Sub Ficha_Vistoria_Pisc()

Printer.PaperSize = 9
Printer.PaintPicture Image1.Picture, 0, 0
Printer.FontName = "Copperplate Gothic Light"
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.FontSize = 12
Printer.FontBold = True
Printer.Print "                                     FICHA DE VISTORIA TÉCNICA (PISCINA)"
Printer.FontBold = False
Printer.Print
Printer.FontSize = 10
Printer.Print
Printer.FontBold = True
Printer.Print Spc(3); "ITENS INDISPENSÁVEIS:"
Printer.FontBold = False
Printer.Print
Call Imprimi_Cliente
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "1) CONFORME CONSULTORIA REALIZADA NO LOCAL ONDE SERÁ REALIZADO O"
Printer.Print Spc(3); "SERVIÇO, DEFINE-SE NESTA, AS PROVIDÊNCIAS A SEREM TOMADAS PELO CLIENTE"
Printer.Print Spc(3); "E OBSERVAÇÕES ADICIONAIS"
Printer.FontBold = True
Printer.Print Spc(3); "DIMENSÕES E CARACTERÍSTCAS DA PISCINA:"
Printer.FontBold = False
Printer.Print Spc(3); "* PROFUNDIDADE:_______________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* COMPRIMENTO:________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* MATERIAL UTILIZANDO P/ CONFECÇÃO DA PISCINA:________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* É POSSIVEL INSTALAR RETORNO EM COBRE: (  ) SIM   (  ) NÃO"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* POSSUI AQUECIMENTO PARA PISC.: (  ) SIM (  ) NÃO TIPO:______________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* PISCINA COM RETORNO: (  ) COBRE (  ) PVC (  ) AQUATERM"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* ETAPA DA PISCINA: (  ) PROJETO (  ) ADAMEN. (  ) PRONTA (  ) REFORMA"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* RALO DE FUNDO: (  ) EXCLUSIVO (  ) COMPARTILHADO"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* PISCINA (  ) COBERTA (  ) DESCOBERTA"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* INCIDÊNCIA DE SOMBRA NOS COLETORES: (  ) ALTA (  ) MÉIDA (  ) BAIXA"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* CROQUI NO VERSO: (DETALHES)"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Spc(3); "ORIGEM DA ÁGUA:"
Printer.FontBold = False
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "(  ) COMPANHIA RESPONSÁVEL PELO ABASTECIMENTO"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "(  ) ÁGUA DE POÇO"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "(  ) CAMINHÃO TANQUE"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "(  ) ÁGUA TRATADA PELO CONDOMÍNIO"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Spc(3); "TIPO DE TRATAMENTO A SER ULTILIZADO:"
Printer.FontBold = False
Printer.Print Spc(3); "COMENTE:______________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Spc(3); "MESSÕES TÉCNICAS CONFORME CROQUI:"
Printer.FontBold = False
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* DISTÂNCIA HORIZONTAL ENTRE PISCINA./ COLETORES:_____________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* DISTÂNCIA HORIZONTAL ENTRE CASA DE MÁQUINAS. / PISCINA:_____________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* DISTÂNCIA HORIZONTAL ENTRE CASA DE MÁQUINAS. / COLETORES:___________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* DISTÂNCIA VERTICAL ENTRE PISCINA. / COLETORES: _____________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* DISTÂNCIA VERTICAL ENTRE CASA DE MÁQ. / COLETORES:__________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* DISTÂNCIA VERTICAL ENTRE CASA DE MÁQ. / PISCINA:____________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* LARGURA DO TELHADO:_________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* COMPRIMENTO DO TELHADO:_____________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* INCLINAÇÃO DO TELHADO:______________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* MATERIAL QUE É FEITO O TELHADO:_____________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* BÚSSOLA:____________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* ALTURA DOS DISPOSITIVOS DE RETORNO:_________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "TIPO DE CASA DE MÁQUINA:"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "(  ) ACIMA DA PISCINA"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "(  ) ABAIXO DA PISCINA:_______________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "(  ) NO NÍVEL DO FUNDO PISCINA:_______________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "(  ) MEIA ALTURA DA PISCINA:__________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Spc(3); "DESCRIÇÃO DO EQUIPAMENTO:"
Printer.FontBold = False
Printer.Print Spc(3); "* COLETOR"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "MODELO:_______________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "QUANTIDADE:___________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Spc(3); "OBSERVAÇÕES GERAIS:"
Printer.FontBold = False
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Spc(3); "IMPORTANTE:"
Printer.FontBold = False
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "LISTA DE MATERIAL SERÁ FEITA MEDIANTE O CROQUI E VISTORIA FEITA NO LOCAL."
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.Print
Printer.Print Spc(3); "____________________________________"
Printer.Print Spc(3); "        CLIENTE OU RESPONSAVEL      "
Printer.Print
Printer.Print Spc(3); "____________________________________"; Spc(5); "___________________,____ DE _______________DE__________"
Printer.Print Spc(3); "        VENDEDOR RESPONSÁVEL        "
Printer.Print
Printer.Print Spc(3); "____________________________________"
Printer.Print Spc(3); "        TÉCNICO RESPONSÁVEL         "
Printer.Print
Printer.EndDoc

End Sub


Sub Ficha_Vistoria_Resid_Ind()

Printer.PaperSize = 9
Printer.PaintPicture Image1.Picture, 0, 0
Printer.FontName = "Copperplate Gothic Light"
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Spc(30); "FICHA DE VISTORIA TÉCNICA"
Printer.FontBold = False
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Spc(3); "ITENS INDISPENSÁVEIS:"
Printer.FontBold = False
Printer.Print
Call Imprimi_Cliente
Printer.Print
Printer.Print Spc(3); "1) CONFORME CONSULTORIA REALIZADA NO LOCAL ONDE SERÁ REALIZADO O"
Printer.Print Spc(3); "SERVIÇO, DEFINE-SE NESTA, AS PROVIDÊNCIAS A SEREM TOMADA PELO CLIENTE"
Printer.Print Spc(3); "E OBSERVAÇÕES ADICIONAIS."
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "FLANGE (   )SIM (   )NÃO VALOR:_______________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "* SUPORTE P/ BOILER: (   ) SIM (   ) NÃO  VALOR:______________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "** SUPORTE P/ CAIXA DÁGUA: (   )SIM (   )NÃO  VALOR:__________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "SUPORTE P/ COLETORES: (   )SIM (   ) NÃO  VALOR:______________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "BARRILHETE: (   ) SIM (   ) NÃO QUANTID.____  VALOR:__________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "MATERIAL: (   ) SIM (   ) NÃO VALOR ATUALIZADO:_______________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "*** PARTE ELÉTRICA P/ SISTEMA: (   ) SIM (   ) NÃO VALOR:_____________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Spc(3); "OBSERVAÇÕES:"
Printer.FontBold = False
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(5); "° A INSTALAÇÃO DEVERÁ SER AGENDA SOMENTE APÓS A CONCLUSÃO DE"
Printer.Print Spc(3); "TODOS A PROVIDENCIAS A SEREM TOMADAS ACIMA."
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(5); "° CASO O TÉCNICO (NO DIA DA INSTALAÇÃO DO SISTEMA), CONSTATE ALGUMA"
Printer.Print Spc(3); "IRREGULARIDADE NAS PROVIDÊNCIAS ACIMA DESCRITAS E QUE O IMPEÇA DE"
Printer.Print Spc(3); "EXECUTAR OS SERVIÇOS, O CLIENTE ESTARÁ COMPREMETIDO A PAGAR COMO"
Printer.Print Spc(3); "MULTA O VALOR DE R$ 100,00, DESTINADO AO RESSARCIMENTO DAS DESPESAS"
Printer.Print Spc(3); "DE DESLOCAMENTO E HORA TÉCNICA."
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Spc(3); "INFORMAÇÕES REFERENTE À INSTALAÇÃO E SERVIÇOS ADICIONAIS:"
Printer.FontItalic = False
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "VALOR DA INSTALAÇÃO: R$______________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "SERVIÇOS ADICIONAIS: R$______________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "VALOR TOTAL: R$______________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print Spc(3); "________________________________________________________,_________DE _______________________DE _________________"
Printer.Print
Printer.Print
Printer.Print
Printer.Print Spc(3); "________________________________"; Spc(25); "__________________________________________________"
Printer.Print Spc(3); "  CLIENTE OU RESPONSÁVEL   "; Spc(27); "TÉCNICO RESPONSÁVEL PELO SERVIÇO"

Printer.Print
Printer.FontBold = True
Printer.Print Spc(3); "(*),(**),(***): CROQUIS RESPECTIVOS NO VERSO"
Printer.FontBold = False
Printer.EndDoc

End Sub

Sub Orcamento()

Printer.PaperSize = 9
Printer.PaintPicture Image1.Picture, 0, 0
Printer.FontName = "Copperplate Gothic Light"
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.FontSize = 12
Printer.FontBold = True
Printer.Print Spc(12); "ORÇAMENTO REFERENTE A (SERVIÇOS PRESTADOS)"
Printer.FontBold = False
Printer.FontSize = 10
Printer.Print
Call Imprimi_Cliente
Printer.FontBold = True
Printer.Print Spc(3); "EQUIPAMENTOS"
Printer.FontBold = False
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "° PISCINA:(   ) SIM (   ) NÃO"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "º BOILER MODELO___________ LITRAG._________ QUANT.__________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "° COLETORES: MODELO_________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "° CDT / CONTROLADORES SIMILARES: MARCA.______________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "MODELO________________________________QUANT.________________"
Printer.Print Spc(3); "SERVIÇOS ADICIONAIS"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "° COMENTE:____________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "VALOR:________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Spc(3); "VALORES REFERENTE A SERVIÇOS PRESTADOS"
Printer.FontBold = False
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "VALOR TOTAL:________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Spc(3); "PAGAMENTO"
Printer.FontBold = False
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "(   ) BOLETO BANCÁRIO"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "(   ) CHEQUE  (   ) À VISTA (   ) DIVIDIDO QUANT. VEZES:_____"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "(   ) DINHEIRO"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Spc(3); "OBSERVAÇÃO:"
Printer.FontBold = False
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "° TODOS OS SERVIÇOS PRESTADOS DEVERÃO SER PAGOS APÓS O TÉRMINO DO  SERVIÇO, TENDO"
Printer.Print Spc(3); "ASSIM GARANTIA DE _________, AFIM DE SOLUCIONAR ALGUM POSSIVEL PROBLEMA QUE VENHA POR "
Printer.Print Spc(3); "VENTURA APARECER."
Printer.Print Spc(3); "° A XSOL NÃO É RESPONSÁVEL PELA QUALIDADE E ENTREGA DOS PRODUTOS ADQUIRIDOS"
Printer.Print Spc(3); "DIRETAMENTE PELO CLIENTE"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.Print Spc(3); "_____________________________________________________,_______DE_______________________DE__________________"
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print Spc(3); "______________________________________"; Spc(27); "__________________________________________"
Printer.Print Spc(3); "          TÉCNICO RESPONSÁVEL"; Spc(32); "RESPONSÁVEL - CIENTE E DE ACORDO  "

Printer.EndDoc

End Sub

Sub Imprimi_Cliente()
'******** Variaveis do Cliente ******************************

Dim Print_NomeCli As String, Print_EndCli As String
Dim Print_BairroCli As String, Print_CepCli As String
Dim Print_CidadeCli As String, Print_TelResCli As String
Dim Print_TelComCli As String, Print_CelularCli As String
Dim Print_CadDataCli As Date

'*****************************************************************************************************************

'Print_CadDataCli = IIf(IsNull(Tabela_Clientes.Fields("Data_Emissao")), "", Tabela_Clientes.Fields("Data_Emissao"))

Print_NomeCli = IIf(IsNull(Tabela_Clientes.Fields("Nome_Cli")), "", Tabela_Clientes.Fields("Nome_Cli"))

Print_EndCli = IIf(IsNull(Tabela_Clientes.Fields("End")), "", Tabela_Clientes.Fields("End"))

Print_BairroCli = IIf(IsNull(Tabela_Clientes.Fields("Bairro")), "", Tabela_Clientes.Fields("Bairro"))

Print_CidadeCli = IIf(IsNull(Tabela_Clientes.Fields("Cidade")), "", Tabela_Clientes.Fields("Cidade"))

Print_CepCli = IIf(IsNull(Tabela_Clientes.Fields("Cep")), "", Tabela_Clientes.Fields("Cep"))

Print_TelResCli = IIf(IsNull(Tabela_Clientes.Fields("Fone_Res")), "", Tabela_Clientes.Fields("Fone_Res"))

Print_TelComCli = IIf(IsNull(Tabela_Clientes.Fields("Fone_Com")), "", Tabela_Clientes.Fields("Fone_Com"))

Print_CelularCli = IIf(IsNull(Tabela_Clientes.Fields("Fone_Celular")), "", Tabela_Clientes.Fields("Fone_Celular"))

'*****************************************************************************************************************

Printer.Print Spc(3); "NOME DO CLIENTE: "; Print_NomeCli
Printer.Print
Printer.Print Spc(3); "ENDEREÇO: "; Print_EndCli
Printer.Print
Printer.Print Spc(3); "BAIRRO: "; Print_BairroCli
Printer.Print
Printer.Print Spc(3); "REFERÊNCIA: ";
Printer.Print
Printer.Print Spc(3); "CEP: "; Print_CepCli
Printer.Print
Printer.Print Spc(3); "CIDADE: "; Print_CidadeCli
Printer.Print
Printer.Print Spc(3); "TEL. RES.: "; Print_TelResCli; Spc(10); " TEL. COM.: "; Print_TelComCli; Spc(10); " CELULAR: "; Print_CelularCli
Printer.Print

End Sub


Sub Ficha_Assistencia()

Printer.PaperSize = 9
Printer.PaintPicture Image1.Picture, 0, 0
Printer.FontName = "Copperplate Gothic Light"
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.FontSize = 12
Printer.FontBold = True
Printer.Print Spc(24); "FICHA DE ASSISTÊNCIA TÉCNICA"
Printer.FontBold = False
Printer.FontSize = 10
Printer.Print
Call Imprimi_Cliente
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "TIPO DE EQUIPAMENTO:"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "TIPO DE SERVIÇO:"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Spc(3); "RELATÓRIO:"
Printer.FontBold = False
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontSize = 6
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "IMPORTANTE:";
Printer.FontBold = False
Printer.Print ; "OS SERVIÇOS EXECUTADOS PODERÃO SER RECLAMADOS NO PRAZO DE "
Printer.Print Spc(3); "03 MESES A CONTAR DA DATA DE SUA REALIZAÇÃO."
Printer.Print Spc(3); "APÓS ESTE PRAZO ENTENDE-SE QUE O PROBLEMA FOI SANADO E NOVAS SOLICITAÇÕES SERÃO"
Printer.Print Spc(3); "COBRADAS."
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print Spc(3); "________________________________________________________,_________DE _______________________DE _________________"
Printer.Print
Printer.Print
Printer.Print
Printer.Print Spc(3); "________________________________"; Spc(25); "__________________________________________________"
Printer.Print Spc(3); "  CLIENTE OU RESPONSÁVEL   "; Spc(27); "TÉCNICO RESPONSÁVEL PELO SERVIÇO"
Printer.EndDoc

End Sub

Sub Ficha_Inst_Piscina()

Printer.PaperSize = 9
Printer.PaintPicture Image1.Picture, 0, 0
Printer.FontName = "Copperplate Gothic Light"
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.FontSize = 12
Printer.FontBold = True
Printer.Print Spc(27); "FICHA DE INSTALAÇÃO (PISCINA)"
Printer.FontBold = False
Printer.FontSize = 10
Printer.Print Spc(3); "EMPRESA EXECUTANTE:"
Printer.Print Spc(3); "NOME DO RESPONSÁVEL:"
Printer.Print
Call Imprimi_Cliente
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "Nº DE COLETORES:"
Printer.Print Spc(3); "VALOR DA INSTALAÇÃO:"
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.Print
Printer.FontBold = True
Printer.Print Spc(3); "IMPORTANTE:"
Printer.FontBold = False
Printer.Print Spc(3); "SEMPRE MANTER O SISTEMA LIGADO JUNTAMENTE COM A CAPA TÉRMICA, PARA QUE O MESMO "
Printer.Print Spc(3); "ATINJA A TEMPERATURA DE CONTRATO 25º à 29 GRAUS."
Printer.Print Spc(3); "QUALQUER TIPO DE MANUTENÇÃO NO SISTEMA DE AQUECIMENTO SOLAR DA PISCINA, É  "
Printer.Print Spc(3); "ACONSELHÁVEL SER EFETUADO APENAS POR PROFISSIONAIS QUALIFICADO DA X-SOL."
Printer.Print
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.Print
Printer.FontBold = True
Printer.Print Spc(3); "GARANTIA DE INSTALAÇÃO"
Printer.FontBold = False
Printer.Print Spc(3); "A INSTALAÇÃO POSSUI 1 ANO DE GARANTIA A PARTIR DA DATA DE CONCLUSÃO DA MESMA, E "
Printer.Print Spc(3); "COBREM:"
Printer.Print Spc(3); "* NÃO FUNCIONAMENTO DO SISTEMA, DECORRENTES A MOTIVOS DE INSTALAÇÃO."
Printer.Print Spc(3); "* VAZAMENTO NOS CASOS DE INTERLIGAÇÃO DO SISTEMA."
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.Print
Printer.Print
Printer.FontBold = True
Printer.Print Spc(3); "LEMBRETE:"
Printer.FontBold = False
Printer.Print Spc(3); "COM O OBJETIVO DE SEMPRE MANTER A BOA PERFORMANCE DO SISTEMA, EXISTE A NECESSIDADE"
Printer.Print Spc(3); "DE NO MÁXIMO 6 MESES MANUTENÇÃO PREVENTIVA."
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print Spc(3); "________________________________________________________,_________DE _______________________DE _________________"
Printer.Print
Printer.Print
Printer.Print
Printer.Print Spc(3); "________________________________"; Spc(25); "__________________________________________________"
Printer.Print Spc(3); "  CLIENTE OU RESPONSÁVEL   "; Spc(27); "TÉCNICO RESPONSÁVEL PELO SERVIÇO"
Printer.EndDoc

End Sub

Sub Ficha_Inst_Res()

Printer.PaperSize = 9
Printer.PaintPicture Image1.Picture, 0, 0
Printer.FontName = "Copperplate Gothic Light"
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.FontSize = 12
Printer.FontBold = True
Printer.Print Spc(31); "FICHA DE INSTALAÇÃO"
Printer.FontBold = False
Printer.Print
Call Imprimi_Cliente
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.Print
Printer.FontSize = 10
Printer.Print Spc(3); "EMPRESA EXECUTANTE:"
Printer.Print Spc(3); "NOME DO RESPONSÁVEL:"
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "TIPO DE EQUIPAMENTO:"
Printer.Print Spc(3); "VALOR DA INSTALAÇÃO:"
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"

Printer.FontBold = True
Printer.Print Spc(3); "IMPORTANTE:"
Printer.FontBold = False
Printer.Print Spc(3); "MANTER A VÁLVULA ANTI-CONGELAMENTO (DISPOSITIVO QUE PROTEGE OS COLETORES CONTRA"
Printer.Print Spc(3); "CONGELAMENTO EM GEADAS), SEMPRE ENERGIZADA. CASO NO ATO DA INSTALAÇÃO O TÉCNICO NÃO ENCONTRE A FIAÇÃO DE ALIMENTAÇÃO DA MESMA (ITEM PEDIDO EM VISTORIA TÉCNICA), O CLIENTE FICA CIENTE QUE O RETORNO SERÁ COBRADO UMA TAXA DE R$100,00. O MESMO SE APLICA COM OS DEMAIS ITENS SOLICITADOS NA VISTORIA."
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontBold = True
Printer.Print Spc(3); "GARANTIA DE INSTALAÇÃO"
Printer.FontBold = False
Printer.Print Spc(3); "A INSTALAÇÃO POSSUI 1 ANO DE GARANTIA A PARTIR DA DATA DE CONCLUSÃO DA MESMA, E "
Printer.Print Spc(3); "COBREM: "
Printer.Print Spc(3); "* NÃO FUNCIONAMENTO DO SISTEMA, DECORRENTES A MOTIVOS DE INSTALAÇÃO."
Printer.Print Spc(3); "* VAZAMENTO NOS CASOS DE INTERLIGAÇÃO DO SISTEMA."
Printer.Print
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.FontBold = True
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.Print
Printer.FontBold = True
Printer.Print Spc(3); "LEMBRETE:"
Printer.FontBold = False
Printer.Print Spc(3); "COM O OBJETIVO DE SEMPRE MANTER A BOA PERFORMANCE DO SISTEMA, EXISTE A NECESSIDADE "
Printer.Print Spc(3); "DE NO MÁXIMO 6 MESES SER REALIZADA UMA MANUTENÇÃO PREVENTIVA."
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print Spc(3); "________________________________________________________,_________DE _______________________DE _________________"
Printer.Print
Printer.Print
Printer.Print
Printer.Print Spc(3); "________________________________"; Spc(25); "__________________________________________________"
Printer.Print Spc(3); "  CLIENTE OU RESPONSÁVEL   "; Spc(27); "TÉCNICO RESPONSÁVEL PELO SERVIÇO"

End Sub
