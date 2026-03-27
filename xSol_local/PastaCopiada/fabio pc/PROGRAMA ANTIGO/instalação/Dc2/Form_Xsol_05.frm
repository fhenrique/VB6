VERSION 5.00
Begin VB.Form Form_Xsol_05 
   Caption         =   "Form1"
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
   Begin VB.OptionButton Option_Pisc 
      Caption         =   "Vistoria Técnica - Piscina"
      Height          =   255
      Left            =   3480
      TabIndex        =   25
      Top             =   3000
      Width           =   2415
   End
   Begin VB.OptionButton Option_Res_Ind 
      Caption         =   "Vistoria Técnica - Resd. Ind"
      Height          =   195
      Left            =   3480
      TabIndex        =   24
      Top             =   2760
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
      Left            =   5280
      Picture         =   "Form_Xsol_05.frx":0000
      Top             =   1920
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
End If

If Option_Pisc = True Then
   Call Ficha_Vistoria_Pisc
   Call Orcamento
   Option_Res_Ind = False
   Option_Pisc = False
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
Printer.FontName = "Courier"
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.FontSize = 14
Printer.Print Spc(3); "FICHA DE VISTORIA TÉCNICA (PISCINA):"
Printer.Print
Call Imprimi_Cliente
Printer.FontSize = 8
Printer.Print "ITENS INDISPENSÁVEIS"
Printer.Print "1) CONFORME CONSULTORIA REALIZADA NO LOCAL ONDE SERÁ REALIZADO O"
Printer.Print "SERVIÇO, DEFINE-SE NESTA, AS PROVIDÊNCIAS A SEREM TOMADAS PELO CLIENTE"
Printer.Print "E OBSERVAÇÕES ADICIONAIS"
Printer.FontBold = True
Printer.Print Spc(3); "DIMENSÕES E CARACTERÍSTCAS DA PISCINA:"
Printer.FontBold = False
Printer.Print Spc(3); "* PROFUNDIDADE:_____________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "* COMPRIMENTO:______________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "* MATERIAL UTILIZANDO P/ CONFECÇÃO DA PISCINA:______________________________________________"
Printer.Print
Printer.Print Spc(3); "* É POSSIVEL INSTALAR RETORNO EM COBRE: (  ) SIM   (  ) NÃO"
Printer.Print
Printer.Print Spc(3); "* POSSUI AQUECIMENTO PARA PISC.: (  ) SIM (  ) NÃO TIPO:____________________________________"
Printer.Print
Printer.Print Spc(3); "* PISCINA COM RETORNO: (  ) COBRE (  ) PVC (  ) AQUATERM"
Printer.Print
Printer.Print Spc(3); "* ETAPA DA PISCINA: (  ) PROJETO (  ) ADAMEN. (  ) PRONTA (  ) REFORMA"
Printer.Print
Printer.Print Spc(3); "* RALO DE FUNDO: (  ) EXCLUSIVO (  ) COMPARTILHADO"
Printer.Print
Printer.Print Spc(3); "* PISCINA (  ) COBERTA (  ) DESCOBERTA"
Printer.Print
Printer.Print Spc(3); "* INCIDÊNCIA DE SOMBRA NOS COLETORES: (  ) ALTA (  ) MÉIDA (  ) BAIXA"
Printer.Print
Printer.Print Spc(3); "* CROQUI NO VERSO: (DETALHES)"
Printer.Print
Printer.FontBold = True
Printer.Print Spc(3); "ORIGEM DA ÁGUA:"
Printer.FontBold = False
Printer.Print
Printer.Print Spc(3); "(  ) COMPANHIA RESPONSÁVEL PELO ABASTECIMENTO"
Printer.Print
Printer.Print Spc(3); "(  ) ÁGUA DE POÇO"
Printer.Print
Printer.Print Spc(3); "(  ) CAMINHÃO TANQUE"
Printer.Print
Printer.Print Spc(3); "(  ) ÁGUA TRATADA PELO CONDOMÍNIO"
Printer.Print
Printer.FontBold = True
Printer.Print Spc(3); "TIPO DE TRATAMENTO A SER ULTILIZADO:"
Printer.FontBold = False
Printer.Print Spc(3); "COMENTE:_____________________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "_____________________________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "MESSÕES TÉCNICAS CONFORME CROQUI:"
Printer.Print
Printer.Print Spc(3); "* DISTÂNCIA HORIZONTAL ENTRE PISCINA./ COLETORES:____________________________________________"
Printer.Print
Printer.Print Spc(3); "* DISTÂNCIA HORIZONTAL ENTRE CASA DE MÁQUINAS. / PISCINA:____________________________________"
Printer.Print
Printer.Print Spc(3); "* DISTÂNCIA HORIZONTAL ENTRE CASA DE MÁQUINAS. / COLETORES:__________________________________"
Printer.Print
Printer.Print Spc(3); "* DISTÂNCIA VERTICAL ENTRE PISCINA. / COLETORES: ____________________________________________"
Printer.Print
Printer.Print Spc(3); "* DISTÂNCIA VERTICAL ENTRE CASA DE MÁQ. / COLETORES:_________________________________________"
Printer.Print
Printer.Print Spc(3); "* DISTÂNCIA VERTICAL ENTRE CASA DE MÁQ. / PISCINA:___________________________________________"
Printer.Print
Printer.Print Spc(3); "* LARGURA DO TELHADO:________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "* COMPRIMENTO DO TELHADO:____________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "* INCLINAÇÃO DO TELHADO:_____________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "* MATERIAL QUE É FEITO O TELHADO:____________________________________________________________"
Printer.Print
Printer.Print Spc(3); "* BÚSSOLA:___________________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "* ALTURA DOS DISPOSITIVOS DE RETORNO:________________________________________________________"
Printer.Print
Printer.Print Spc(3); "TIPO DE CASA DE MÁQUINA:"
Printer.Print
Printer.Print Spc(3); "(  ) ACIMA DA PISCINA"
Printer.Print
Printer.Print Spc(3); "(  ) ABAIXO DA PISCINA:______________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "(  ) NO NÍVEL DO FUNDO PISCINA:______________________________________________________________"
Printer.Print
Printer.Print Spc(3); "(  ) MEIA ALTURA DA PISCINA:_________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "DESCRIÇÃO DO EQUIPAMENTO:"
Printer.Print Spc(3); "* COLETOR"
Printer.Print
Printer.Print Spc(3); "MODELO:"
Printer.Print
Printer.Print Spc(3); "QUANTIDADE:"
Printer.Print
Printer.Print Spc(3); "OBSERVAÇÕES GERAIS:"
Printer.Print
Printer.Print Spc(3); "______________________________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "______________________________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "______________________________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "______________________________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "______________________________________________________________________________________________"
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
Printer.FontName = "Courier"
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.FontSize = 14
Printer.Print Spc(3); "FICHA DE VISTORIA TÉCNICA"
Printer.Print
Printer.FontSize = 8
Printer.FontItalic = True
Printer.Print Spc(3); "ITENS INDISPENSÁVEIS:"
Printer.FontItalic = False
Printer.Print
Call Imprimi_Cliente
Printer.Print Spc(3); "1) CONFORME CONSULTORIA REALIZADA NO LOCAL ONDE SERÁ REALIZADO O"
Printer.Print Spc(3); "SERVIÇO, DEFINE-SE NESTA, AS PROVIDÊNCIAS A SEREM TOMADA PELO CLIENTE"
Printer.Print Spc(3); "E OBSERVAÇÕES ADICIONAIS."
Printer.Print
Printer.Print Spc(3); "FLANGE (   )SIM (   )NÃO VALOR:______________________________________________________"
Printer.Print
Printer.Print Spc(3); "* SUPORTE P/ BOILER: (   ) SIM (   ) NÃO  VALOR:_____________________________________"
Printer.Print
Printer.Print Spc(3); "** SUPORTE P/ CAIXA DÁGUA: (   )SIM (   )NÃO  VALOR:_________________________________"
Printer.Print
Printer.Print Spc(3); "SUPORTE P/ COLETORES: (   )SIM (   ) NÃO  VALOR:_____________________________________"
Printer.Print
Printer.Print Spc(3); "BARRILHETE: (   ) SIM (   ) NÃO QUANTID.____  VALOR:_________________________________"
Printer.Print
Printer.Print Spc(3); "MATERIAL: (   ) SIM (   ) NÃO VALOR ATUALIZADO:______________________________________"
Printer.Print
Printer.Print Spc(3); "*** PARTE ELÉTRICA P/ SISTEMA: (   ) SIM (   ) NÃO VALOR:____________________________"
Printer.Print
Printer.Print Spc(3); "OBSERVAÇÕES:"
Printer.Print Spc(5); "° A INSTALAÇÃO DEVERÁ SER AGENDA SOMENTE APÓS A CONCLUSÃO DE"
Printer.Print Spc(3); "TODOS A PROVIDENCIAS A SEREM TOMADAS ACIMA."
Printer.Print
Printer.Print Spc(5); "° CASO O TÉCNICO (NO DIA DA INSTALAÇÃO DO SISTEMA), CONSTATE ALGUMA"
Printer.Print Spc(3); "IRREGULARIDADE NAS PROVIDÊNCIAS ACIMA DESCRITAS E QUE O IMPEÇA DE"
Printer.Print Spc(3); "EXECUTAR OS SERVIÇOS, O CLIENTE ESTARÁ COMPREMETIDO A PAGAR COMO"
Printer.Print Spc(3); "MULTA O VALOR DE R$ 100,00, DESTINADO AO RESSARCIMENTO DAS DESPESAS"
Printer.Print Spc(3); "DE DESLOCAMENTO E HORA TÉCNICA."
Printer.Print
Printer.FontItalic = True
Printer.Print Spc(3); "INFORMAÇÕES REFERENTE À INSTALAÇÃO E SERVIÇOS ADICIONAIS:"
Printer.FontItalic = True
Printer.Print
Printer.Print Spc(3); "VALOR DA INSTALAÇÃO: R$______________"
Printer.Print
Printer.Print Spc(3); "SERVIÇOS ADICIONAIS: R$______________"
Printer.Print
Printer.Print Spc(3); "VALOR TOTAL: R$______________"
Printer.Print
Printer.Print
Printer.Print
Printer.Print Spc(3); "____________________________,_______DE__________________________DE__________________"
Printer.Print
Printer.Print
Printer.Print Spc(3); "________________________________________"; Spc(3); "________________________________________"
Printer.Print
Printer.Print Spc(3); "         CLIENTE OU RESPONSÁVEL         "; Spc(3); "  RESPONSÁVEL PELA EXECUÇÃO DO SERVIÇO  "
Printer.Print
Printer.Print Spc(3); "(*),(**),(***): CROQUIS RESPECTIVOS NO VERSO"
Printer.EndDoc

End Sub

Sub Orcamento()

Printer.PaperSize = 9
Printer.PaintPicture Image1.Picture, 0, 0
Printer.FontName = "Courier"
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.FontSize = 14
Printer.FontItalic = True
Printer.Print Spc(3); "ORÇAMENTO REFERENTE A (SERVIÇOS PRESTADOS)"
Printer.FontItalic = False
Printer.Print
Printer.FontSize = 8
Call Imprimi_Cliente
Printer.Print Spc(3); "EQUIPAMENTOS"
Printer.Print
Printer.Print Spc(3); "° PISCINA:(   ) SIM (   ) NÃO"
Printer.Print
Printer.Print Spc(3); "º BOILER MODELO___________ LITRAG._________ QUANT.__________"
Printer.Print
Printer.Print Spc(3); "° COLETORES: MODELO_________________________________________"
Printer.Print
Printer.Print Spc(3); "° CDT / CONTROLADORES SIMILARES: MARCA.______________________"
Printer.Print
Printer.Print Spc(3); "MODELO________________________________QUANT.________________"
Printer.Print Spc(3); "SERVIÇOS ADICIONAIS"
Printer.Print
Printer.Print Spc(3); "° COMENTE:__________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "____________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "____________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "____________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "____________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "____________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "____________________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "VALOR:______________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "VALORES REFERENTE A SERVIÇOS PRESTADOS"
Printer.Print
Printer.Print Spc(3); "VALOR TOTAL:________________________________________________________________"
Printer.Print
Printer.Print Spc(3); "PAGAMENTO"
Printer.Print
Printer.Print Spc(3); "(   ) BOLETO BANCÁRIO"
Printer.Print
Printer.Print Spc(3); "(   ) CHEQUE  (   ) À VISTA (   ) DIVIDIDO QUANT. VEZES:_____"
Printer.Print
Printer.Print Spc(3); "(   ) DINHEIRO"
Printer.Print
Printer.Print Spc(3); "OBSERVAÇÃO:"
Printer.Print
Printer.Print Spc(3); "° TODOS OS SERVIÇOS PRESTADOS DEVERÃO SER PAGOS APÓS O "
Printer.Print Spc(3); "TÉRMINO DO SERVIÇO, TENDO ASSIM GARANTIA DE _________, AFIM DE "
Printer.Print Spc(3); "SOLUCIONAR ALGUM POSSIVEL PROBLEMA QUE VENHA POR VENTURA APARECER."
Printer.Print Spc(3); "° A XSOL NÃO É RESPONSÁVEL PELA QUALIDADE E ENTREGA DOS"
Printer.Print Spc(3); "PRODUTOS ADQUIRIDOS DIRETAMENTE PELO CLIENTE"
Printer.Print
Printer.Print Spc(3); "________________________,_______DE_______________________DE__________________"
Printer.Print
Printer.Print Spc(3); "______________________________________"; Spc(3); "______________________________________"
Printer.Print Spc(3); "          TÉCNICO RESPONSÁVEL         "; Spc(3); "   RESPONSÁVEL - CIENTE E DE ACORODO  "
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

Print_CadDataCli = IIf(IsNull(Tabela_Clientes.Fields("Dt_Emissao")), "", Tabela_Clientes.Fields("Dt_Emissao"))

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
