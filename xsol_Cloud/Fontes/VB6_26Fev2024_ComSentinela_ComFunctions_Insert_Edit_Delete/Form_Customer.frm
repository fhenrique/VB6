VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_Customer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   10170
   ClientLeft      =   1290
   ClientTop       =   825
   ClientWidth     =   17370
   Icon            =   "Form_Customer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   17370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Caption         =   "Relatório de Clientes"
      Height          =   975
      Left            =   4005
      TabIndex        =   60
      Top             =   9075
      Width           =   3810
      Begin VB.OptionButton optnPorCliente 
         Caption         =   "Por Cliente"
         Height          =   195
         Left            =   1425
         TabIndex        =   66
         Top             =   630
         Width           =   1125
      End
      Begin VB.OptionButton optRelClientesArquivoMorto 
         Caption         =   "Arquivo morto"
         Height          =   195
         Left            =   1425
         TabIndex        =   64
         Top             =   315
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optRelClientesPorCidade 
         Caption         =   "Por Cidade"
         Height          =   195
         Left            =   165
         TabIndex        =   63
         Top             =   645
         Width           =   1125
      End
      Begin VB.OptionButton optRelClientesGeral 
         Caption         =   "Geral"
         Height          =   195
         Left            =   150
         TabIndex        =   62
         Top             =   315
         Width           =   885
      End
      Begin VB.CommandButton btnRelatorioGerar 
         Caption         =   "Gerar"
         Height          =   495
         Left            =   2850
         TabIndex        =   61
         Top             =   315
         Width           =   795
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Exibição"
      Height          =   900
      Left            =   15375
      TabIndex        =   56
      Top             =   9180
      Width           =   1935
      Begin VB.OptionButton Option1 
         Caption         =   "Arquivo morto"
         Height          =   375
         Left            =   150
         TabIndex        =   58
         Top             =   495
         Width           =   1710
      End
      Begin VB.OptionButton optnTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   150
         TabIndex        =   57
         Top             =   255
         Value           =   -1  'True
         Width           =   1530
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Equipamentos (Duplo clique na grade para abrir o cadastro de equipamentos)"
      Height          =   2535
      Left            =   105
      TabIndex        =   50
      Top             =   6480
      Width           =   9015
      Begin MSFlexGridLib.MSFlexGrid Grid_Equipment 
         Height          =   2265
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   3995
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Endereços (Duplo clique na grade para abrir o cadastro de endereços)"
      Height          =   2640
      Left            =   120
      TabIndex        =   48
      Top             =   3735
      Width           =   9015
      Begin MSFlexGridLib.MSFlexGrid Grid_Address 
         Height          =   2265
         Left            =   90
         TabIndex        =   49
         ToolTipText     =   "Duplo clique para abrir o cadastro de endereços"
         Top             =   285
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   3995
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   7890
      TabIndex        =   47
      Top             =   9075
      Width           =   1230
      Begin VB.CommandButton btnFechar 
         Caption         =   "Fechar"
         Height          =   290
         Left            =   150
         TabIndex        =   65
         Top             =   600
         Width           =   885
      End
      Begin VB.CommandButton Cmd_Serv 
         Caption         =   "A&tende"
         Height          =   290
         Left            =   150
         TabIndex        =   52
         Top             =   195
         Width           =   885
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Edição"
      Height          =   975
      Left            =   105
      TabIndex        =   42
      Top             =   9075
      Width           =   3840
      Begin VB.CommandButton Cmd_Exclui 
         BackColor       =   &H008080FF&
         Caption         =   "E&xcluir"
         Height          =   495
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   270
         Width           =   750
      End
      Begin VB.CommandButton Cmd_Edita 
         Caption         =   "&Editar"
         Height          =   495
         Left            =   1095
         TabIndex        =   45
         Top             =   285
         Width           =   780
      End
      Begin VB.CommandButton Cmd_Salva 
         Caption         =   "&Salvar"
         Height          =   495
         Left            =   2025
         TabIndex        =   44
         Top             =   270
         Width           =   765
      End
      Begin VB.CommandButton Cmd_Inclui 
         BackColor       =   &H0080FF80&
         Caption         =   "&Incluir"
         Height          =   495
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   285
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ordenação"
      Height          =   990
      Left            =   15360
      TabIndex        =   35
      Top             =   8190
      Width           =   1335
      Begin VB.OptionButton Opt_Sort_Name 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   630
         Width           =   990
      End
      Begin VB.OptionButton Opt_Sort_Code 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pesquisa"
      Height          =   1860
      Left            =   9240
      TabIndex        =   28
      Top             =   8205
      Width           =   6015
      Begin VB.TextBox Txt_Search_Address 
         Height          =   285
         Left            =   960
         TabIndex        =   53
         Top             =   1440
         Width           =   4215
      End
      Begin VB.CommandButton But_Search_Clean 
         Height          =   495
         Left            =   5280
         Picture         =   "Form_Customer.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Limpar"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton But_Search 
         Height          =   495
         Left            =   5280
         Picture         =   "Form_Customer.frx":096E
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Pesquisar"
         Top             =   1230
         Width           =   615
      End
      Begin VB.TextBox Txt_Search_Name 
         Height          =   285
         Left            =   960
         TabIndex        =   34
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox Txt_Search_CPF 
         Height          =   285
         Left            =   960
         TabIndex        =   32
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Txt_Search_Code 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   29
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Lbl_Search_Address 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "End:"
         Height          =   195
         Left            =   405
         TabIndex        =   54
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label Lbl_Search_Name 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   270
         TabIndex        =   33
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Lbl_Search_CPF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CPF:"
         Height          =   195
         Left            =   390
         TabIndex        =   31
         Top             =   720
         Width           =   345
      End
      Begin VB.Label Lbl_Search_Code 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   195
         TabIndex        =   30
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.CommandButton But_Close 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16800
      Picture         =   "Form_Customer.frx":0A51
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8310
      UseMaskColor    =   -1  'True
      Width           =   525
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de Clientes"
      Height          =   8100
      Left            =   9240
      TabIndex        =   25
      Top             =   45
      Width           =   8055
      Begin MSFlexGridLib.MSFlexGrid Grid_Customer 
         Height          =   7335
         Left            =   135
         TabIndex        =   26
         Top             =   195
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   12938
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         Redraw          =   -1  'True
         SelectionMode   =   1
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "ARQUIVO MORTO"
         Height          =   225
         Left            =   3555
         TabIndex        =   69
         Top             =   7650
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Legenda:"
         Height          =   240
         Left            =   165
         TabIndex        =   68
         Top             =   7650
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "EXCLUÍDO MANUALMENTE"
         Height          =   225
         Left            =   960
         TabIndex        =   67
         Top             =   7650
         Width           =   2310
      End
   End
   Begin VB.Frame Frame_Cliente 
      Caption         =   "Dados do Cliente"
      Height          =   3630
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   9015
      Begin VB.CommandButton cmdJustificativa 
         Caption         =   "Ver justificativa"
         Height          =   240
         Left            =   7050
         TabIndex        =   59
         Top             =   3225
         Width           =   1215
      End
      Begin VB.CheckBox chkArquivoMorto 
         Caption         =   "Arquivo Morto"
         Height          =   270
         Left            =   5400
         TabIndex        =   55
         Top             =   3210
         Width           =   1605
      End
      Begin VB.TextBox Txt_Next_Visit 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   41
         Top             =   3150
         Width           =   1215
      End
      Begin VB.TextBox Text_Cod 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   24
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Text_Registro 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   8160
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text_Cnpj 
         Height          =   285
         Left            =   2880
         TabIndex        =   8
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Text_Dt_Contrato 
         Height          =   285
         Left            =   5400
         TabIndex        =   7
         Top             =   2790
         Width           =   1095
      End
      Begin VB.TextBox Text_Contrato 
         Height          =   285
         Left            =   2880
         TabIndex        =   5
         Top             =   2790
         Width           =   1215
      End
      Begin VB.ComboBox Combo_Solicitante 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2430
         Width           =   4335
      End
      Begin VB.TextBox Text_Solicitante 
         Height          =   285
         Left            =   2880
         TabIndex        =   15
         Top             =   2460
         Width           =   4215
      End
      Begin VB.TextBox Text_NomeCli 
         Height          =   285
         Left            =   2880
         TabIndex        =   2
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox Text_Dt_Emis 
         Height          =   285
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text_Cpf 
         Height          =   285
         Left            =   2880
         TabIndex        =   6
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text_E_Mail 
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Top             =   2070
         Width           =   2535
      End
      Begin VB.TextBox Text_Atendente 
         Height          =   285
         Left            =   6240
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Lbl_Dt_Visita 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Próxima Visita:"
         Height          =   195
         Left            =   1365
         TabIndex        =   40
         Top             =   3150
         Width           =   1410
      End
      Begin VB.Label Lbl_Cod 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código do Cliente:"
         Height          =   195
         Left            =   1485
         TabIndex        =   23
         Top             =   1320
         Width           =   1290
      End
      Begin VB.Label Lbl_Registro 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Registro:"
         Height          =   195
         Left            =   7425
         TabIndex        =   21
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ou"
         Height          =   195
         Left            =   2235
         TabIndex        =   20
         Top             =   720
         Width           =   180
      End
      Begin VB.Label Lbl_Cnpj 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cnpj:"
         Height          =   195
         Left            =   2415
         TabIndex        =   19
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Lbl_Dt_Contrato 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dt. Contrato:"
         Height          =   195
         Left            =   4395
         TabIndex        =   18
         Top             =   2790
         Width           =   900
      End
      Begin VB.Label Lbl_Contrato 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
         Height          =   195
         Left            =   2145
         TabIndex        =   17
         Top             =   2790
         Width           =   600
      End
      Begin VB.Label Llb_Solitante 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Empresa solicitante:"
         Height          =   195
         Left            =   1380
         TabIndex        =   16
         Top             =   2505
         Width           =   1410
      End
      Begin VB.Label Lbl_Nome 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nome do Cliente / Razão Social:"
         Height          =   195
         Left            =   450
         TabIndex        =   14
         Top             =   1680
         Width           =   2325
      End
      Begin VB.Label Lbl_Dt_Emis 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Emissão:"
         Height          =   195
         Left            =   1755
         TabIndex        =   13
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Lbl_Cpf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cpf:"
         Height          =   195
         Left            =   2490
         TabIndex        =   12
         Top             =   600
         Width           =   285
      End
      Begin VB.Label Lbl_E_Mail 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "E-Mail:"
         Height          =   195
         Left            =   2295
         TabIndex        =   11
         Top             =   2070
         Width           =   480
      End
      Begin VB.Label Lbl_Atendente 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Atendente:"
         Height          =   195
         Left            =   5355
         TabIndex        =   10
         Top             =   600
         Width           =   780
      End
   End
End
Attribute VB_Name = "Form_Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public W_DB As Connection
Private conn As ADODB.Connection
Public Tabela_Empresas As Recordset
Public Atendente As String
Public codCliente As String
Public strJustificativaArquivoMorto As String
Public ehProspect As Boolean
Public strFoneRes As String
Public strFoneCom As String
Public strFoneCel As String
Public excluindo As Boolean


Private Sub executaRelatorio(dbname As String, titulo1Rel As String)

    'RelClientesAtendimentos.DAODataControl1.DatabaseName = dbname

    'RelClientesAtendimentos.lblTipoRelatorio.Caption = titulo1Rel
    'RelClientesAtendimentos.DAODataControl1.RecordSource = "Select cli.Codigo, cli.Data_Emissao, cli.Nome_Cli, cli.CPF, cli.CNPJ, cli.Fone_Celular, cli.Fone_Res, cli.Fone_Com, cli.E_Mail, cli.Aquivo_Morto,end.Endereco, end.Bairro, end.Ref, end.CEP From Clientes cli Inner join Endereco end on end.Cod_Cliente = cli.Codigo Where Aquivo_Morto = 'N' Order by 3"
    'RelClientesAtendimentos.Show (1)
    
End Sub

Private Sub executaRelatorioArquivoMorto(dbname As String, titulo1Rel As String)

    relClientes.DAODataControl1.DatabaseName = dbname

    relClientes.lblTipoRelatorio.Caption = titulo1Rel
    relClientes.DAODataControl1.RecordSource = "Select cli.Codigo, cli.Data_Emissao, cli.Nome_Cli, cli.CPF, cli.CNPJ, cli.Fone_Celular, cli.Fone_Res, cli.Fone_Com, cli.E_Mail, cli.Aquivo_Morto, cli.Aquivo_Morto_Justificativa, end.Endereco, end.Bairro, end.Ref, end.CEP From Clientes cli Inner join Endereco end on end.Cod_Cliente = cli.Codigo Where Aquivo_Morto = 'S' Order by 3"
    relClientes.Show (1)
    
End Sub

Private Sub executaRelatorioAtendimento(dbname As String, titulo1Rel As String)

    relAtendimentosPorCliente.DAODataControl1.DatabaseName = dbname

    'relAtendimentosPorCliente.lblTipoRelatorio.Caption = titulo1Rel
    relAtendimentosPorCliente.DAODataControl1.RecordSource = "Select a.Codigo as OS, a.Dt_Atend, a.Tipo_Serv, a.Valor_Serv, a.Status, b.endereco, c.equipamento From Atendimento a, Endereco b, Equipamento c Where a.cod_cliente = " & Text_Cod.Text & " and b.codigo = a.cod_end and c.codigo = a.cod_equip"
    
    'Dados do cliente, no cabeçalho
    relAtendimentosPorCliente.lblCliente.Caption = Text_NomeCli.Text
    relAtendimentosPorCliente.lblCPF.Caption = Text_Cpf.Text
    relAtendimentosPorCliente.lblContato.Caption = Text_Contrato.Text
    relAtendimentosPorCliente.lblEmail.Caption = Text_E_Mail.Text
    relAtendimentosPorCliente.lblTelComercial.Caption = strFoneCom
    relAtendimentosPorCliente.lblTelResidencial.Caption = strFoneRes
    relAtendimentosPorCliente.lblTelCelular.Caption = strFoneCel
    
    relAtendimentosPorCliente.Show (1)
    
End Sub

Private Sub btnFechar_Click()
    If ehProspect Then
        frmOrcamento.strNomeProspect = Text_NomeCli.Text
    End If
    
    Unload Me
End Sub

Private Sub btnRelatorioGerar_Click()
    If optRelClientesGeral.value = True Then
        executaRelatorio W_DBname, "Todos"
    End If
    
    If optRelClientesPorCidade.value = True Then
        frmFiltroRelatorio_Cliente.Show (1)
    End If
    
    If optRelClientesArquivoMorto Then
        executaRelatorioArquivoMorto W_DBname, "Arquivo Morto"
    End If
    
    If optnPorCliente Then
        executaRelatorioAtendimento W_DBname, "Atendimentos por cliente"
    End If
    
End Sub

Private Sub But_Close_Click()
    Unload Me
End Sub

Private Sub But_Search_Clean_Click()
    Txt_Search_Code.Text = ""
    Txt_Search_CPF.Text = ""
    Txt_Search_Name.Text = ""
    Txt_Search_Address.Text = ""
End Sub

Private Sub But_Search_Click()
    With Grid_Customer
        Call Fill_Grid(Return_Recordset(Return_SqlString(3, Txt_Search_Code.Text, Txt_Search_CPF.Text, Txt_Search_Name.Text, Txt_Search_Address.Text)))
    End With
End Sub

Private Sub Cmd_Edita_Click()

    If Cmd_Edita.Caption = "&Editar" Then
            
        Cmd_Edita.Caption = "&Cancela"
    
        Cmd_Inclui.enabled = False
        
        Call DesabilitaNavegacao
        Call Destrava_Registro
        Call Cad_Ed
        Call SetComboItem(Combo_Solicitante, Text_Solicitante.Text)
               
        Text_NomeCli.SetFocus
        
    ElseIf Cmd_Edita.Caption = "&Cancela" Then
        Cmd_Edita.Caption = "&Editar"
        
        Cmd_Inclui.enabled = True
        
        Call Grid_Customer_Click
        
        Call HabilitaNavegacao
        Call Cad_Ed_Fim
        Call Trava_Registro
    End If

End Sub

Private Sub Cmd_Exclui_Click()

    If MsgBox("Confirma a Exclusão do Cliente " & Text_NomeCli.Text & " ? ", vbYesNo, "Excluir Registros") = vbYes Then
    
        'velho: Call Return_Recordset(Return_SqlString(2, Text_Cod.Text)).Delete
        
        'novo
        chkArquivoMorto.value = 1
        excluindo = True
        
        Cmd_Salva_Click
    End If
End Sub

'Private Sub Cmd_Finaliza_Click()
'    Cod_Cli = CInt(Text_Cod.Text)
'    Nome_Cli_Atend = Text_NomeCli.Text
'
'    Form_Xsol_06.Show vbModal
'End Sub

Private Sub Cmd_Inclui_Click()
Call Cad_Ed

If Cmd_Inclui.Caption = "&Incluir" Then
   
   Cmd_Inclui.Caption = "&Cancela"
   
   Cmd_Edita.enabled = False
   
   Call DesabilitaNavegacao
   Call LimpaControles
   Call Destrava_Registro
   
   If Text_NomeCli.Visible Then
    Text_NomeCli.SetFocus
   End If
   
   If ehProspect Then
    Text_NomeCli.BackColor = vbYellow
    Text_E_Mail.BackColor = vbYellow
   End If
   
   
ElseIf Cmd_Inclui.Caption = "&Cancela" Then
   
   Cmd_Inclui.Caption = "&Incluir"
   
   Cmd_Edita.enabled = True
   
   If ehProspect Then
    Grid_Customer.row = 0
    'Grid_Customer.RowPos = 0
   End If
   
   If Not ehProspect Then
    Call Grid_Customer_Click
   End If
   
   Call HabilitaNavegacao
   Call Trava_Registro
   
   Call Cad_Ed_Fim

   If ehProspect Then
    Text_NomeCli.BackColor = vbWhite
    Text_E_Mail.BackColor = vbWhite
   End If


End If

End Sub


Sub pintaLinhas()
    
    Dim contador1 As Integer
    
    Dim contador3 As Integer
    
    'primeiro pinta as linhas de branco
    'despintaLinhasEquipamentosDaOs
    
    For contador1 = 2 To Grid_Customer.Rows - 1
        If (Grid_Customer.TextMatrix(contador1, 4) = "S") Then
          
            'MsgBox "achou"
            Grid_Customer.row = contador1

            For contador3 = 0 To 3
                Grid_Customer.col = contador3
                
                
                If Grid_Customer.TextMatrix(contador1, 5) = "EXCLUIDO MANUALMENTE" Then
                    Grid_Customer.CellBackColor = &H8080FF
                End If
                
                If (Grid_Customer.TextMatrix(contador1, 5) <> "") And (Grid_Customer.TextMatrix(contador1, 5) <> "EXCLUIDO MANUALMENTE") Then
                    Grid_Customer.CellBackColor = vbWhite
                End If
                
                If (Grid_Customer.TextMatrix(contador1, 5) <> " ") And (Grid_Customer.TextMatrix(contador1, 5) <> "EXCLUIDO MANUALMENTE") Then
                    Grid_Customer.CellBackColor = &HFFFF&
                End If
                
                '&H80FF80 = verde claro
                '&H008080FF& = VERMELHO CLARO
                '&H0000FFFF& = AMARELO
            Next contador3

            'If Form_Service.Visible Then
            '    Grid_Equipment.TextMatrix(contador1, 2) = "S"
            'End If
        End If
    Next contador1

End Sub

Private Sub Cmd_Salva_Click()
    
    Dim rs As ADODB.Recordset
    
    
    If (chkArquivoMorto.value = 1) And (excluindo = False) And (cmdJustificativa.enabled = False) Then
        
justifica:
        strJustificativaArquivoMorto = InputBox("Informe o motivo pelo qual o cliente está indo para o Arquivo Morto.", "Justificativa para inclusão em Arquivo Morto")
        If Trim(strJustificativaArquivoMorto) = "" Then
            MsgBox "Justificativa não informada!", vbOKOnly, "xSol Aquecedores"
            GoTo justifica
        End If
    End If
    
    If (chkArquivoMorto.value = 1) And (excluindo) Then
        strJustificativaArquivoMorto = "EXCLUIDO MANUALMENTE"
        excluindo = False
    End If
    
    
    If Not Fill_Mandatory_Fields Then
        Exit Sub
    End If
    
    If Cmd_Inclui.Caption = "&Cancela" Then
        Set rs = Return_Recordset(Return_SqlString(4))
        rs.AddNew
        Text_Cod.Text = rs.Fields("codigo")
    Else
        Set rs = Return_Recordset(Return_SqlString(2, Text_Cod.Text))
    End If

    Call Record_Record(rs)

    rs.Update

    rs.Close
    Set rs = Nothing
    
    Call Trava_Registro
    Call HabilitaNavegacao
    
    If Cmd_Inclui.Caption = "&Cancela" Then
       Cmd_Inclui.Caption = "&Incluir"
    ElseIf Cmd_Edita.Caption = "&Cancela" Then
       Cmd_Edita.Caption = "&Editar"
    End If
     
    Call Cad_Ed_Fim
    Cmd_Salva.enabled = False
    Cmd_Edita.enabled = True
    Cmd_Inclui.enabled = True
    Cmd_Exclui.enabled = True

    With Grid_Customer
        Call Fill_Grid(Return_Recordset(Return_SqlString(3, Txt_Search_Code.Text, Txt_Search_CPF.Text, Txt_Search_Name.Text, Txt_Search_Address.Text)), Text_Cod.Text)
    End With

End Sub

Private Function Fill_Mandatory_Fields() As Boolean
    
    Fill_Mandatory_Fields = True
    
    If Me.Caption = "Cadastro de Clientes - XSol" Then
        If Trim(Text_NomeCli.Text) = "" Then
           MsgBox "É necessário digitar o Nome do Cliente / Razão Social! ", vbOKOnly, "XSol Aquecedores"
           Text_NomeCli.SetFocus
           Fill_Mandatory_Fields = False
           Exit Function
        End If
        
        If Combo_Solicitante.Text = "" And Text_Solicitante.Text = "" Then
          MsgBox "É necessário selecionar o Solicitante! ", vbOKOnly, "XSol Aquecedores"
           Combo_Solicitante.SetFocus
           Fill_Mandatory_Fields = False
           Exit Function
        End If
    ElseIf Me.Caption = "Cadastro de Prospects - XSol" Then
        If Trim(Text_NomeCli.Text) = "" Then
           MsgBox "É necessário digitar o Nome do Prospect! ", vbOKOnly, "XSol Aquecedores"
           Text_NomeCli.SetFocus
           Fill_Mandatory_Fields = False
           Exit Function
        End If
        
        If Trim(Text_E_Mail.Text) = "" Then
           MsgBox "É necessário digitar o email do Prospect! ", vbOKOnly, "XSol Aquecedores"
           Text_E_Mail.SetFocus
           Fill_Mandatory_Fields = False
           Exit Function
        End If
              
    End If

End Function

Private Sub Cmd_Serv_Click()
    Atendente = Text_Atendente.Text
    
    If Trim(Text_Cod) = "" Then
        MsgBox "Por favor escolha um cliente", vbOKOnly, "Xsol Aquecedores"
        Exit Sub
    End If
    
    If chkArquivoMorto.value = 1 Then
        MsgBox "Não é possível realizar atendimento para cliente no Arquivo Morto!", vbOKOnly, "XSol Aquecedores"
        Exit Sub
    End If
    
    Cod_Cli = CInt(Text_Cod.Text)
    Nome_Cli_Atend = Text_NomeCli.Text
    

    If Grid_Address.RowSel = 1 Then
        MsgBox "Informe o endereço do atendimento"
        Exit Sub
    End If

    strFoneRes = Grid_Address.TextMatrix(Grid_Address.RowSel, 6)
    strFoneCom = Grid_Address.TextMatrix(Grid_Address.RowSel, 7)
    strFoneCel = Grid_Address.TextMatrix(Grid_Address.RowSel, 8)
    
    
    Form_Service.Show vbModal
End Sub

Private Sub cmdJustificativa_Click()
    MsgBox Grid_Customer.TextMatrix(Grid_Customer.RowSel, 5), vbOKOnly, "xSOL Aquecedores"
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

Private Sub Form_Activate()
    If ehProspect Then
        Text_NomeCli.SetFocus
        Frame9.enabled = False
        Cmd_Serv.enabled = False
    Else
        Frame9.enabled = True
        Cmd_Serv.enabled = True
    End If
    
    If (MDIForm_Xsol.acessibilidade_TextoNegrito = True) Or (MDIForm_Xsol.acessibilidade_TextoNegrito = False) Then
        Combo_Solicitante.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    
    
        Txt_Search_Address.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Txt_Search_Name.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Txt_Search_CPF.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Txt_Search_Code.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Txt_Next_Visit.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Cod.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Registro.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Cnpj.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Dt_Contrato.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Contrato.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Solicitante.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_NomeCli.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Dt_Emis.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Cpf.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_E_Mail.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Atendente.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        
        Grid_Address.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Grid_Equipment.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Grid_Customer.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        
        Lbl_Search_Address.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Search_Name.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Search_CPF.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Search_Code.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Dt_Visita.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Cod.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Registro.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label2.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Cnpj.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Dt_Contrato.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Contrato.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Llb_Solitante.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Nome.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Dt_Emis.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Cpf.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_E_Mail.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Atendente.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame1.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame_Cliente.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame9.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame8.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame7.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame6.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame5.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame4.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame3.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame2.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Grid_Equipment.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Grid_Address.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Grid_Customer.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        optRelClientesArquivoMorto.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        optRelClientesPorCidade.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        optRelClientesGeral.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Option1.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        optnTodos.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Opt_Sort_Name.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Opt_Sort_Code.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        chkArquivoMorto.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        btnRelatorioGerar.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        btnFechar.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Serv.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Exclui.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Edita.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Salva.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Inclui.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        But_Search_Clean.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        But_Search.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        But_Close.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        cmdJustificativa.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    End If
    
End Sub

Private Sub Form_Load()

Dim hMenu As Long
Dim lItemCount As Long

'Call Show_Customer_Equipment(Return_Recordset(Return_SqlString(8, Grid_Customer.TextMatrix(Grid_Customer.RowSel, 0))))

hMenu = GetSystemMenu(Me.hWnd, 0)
If hMenu Then
   lItemCount = GetMenuItemCount(hMenu)
   Call RemoveMenu(hMenu, lItemCount - 1, MF_REMOVE Or _
   MF_BYPOSITION)
    Call RemoveMenu(hMenu, lItemCount - 2, MF_REMOVE Or _
    MF_BYPOSITION)
   Call DrawMenuBar(Me.hWnd)
End If

'Call Ajusta_Form(Me)

Text_NomeCli.MaxLength = 80
Text_Cpf.MaxLength = 15
Text_Cnpj.MaxLength = 25
Text_Atendente.MaxLength = 20
Text_E_Mail.MaxLength = 50
Text_Solicitante.MaxLength = 50
Text_Contrato.MaxLength = 25


'removido, pois já tem o cloud_ConfigFile que faz isso
'Call Arquivos_Dados

'Set W_DB = New Connection
'Set conn = New ADODB.Connection
'conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
'W_DB.Mode = adModeReadWrite
'W_DB.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"


'Se precisar
'Set cloud_registro = conexao.Execute("select Tipo_Cliente from Empresa Order By Tipo_Cliente")

Text_Dt_Emis.enabled = False

Call Trava_Registro

Call Configure_Grid

Call Configure_Controls

Call Fill_Grid(Return_Recordset(Return_SqlString(1)))

If Not ehProspect Then
Me.Caption = "Cadastro de Clientes - Xsol"
    Me.Width = 17610
Else
    Me.Caption = "Cadastro de Prospects - Xsol"
    Me.Width = 9350
    Cmd_Inclui_Click
    Text_NomeCli.BackColor = vbYellow
    Text_E_Mail.BackColor = vbYellow
End If
 
 'cloud parou aqui
End Sub

Private Sub Configure_Controls()
    Opt_Sort_Code.value = True
End Sub

Private Sub Configure_Grid()

    With Grid_Customer
        .Cols = 6
        .FixedCols = 0
        .ColWidth(0) = 650
        .ColWidth(1) = 4000
        .ColWidth(2) = 1500
        .ColWidth(3) = 1300
        .ColWidth(4) = 900
        .ColWidth(5) = 1500
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Nome"
        .TextMatrix(0, 2) = "CPF"
        .TextMatrix(0, 3) = "Celular"
        .TextMatrix(0, 4) = "Arq. Morto"
        .TextMatrix(0, 5) = "Justificativa"
        .RowHeight(1) = 0
    End With
    
    With Grid_Address
        .Cols = 9
        .ColWidth(0) = 630
        .ColWidth(1) = 3870
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .ColWidth(4) = 3200
        .ColWidth(5) = 3200
        .ColWidth(6) = 3200
        .ColWidth(7) = 3200
        .ColWidth(8) = 3200
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignLeftCenter
        .ColAlignment(8) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Endereço"
        .TextMatrix(0, 2) = "Bairro"
        .TextMatrix(0, 3) = "Cidade"
        .TextMatrix(0, 4) = "Contato"
        .TextMatrix(0, 5) = "Tipo de telhado"
        .TextMatrix(0, 6) = "Fone Residencial"
        .TextMatrix(0, 7) = "Fone Comercial"
        .TextMatrix(0, 8) = "Fone Celular"
        .RowHeight(1) = 0
    End With
    
    With Grid_Equipment
        .Cols = 4
        .ColWidth(0) = 630
        .ColWidth(1) = 6800
        .ColWidth(2) = 800
        .ColWidth(3) = 500
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Equipamento"
        .TextMatrix(0, 2) = "Cód. End."
        .TextMatrix(0, 3) = "Ficha"
        .RowHeight(1) = 0
    End With
End Sub


' queryType = 1 => all
'           = 2 => by codigo
'           = 3 => by codigo, cpf, name or address

Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String, Optional param3 As String, Optional param4 As String) As String
    Dim whereClause As String
    
    'ajustar para mysql
    If queryType = 1 Then
        'mysql ok
        Return_SqlString = "Select codigo,nome_cli,cpf,fone_celular, Aquivo_Morto, Aquivo_Morto_Justificativa FROM Clientes c "
        
    ElseIf queryType = 2 Then
        'mysql ok
        Return_SqlString = "SELECT * FROM `Clientes` c WHERE codigo = " & param1
    ElseIf queryType = 3 Then
        
        'mysql ok
        Return_SqlString = "SELECT c.codigo,nome_cli,cpf,fone_celular,Aquivo_Morto, Aquivo_Morto_Justificativa FROM `Clientes` c LEFT JOIN `Endereco` e ON c.codigo = e.cod_cliente"
        
        If param1 <> "" Then
            whereClause = " WHERE c.codigo = " & param1
        End If
        If param2 <> "" Then
            whereClause = IIf(whereClause = "", " WHERE", " AND") & " cpf LIKE '%" & param2 & "%'"
        End If
        If param3 <> "" Then
            whereClause = IIf(whereClause = "", " WHERE", " AND") & " nome_cli LIKE '%" & param3 & "%'"
        End If
        If param4 <> "" Then
            whereClause = IIf(whereClause = "", " WHERE", " AND") & " endereco LIKE '%" & param4 & "%'"
        End If
        Return_SqlString = Return_SqlString + whereClause
    ElseIf queryType = 4 Then
        'mysql ok
        Return_SqlString = "SELECT * FROM `Clientes` c"
    ElseIf queryType = 5 Then
        'mysql ok
        Return_SqlString = "SELECT * FROM `Atendimento` WHERE cod_cliente = " & param1
    ElseIf queryType = 6 Then
        'mysql ok
        Return_SqlString = "SELECT MIN(dt_prox_visita) AS dt_prox_visita FROM `Visitas` WHERE dt_expiracao IS NULL AND cod_cliente = " & param1
    ElseIf queryType = 7 Then
        'mysql ok
        Return_SqlString = "SELECT * FROM `Endereco` WHERE cod_cliente = " & param1
    ElseIf queryType = 8 Then
        'mysql ok
        Return_SqlString = "SELECT * FROM `Equipamento` WHERE cod_cliente = " & param1
    ElseIf queryType = 9 Then
        'mysql ok
        Return_SqlString = "SELECT * FROM `Clientes` WHERE Aquivo_Morto = 'S'"
    ElseIf queryType = 10 Then
        'mysql ok
        Return_SqlString = "SELECT * FROM `Clientes` WHERE Aquivo_Morto in('S','N','')"
    ElseIf queryType = 11 Then
        'mysql ok
        Return_SqlString = "select * from `Equipamento` where cod_cliente = " & param1 & " and cod_endereco = '" & param2 & "'"

    End If

    If queryType <= 4 Then
        If Opt_Sort_Code.value Then
            Return_SqlString = Return_SqlString + " ORDER BY c.Codigo desc"
        ElseIf Opt_Sort_Name Then
            Return_SqlString = Return_SqlString + " ORDER BY Nome_Cli"
        End If
    End If
    
    'mysql Limit 50
    If (queryType = 1) Or (queryType = 3) Or (queryType = 4) Then
        Return_SqlString = Return_SqlString + " Limit 50"
    End If
    
    


End Function

Private Function Return_Recordset(strSql As String) As ADODB.Recordset
On Error GoTo reconecta
    Dim rs As New ADODB.Recordset
    
inicio:
    
    Set cloud_registro = cloud_conexao.Execute(strSql)
    
    Set Return_Recordset = cloud_registro
    Exit Function
reconecta:

    If cloud_status_conexao = 0 Then
        cloud_conexao.Close
    End If
    
    If ConectaMySQL Then
        GoTo inicio
    Else
        MsgBox "O sistema perdeu a conexão online, entre novamente!", vbCritical, "Falha na conexão com mysql online"
        Unload MDIForm_Xsol
        Exit Function
    End If
End Function

Private Sub Fill_Grid(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim lastCode As Integer
    Dim count As Integer
    
    If cloud_status_conexao = 0 Then
        Exit Sub
    End If
    
    lastCode = -1
    selectedRow = 2
    count = 0
    With Grid_Customer
        .Rows = 2
        Do Until rs.EOF
            If lastCode <> rs.Fields("codigo") Then
               lastCode = rs.Fields("codigo")
               count = count + 1
               .AddItem (rs.Fields("codigo") & vbTab & IIf(IsNull(rs.Fields("Nome_Cli")), "", rs.Fields("Nome_Cli")) & vbTab & IIf(IsNull(rs.Fields("cpf")), "", rs.Fields("cpf")) & vbTab & rs.Fields("fone_celular") & vbTab & rs.Fields("Aquivo_Morto") & vbTab & rs.Fields("Aquivo_Morto_Justificativa"))
               If code <> "" Then
                  If code = rs.Fields("codigo") Then
                     selectedRow = .Rows - 1
                  End If
               End If
            End If
            rs.MoveNext
        Loop
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            .TopRow = selectedRow
            .col = 0
            .ColSel = .Cols - 1
            Call Grid_Customer_Click
            Cmd_Edita.enabled = True
            Cmd_Exclui.enabled = True
        Else
            Cmd_Edita.enabled = False
            Cmd_Exclui.enabled = False
            Call LimpaControles
        End If
    End With
    
    Text_Registro.Text = count
    rs.Close
    Set rs = Nothing
    
    pintaLinhas
End Sub

Sub Show_Record(rs As ADODB.Recordset)

    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If

    Text_Dt_Emis.Text = IIf(IsNull(rs.Fields("Data_Emissao")), "", rs.Fields("Data_Emissao"))
    
    Text_Cod.Text = UCase(IIf(IsNull(rs.Fields("Codigo")), "", rs.Fields("Codigo")))
    
    Text_NomeCli.Text = UCase(IIf(IsNull(rs.Fields("Nome_Cli")), "", rs.Fields("Nome_Cli")))
    
    Text_Cpf.Text = IIf(IsNull(rs.Fields("CPF")), "", rs.Fields("CPF"))
    
    Text_Cnpj.Text = IIf(IsNull(rs.Fields("CNPJ")), "", rs.Fields("CNPJ"))
    
    Text_Atendente.Text = UCase(IIf(IsNull(rs.Fields("Atendente")), "", rs.Fields("Atendente")))
    
'    Text_End.Text = UCase(IIf(IsNull(rs.Fields("End")), "", rs.Fields("End")))
'
'    Text_Bairro.Text = UCase(IIf(IsNull(rs.Fields("Bairro")), "", rs.Fields("Bairro")))
'
'    Text_Cidade.Text = UCase(IIf(IsNull(rs.Fields("Cidade")), "", rs.Fields("Cidade")))
'
'    Text_Uf.Text = UCase(IIf(IsNull(rs.Fields("Uf")), "", rs.Fields("Uf")))
'
'    Text_Cep.Text = IIf(IsNull(rs.Fields("Cep")), "", rs.Fields("Cep"))
'
'    Text_Ref.Text = UCase(IIf(IsNull(rs.Fields("Ref")), "", rs.Fields("Ref")))
'
'    Text_Zona.Text = UCase(IIf(IsNull(rs.Fields("Zona")), "", rs.Fields("Zona")))
'
    
    Text_E_Mail.Text = IIf(IsNull(rs.Fields("E_Mail")), "", rs.Fields("E_Mail"))
    
'    Text_Tipo_Equip.Text = UCase(IIf(IsNull(rs.Fields("Tipo_Eq_Mod")), "", rs.Fields("Tipo_Eq_Mod")))
    
    Text_Solicitante.Text = UCase(IIf(IsNull(rs.Fields("Solicitante")), "", rs.Fields("Solicitante")))
    
    Text_Contrato.Text = IIf(IsNull(rs.Fields("Contrato")), "", rs.Fields("Contrato"))
    
    Text_Dt_Contrato.Text = IIf(IsNull(rs.Fields("Dt_Contrato")), "", rs.Fields("Dt_Contrato"))
    
    Txt_Next_Visit.Text = GET_CLIENT_NEXT_VISIT(rs.Fields("Codigo"))
    
    If rs.Fields("Aquivo_Morto") = "S" Then
        chkArquivoMorto.value = 1
        cmdJustificativa.enabled = True
    Else
        chkArquivoMorto.value = 0
        cmdJustificativa.enabled = False
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Show_Customer_Address(rs As ADODB.Recordset)
    With Grid_Address
        .Rows = 2
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If
        Do Until rs.EOF
            .AddItem (rs.Fields("codigo") & vbTab & IIf(IsNull(rs.Fields("Endereco")), "", rs.Fields("Endereco")) & vbTab & IIf(IsNull(rs.Fields("Bairro")), "", rs.Fields("Bairro")) & vbTab & IIf(IsNull(rs.Fields("Cidade")), "", rs.Fields("Cidade")) & vbTab & IIf(IsNull(rs.Fields("Contato")), "", rs.Fields("Contato")) & vbTab & rs.Fields("Tipo_Telhado") & vbTab & rs.Fields("FONERES") & vbTab & rs.Fields("FONECOM") & vbTab & rs.Fields("FONECEL"))
            rs.MoveNext
        Loop
        .Refresh
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Show_Customer_Equipment(rs As ADODB.Recordset)

    If cloud_status_conexao = 0 Then
        Exit Sub
    End If
    
    With Grid_Equipment
        .Rows = 2
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("codigo") & vbTab & IIf(IsNull(rs.Fields("Equipamento")), "", rs.Fields("Equipamento")) & vbTab & IIf(IsNull(rs.Fields("Cod_Endereco")), "", rs.Fields("Cod_Endereco")) & vbTab & IIf(IsNull(rs.Fields("ExibeNaFicha")), "", rs.Fields("ExibeNaFicha")))
            rs.MoveNext
        Loop
        .Refresh
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Private Function GET_CLIENT_NEXT_VISIT(codigo As String) As String
    Dim rs As ADODB.Recordset
    
    Set rs = Return_Recordset(Return_SqlString(6, codigo))
    If rs.RecordCount > 0 Then
        GET_CLIENT_NEXT_VISIT = IIf(IsNull(rs.Fields("dt_prox_visita")), "", rs.Fields("dt_prox_visita"))
    Else
        GET_CLIENT_NEXT_VISIT = ""
    End If
End Function

Sub Trava_Registro()

Cmd_Salva.enabled = False
'Combo_Zona.Visible = False
'Combo_Uf.Visible = False
Combo_Solicitante.Visible = False

Text_NomeCli.Locked = True
Text_Cpf.Locked = True
Text_Cnpj.Locked = True
Text_Atendente.Locked = True
'Text_End.Locked = True
'Text_Bairro.Locked = True
'Text_Cidade.Locked = True
'Text_Uf.Locked = True
'Text_Cep.Locked = True
'Text_Ref.Locked = True
'Text_Zona.Locked = True
Text_E_Mail.Locked = True
'Text_Tipo_Equip.Locked = True
Text_Solicitante.Locked = True
Text_Contrato.Locked = True
Text_Dt_Contrato.Locked = True
chkArquivoMorto.enabled = False

End Sub

Sub Destrava_Registro()

Cmd_Salva.enabled = True
Combo_Solicitante.Visible = True

Text_NomeCli.Locked = False
Text_Cpf.Locked = False
Text_Cnpj.Locked = False
Text_Atendente.Locked = False
Text_E_Mail.Locked = False
Text_Solicitante.Locked = False
Text_Contrato.Locked = False
Text_Dt_Contrato.Locked = False
chkArquivoMorto.enabled = True

End Sub

Sub LimpaControles()

Text_Dt_Emis.Text = ""
Text_Cod.Text = ""
Text_NomeCli.Text = ""
Text_Cpf.Text = ""
Text_Cnpj.Text = ""
Text_Atendente.Text = ""
Text_E_Mail.Text = ""
Text_Contrato.Text = ""
Text_Solicitante.Text = ""
Text_Dt_Contrato.Text = ""

Grid_Address.Clear
Grid_Equipment.Clear

End Sub

Private Sub Record_Record(rs As ADODB.Recordset)
  
   If Cmd_Inclui.Caption = "&Cancela" Then
      rs.Fields("Data_Emissao").value = Date
   End If
   
   If Trim(Text_NomeCli.Text) = "" Then
      rs.Fields("Nome_Cli").value = Null
   Else
      rs.Fields("Nome_Cli").value = Text_NomeCli.Text
   End If
   
   If Trim(Text_Cpf.Text) = "" Then
      rs.Fields("CPF").value = Null
   Else
      rs.Fields("CPF").value = Text_Cpf.Text
   End If
   
   If Trim(Text_Cnpj.Text) = "" Then
      rs.Fields("CNPJ").value = Null
   Else
      rs.Fields("CNPJ").value = Text_Cnpj.Text
   End If
   
   If Trim(Text_Atendente.Text) = "" Then
      rs.Fields("Atendente").value = Null
   Else
      rs.Fields("Atendente").value = Text_Atendente.Text
   End If
   
   
   If Trim(Text_E_Mail.Text) = "" Then
      rs.Fields("E_Mail").value = Null
   Else
      rs.Fields("E_Mail").value = Text_E_Mail.Text
   End If
   
   If Trim(Text_Solicitante.Text) = "" Then
      rs.Fields("Solicitante").value = Null
   Else
      rs.Fields("Solicitante").value = Text_Solicitante.Text
   End If
   
   If Trim(Text_Contrato.Text) = "" Then
      rs.Fields("Contrato").value = Null
   Else
      rs.Fields("Contrato").value = Text_Contrato.Text
   End If
   
   If Trim(Text_Dt_Contrato.Text) = "" Then
      rs.Fields("Dt_Contrato").value = Null
   Else
      rs.Fields("Dt_Contrato").value = Text_Dt_Contrato.Text
   End If
   
   If chkArquivoMorto.value = 0 Then
    rs.Fields("Aquivo_Morto").value = "N"
    rs.Fields("Aquivo_Morto_Justificativa").value = ""
    cmdJustificativa.enabled = False
   Else
    rs.Fields("Aquivo_Morto").value = "S"
    rs.Fields("Aquivo_Morto_Justificativa").value = strJustificativaArquivoMorto
    cmdJustificativa.enabled = True
   End If
   
   If Not ehProspect Then
    rs.Fields("Tipo").value = "CLIENTE"
   Else
    rs.Fields("Tipo").value = "PROSPECT"
   End If
   
End Sub

Sub DesabilitaNavegacao()

   But_Close.enabled = False
   Cmd_Exclui.enabled = False
   
   But_Search_Clean.enabled = False
   But_Search.enabled = False
   Txt_Search_Code.enabled = False
   Txt_Search_CPF.enabled = False
   Txt_Search_Name.enabled = False
   Opt_Sort_Code.enabled = False
   Opt_Sort_Name.enabled = False
   Grid_Address.enabled = False
   'Grid_Equipment.enabled = False
   Grid_Customer.enabled = False
End Sub

Sub HabilitaNavegacao()

   But_Close.enabled = True
   Cmd_Exclui.enabled = True
   
   But_Search_Clean.enabled = True
   But_Search.enabled = True
   Txt_Search_Code.enabled = True
   Txt_Search_CPF.enabled = True
   Txt_Search_Name.enabled = True
   Opt_Sort_Code.enabled = True
   Opt_Sort_Name.enabled = True
   Grid_Address.enabled = True
   Grid_Equipment.enabled = True
   Grid_Customer.enabled = True
   
End Sub

Private Sub Form_Terminate()

W_DB.Close

End Sub


Sub Equipamentos_Show_Record_Detail(rs As ADODB.Recordset)

    With Grid_Equipment
        .Rows = 2
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If
        Do Until rs.EOF
            .AddItem (rs.Fields("Codigo") & vbTab & IIf(IsNull(rs.Fields("Equipamento")), "", rs.Fields("Equipamento")) & vbTab & IIf(IsNull(rs.Fields("Cod_Endereco")), "", rs.Fields("Cod_Endereco")) & vbTab & IIf(IsNull(rs.Fields("ExibeNaFicha")), "", rs.Fields("ExibeNaFicha")))
            rs.MoveNext
        Loop
        .Refresh
    End With
    
    rs.Close
    Set rs = Nothing
    
    'rs.Close
    'Set rs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
    'If (Grid_Address.Rows <= 2) And (cloud_status_conexao <> 0) Then
    '    Cancel = True
    '    MsgBox "Cadastre ao menos 01 endereço"
    'End If
End Sub

Private Sub Grid_Address_Click()

    Call Equipamentos_Show_Record_Detail(Return_Recordset(Return_SqlString(11, Grid_Customer.TextMatrix(Grid_Customer.RowSel, 0), Grid_Address.TextMatrix(Grid_Address.RowSel, 0))))

End Sub

Private Sub Grid_Address_DblClick()
    Cod_Cli = CInt(Text_Cod.Text)
    Nome_Cli_Atend = Text_NomeCli.Text
    Cod_Address = Grid_Address.TextMatrix(Grid_Address.RowSel, 0)
    Form_Address.Show vbModal
    Call Show_Customer_Address(Return_Recordset(Return_SqlString(7, Str(Cod_Cli))))
End Sub

Private Sub Grid_Equipment_DblClick()
    Cod_Cli = CInt(Text_Cod.Text)
    Nome_Cli_Atend = Text_NomeCli.Text
    Cod_Equipment = Grid_Equipment.TextMatrix(Grid_Equipment.RowSel, 0)
    Form_Equipment.Show

    
End Sub

Private Sub Grid_Customer_Click()
    With Grid_Customer
        If .RowSel = 1 Then
            Exit Sub
        End If
        
        Call Show_Record(Return_Recordset(Return_SqlString(2, .TextMatrix(.RowSel, 0))))
        Call Show_Customer_Address(Return_Recordset(Return_SqlString(7, .TextMatrix(.RowSel, 0))))
        Call Show_Customer_Equipment(Return_Recordset(Return_SqlString(8, .TextMatrix(.RowSel, 0))))
        codCliente = .TextMatrix(.RowSel, 0)
        
    End With
End Sub

Public Sub Refresh_Grid_Equipment()
    Dim rs As ADODB.Recordset
    
    Call ConnectionRefresh
    With Grid_Customer
        Call Show_Customer_Equipment(Return_Recordset(Return_SqlString(8, .TextMatrix(.RowSel, 0))))
    End With
End Sub

Public Sub Refresh_Grid_Address()
    Dim rs As ADODB.Recordset
    
    Call ConnectionRefresh
    With Grid_Customer
        Call Show_Customer_Address(Return_Recordset(Return_SqlString(7, .TextMatrix(.RowSel, 0))))
    End With
End Sub

Public Sub Refresh_Record_Shown()
    Call Grid_Customer_Click
End Sub

Private Sub Grid_Customer_SelChange()
    Call Grid_Customer_Click
End Sub

Private Sub Opt_Sort_Code_Click()
    With Grid_Customer
        Call Fill_Grid(Return_Recordset(Return_SqlString(3, Txt_Search_Code.Text, Txt_Search_CPF.Text, Txt_Search_Name.Text, Txt_Search_Address.Text)))
    End With
End Sub

Private Sub Opt_Sort_Name_Click()
    With Grid_Customer
        Call Fill_Grid(Return_Recordset(Return_SqlString(3, Txt_Search_Code.Text, Txt_Search_CPF.Text, Txt_Search_Name.Text, Txt_Search_Address.Text)))
    End With
End Sub

Private Sub Text_Bairro_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If

End Sub

Private Sub Option1_Click()
    With Grid_Customer
        Call Fill_Grid(Return_Recordset(Return_SqlString(9, 9, Txt_Search_CPF.Text, Txt_Search_Name.Text, Txt_Search_Address.Text)))
    End With
End Sub

Private Sub optnTodos_Click()
    With Grid_Customer
        Call Fill_Grid(Return_Recordset(Return_SqlString(10, 10, Txt_Search_CPF.Text, Txt_Search_Name.Text, Txt_Search_Address.Text)))
    End With
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

Private Sub Text_Cnpj_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If CGCValido(Text_Cnpj.Text) = False Then
      MsgBox "CNPJ Invalido!!!", vbOKOnly, "Xsol Aquecedores"
      Text_Cnpj.Text = ""
   Else
      SendKeys "{TAB}"
   End If
End If

End Sub

Private Sub Text_Cpf_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If CPFValido(Text_Cpf.Text) = False Then
      MsgBox "CPF Invalido!!!", vbOKOnly, "Xsol Aquecedores"
      Text_Cpf.Text = ""
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

    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 180 Then
        KeyAscii = 0
    End If

End Sub

Private Sub Text_Uf_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If

End Sub

Sub Cad_Ed()

If Me.Caption = "Cadastro de Clientes - XSol" Then
    Call Lista_Empresa
    
    Text_NomeCli.TabIndex = 1
    Text_NomeCli.BackColor = vbYellow
    
    Combo_Solicitante.TabIndex = 6
    Combo_Solicitante.BackColor = vbYellow
    
    Text_Contrato.TabIndex = 7
    Text_Dt_Contrato.TabIndex = 15
    
    Cmd_Salva.TabIndex = 8
ElseIf Me.Caption = "Cadastro de Prospects - XSol" Then
    Text_NomeCli.BackColor = vbYellow
    Text_E_Mail.BackColor = vbYellow
End If

End Sub
Sub Cad_Ed_Fim()

   Text_NomeCli.BackColor = vbWhite
   Combo_Solicitante.BackColor = vbWhite
   
   If ehProspect Then
    Text_NomeCli.BackColor = vbYellow
    Text_E_Mail.BackColor = vbYellow
    
  End If
   

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

Private Sub Txt_Search_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And KeyAscii <> 13 And Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
        'Txt_Search_Code.Text = ""
        Exit Sub
    End If
    If KeyAscii = 13 Then
        With Grid_Customer
            Call Fill_Grid(Return_Recordset(Return_SqlString(3, Txt_Search_Code.Text, Txt_Search_CPF.Text, Txt_Search_Name.Text, Txt_Search_Address.Text)))
        End With
    End If
End Sub

Private Sub Txt_Search_CPF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With Grid_Customer
            Call Fill_Grid(Return_Recordset(Return_SqlString(3, Txt_Search_Code.Text, Txt_Search_CPF.Text, Txt_Search_Name.Text, Txt_Search_Address.Text)))
        End With
    End If
End Sub

Private Sub Txt_Search_Name_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With Grid_Customer
            Call Fill_Grid(Return_Recordset(Return_SqlString(3, Txt_Search_Code.Text, Txt_Search_CPF.Text, Txt_Search_Name.Text, Txt_Search_Address.Text)))
        End With
    End If
End Sub

Private Sub Txt_Search_Address_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With Grid_Customer
            Call Fill_Grid(Return_Recordset(Return_SqlString(3, Txt_Search_Code.Text, Txt_Search_CPF.Text, Txt_Search_Name.Text, Txt_Search_Address.Text)))
        End With
    End If
End Sub
