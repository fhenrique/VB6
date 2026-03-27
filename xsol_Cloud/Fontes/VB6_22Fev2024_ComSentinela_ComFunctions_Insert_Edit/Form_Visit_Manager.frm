VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Visit_Manager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Próximas Visitas"
   ClientHeight    =   10200
   ClientLeft      =   1815
   ClientTop       =   555
   ClientWidth     =   17670
   Icon            =   "Form_Visit_Manager.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   17670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame8 
      Caption         =   "Relatório de atendimentos"
      Height          =   975
      Left            =   14130
      TabIndex        =   44
      Top             =   9120
      Width           =   3465
      Begin VB.CommandButton btnGerarRelatorio 
         Caption         =   "Gerar"
         Height          =   600
         Left            =   1785
         TabIndex        =   50
         Top             =   255
         Width           =   645
      End
      Begin MSComCtl2.DTPicker dtpDataDe 
         Height          =   300
         Left            =   435
         TabIndex        =   47
         Top             =   255
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   158334977
         CurrentDate     =   44239
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
         Left            =   2745
         Picture         =   "Form_Visit_Manager.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   525
      End
      Begin MSComCtl2.DTPicker dtpDataAte 
         Height          =   300
         Left            =   435
         TabIndex        =   49
         Top             =   600
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   158334977
         CurrentDate     =   44239
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Até"
         Height          =   195
         Left            =   105
         TabIndex        =   48
         Top             =   675
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "De"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   210
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Pesquisa do cliente"
      Height          =   1575
      Left            =   10410
      TabIndex        =   23
      Top             =   7560
      Width           =   7215
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
         TabIndex        =   28
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Txt_Search_CPF 
         Height          =   285
         Left            =   960
         TabIndex        =   27
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Txt_Search_Name 
         Height          =   285
         Left            =   960
         TabIndex        =   26
         Top             =   1080
         Width           =   4920
      End
      Begin VB.CommandButton But_Search 
         Height          =   435
         Left            =   5970
         Picture         =   "Form_Visit_Manager.frx":03DA
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Pesquisar"
         Top             =   705
         Width           =   435
      End
      Begin VB.CommandButton But_Search_Clean 
         Height          =   420
         Left            =   5970
         Picture         =   "Form_Visit_Manager.frx":0D1C
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Limpar"
         Top             =   210
         Width           =   435
      End
      Begin VB.Label Lbl_Search_Code 
         Alignment       =   1  'Right Justify
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Lbl_Search_CPF 
         Alignment       =   1  'Right Justify
         Caption         =   "CPF:"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Lbl_Search_Name 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Controles"
      Height          =   975
      Left            =   10395
      TabIndex        =   16
      Top             =   9120
      Width           =   3645
      Begin VB.CommandButton But_Alterar 
         Caption         =   "&Editar"
         Height          =   495
         Left            =   990
         TabIndex        =   20
         Top             =   285
         Width           =   720
      End
      Begin VB.CommandButton But_Salvar 
         Caption         =   "&Salvar"
         Height          =   495
         Left            =   1860
         TabIndex        =   19
         Top             =   285
         Width           =   720
      End
      Begin VB.CommandButton But_Incluir 
         BackColor       =   &H0080FF80&
         Caption         =   "&Incluir"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   285
         Width           =   750
      End
      Begin VB.CommandButton But_Excluir 
         BackColor       =   &H008080FF&
         Caption         =   "E&xcluir"
         Height          =   495
         Left            =   2730
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   285
         Width           =   780
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Histórico de Modificações (duplo clique p/ ver endereço do atendimento)"
      Height          =   4695
      Left            =   10410
      TabIndex        =   2
      Top             =   2835
      Width           =   7215
      Begin VB.Frame panEnderecos 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Endereço do atendimento:"
         Height          =   2040
         Left            =   15
         TabIndex        =   51
         Top             =   1365
         Visible         =   0   'False
         Width           =   7185
         Begin MSFlexGridLib.MSFlexGrid dgEnderecos 
            Height          =   1395
            Left            =   45
            TabIndex        =   53
            Top             =   225
            Width           =   7080
            _ExtentX        =   12488
            _ExtentY        =   2461
            _Version        =   393216
         End
         Begin VB.CommandButton btnFecharPanEndereco 
            Caption         =   "Fechar"
            Height          =   270
            Left            =   6495
            TabIndex        =   52
            Top             =   1695
            Width           =   630
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Hist 
         Height          =   4335
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   7646
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Próximas Visitas"
      Height          =   7215
      Left            =   105
      TabIndex        =   1
      Top             =   2835
      Width           =   10215
      Begin MSFlexGridLib.MSFlexGrid Grid_Prox_Visita 
         Height          =   6855
         Left            =   90
         TabIndex        =   14
         Top             =   225
         Width           =   10020
         _ExtentX        =   17674
         _ExtentY        =   12091
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados"
      Height          =   2775
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   17505
      Begin VB.Frame Frame7 
         Caption         =   "Atendimentos neste dia"
         Height          =   750
         Left            =   5235
         TabIndex        =   42
         Top             =   1710
         Width           =   2070
         Begin VB.TextBox txtQtdAtendimentosPorDia 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   750
            MaxLength       =   3
            TabIndex        =   43
            Top             =   330
            Width           =   510
         End
      End
      Begin VB.TextBox txtCelular 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4665
         TabIndex        =   40
         Top             =   1080
         Width           =   2220
      End
      Begin VB.TextBox txtTelefone 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   39
         Top             =   1080
         Width           =   2040
      End
      Begin MSComCtl2.DTPicker Txt_Dt_Criacao 
         Height          =   330
         Left            =   1920
         TabIndex        =   37
         Top             =   1785
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   156827649
         CurrentDate     =   44033
      End
      Begin MSComCtl2.DTPicker Mask_Dt_Prox_Visita 
         Height          =   300
         Left            =   1920
         TabIndex        =   36
         Top             =   1425
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   156827649
         CurrentDate     =   44033
      End
      Begin VB.Frame Frame6 
         Caption         =   "Limite de atendimento diário"
         Height          =   750
         Left            =   7740
         TabIndex        =   33
         Top             =   1710
         Width           =   2235
         Begin VB.CommandButton btnLigacoesDiarias 
            Caption         =   "&Salvar"
            Enabled         =   0   'False
            Height          =   285
            Left            =   1140
            TabIndex        =   35
            Top             =   330
            Width           =   600
         End
         Begin VB.TextBox txtQtdLigacoesDiarias 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   480
            MaxLength       =   3
            TabIndex        =   34
            Top             =   330
            Width           =   510
         End
      End
      Begin VB.TextBox Txt_OS 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3660
         TabIndex        =   22
         Top             =   360
         Width           =   1230
      End
      Begin VB.TextBox Txt_Obs 
         BackColor       =   &H8000000F&
         Height          =   2175
         Left            =   10290
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   360
         Width           =   7095
      End
      Begin VB.TextBox Txt_Dt_Expiracao 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   11
         Top             =   2160
         Width           =   2070
      End
      Begin VB.TextBox Txt_Nome_Cliente 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3660
         TabIndex        =   7
         Top             =   720
         Width           =   3225
      End
      Begin VB.TextBox Txt_Cod_Cliente 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   720
         Width           =   1320
      End
      Begin VB.TextBox Txt_Cod_Visita 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   1320
      End
      Begin MSMask.MaskEdBox Mask_Hora_Prox_Visita 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   3315
         TabIndex        =   32
         Top             =   1440
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   -2147483633
         Enabled         =   0   'False
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "/"
         Height          =   255
         Left            =   4185
         TabIndex        =   41
         Top             =   1155
         Width           =   135
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Telefone / Celular:"
         Height          =   255
         Left            =   450
         TabIndex        =   38
         Top             =   1095
         Width           =   1365
      End
      Begin VB.Label Lbl_OS 
         Alignment       =   1  'Right Justify
         Caption         =   "OS:"
         Height          =   255
         Left            =   3255
         TabIndex        =   21
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Observação.:"
         Height          =   255
         Left            =   9210
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Lbl_Dt_Expiracao 
         Alignment       =   1  'Right Justify
         Caption         =   "Data de Expiração:"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Lbl_Dt_Prox_Visita 
         Alignment       =   1  'Right Justify
         Caption         =   "Data da Próxima Visita:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Lbl_Dt_Criacao 
         Alignment       =   1  'Right Justify
         Caption         =   "Data de Criação:"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Lbl_Cliente 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Lbl_Cod_Visita 
         Alignment       =   1  'Right Justify
         Caption         =   "Código da Visita:"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form_Visit_Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public W_DB As ADODB.Connection
Public Configuracao_DBname As String
Public enderecoAtendimento As String

Private Sub Record_Record(rs As ADODB.Recordset)
    If rs.RecordCount = 0 Then
        rs.AddNew
    End If
    rs.Fields("LimiteLigacoesDiarias").value = txtQtdLigacoesDiarias.Text
End Sub

Private Sub btnFecharPanEndereco_Click()
    panEnderecos.Visible = False
End Sub

Private Sub btnGerarRelatorio_Click()

    If dtpDataAte.value < dtpDataDe.value Then
        MsgBox "A data final deve ser maior que a inicial."
        Exit Sub
    End If
    
    'Format(Date, "MM/dd/yyyy") '04/29/2009
    relProximasVisitas.lblPeriodo.Caption = "Período: " & Format(dtpDataDe.value, "dd/MM/yyyy") & " a " & Format(dtpDataAte.value, "dd/MM/yyyy")
    relProximasVisitas.DAODataControl1.DatabaseName = W_DBname
    relProximasVisitas.DAODataControl1.RecordSource = "SELECT v.Codigo as Codigo, c.codigo as CliCodigo, nome_cli as CliNome, Dt_Prox_Visita, Fone_celular, Fone_res FROM Clientes c LEFT JOIN (SELECT * FROM Visitas WHERE dt_Expiracao IS NULL) v ON c.codigo = v.cod_cliente WHERE v.Dt_Expiracao IS NULL AND c.Aquivo_Morto = 'N' AND Dt_Prox_Visita >= cDate('" & dtpDataDe.value & "') AND Dt_Prox_Visita <= cDate('" & dtpDataAte.value & "') ORDER BY Dt_Prox_Visita, Hora_Prox_Visita, Nome_Cli"
    relProximasVisitas.Show (1)

End Sub

Private Sub btnLigacoesDiarias_Click()
    
    If Not IsNumeric(txtQtdLigacoesDiarias.Text) Then
        MsgBox "Informe um número inteiro positivo para o limite de ligações diárias.", vbOKOnly, "XSol Aquecedores"
        txtQtdLigacoesDiarias.SetFocus
        Exit Sub
    End If
    
    Dim rsLiga As ADODB.Recordset
    Set rsLiga = Return_Recordset(Return_SqlString(5)) 'Este recordset está gravando na tabela Configuracao, campo LimiteLigacoesDiarias
    
    Call Record_Record(rsLiga)
    rsLiga.Update
    
    rsLiga.Close
    Set rsLiga = Nothing
       
    MsgBox "Quantidade de ligações gravada!", vbOKOnly, "Xsol Aquecedores"
    But_Alterar_Click
    
    
    
End Sub

Private Sub But_Alterar_Click()
    Txt_Obs.Locked = But_Alterar.Caption <> "&Editar"
    Txt_Obs.BackColor = IIf(Not Txt_Obs.Locked, RGB(255, 255, 255), COLOR_BTNFACE)
    
    txtQtdLigacoesDiarias.Locked = But_Alterar.Caption <> "&Editar"
    txtQtdLigacoesDiarias.BackColor = IIf(Not Txt_Obs.Locked, RGB(255, 255, 255), COLOR_BTNFACE)
    
    btnLigacoesDiarias.enabled = But_Alterar.Caption = "&Editar"
    
    But_Alterar.Caption = IIf(Not Txt_Obs.Locked, "&Cancelar", "&Editar")
    If Not Txt_Obs.Locked Then
        txtQtdLigacoesDiarias.enabled = True
        Mask_Hora_Prox_Visita.enabled = True
        Call Clean_TextBox
    Else
        Call Show_Record_Detail(Return_Recordset2(Return_SqlString(5)))
        txtQtdLigacoesDiarias.enabled = False
        If Trim(Txt_Cod_Cliente.Text) <> "" Then
            Call Grid_Hist_Click
        End If
    End If
    Call Configure_Controls(IIf(Not Txt_Obs.Locked, 2, 0))
End Sub

Private Sub But_Close_Click()
    Unload Me
End Sub

Private Sub But_Excluir_Click()
    If MsgBox("Confirma a Exclusão da Visita ao Cliente " & Txt_Nome_Cliente.Text & " ? ", vbYesNo, "Excluir Visita") = vbYes Then
        Call Expire_Visit(, Txt_Cod_Visita.Text)
        Call Fill_Grid(Return_Recordset(Return_SqlString(1)), Txt_Cod_Cliente.Text)
    End If
End Sub

Private Sub But_Incluir_Click()
    Mask_Dt_Prox_Visita.enabled = But_Incluir.Caption = "&Incluir"
    Mask_Hora_Prox_Visita.enabled = But_Incluir.Caption = "&Incluir"
    
    Txt_Obs.Locked = Not Mask_Dt_Prox_Visita.enabled
'    Mask_Dt_Prox_Visita.BackColor = IIf(Not Txt_Obs.Locked, RGB(255, 255, 255), COLOR_BTNFACE) 'GetSysColor(COLOR_BTNFACE))
    Mask_Hora_Prox_Visita.BackColor = IIf(Not Txt_Obs.Locked, RGB(255, 255, 255), COLOR_BTNFACE) 'GetSysColor(COLOR_BTNFACE))
    
    'Txt_Obs.BackColor = Mask_Dt_Prox_Visita.BackColor
    But_Incluir.Caption = IIf(Not Txt_Obs.Locked, "&Cancelar", "&Incluir")
    If Not Txt_Obs.Locked Then
        Call Clean_TextBox
    Else
        Call Grid_Hist_Click
    End If
    Call Configure_Controls(IIf(Not Txt_Obs.Locked, 1, 0))
    
    If Mask_Dt_Prox_Visita.enabled Then
        Mask_Dt_Prox_Visita.SetFocus
        Mask_Dt_Prox_Visita.value = Format(DateTime.Now, "Short Date")
    End If
    
End Sub

Private Sub Clean_TextBox()
    If Mask_Dt_Prox_Visita.enabled Then
        Txt_Cod_Visita.Text = ""
        'Mask_Dt_Prox_Visita.Text = "__/__/____"
        'Txt_Dt_Criacao.Text = ""
        Txt_Dt_Expiracao.Text = ""
        Txt_Obs.Text = ""
    End If
End Sub

Private Sub But_Salvar_Click()
    
    Dim rsQtd As ADODB.Recordset
    Dim qtd As Integer
    Set rsQtd = Return_Recordset(Return_SqlString(6, Mask_Dt_Prox_Visita.value))
    qtd = rsQtd.Fields("qtd")
    rsQtd.Close
    Set rsQtd = Nothing
    If qtd >= Int(txtQtdLigacoesDiarias.Text) Then
        MsgBox "O limite de atendimentos diário (" & txtQtdLigacoesDiarias.Text & ") foi atingido.", vbOKOnly, "Xsol Aquecedores"
        Exit Sub
    End If
    'Fim verificação de quantidade de atendimentos diários
    
    'MsgBox "validar aqui se foi (ou foram ) informados os equipamentos que serão atendidos."
    
    If Weekday(Mask_Dt_Prox_Visita.value, vbSunday) = vbSunday Then
        MsgBox "A data informada é um domingo, informe outro dia.", vbOKOnly, "Xsol Aquecedores"
        If Mask_Dt_Prox_Visita.enabled Then
            Mask_Dt_Prox_Visita.SetFocus
        End If
        Exit Sub
    End If
    
    Dim Insert As Boolean
    
    Insert = False
    If Mask_Dt_Prox_Visita.enabled Then
        Insert = True
        If Not Validate_Fields() Then
            Exit Sub
        End If
        Mask_Dt_Prox_Visita.enabled = False
'        Mask_Dt_Prox_Visita.BackColor = COLOR_BTNFACE
    End If
    Txt_Obs.Locked = True
    Txt_Obs.BackColor = COLOR_BTNFACE
    But_Incluir.Caption = "&Incluir"
    But_Alterar.Caption = "&Editar"
    If Insert Then
        Call Insert_New_Visit(Txt_Cod_Cliente.Text)
    Else
        Call Update_Visit(Txt_Cod_Visita.Text)
    End If
    
    Call Fill_Grid(Return_Recordset(Return_SqlString(1)), Txt_Cod_Cliente.Text)
    Call Configure_Controls(IIf(Not Txt_Obs.Locked, 1, 0))
    
    Mask_Hora_Prox_Visita.Text = "__:__"
    
End Sub

Private Function retornaProximoCodigoVisita() As Long

    Dim intProximoCodigoVisita As Long
    Dim rsProximoCodigoVisita As Recordset
    
    Set rsProximoCodigoVisita = Return_Recordset(Return_SqlString(10))
    intProximoCodigoVisita = rsProximoCodigoVisita.Fields("proximo")
    
    retornaProximoCodigoVisita = intProximoCodigoVisita
End Function

Private Function consultaEnderecoCliente(strCliente As String) As Recordset
    Set consultaEnderecoCliente = Return_Recordset(Return_SqlString(9, strCliente))
End Function

Private Sub ConfigureGridEnderecos()
    With frmEnderecosCliente.gridEnderecos
        .Cols = 1
        .FixedCols = 0
        .ColWidth(0) = 7000
        
        .ColAlignment(0) = flexAlignLeftTop
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Endereços"
        .RowHeight(1) = 0
        
    End With

End Sub

Private Sub Insert_New_Visit(clientCode As String)
    Dim rs As ADODB.Recordset
    Dim rs_enderecos As Recordset
    Dim strEnderecoAtendimento As String
    Dim intProximoCodVisita As Long
    
    Set rs_enderecos = consultaEnderecoCliente(clientCode)
    
    If rs_enderecos.RecordCount = 1 Then
        strEnderecoAtendimento = rs_enderecos.Fields("endereco")
    Else
        'hoje aqui 2, mostra tela com endereços do cliente para selecionar
        
        Call ConfigureGridEnderecos
        
        With frmEnderecosCliente.gridEnderecos
            .Rows = 1
            Do Until rs_enderecos.EOF
                .AddItem (rs_enderecos.Fields("endereco"))
                rs_enderecos.MoveNext
            Loop
        End With
        
        frmEnderecosCliente.Show (1)
        strEnderecoAtendimento = frmEnderecosCliente.strEnderecoAtendimento
        
    End If
    
    
    Set rs = Return_Recordset(Return_SqlString(2, clientCode))
    Call Expire_Visit(clientCode)
    rs.AddNew
    
    rs.Fields("Cod_Cliente") = clientCode
    rs.Fields("Dt_Criacao") = Format(Now, "Short Date")
    rs.Fields("Dt_Prox_Visita") = Mask_Dt_Prox_Visita.value
    
    If Mask_Hora_Prox_Visita.Text <> "__:__" Then
        rs.Fields("Hora_Prox_Visita") = Mask_Hora_Prox_Visita.Text
    End If
    
    rs.Fields("Obs") = Txt_Obs.Text
    rs.Fields("Endereco_Da_Visita") = strEnderecoAtendimento
    
    'pega aqui o último codigo da visitae soma + 1
    intProximoCodVisita = retornaProximoCodigoVisita
    
    rs.Fields("cod_visita") = intProximoCodVisita
    
    rs.Update
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Update_Visit(code As String)
    Dim rs As ADODB.Recordset
    
    Set rs = Return_Recordset(Return_SqlString(3, code))
    rs.Fields("Obs") = Txt_Obs.Text
    
    rs.Update
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Expire_Visit(Optional clientCode As String, Optional code As String)
    If clientCode <> "" Then
        W_DB.Execute "UPDATE Visitas SET Dt_Expiracao = Now() WHERE Dt_Expiracao IS NULL AND cod_cliente = " + clientCode
    Else
        W_DB.Execute "UPDATE Visitas SET Dt_Expiracao = Now() WHERE codigo = " + code
    End If
End Sub

Private Function Validate_Fields() As Boolean
    Dim dias As Integer
    
    Validate_Fields = False
    
    dias = DateDiff("d", Date, vData(Mask_Dt_Prox_Visita))
    If dias < 0 Then
        MsgBox "A data deve ser maior ou igual a hoje!", vbOKOnly, "XSol Aquecedores"
        Mask_Dt_Prox_Visita.SetFocus
        Exit Function
    End If
    
    Validate_Fields = True
End Function

Private Sub But_Search_Clean_Click()
    Txt_Search_Code.Text = ""
    Txt_Search_CPF.Text = ""
    Txt_Search_Name.Text = ""
End Sub

Private Sub But_Search_Click()
    With Grid_Prox_Visita
        Call Fill_Grid(Return_Recordset(Return_SqlString(4, Txt_Search_Code.Text, Txt_Search_CPF.Text, Txt_Search_Name.Text)))
    End With
End Sub

Private Sub Form_Activate()
    If (MDIForm_Xsol.acessibilidade_TextoNegrito = True) Or (MDIForm_Xsol.acessibilidade_TextoNegrito = False) Then
Txt_Search_Code.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Txt_Search_CPF.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Txt_Search_Name.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtQtdAtendimentosPorDia.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtCelular.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtTelefone.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Txt_Dt_Criacao.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtQtdLigacoesDiarias.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Txt_OS.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Txt_Obs.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Txt_Dt_Expiracao.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Txt_Nome_Cliente.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Txt_Cod_Cliente.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Txt_Cod_Visita.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
dgEnderecos.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Grid_Hist.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Grid_Prox_Visita.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
dtpDataDe.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
dtpDataAte.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Mask_Dt_Prox_Visita.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    End If

End Sub

Private Sub Grid_Hist_DblClick()
    If Grid_Hist.Rows < 3 Then
        Exit Sub
    End If
    
    If (panEnderecos.Visible = False) And (enderecoAtendimento = "") Then
        panEnderecos.Visible = True
    Else
            MsgBox enderecoAtendimento
    End If
    
    
End Sub

Private Sub Mask_Dt_Prox_Visita_Change()
        '15/09
        If Weekday(Mask_Dt_Prox_Visita.value, vbSunday) = vbSunday Then
            MsgBox "O dia selecionado é um domingo!", vbOKOnly, "xSol Aquecedores"
        End If
        carregaQtdAtendimentosNoDia
End Sub

Private Sub Txt_Search_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call But_Search_Click
    End If
End Sub

Private Sub Txt_Search_CPF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call But_Search_Click
    End If
End Sub

Private Sub Txt_Search_Name_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call But_Search_Click
    End If
End Sub

Sub Show_Record_Detail(rs As ADODB.Recordset)

    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    txtQtdLigacoesDiarias.Text = rs.Fields("LimiteLigacoesDiarias")
        
    rs.Close
    Set rs = Nothing
End Sub

Sub Show_Record_DetailDiario(rs As ADODB.Recordset)

    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    txtQtdAtendimentosPorDia.Text = rs.Fields("qtd")
        
    rs.Close
    Set rs = Nothing
End Sub

Sub ConfiguracaoArquivos_Dados()
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
    Configuracao_DBname = Path + FileName
End Sub

Public Function Return_Recordset2(strSql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    Call ConfiguracaoArquivos_Dados
    
    Set conn = Nothing
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & Configuracao_DBname & ";"
    
    
    rs.Open strSql, conn, adOpenStatic, adLockOptimistic
    
    Set Return_Recordset2 = rs
End Function

Public Function Return_Recordset3(strSql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    Call ConfiguracaoArquivos_Dados
    
    Set conn = Nothing
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & Configuracao_DBname & ";"
    
    
    rs.Open strSql, conn, adOpenStatic, adLockOptimistic
    
    Set Return_Recordset3 = rs
End Function


Private Sub carregaQtdAtendimentosNoDia()
    Call Show_Record_DetailDiario(Return_Recordset3(Return_SqlString(7, Mask_Dt_Prox_Visita.value)))
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
        
    Set W_DB = New ADODB.Connection
    W_DB.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"

    Call Configure_Grid
    Call Fill_Grid(Return_Recordset(Return_SqlString(1)))


    'Trazendo o limite de atendimentos diários
    Call Show_Record_Detail(Return_Recordset2(Return_SqlString(5)))
    
    'Trazendo q quantidade de atedimentos no dia
    Call carregaQtdAtendimentosNoDia
    
    dtpDataDe.value = DateTime.Now
    dtpDataAte.value = DateTime.Now + 1
    
    

End Sub

Private Sub Configure_Controls(Optional status As Integer)
    Dim enabled As Boolean
    
    enabled = False
    But_Incluir.enabled = Grid_Prox_Visita.RowSel > 0 And status <= 1
    If status = 0 Then
        With Grid_Hist
            If .RowSel > 0 Then
                enabled = Grid_Hist.TextMatrix(.RowSel, 3) = ""
            End If
        End With
    End If
    But_Alterar.enabled = enabled Or status = 2
    But_Salvar.enabled = Not Txt_Obs.Locked
    But_Excluir.enabled = enabled
    But_Close.enabled = status = 0
    Grid_Prox_Visita.enabled = status = 0
    Grid_Hist.enabled = status = 0
End Sub

Private Function GET_BACKCOLOR(d As Date) As Long
    If d < #1/1/1900# Then
        GET_BACKCOLOR = RGB(250, 250, 250)
    ElseIf d < Date Then
        GET_BACKCOLOR = RGB(250, 160, 160)
    ElseIf d = Date Then
        GET_BACKCOLOR = RGB(255, 255, 71)
    ElseIf d < DateAdd("d", 7, Date) Then
        GET_BACKCOLOR = RGB(250, 192, 144)
    ElseIf d < DateAdd("m", 1, Date) Then
        GET_BACKCOLOR = RGB(0, 255, 0)
    Else
        GET_BACKCOLOR = 0
    End If
End Function

Private Sub Fill_Grid(rs As ADODB.Recordset, Optional code As String)
    Dim count As Integer
    Dim col As Integer
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim dt As Date
    
    If rs.RecordCount = 0 Then
        MsgBox "Cliente não localizado ou em Arquivo Morto", vbOKOnly, "Xsol Aquecedores"
        Exit Sub
    End If
    
    count = 2
    selectedRow = 2
    With Grid_Prox_Visita
        .Rows = 2
        Do Until rs.EOF
            'Format(MyTime, "h:m:s")
            .AddItem (Get_Field_Val(rs.Fields("v.codigo")) & vbTab & rs.Fields("c.Codigo") & vbTab & Get_Field_Val(rs.Fields("Nome_Cli")) & vbTab + CStr(Get_Field_Val(rs.Fields("Dt_Prox_Visita"))) & vbTab + Format(CStr(Get_Field_Val(rs.Fields("Hora_Prox_Visita"))), "Short Time") & vbTab & Get_Field_Val(rs.Fields("Fone_Celular")) & vbTab & Get_Field_Val(rs.Fields("Fone_Res")) & vbTab & Get_Field_Val(rs.Fields("Endereco_Da_Visita")))
            bkColor = GET_BACKCOLOR(IIf(IsNull(rs.Fields("Dt_Prox_Visita")), vbNull, rs.Fields("Dt_Prox_Visita")))
            If bkColor <> 0 Then
                .row = .Rows - 1
                For col = .FixedCols To .Cols - 1
                    .col = col
                    .CellBackColor = bkColor
                Next
            End If
            If code <> "" Then
                If code = rs.Fields("Cod_Cliente") Then
                    selectedRow = count
                End If
                count = count + 1
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
            Call Grid_Prox_Visita_Click
'            Cmd_Edita.Enabled = True
'            Cmd_Exclui.Enabled = True
        Else
'            Cmd_Edita.Enabled = False
'            Cmd_Exclui.Enabled = False
'            Cmd_Atende.Enabled = False
'            Cmd_Finaliza.Enabled = False
            Call LimpaControles
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Fill_Grid_Hist(rs As ADODB.Recordset, Optional code As String)
    Dim col As Integer
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim dtExp As String
    
    selectedRow = 0
    With Grid_Hist
        .Rows = 2
        Do Until rs.EOF
            If IsNull(rs.Fields("Dt_Expiracao")) Then
                dtExp = ""
            Else
                dtExp = CStr(rs.Fields("Dt_Expiracao"))
            End If
            
            .AddItem (rs.Fields("Codigo") & vbTab & Get_Field_Val(rs.Fields("Dt_Prox_Visita")) & vbTab & rs.Fields("Dt_Criacao") & vbTab & dtExp)
            If IsNull(rs.Fields("Dt_Expiracao")) Then
                .row = .Rows - 1
                bkColor = RGB(0, 255, 0)
                For col = .FixedCols To .Cols - 1
                    .col = col
                    .CellBackColor = bkColor
                Next
            End If
            If code <> "" Then
                If code = rs.Fields("codigo") Then
                    selectedRow = .Rows - 1
                End If
            End If
            rs.MoveNext
        Loop
        If selectedRow = 0 Then
           selectedRow = .Rows - 1
        End If
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            .TopRow = selectedRow
            .col = 0
            .ColSel = .Cols - 1
            Call Grid_Hist_Click
'            Cmd_Edita.Enabled = True
'            Cmd_Exclui.Enabled = True
        Else
            Txt_Obs = ""
'            Cmd_Edita.Enabled = False
'            Cmd_Exclui.Enabled = False
'            Cmd_Atende.Enabled = False
'            Cmd_Finaliza.Enabled = False
'            Call LimpaControles
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Fill_Grid_Endereco(rs As ADODB.Recordset, Optional code As String)
    Dim col As Integer
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim dtExp As String
    
    Dim rsConteudoEndereco As Recordset
    Dim conteudoEndereco As String
    
    If Form_Visit_Manager.Visible = False Then
        Exit Sub
    End If
    
    
    dgEnderecos.Clear
    enderecoAtendimento = ""
    
    If IsNull(rs.Fields("Endereco_da_visita")) Then

        Set rsConteudoEndereco = consultaEnderecoCliente(Txt_Cod_Cliente.Text)
        
        If rsConteudoEndereco.RecordCount = 1 Then
            enderecoAtendimento = rsConteudoEndereco.Fields("endereco")
            Exit Sub
        End If
        
        'dgEnderecos.AddItem (conteudoEndereco)
        
        'koka
        enderecoAtendimento = "Não foi informado endereço para este atendimento!"
        Exit Sub
    End If
    
    enderecoAtendimento = ""
    
    selectedRow = 0
    With dgEnderecos
        .Rows = 2
        .Clear
        Do Until rs.EOF
            .AddItem (rs.Fields("Endereco_da_visita"))
            rs.MoveNext
        Loop
        
        .Refresh
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Configure_Grid()
    With Grid_Prox_Visita
        .Cols = 8
        .FixedCols = 0
        .ColWidth(0) = 750
        .ColWidth(1) = 1000
        .ColWidth(2) = 5700
        .ColWidth(3) = 1100
        .ColWidth(4) = 700
        .ColWidth(5) = 1500
        .ColWidth(6) = 1500
        .ColWidth(7) = 3500 'Endereco do atendimento
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Cód Cliente"
        .TextMatrix(0, 2) = "Nome"
        .TextMatrix(0, 3) = "Próx Visita"
        .TextMatrix(0, 4) = "Horário"
        .TextMatrix(0, 5) = "Celular"
        .TextMatrix(0, 6) = "Telefone"
        .TextMatrix(0, 7) = "End. do atendimento"
        .RowHeight(1) = 0
    End With
    
    With Grid_Hist
        .Cols = 4
        .FixedCols = 0
        .ColWidth(0) = 800
        .ColWidth(1) = 1850
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Visita"
        .TextMatrix(0, 2) = "Dt Criação"
        .TextMatrix(0, 3) = "Dt Expiração"
        .RowHeight(1) = 0
    End With
    
    '07/03/2022, implement do panEnderecos
    With dgEnderecos
        .Cols = 1
        .FixedCols = 0
        .ColWidth(0) = 7000
        
        .ColAlignment(0) = flexAlignLeftCenter
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Endereço"
        .RowHeight(1) = 0
    End With
    
End Sub


Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String, Optional param3 As String) As String
    Dim whereClause As String
    
    If queryType = 1 Then
        Return_SqlString = "SELECT * FROM Clientes c LEFT JOIN (SELECT * FROM Visitas WHERE dt_Expiracao IS NULL) v ON c.codigo = v.cod_cliente WHERE v.Dt_Expiracao IS NULL AND c.Aquivo_Morto = 'N' ORDER BY Dt_Prox_Visita, Hora_Prox_Visita, Nome_Cli"
    ElseIf queryType = 2 Then 'Tabela Visitas filtrando Cod_Cliente, ordenano por código
        Return_SqlString = "SELECT * FROM Visitas WHERE Cod_Cliente = " & param1 & " ORDER BY Codigo"
    ElseIf queryType = 3 Then 'Tabela visitas filtrando somente Codigo
        Return_SqlString = "SELECT * FROM Visitas WHERE Codigo = " & param1
    ElseIf queryType = 4 Then 'Tabela Clientes com Join em Visitas
        Return_SqlString = "SELECT * FROM Clientes c LEFT JOIN (SELECT * FROM Visitas WHERE dt_Expiracao IS NULL) v ON c.codigo = v.cod_cliente"
        'MsgBox Return_SqlString
        If param1 <> "" Then
            whereClause = IIf(whereClause = "", " WHERE", " AND") & " c.codigo = " & param1 & " AND c.Aquivo_Morto = 'N'"
        End If
        If param2 <> "" Then
            whereClause = IIf(whereClause = "", " WHERE", " AND") & " cpf LIKE '%" & param2 & "%'"
        End If
        If param3 <> "" Then
            whereClause = IIf(whereClause = "", " WHERE", " AND") & " nome_cli LIKE '%" & param3 & "%'"
        End If
        Return_SqlString = Return_SqlString & whereClause & " ORDER BY Dt_Prox_Visita, Hora_Prox_Visita"
    ElseIf queryType = 5 Then 'Tabela Configuracao
        Return_SqlString = "SELECT LimiteLigacoesDiarias FROM Configuracao"
    ElseIf queryType = 6 Then 'Tabela Visitas, filtrando Dt_Prox_Visita
        Return_SqlString = "SELECT count(*) as qtd FROM Clientes c LEFT JOIN (SELECT * FROM Visitas WHERE dt_Expiracao IS NULL) v ON c.codigo = v.cod_cliente WHERE v.Dt_Expiracao IS NULL AND c.Aquivo_Morto = 'N' AND v.Dt_Prox_Visita = datevalue('" & Format(param1, "Short Date") & "')"
    ElseIf queryType = 7 Then 'Tabela Visitas, filtrando Dt_Prox_Visita
        'Return_SqlString = "SELECT Count(*) as qtd FROM Visitas Where dt_Expiracao IS NULL AND Dt_Prox_Visita = datevalue('" & Format(param1, "Short Date") & "')"
        Return_SqlString = "SELECT count(*) as qtd FROM Clientes c LEFT JOIN (SELECT * FROM Visitas WHERE dt_Expiracao IS NULL) v ON c.codigo = v.cod_cliente WHERE v.Dt_Expiracao IS NULL AND c.Aquivo_Morto = 'N' AND v.Dt_Prox_Visita = datevalue('" & Format(param1, "Short Date") & "')"
    ElseIf queryType = 8 Then
        Return_SqlString = "Select Endereco_da_visita From visitas where codigo = " & param1
    ElseIf queryType = 9 Then
        Return_SqlString = "SELECT endereco & '-' & bairro & '-' & cidade as endereco From endereco Where Cod_cliente = " & param1
    ElseIf queryType = 10 Then
        Return_SqlString = "SELECT max(codigo) +1 as proximo from visitas"
    End If
End Function

Private Function Return_Recordset(strSql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    rs.Open strSql, W_DB, adOpenStatic, adLockOptimistic
    
    Set Return_Recordset = rs
End Function

Private Sub Grid_Hist_Click()
        
        'krai
        If panEnderecos.Visible = True Then
            panEnderecos.Visible = False
        End If
        
        
        With Grid_Hist
            If .RowSel = 1 Then
                Exit Sub
            End If
            Call Show_Visit_Detail(Return_Recordset(Return_SqlString(3, .TextMatrix(.RowSel, 0))))
            Call Show_Visit_DetailEndereco(Return_Recordset(Return_SqlString(8, Grid_Hist.TextMatrix(Grid_Hist.RowSel, 0))))
        End With
        Call Configure_Controls
        
        Call Fill_Grid_Endereco(Return_Recordset(Return_SqlString(8, Grid_Hist.TextMatrix(Grid_Hist.RowSel, 0))))

        
End Sub

Private Sub Show_Visit_Detail(rs As ADODB.Recordset)
    Txt_Cod_Visita.Text = rs.Fields("Codigo")
    Txt_OS.Text = IIf(IsNull(rs.Fields("Cod_OS")), "", rs.Fields("Cod_OS"))
    Mask_Dt_Prox_Visita.value = rs.Fields("Dt_Prox_Visita")
    Txt_Dt_Criacao.value = Format(rs.Fields("Dt_Criacao"), "Short Date")
    Txt_Dt_Expiracao.Text = IIf(IsNull(rs.Fields("Dt_Expiracao")), "", rs.Fields("Dt_Expiracao"))
    Txt_Obs.Text = IIf(IsNull(rs.Fields("Obs")), "", rs.Fields("Obs"))
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Show_Visit_DetailEndereco(rs As ADODB.Recordset)
    Set rs = Nothing
End Sub

Private Sub Grid_Hist_SelChange()

    Call Grid_Hist_Click
End Sub

Private Sub Grid_Prox_Visita_Click()
        With Grid_Prox_Visita
            If .RowSel = 1 Then
                Exit Sub
            End If
            Call Show_Visit_Hist
        End With
End Sub

Private Sub Show_Visit_Hist()
    With Grid_Prox_Visita
        Txt_Cod_Cliente = .TextMatrix(.RowSel, 1)
        Txt_Nome_Cliente = .TextMatrix(.RowSel, 2)
        txtTelefone.Text = .TextMatrix(.RowSel, 5)
        txtCelular.Text = .TextMatrix(.RowSel, 6)
        Call Fill_Grid_Hist(Return_Recordset(Return_SqlString(2, .TextMatrix(.RowSel, 1))))
    End With
End Sub

Private Sub Grid_Prox_Visita_SelChange()
    
    If panEnderecos.Visible = True Then
        panEnderecos.Visible = False
    End If

    Call Grid_Prox_Visita_Click
    carregaQtdAtendimentosNoDia
End Sub

Sub LimpaControles()
    Txt_Cod_Visita.Text = ""
    Txt_Cod_Cliente.Text = ""
    Txt_Nome_Cliente.Text = ""
    Txt_OS.Text = ""
    'Mask_Dt_Prox_Visita.Text = "__/__/____"
    'Txt_Dt_Criacao.Text = ""
    Txt_Dt_Expiracao.Text = ""
    Txt_Obs.Text = ""
    
    Grid_Hist.Rows = 2
End Sub
