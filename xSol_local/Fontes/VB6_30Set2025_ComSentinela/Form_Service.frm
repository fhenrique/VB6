VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form_Service 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gerenciamento de Atendimentos"
   ClientHeight    =   10335
   ClientLeft      =   915
   ClientTop       =   420
   ClientWidth     =   18210
   Icon            =   "Form_Service.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10335
   ScaleWidth      =   18210
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame Frame5 
      Height          =   345
      Left            =   16005
      TabIndex        =   89
      Top             =   -75
      Width           =   1995
      Begin VB.CheckBox chkTodosAtendimentos 
         Caption         =   "Mostrar todos os anos"
         Height          =   195
         Left            =   45
         TabIndex        =   90
         Top             =   120
         Width           =   1905
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Fichas"
      Height          =   3885
      Left            =   16125
      TabIndex        =   50
      Top             =   6360
      Width           =   1995
      Begin VB.CommandButton btnOpcoes 
         Caption         =   "&Opcões"
         Height          =   375
         Left            =   1185
         TabIndex        =   73
         Top             =   3300
         Width           =   720
      End
      Begin VB.CheckBox Chk_Phone 
         Caption         =   "Mostrar Telefone"
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   2085
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox Chk_PDF 
         Caption         =   "Salvar em PDF"
         Height          =   195
         Left            =   135
         TabIndex        =   57
         Top             =   2745
         Width           =   1980
      End
      Begin VB.OptionButton Option_Inst_Res 
         Caption         =   "Instalação de Resid."
         Height          =   255
         Left            =   75
         TabIndex        =   54
         Top             =   405
         Width           =   1800
      End
      Begin VB.OptionButton Option_Inst_Pisc 
         Caption         =   "Instalação de Piscina"
         Height          =   195
         Left            =   75
         TabIndex        =   53
         Top             =   945
         Width           =   1860
      End
      Begin VB.OptionButton Option_Assistencia 
         Caption         =   "Assistência Técnica"
         Height          =   195
         Left            =   75
         TabIndex        =   52
         Top             =   1530
         Width           =   2055
      End
      Begin VB.CommandButton Cmd_Print 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   195
         TabIndex        =   51
         Top             =   3300
         Width           =   780
      End
   End
   Begin VB.Frame Fr_Finalizacao 
      Caption         =   "Dados de Finalização"
      Height          =   3990
      Left            =   9105
      TabIndex        =   46
      Top             =   6345
      Width           =   6945
      Begin VB.Frame panPecas 
         Enabled         =   0   'False
         Height          =   2520
         Left            =   30
         TabIndex        =   74
         Top             =   660
         Width           =   6885
         Begin VB.TextBox txtValorGeral 
            Height          =   285
            Left            =   5295
            Locked          =   -1  'True
            TabIndex        =   87
            Top             =   2175
            Width           =   1065
         End
         Begin VB.TextBox txtTotalValorPecas 
            Height          =   285
            Left            =   75
            Locked          =   -1  'True
            TabIndex        =   86
            Top             =   2175
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtQtdPecas 
            Height          =   285
            Left            =   6165
            TabIndex        =   83
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtValorPeca 
            Height          =   285
            Left            =   4890
            TabIndex        =   82
            Top             =   495
            Width           =   735
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H008080FF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   6390
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   1425
            Width           =   390
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H0000C000&
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   6390
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   870
            Width           =   390
         End
         Begin VB.ComboBox cboPecas 
            Height          =   315
            Left            =   915
            TabIndex        =   76
            Text            =   "CboPecas"
            Top             =   495
            Width           =   3420
         End
         Begin MSFlexGridLib.MSFlexGrid gridMaterial 
            Height          =   1320
            Left            =   45
            TabIndex        =   79
            Top             =   870
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   2328
            _Version        =   393216
            Rows            =   0
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            SelectionMode   =   1
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Valor total de peças utilizadas no atendimento:"
            Height          =   195
            Left            =   1890
            TabIndex        =   85
            Top             =   2235
            Width           =   3285
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Qtd:"
            Height          =   195
            Left            =   5715
            TabIndex        =   84
            Top             =   570
            Width           =   300
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   195
            Left            =   4440
            TabIndex        =   78
            Top             =   585
            Width           =   405
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   90
            TabIndex        =   77
            Top             =   585
            Width           =   765
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Peças"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   60
            TabIndex        =   75
            Top             =   105
            Width           =   735
         End
      End
      Begin VB.TextBox Text_Obs 
         Height          =   585
         Left            =   600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   3285
         Width           =   6270
      End
      Begin VB.ComboBox Combo_C 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4380
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   300
         Width           =   2520
      End
      Begin VB.ComboBox Combo_G_C 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Text            =   "Combo_G_C"
         Top             =   315
         Width           =   2025
      End
      Begin VB.Label Lbl_Obs 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Obs.:"
         Height          =   195
         Left            =   165
         TabIndex        =   49
         Top             =   3660
         Width           =   375
      End
      Begin VB.Label Lbl_Cob 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cobrança:"
         Height          =   195
         Left            =   3615
         TabIndex        =   48
         Top             =   375
         Width           =   735
      End
      Begin VB.Label Lbl_Cob_Garant 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Garantia / Cobrança:"
         Height          =   195
         Left            =   45
         TabIndex        =   47
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ação"
      Height          =   1095
      Left            =   6000
      TabIndex        =   44
      Top             =   9240
      Width           =   3015
      Begin VB.CommandButton Cmd_Cancela 
         Caption         =   "&Cancelar"
         Height          =   495
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Finaliza 
         BackColor       =   &H000080FF&
         Caption         =   "&Finalizar"
         Height          =   495
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Lista de Atendimentos"
      Height          =   6240
      Left            =   9120
      TabIndex        =   24
      Top             =   0
      Width           =   9015
      Begin MSFlexGridLib.MSFlexGrid Grid_Service 
         Height          =   5865
         Left            =   120
         TabIndex        =   25
         Top             =   285
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   10345
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Edição"
      Height          =   1095
      Left            =   120
      TabIndex        =   19
      Top             =   9240
      Width           =   5805
      Begin VB.CommandButton Cmd_Inclui 
         BackColor       =   &H0080FF80&
         Caption         =   "&Incluir"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Salva 
         Caption         =   "&Salvar"
         Height          =   495
         Left            =   3000
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Edita 
         Caption         =   "&Editar"
         Height          =   495
         Left            =   1560
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Exclui 
         BackColor       =   &H008080FF&
         Caption         =   "E&xcluir"
         Height          =   495
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Atendimento"
      Height          =   9180
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   8895
      Begin VB.ComboBox Combo_Servico2 
         Height          =   315
         Left            =   2400
         TabIndex        =   65
         Text            =   "Combo_Servico2"
         Top             =   1890
         Width           =   4950
      End
      Begin VB.TextBox txtObsPreFinalizacao 
         Enabled         =   0   'False
         Height          =   465
         Left            =   4485
         MaxLength       =   110
         MultiLine       =   -1  'True
         TabIndex        =   62
         Top             =   2505
         Width           =   4110
      End
      Begin VB.TextBox Text_Hr_Serv 
         Height          =   285
         Left            =   2280
         MaxLength       =   36
         TabIndex        =   3
         Top             =   2655
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Endereços"
         Height          =   3405
         Left            =   240
         TabIndex        =   28
         Top             =   3405
         Width           =   8415
         Begin MSFlexGridLib.MSFlexGrid Grid_Address 
            Height          =   1665
            Left            =   60
            TabIndex        =   29
            Top             =   1695
            Width           =   8310
            _ExtentX        =   14658
            _ExtentY        =   2937
            _Version        =   393216
            Rows            =   1
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            SelectionMode   =   1
         End
         Begin VB.Label Label20 
            Caption         =   "Duplo clique no endereço para visualizar somente os atendimentos dele."
            ForeColor       =   &H00808080&
            Height          =   165
            Left            =   120
            TabIndex        =   88
            Top             =   1500
            Width           =   6510
         End
         Begin VB.Label Lbl_Address_Cod 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   840
            TabIndex        =   56
            Top             =   330
            Width           =   585
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   195
            TabIndex        =   55
            Top             =   330
            Width           =   540
         End
         Begin VB.Label Lbl_Zona 
            AutoSize        =   -1  'True
            Caption         =   "Zona"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   6840
            TabIndex        =   43
            Top             =   990
            Width           =   375
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Zona:"
            Height          =   195
            Left            =   6315
            TabIndex        =   42
            Top             =   990
            Width           =   420
         End
         Begin VB.Label Lbl_Ref 
            AutoSize        =   -1  'True
            Caption         =   "Ref"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   2520
            TabIndex        =   41
            Top             =   990
            Width           =   255
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ref.:"
            Height          =   195
            Left            =   2070
            TabIndex        =   40
            Top             =   990
            Width           =   345
         End
         Begin VB.Label Lbl_CEP 
            AutoSize        =   -1  'True
            Caption         =   "CEP"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   840
            TabIndex        =   39
            Top             =   990
            Width           =   315
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
            Height          =   195
            Left            =   375
            TabIndex        =   38
            Top             =   990
            Width           =   360
         End
         Begin VB.Label Lbl_UF 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   7680
            TabIndex        =   37
            Top             =   645
            Width           =   210
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Left            =   7320
            TabIndex        =   36
            Top             =   645
            Width           =   255
         End
         Begin VB.Label Lbl_Cidade 
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   4320
            TabIndex        =   35
            Top             =   645
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Left            =   3675
            TabIndex        =   34
            Top             =   645
            Width           =   540
         End
         Begin VB.Label Lbl_Bairro 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   840
            TabIndex        =   33
            Top             =   645
            Width           =   405
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Left            =   285
            TabIndex        =   32
            Top             =   645
            Width           =   450
         End
         Begin VB.Label Lbl_End 
            AutoSize        =   -1  'True
            Caption         =   "End:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   2520
            TabIndex        =   31
            Top             =   330
            Width           =   330
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "End.:"
            Height          =   195
            Left            =   2040
            TabIndex        =   30
            Top             =   315
            Width           =   375
         End
      End
      Begin VB.TextBox Text_Tecnico_Cod 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   27
         Top             =   1095
         Width           =   415
      End
      Begin VB.TextBox Text_Status 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   330
         Width           =   2295
      End
      Begin VB.ComboBox Combo_Ficha 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3030
         Width           =   1935
      End
      Begin VB.ComboBox Combo_Tecnico 
         Height          =   315
         Left            =   2760
         TabIndex        =   1
         Text            =   "Combo_Tecnico"
         Top             =   1095
         Width           =   4590
      End
      Begin VB.ComboBox Combo_Servico 
         Height          =   315
         Left            =   2400
         TabIndex        =   2
         Text            =   "Combo_Servico"
         Top             =   1515
         Width           =   4965
      End
      Begin VB.TextBox Text_Valor 
         Height          =   285
         Left            =   4980
         TabIndex        =   5
         Top             =   3105
         Width           =   1605
      End
      Begin VB.TextBox Text_OS 
         Alignment       =   2  'Center
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
         Left            =   2280
         TabIndex        =   11
         Top             =   330
         Width           =   855
      End
      Begin VB.TextBox Text_Nome_Cliente 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   10
         Top             =   735
         Width           =   4095
      End
      Begin VB.TextBox Text_Cod_Cliente 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   9
         Top             =   735
         Width           =   855
      End
      Begin MSMask.MaskEdBox Mask_Dt_Serv 
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
         Left            =   2295
         TabIndex        =   63
         Top             =   2295
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Equipment 
         Height          =   1515
         Left            =   195
         TabIndex        =   67
         Top             =   7200
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   2672
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin VB.Label lblOsSelecionada 
         Caption         =   "..."
         Height          =   195
         Left            =   3525
         TabIndex        =   72
         Top             =   8700
         Width           =   855
      End
      Begin VB.Label Lbl_Equip_Cod 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2115
         TabIndex        =   71
         Top             =   6600
         Width           =   1125
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Legenda:"
         Height          =   195
         Left            =   165
         TabIndex        =   70
         Top             =   8715
         Width           =   675
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         Caption         =   "Equipamento(s) atendido(s) pela OS "
         Height          =   195
         Left            =   990
         TabIndex        =   69
         Top             =   8715
         Width           =   2580
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Equipamentos do cliente no endereço acima:"
         Height          =   195
         Left            =   210
         TabIndex        =   68
         Top             =   6960
         Width           =   3195
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   195
         Left            =   2235
         TabIndex        =   66
         Top             =   1965
         Width           =   90
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Left            =   2235
         TabIndex        =   64
         Top             =   1590
         Width           =   90
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Obs. Pré-finalização:"
         Height          =   195
         Left            =   4470
         TabIndex        =   61
         Top             =   2250
         Width           =   1455
      End
      Begin VB.Label Lbl_Hr_Serv 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Período:"
         Height          =   195
         Left            =   1380
         TabIndex        =   58
         Top             =   2715
         Width           =   615
      End
      Begin VB.Label Lbl_Tec 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Técnico:"
         Height          =   195
         Left            =   1395
         TabIndex        =   18
         Top             =   1095
         Width           =   630
      End
      Begin VB.Label Lbl_Serv 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Serviço Solicitado:"
         Height          =   195
         Left            =   705
         TabIndex        =   17
         Top             =   1575
         Width           =   1320
      End
      Begin VB.Label Lbl_DtServ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Serv.:"
         Height          =   195
         Left            =   1215
         TabIndex        =   16
         Top             =   2355
         Width           =   810
      End
      Begin VB.Label Lbl_Atend 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ficha de Atendimento:"
         Height          =   195
         Left            =   435
         TabIndex        =   15
         Top             =   3030
         Width           =   1590
      End
      Begin VB.Label Lbl_Valor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Left            =   4470
         TabIndex        =   14
         Top             =   3105
         Width           =   405
      End
      Begin VB.Label Lbl_Registro 
         Alignment       =   1  'Right Justify
         Caption         =   "OS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1410
         TabIndex        =   13
         Top             =   330
         Width           =   615
      End
      Begin VB.Label Lbl_Cliente 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   1500
         TabIndex        =   12
         Top             =   735
         Width           =   525
      End
   End
   Begin VB.Image Image2 
      Height          =   195
      Left            =   15570
      Picture         =   "Form_Service.frx":000C
      Top             =   8055
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   14400
      Picture         =   "Form_Service.frx":0270
      Top             =   8760
      Visible         =   0   'False
      Width           =   3000
   End
End
Attribute VB_Name = "Form_Service"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conn As ADODB.Connection
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'Const PDF_PRINTER_NAME = "Solid PDF Creator"
Const PDF_PRINTER_NAME = "CutePDF Writer"
Const strChecked = "þ"
Const strUnChecked = "q"


Public publicValorAtendimento As Double 'usado para guardar o valor original do atendimento, sem as peças
Public publicAlterouPecas As Boolean
Public pubTecnicoBloqueado As Boolean
Public strCodEndereco As String
Public strEndereco As String
Public strBairro As String
Public strCidade As String
Public strUF As String
Public strCEP As String

Public qtdServicos As Integer
Public qtdEquipamentos As Integer


Private Function PrinterIndex(ByVal printerName As String) As Integer
    Dim I As Integer
    
    For I = 0 To Printers.count - 1
        If LCase(Printers(I).DeviceName) Like LCase(printerName) Then
            PrinterIndex = I
            Exit Function
        End If
    Next
    PrinterIndex = -1
End Function

Private Sub btnOpcoes_Click()
    FrmFichasAtendimentoOpcoes.Show (1)
End Sub

Private Sub cboPecas_Click()
    CarregaPreco (Mid(cboPecas.Text, 1, 3))
End Sub

Private Sub chkTodosAtendimentos_Click()
    If chkTodosAtendimentos.value = 0 Then
        'Mostra somente registros de até um ano
        Call Show_Records(Return_Recordset(Return_SqlString(1, Str(Cod_Cli))))
    Else
        'Mostra todos os registros
        Call Show_Records(Return_Recordset(Return_SqlString(26, Str(Cod_Cli))))
    End If
End Sub

Private Sub Cmd_Cancela_Click()
    If Cmd_Cancela.Caption = "&Cancela" Then
        Call Object_Controller(6)
    Else
        Call Object_Controller(7)
    End If
End Sub

Private Sub Combo_Tecnico_Click()
    Dim rs As ADODB.Recordset
    
    If Not Combo_Tecnico.enabled Or Combo_Tecnico.ListIndex < 0 Then
        Exit Sub
    End If
    
    Set rs = Return_Recordset(Return_SqlString(11, Combo_Tecnico.Text))
    Text_Tecnico_Cod.Text = rs.Fields("Codigo")
    rs.Close
    Set rs = Nothing
End Sub

Private Sub carregaGarantias(rs As ADODB.Recordset, Optional code As String)
    
    With Combo_G_C
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("Descricao"))
            rs.MoveNext
        Loop
    End With
    
    rs.Close
    Set rs = Nothing
    
End Sub

Private Sub carregaCobrancas(rs As ADODB.Recordset, Optional code As String)
    
    With Combo_C
        '.Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("Descricao"))
            rs.MoveNext
        Loop
    End With
    
    rs.Close
    Set rs = Nothing
    
End Sub


Private Sub Command1_Click()
    
    Dim valorTotal As Double
    Dim contador As Integer
    Dim ValorAcrescentar As Double
     
    If txtQtdPecas.Text = "" Then
        txtQtdPecas.Text = "1"
    End If
     
    If Trim(txtValorPeca.Text) = "" Then
        MsgBox "Informe o valor da peça!"
        Exit Sub
    End If
    
    valorTotal = CDbl(txtQtdPecas.Text) * CDbl(txtValorPeca.Text)
    valorTotal = Format(valorTotal, "###,###,###.##")
    
    With gridMaterial
        .AddItem (cboPecas.Text & vbTab & txtValorPeca.Text & vbTab & txtQtdPecas.Text)
        .Refresh
    End With
    
    For contador = 2 To gridMaterial.Rows - 1
        ValorAcrescentar = ValorAcrescentar + (gridMaterial.TextMatrix(contador, 2) * gridMaterial.TextMatrix(contador, 1))
    Next contador
    
    txtValorGeral.Text = ValorAcrescentar
    txtValorGeral.Text = FormatCurrency(txtValorGeral.Text)
    Text_Valor.Text = publicValorAtendimento + ValorAcrescentar
    Text_Valor.Text = FormatCurrency(Text_Valor.Text)
    
    publicAlterouPecas = True
    
    
End Sub

Private Sub Command2_Click()
    
    'botão para remover peças do atendimento
    
    Dim contador As Integer
    Dim ValorAcrescentar As Double
    
    With gridMaterial
        If .RowSel > 1 Then
            
            'Parei aqui, vou remover isso aí em baixo e levar para o botão de salvar, tira tudo e depois pooe tudo
            'Remove a peca do banco de dados
            
            Call removePeca(Text_OS.Text, Mid(gridMaterial.TextMatrix(.RowSel, 0), 1, 2))
            
            'Remove do grid
            .RemoveItem .RowSel
            .Refresh
        Else
            MsgBox "Não há peça selecionada para remover"
            Exit Sub
        End If
    End With
    
    
    For contador = 2 To gridMaterial.Rows - 1
        ValorAcrescentar = ValorAcrescentar + (gridMaterial.TextMatrix(contador, 2) * gridMaterial.TextMatrix(contador, 1))
    Next contador
    
    Text_Valor.Text = publicValorAtendimento + ValorAcrescentar
    Text_Valor.Text = FormatCurrency(Text_Valor.Text)
    
    gridMaterial.Refresh
    
    If (gridMaterial.Rows - 2) > 0 Then
        txtValorGeral.Text = ValorAcrescentar
        txtValorGeral.Text = FormatCurrency(txtValorGeral.Text)
    Else
        txtValorGeral.Text = ""
    End If
    
    gridMaterial.Refresh
    
    publicAlterouPecas = True
    
    
    
End Sub

Private Sub Form_Activate()
    If (MDIForm_Xsol.acessibilidade_TextoNegrito = True) Or (MDIForm_Xsol.acessibilidade_TextoNegrito = False) Then
        Lbl_Obs.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Cob.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Cob_Garant.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Address_Cod.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label8.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Zona.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label7.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Ref.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label6.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Cep.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label5.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_UF.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label4.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Cidade.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label3.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Bairro.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label2.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_End.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label1.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        lblOsSelecionada.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Equip_Cod.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label14.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label13.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label9.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label12.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label11.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label10.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Hr_Serv.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Tec.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Serv.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_DtServ.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Atend.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Valor.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Registro.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Cliente.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Combo_C.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Combo_G_C.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Combo_Servico2.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Combo_Ficha.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Combo_Tecnico.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Combo_Servico.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Obs.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        txtObsPreFinalizacao.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Hr_Serv.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Tecnico_Cod.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Status.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Valor.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_OS.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Nome_Cliente.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Cod_Cliente.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Grid_Service.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Grid_Address.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Grid_Equipment.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame9.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Fr_Finalizacao.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame3.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame6.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame4.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame1.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame2.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        btnOpcoes.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Print.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Cancela.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Finaliza.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Inclui.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Salva.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Edita.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Exclui.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Chk_Phone.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Chk_PDF.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Option_Inst_Res.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Option_Inst_Pisc.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Option_Assistencia.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    End If
    
    global_abrindoForm = False

End Sub

Private Sub Form_Load()
   
   Me.Caption = Me.Caption & " - " & versaoAtual
   
   Set conn = New ADODB.Connection
   conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"

   Call carregaGarantias(Return_Recordset(Return_SqlString(19)))
   Call carregaCobrancas(Return_Recordset(Return_SqlString(20)))


   Text_Cod_Cliente.Text = Cod_Cli
   Text_Nome_Cliente.Text = Nome_Cli_Atend
   
   Call Configure_Grid
   Call Object_Controller(1)
   
   'endereços do cliente
   Call Show_Customer_Address(Return_Recordset(Return_SqlString(6, Str(Cod_Cli))))
   
   'carrega aqui o grid_services
   Call Show_Records(Return_Recordset(Return_SqlString(1, Str(Cod_Cli))))

    
    pintaLinhasEquipamentosDaOs
   
    '-- Atualizacao 9 - Peças do atendimento
        'carrega as pecas no combo
        Dim rs_pecas As Recordset
        Set rs_pecas = RetornaPecas()
        
        Dim contador As Integer
        Dim qtdPecasTotal As Integer
        
        Set rs_pecas = Return_Recordset("SELECT Codigo & '   -' & Descricao as Peca From Material")
        qtdPecasTotal = rs_pecas.RecordCount
        
        With cboPecas
            .Clear
            For contador = 1 To qtdPecasTotal
                .AddItem (rs_pecas("Peca"))
                rs_pecas.MoveNext
            Next contador
        End With
        
        publicAlterouPecas = False
     
    
     
End Sub

Private Sub Cmd_Edita_Click()
    If Cmd_Edita.Caption = "&Cancela" Then
        Call Object_Controller(6)
    Else
        Call Object_Controller(4)
        Call setGridDefault(Grid_Address)
        Call setGridDefault(Grid_Equipment)
    End If
End Sub

Private Sub Cmd_Inclui_Click()
    If Cmd_Inclui.Caption = "&Cancela" Then
        Call Object_Controller(6)
    Else
        Combo_Servico.Text = ""
        Combo_Servico2.Text = ""
        Call Object_Controller(3)
        Call setGridDefault(Grid_Address)
        Call setGridDefault(Grid_Equipment)
        despintaLinhasEquipamentosDaOs
    End If
End Sub

Private Sub Cmd_Salva_Click()
    Dim rs As ADODB.Recordset
    'Dim rs2 As ADODB.Recordset
    
    Dim rspecas As ADODB.Recordset
    Dim contador As Integer
    
    Dim diasNaoDisponivel As String
    Dim mesAtual As Integer
    Dim strMesNaoDisponivel As String
    Dim intMesNaoDisponivel As Integer
    
    'SOMENTE PARA A AGENDA DO TECNICO, PARA VER SE ELE ESTÁ ATIVO
    Dim dbFun As DAO.Database
    Dim rsFun As DAO.Recordset
    Dim strSql As String
    
    If Not Mandatory_Fields_Filled Then
        Exit Sub
    End If
    
    'se não tiver equipamento sai também
    If Grid_Equipment.Rows = 2 Then
        MsgBox "É necessário criar ao menos 1 equipamento para este cliente!", vbInformation, "Cliente sem equipamento!"
        Exit Sub
    End If
    
    
    Call Object_Controller(4)
        
    'só entra na valida domingo e na valida agenda
    'se o combobox garantia / cobrança estiver enabled = false
    If Combo_G_C.enabled = False Then

        'É domingo?
        If Weekday(Mask_Dt_Serv.Text, vbSunday) = vbSunday Then
            MsgBox "A data informada é um domingo, informe outro dia.", vbOKOnly, "Xsol Aquecedores"
            Mask_Dt_Serv.SetFocus
            Exit Sub
        End If

        'INICIA Valida agenda
        'strSql = "SELECT count(*) as qtd FROM Tecnico WHERE Codigo = " & Text_Tecnico_Cod.Text & " AND (DiasNaoDisponivel LIKE '*" & Mid(Mask_Dt_Serv.Text, 1, 2) & "*' OR DiasNaoDisponivel = 'M')"
        Dim mesDoAtendimento As String
        mesDoAtendimento = converteNumeroEmNome(Mid(Mask_Dt_Serv.Text, 4, 2))
        
        
        
        strSql = "SELECT count(*) as qtd FROM tecnicos_agenda WHERE Codigo = '" & Text_Tecnico_Cod.Text & "' AND mes = '" & mesDoAtendimento & "' and Dias LIKE '*" & Mid(Mask_Dt_Serv.Text, 1, 2) & "*'"
        
        Set dbFun = OpenDatabase(W_DBname)
        Set rsFun = dbFun.OpenRecordset(strSql)
        
        
        'Verifica aqui a agenda do técnico
        If rsFun.Fields("qtd") > 0 Then
                MsgBox "O técnico " & Combo_Tecnico.Text & " não está disponível em " & Mask_Dt_Serv.Text
                
                rsFun.Close
                Set rsFun = Nothing
                
                dbFun.Close
                Set dbFun = Nothing
                
                Combo_Tecnico.SetFocus
                
                Exit Sub
        Else
            rsFun.Close
            Set rsFun = Nothing
            
            dbFun.Close
            Set dbFun = Nothing
        End If
    End If
    'FIM Valida agenda
    
    
    If Cmd_Inclui.Caption = "&Cancela" Then
        Set rs = Return_Recordset(Return_SqlString(3))
        rs.AddNew
        Text_OS.Text = rs.Fields("codigo")
        rs.Fields("Cod_Cliente") = Text_Cod_Cliente.Text
        rs.Fields("Dt_Status") = Null
        
    Else
        'edita
        Set rs = Return_Recordset(Return_SqlString(13, Text_OS.Text))
        
        'atualiza aqui o ExibeNaFicha
        Dim strExibeNaFicha As String
        strExibeNaFicha = Grid_Equipment.TextMatrix(Grid_Equipment.RowSel, 2)
        
        Call atualizaEquipamento(Grid_Equipment.TextMatrix(Grid_Equipment.RowSel, 0), strExibeNaFicha)
        
    End If
    
    
    Call Record_Record(rs)

    rs.Update

    rs.Close
    Set rs = Nothing
    
    
    'Implmentação das peças do atendimento, insert na tabela ATENDIMENTO_PECAS
    If publicAlterouPecas Then
        Set rspecas = Return_Recordset(Return_SqlString(22))
        
        For contador = 2 To gridMaterial.Rows - 1
            rspecas.AddNew
            rspecas.Fields("CODIGO_ATENDIMENTO") = Text_OS.Text
            rspecas.Fields("CODIGO_MATERIAL") = Mid(gridMaterial.TextMatrix(contador, 0), 1, 3)
            rspecas.Fields("QTD") = gridMaterial.TextMatrix(contador, 2)
            rspecas.Fields("VALOR") = gridMaterial.TextMatrix(contador, 1)
        Next contador
        
        If rspecas.State = 1 Then
            rspecas.Update
        End If
        
        rspecas.Close
        Set rspecas = Nothing
        
        txtValorPeca.Text = ""
        txtQtdPecas.Text = ""
        
        gridMaterial.Rows = 2
        
        txtValorGeral.Text = ""
        publicAlterouPecas = False
    End If
    
    Call Object_Controller(6)
    Call Show_Records(Return_Recordset(Return_SqlString(1, Text_Cod_Cliente.Text)))
    
    Call Grid_Service_Click
    
End Sub

Private Function converteNumeroEmNome(numeroDoMes As String) As String
    If numeroDoMes = 1 Then
        converteNumeroEmNome = "JAN"
    ElseIf numeroDoMes = 2 Then
       converteNumeroEmNome = "FEV"
    ElseIf numeroDoMes = 3 Then
       converteNumeroEmNome = "MAR"
    ElseIf numeroDoMes = 4 Then
       converteNumeroEmNome = "ABR"
    ElseIf numeroDoMes = 5 Then
       converteNumeroEmNome = "MAI"
    ElseIf numeroDoMes = 6 Then
       converteNumeroEmNome = "JUN"
    ElseIf numeroDoMes = 7 Then
       converteNumeroEmNome = "JUL"
    ElseIf numeroDoMes = 8 Then
       converteNumeroEmNome = "AGO"
    ElseIf numeroDoMes = 9 Then
       converteNumeroEmNome = "SET"
    ElseIf numeroDoMes = 10 Then
       converteNumeroEmNome = "OUT"
    ElseIf numeroDoMes = 11 Then
       converteNumeroEmNome = "NOV"
    ElseIf numeroDoMes = 12 Then
       converteNumeroEmNome = "DEZ"
    End If
End Function

Private Sub Cmd_Exclui_Click()
    
    If Grid_Service.Rows - 2 = 0 Then
        Cmd_Exclui.enabled = False
    Else
    
    If MsgBox("Confirma a Exclusão da OS " & Text_OS.Text & " do Cliente " & Text_Nome_Cliente.Text & " ? ", vbYesNo, "Excluir Registros") = vbYes Then
        Call Return_Recordset(Return_SqlString(13, Text_OS.Text)).Delete
        'Call Object_Controller(1)
        Call Show_Records(Return_Recordset(Return_SqlString(1, Text_Cod_Cliente.Text)))
    End If
    
    End If
End Sub

Private Sub Cmd_Finaliza_Click()
        

    
    If Cmd_Finaliza.Caption = "&Cancela" Then
        panPecas.enabled = False
        Call Object_Controller(6)
    Else
        panPecas.enabled = True
        Call Object_Controller(5)
    End If
End Sub

Private Function Set_Printer(ByVal printerName As String) As Boolean
    Dim prtidx As Integer
    
    Set_Printer = False

    Rem -- Find the index of the printer
    prtidx = PrinterIndex(printerName)
    If prtidx < 0 Then
        MsgBox "Não foi encontrada a impressora '" & printerName & "'. Por favor instale a impressora", vbCritical, "Xsol Aquecedores"
        Exit Function
    End If

    Rem -- Set the current printer
    Set Printer = Printers(prtidx)

    Set_Printer = True
End Function

Private Sub Cmd_Print_Click()

    If (Option_Inst_Res.value = False) And (Option_Inst_Pisc.value = False) And (Option_Assistencia.value = False) Then
        MsgBox "Informe um tipo de ficha! Residência, Piscina ou Assistência Técnica"
        Exit Sub
    End If
    
    If Text_Hr_Serv.Text = "" Then
        MsgBox "Informe o período do atendimento"
        Cmd_Edita_Click
        Text_Hr_Serv.SetFocus
        Exit Sub
    End If

    If Chk_Phone.value = 0 Then
        If MsgBox("A opção de exibir o telefone na ficha está desabilitada, deseja habilitar desta vez?", vbYesNo) = vbYes Then
            Chk_Phone.value = 1
        End If
    End If

    If Not Set_Printer(IIf(Chk_PDF.value = 1, PDF_PRINTER_NAME, Default_Printer)) Then
        Exit Sub
    End If
    
    If Option_Assistencia Then
        relFicha_Assistencia_Tecnica.lblNumeroOS.Caption = "OS: X" & Text_OS.Text
        relFicha_Assistencia_Tecnica.lblDataDoServico.Caption = Mask_Dt_Serv.Text
        relFicha_Assistencia_Tecnica.lblPeriodo.Caption = Text_Hr_Serv.Text
        relFicha_Assistencia_Tecnica.lblTecnicoResponsavel.Caption = Combo_Tecnico.Text
        relFicha_Assistencia_Tecnica.lblCodigoCliente.Caption = Text_Cod_Cliente.Text
        relFicha_Assistencia_Tecnica.lblNomeCliente.Caption = Text_Nome_Cliente.Text
        relFicha_Assistencia_Tecnica.lblEndereco.Caption = Lbl_End.Caption
        relFicha_Assistencia_Tecnica.lblBairro.Caption = Lbl_Bairro.Caption
        relFicha_Assistencia_Tecnica.lblZona.Caption = Lbl_Zona.Caption
        relFicha_Assistencia_Tecnica.lblCep.Caption = Lbl_Cep.Caption
        relFicha_Assistencia_Tecnica.lblCidade.Caption = Lbl_Cidade.Caption
        relFicha_Assistencia_Tecnica.lblReferencia.Caption = Lbl_Ref.Caption
        
        'se mostra ou não os telefones
        If Chk_Phone.value = 1 Then
            relFicha_Assistencia_Tecnica.lblTelRes.Caption = global_Ficha_telefone1
            relFicha_Assistencia_Tecnica.lblTelCom.Caption = global_Ficha_telefone2
            relFicha_Assistencia_Tecnica.lblCelular.Caption = global_Ficha_telefone3
        Else
            relFicha_Assistencia_Tecnica.lblViewTelRes.Visible = False
            relFicha_Assistencia_Tecnica.lblViewTelCom.Visible = False
            relFicha_Assistencia_Tecnica.lblViewTelCel.Visible = False
            relFicha_Assistencia_Tecnica.lblTelRes.Visible = False
            relFicha_Assistencia_Tecnica.lblTelCom.Visible = False
            relFicha_Assistencia_Tecnica.lblCelular.Visible = False
        End If
        
        'EQUIPAMENTOS
        If Form_Service.Grid_Equipment.Rows > 2 Then
            
            'Se só tem 1 equipamento    ok
            If Grid_Equipment.Rows - 2 = 1 Then
                
                relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                
                'relFicha_Assistencia_Tecnica.
                
                GoTo servicos
            End If
            
            'Se tem 2 equipamentos
            If Grid_Equipment.Rows - 2 = 2 Then
            
                relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
            
                'equipamento 1 = S
                'equipamento 2 = N  feito
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And ((Grid_Equipment.TextMatrix(3, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    GoTo servicos
                End If
            
                'equipamento 1 = N
                'equipamento 2 = S  feito
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "S")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(3, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = N
                'equipamento 2 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = False
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    GoTo servicos
                End If
                
                'equipamento 1 = S
                'equipamento 2 = S
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And ((Grid_Equipment.TextMatrix(3, 2) = "S")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(3, 1)
                    GoTo servicos
                End If
                
            End If
            
            '----------
            'Se tem 3 equipamentos
            If Grid_Equipment.Rows - 2 = 3 Then
            
                'equipamento 1 = S
                'equipamento 2 = S
                'equipamento 3 = S
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And ((Grid_Equipment.TextMatrix(3, 2) = "S")) And ((Grid_Equipment.TextMatrix(4, 2) = "S")) Then
                    
                    '------------------------ Mostra os 3 equipamentos
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = True
                    '-------------------------
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(3, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Caption = Grid_Equipment.TextMatrix(4, 1)
                    GoTo servicos
                End If
            
                'equipamento 1 = N
                'equipamento 2 = N  FEITO
                'equipamento 3 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "N")) And ((Grid_Equipment.TextMatrix(4, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = False
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    GoTo servicos
                End If
                
                'equipamento 1 = S
                'equipamento 2 = S  FEITO
                'equipamento 3 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And ((Grid_Equipment.TextMatrix(3, 2) = "S")) And ((Grid_Equipment.TextMatrix(4, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(3, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = S
                'equipamento 2 = N  FEITO
                'equipamento 3 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And ((Grid_Equipment.TextMatrix(3, 2) = "N")) And ((Grid_Equipment.TextMatrix(4, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    GoTo servicos
                End If
                
                
                'equipamento 1 = S
                'equipamento 2 = N  FEITO
                'equipamento 3 = S
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And ((Grid_Equipment.TextMatrix(3, 2) = "N")) And ((Grid_Equipment.TextMatrix(4, 2) = "S")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(4, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = N
                'equipamento 2 = N  FEITO
                'equipamento 3 = S
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "N")) And ((Grid_Equipment.TextMatrix(4, 2) = "S")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(4, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = N
                'equipamento 2 = S FEITO
                'equipamento 3 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "S")) And ((Grid_Equipment.TextMatrix(4, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(3, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = N
                'equipamento 2 = S FEITO
                'equipamento 3 = S
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "S")) And ((Grid_Equipment.TextMatrix(4, 2) = "S")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(3, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(4, 1)
                    GoTo servicos
                End If
                
                '5 equipamentos, último = S
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "N")) And ((Grid_Equipment.TextMatrix(4, 2) = "N")) And ((Grid_Equipment.TextMatrix(5, 2) = "N")) And ((Grid_Equipment.TextMatrix(6, 2) = "S")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(6, 1)
                    GoTo servicos
                End If
                
                
            End If
            
            'Se tem 4 equipamentos
            If Grid_Equipment.Rows - 2 = 4 Then
                
                'equipamento 1 = N
                'equipamento 2 = N
                'equipamento 3 = N
                'equipamento 4 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "N")) And ((Grid_Equipment.TextMatrix(4, 2) = "N")) And ((Grid_Equipment.TextMatrix(5, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    'relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    'relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(3, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = S
                'equipamento 2 = N
                'equipamento 3 = N
                'equipamento 4 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And ((Grid_Equipment.TextMatrix(3, 2) = "N")) And ((Grid_Equipment.TextMatrix(4, 2) = "N")) And ((Grid_Equipment.TextMatrix(5, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    'relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(3, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = S
                'equipamento 2 = S
                'equipamento 3 = N
                'equipamento 4 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And ((Grid_Equipment.TextMatrix(3, 2) = "S")) And ((Grid_Equipment.TextMatrix(4, 2) = "N")) And ((Grid_Equipment.TextMatrix(5, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(3, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = S
                'equipamento 2 = S
                'equipamento 3 = S
                'equipamento 4 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And ((Grid_Equipment.TextMatrix(3, 2) = "S")) And ((Grid_Equipment.TextMatrix(4, 2) = "S")) And ((Grid_Equipment.TextMatrix(5, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(3, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Caption = Grid_Equipment.TextMatrix(4, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = S
                'equipamento 2 = S
                'equipamento 3 = S
                'equipamento 4 = S
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And ((Grid_Equipment.TextMatrix(3, 2) = "S")) And ((Grid_Equipment.TextMatrix(4, 2) = "S")) And ((Grid_Equipment.TextMatrix(5, 2) = "S")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento4.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento4.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(3, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Caption = Grid_Equipment.TextMatrix(4, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento4.Caption = Grid_Equipment.TextMatrix(5, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = S
                'equipamento 2 = N
                'equipamento 3 = S
                'equipamento 4 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And ((Grid_Equipment.TextMatrix(3, 2) = "N")) And ((Grid_Equipment.TextMatrix(4, 2) = "S")) And ((Grid_Equipment.TextMatrix(5, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(3, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = S
                'equipamento 2 = N
                'equipamento 3 = N
                'equipamento 4 = S
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And ((Grid_Equipment.TextMatrix(3, 2) = "N")) And ((Grid_Equipment.TextMatrix(4, 2) = "N")) And ((Grid_Equipment.TextMatrix(5, 2) = "S")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(5, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = N
                'equipamento 2 = S
                'equipamento 3 = S
                'equipamento 4 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "S")) And ((Grid_Equipment.TextMatrix(4, 2) = "S")) And ((Grid_Equipment.TextMatrix(5, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(3, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = N
                'equipamento 2 = S
                'equipamento 3 = N
                'equipamento 4 = S
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "S")) And ((Grid_Equipment.TextMatrix(4, 2) = "N")) And ((Grid_Equipment.TextMatrix(5, 2) = "S")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(3, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = N
                'equipamento 2 = N
                'equipamento 3 = S
                'equipamento 4 = S
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "N")) And ((Grid_Equipment.TextMatrix(4, 2) = "S")) And ((Grid_Equipment.TextMatrix(5, 2) = "S")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(3, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = N
                'equipamento 2 = S
                'equipamento 3 = N
                'equipamento 4 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "S")) And ((Grid_Equipment.TextMatrix(4, 2) = "N")) And ((Grid_Equipment.TextMatrix(5, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(3, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = N
                'equipamento 2 = N
                'equipamento 3 = S
                'equipamento 4 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "N")) And ((Grid_Equipment.TextMatrix(4, 2) = "S")) And ((Grid_Equipment.TextMatrix(5, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False

                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(4, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = N
                'equipamento 2 = N
                'equipamento 3 = N
                'equipamento 4 = S
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "N")) And ((Grid_Equipment.TextMatrix(4, 2) = "N")) And ((Grid_Equipment.TextMatrix(5, 2) = "S")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False

                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(5, 1)
                    GoTo servicos
                End If
                
                
                
                
            End If
            
            
            '-- SE TEM 5 EQUIPAMENTOS
            If Grid_Equipment.Rows - 2 = 5 Then
                
                'equipamento 1 = N
                'equipamento 2 = N
                'equipamento 3 = N
                'equipamento 4 = N
                'equipamento 4 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "N")) And ((Grid_Equipment.TextMatrix(4, 2) = "N")) And ((Grid_Equipment.TextMatrix(5, 2) = "N")) And ((Grid_Equipment.TextMatrix(6, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    'relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(3, 1)
                    'relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(4, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = S
                'equipamento 2 = N
                'equipamento 3 = N
                'equipamento 4 = N
                'equipamento 4 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And ((Grid_Equipment.TextMatrix(3, 2) = "N")) And ((Grid_Equipment.TextMatrix(4, 2) = "N")) And ((Grid_Equipment.TextMatrix(5, 2) = "N")) And ((Grid_Equipment.TextMatrix(6, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    'relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(4, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = S
                'equipamento 2 = S
                'equipamento 3 = N
                'equipamento 4 = N
                'equipamento 4 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And ((Grid_Equipment.TextMatrix(3, 2) = "S")) And ((Grid_Equipment.TextMatrix(4, 2) = "N")) And ((Grid_Equipment.TextMatrix(5, 2) = "N")) And ((Grid_Equipment.TextMatrix(6, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(3, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = S
                'equipamento 2 = S
                'equipamento 3 = S
                'equipamento 4 = N
                'equipamento 4 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And ((Grid_Equipment.TextMatrix(3, 2) = "S")) And ((Grid_Equipment.TextMatrix(4, 2) = "S")) And ((Grid_Equipment.TextMatrix(5, 2) = "N")) And ((Grid_Equipment.TextMatrix(6, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(2, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(3, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Caption = Grid_Equipment.TextMatrix(4, 1)
                    GoTo servicos
                End If
                
                
                'equipamento 1 = N
                'equipamento 2 = S
                'equipamento 3 = S
                'equipamento 4 = N
                'equipamento 4 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "S")) And ((Grid_Equipment.TextMatrix(4, 2) = "S")) And ((Grid_Equipment.TextMatrix(5, 2) = "N")) And ((Grid_Equipment.TextMatrix(6, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(3, 1)
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(4, 1)
                    GoTo servicos
                End If
                
                'equipamento 1 = N
                'equipamento 2 = S
                'equipamento 3 = N
                'equipamento 4 = N
                'equipamento 4 = N
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And ((Grid_Equipment.TextMatrix(3, 2) = "S")) And ((Grid_Equipment.TextMatrix(4, 2) = "N")) And ((Grid_Equipment.TextMatrix(5, 2) = "N")) And ((Grid_Equipment.TextMatrix(6, 2) = "N")) Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                    
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Caption = Grid_Equipment.TextMatrix(3, 1)
                    'relFicha_Assistencia_Tecnica.lblEquipamento2.Caption = Grid_Equipment.TextMatrix(4, 1)
                    GoTo servicos
                End If
                
                
                
                

            End If
            
        End If
        
        

servicos:
        'SERVIÇOS
        
        'Tem somente 1 serviço
        If (Combo_Servico.Text <> "") And (Combo_Servico2.Text = "") Then
            relFicha_Assistencia_Tecnica.lblViewTipoServico2.Visible = False
            relFicha_Assistencia_Tecnica.lblTipoServico2.Visible = False
            relFicha_Assistencia_Tecnica.lblTipoServico1.Caption = Combo_Servico.Text
        End If
        
        'Tem os 2 serviços
        If (Combo_Servico.Text <> "") And (Combo_Servico2.Text <> "") Then
            relFicha_Assistencia_Tecnica.lblTipoServico1.Caption = Combo_Servico.Text
            relFicha_Assistencia_Tecnica.lblTipoServico2.Caption = Combo_Servico2.Text
        End If
        
        If Trim(txtObsPreFinalizacao.Text) <> "" Then
            relFicha_Assistencia_Tecnica.lblObs.Caption = "OBS.: " & txtObsPreFinalizacao.Text
        Else
            relFicha_Assistencia_Tecnica.lblObs.Caption = "OBS.:"
        End If
        
        
        relFicha_Assistencia_Tecnica.lblValorInstalacao.Caption = FormatCurrency(Text_Valor.Text)
        relFicha_Assistencia_Tecnica.Caption = Text_Nome_Cliente.Text
        
ajuste_posicao:

        'Ajuste de posições
        
        '1 serviço
        '1 equipamento
        If (Combo_Servico2.Text = "") And (Form_Service.Grid_Equipment.Rows - 2 = 1) Then
        
            'equipamento1 = S feito
            If (Grid_Equipment.TextMatrix(2, 2) = "S") Then
                relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2350
                relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2350
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2300
                relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2300
                
                
                relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2750
                relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2750
                relFicha_Assistencia_Tecnica.lblObs.Top = 3000
                
                relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3300
                relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3300
                relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3800
                'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                GoTo printficha
            End If
            
            'equipamento1 = N
            If (Grid_Equipment.TextMatrix(2, 2) = "N") Then
                relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = False
                relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = False
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2200
                relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2200

                relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2400
                relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2400
                
                relFicha_Assistencia_Tecnica.lblObs.Top = 3050
                
                relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3300
                relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3300
                relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3800
                'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                GoTo printficha
            End If
            
            
        End If
        
        '1 servico
        '2 equipamentos
        If (Combo_Servico2.Text = "") And (Form_Service.Grid_Equipment.Rows - 2 = 2) Then
            If Combo_Servico2.Text = "" Then
                
                '1o. EQUIPAMENTO = N
                '2o. EQUIPAMENTO = N
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And (Grid_Equipment.TextMatrix(3, 2) = "N") Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2500
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2500
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2780
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2780
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 2800
                    
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3000
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3000
                    
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                    
                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
                
                
                '1o. EQUIPAMENTO = S
                '2o. EQUIPAMENTO = N
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And (Grid_Equipment.TextMatrix(3, 2) = "N") Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2400
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2400
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2650
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2650
                    
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 2900
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3200
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3200
                    
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
                
                
                '1o. EQUIPAMENTO = N
                '2o. EQUIPAMENTO = S
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And (Grid_Equipment.TextMatrix(3, 2) = "S") Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2380
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2380
                    
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2700
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2700
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 3300
                    
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3000
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3000
                    
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3550
                    
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                    
                    
                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
                
                '1o. EQUIPAMENTO = S
                '2o. EQUIPAMENTO = S
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And (Grid_Equipment.TextMatrix(3, 2) = "S") Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2400
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2400
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2630
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2630
                    
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2860
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2860
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 3100
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3400
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3400
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                    
                    

                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
            End If
        End If
        
        
        '1 servico
        '3 ou mais equipamentos
        If (Combo_Servico2.Text = "") And (Form_Service.Grid_Equipment.Rows - 2 > 2) Then
        
                If Form_Service.Grid_Equipment.Rows - 2 = 4 Then
                    GoTo quatro_equipamentos
                End If
        
                If Form_Service.Grid_Equipment.Rows - 2 = 5 Then
                    GoTo cinco_equipamentos
                End If
        
                'SSS
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And (Grid_Equipment.TextMatrix(3, 2) = "S") And (Grid_Equipment.TextMatrix(4, 2) = "S") Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2600
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2600
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 3100
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 3100
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 3400
                    
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3800
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3800
                    
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 4100
                    GoTo printficha
                End If
                
                'NNN
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And (Grid_Equipment.TextMatrix(3, 2) = "N") And (Grid_Equipment.TextMatrix(4, 2) = "N") Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2500
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2500
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2580
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2580
                    relFicha_Assistencia_Tecnica.lblObs.Top = 2800
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3000
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3000
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
                
                'SSN
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And (Grid_Equipment.TextMatrix(3, 2) = "S") And (Grid_Equipment.TextMatrix(4, 2) = "N") Then
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2390
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2390
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2590
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2590
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 3050
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 3050
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 2800
                    
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3280
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3280
                    
                    
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
                
                'SNN
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And (Grid_Equipment.TextMatrix(3, 2) = "N") And (Grid_Equipment.TextMatrix(4, 2) = "N") Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2500
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2500
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2820
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2820
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 3120
                    
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3400
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3400
                    
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
                
                'SNS
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And (Grid_Equipment.TextMatrix(3, 2) = "N") And (Grid_Equipment.TextMatrix(4, 2) = "S") Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2400
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2400
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2650
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2650
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2880
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2880
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 3150
                    
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3450
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3450
                    
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
                
                'NNS
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And (Grid_Equipment.TextMatrix(3, 2) = "N") And (Grid_Equipment.TextMatrix(4, 2) = "S") Then
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2600
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2600
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2350
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2350

                    relFicha_Assistencia_Tecnica.lblObs.Top = 2800
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3000
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3000
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
                
                'NSN
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And (Grid_Equipment.TextMatrix(3, 2) = "S") And (Grid_Equipment.TextMatrix(4, 2) = "N") Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2380
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2380
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2580
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2580
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 2800
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3000
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3000
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
                
                'NSS
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And (Grid_Equipment.TextMatrix(3, 2) = "S") And (Grid_Equipment.TextMatrix(4, 2) = "S") Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2380
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2380
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2600
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2600
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2800
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2800
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 3050
                    
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3350
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3350
                    
                    
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
        End If
        
        '2 serviços
        '3 equipamento
        If qtdServicos = 2 And qtdEquipamentos = 3 Then
        
            'N, N, N
            If (Grid_Equipment.TextMatrix(2, 2) = "N") And (Grid_Equipment.TextMatrix(3, 2) = "N") And (Grid_Equipment.TextMatrix(4, 2) = "N") Then
                relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = False
                relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = False
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                
                relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2750
                relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2750
                relFicha_Assistencia_Tecnica.lblObs.Top = 3000
                
                relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3300
                relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3300
                relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3800
                'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                GoTo printficha
            End If
            
            
            'N, N, S
            If (Grid_Equipment.TextMatrix(2, 2) = "N") And (Grid_Equipment.TextMatrix(3, 2) = "N") And (Grid_Equipment.TextMatrix(4, 2) = "S") Then
                relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = False
                relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = False
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = False
                relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = False
                
                relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2750
                relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2750
                relFicha_Assistencia_Tecnica.lblObs.Top = 3000
                
                relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3300
                relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3300
                relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3800
                'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                GoTo printficha
            End If
            
            'S, S, S
            If (Grid_Equipment.TextMatrix(2, 2) = "S") And (Grid_Equipment.TextMatrix(3, 2) = "S") And (Grid_Equipment.TextMatrix(4, 2) = "S") Then
                relFicha_Assistencia_Tecnica.lblViewEquipamento1.Visible = True
                relFicha_Assistencia_Tecnica.lblEquipamento1.Visible = True
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento2.Visible = True
                relFicha_Assistencia_Tecnica.lblEquipamento2.Visible = True
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento3.Visible = True
                relFicha_Assistencia_Tecnica.lblEquipamento3.Visible = True
                
                relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 3050
                relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 3050
                
                relFicha_Assistencia_Tecnica.lblViewTipoServico2.Top = 3350
                relFicha_Assistencia_Tecnica.lblTipoServico2.Top = 3300
                
                relFicha_Assistencia_Tecnica.lblObs.Top = 3550
                
                relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3800
                relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3800
                
                relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 4100
                'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                GoTo printficha
            End If

            
            
        End If
        
quatro_equipamentos:

        
        
cinco_equipamentos:
        If Form_Service.Grid_Equipment.Rows - 2 = 5 Then
            
            'NNNNN
            If (Grid_Equipment.TextMatrix(2, 2) = "N") And (Grid_Equipment.TextMatrix(3, 2) = "N") And (Grid_Equipment.TextMatrix(4, 2) = "N") And (Grid_Equipment.TextMatrix(5, 2) = "N") And (Grid_Equipment.TextMatrix(6, 2) = "N") Then
                relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2480
                relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2480
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2600
                relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2600
                
                relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2500
                relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2500
                
                relFicha_Assistencia_Tecnica.lblObs.Top = 3050
                
                relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3350
                relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3350
                
                
                relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                GoTo printficha
            End If
            
            'NSNNN
            If (Grid_Equipment.TextMatrix(2, 2) = "N") And (Grid_Equipment.TextMatrix(3, 2) = "S") And (Grid_Equipment.TextMatrix(4, 2) = "N") And (Grid_Equipment.TextMatrix(5, 2) = "N") And (Grid_Equipment.TextMatrix(6, 2) = "N") Then
                'relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2480
                'relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2480
                
                'relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2600
                'relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2600
                
                'relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2500
                'relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2500
                
                'relFicha_Assistencia_Tecnica.lblObs.Top = 2800
                
                relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 2500
                relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 2500
                
                
                'relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                GoTo printficha
            End If
            
            
            
            
            'SSNN
            If (Grid_Equipment.TextMatrix(2, 2) = "S") And (Grid_Equipment.TextMatrix(3, 2) = "S") And (Grid_Equipment.TextMatrix(4, 2) = "N") And (Grid_Equipment.TextMatrix(5, 2) = "N") Then
                relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2480
                relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2480
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2600
                relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2600
                
                relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2800
                relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2800
                
                relFicha_Assistencia_Tecnica.lblObs.Top = 3050
                
                relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3350
                relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3350
                
                
                relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                GoTo printficha
            End If
            
            'SNSN
            If (Grid_Equipment.TextMatrix(2, 2) = "S") And (Grid_Equipment.TextMatrix(3, 2) = "N") And (Grid_Equipment.TextMatrix(4, 2) = "S") And (Grid_Equipment.TextMatrix(5, 2) = "N") Then
                relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2480
                relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2480
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2600
                relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2600
                
                relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2800
                relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2800
                
                relFicha_Assistencia_Tecnica.lblObs.Top = 3050
                
                relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3350
                relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3350
                
                
                relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                GoTo printficha
            End If
            
            'SNNS
            If (Grid_Equipment.TextMatrix(2, 2) = "S") And (Grid_Equipment.TextMatrix(3, 2) = "N") And (Grid_Equipment.TextMatrix(4, 2) = "N") And (Grid_Equipment.TextMatrix(5, 2) = "S") Then
                relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2400
                relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2400
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2600
                relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2600
                
                relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2800
                relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2800
                
                relFicha_Assistencia_Tecnica.lblObs.Top = 3050
                
                relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3350
                relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3350
                
                
                relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                GoTo printficha
            End If
            
            'NSSN
            If (Grid_Equipment.TextMatrix(2, 2) = "N") And (Grid_Equipment.TextMatrix(3, 2) = "S") And (Grid_Equipment.TextMatrix(4, 2) = "S") And (Grid_Equipment.TextMatrix(5, 2) = "N") Then
                relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2480
                relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2480
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2600
                relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2600
                
                relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2800
                relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2800
                
                relFicha_Assistencia_Tecnica.lblObs.Top = 3050
                
                relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3350
                relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3350
                
                
                relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                GoTo printficha
            End If
            
            'NNSS
            If (Grid_Equipment.TextMatrix(2, 2) = "N") And (Grid_Equipment.TextMatrix(3, 2) = "N") And (Grid_Equipment.TextMatrix(4, 2) = "S") And (Grid_Equipment.TextMatrix(5, 2) = "S") Then
                relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2480
                relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2480
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2600
                relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2600
                
                relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2800
                relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2800
                
                relFicha_Assistencia_Tecnica.lblObs.Top = 3050
                
                relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3350
                relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3350
                
                
                relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                GoTo printficha
            End If
            
            'NSNS
            If (Grid_Equipment.TextMatrix(2, 2) = "S") And (Grid_Equipment.TextMatrix(3, 2) = "S") And (Grid_Equipment.TextMatrix(4, 2) = "N") And (Grid_Equipment.TextMatrix(5, 2) = "N") Then
                relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2480
                relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2480
                
                relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2600
                relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2600
                
                relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2800
                relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2800
                
                relFicha_Assistencia_Tecnica.lblObs.Top = 3050
                
                relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3350
                relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3350
                
                
                relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                GoTo printficha
            End If
            
            
            
        End If
        
        
        '2 servicos
        '1 equipamento
        If (Combo_Servico2.Text <> "") And (Form_Service.Grid_Equipment.Rows - 2 = 1) Then
            relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2800
            relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2800
            
            relFicha_Assistencia_Tecnica.lblViewTipoServico2.Top = 3080
            relFicha_Assistencia_Tecnica.lblTipoServico2.Top = 3080
            
            relFicha_Assistencia_Tecnica.lblObs.Top = 3280
            
            relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 4300
            relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 4300
            
            relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3910
            
            'relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 4400
            
            GoTo printficha
        End If
        
        '2 servicos
        '2 equipamentos
        If (qtdServicos = 2) And (qtdEquipamentos = 2) Then
            
                '1o. EQUIPAMENTO = S    feito
                '2o. EQUIPAMENTO = S
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And (Grid_Equipment.TextMatrix(3, 2) = "S") Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2400
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2400
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2630
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2630
                    
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2900
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2900
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico2.Top = 3130
                    relFicha_Assistencia_Tecnica.lblTipoServico2.Top = 3130
                    
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3500
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3500
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 3900
                    
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 4200
                    
                    

                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
                
                '1o. EQUIPAMENTO = S    falta
                '2o. EQUIPAMENTO = N
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And (Grid_Equipment.TextMatrix(3, 2) = "N") Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2400
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2400
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2630
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2630
                    
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2900
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2900
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico2.Top = 3130
                    relFicha_Assistencia_Tecnica.lblTipoServico2.Top = 3130
                    
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 3800
                    
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3850
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3850
                    
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 4200
                    
                    

                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
                
                '1o. EQUIPAMENTO = N    falta
                '2o. EQUIPAMENTO = S
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And (Grid_Equipment.TextMatrix(3, 2) = "S") Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2400
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2400
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2630
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2630
                    
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2900
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2900
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico2.Top = 3130
                    relFicha_Assistencia_Tecnica.lblTipoServico2.Top = 3130
                    
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 3500
                    
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3850
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3850
                    
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 4200
                    
                    

                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
            
            
                '1o. EQUIPAMENTO = N
                '2o. EQUIPAMENTO = N
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And (Grid_Equipment.TextMatrix(3, 2) = "N") Then
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2400
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2400
                    relFicha_Assistencia_Tecnica.lblViewTipoServico2.Top = 2650
                    relFicha_Assistencia_Tecnica.lblTipoServico2.Top = 2650
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 2860
                    
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3100
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3100
                    
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                    
                    

                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
            
            
            
            
            relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2500
            relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2500
            relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2800
            relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2800
            relFicha_Assistencia_Tecnica.lblViewTipoServico2.Top = 3080
            relFicha_Assistencia_Tecnica.lblTipoServico2.Top = 3080
            relFicha_Assistencia_Tecnica.lblObs.Top = 3280
            relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3490
            relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3490
            
            relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3910
            'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
            
        End If
        
        '2 servicos
        '3 equipamentos
        If (Form_Service.Grid_Equipment.Rows - 2 = 3) And (Combo_Servico2.Text <> "") Then
            
                '1o. EQUIPAMENTO = S
                '2o. EQUIPAMENTO = S
                '3o. EQUIPAMENTO = N
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And (Grid_Equipment.TextMatrix(3, 2) = "S") And (Grid_Equipment.TextMatrix(4, 2) = "N") Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2400
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2400
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2630
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2630
                    
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 3100
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 3100
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 2860
                    
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3500
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3500
                    
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                    
                    

                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
            
                '1o. EQUIPAMENTO = S
                '2o. EQUIPAMENTO = N
                '3o. EQUIPAMENTO = S
                If (Grid_Equipment.TextMatrix(2, 2) = "S") And (Grid_Equipment.TextMatrix(3, 2) = "N") And (Grid_Equipment.TextMatrix(4, 2) = "S") Then
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2400
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2400
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2600
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2600
                    
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2800
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2800
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico2.Top = 3030
                    relFicha_Assistencia_Tecnica.lblTipoServico2.Top = 3030
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 3200
                    
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3500
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3500
                    
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                    
                    

                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
                
                '1o. EQUIPAMENTO = N
                '2o. EQUIPAMENTO = S
                '3o. EQUIPAMENTO = S
                If (Grid_Equipment.TextMatrix(2, 2) = "N") And (Grid_Equipment.TextMatrix(3, 2) = "S") And (Grid_Equipment.TextMatrix(4, 2) = "S") Then
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2380
                    relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2380
                    
                    relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2600
                    relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2600
                    
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2800
                    relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2800
                    
                    relFicha_Assistencia_Tecnica.lblViewTipoServico2.Top = 3050
                    relFicha_Assistencia_Tecnica.lblTipoServico2.Top = 3050
                    
                    relFicha_Assistencia_Tecnica.lblObs.Top = 3290
                    
                    relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3500
                    relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3500
                    
                    relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3750
                    
                    

                    'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
                    GoTo printficha
                End If
            
                relFicha_Assistencia_Tecnica.lblViewEquipamento2.Top = 2500
                relFicha_Assistencia_Tecnica.lblEquipamento2.Top = 2500
                relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2800
                relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2800
                relFicha_Assistencia_Tecnica.lblViewTipoServico2.Top = 3080
                relFicha_Assistencia_Tecnica.lblTipoServico2.Top = 3080
                relFicha_Assistencia_Tecnica.lblObs.Top = 3280
                relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3490
                relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3490
                
                relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 3910
                'relFicha_Assistencia_Tecnica.lineBrinde.Visible = True
        End If
        
        
        
printficha:

        'jaca , Função que ajusta a posicao dos serviços e equipamentos mostradas no relatorio
        Call ajustaPosicaoServicoEquipamento

        relFicha_Assistencia_Tecnica.Show (1)
       'Call Ficha(3)
    End If
    
    If Option_Inst_Pisc Then
       Call Ficha(2)
    End If
    
    If Option_Inst_Res Then
       Call Ficha(1)
    End If

    Printer.EndDoc
    
    Option_Assistencia = False
    Option_Inst_Pisc = False
    Option_Inst_Res = False
End Sub

Private Sub Combo_G_C_Click()
    If Not Combo_G_C.enabled Then
        Exit Sub
    End If
    Combo_C.enabled = Combo_G_C.Text = "COBRANÇA"
    Combo_C.BackColor = IIf(Combo_C.enabled, vbYellow, COLOR_BTNFACE)
End Sub

Private Function buscaCodigoTpServico(strTpServico As String) As String

    Dim rs As ADODB.Recordset
    Dim codigoTpServico As String
    
    Set rs = Return_Recordset(Return_SqlString(18, strTpServico))

    If rs.RecordCount > 0 Then
        codigoTpServico = rs.Fields("Val")
        buscaCodigoTpServico = codigoTpServico
    Else
        buscaCodigoTpServico = "0"
    End If
    
    rs.Close
    Set rs = Nothing

End Function




Private Sub Record_equipamentosDaOS(rs As ADODB.Recordset, tipo As String, numeroDoEquipamento As String, numeroColunaEquip_Codigo As Integer)
    
    If tipo = "REMOVE" Then
        If numeroColunaEquip_Codigo = 1 Then
            rs.Fields("Equip1_Codigo").value = 0
        ElseIf numeroColunaEquip_Codigo = 2 Then
            rs.Fields("Equip2_Codigo").value = 0
        ElseIf numeroColunaEquip_Codigo = 3 Then
            rs.Fields("Equip3_Codigo").value = 0
        ElseIf numeroColunaEquip_Codigo = 4 Then
            rs.Fields("Equip4_Codigo").value = 0
        ElseIf numeroColunaEquip_Codigo = 5 Then
            rs.Fields("Equip5_Codigo").value = 0
        ElseIf numeroColunaEquip_Codigo = 6 Then
            rs.Fields("Equip6_Codigo").value = 0
        ElseIf numeroColunaEquip_Codigo = 7 Then
            rs.Fields("Equip7_Codigo").value = 0
        ElseIf numeroColunaEquip_Codigo = 8 Then
            rs.Fields("Equip8_Codigo").value = 0
        ElseIf numeroColunaEquip_Codigo = 9 Then
            rs.Fields("Equip9_Codigo").value = 0
        ElseIf numeroColunaEquip_Codigo = 10 Then
            rs.Fields("Equip10_Codigo").value = 0
        End If
    Else
        If numeroColunaEquip_Codigo = 1 Then
            rs.Fields("Equip1_Codigo").value = numeroDoEquipamento
        ElseIf numeroColunaEquip_Codigo = 2 Then
            rs.Fields("Equip2_Codigo").value = numeroDoEquipamento
        ElseIf numeroColunaEquip_Codigo = 3 Then
            rs.Fields("Equip3_Codigo").value = numeroDoEquipamento
        ElseIf numeroColunaEquip_Codigo = 4 Then
            rs.Fields("Equip4_Codigo").value = numeroDoEquipamento
        ElseIf numeroColunaEquip_Codigo = 5 Then
            rs.Fields("Equip5_Codigo").value = numeroDoEquipamento
        ElseIf numeroColunaEquip_Codigo = 6 Then
            rs.Fields("Equip6_Codigo").value = numeroDoEquipamento
        ElseIf numeroColunaEquip_Codigo = 7 Then
            rs.Fields("Equip7_Codigo").value = numeroDoEquipamento
        ElseIf numeroColunaEquip_Codigo = 8 Then
            rs.Fields("Equip8_Codigo").value = numeroDoEquipamento
        ElseIf numeroColunaEquip_Codigo = 9 Then
            rs.Fields("Equip9_Codigo").value = numeroDoEquipamento
        ElseIf numeroColunaEquip_Codigo = 10 Then
            rs.Fields("Equip10_Codigo").value = numeroDoEquipamento
        End If
    End If

End Sub

Private Sub Record_RemoveEquipamentosDaOS(rs As ADODB.Recordset, numeroDoEquipamento As String, numeroColunaEquip_Codigo As Integer)
    
    If numeroColunaEquip_Codigo = 1 Then
        rs.Fields("Equip1_Codigo").value = numeroDoEquipamento
    ElseIf numeroColunaEquip_Codigo = 2 Then
        rs.Fields("Equip2_Codigo").value = numeroDoEquipamento
    ElseIf numeroColunaEquip_Codigo = 3 Then
        rs.Fields("Equip3_Codigo").value = numeroDoEquipamento
    ElseIf numeroColunaEquip_Codigo = 4 Then
        rs.Fields("Equip4_Codigo").value = numeroDoEquipamento
    ElseIf numeroColunaEquip_Codigo = 5 Then
        rs.Fields("Equip5_Codigo").value = numeroDoEquipamento
    ElseIf numeroColunaEquip_Codigo = 6 Then
        rs.Fields("Equip6_Codigo").value = numeroDoEquipamento
    ElseIf numeroColunaEquip_Codigo = 7 Then
        rs.Fields("Equip7_Codigo").value = numeroDoEquipamento
    ElseIf numeroColunaEquip_Codigo = 8 Then
        rs.Fields("Equip8_Codigo").value = numeroDoEquipamento
    ElseIf numeroColunaEquip_Codigo = 9 Then
        rs.Fields("Equip9_Codigo").value = numeroDoEquipamento
    ElseIf numeroColunaEquip_Codigo = 10 Then
        rs.Fields("Equip10_Codigo").value = numeroDoEquipamento
    End If

End Sub

Private Sub Record_Record(rs As ADODB.Recordset)
    Dim visitRs As ADODB.Recordset
    
    If Trim(txtValorGeral.Text) <> "" Then
        rs.Fields("Valor_Serv").value = rs.Fields("Valor_Serv").value + cvDbl(txtValorGeral.Text)
        panPecas.enabled = False
    End If
    
    If Cmd_Finaliza.Caption <> "&Cancela" And Cmd_Cancela.Caption <> "&Cancela" Then
    
        rs.Fields("Cod_Cliente").value = Text_Cod_Cliente.Text
        rs.Fields("Cod_Tecnico").value = Text_Tecnico_Cod.Text
        rs.Fields("Cod_End").value = Lbl_Address_Cod.Caption
        
        Grid_Equipment_Click
        
        
        If (Lbl_Equip_Cod.Caption = "") Or (Lbl_Equip_Cod.Caption = "Código") Then
            rs.Fields("Cod_Equip").value = Null
        Else
            rs.Fields("Cod_Equip").value = Lbl_Equip_Cod.Caption
        End If
        
        
        rs.Fields("Dt_Atend").value = Date
        rs.Fields("Tipo_Serv").value = UCase(Combo_Servico.Text)
        rs.Fields("Tipo_Serv2").value = UCase(Combo_Servico2.Text)
        rs.Fields("Dt_Atend_Serv").value = vData(Mask_Dt_Serv.Text)
        rs.Fields("Hr_Atend_Serv").value = Set_Field_Val(Text_Hr_Serv)
        rs.Fields("Ficha").value = UCase(Combo_Ficha.Text)
        If Trim(val(Text_Valor.Text)) = "" Then
           rs.Fields("Valor_Serv").value = Null
        Else
           rs.Fields("Valor_Serv").value = cvDbl(Text_Valor.Text)
        End If
        rs.Fields("Status").value = "ABERTO"
        
        rs.Fields("ObsAtendimento").value = txtObsPreFinalizacao.Text
        
        'Gera visita automática após um ano, 01 ano, 1 ano, 365 dias
        'A função buscaCodigoTpServico busca o código do tipo de serviço, correspondente ao texto do Combo_Servico e manda para a função Geretare_Next

        If Generate_Next_Visit(buscaCodigoTpServico(Combo_Servico.Text)) Then
            
            'MsgBox Grid_Service.TextMatrix(Grid_Address.RowSel, 1)
            
            Call Remove_Next_Visits(rs.Fields("Codigo"), Text_Cod_Cliente, Lbl_Address_Cod, Lbl_Equip_Cod, Combo_Servico.Text)
            
            Set visitRs = Return_Recordset(Return_SqlString(10))
            visitRs.AddNew
            visitRs.Fields("Dt_Prox_Visita").value = DateAdd("yyyy", 1, vData(Mask_Dt_Serv.Text))
            visitRs.Fields("Cod_Cliente").value = Text_Cod_Cliente.Text
            visitRs.Fields("Cod_OS").value = Text_OS.Text
            visitRs.Fields("Dt_Criacao").value = Now
            visitRs.Update
                        
            visitRs.Close
            Set visitRs = Nothing
        End If
    Else
        If Cmd_Finaliza.Caption = "&Cancela" Then
            rs.Fields("G_C") = UCase(Combo_G_C.Text)
            If (Combo_C.ListIndex = -1) And (UCase(Combo_G_C.Text) = "") Then
                rs.Fields("G_C") = "" 'jaca
                rs.Fields("Tipo_Cobranca") = " "
            Else
                rs.Fields("G_C") = UCase(Combo_G_C.Text) 'jaca
                If UCase(Combo_C.Text) <> "" Then
                    rs.Fields("Tipo_Cobranca") = UCase(Combo_C.Text)
                Else
                    rs.Fields("Tipo_Cobranca") = " "
                End If
            End If
            rs.Fields("Status").value = "BAIXADO"
        Else
            Call Remove_Service_Next_Visit(rs.Fields("Codigo"))
            rs.Fields("Status").value = "CANCELADO"
        End If
        
        If Text_Obs.Text = "" Then
            rs.Fields("Obs").value = Null
        Else
            rs.Fields("Obs").value = Text_Obs.Text
        End If
        'rs.Fields("Dt_Status").value = Date
        rs.Fields("Dt_Status").value = Date

    End If
End Sub

Sub Remove_Next_Visits(code As Integer, clientCod As String, addressCod As String, equipCod As String, servTyp As String)
    Dim sql As String
    
    sql = "UPDATE Visitas AS v, Atendimento AS a SET v.dt_expiracao = Now(), v.obs = 'Expirada após a inserção da OS " & code & "'"
    sql = sql & " WHERE v.dt_expiracao Is Null AND v.cod_cliente=" & clientCod & " AND a.codigo=v.cod_os AND a.tipo_serv='" & servTyp & "'"
    sql = sql & " AND a.cod_end = " & addressCod
    If equipCod = "" Then
        sql = sql & " AND a.cod_equip IS NULL"
    Else
        sql = sql & " AND a.cod_equip = " & equipCod
    End If
    conn.Execute sql
End Sub

Sub Remove_Service_Next_Visit(os As Integer)
    Dim sql As String

    sql = "UPDATE Visitas AS v SET v.dt_expiracao = Now(), v.obs = 'Expirada após o cancelamento da OS " & os & "'"
    sql = sql & " WHERE v.cod_os=" & os
    
    conn.Execute sql
End Sub

Sub Clean_Objects()
    Text_OS.Text = ""
    Text_Tecnico_Cod.Text = ""
    Combo_Tecnico.ListIndex = -1
    Combo_Servico.ListIndex = -1
    Combo_Servico2.ListIndex = -1
    'Mask_Dt_Serv.Text = "__/__/____"
    Text_Hr_Serv.Text = ""
    Combo_Ficha.ListIndex = -1
    Text_Valor.Text = ""
    txtObsPreFinalizacao.Text = ""
    
    Lbl_Address_Cod.Caption = ""
    Lbl_End.Caption = ""
    Lbl_Bairro.Caption = ""
    Lbl_Cidade.Caption = ""
    Lbl_UF.Caption = ""
    Lbl_Cep.Caption = ""
    Lbl_Ref.Caption = ""
    Lbl_Zona.Caption = ""
    
    Grid_Address.row = 1
    Grid_Address.RowSel = 1
    Grid_Equipment.row = 1
    Grid_Equipment.RowSel = 1
    
    'Lbl_Equip_Cod.Caption = ""
    'Lbl_Equip.Caption = ""
    
    Combo_G_C.ListIndex = -1
    Combo_C.ListIndex = -1
    Text_Obs.Text = ""
End Sub

Private Function Mandatory_Fields_Filled() As Boolean
    Dim dias As Integer

    Mandatory_Fields_Filled = True

    If Combo_Tecnico.enabled Then
        If Trim(Combo_Tecnico.Text) = "" Then
            MsgBox "É necessário selecionar o técnico!", vbOKOnly, "XSol Aquecedores"
            Combo_Tecnico.SetFocus
            Mandatory_Fields_Filled = False
            Exit Function
        End If
    
        If Trim(Combo_Servico.Text) = "" Then
            MsgBox "É necessário selecionar o Serviço!", vbOKOnly, "XSol Aquecedores"
            Combo_Servico.SetFocus
            Mandatory_Fields_Filled = False
            Exit Function
        End If
        
        dias = DateDiff("d", Date, vData(Mask_Dt_Serv))
        If (dias < 0 Or Trim(Mask_Dt_Serv.Text) = "__/__/____" Or Not IsDate(Mask_Dt_Serv.Text)) And Cmd_Inclui.Caption = "&Cancela" Then
            MsgBox "É necessário digitar um dia valido!", vbOKOnly, "XSol Aquecedores"
            Mask_Dt_Serv.SetFocus
            Mandatory_Fields_Filled = False
            Exit Function
        End If
    
        If Trim(Combo_Ficha.Text) = "" Then
            MsgBox "É necessário selecionar a Ficha!", vbOKOnly, "XSol Aquecedores"
            Combo_Ficha.SetFocus
            Mandatory_Fields_Filled = False
            Exit Function
        End If
        
        If Lbl_Address_Cod.Caption = "" Then
            MsgBox "É necessário selecionar um endereço!", vbOKOnly, "XSol Aquecedores"
            Mandatory_Fields_Filled = False
            Exit Function
        End If
        
        'parei aqui
'        If Lbl_Equip_Cod.Caption = "Código" Then
'            MsgBox "É necessário selecionar um e!", vbOKOnly, "XSol Aquecedores"
'            Mandatory_Fields_Filled = False
'            Exit Function
'        End If
        
    Else
        If Cmd_Finaliza.Caption = "&Cancela" Then
            If Trim(Combo_G_C.Text) = "" Then
                MsgBox "É necessário selecionar a Garantia / Cobrança!", vbOKOnly, "XSol Aquecedores"
                Combo_G_C.SetFocus
                Mandatory_Fields_Filled = False
                Exit Function
            End If
            
            If Trim(Combo_G_C.Text) = "COBRANÇA" And Trim(Combo_C.Text) = "" Then
                MsgBox "É necessário selecionar tipo de cobrança!", vbOKOnly, "XSol Aquecedores"
                'Combo_C.SetFocus
                Mandatory_Fields_Filled = False
                Exit Function
            End If
        End If
    End If
End Function

Private Sub Object_Controller(cType As Integer)
        
    
    Call enableObjects(cType)
    
    Cmd_Inclui.enabled = cType <= 3
    Cmd_Edita.enabled = cType = 2 Or cType = 4
    Cmd_Exclui.enabled = cType = 2
    Cmd_Print.enabled = cType = 2
    If cType <> 2 Then
        Cmd_Finaliza.enabled = cType = 5
        Cmd_Cancela.enabled = cType = 7
        Grid_Service.enabled = cType = 1 Or cType = 6
    End If
    
    Lbl_Cob_Garant.Visible = cType <> 7
    Lbl_Cob.Visible = cType <> 7
    Combo_G_C.Visible = cType <> 7
    Combo_C.Visible = cType <> 7
    
    If cType = 1 Then
        Call Combo_Set_Initialize
        Call Clean_Objects
    ElseIf cType = 2 Then  'record selected
    ElseIf cType = 3 Then  'insert a record
        Cmd_Inclui.Caption = "&Cancela"
        Call insEdit_Controller
        Call Clean_Objects
        Call Combo_Initialize(11)
    ElseIf cType = 4 Then  'Edit a record
        Cmd_Edita.Caption = "&Cancela"
        'Combo_G_C.enabled = True
        'Combo_C.enabled = True
        
        Call insEdit_Controller
    ElseIf cType = 5 Then  'Close a service
        Cmd_Finaliza.Caption = "&Cancela"
        
        Combo_G_C.BackColor = vbYellow
        Combo_G_C.SetFocus
        
        Call Combo_Finaliza
    ElseIf cType = 6 Then  'Cancel a insertion or edition or after saving
        Cmd_Inclui.Caption = "&Incluir"
        Cmd_Edita.Caption = "&Editar"
        Cmd_Finaliza.Caption = "&Finalizar"
        Cmd_Cancela.Caption = "&Cancelar"
        
        Combo_Tecnico.BackColor = vbWhite
        Combo_Servico.BackColor = vbWhite
        Combo_Ficha.BackColor = vbWhite
        Mask_Dt_Serv.BackColor = vbWhite
        'Combo_G_C.BackColor = vbWhite
        'Combo_C.BackColor = vbWhite
                
        Combo_C.enabled = False
        txtObsPreFinalizacao.enabled = False
                
        Call Clean_Objects
        Cmd_Inclui.enabled = True
        Call Combo_Set_Initialize
        Call Grid_Service_Click
    ElseIf cType = 7 Then   'Canceling a service
        Cmd_Cancela.Caption = "&Cancela"
        
        Text_Obs.SetFocus
    End If
End Sub

Private Sub enableObjects(cType As Integer)
    Dim enabled As Boolean
    
    enabled = cType = 3 Or cType = 4
    
    Combo_Tecnico.enabled = enabled
    Combo_Servico.enabled = enabled
    Combo_Servico2.enabled = enabled
    Mask_Dt_Serv.enabled = enabled
    Text_Hr_Serv.enabled = enabled
    Combo_Ficha.enabled = enabled
    Text_Valor.enabled = enabled
    
    'Grid_Address.enabled = enabled
    Grid_Equipment.enabled = enabled
    
    enabled = cType = 5 Or cType = 7
    
    Combo_G_C.enabled = enabled
    Text_Obs.Locked = Not enabled
    
    
    Combo_G_C.BackColor = IIf(Combo_G_C.enabled, &H80000005, COLOR_BTNFACE)
    Combo_C.BackColor = IIf(Combo_C.enabled, &H80000005, COLOR_BTNFACE)
    Text_Obs.BackColor = IIf(Not Text_Obs.Locked, &H80000005, COLOR_BTNFACE)
    
    'Text_Obs.enabled = enabled
    
    Cmd_Salva.enabled = (cType >= 3 And cType <= 5) Or cType = 7
End Sub

Private Sub insEdit_Controller()
    Combo_Tecnico.BackColor = vbYellow
    Combo_Servico.BackColor = vbYellow
    Mask_Dt_Serv.BackColor = vbYellow
    
    If (Grid_Service.Rows > 2) And (Cmd_Salva.enabled = False) Then
        Mask_Dt_Serv.Text = Grid_Service.TextMatrix(Grid_Service.RowSel, 1)
    ElseIf Cmd_Salva.enabled = False Then
        Mask_Dt_Serv.Text = "__/__/____"
    End If
    
    
    Combo_Ficha.BackColor = vbYellow
    txtObsPreFinalizacao.enabled = True
    Combo_Tecnico.SetFocus
End Sub

Private Sub Configure_Grid()
    With Grid_Address
        .Cols = 6
        .ColWidth(0) = 630
        .ColWidth(1) = 4800
        .ColWidth(2) = 2940
        .ColWidth(3) = 2940
        .ColWidth(4) = 2940
        .ColWidth(5) = 2940
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Endereço"
        .TextMatrix(0, 2) = "Cidade"
        .TextMatrix(0, 3) = "Fone res."
        .TextMatrix(0, 4) = "Fone com."
        .TextMatrix(0, 5) = "Celular"
        
        .RowHeight(1) = 0
    End With
    
    With Grid_Equipment
        .Cols = 4
        .ColWidth(0) = 630
        .ColWidth(1) = 6200
        .ColWidth(2) = 570
        .ColWidth(3) = 780
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Equipamento"
        .TextMatrix(0, 2) = "Ficha"
        .TextMatrix(0, 3) = "Cód.End."
        .RowHeight(1) = 0
    End With
    
    '8715
    With Grid_Service
    'Equip1_Codigo
        .Cols = 16
        .ColWidth(0) = 690
        .ColWidth(1) = 1200
        .ColWidth(2) = 3900
        .ColWidth(3) = 850
        .ColWidth(4) = 1075
        .ColWidth(5) = 1100
        .ColWidth(6) = 690
        .ColWidth(7) = 690
        .ColWidth(8) = 690
        .ColWidth(9) = 690
        .ColWidth(10) = 690
        .ColWidth(11) = 690
        .ColWidth(12) = 690
        .ColWidth(13) = 690
        .ColWidth(14) = 690
        .ColWidth(15) = 690
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignCenterCenter
        .ColAlignment(10) = flexAlignCenterCenter
        .ColAlignment(11) = flexAlignCenterCenter
        .ColAlignment(12) = flexAlignCenterCenter
        .ColAlignment(13) = flexAlignCenterCenter
        .ColAlignment(14) = flexAlignCenterCenter
        .ColAlignment(15) = flexAlignCenterCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "OS"
        .TextMatrix(0, 1) = "Dt Atendimento"
        .TextMatrix(0, 2) = "Tipo Serv"
        .TextMatrix(0, 3) = "Valor"
        .TextMatrix(0, 4) = "Status"
        .TextMatrix(0, 5) = "Dt Status"
        .TextMatrix(0, 6) = "Equip 1"
        .TextMatrix(0, 7) = "Equip 2"
        .TextMatrix(0, 8) = "Equip 3"
        .TextMatrix(0, 9) = "Equip 4"
        .TextMatrix(0, 10) = "Equip 5"
        .TextMatrix(0, 11) = "Equip 6"
        .TextMatrix(0, 12) = "Equip 7"
        .TextMatrix(0, 13) = "Equip 8"
        .TextMatrix(0, 14) = "Equip 9"
        .TextMatrix(0, 15) = "Equip 10"
        
        .RowHeight(1) = 0
    End With
    
    'pirao de ovo 1
    With gridMaterial
        .Cols = 3
        .ColWidth(0) = 4500
        .ColWidth(1) = 930
        .ColWidth(2) = 630
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Peça"
        .TextMatrix(0, 1) = "Valor"
        .TextMatrix(0, 2) = "Qtd."
        .RowHeight(1) = 0
    End With
End Sub

Private Function Return_Recordset(strSql As String) As ADODB.Recordset
    
    Dim rs As New ADODB.Recordset
    Call Form_Fornecedores.FornecedorArquivos_Dados
    
    Set conn = Nothing
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
    
    'MsgBox strSql
    rs.Open strSql, conn, adOpenStatic, adLockOptimistic
    
    Set Return_Recordset = rs
    
    
End Function



Private Sub Show_Records(rs As ADODB.Recordset, Optional code As Integer)
    Dim selectedRow As Integer
    Dim col As Integer
    Dim bkColor As Long
    Dim qtdRows As Integer
        
    'If global_abrindoForm = True Then
    '    Exit Sub
    'End If
        
    selectedRow = 2
    With Grid_Service
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("codigo") & vbTab & IIf(IsNull(rs.Fields("Dt_Atend_Serv")), "", rs.Fields("Dt_Atend_Serv")) & vbTab & IIf(IsNull(rs.Fields("Tipo_Serv")), "", rs.Fields("Tipo_Serv")) & vbTab & IIf(IsNull(rs.Fields("Valor_Serv")), "", rs.Fields("Valor_Serv")) & vbTab & IIf(IsNull(rs.Fields("Status")), "", rs.Fields("Status"))) & vbTab & rs.Fields("Dt_Status") & vbTab & rs.Fields("Equip1_Codigo") & vbTab & IIf(IsNull(rs.Fields("Equip2_Codigo")), "", rs.Fields("Equip2_Codigo")) & vbTab & IIf(IsNull(rs.Fields("Equip3_Codigo")), "", rs.Fields("Equip3_Codigo")) & vbTab & IIf(IsNull(rs.Fields("Equip4_Codigo")), "", rs.Fields("Equip4_Codigo")) & vbTab & IIf(IsNull(rs.Fields("Equip5_Codigo")), "", rs.Fields("Equip5_Codigo")) & vbTab & IIf(IsNull(rs.Fields("Equip6_Codigo")), "", rs.Fields("Equip6_Codigo")) & vbTab & IIf(IsNull(rs.Fields("Equip7_Codigo")), "", rs.Fields("Equip7_Codigo")) _
            & vbTab & IIf(IsNull(rs.Fields("Equip8_Codigo")), "", rs.Fields("Equip8_Codigo")) & vbTab & IIf(IsNull(rs.Fields("Equip9_Codigo")), "", rs.Fields("Equip9_Codigo")) & vbTab & IIf(IsNull(rs.Fields("Equip10_Codigo")), "", rs.Fields("Equip10_Codigo"))
            
            
            '.AddItem (rs.Fields("Equip1_Codigo"))
            'MsgBox .TextMatrix(22, 6)
            
            If code = rs.Fields("codigo") Then
                selectedRow = .Rows - 1
            End If
            bkColor = GET_BACKCOLOR_BY_STATUS(rs.Fields("Status"))
            If bkColor <> 0 Then
                .row = .Rows - 1
                For col = .FixedCols To .Cols - 1
                    .col = col
                    .CellBackColor = bkColor
                Next
            End If
            rs.MoveNext
        Loop
        
        
        qtdRows = .Rows '
        
        .Refresh
        
        If rs.RecordCount > 0 Then
            .row = qtdRows - 1 'Para posicionar o cursor no último registro do grid
            .RowSel = selectedRow
            .TopRow = 2 'selectedRow
            .col = 0
            .ColSel = .Cols - 1
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Grid_Address_Click()
    With Grid_Address
        If .RowSel = 1 Then
            Exit Sub
        End If
        Call Show_Address_Record_Detail(Return_Recordset(Return_SqlString(7, .TextMatrix(.RowSel, 0))))
    End With
End Sub

Private Sub Grid_Address_DblClick()
    strCodEndereco = Lbl_Address_Cod.Caption
    strEndereco = Lbl_End.Caption
    strBairro = Lbl_Bairro.Caption
    strCidade = Lbl_Cidade.Caption
    strUF = Lbl_UF.Caption
    strCEP = Lbl_Cep.Caption
    frmAtendimentosEndereco.Show (1)
    'Chamar a tela que vai mostrar somente os atendimentos do endereço duplo clicado

End Sub

Private Sub Grid_Address_SelChange()
    'If global_abrindoForm = True Then
    '    Exit Sub
    'End If
    
    Call Grid_Address_Click
    
    If Grid_Address.TextMatrix(Grid_Address.RowSel, 0) <> "" Then
        Call Show_Customer_Equipment(Return_Recordset(Return_SqlString(24, Str(Cod_Cli), Grid_Address.TextMatrix(Grid_Address.RowSel, 0))))
    End If
    
    global_Ficha_telefone1 = Grid_Address.TextMatrix(Grid_Address.RowSel, 3)
    global_Ficha_telefone2 = Grid_Address.TextMatrix(Grid_Address.RowSel, 4)
    global_Ficha_telefone3 = Grid_Address.TextMatrix(Grid_Address.RowSel, 5)
    
End Sub

Sub Show_Address_Record_Detail(rs As ADODB.Recordset)

    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    Lbl_Address_Cod.Caption = rs.Fields("Codigo")
    Lbl_End.Caption = Get_Field_Val(rs.Fields("Endereco"))
    Lbl_Bairro.Caption = Get_Field_Val(rs.Fields("Bairro"))
    Lbl_Cidade.Caption = Get_Field_Val(rs.Fields("Cidade"))
    Lbl_UF.Caption = Get_Field_Val(rs.Fields("UF"))
    Lbl_Cep.Caption = Get_Field_Val(rs.Fields("CEP"))
    Lbl_Ref.Caption = Get_Field_Val(rs.Fields("Ref"))
    Lbl_Zona.Caption = Get_Field_Val(rs.Fields("Zona"))
    
    rs.Close
    Set rs = Nothing
End Sub

Sub Show_Equipment_Record_Detail(rs As ADODB.Recordset)

    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    Lbl_Equip_Cod.Caption = rs.Fields("Codigo")
    'Lbl_Equip.Caption = Get_Field_Val(rs.Fields("Equipamento"))
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Grid_Equipment_Click()
    With Grid_Equipment
        If .RowSel = 1 Then
            Exit Sub
        End If
        Call Show_Equipment_Record_Detail(Return_Recordset(Return_SqlString(9, .TextMatrix(.RowSel, 0))))
    End With
End Sub






Public Sub gravaEquipamentosDaOS(tipo As String, numeroDaOS As String, numeroDoEquipamento As String, numeroColunaEquip_Codigo As Integer)
    Dim rs As ADODB.Recordset
    Set rs = Return_Recordset(Return_SqlString(21, lblOsSelecionada.Caption))
    
    Call Record_equipamentosDaOS(rs, tipo, numeroDoEquipamento, numeroColunaEquip_Codigo)
        
    rs.Update
    
    rs.Close
    Set rs = Nothing
    
End Sub

Public Sub desgravaEquipamentosDaOS(numeroDaOS As String, numeroDoEquipamento As String, numeroColunaEquip_Codigo As Integer)
    Dim rs As ADODB.Recordset
    Set rs = Return_Recordset(Return_SqlString(21, lblOsSelecionada.Caption))
    
    Call Record_RemoveEquipamentosDaOS(rs, numeroDoEquipamento, numeroColunaEquip_Codigo)
        
    rs.Update
    
    rs.Close
    Set rs = Nothing
    
End Sub




Private Sub Grid_Equipment_DblClick()

    If Cmd_Inclui.Caption = "&Cancela" Then
        MsgBox "Você precisa estar em modo de Edição para alterar esta tabela."
        Exit Sub
    End If


    If Grid_Service.Rows <= 2 Then
        MsgBox "Este é o primeiro atendimento do cliente " & Text_Nome_Cliente.Text & "." & Chr(13) & "É necessário salvar este atendimento e posteriormente editar para incluir o(s) equipamento(s)."
        Exit Sub
    End If

    With Grid_Equipment
    
        If .Rows = 2 Then
            Exit Sub
        End If
    

        If .TextMatrix(.RowSel, 2) = "S" Then
            .TextMatrix(.RowSel, 2) = "N"
            gravaEquipamentosDaOS "REMOVE", lblOsSelecionada.Caption, .TextMatrix(.RowSel, 0), .RowSel - 1
        Else
            .TextMatrix(.RowSel, 2) = "S"
            gravaEquipamentosDaOS "INSERE", lblOsSelecionada.Caption, .TextMatrix(.RowSel, 0), .RowSel - 1
        End If
    End With
    
    'atualiza de novo aqui o grid service parei aqui
    'comentei aqui
    'Call Show_Records(Return_Recordset(Return_SqlString(1, Str(Cod_Cli))))
    
    'pintaLinhasEquipamentosDaOs

End Sub

Private Sub Grid_Equipment_SelChange()
    Call Grid_Equipment_Click
End Sub

Private Sub Grid_Service_Click()
    With Grid_Service
        If .RowSel = 1 Then
            Exit Sub
        End If
        
        Call Show_Record_Detail(Return_Recordset(Return_SqlString(2, .TextMatrix(.RowSel, 0))))
    
        If Form_Service.Visible = True Then
            If .Rows > 3 Then
                pintaLinhasEquipamentosDaOs
            End If
        End If
    End With
    
    publicValorAtendimento = Text_Valor.Text
    
    pecasDaOS (Grid_Service.TextMatrix(Grid_Service.RowSel, 0))
    Call contaServicoEquipamento
    
End Sub

Private Sub Grid_Service_SelChange()

    Call Grid_Service_Click

End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String, Optional param3 As String) As String
    Dim whereClause As String
    
    If queryType = 1 Then
        'mostra somente registros de 1 ano
        Return_SqlString = "SELECT * FROM Atendimento WHERE cod_cliente =  " & param1 & " AND dt_atend >= DATEADD('yyyy',-1,Date()) ORDER BY Codigo"
    ElseIf queryType = 2 Then
        Return_SqlString = "SELECT * FROM Atendimento a, Tecnico t WHERE a.Cod_Tecnico = t.Codigo AND a.codigo = " & param1
    ElseIf queryType = 3 Then
        Return_SqlString = "SELECT * FROM Atendimento"
    ElseIf queryType = 4 Then
        Return_SqlString = "SELECT Nome AS Val FROM Tecnico Where Status = 'A' "
    ElseIf queryType = 5 Then
        Return_SqlString = "SELECT Tipo_Serv AS Val FROM Servico Order by 1"
    ElseIf queryType = 6 Then
        Return_SqlString = "SELECT * FROM Endereco WHERE cod_cliente = " & param1 & " ORDER BY Codigo"
    ElseIf queryType = 7 Then
        Return_SqlString = "SELECT * FROM Endereco WHERE Codigo = " & param1
    ElseIf queryType = 8 Then
        Return_SqlString = "SELECT * FROM Equipamento WHERE cod_cliente = " & param1 & " ORDER BY Codigo"
    ElseIf queryType = 9 Then
        Return_SqlString = "SELECT * FROM Equipamento WHERE Codigo = " & param1
    ElseIf queryType = 10 Then
        Return_SqlString = "SELECT * FROM Visitas"
    ElseIf queryType = 11 Then
        Return_SqlString = "SELECT * FROM Tecnico WHERE Nome = '" & param1 & "'"
    ElseIf queryType = 12 Then
        Return_SqlString = "SELECT * FROM Clientes WHERE Codigo = " & param1
    ElseIf queryType = 13 Then
        Return_SqlString = "SELECT * FROM Atendimento WHERE codigo = " & param1
    ElseIf queryType = 14 Then
        Return_SqlString = "SELECT Nome AS Val FROM Tecnico WHERE Status = 'A'"
    ElseIf queryType = 15 Then
        Return_SqlString = "SELECT * FROM Visitas v WHERE v.cod_cliente = " & param1 & " AND "
    ElseIf queryType = 16 Then
        Return_SqlString = "SELECT count(*) FROM Tecnico  WHERE Codigo = " & param1 & " AND (DiasNaoDisponivel LIKE '*" & param2 & "*' OR DiasNaoDisponivel = 'M')"
    ElseIf queryType = 17 Then
        Return_SqlString = "SELECT Tipo_Serv AS Val FROM Servico Order by 1"
    ElseIf queryType = 18 Then
        Return_SqlString = "SELECT Codigo AS Val FROM Servico Where Tipo_Serv = '" & param1 & "'"
    ElseIf queryType = 19 Then
        Return_SqlString = "SELECT Descricao from Garantias order by 1"
    ElseIf queryType = 20 Then
        Return_SqlString = "SELECT Descricao from Cobrancas order by 1"
    ElseIf queryType = 21 Then
        Return_SqlString = "SELECT * FROM Atendimento WHERE codigo = " & param1
    ElseIf queryType = 22 Then
        Return_SqlString = "SELECT * FROM Atendimento_pecas"
    ElseIf queryType = 23 Then
        Return_SqlString = "SELECT * FROM Atendimento_pecas WHERE codigo_atendimento = " & param1 & " and codigo_material = " & param2
    ElseIf queryType = 24 Then
        Return_SqlString = "SELECT * FROM Equipamento WHERE Cod_Cliente = " & param1 & " AND Cod_endereco = '" & param2 & "'"
    ElseIf queryType = 25 Then
        Return_SqlString = "SELECT * FROM Equipamento WHERE Codigo = " & param1
    ElseIf queryType = 26 Then
        'Mostra todos os registros
        Return_SqlString = "SELECT * FROM Atendimento WHERE cod_cliente = " & param1 & " ORDER BY Codigo"
    End If
End Function

Sub despintaLinhasEquipamentosDaOs()
    Dim contador2 As Integer
    Dim X As Integer
    
    For contador2 = 2 To Grid_Equipment.Rows - 1
        Grid_Equipment.row = contador2
        For X = 0 To 3
            Grid_Equipment.col = X
            Grid_Equipment.CellBackColor = vbWhite

            Grid_Equipment.TextMatrix(contador2, 2) = "N"
        Next X
    Next contador2
End Sub

Sub pintaLinhasEquipamentosDaOs()
    
    'para teste
    'Exit Sub
    
    Dim contador1 As Integer
    
    Dim contador3 As Integer
    
    'primeiro pinta as linhas de branco
    despintaLinhasEquipamentosDaOs
    
    For contador1 = 2 To Grid_Equipment.Rows - 1
        If (Grid_Equipment.TextMatrix(contador1, 0) = Grid_Service.TextMatrix(Grid_Service.RowSel, 6)) Or (Grid_Equipment.TextMatrix(contador1, 0) = Grid_Service.TextMatrix(Grid_Service.RowSel, 7)) Or (Grid_Equipment.TextMatrix(contador1, 0) = Grid_Service.TextMatrix(Grid_Service.RowSel, 8)) Or (Grid_Equipment.TextMatrix(contador1, 0) = Grid_Service.TextMatrix(Grid_Service.RowSel, 9)) Or (Grid_Equipment.TextMatrix(contador1, 0) = Grid_Service.TextMatrix(Grid_Service.RowSel, 10)) Or (Grid_Equipment.TextMatrix(contador1, 0) = Grid_Service.TextMatrix(Grid_Service.RowSel, 11)) Or (Grid_Equipment.TextMatrix(contador1, 0) = Grid_Service.TextMatrix(Grid_Service.RowSel, 12)) Or (Grid_Equipment.TextMatrix(contador1, 0) = Grid_Service.TextMatrix(Grid_Service.RowSel, 13)) Or (Grid_Equipment.TextMatrix(contador1, 0) = Grid_Service.TextMatrix(Grid_Service.RowSel, 14)) _
        Or (Grid_Equipment.TextMatrix(contador1, 0) = Grid_Service.TextMatrix(Grid_Service.RowSel, 15)) Then
          
            'MsgBox "achou"
            Grid_Equipment.row = contador1
            
            For contador3 = 0 To 3
                Grid_Equipment.col = contador3
                Grid_Equipment.CellBackColor = &H80FF80
            Next contador3
            
            If Form_Service.Visible Then
                Grid_Equipment.TextMatrix(contador1, 2) = "S"
            End If
        End If
    Next contador1

End Sub

Sub Show_Record_Detail(rs As ADODB.Recordset)

    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    Text_OS.Text = rs.Fields("a.Codigo")
    lblOsSelecionada.Caption = rs.Fields("a.Codigo")
    Text_Status.Text = Get_Field_Val(rs.Fields("a.Status"))
    Cmd_Finaliza.enabled = rs.Fields("a.Status") = "ABERTO"
    Cmd_Cancela.enabled = Cmd_Finaliza.enabled
    
    Text_Tecnico_Cod.Text = rs.Fields("Cod_Tecnico")
    Call Form_Estoque.carregaTecnicos(Form_Estoque.Return_Recordset(Form_Estoque.Return_SqlString(4)))
    Combo_Tecnico.Text = Get_Field_Val(rs.Fields("Nome"))
    Combo_Servico.Text = Get_Field_Val(rs.Fields("Tipo_Serv"))
    
    If rs.Fields("Tipo_Serv2") <> "" Then
        Combo_Servico2.Text = Get_Field_Val(rs.Fields("Tipo_Serv2"))
    Else
        Combo_Servico2.Text = ""
    End If
    
    Mask_Dt_Serv.Text = Get_Field_Val(rs.Fields("Dt_Atend_Serv"))
    Text_Hr_Serv.Text = Get_Field_Val(rs.Fields("Hr_Atend_Serv"))
    If Not IsNull(rs.Fields("Ficha")) Then
        'Combo_Ficha.enabled = True
        Combo_Ficha.Text = rs.Fields("Ficha")
    End If
    Text_Valor.Text = Get_Field_Val(rs.Fields("Valor_Serv"))
    
    If rs.Fields("G_C") = "" Or IsNull(rs.Fields("G_C")) Then
        Combo_G_C.ListIndex = -1
    Else
        'If Combo_G_C.enabled Then
        'If Combo_G_C.enabled Then
            Combo_G_C.enabled = True
            Combo_G_C.Text = rs.Fields("G_C") 'jaca
       ' End If
        
        'End If
    End If
    
    If rs.Fields("Tipo_Cobranca") = " " Or IsNull(rs.Fields("Tipo_Cobranca")) Then
        Combo_C.ListIndex = -1
    Else
        Combo_C.Text = rs.Fields("Tipo_Cobranca")
    End If
    
    Text_Obs.Text = Get_Field_Val(rs.Fields("Obs"))
    

    Call setGridSelRow(Grid_Address, rs.Fields("Cod_End"))
    
    
    Call Show_Address_Record_Detail(Return_Recordset(Return_SqlString(7, rs.Fields("Cod_End"))))
    
    
    'aqui dentro é a merda
    'If Form_Service.Visible = True Then
    '    Call setGridSelRow(Grid_Equipment, IIf(IsNull(rs.Fields("Cod_Equip")), "", rs.Fields("Cod_Equip")))
    'End If
    
    
    If Not IsNull(rs.Fields("Cod_Equip")) Then
        Call Show_Equipment_Record_Detail(Return_Recordset(Return_SqlString(9, rs.Fields("Cod_Equip"))))
    Else
        Lbl_Equip_Cod.Caption = ""
    End If
    
    If (Not IsNull(rs.Fields("ObsAtendimento"))) Or (Trim(rs.Fields("ObsAtendimento")) = "") Then
        txtObsPreFinalizacao.Text = rs.Fields("ObsAtendimento")
    Else
        txtObsPreFinalizacao.Text = ""
    End If
    
    
    rs.Close
    Set rs = Nothing
    
    Call Object_Controller(2)
    

    
End Sub

Private Sub Show_Customer_Address(rs As ADODB.Recordset, Optional code As Integer)
    Dim selectedRow As Integer
    
    selectedRow = 2
    With Grid_Address
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

    'global_Ficha_telefone1 = Grid_Address.TextMatrix(Grid_Address.RowSel, 3)
    'global_Ficha_telefone2 = Grid_Address.TextMatrix(Grid_Address.RowSel, 4)
    'global_Ficha_telefone3 = Grid_Address.TextMatrix(Grid_Address.RowSel, 5)


        Do Until rs.EOF
            .AddItem (rs.Fields("codigo") & vbTab & IIf(IsNull(rs.Fields("Endereco")), "", rs.Fields("Endereco")) & vbTab & IIf(IsNull(rs.Fields("Cidade")), "", rs.Fields("Cidade")) & vbTab & IIf(IsNull(rs.Fields("foneres")), "", rs.Fields("foneres")) & vbTab & IIf(IsNull(rs.Fields("fonecom")), "", rs.Fields("fonecom")) & vbTab & IIf(IsNull(rs.Fields("fonecel")), "", rs.Fields("fonecel")))
            If code = rs.Fields("codigo") Then
                selectedRow = .Rows - 1
            End If
            rs.MoveNext
        Loop
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            '.TopRow = selectedRow
            .col = 0
            .ColSel = .Cols - 1
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Show_Customer_Equipment(rs As ADODB.Recordset, Optional code As Integer)
    Dim selectedRow As Integer
    
    selectedRow = 2
    With Grid_Equipment
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("codigo") & vbTab & IIf(IsNull(rs.Fields("Equipamento")), "", rs.Fields("Equipamento")) & vbTab & IIf(IsNull(rs.Fields("ExibeNaFicha")), "", rs.Fields("ExibeNaFicha")) & vbTab & IIf(IsNull(rs.Fields("Cod_Endereco")), "", rs.Fields("Cod_Endereco")))
            If code = rs.Fields("codigo") Then
                selectedRow = .Rows - 1
            End If
            rs.MoveNext
        Loop
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            '.TopRow = selectedRow
            .col = 0
            .ColSel = .Cols - 1
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Sub Combo_Set_Initialize()
    Call Combo_Initialize(1)
    Call Combo_Initialize(2)
    Call Combo_Initialize(14)
    
    Combo_Ficha.Clear
    Combo_Ficha.AddItem "ASSISTENCIA"
    Combo_Ficha.AddItem "INSTALAÇAO"
    Combo_Ficha.AddItem "VISTORIA"
'
'    Combo_G_C.Clear
'    Combo_G_C.AddItem "GARANTIA SOLETROL"
'    Combo_G_C.AddItem "GARANTIA LUCSOL"
'    Combo_G_C.AddItem "GARANTIA BOSCH"
'    Combo_G_C.AddItem "GARANTIA XSOL"
'    Combo_G_C.AddItem "GARANTIA PRESTADORA"
'    Combo_G_C.AddItem "COBRANÇA"
'
'    Combo_C.Clear
'    Combo_C.AddItem "DINHEIRO"
'    Combo_C.AddItem "BOLETO"
'    Combo_C.AddItem "CHEQUE A VISTA"
'    Combo_C.AddItem "CHEQUE PARCELADO 2 VEZES"
'    Combo_C.AddItem "CHEQUE PARCELADO 3 VEZES"
End Sub

Sub Combo_Finaliza()
'    Combo_G_C.Clear
'    Combo_G_C.AddItem "GARANTIA BOSCH"
'    Combo_G_C.AddItem "GARANTIA XSOL"
'    Combo_G_C.AddItem "GARANTIA PRESTADORA"
'    Combo_G_C.AddItem "COBRANÇA"
End Sub

Sub Combo_Initialize(comboType As Integer)
    Dim combo As ComboBox
    Dim rs As ADODB.Recordset
    
    If comboType = 1 Or comboType = 11 Then
        Set combo = Combo_Tecnico
    ElseIf comboType = 2 Then
        Set combo = Combo_Servico
    ElseIf comboType = 14 Then
        Set combo = Combo_Servico2
    End If
    
    Set rs = Return_Recordset(Return_SqlString(comboType + 3))
    
    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    combo.Clear
    Do While Not rs.EOF
       combo.AddItem rs.Fields("Val")
       rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Sub Ficha(report_Type As Integer)
    Dim Tabela_Atendimento As ADODB.Recordset
    Dim Tabela_Clientes As ADODB.Recordset
    Dim Tabela_Endereco As ADODB.Recordset
    Dim Tabela_Equipamento As ADODB.Recordset
    
    Set Tabela_Atendimento = Return_Recordset(Return_SqlString(2, Text_OS.Text))
    Set Tabela_Clientes = Return_Recordset(Return_SqlString(12, Text_Cod_Cliente.Text))
    Set Tabela_Endereco = Return_Recordset(Return_SqlString(7, Lbl_Address_Cod.Caption))
    If Lbl_Equip_Cod.Caption <> "" Then
        Set Tabela_Equipamento = Return_Recordset(Return_SqlString(9, Lbl_Equip_Cod.Caption))
    End If
    
    Printer.FontName = "Copperplate Gothic Light"
    
    Call Print_Header
    
    Printer.FontSize = 9
    Printer.Print
    Call Print_Title(report_Type)
    Printer.Print
    
    Call Print_OS(Tabela_Atendimento)
    Printer.Print
    
    Call DesligaCorDetaque
    Call Print_Executante(Tabela_Atendimento, Tabela_Clientes)
    Call Print_Cliente(Tabela_Clientes, Tabela_Endereco)
    Call Print_Equipamento(report_Type, Tabela_Equipamento, Tabela_Atendimento)
    If report_Type = 3 Then
        Call Print_Relatorio
    End If
    
    
    Call Pint_Pecas_Trocadas
    
    Call Print_Importante(report_Type)
    
    If report_Type <> 3 Then
        Call Print_Garantia
        Call Print_Lembrete
        Printer.Print
    End If
    
    Call Print_DataAssinatura(report_Type)
    
    Printer.EndDoc
    
    Tabela_Atendimento.Close
    Set Tabela_Atendimento = Nothing
    Tabela_Clientes.Close
    Set Tabela_Clientes = Nothing
    Tabela_Endereco.Close
    Set Tabela_Endereco = Nothing
    If Not Tabela_Equipamento Is Nothing Then
        Tabela_Equipamento.Close
        Set Tabela_Equipamento = Nothing
    End If
End Sub

Sub Print_Header()

    Printer.PaintPicture Image1.Picture, 0, 0
    Printer.Print
    Printer.Print
    Printer.FontBold = True
    Printer.ForeColor = vbBlack
    
    Printer.FontSize = 10
    Printer.Print Spc(40); "FONES:";
    
    Printer.FontSize = 6
    Printer.Print Spc(2);
    Printer.FontSize = 10
    Printer.Print "(11) 2116-6100"
    Printer.FontSize = 9
    Printer.Print Spc(46);
    Printer.FontSize = 10
    Printer.Print "                  (11) 94077-7778"
    
    'Printer.FontSize = 9
    'Printer.Print Spc(79);
    'Printer.FontSize = 10
    'Printer.Print "                  (11) 2301-3651"

    
    
    'Printer.PaintPicture Image2.Picture, 11100, 800
    'whatsapp icone, icone zap, zap, icone zap na ficha
    Printer.PaintPicture Image2.Picture, 7500, 600
    
    Printer.FontBold = False
End Sub

Sub Print_Title(report_Type As Integer)

    Printer.Print
    Printer.FontSize = 12
    Printer.FontBold = True
    If report_Type = 1 Then
        Printer.Print Spc(32); "FICHA DE INSTALAÇÃO"
    ElseIf report_Type = 2 Then
        Printer.Print Spc(27); "FICHA DE INSTALAÇÃO (PISCINA)"
    Else
        Printer.Print Spc(27); "FICHA DE ASSISTÊNCIA TÉCNICA"
    End If
    Printer.FontBold = False

End Sub

Public Sub ligaCorDetaque()
    'ALTERANDO A COR DA FONTE DO DESTAQUE
    If MDIForm_Xsol.fichaDestaqueForecolor = "VERMELHO" Then
        Printer.ForeColor = vbRed
    ElseIf MDIForm_Xsol.fichaDestaqueForecolor = "AZUL" Then
        Printer.ForeColor = vbBlue
    ElseIf MDIForm_Xsol.fichaDestaqueForecolor = "AMARELO" Then
        Printer.ForeColor = vbYellow
    ElseIf MDIForm_Xsol.fichaDestaqueForecolor = "CIANO" Then
        Printer.ForeColor = vbCyan
    End If
    
    If MDIForm_Xsol.fichaDestaqueNegrigo = True Then
        Printer.Font.Bold = True
    Else
        Printer.Font.Bold = False
    End If
    

End Sub

Public Sub DesligaCorDetaque()
    Printer.ForeColor = vbBlack
    Printer.Font.Bold = False
End Sub


Sub Print_OS(Tabela_Atendimento As ADODB.Recordset)
    '******** Variaveis do Cliente ******************************
    
    Dim Print_OS As String
    Dim Print_DtServ As String
    Dim Print_HrServ As String
    
'    Dim Destaque_Forecolor As String
'    Dim Destaque_BackColor As String
'    Dim Destaque_Negrito As Boolean
    
    '*****************************************************************************************************************
    
    Print_OS = IIf(IsNull(Tabela_Atendimento.Fields("a.Codigo")), "", Tabela_Atendimento.Fields("a.Codigo"))
    
    Print_DtServ = IIf(IsNull(Tabela_Atendimento.Fields("Dt_Atend_Serv")), "", Tabela_Atendimento.Fields("Dt_Atend_Serv"))
    
    Print_HrServ = IIf(IsNull(Tabela_Atendimento.Fields("Hr_Atend_Serv")), "", Tabela_Atendimento.Fields("Hr_Atend_Serv"))
    
    '*****************************************************************************************************************
    Printer.FontSize = 12
    
    Printer.FontName = "Copperplate Gothic Light"
    Printer.Print Spc(35);
    
    Call Print_LabelOS("OS")
    Call Print_ValRed(Print_OS, True)
    
    Call Print_Label("DATA DO SERVIÇO")
    Call ligaCorDetaque
    Call Print_Val(Print_DtServ, False)
    Call DesligaCorDetaque
    
    Printer.Print Spc(40);
    Call Print_Label("PERÍODO")
    Call ligaCorDetaque
    Call Print_Val(Print_HrServ, True)
    Call DesligaCorDetaque
    
    Printer.ForeColor = vbBlack
    Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"

End Sub

Sub Print_Executante(Tabela_Atendimento As ADODB.Recordset, Tabela_Clientes As ADODB.Recordset)
'******** Variaveis do Cliente ******************************

Dim Print_EmpExec As String
Dim Print_Nome As String, Print_Atendente As String

'*****************************************************************************************************************

Print_EmpExec = IIf(IsNull(Tabela_Atendimento.Fields("Razao")), "", Tabela_Atendimento.Fields("Razao"))

Print_Nome = IIf(IsNull(Tabela_Atendimento.Fields("Nome")), "", Tabela_Atendimento.Fields("Nome"))

Print_Atendente = IIf(IsNull(Tabela_Clientes.Fields("Atendente")), "", Tabela_Clientes.Fields("Atendente"))

'*****************************************************************************************************************
Printer.FontSize = 9
'Printer.Print

Call Print_Label("EMPRESA EXECUTANTE")
Call Print_ValBlue(Print_EmpExec, True)

Call Print_Label("TÉCNICO RESPONSÁVEL")
Call ligaCorDetaque
Call Print_Val(Print_Nome, True)
Call DesligaCorDetaque

If Form_Customer.Atendente <> "" Then
    Call Print_Label("ATENDENTE")
    Call Print_Val(Print_Atendente, True)
End If

Printer.ForeColor = vbBlack
Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"

End Sub

Sub Print_Cliente(Tabela_Clientes As ADODB.Recordset, Tabela_Endereco As ADODB.Recordset)
    '******** Variaveis do Cliente ******************************
    Dim Print_Codigo As String
    Dim Print_NomeCli As String, Print_EndCli As String
    Dim Print_BairroCli As String, Print_CepCli As String
    Dim Print_Zona As String
    Dim Print_Ref As String
    Dim Print_CidadeCli As String, Print_TelResCli As String
    Dim Print_TelComCli As String, Print_CelularCli As String
    
    '*****************************************************************************************************************
    
    Print_Codigo = IIf(IsNull(Tabela_Clientes.Fields("codigo")), "", Tabela_Clientes.Fields("codigo"))
    
    Print_NomeCli = IIf(IsNull(Tabela_Clientes.Fields("Nome_Cli")), "", Tabela_Clientes.Fields("Nome_Cli"))
    
    Print_EndCli = IIf(IsNull(Tabela_Endereco.Fields("Endereco")), "", Tabela_Endereco.Fields("Endereco"))
    
    Print_BairroCli = IIf(IsNull(Tabela_Endereco.Fields("Bairro")), "", Tabela_Endereco.Fields("Bairro"))
    
    Print_Zona = IIf(IsNull(Tabela_Endereco.Fields("Zona")), "", Tabela_Endereco.Fields("Zona"))
    
    Print_CepCli = IIf(IsNull(Tabela_Endereco.Fields("Cep")), "", Tabela_Endereco.Fields("Cep"))
    
    Print_Ref = IIf(IsNull(Tabela_Endereco.Fields("Ref")), "", Tabela_Endereco.Fields("Ref"))
    
    Print_CidadeCli = IIf(IsNull(Tabela_Endereco.Fields("Cidade")), "", Tabela_Endereco.Fields("Cidade"))
    
    'Telefone, foneres, fonecom, fonecelular
    If Form_Customer.strFoneRes = "" Then
        Print_TelResCli = ""
        Print_TelComCli = ""
        Print_CelularCli = ""
    Else
        Print_TelResCli = Form_Customer.strFoneRes
        Print_TelComCli = Form_Customer.strFoneCom
        Print_CelularCli = Form_Customer.strFoneCel
    End If
    
    '*****************************************************************************************************************
    
    Printer.FontSize = 9
    
    Printer.Print
    Call Print_Label("CÓDIGO DO CLIENTE")
    Call Print_ValBlue(Print_Codigo, True)
    Call Print_Label("NOME DO CLIENTE")
    Call Print_ValBlue(Print_NomeCli, True)
    
    Printer.Print
    
    Call Print_Label("ENDEREÇO")
    Call Print_ValBlue(Print_EndCli, True)
    
    Call Print_Label("BAIRRO")
    Call Print_ValBlue(Print_BairroCli, True)
    
    Call Print_Label("ZONA")
    Call Print_ValBlue(Print_Zona, True)
    
    Call Print_Label("CEP")
    Call Print_ValBlue(Print_CepCli, True)
    
    Call Print_Label("CIDADE")
    Call Print_ValBlue(Print_CidadeCli, True)
    
    Call Print_Label("REFERÊNCIA")
    Call Print_ValBlue(Print_Ref, True)
    'Printer.Print
    
    If Chk_Phone.value = 1 Then
      Call Print_Label("TEL. RES.")
      Call Print_ValBlue(Print_TelResCli, False)
      Printer.Print Spc(5);
      Call Print_Label("TEL. COM.")
      Call Print_ValBlue(Print_TelComCli, False)
      Printer.Print Spc(5);
      Call Print_Label("CELULAR")
      Call Print_ValBlue(Print_CelularCli, True)
    End If
    
    Printer.ForeColor = vbBlack
    Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"

End Sub

Sub Print_Equipamento(report_Type As Integer, Tabela_Equipamento As ADODB.Recordset, Tabela_Atendimento As ADODB.Recordset)

    Dim Print_Equip As String, Print_ValServ As String, Print_TipoServ As String
    Dim contador As Integer
    
    '*****************************************************************************************************************
    
    If Not Tabela_Equipamento Is Nothing Then
        Print_Equip = IIf(IsNull(Tabela_Equipamento.Fields("Equipamento")), "", Tabela_Equipamento.Fields("Equipamento"))
    Else
        Print_Equip = ""
    End If
    Print_ValServ = IIf(IsNull(Tabela_Atendimento.Fields("Valor_Serv")), "", "R$ " + FormatNumber(Tabela_Atendimento.Fields("Valor_Serv"), 2))
    
    Print_TipoServ = IIf(IsNull(Tabela_Atendimento.Fields("Tipo_Serv")), "", Tabela_Atendimento.Fields("Tipo_Serv"))
    
    '*****************************************************************************************************************
    Printer.FontSize = 9
    Printer.Print
    
    If report_Type <> 2 Then
        'Call Print_Label("TIPO DE EQUIPAMENTO")
        'Call Print_Val(Print_Equip, True)
    Else
        Call Print_Label("Nº DE COLETORES")
        Printer.Print
    End If
    
    For contador = 2 To (Grid_Equipment.Rows - 1)
        If Grid_Equipment.TextMatrix(contador, 2) = "S" Then
            Call Print_ValBlack("      EQUIPAMENTO " & Str(contador - 1) & ": ", False)
            Call Print_ValBlue(Grid_Equipment.TextMatrix(contador, 1), True)
        End If
    Next contador
    
    If report_Type = 3 Then
        Call Print_Label("TIPO DE SERVIÇO 1")
        Call Print_ValBlue(Print_TipoServ, True)
        If Combo_Servico2.Text <> "" Then
            Call Print_Label("TIPO DE SERVIÇO 2")
            Call Print_ValBlue(Combo_Servico2.Text, True)
        End If
    End If
    
    
    Call Print_LabelRed("OBS.")
    Call Print_ValRed(Trim(txtObsPreFinalizacao.Text), True)
    
    Call Print_Label("VALOR DA INSTALAÇÃO")
    Call Print_ValBlue(Print_ValServ, True)
    
    Printer.ForeColor = vbBlack
    'Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"

End Sub

Sub Print_Relatorio()
    
    Printer.FontSize = 6
    Printer.Print
    Printer.FontSize = 9
    Printer.FontBold = True
    Printer.Print Spc(3); "RELATÓRIO:"
    Printer.FontBold = False
    Printer.FontSize = 6
    Printer.Print
    Printer.FontSize = 10
    Printer.Print Spc(3); "__________________________________________________________________________________________________________________"
    Printer.FontSize = 8
    Printer.Print
    Printer.FontSize = 10
    Printer.Print Spc(3); "__________________________________________________________________________________________________________________"
    Printer.FontSize = 8
    Printer.Print
    Printer.FontSize = 10
    Printer.Print Spc(3); "__________________________________________________________________________________________________________________"
    Printer.FontSize = 8
    Printer.Print
    Printer.FontSize = 10
    Printer.Print Spc(3); "__________________________________________________________________________________________________________________"
    Printer.FontSize = 8
    Printer.Print
    Printer.FontSize = 10
    Printer.Print Spc(3); "__________________________________________________________________________________________________________________"
    Printer.FontSize = 8
    Printer.Print
    Printer.FontSize = 10
    Printer.Print Spc(3); "__________________________________________________________________________________________________________________"
    Printer.FontSize = 8
    Printer.Print
    Printer.FontSize = 10
    Printer.Print Spc(3); "__________________________________________________________________________________________________________________"
    Printer.FontSize = 8
    Printer.Print
    Printer.FontSize = 10
    Printer.Print Spc(3); "__________________________________________________________________________________________________________________"
    Printer.FontSize = 8
    Printer.Print
    Printer.FontSize = 10
    Printer.Print Spc(3); "__________________________________________________________________________________________________________________"
    Printer.FontSize = 8
    Printer.Print
    Printer.FontSize = 10
    Printer.Print Spc(3); "__________________________________________________________________________________________________________________"
    Printer.FontSize = 8
    Printer.Print
    Printer.FontSize = 10
    Printer.Print Spc(3); "__________________________________________________________________________________________________________________"
    Printer.FontSize = 8
    Printer.Print
    Printer.FontSize = 10
    Printer.Print Spc(3); "__________________________________________________________________________________________________________________"
    'Printer.FontSize = 8
    'Printer.Print
    'Printer.FontSize = 10
    'Printer.Print Spc(3); "__________________________________________________________________________________________________________________"
    'Printer.FontSize = 8
    'Printer.Print
    'Printer.FontSize = 10
    'Printer.Print Spc(3); "__________________________________________________________________________________________________________________"
End Sub

Sub Print_Importante(report_Type As Integer)
    
    Printer.FontBold = True
    Printer.Print
    
    If report_Type <> 3 Then
        Printer.ForeColor = vbRed
        Printer.Print Spc(3); "IMPORTANTE:"
        Printer.FontBold = False
        Printer.ForeColor = RGB(170, 60, 57)
    Else
        Printer.ForeColor = vbRed
        Printer.Print Spc(3); "IMPORTANTE:";
        Printer.FontBold = False
    End If
    
    Printer.ForeColor = vbBlack
    
    If report_Type = 1 Then
        Printer.Print Spc(3); "MANTER A VÁLVULA ANTI-CONGELAMENTO (DISPOSITIVO QUE PROTEGE OS COLETORES CONTRA"
        Printer.Print Spc(3); "CONGELAMENTO EM GEADAS), SEMPRE ENERGIZADA. CASO NO ATO DA INSTALAÇÃO O TÉCNICO NÃO "
        Printer.Print Spc(3); "ENCONTRE A FIAÇÃO DE ALIMENTAÇÃO DA MESMA (ITEM PEDIDO EM VISTORIA TÉCNICA), O "
        Printer.Print Spc(3); "CLIENTE FICA CIENTE QUE O RETORNO SERÁ COBRADO UMA TAXA REFERENTE A UMA VISITA "
        Printer.Print Spc(3); "TÉCNICA. O MESMO SE APLICA COM OS DEMAIS ITENS SOLICITADOS NA VISTORIA."
    ElseIf report_Type = 2 Then
        Printer.Print Spc(3); "SEMPRE MANTER O SISTEMA LIGADO JUNTAMENTE COM A CAPA TÉRMICA, PARA QUE O MESMO "
        Printer.Print Spc(3); "ATINJA A TEMPERATURA DE CONTRATO 25º à 29 GRAUS."
        Printer.Print Spc(3); "QUALQUER TIPO DE MANUTENÇÃO NO SISTEMA DE AQUECIMENTO SOLAR DA PISCINA, É  "
        Printer.Print Spc(3); "ACONSELHÁVEL SER EFETUADO APENAS POR PROFISSIONAIS QUALIFICADO DA XSOL."
    Else
        Printer.Print ; "OS SERVIÇOS EXECUTADOS PODERÃO SER RECLAMADOS NO PRAZO DE "
        Printer.Print Spc(3); "03 MESES A CONTAR DA DATA DE SUA REALIZAÇÃO."
        Printer.Print Spc(3); "APÓS ESTE PRAZO ENTENDE-SE QUE O PROBLEMA FOI SANADO E NOVAS SOLICITAÇÕES SERÃO"
        Printer.Print Spc(3); "COBRADAS."
        Printer.Print
        Printer.Print
    End If
    
    Printer.ForeColor = vbBlack
    'Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
End Sub

Sub Pint_Pecas_Trocadas()
    
    Printer.FontBold = True
    Printer.Print
    
    Printer.ForeColor = vbRed
    Printer.Print Spc(3); "PEÇAS SUBSTITUÍDAS"
    Printer.FontSize = 8
    Printer.Print

    Printer.FontBold = False
    Printer.ForeColor = RGB(170, 60, 57)

    Printer.ForeColor = vbBlack
    
    Printer.FontSize = 10
    Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
    Printer.FontSize = 8
    Printer.Print
    
    Printer.FontSize = 10
    Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
    Printer.FontSize = 8
    Printer.Print
    
    Printer.FontSize = 10
    Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
    Printer.FontSize = 8
    Printer.Print
    
    Printer.FontSize = 10
    Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
    Printer.FontSize = 8
    Printer.Print
    
    Printer.FontSize = 10
    Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
    Printer.FontSize = 8
    Printer.Print
    
    
End Sub

Sub Print_Garantia()

    Printer.FontBold = True
    Printer.ForeColor = vbBlack
    Printer.Print
    Printer.Print Spc(3); "GARANTIA DE INSTALAÇÃO"
    Printer.FontBold = False
    Printer.Print Spc(3); "A INSTALAÇÃO POSSUI 3 MESES DE GARANTIA A PARTIR DA DATA DE CONCLUSÃO DA MESMA, E "
    Printer.Print Spc(3); "COBRE: "
    Printer.Print Spc(3); "* O NÃO FUNCIONAMENTO DO SISTEMA, DECORRENTES A MOTIVOS DE INSTALAÇÃO."
    Printer.Print Spc(3); "* VAZAMENTO NOS CASOS DE INTERLIGAÇÃO DO SISTEMA."
    Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"

End Sub

Sub Print_Lembrete()

    Printer.FontBold = True
    Printer.ForeColor = vbBlack
    Printer.Print
    Printer.FontBold = True
    Printer.Print Spc(3); "LEMBRETE:"
    Printer.FontBold = False
    Printer.Print Spc(3); "COM O OBJETIVO DE SEMPRE MANTER A BOA PERFORMANCE DO SISTEMA, EXISTE A NECESSIDADE "
    Printer.Print Spc(3); "DE NO MÁXIMO 6 MESES SER REALIZADA UMA MANUTENÇÃO PREVENTIVA."
    Printer.Print
    Printer.FontBold = True
    Printer.Print Spc(3); "OBSERVAÇÕES:"
    Printer.Print
    Printer.FontBold = False
    Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
    Printer.Print
    Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
    Printer.Print
    Printer.Print Spc(3); "______________________________________________________________________________________________________________________________"
End Sub

Sub Print_DataAssinatura(report_Type As Integer)

    If report_Type <> 3 Then
        Printer.Print
    End If
    
    Printer.Print Spc(3); "________________________________________________________ , _________ DE _______________________ DE ______________"
    
    If report_Type <> 3 Then
        Printer.Print
        Printer.Print
    End If
    Printer.Print
    Printer.Print
    Printer.Print Spc(3); "__________________________________-__________________________"; Spc(7); "___________________________________"
    Printer.Print Spc(3); "  CLIENTE OU RESPONSÁVEL    -                   RG   "; Spc(24); "TÉCNICO RESPONSÁVEL"

End Sub

Sub Print_Label(Lbl As String)
    Printer.Font.Size = 9
    Printer.FontBold = True
    Printer.ForeColor = vbBlack
    Printer.Print Spc(3); Lbl + ": ";
End Sub

Sub Print_LabelBlue(Lbl As String)
    Printer.Font.Size = 9
    Printer.FontBold = True
    Printer.ForeColor = vbBlue
    Printer.Print Spc(3); Lbl + ": ";
End Sub

Sub Print_LabelOS(Lbl As String)
    Printer.Font.Size = 12
    Printer.FontBold = True
    Printer.ForeColor = vbRed
    Printer.Print Spc(3); Lbl + ": ";
End Sub

Sub Print_LabelRed(Lbl As String)

    Printer.FontBold = True
    Printer.ForeColor = vbRed
    Printer.Print Spc(3); Lbl + ": ";

End Sub

Sub Print_ValBlack(val As String, new_Line As Boolean)
    Printer.Font.Size = 9
    Printer.FontBold = False
    Printer.ForeColor = vbBlack
    Printer.Print val;
    If new_Line Then
        Printer.Print
    End If
End Sub

Sub Print_Val(val As String, new_Line As Boolean)
    Printer.Font.Size = 9
    
    If MDIForm_Xsol.fichaDestaqueNegrigo = False Then
        Printer.FontBold = False
    Else
        Printer.Font.Bold = True
    End If
    
    If (Printer.ForeColor <> 255) And (Printer.ForeColor <> 0) And (Printer.ForeColor <> 65535) And (Printer.ForeColor <> 16776960) Then
        Printer.ForeColor = RGB(0, 0, 255)
    End If
    
'    Printer.FillColor = vbYellow
    Printer.Print val;
    If new_Line Then
        Printer.Print
    End If
End Sub

Sub Print_ValBlue(val As String, new_Line As Boolean)
    Printer.Font.Size = 9
    
    If MDIForm_Xsol.fichaDestaqueNegrigo = False Then
        Printer.FontBold = False
    Else
        Printer.Font.Bold = True
    End If
    
    
        Printer.ForeColor = vbBlue
    
    
'    Printer.FillColor = vbYellow
    Printer.Print val;
    If new_Line Then
        Printer.Print
    End If
End Sub


Sub Print_ValRed(val As String, new_Line As Boolean)
    Printer.FontBold = True
    Printer.ForeColor = vbRed
    Printer.Print val;
    If new_Line Then
        Printer.Print
    End If
End Sub

Private Function RetornaPecas() As Recordset
    Set RetornaPecas = Return_Recordset("SELECT Codigo & '   -' & Descricao as Peca From Material")
End Function

Private Function RetornaRecordset(strQuery As String) As Recordset
    Set RetornaRecordset = Return_Recordset(strQuery)
End Function

Private Sub CarregaPreco(strCodMaterial)
    Dim rsPreco As Recordset
    Set rsPreco = RetornaRecordset("Select ValorVenda from materialfornecedor Where CodMaterial = " & strCodMaterial)
    
    txtValorPeca.Text = rsPreco("ValorVenda")
    
    
End Sub


Private Sub Option_Assistencia_Click()
    Chk_PDF.enabled = False
End Sub

Private Sub Option_Inst_Pisc_Click()
    Chk_PDF.enabled = True
End Sub

Private Sub Option_Inst_Res_Click()
    Chk_PDF.enabled = True
End Sub

Private Sub txtObsPreFinalizacao_KeyPress(KeyAscii As Integer)
    'MsgBox KeyAscii
    If KeyAscii = 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    
End Sub

Private Sub txtQtdPecas_LostFocus()
On Error GoTo erro
    
    If (Trim(txtQtdPecas.Text) <> "") And (Trim(txtQtdPecas.Text) > 0) And (Trim(txtValorPeca.Text) <> "") Then
        txtTotalValorPecas.Text = txtQtdPecas.Text * txtValorPeca.Text
        Exit Sub
    Else
        Exit Sub
    End If
   
erro:
   MsgBox ("Digite um valor inteiro válido!")
   txtQtdPecas.Text = ""
   txtValorPeca.Text = ""
   txtTotalValorPecas.Text = ""

End Sub

Private Sub calculaValorPecas()
On Error GoTo erro
    
    If (Trim(txtQtdPecas.Text) <> "") And (Trim(txtQtdPecas.Text) > 0) And (Trim(txtValorPeca.Text) <> "") Then
        txtTotalValorPecas.Text = txtQtdPecas.Text * txtValorPeca.Text
        Exit Sub
    Else
        txtTotalValorPecas.Text = ""
        Exit Sub
    End If
   
erro:
   MsgBox ("Digite um valor inteiro válido!")
   txtQtdPecas.Text = ""
   txtValorPeca.Text = ""
   txtTotalValorPecas.Text = ""

End Sub

Private Sub pecasDaOS(strOS As String)
        
        Dim strSql As String
        
        Dim rs_pecasOS As Recordset
        Set rs_pecasOS = RetornaPecas()
        
        Dim contador As Integer
        Dim qtdPecasTotal As Integer
        Dim valorTotal As Integer
        
        strSql = "Select tb1.CODIGO_MATERIAL, tb1.QTD, tb1.VALOR , tb2.descricao from atendimento_pecas as tb1 Inner Join Material as tb2 on tb2.Codigo = tb1.codigo_material where tb1.codigo_atendimento = " & strOS
        
        Set rs_pecasOS = Return_Recordset(strSql)
        qtdPecasTotal = rs_pecasOS.RecordCount
        
        
        valorTotal = 0
        
        With gridMaterial
            .Rows = 2
            
            For contador = 1 To qtdPecasTotal
                .AddItem (rs_pecasOS("CODIGO_MATERIAL") & "-" & rs_pecasOS("DESCRICAO") & vbTab & rs_pecasOS("VALOR") & vbTab & rs_pecasOS("QTD"))
                valorTotal = valorTotal + (rs_pecasOS("QTD") * rs_pecasOS("VALOR"))
                rs_pecasOS.MoveNext
            Next contador
            
            .Refresh
        End With
        
        txtValorGeral.Text = valorTotal
        txtValorGeral.Text = FormatCurrency(txtValorGeral.Text)
        
    
End Sub

Private Sub removePeca(codAtendimento As Integer, codMaterial As Integer)
    Dim rsRemovePeca As Recordset
    Dim contador As Integer
    
    Set rsRemovePeca = Return_Recordset(Return_SqlString(23, CDbl(codAtendimento), CDbl(codMaterial)))
    
    If rsRemovePeca.RecordCount > 0 Then
        rsRemovePeca.Delete adAffectCurrent
        rsRemovePeca.Update
    End If
    
    rsRemovePeca.Close
    Set rsRemovePeca = Nothing
End Sub

Private Sub atualizaEquipamento(codEquipamento As String, exibeNaFicha As String)
    Dim rsAtu As ADODB.Recordset
    
    Set rsAtu = Return_Recordset(Return_SqlString(25, codEquipamento))

    rsAtu.Fields("ExibeNaFicha") = exibeNaFicha
    
    rsAtu.Update

    rsAtu.Close
    Set rsAtu = Nothing
    
    
End Sub

Private Sub ajustaPosicaoServicoEquipamento()
    If qtdEquipamentos = 1 Then
        relFicha_Assistencia_Tecnica.lblViewTipoServico1.Top = 2800
        relFicha_Assistencia_Tecnica.lblTipoServico1.Top = 2800
        
        relFicha_Assistencia_Tecnica.lblObs.Top = 3600
        
        relFicha_Assistencia_Tecnica.lblViewValorInstalacao.Top = 3330
        relFicha_Assistencia_Tecnica.lblValorInstalacao.Top = 3330
        
        relFicha_Assistencia_Tecnica.lblViewEquipamento1.Top = 2400
        relFicha_Assistencia_Tecnica.lblEquipamento1.Top = 2400
        
        relFicha_Assistencia_Tecnica.lblViewRelatorio.Top = 4200
    End If
    
    
    'relFicha_Assistencia_Tecnica.lblEquipamento1
End Sub

Private Sub contaServicoEquipamento()
    Dim contador As Integer
     
    'Contando a quantidade de serviços
    If (Combo_Servico.Text = "") And (Combo_Servico2.Text = "") Then
        qtdServicos = 0
    End If
    
    If (Combo_Servico.Text <> "") And (Combo_Servico2.Text = "") Then
        qtdServicos = 1
    End If
    
    If (Combo_Servico.Text <> "") And (Combo_Servico2.Text <> "") Then
        qtdServicos = 2
    End If
    
    'Contando a quantidade de equipamentos
    qtdEquipamentos = 0
    For contador = 2 To Grid_Equipment.Rows - 1
        If (Grid_Equipment.TextMatrix(contador, 2) = "S") Then
                qtdEquipamentos = qtdEquipamentos + 1
        End If
    Next contador
        
    
End Sub
