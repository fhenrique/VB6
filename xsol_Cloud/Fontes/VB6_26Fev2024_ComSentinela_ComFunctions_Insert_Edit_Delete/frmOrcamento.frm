VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrcamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "xSol Aquecedores"
   ClientHeight    =   10455
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   20310
   ControlBox      =   0   'False
   Icon            =   "frmOrcamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10455
   ScaleWidth      =   20310
   Begin VB.Frame Frame1 
      Height          =   10485
      Left            =   -15
      TabIndex        =   0
      Top             =   -75
      Width           =   20280
      Begin VB.Frame frameStatus 
         Caption         =   "STATUS"
         Height          =   630
         Left            =   15045
         TabIndex        =   95
         Top             =   7815
         Width           =   5115
         Begin VB.TextBox txtStatusMensagem 
            Height          =   315
            Left            =   1590
            MaxLength       =   50
            TabIndex        =   97
            Top             =   240
            Width           =   3405
         End
         Begin VB.ComboBox cboStatus 
            Height          =   315
            ItemData        =   "frmOrcamento.frx":08CA
            Left            =   90
            List            =   "frmOrcamento.frx":08D7
            Style           =   2  'Dropdown List
            TabIndex        =   96
            Top             =   240
            Width           =   1410
         End
      End
      Begin VB.Frame panClone 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   19230
         TabIndex        =   91
         Top             =   7350
         Width           =   840
         Begin VB.CommandButton btnClonar 
            Caption         =   "Clonar"
            Height          =   270
            Left            =   60
            TabIndex        =   92
            Top             =   45
            Width           =   735
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Desconto"
         Height          =   615
         Left            =   13680
         TabIndex        =   89
         Top             =   7815
         Width           =   1245
         Begin VB.CommandButton btnDesconto 
            Caption         =   "..."
            Height          =   330
            Left            =   825
            TabIndex        =   99
            Top             =   195
            Width           =   375
         End
         Begin VB.TextBox txtDesconto 
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   90
            Text            =   "855,00"
            Top             =   210
            Width           =   690
         End
      End
      Begin VB.TextBox lblNumeroOrcamento 
         Height          =   285
         Left            =   3045
         TabIndex        =   88
         Top             =   705
         Width           =   1125
      End
      Begin VB.Frame Frame12 
         Caption         =   "Pesquisar orçamentos"
         Height          =   1800
         Left            =   13665
         TabIndex        =   67
         Top             =   8595
         Width           =   6495
         Begin VB.CommandButton btnRestaurar 
            Caption         =   "Atualizar"
            Height          =   315
            Left            =   3870
            TabIndex        =   101
            Top             =   1320
            Width           =   885
         End
         Begin VB.CommandButton btnPesquisar 
            Caption         =   "Pesquisar"
            Height          =   315
            Left            =   5220
            TabIndex        =   77
            Top             =   1320
            Width           =   1005
         End
         Begin VB.ComboBox cboClientesProspects 
            Height          =   315
            Left            =   1065
            TabIndex        =   76
            Top             =   765
            Width           =   5190
         End
         Begin VB.OptionButton optnProspect 
            Caption         =   "Prospect"
            Height          =   195
            Left            =   180
            TabIndex        =   75
            Top             =   1365
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.OptionButton optnCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   120
            TabIndex        =   74
            Top             =   900
            Width           =   900
         End
         Begin VB.TextBox txtCodigo 
            Height          =   300
            Left            =   780
            MaxLength       =   50
            TabIndex        =   73
            Top             =   300
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker dtpDataDe 
            Height          =   270
            Left            =   1680
            TabIndex        =   70
            Top             =   1155
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   476
            _Version        =   393216
            Format          =   161218561
            CurrentDate     =   44349
         End
         Begin MSComCtl2.DTPicker dtpDataAte 
            Height          =   270
            Left            =   1680
            TabIndex        =   71
            Top             =   1440
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            Format          =   161218561
            CurrentDate     =   44349
         End
         Begin VB.Label Label21 
            Caption         =   "Abrir o panel, tem mais coisa em baixo"
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   2025
            TabIndex        =   100
            Top             =   300
            Visible         =   0   'False
            Width           =   2610
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   165
            TabIndex        =   72
            Top             =   390
            Width           =   540
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Até"
            Height          =   195
            Left            =   1320
            TabIndex        =   69
            Top             =   1515
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   1200
            TabIndex        =   68
            Top             =   1215
            Visible         =   0   'False
            Width           =   390
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Imagens"
         Height          =   1830
         Left            =   6735
         TabIndex        =   64
         Top             =   7560
         Width           =   6795
         Begin MSComDlg.CommonDialog CommonDialog2 
            Left            =   3240
            Top             =   1020
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DialogTitle     =   "Imagem para relatório"
         End
         Begin VB.TextBox txtTextoImg2 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Left            =   870
            MaxLength       =   500
            TabIndex        =   85
            Top             =   1440
            Width           =   5805
         End
         Begin VB.CommandButton btnImagem2 
            Caption         =   "..."
            Height          =   300
            Left            =   6285
            TabIndex        =   84
            Top             =   1035
            Width           =   375
         End
         Begin VB.TextBox txtTextoImg1 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   870
            MaxLength       =   500
            TabIndex        =   82
            Top             =   630
            Width           =   5805
         End
         Begin VB.TextBox txtImagem2 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Left            =   870
            Locked          =   -1  'True
            MaxLength       =   500
            TabIndex        =   78
            Top             =   1050
            Width           =   5400
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   3285
            Top             =   180
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DialogTitle     =   "Imagem para relatório"
         End
         Begin VB.CommandButton btnImagem 
            Caption         =   "..."
            Height          =   300
            Left            =   6285
            TabIndex        =   66
            Top             =   270
            Width           =   375
         End
         Begin VB.TextBox txtImagem 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   870
            Locked          =   -1  'True
            MaxLength       =   500
            TabIndex        =   65
            Top             =   255
            Width           =   5400
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Texto:"
            Height          =   195
            Left            =   315
            TabIndex        =   86
            Top             =   1500
            Width           =   450
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Texto:"
            Height          =   195
            Left            =   360
            TabIndex        =   83
            Top             =   660
            Width           =   450
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Imagem 2:"
            Height          =   195
            Left            =   75
            TabIndex        =   81
            Top             =   1110
            Width           =   735
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Imagem 1:"
            Height          =   195
            Left            =   105
            TabIndex        =   79
            Top             =   285
            Width           =   735
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Teste"
         Height          =   375
         Left            =   8280
         TabIndex        =   60
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton optnModelo2 
         Caption         =   "Imprimir orçamento"
         Height          =   195
         Left            =   4170
         TabIndex        =   57
         Top             =   10020
         Width           =   1680
      End
      Begin VB.OptionButton optnModelo1 
         Caption         =   "Imprimir proposta"
         Height          =   195
         Left            =   4170
         TabIndex        =   56
         Top             =   9705
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.Frame Frame10 
         Caption         =   "Observações:"
         Height          =   1830
         Left            =   165
         TabIndex        =   52
         Top             =   7560
         Width           =   6495
         Begin VB.CommandButton btnEditarObs 
            Caption         =   "Editar"
            Height          =   300
            Left            =   5550
            TabIndex        =   98
            Top             =   1050
            Width           =   855
         End
         Begin VB.CommandButton btnRemoverObservacao 
            Caption         =   "Remover"
            Height          =   300
            Left            =   5550
            TabIndex        =   54
            Top             =   675
            Width           =   855
         End
         Begin VB.CommandButton btnAdicionarObservacao 
            Caption         =   "Adicionar"
            Height          =   300
            Left            =   5550
            TabIndex        =   53
            Top             =   285
            Width           =   870
         End
         Begin MSFlexGridLib.MSFlexGrid gridObservacoes 
            Height          =   1515
            Left            =   105
            TabIndex        =   55
            Top             =   255
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   2672
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            SelectionMode   =   1
         End
      End
      Begin VB.TextBox txtValidade 
         Height          =   315
         Left            =   5310
         MaxLength       =   50
         TabIndex        =   51
         Top             =   690
         Width           =   2640
      End
      Begin VB.Frame Frame9 
         Caption         =   "Forma(s) de pagamento:"
         Height          =   1920
         Left            =   6375
         TabIndex        =   46
         Top             =   5625
         Width           =   6045
         Begin VB.CommandButton btnEditarFormaPagamento 
            Caption         =   "Editar"
            Height          =   300
            Left            =   3240
            TabIndex        =   94
            Top             =   1530
            Width           =   960
         End
         Begin VB.CommandButton btnRemoverFormaPagamento 
            Caption         =   "Remover forma"
            Height          =   330
            Left            =   1755
            TabIndex        =   48
            Top             =   1500
            Width           =   1320
         End
         Begin VB.CommandButton btnAdicionarFormaPagamento 
            Caption         =   "Adicionar forma"
            Height          =   330
            Left            =   120
            TabIndex        =   47
            Top             =   1500
            Width           =   1440
         End
         Begin MSFlexGridLib.MSFlexGrid gridFormasPagamento 
            Height          =   1215
            Left            =   105
            TabIndex        =   49
            Top             =   270
            Width           =   5820
            _ExtentX        =   10266
            _ExtentY        =   2143
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            SelectionMode   =   1
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Garantia(s):"
         Height          =   1890
         Left            =   165
         TabIndex        =   42
         Top             =   5640
         Width           =   6015
         Begin VB.CommandButton btnEditarGarantia 
            Caption         =   "Editar"
            Height          =   300
            Left            =   2520
            TabIndex        =   93
            Top             =   1500
            Width           =   960
         End
         Begin VB.CommandButton btnAdicionarGarantia 
            Caption         =   "Adicionar"
            Height          =   300
            Left            =   90
            TabIndex        =   44
            Top             =   1500
            Width           =   1050
         End
         Begin VB.CommandButton btnRemoverGarantia 
            Caption         =   "Remover"
            Height          =   300
            Left            =   1350
            TabIndex        =   43
            Top             =   1500
            Width           =   1050
         End
         Begin MSFlexGridLib.MSFlexGrid gridGarantias 
            Height          =   1200
            Left            =   105
            TabIndex        =   45
            Top             =   270
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   2117
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            SelectionMode   =   1
         End
      End
      Begin VB.TextBox txtTempoServico 
         Height          =   330
         Left            =   8565
         MaxLength       =   100
         TabIndex        =   41
         Top             =   5250
         Width           =   3360
      End
      Begin VB.TextBox txtEntrega 
         Height          =   315
         Left            =   840
         MaxLength       =   200
         TabIndex        =   39
         Top             =   5250
         Width           =   6225
      End
      Begin VB.Frame Frame7 
         Caption         =   "Ação"
         Height          =   780
         Left            =   9630
         TabIndex        =   31
         Top             =   9540
         Width           =   2475
         Begin VB.CommandButton btnImprimir 
            BackColor       =   &H000080FF&
            Caption         =   "&Imprimir"
            Height          =   330
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   285
            Width           =   975
         End
         Begin VB.CommandButton btnFechar 
            Caption         =   "&Fechar"
            Height          =   330
            Left            =   1335
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   270
            Width           =   990
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Lista de Orçamentos já realizados"
         Height          =   6990
         Left            =   12555
         TabIndex        =   29
         Top             =   570
         Width           =   7605
         Begin MSFlexGridLib.MSFlexGrid gridOrcamentos 
            Height          =   6570
            Left            =   75
            TabIndex        =   30
            Top             =   240
            Width           =   7470
            _ExtentX        =   13176
            _ExtentY        =   11589
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            SelectionMode   =   1
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Ação"
         Height          =   795
         Left            =   165
         TabIndex        =   24
         Top             =   9525
         Width           =   3840
         Begin VB.CommandButton Cmd_Inclui 
            BackColor       =   &H0080FF80&
            Caption         =   "&Incluir"
            Height          =   360
            Left            =   105
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   270
            Width           =   795
         End
         Begin VB.CommandButton Cmd_Salva 
            Caption         =   "&Salvar"
            Enabled         =   0   'False
            Height          =   360
            Left            =   1980
            TabIndex        =   27
            Top             =   255
            Width           =   765
         End
         Begin VB.CommandButton Cmd_Edita 
            Caption         =   "&Editar"
            Height          =   360
            Left            =   1035
            TabIndex        =   26
            Top             =   270
            Width           =   780
         End
         Begin VB.CommandButton Cmd_Exclui 
            BackColor       =   &H008080FF&
            Caption         =   "E&xcluir"
            Height          =   360
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   270
            Width           =   750
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Soluções propostas"
         Height          =   1860
         Left            =   4635
         TabIndex        =   22
         Top             =   3315
         Width           =   7815
         Begin VB.CommandButton btnEditarSolucao 
            Caption         =   "Editar"
            Height          =   300
            Left            =   2325
            TabIndex        =   87
            Top             =   1440
            Width           =   960
         End
         Begin VB.CommandButton btnRemoverSolucao 
            Caption         =   "Remover"
            Height          =   300
            Left            =   1170
            TabIndex        =   35
            Top             =   1440
            Width           =   990
         End
         Begin VB.CommandButton btnAdicionarSolucao 
            Caption         =   "Adicionar"
            Height          =   300
            Left            =   105
            TabIndex        =   34
            Top             =   1455
            Width           =   960
         End
         Begin MSFlexGridLib.MSFlexGrid gridSolucoes 
            Height          =   1140
            Left            =   105
            TabIndex        =   23
            Top             =   270
            Width           =   7620
            _ExtentX        =   13441
            _ExtentY        =   2011
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            SelectionMode   =   1
         End
         Begin VB.Label lblSolucoesTotalGeral 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5760
            TabIndex        =   37
            Top             =   1515
            Width           =   1065
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TOTAL"
            Height          =   195
            Left            =   5070
            TabIndex        =   36
            Top             =   1500
            Width           =   525
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Necessidade"
         Height          =   1830
         Left            =   165
         TabIndex        =   20
         Top             =   3315
         Width           =   4290
         Begin VB.TextBox txtNecessidade 
            BackColor       =   &H00FFFFFF&
            Height          =   1440
            Left            =   105
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   270
            Width           =   4065
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cliente / Prospect"
         Height          =   2175
         Left            =   150
         TabIndex        =   6
         Top             =   1080
         Width           =   12300
         Begin VB.CommandButton btnConverterProspect 
            Caption         =   "Converter pospect em cliente"
            Height          =   225
            Left            =   9120
            TabIndex        =   62
            Top             =   0
            Width           =   2655
         End
         Begin VB.CommandButton btnProspect 
            Caption         =   "Prospect"
            Enabled         =   0   'False
            Height          =   225
            Left            =   3960
            TabIndex        =   61
            Top             =   585
            Width           =   975
         End
         Begin VB.CheckBox chkComercial 
            Caption         =   "Comercial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9240
            TabIndex        =   16
            ToolTipText     =   "Marque caso queira que apareça no orçamento impresso."
            Top             =   1815
            Width           =   2550
         End
         Begin VB.CheckBox chkCelular 
            Caption         =   "Celular"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   14
            ToolTipText     =   "Marque caso queira que apareça no orçamento impresso."
            Top             =   1800
            Width           =   3630
         End
         Begin VB.ComboBox cboClientes 
            Height          =   315
            Left            =   900
            TabIndex        =   7
            Top             =   270
            Width           =   4050
         End
         Begin MSFlexGridLib.MSFlexGrid gridEnderecos 
            Height          =   900
            Left            =   135
            TabIndex        =   11
            ToolTipText     =   "Clique duas vezes no endereço que será usado para este orçamento."
            Top             =   915
            Width           =   11820
            _ExtentX        =   20849
            _ExtentY        =   1588
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            SelectionMode   =   1
         End
         Begin VB.CheckBox chkResidencial 
            Caption         =   "Residencial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4950
            TabIndex        =   15
            ToolTipText     =   "Marque caso queira que apareça no orçamento impresso."
            Top             =   1800
            Width           =   3825
         End
         Begin VB.Label lblTelefone 
            AutoSize        =   -1  'True
            Caption         =   "Telefone(s):"
            Height          =   195
            Left            =   135
            TabIndex        =   63
            Top             =   1860
            Width           =   840
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   6060
            TabIndex        =   59
            Top             =   225
            Width           =   540
         End
         Begin VB.Label lblCodigo 
            AutoSize        =   -1  'True
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6660
            TabIndex        =   58
            Top             =   225
            Width           =   195
         End
         Begin VB.Label lblCPF 
            AutoSize        =   -1  'True
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   9045
            TabIndex        =   19
            Top             =   255
            Width           =   195
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   9045
            TabIndex        =   18
            Top             =   585
            Width           =   195
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Email:"
            Height          =   195
            Left            =   8595
            TabIndex        =   17
            Top             =   585
            Width           =   420
         End
         Begin VB.Label lblContato 
            AutoSize        =   -1  'True
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6660
            TabIndex        =   13
            Top             =   555
            Width           =   195
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Contato:"
            Height          =   195
            Left            =   6000
            TabIndex        =   12
            Top             =   555
            Width           =   600
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Endereço(s):"
            Height          =   195
            Left            =   150
            TabIndex        =   10
            Top             =   690
            Width           =   900
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CPF:"
            Height          =   195
            Left            =   8670
            TabIndex        =   9
            Top             =   255
            Width           =   345
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Nome:"
            Height          =   255
            Left            =   180
            TabIndex        =   8
            Top             =   345
            Width           =   615
         End
      End
      Begin MSComCtl2.DTPicker dtpData 
         Height          =   315
         Left            =   645
         TabIndex        =   3
         Top             =   675
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   16777215
         Format          =   160956417
         CurrentDate     =   44025
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Validade:"
         Height          =   195
         Left            =   4515
         TabIndex        =   50
         Top             =   780
         Width           =   660
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tempo de serviço:"
         Height          =   195
         Left            =   7200
         TabIndex        =   40
         Top             =   5355
         Width           =   1320
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Entrega:"
         Height          =   195
         Left            =   180
         TabIndex        =   38
         Top             =   5370
         Width           =   600
      End
      Begin VB.Label lblNumeroOrcamento_old 
         BackColor       =   &H00C0C0C0&
         Caption         =   "..."
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
         Left            =   3045
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Número:"
         Height          =   255
         Left            =   2370
         TabIndex        =   4
         Top             =   750
         Width           =   615
      End
      Begin VB.Label Lbl_Registro 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   195
         Left            =   165
         TabIndex        =   2
         Top             =   795
         Width           =   390
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Orçamentos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   45
         TabIndex        =   1
         Top             =   150
         Width           =   20160
      End
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Left            =   0
      TabIndex        =   80
      Top             =   0
      Width           =   90
   End
End
Attribute VB_Name = "frmOrcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conn As ADODB.Connection
Private Path As String
Private FileName As String
Public strRel1_lblTelefone As String

Public OrcamentoSolucaoDescricao As String
Public OrcamentoSolucaoQtd As Integer
Public OrcamentoSolucaoValUnitario As Double
Public OrcamentoSolucaoValTotal As Double

Public OrcamentoGarantiaCodigo As Integer
Public OrcamentoGarantiaDescricao As String

Public OrcamentoFormaPagamentoCodigo As Integer
Public OrcamentoFormaPagamentoDescricao As String
Public OrcamentoFormaPagamentoCancelado As Boolean

Public OrcamentoObservacaoCodigo As Integer
Public OrcamentoObservacaoDescricao As String

Public bEditando As Boolean
Public dbname As String

Public qtdSolucoesInicio As Integer
Public qtdGarantiasInicio As Integer
Public qtdFormasPagtoInicio As Integer
Public qtdObsInicio As Integer
Public strNomeProspect As String
Public globalProximoNumeroOrcamento As Integer
'VARIÁVEIS PARA SEREM USADAS NA IMPRESSÃO DO RELATÓRIO (EXCETO PARA AS SOLUÇÕES)
Public relTelCelular As String
Public relTelResidencial As String
Public relTelComercial As String
Public relTelefone As String 'Formado com a concatenação de relTelCelular, relTelResidencial e relTelComercial

Public pblc_strClienteTipo As String
Public pblc_strTipoPesquisa As String
Public bln_editando As Boolean
Public bln_cancelado As Boolean
Public pblc_qtdObservacoes As Integer

Public descontoPorValor As Boolean
Public descontoPorPorcentagem As Boolean

Public descontoValorConcedido As Double
Public editandoSolucaso As Boolean
Public strStatusOrcamento As String
Public strStatusDescricao As String

Private Sub btnClonar_Click()
 If MsgBox("Confirma a clonagem do orçamento " & lblNumeroOrcamento.Text, vbYesNo, "Clonagem de orçamento") = vbYes Then
    clonagem
    MsgBox "O orçamento clonado possui o código: " & globalProximoNumeroOrcamento
 End If
End Sub

Private Sub btnConverterProspect_Click()
    
    If lblCodigo.Caption <> "" Then
        converteProspect 'pendenre
    Else
        MsgBox "Selecione um prospect"
    End If

End Sub

Private Sub btnDesconto_Click()
    Dim contador As Integer
    
    If txtDesconto.Text <> "" Then
        frmDesconto.valorQueveio = txtDesconto.Text
    End If
    
    'For contador = 1 To Len(txtDesconto.Text)
    '    'MsgBox Mid(txtDesconto.Text, contador, 1)
    '    If Mid(txtDesconto.Text, contador, 1) = "%" Then
    '        frmDesconto.tipoDesconto = "porcentagem"
    '    End If
    'Next contador
    
    
    frmDesconto.Show (1)
    
    If descontoValorConcedido > 0 Then
        If descontoPorValor Then
            Frame13.Caption = "Desconto R$"
        Else
            Frame13.Caption = "Desconto %"
        End If
    Else
        txtDesconto.Text = "0,00"
        Frame13.Caption = "Desconto"
    End If
    
            txtDesconto.Text = descontoValorConcedido
    
End Sub

Private Sub btnEditarFormaPagamento_Click()
    If gridFormasPagamento.Rows = 2 Then
        MsgBox "Não há forma de pagamento selecionada para editar"
        Exit Sub
    End If
       
    bln_editando = True
    Call editaFormaDePagamento(lblNumeroOrcamento.Text, gridFormasPagamento.TextMatrix(gridFormasPagamento.RowSel, 0))

    If (bln_editando) And (OrcamentoFormaPagamentoCancelado = False) Then
        With gridFormasPagamento
            gridFormasPagamento.TextMatrix(gridFormasPagamento.RowSel, 0) = OrcamentoFormaPagamentoDescricao
        End With
    End If

    bln_editando = False

End Sub

Private Sub editaFormaDePagamento(strNumeroOrcamento As String, strDescricao As String)
    frmOrcamentoFormasPagamento.blnEditando = True
    frmOrcamentoFormasPagamento.strNumeroOrcamento = strNumeroOrcamento
    frmOrcamentoFormasPagamento.strDescricaoFormaDePagamento = strDescricao
    frmOrcamentoFormasPagamento.Show (1)
End Sub


Private Sub editaObservacoes2(strNumeroOrcamento As String, strDescricao As String)
    frmOrcamentoObservacoes.blnEditando = True
    frmOrcamentoObservacoes.strNumeroOrcamento = strNumeroOrcamento
    frmOrcamentoObservacoes.txtDescricao = strDescricao
    frmOrcamentoObservacoes.Show (1)
    
End Sub

Private Sub btnEditarGarantia_Click()
        
    If gridGarantias.Rows = 2 Then
        MsgBox "Não há solução selecionada para editar"
        Exit Sub
    End If
       
    bln_editando = True
    Call editaGarantia(lblNumeroOrcamento.Text, gridGarantias.TextMatrix(gridGarantias.RowSel, 0))

    If bln_editando And Not frmOrcamentoGarantias.cancelou Then
        With gridGarantias
            gridGarantias.TextMatrix(gridGarantias.RowSel, 0) = OrcamentoGarantiaDescricao
        End With
    End If

    bln_editando = False


End Sub

Private Sub btnEditarObs_Click()
    If gridObservacoes.Rows = 2 Then
        MsgBox "Não há observação selecionada para a edição."
        Exit Sub
    End If
       
    bln_editando = True
    
    Call editaObservacoes2(lblNumeroOrcamento.Text, gridObservacoes.TextMatrix(gridObservacoes.RowSel, 0))
    gridObservacoes.TextMatrix(gridObservacoes.RowSel, 0) = OrcamentoObservacaoDescricao
    
    



'    If (bln_editando) And (OrcamentoFormaPagamentoCancelado = False) Then
'        With gridFormasPagamento
'            gridFormasPagamento.TextMatrix(gridFormasPagamento.RowSel, 0) = OrcamentoFormaPagamentoDescricao
'        End With
'    End If

    bln_editando = False

End Sub

Private Sub btnEditarSolucao_Click()
    editandoSolucaso = True
       
    If gridSolucoes.Rows = 2 Then
        MsgBox "Não há solução selecionada para editar"
        Exit Sub
    End If
       
    bln_editando = True
    bln_cancelado = False
    Call editaSolucao(lblNumeroOrcamento.Text, gridSolucoes.TextMatrix(gridSolucoes.RowSel, 1), gridSolucoes.TextMatrix(gridSolucoes.RowSel, 2), gridSolucoes.TextMatrix(gridSolucoes.RowSel, 3), gridSolucoes.TextMatrix(gridSolucoes.RowSel, 4))


    If (bln_editando) And (bln_cancelado = False) Then
        With gridSolucoes
            gridSolucoes.TextMatrix(gridSolucoes.RowSel, 1) = OrcamentoSolucaoDescricao
            gridSolucoes.TextMatrix(gridSolucoes.RowSel, 2) = OrcamentoSolucaoQtd
            gridSolucoes.TextMatrix(gridSolucoes.RowSel, 3) = OrcamentoSolucaoValUnitario
            gridSolucoes.TextMatrix(gridSolucoes.RowSel, 4) = OrcamentoSolucaoValTotal
        End With
    End If
    
    bln_cancelado = False

    bln_editando = False
    
    editandoSolucaso = False

End Sub

Private Sub btnFechar_Click()
    Unload Me
End Sub

Private Sub carregaGridFormaPagamento()
    If Trim(OrcamentoFormaPagamentoDescricao) <> "" Then
        With gridFormasPagamento
            OrcamentoFormaPagamentoCodigo = (.Rows + 1) - 2
            .AddItem (OrcamentoFormaPagamentoDescricao)
            .Refresh
        End With
    End If
End Sub
Private Sub btnAdicionarFormaPagamento_Click()
    bln_editando = False
    
    If gridFormasPagamento.Rows - 2 >= 3 Then
        MsgBox "Formas de pagamento: Máximo de 3 linhas atingindo"
        Exit Sub
    End If
    
    frmOrcamentoFormasPagamento.Show (1)
    Call carregaGridFormaPagamento
End Sub

Private Sub carregaGridGarantias()
    If Trim(OrcamentoGarantiaDescricao) <> "" Then
        With gridGarantias
            OrcamentoGarantiaCodigo = (.Rows + 1) - 2
            .AddItem (OrcamentoGarantiaDescricao)
            .Refresh
        End With
    End If
End Sub

Private Sub carregaGridObservacoes()
    If Trim(OrcamentoObservacaoDescricao) <> "" Then
        With gridObservacoes
            OrcamentoObservacaoCodigo = (.Rows + 1) - 2
            .AddItem (OrcamentoObservacaoDescricao)
            .Refresh
        End With
    End If
End Sub

Private Sub removeGridGarantias(nIndice As Integer)
    With gridGarantias
        .RemoveItem (nIndice)
        .Refresh
    End With
End Sub

Private Sub removeGridSolucoes(nIndice As Integer)
    With gridSolucoes
        .RemoveItem (nIndice)
        .Refresh
    End With
End Sub

Private Sub removeGridFormasPagamento(nIndice As Integer)
    With gridFormasPagamento
        .RemoveItem (nIndice)
        .Refresh
    End With
End Sub

Private Sub removeGridObservacoes(nIndice As Integer)
    With gridObservacoes
        .RemoveItem (nIndice)
        .Refresh
    End With
End Sub

Private Sub btnAdicionarGarantia_Click()
    If gridGarantias.Rows - 2 >= 3 Then
        MsgBox "Garantias: Máximo de 3 linhas atingindo"
        Exit Sub
    End If
    
    frmOrcamentoGarantias.Show (1)
    Call carregaGridGarantias
End Sub

Private Sub btnAdicionarObservacao_Click()
    If gridObservacoes.Rows - 2 >= 4 Then
        MsgBox "Limite: 4 linhas", vbOKOnly, "XSol Aquecedores"
        Exit Sub
    End If
    frmOrcamentoObservacoes.Show (1)
    Call carregaGridObservacoes
End Sub

Private Sub carregaGridSolucoes()

Dim contador As Integer
Dim ultimoitem As Integer
    
    'pegar o valor da ultima coluna do item
    ultimoitem = 0
    For contador = 1 To gridSolucoes.Rows - 2
        ultimoitem = CDbl(gridSolucoes.TextMatrix(contador + 1, 0))
    Next contador
    



    If Trim(OrcamentoSolucaoDescricao) <> "" Then
        With gridSolucoes
            .AddItem (ultimoitem + 1 & vbTab & OrcamentoSolucaoDescricao & vbTab & OrcamentoSolucaoQtd & vbTab & FormatCurrency(OrcamentoSolucaoValUnitario) & vbTab & FormatCurrency(OrcamentoSolucaoValTotal))
            .Refresh
        End With
    End If
End Sub

Private Sub totalGeralSolucoes()
    Dim contadorGrid As Integer
    Dim calculoTotal As Double
    For contadorGrid = 1 To (gridSolucoes.Rows - 2)
        calculoTotal = calculoTotal + CDbl(gridSolucoes.TextMatrix(contadorGrid + 1, 4))
    Next contadorGrid
    
    lblSolucoesTotalGeral.Caption = "R$" & Format(calculoTotal, "###,###,##0.00")
End Sub

Private Sub btnAdicionarSolucao_Click()

    frmOrcamentoSolucoes.Show (1)
    Call carregaGridSolucoes
    
    Call totalGeralSolucoes
End Sub

Private Sub executaRelatorio(nModeloRelatorio As Integer, pathDB As String)
    Dim X As Integer
    Dim calculoUnitario As Double
    Dim calculoTotal As Double
    Dim ValDesconto As Double
    Dim contador As Integer
    
'On Error GoTo continue
    
    If (chkCelular.value = 0) And (chkResidencial.value = 0) And (chkComercial.value = 0) Then
        MsgBox "Não há um telefone informado! Clique em editar, selecione um telefone para contato com o cliente e depois emita o relatório."
        Exit Sub
    End If
    
    If nModeloRelatorio = 1 Then
    
        If descontoPorPorcentagem Then
            relOrcamentoModelo01.vTipoDesconto = "porcentagem"
        Else
            relOrcamentoModelo01.vTipoDesconto = "valor"
        End If
    
        'Impressão das observações
        relOrcamentoModelo01.lblObservacoesLinha.Caption = ""
        For contador = 2 To (gridObservacoes.Rows - 1)
            relOrcamentoModelo01.lblObservacoesLinha.Caption = relOrcamentoModelo01.lblObservacoesLinha.Caption & gridObservacoes.TextMatrix(contador, 0) & Chr(13)
        Next contador
    
        If txtValidade.Text <> "" Then
            relOrcamentoModelo01.lblValidade.Caption = txtValidade.Text
        Else
            relOrcamentoModelo01.lblValidade.Caption = "A informar"
        End If
    
        If gridFormasPagamento.Rows < 3 Then
            relOrcamentoModelo01.lblFormaPagtoLinha1.Caption = "A informar"
        Else
            If gridFormasPagamento.Rows < 4 Then
                relOrcamentoModelo01.lblFormaPagtoLinha1.Caption = gridFormasPagamento.TextMatrix(2, 0)
            Else
                relOrcamentoModelo01.lblFormaPagtoLinha1.Caption = gridFormasPagamento.TextMatrix(3, 0)
            End If
        End If
    
        If txtImagem.Text = "" And txtTextoImg1.Text = "" And txtImagem2.Text = "" And txtTextoImg2.Text = "" Then
            relOrcamentoModelo01.imgImagem.Visible = False
            relOrcamentoModelo01.imgImagem.Visible = False
        Else
            relOrcamentoModelo01.imgImagem.Visible = True
        End If
        
        'Se só tem 1 imagem
        If (Trim(txtImagem.Text) <> "") And (Trim(txtImagem2.Text) = "") Then
            'relOrcamentoModelo01.GroupFooter2.Height = 1335
        End If
        
        'Se tem 2 imagens
        If (Trim(txtImagem.Text) <> "") And (Trim(txtImagem2.Text) <> "") Then
            'relOrcamentoModelo01.GroupFooter2.Height = 10470
        End If
        
        
        relOrcamentoModelo01.DAODataControl1.DatabaseName = pathDB
        relOrcamentoModelo01.DAODataControl1.RecordSource = "SELECT A.Descricao, A.QUANTIDADE, A.VALOR_UNITARIO, A.VALOR_TOTAL FROM ORCAMENTO_SOLUCAO AS A WHERE A.Cod_Orcamento=" & lblNumeroOrcamento.Text & " ORDER BY A.codigo"
        
        relOrcamentoModelo01.lblData.Caption = Format(Str(dtpData.value), "dd, mmm yyyy")
        relOrcamentoModelo01.lblData.Caption = Format(Str(dtpData.value), "dd") & " de " & Format(Str(dtpData.value), "mmmm") & " de " & Format(Str(dtpData.value), "yyyy")
        
        
        
        relOrcamentoModelo01.lblPropostaNumero.Caption = "Proposta N° X" & lblNumeroOrcamento.Text & "/" & Mid(FormatDateTime(Now, vbShortDate), 9, 2)
        
        relOrcamentoModelo01.lblCliente.Caption = lblCodigo.Caption & " - " & cboClientes.Text
        
        relOrcamentoModelo01.lblCPF.Caption = lblCPF.Caption
        
        With gridEnderecos
            For X = 2 To .Rows - 1
                If .TextMatrix(X, 3) = "S" Then
                    relOrcamentoModelo01.lblEndereco.Caption = .TextMatrix(X, 0)
                    relOrcamentoModelo01.lblBairro.Caption = .TextMatrix(X, 1) & " / " & .TextMatrix(X, 2)
                End If
            Next X
        End With
        
        relOrcamentoModelo01.lblContato.Caption = lblContato.Caption
        
        relTelefone = relTelCelular + " " + relTelResidencial + " " + relTelComercial
        
        'pega o telefone selecionado para informar no relatório
        If chkCelular.value = 1 Then
            strRel1_lblTelefone = chkCelular.Caption
        End If
        
        If chkComercial.value = 1 Then
            strRel1_lblTelefone = chkComercial.Caption
        End If
        
        If chkResidencial.value = 1 Then
            strRel1_lblTelefone = chkResidencial.Caption
        End If
        
        If (chkCelular.value = 1) And (chkComercial.value = 1) Then
            strRel1_lblTelefone = chkCelular.Caption & " - " & chkComercial.Caption
        End If
        
        If (chkCelular.value = 1) And (chkResidencial.value = 1) Then
            strRel1_lblTelefone = chkCelular.Caption & " - " & chkResidencial.Caption
        End If
        
        If (chkComercial.value = 1) And (chkResidencial.value = 1) Then
            strRel1_lblTelefone = chkComercial.Caption & " - " & chkResidencial.Caption
        End If
        
        
        
        
        relOrcamentoModelo01.lblTelefone.Caption = strRel1_lblTelefone
        
        relOrcamentoModelo01.lblEmail.Caption = lblEmail.Caption
        relOrcamentoModelo01.lblNecessidade.Caption = txtNecessidade.Text
        
        'O calculo do TOTAL está feito no before print da banda Detail do relatório
            '----
           'MsgBox DAODataControl1.Recordset.RecordCount
        
           '----------
           If txtDesconto.Text = "" Then
                txtDesconto.Text = "0"
           End If
           
           

                If descontoPorValor Then
                     If (Trim(txtDesconto.Text) <> "") And (txtDesconto.Text > 0) Then
                         ValDesconto = CDbl(txtDesconto.Text)
                         calculoTotal = CDbl(lblSolucoesTotalGeral.Caption) - ValDesconto
                         relOrcamentoModelo01.lblTotalGeral.Caption = FormatCurrency(calculoTotal)
                     Else
                         calculoTotal = CDbl(lblSolucoesTotalGeral.Caption)
                         relOrcamentoModelo01.lblTotalGeral.Caption = FormatCurrency(calculoTotal)
                     End If
                Else
                     If (Trim(txtDesconto.Text) <> "") And (txtDesconto.Text > 0) Then
                         ValDesconto = CDbl(txtDesconto.Text)
                         calculoTotal = CDbl(lblSolucoesTotalGeral.Caption) - (CDbl(lblSolucoesTotalGeral.Caption) * ValDesconto / 100)
                         relOrcamentoModelo01.lblTotalGeral.Caption = FormatCurrency(calculoTotal)
                     Else
                         calculoTotal = CDbl(lblSolucoesTotalGeral.Caption)
                         relOrcamentoModelo01.lblTotalGeral.Caption = FormatCurrency(calculoTotal)
                     End If
                End If

                
        
        relOrcamentoModelo01.lblEntrega.Caption = txtEntrega.Text
        relOrcamentoModelo01.lblTempoServico.Caption = txtTempoServico.Text
        
        'Garantias
        With gridGarantias
            If .Rows < 3 Then
                relOrcamentoModelo01.lblGarantiaLinha1.Caption = "A informar"
            Else
                If gridGarantias.Rows < 4 Then
                    relOrcamentoModelo01.lblGarantiaLinha1.Caption = .TextMatrix(2, 0)
                    relOrcamentoModelo01.lblGarantiaLinha2.Visible = False
                Else
                    relOrcamentoModelo01.lblGarantiaLinha1.Caption = .TextMatrix(2, 0)
                    relOrcamentoModelo01.lblGarantiaLinha2.Visible = True
                    relOrcamentoModelo01.lblGarantiaLinha2.Caption = .TextMatrix(3, 0)
                End If
            End If
        
            '    relOrcamentoModelo01.lblGarantiaLinha2.Caption = ""
            'End If
            
            'If .Rows - 1 > 3 Then
            '    relOrcamentoModelo01.lblGarantiaLinha3.Caption = .TextMatrix(4, 0)
            'Else
            '    relOrcamentoModelo01.lblGarantiaLinha3.Caption = ""
            'End If
        End With
        
        
        'Formas de pagamento
        With gridFormasPagamento
        
        
            If .Rows < 3 Then
                relOrcamentoModelo01.lblFormaPagtoLinha1.Caption = "A informar"
            Else
                If .Rows < 4 Then
                    relOrcamentoModelo01.lblFormaPagtoLinha1.Caption = .TextMatrix(2, 0)
                    relOrcamentoModelo01.lblFormaPagtoLinha2.Visible = False
                    relOrcamentoModelo01.lblFormaPagtoLinha3.Visible = False
                Else
                    relOrcamentoModelo01.lblFormaPagtoLinha1.Caption = .TextMatrix(2, 0)
                    relOrcamentoModelo01.lblFormaPagtoLinha2.Visible = True
                    relOrcamentoModelo01.lblFormaPagtoLinha3.Visible = False
                    relOrcamentoModelo01.lblFormaPagtoLinha2.Caption = .TextMatrix(3, 0)
                End If
                
                If .Rows = 5 Then
                    relOrcamentoModelo01.lblFormaPagtoLinha1.Caption = .TextMatrix(2, 0)
                    relOrcamentoModelo01.lblFormaPagtoLinha2.Visible = True
                    relOrcamentoModelo01.lblFormaPagtoLinha3.Visible = True
                    relOrcamentoModelo01.lblFormaPagtoLinha2.Caption = .TextMatrix(3, 0)
                    relOrcamentoModelo01.lblFormaPagtoLinha3.Caption = .TextMatrix(4, 0)
                End If
            End If
        
            '    relOrcamentoModelo01.lblFormaPagtoLinha2.Caption = ""
            'End If
            
            'If .Rows - 1 > 3 Then
            '    relOrcamentoModelo01.lblFormaPagtoLinha3.Caption = .TextMatrix(4, 0)
            'Else
            '    relOrcamentoModelo01.lblFormaPagtoLinha3.Caption = ""
            'End If
        End With
        
        
        relOrcamentoModelo01.lblValidade.Caption = txtValidade.Text
        
        'Observações
        With gridObservacoes
            'If .Rows - 1 > 1 Then
            '    relOrcamentoModelo01.lblObservacoesLinha1.Caption = .TextMatrix(2, 0)
            'Else
            '    relOrcamentoModelo01.lblObservacoesLinha1.Caption = ""
            'End If
            
            'If .Rows - 1 > 2 Then
            '    relOrcamentoModelo01.lblObservacoesLinha2.Caption = .TextMatrix(3, 0)
            'Else
            '    relOrcamentoModelo01.lblObservacoesLinha2.Caption = ""
            'End If
            
            'If .Rows - 1 > 3 Then
            '    relOrcamentoModelo01.lblObservacoesLinha3.Caption = .TextMatrix(4, 0)
            'Else
            '    relOrcamentoModelo01.lblObservacoesLinha3.Caption = ""
            'End If
            
            'If .Rows - 1 > 4 Then
            '    relOrcamentoModelo01.lblObservacoesLinha4.Caption = .TextMatrix(5, 0)
            'Else
            '    relOrcamentoModelo01.lblObservacoesLinha4.Caption = ""
            'End If
            
            'If .Rows - 1 > 5 Then
            '    relOrcamentoModelo01.lblObservacoesLinha5.Caption = .TextMatrix(6, 0)
            'Else
            '    relOrcamentoModelo01.lblObservacoesLinha5.Caption = ""
            'End If
        End With
        
        If Trim(txtImagem.Text) <> "" Then
            relOrcamentoModelo01.imgImagem.Picture = LoadPicture(txtImagem.Text)
        End If
        
        
        If Trim(txtImagem2.Text) <> "" Then
            relOrcamentoModelo01.ImgImagem2.Picture = LoadPicture(txtImagem2.Text)
        End If
        
        
        If (Trim(txtDesconto.Text) = "") Or (txtDesconto.Text = "0,00") Or (txtDesconto.Text = "0") Then
            'NÃO TEM DESCONTO
            relOrcamentoModelo01.lblDesconto.Visible = False
            relOrcamentoModelo01.lblDescontoValor.Visible = False
            relOrcamentoModelo01.lblSubtotal.Visible = False
            relOrcamentoModelo01.lblDescontoValor.Visible = False
            relOrcamentoModelo01.lblValorOriginal.Visible = False
            relOrcamentoModelo01.lblDescontoValorNegativo.Visible = False
            relOrcamentoModelo01.linhaDesconto.Visible = False
        Else
            'TEM DESCONTO

            relOrcamentoModelo01.lblDesconto.Visible = True
            relOrcamentoModelo01.lblDescontoValor.Visible = True
            relOrcamentoModelo01.lblSubtotal.Visible = True
            relOrcamentoModelo01.lblDescontoValor.Visible = True
            relOrcamentoModelo01.lblValorOriginal.Visible = True
            relOrcamentoModelo01.lblDescontoValorNegativo.Visible = True
            relOrcamentoModelo01.linhaDesconto.Visible = True

            relOrcamentoModelo01.lblDescontoValor.Caption = txtDesconto.Text
            
            If descontoPorPorcentagem Then
                relOrcamentoModelo01.lblDescontoValor.Caption = txtDesconto.Text & " % " '& Str((CDbl(lblSolucoesTotalGeral.Caption) * ValDesconto / 100)) & ")"
                relOrcamentoModelo01.lblValorOriginal.Caption = FormatCurrency(CDbl(lblSolucoesTotalGeral.Caption))
                relOrcamentoModelo01.lblDescontoValorNegativo = "-" & FormatCurrency(CDbl(lblSolucoesTotalGeral.Caption) * ValDesconto / 100)
            Else
                relOrcamentoModelo01.lblValorOriginal.Caption = FormatCurrency(CDbl(lblSolucoesTotalGeral.Caption))
                relOrcamentoModelo01.lblDescontoValor.Caption = "R$" & txtDesconto.Text
                relOrcamentoModelo01.lblDescontoValorNegativo = "-" & FormatCurrency(CDbl(ValDesconto))
            End If
            'relOrcamentoModelo01.lblTotalGeral.Caption = relOrcamentoModelo01.lblTotalGeral.Caption - txtDesconto.Text
        End If
        
        
        relOrcamentoModelo01.Show (1)
    
    
    ElseIf nModeloRelatorio = 2 Then
        
        'CALCULA O VALOR UNITÁRIO E TOTAL
        gridSolucoes.RowSel = 2
        For X = 1 To gridSolucoes.Rows - 2
            calculoUnitario = calculoUnitario + gridSolucoes.TextMatrix(gridSolucoes.RowSel, 2)
            calculoTotal = calculoTotal + gridSolucoes.TextMatrix(gridSolucoes.RowSel, 3)
            'gridSolucoes.RowSel = gridSolucoes.Rows + 1 'Vai dar erro então vai para o continue, em debug tem que puxar
        Next X
        
continue:
    
    
        If nModeloRelatorio = 2 Then
        
            
            'relOrcamentoModelo02.lblTotalGeral_ValUnitario = FormatCurrency(calculoUnitario)
            relOrcamentoModelo02.lblTotalGeral_ValTotal = FormatCurrency(calculoTotal)
        
            relOrcamentoModelo02.DAODataControl1.DatabaseName = pathDB
            relOrcamentoModelo02.DAODataControl1.RecordSource = "Select Quantidade, Descricao, Valor_Unitario, Valor_Total  from orcamento_solucao Where COD_ORCAMENTO = " & lblNumeroOrcamento.Text
            
            'relOrcamentoModelo02.lblDataOrcamento.Caption = Format(Str(dtpData.value), "Long date")
            relOrcamentoModelo02.lblDataOrcamento.Caption = Format(Str(dtpData.value), "dd") & " de " & Format(Str(dtpData.value), "mmmm") & " de " & Format(Str(dtpData.value), "yyyy")
            'relOrcamentoModelo02.lblPropostaNumero.Caption = lblNumeroOrcamento.Text
            
            
            relOrcamentoModelo02.lblNomeCliente.Caption = lblCodigo.Caption & " - " & cboClientes.Text & " - " & strRel1_lblTelefone
            relOrcamentoModelo02.lblTpoServico.Caption = txtNecessidade.Text
            relOrcamentoModelo02.lblNumeroOrcamento.Caption = "X" & lblNumeroOrcamento.Text
            'relOrcamentoModelo02.lblEnderecoCliente.Caption =
            
            'Endereço no orçamento 2
            With gridEnderecos
                For X = 2 To .Rows - 1
                    If .TextMatrix(X, 3) = "S" Then
                        relOrcamentoModelo02.lblEnderecoCliente.Caption = .TextMatrix(X, 0)
                        relOrcamentoModelo02.lblEnderecoCliente.Caption = relOrcamentoModelo02.lblEnderecoCliente.Caption & " " & .TextMatrix(X, 1) & " - " & .TextMatrix(X, 2)
                    End If
                Next X
            End With
            
            'Formas de pagamento no orçamento 2
            With gridFormasPagamento
                If .Rows - 1 > 1 Then
                    relOrcamentoModelo02.lblFormaPagamento_Linha1.Caption = .TextMatrix(2, 0)
                End If
                
                If .Rows - 1 > 2 Then
                    relOrcamentoModelo02.lblFormaPagamento_Linha2.Caption = .TextMatrix(3, 0)
                Else
                    relOrcamentoModelo02.lblFormaPagamento_Linha2.Caption = ""
                End If
                
                If .Rows - 1 > 3 Then
                    relOrcamentoModelo02.lblFormaPagamento_Linha3.Caption = .TextMatrix(4, 0)
                Else
                    relOrcamentoModelo02.lblFormaPagamento_Linha3.Caption = ""
                End If
                
            End With
            
            'Garantias no orçamento 2
            With gridGarantias
                If .Rows - 1 > 1 Then
                    relOrcamentoModelo02.lblGarantia_Linha1.Caption = .TextMatrix(2, 0)
                End If
            
                If .Rows - 1 > 2 Then
                    relOrcamentoModelo02.lblGarantia_Linha2.Caption = .TextMatrix(3, 0)
                Else
                    relOrcamentoModelo02.lblGarantia_Linha2.Caption = ""
                End If
                
            End With
            
            relOrcamentoModelo02.lblValidade.Caption = txtValidade.Text
            
            'Observações no orçamento 2
            With gridObservacoes
                relOrcamentoModelo02.lblObs_Linha1.Caption = ""
                
                For X = 2 To .Rows - 1
                        relOrcamentoModelo02.lblObs_Linha1.Caption = relOrcamentoModelo02.lblObs_Linha1.Caption & .TextMatrix(X, 0) & ". "
                Next X
                'If .Rows - 1 > 1 Then
                '    relOrcamentoModelo02.lblObs_Linha1.Caption = .TextMatrix(2, 0)
                'Else
                '    relOrcamentoModelo02.lblObs_Linha1.Caption = "Sem observações"
                'End If
                
                'If .Rows - 1 > 2 Then
                '    relOrcamentoModelo02.lblObs_Linha2.Caption = .TextMatrix(3, 0)
                'Else
                '    relOrcamentoModelo02.lblObs_Linha2.Caption = ""
                'End If
                
                'If .Rows - 1 > 3 Then
                '    relOrcamentoModelo02.lblObs_Linha3.Caption = .TextMatrix(4, 0)
                'Else
                '    relOrcamentoModelo02.lblObs_Linha3.Caption = ""
                'End If
                
                'If .Rows - 1 > 4 Then
                '    relOrcamentoModelo02.lblObs_Linha4.Caption = .TextMatrix(5, 0)
                'Else
                '    relOrcamentoModelo02.lblObs_Linha4.Caption = ""
                'End If
                
'                If .Rows - 1 > 5 Then
'                    relOrcamentoModelo02.lblObs_Linha5.Caption = .TextMatrix(6, 0)
'                Else
'                    relOrcamentoModelo02.lblObs_Linha5.Caption = ""
'                End If
            End With
            
            
            If txtDesconto.Text = "0,00" Then
                relOrcamentoModelo02.lblDesconto.Visible = False
                relOrcamentoModelo02.lblDescontoValor.Visible = False
            Else
                relOrcamentoModelo02.lblDesconto.Visible = True
                relOrcamentoModelo02.lblDescontoValor.Visible = True
                relOrcamentoModelo02.lblDescontoValor.Caption = txtDesconto.Text
            End If
            
            If (Trim(txtDesconto.Text) <> "") Then
                If txtDesconto.Text > 0 Then
                    ValDesconto = CDbl(txtDesconto.Text)
                    calculoTotal = CDbl(lblSolucoesTotalGeral.Caption) - ValDesconto
                    relOrcamentoModelo02.lblTotalGeral_ValTotal.Caption = FormatCurrency(calculoTotal)
                End If
            Else
                calculoTotal = CDbl(lblSolucoesTotalGeral.Caption)
                relOrcamentoModelo02.lblTotalGeral_ValTotal.Caption = FormatCurrency(calculoTotal)
            End If

           If txtDesconto.Text = "" Then 'paradero
                txtDesconto.Text = "0"
           End If
           
           If descontoPorValor Then
                If (Trim(txtDesconto.Text) <> "") And (txtDesconto.Text > 0) Then
                    ValDesconto = CDbl(txtDesconto.Text)
                    calculoTotal = CDbl(lblSolucoesTotalGeral.Caption) - ValDesconto
                    relOrcamentoModelo02.lblTotalGeral_ValTotal.Caption = FormatCurrency(calculoTotal)
                Else
                    calculoTotal = CDbl(lblSolucoesTotalGeral.Caption)
                    relOrcamentoModelo02.lblTotalGeral_ValTotal.Caption = FormatCurrency(calculoTotal)
                End If
           Else
                If (Trim(txtDesconto.Text) <> "") And (txtDesconto.Text > 0) Then
                    ValDesconto = CDbl(txtDesconto.Text)
                    calculoTotal = CDbl(lblSolucoesTotalGeral.Caption) - (CDbl(lblSolucoesTotalGeral.Caption) * ValDesconto / 100)
                    relOrcamentoModelo02.lblTotalGeral_ValTotal.Caption = FormatCurrency(calculoTotal)
                Else
                    calculoTotal = CDbl(lblSolucoesTotalGeral.Caption)
                    relOrcamentoModelo02.lblTotalGeral_ValTotal.Caption = FormatCurrency(calculoTotal)
                End If
           End If

            'ajusta os descontos
        If (Trim(txtDesconto.Text) = "") Or (txtDesconto.Text = "0,00") Or (txtDesconto.Text = "0") Then
            'NÃO TEM DESCONTO
            relOrcamentoModelo02.lblDesconto.Visible = False
            relOrcamentoModelo02.lblDescontoValor.Visible = False
            relOrcamentoModelo02.lblSubtoal.Visible = False
            relOrcamentoModelo02.lblDescontoValor.Visible = False
            relOrcamentoModelo02.lblValorOriginal.Visible = False
            relOrcamentoModelo02.lblDescontoValorNegativo.Visible = False
            relOrcamentoModelo02.linhaDesconto.Visible = False
        Else
            'TEM DESCONTO

            relOrcamentoModelo02.lblDesconto.Visible = True
            relOrcamentoModelo02.lblDescontoValor.Visible = True
            relOrcamentoModelo02.lblSubtoal.Visible = True
            relOrcamentoModelo02.lblDescontoValor.Visible = True
            relOrcamentoModelo02.lblValorOriginal.Visible = True
            relOrcamentoModelo02.lblDescontoValorNegativo.Visible = True
            relOrcamentoModelo02.linhaDesconto.Visible = True
'
            relOrcamentoModelo02.lblDescontoValor.Caption = txtDesconto.Text
'
            If descontoPorPorcentagem Then
                relOrcamentoModelo02.lblDescontoValor.Caption = txtDesconto.Text & " % " '& Str((CDbl(lblSolucoesTotalGeral.Caption) * ValDesconto / 100)) & ")"
                relOrcamentoModelo02.lblValorOriginal.Caption = FormatCurrency(CDbl(lblSolucoesTotalGeral.Caption))
                relOrcamentoModelo02.lblDescontoValorNegativo = "-" & FormatCurrency(CDbl(lblSolucoesTotalGeral.Caption) * ValDesconto / 100)
            Else
                relOrcamentoModelo02.lblValorOriginal.Caption = FormatCurrency(CDbl(lblSolucoesTotalGeral.Caption))
                relOrcamentoModelo02.lblDescontoValor.Caption = "R$" & txtDesconto.Text
                relOrcamentoModelo02.lblDescontoValorNegativo = "-" & FormatCurrency(CDbl(ValDesconto))
            End If
            'relOrcamentoModelo02.lblTotalGeral_ValTotal.Caption = relOrcamentoModelo02.lblTotalGeral_ValTotal.Caption - txtDesconto.Text
        End If
            
            
            
            relOrcamentoModelo02.Show (1)
        Else
            relOrcamentoModelo01.Show (1)
        End If
    End If

End Sub

Private Sub btnImagem_Click()
    CommonDialog1.ShowOpen
    txtImagem.Text = CommonDialog1.FileName
End Sub

Private Sub btnImagem2_Click()
    CommonDialog2.ShowOpen
    txtImagem2.Text = CommonDialog2.FileName
End Sub

Private Sub btnImprimir_Click()
    Dim achouEndereco As Boolean
    Dim contador As Integer
    
    pblc_qtdObservacoes = gridObservacoes.Rows - 2
    
    achouEndereco = False
    
    With gridEnderecos
        For contador = 2 To .Rows - 1
            If .TextMatrix(contador, 3) = "S" Then
                achouEndereco = True
            End If
        Next contador
    End With
    
    If achouEndereco = False Then
        MsgBox "Confirme o endereço do orçamento" & Chr(13) & Chr(13) & "1. Edite o orçamento" & Chr(13) & "2.Dê um clique duplo na linha do endereço" & Chr(13) & "para alterar o campo USAR para S" & Chr(13) & "3. Salve o orçamento e imprima-o"
        Exit Sub
    End If
    
    If optnModelo1.value = False And optnModelo2.value = False Then
        MsgBox "Selecione uma das opções para o modelo do orçamento"
        optnModelo1.SetFocus
        Exit Sub
    End If
    
    
    If optnModelo1.value = True Then
        Call executaRelatorio(1, dbname)
    ElseIf optnModelo2.value = True Then
        Call executaRelatorio(2, dbname)
    End If
    
    
    
End Sub

Private Sub btnPesquisar_Click()
    If pblc_strTipoPesquisa = "CLIENTE" Then
        Call carregaGridOrcamentos(Return_Recordset(Return_SqlString(23, cboClientesProspects.Text)))
    ElseIf pblc_strTipoPesquisa = "PROSPECT" Then
        Call carregaGridOrcamentos(Return_Recordset(Return_SqlString(24, cboClientesProspects.Text)))
    ElseIf pblc_strTipoPesquisa = "DATA" Then
        Call carregaGridOrcamentos(Return_Recordset(Return_SqlString(25)))
    ElseIf pblc_strTipoPesquisa = "CODIGO" Then
        If RTrim(LTrim(txtCodigo.Text)) = "" Then
            Call carregaGridOrcamentos(Return_Recordset(Return_SqlString(10, txtCodigo.Text)))
            Exit Sub
        End If
        Call carregaGridOrcamentos(Return_Recordset(Return_SqlString(26, txtCodigo.Text)))
    End If
End Sub

Private Sub btnProspect_Click()
    'abro a tela de clientes e deixo cadastrar o nome, endereço, email e telefone do prospect
    Form_Customer.ehProspect = True
    Form_Customer.Show (1)
    Form_Customer.ehProspect = False
    'Call carregaGridOrcamentos(Return_Recordset(Return_SqlString(10)))
    Call carregaCboClientes(Return_Recordset(Return_SqlString(1)))
    
    If Trim(strNomeProspect) <> "" Then
        cboClientes.Text = strNomeProspect
        cboClientes_Click
    End If
    
End Sub

Private Sub btnRemoverFormaPagamento_Click()
    'Remove a solução selecionada no grid Formas de pagamento
    If gridFormasPagamento.RowSel = 1 Then
        MsgBox "Selecione a forma de pagamento que deseja remover"
        Exit Sub
    End If
    If gridFormasPagamento.Rows = 2 Or gridFormasPagamento.Rows = 3 Then
        gridFormasPagamento.Clear
        Call ConfigureGridFormasPagamento
        Exit Sub
    End If
    Call removeGridFormasPagamento(gridFormasPagamento.row)
    
End Sub

Private Sub btnRemoverGarantia_Click()
    'Remove a solução selecionada no grid Garantias
    If gridGarantias.RowSel = 1 Then
        MsgBox "Selecione a garantia que deseja remover"
        Exit Sub
    End If
    If gridGarantias.Rows = 2 Or gridGarantias.Rows = 3 Then
        gridGarantias.Clear
        Call ConfigureGridGarantias
        Exit Sub
    End If
    Call removeGridGarantias(gridGarantias.row)
End Sub

Private Sub btnRemoverObservacao_Click()
    'Remove a solução selecionada no grid observações
    If gridObservacoes.RowSel = 1 Then
        MsgBox "Selecione a observação que deseja remover"
        Exit Sub
    End If
    If gridObservacoes.Rows = 2 Or gridObservacoes.Rows = 3 Then
        gridObservacoes.Clear
        Call ConfigureGridObservacoes
        Exit Sub
    End If
    Call removeGridObservacoes(gridObservacoes.row)
End Sub

Private Sub btnRemoverSolucao_Click()
    'Remove a solução selecionada no grid soluções propostas
    If gridSolucoes.RowSel = 1 Then
        MsgBox "Selecione a solução que deseja remover"
        Exit Sub
    End If
    If gridSolucoes.Rows = 2 Or gridSolucoes.Rows = 3 Then
        gridSolucoes.Clear
        Call ConfigureGridSolucoes
        Exit Sub
    End If
    Call removeGridSolucoes(gridSolucoes.row)
    Call totalGeralSolucoes
End Sub

Private Function retornaProximoNumeroOrcamento(rs As ADODB.Recordset) As Integer
    
    If IsNull((rs.Fields("MAXCODE"))) Then
        retornaProximoNumeroOrcamento = 1
    Else
        retornaProximoNumeroOrcamento = rs.Fields("MAXCODE") + 1
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub limpaTela()
    'lblNumeroOrcamento.Text = "..."
    lblNumeroOrcamento.Text = "..."
    txtValidade.Text = ""
    cboClientes.Text = ""
    lblCodigo.Caption = ""
    lblCPF.Caption = ""
    lblContato.Caption = ""
    lblEmail.Caption = ""
    gridEnderecos.Clear
    
    chkCelular.Caption = "Tel. Celular"
    chkResidencial.Caption = "Tel. Resid."
    chkComercial.Caption = "Tel. Com."
    
    chkCelular.value = False
    chkResidencial.value = False
    chkComercial.value = False
    
    txtNecessidade.Text = ""
    gridSolucoes.Clear
    lblSolucoesTotalGeral.Caption = "..."
    txtEntrega.Text = ""
    txtTempoServico.Text = ""
    gridGarantias.Clear
    gridFormasPagamento.Clear
    gridObservacoes.Clear
    
    Call ConfigureGridEnderecos
    Call ConfigureGridSolucoes
    Call ConfigureGridGarantias
    Call ConfigureGridFormasPagamento
    Call ConfigureGridObservacoes
    
    qtdSolucoesInicio = 0
    qtdGarantiasInicio = 0
    qtdFormasPagtoInicio = 0
    qtdObsInicio = 0
    
End Sub

Private Sub HabilitarComponentes(bBloqueio As Boolean)
    dtpData.enabled = bBloqueio
    txtValidade.enabled = bBloqueio
    cboClientes.enabled = bBloqueio
    btnProspect.enabled = bBloqueio
    gridEnderecos.enabled = bBloqueio
    'chkCelular.enabled = bBloqueio
    'chkResidencial.enabled = bBloqueio
    'chkComercial.enabled = bBloqueio
    txtNecessidade.enabled = bBloqueio
    'gridSolucoes.enabled = bBloqueio
    btnAdicionarSolucao.enabled = bBloqueio
    btnRemoverSolucao.enabled = bBloqueio
    btnEditarSolucao.enabled = bBloqueio
    txtEntrega.enabled = bBloqueio
    txtTempoServico.enabled = bBloqueio
    gridGarantias.enabled = bBloqueio
    btnAdicionarGarantia.enabled = bBloqueio
    btnRemoverGarantia.enabled = bBloqueio
    gridFormasPagamento.enabled = bBloqueio
    btnAdicionarFormaPagamento.enabled = bBloqueio
    btnRemoverFormaPagamento.enabled = bBloqueio
    gridObservacoes.enabled = bBloqueio
    btnAdicionarObservacao.enabled = bBloqueio
    btnRemoverObservacao.enabled = bBloqueio
    btnImagem.enabled = bBloqueio
    btnImagem2.enabled = bBloqueio
    txtTextoImg1.enabled = bBloqueio
    txtTextoImg2.enabled = bBloqueio
    txtDesconto.enabled = bBloqueio
    btnEditarGarantia.enabled = bBloqueio
    btnEditarFormaPagamento.enabled = bBloqueio
    
    'cboStatus.Locked = Not bBloqueio
    'txtStatusMensagem.Locked = Not bBloqueio
    
    btnEditarObs.enabled = bBloqueio
    
    btnDesconto.enabled = bBloqueio
End Sub

Private Sub carregaInfoCliente(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    If rs.EOF Or rs.BOF Then
       Exit Sub
    Else
        lblCodigo.Caption = rs.Fields("Codigo")
        If IsNull(rs.Fields("CPF")) Then
            lblCPF.Caption = "..."
        Else
            lblCPF.Caption = rs.Fields("CPF")
        End If
    
        If IsNull(rs.Fields("Contrato")) Then
            lblContato.Caption = "..."
        Else
            lblContato.Caption = rs.Fields("Contrato")
        End If
        
        If IsNull(rs.Fields("e_mail")) Then
            lblEmail.Caption = "..."
        Else
            lblEmail.Caption = rs.Fields("e_mail")
        End If
        
        If IsNull(rs.Fields("Fone_celular")) Then
            chkCelular.Caption = "..."
        Else
            chkCelular.Caption = rs.Fields("Fone_celular")
        End If
        
        If IsNull(rs.Fields("Fone_res")) Then
            chkResidencial.Caption = "..."
        Else
            chkResidencial.Caption = rs.Fields("Fone_res")
        End If
        
        If IsNull(rs.Fields("Fone_com")) Then
            chkComercial.Caption = "..."
        Else
            chkComercial.Caption = rs.Fields("Fone_com")
        End If
        
        If rs.Fields("Tipo") = "CLIENTE" Then
            pblc_strClienteTipo = "CLIENTE"
        Else
            pblc_strClienteTipo = "PROSPECT"
        End If
        
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub carregaInfoOrcamento(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        dtpData.value = rs.Fields("Data")
        'lblNumeroOrcamento.Text = rs.Fields("CodOrcamento")
        lblNumeroOrcamento.Text = rs.Fields("CodOrcamento")
        txtValidade.Text = rs.Fields("Validade")
        cboClientes.Text = rs.Fields("NOME_CLI")
        lblCodigo.Caption = rs.Fields("CODCLI")
        
        If IsNull(rs.Fields("CPF")) Then
            lblCPF.Caption = "..."
        Else
            lblCPF.Caption = rs.Fields("CPF")
        End If
        
        
        If IsNull(rs.Fields("E_MAIL")) Then
            lblEmail.Caption = "..."
        Else
            lblEmail.Caption = rs.Fields("E_MAIL")
        End If
        
        If IsNull(rs.Fields("Fone_celular")) Then
            chkCelular.Caption = "..."
            telCelular = ""
        Else
            chkCelular.Caption = rs.Fields("Fone_celular")
            telCelular = rs.Fields("Fone_celular")
        End If
        
        If IsNull(rs.Fields("Fone_res")) Then
            chkResidencial.Caption = "..."
            telResidencial = ""
        Else
            chkResidencial.Caption = rs.Fields("Fone_res")
            telResidencial = rs.Fields("Fone_res")
        End If
        
        If IsNull(rs.Fields("Fone_com")) Then
            chkComercial.Caption = "..."
            telComercial = ""
        Else
            chkComercial.Caption = rs.Fields("Fone_com")
            telComercial = rs.Fields("Fone_com")
        End If
        
        txtNecessidade.Text = rs.Fields("Necessidade")
        txtEntrega.Text = rs.Fields("Entrega")
        txtTempoServico.Text = rs.Fields("Tempo_servico")
        
        If Not IsNull(rs.Fields("IMG_PATH")) Then
             txtImagem.Text = rs.Fields("IMG_PATH")
        Else
            txtImagem.Text = ""
        End If
        
        If Not IsNull(rs.Fields("IMG_TEXT")) Then
             txtTextoImg1.Text = rs.Fields("IMG_TEXT")
        Else
            txtTextoImg1.Text = ""
        End If
        
        If Not IsNull(rs.Fields("IMG2_PATH")) Then
             txtImagem2.Text = rs.Fields("IMG2_PATH")
        Else
            txtImagem.Text = ""
        End If
        
        If Not IsNull(rs.Fields("IMG2_TEXT")) Then
             txtTextoImg2.Text = rs.Fields("IMG2_TEXT")
        Else
            txtTextoImg1.Text = ""
        End If
        
        
        If Not IsNull(rs.Fields("Desconto")) Then
            txtDesconto.Text = rs.Fields("Desconto")
            If Not IsNull(rs.Fields("descontotipo")) Then
                If rs.Fields("descontotipo") = "porcento" Then
                    Frame13.Caption = "Desconto %"
                    descontoPorPorcentagem = True
                    descontoPorValor = False
                Else
                    Frame13.Caption = "Desconto R$"
                    descontoPorPorcentagem = False
                    descontoPorValor = True
                End If
            Else
                Frame13.Caption = "Desconto R$"
                descontoPorPorcentagem = False
                descontoPorValor = False
            End If
        Else
            txtDesconto.Text = ""
        End If
        
        If Trim(txtDesconto.Text) <> "" Then
            txtDesconto.Text = FormatNumber(txtDesconto.Text, 2, vbUseDefault, vbUseDefault, vbUseDefault)
        End If
        
        pblc_strClienteTipo = rs.Fields("Tipo")
        
        If Not IsNull(rs.Fields("STATUS")) Then
            If rs.Fields("STATUS") <> "" Then
                If rs.Fields("STATUS") = "Enviado" Then
                    cboStatus.ListIndex = 0
                    cboStatus.BackColor = vbYellow
                    If Not IsNull(rs.Fields("STATUS_MSG")) Then
                        txtStatusMensagem.Text = rs.Fields("STATUS_MSG")
                    Else
                        txtStatusMensagem.Text = ""
                    End If
                ElseIf rs.Fields("STATUS") = "Aprovado" Then
                    cboStatus.ListIndex = 1
                    cboStatus.BackColor = vbGreen
                    If Not IsNull(rs.Fields("STATUS_MSG")) Then
                        txtStatusMensagem.Text = rs.Fields("STATUS_MSG")
                    Else
                        txtStatusMensagem.Text = ""
                    End If
                ElseIf rs.Fields("STATUS") = "Reprovado" Then
                    cboStatus.ListIndex = 2
                    cboStatus.BackColor = vbRed
                    If Not IsNull(rs.Fields("STATUS_MSG")) Then
                        txtStatusMensagem.Text = rs.Fields("STATUS_MSG")
                    Else
                        txtStatusMensagem.Text = ""
                    End If
                End If
            End If
        Else
            cboStatus.ListIndex = -1
            cboStatus.BackColor = vbWhite
            txtStatusMensagem.Text = ""
        End If
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    If pblc_strClienteTipo = "CLIENTE" Then
        btnConverterProspect.enabled = False
    Else
        btnConverterProspect.enabled = True
    End If
    
End Sub

Private Sub carregaInfoSolucoes(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    selectedRow = 2
    With gridSolucoes
        Dim contador As Integer
        contador = 0
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            contador = contador + 1
            .AddItem (contador) & vbTab & rs.Fields("Descricao") & vbTab & IIf(IsNull(rs.Fields("Quantidade")), "", rs.Fields("Quantidade")) & vbTab & IIf(IsNull(rs.Fields("Valor_unitario")), "", "R$ " & rs.Fields("Valor_unitario")) & vbTab & "R$ " & Format(rs.Fields("Valor_Total"), "##,##,##0.00") & vbTab & rs.Fields("Codigo")
            
            '.AddItem (OrcamentoSolucaoDescricao & vbTab & OrcamentoSolucaoQtd & vbTab & Format(Str(OrcamentoSolucaoValUnitario), "###,###,##0.00") & vbTab & Format(Str(OrcamentoSolucaoValTotal), "##,##,##0.00"))
            
            If code = rs.Fields("Descricao") Then
                selectedRow = .Rows - 1
            End If
            rs.MoveNext
        Loop
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            If .Rows > 9 Then
                .TopRow = selectedRow
            End If
            .col = 0
            .ColSel = .Cols - 1
        Else
        End If
        
    End With
    rs.Close
    Set rs = Nothing
    
    Call totalGeralSolucoes
End Sub

Private Sub carregaInfoGarantias(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    selectedRow = 2
    With gridGarantias
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("Garantia_Descricao"))
            If code = rs.Fields("Garantia_Descricao") Then
                selectedRow = .Rows - 1
            End If
            rs.MoveNext
        Loop
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            If .Rows > 9 Then
                .TopRow = selectedRow
            End If
            .col = 0
            .ColSel = .Cols - 1
        Else
        End If
        
    End With
    rs.Close
    Set rs = Nothing
    
End Sub

Private Sub carregaInfoFormasPagto(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    selectedRow = 2
    With gridFormasPagamento
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("Forma_Pagto_Descricao"))
            If code = rs.Fields("Forma_Pagto_Descricao") Then
                selectedRow = .Rows - 1
            End If
            rs.MoveNext
        Loop
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            If .Rows > 9 Then
                .TopRow = selectedRow
            End If
            .col = 0
            .ColSel = .Cols - 1
        Else
        End If
        
    End With
    rs.Close
    Set rs = Nothing
    
End Sub

Private Sub carregaInfoObservacoes(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    selectedRow = 2
    With gridObservacoes
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("OBSERVACAO_DESCRICAO"))
            If code = rs.Fields("OBSERVACAO_DESCRICAO") Then
                selectedRow = .Rows - 1
            End If
            rs.MoveNext
        Loop
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            If .Rows > 9 Then
                .TopRow = selectedRow
            End If
            .col = 0
            .ColSel = .Cols - 1
        Else
        End If
        
    End With
    rs.Close
    Set rs = Nothing
    
End Sub

Private Sub carregaGridEnderecos(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    selectedRow = 2
    With gridEnderecos
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("Endereco") & vbTab & IIf(IsNull(rs.Fields("Bairro")), "", rs.Fields("Bairro")) & vbTab & IIf(IsNull(rs.Fields("Cidade")), "", rs.Fields("Cidade")) & vbTab & IIf(IsNull(rs.Fields("COD_ORCAMENTO")), "N", "S"))
            If code = rs.Fields("Endereco") Then
                selectedRow = .Rows - 1
            End If
            
            rs.MoveNext
        Loop
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            If .Rows > 9 Then
                .TopRow = selectedRow
            End If
            .col = 0
            .ColSel = .Cols - 1
        Else
        End If
        
    End With
    rs.Close
    Set rs = Nothing
End Sub

Private Sub carregaGridOrcamentos(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    'Dim bkColor As Long
    Dim col As Integer
    
    selectedRow = 2
    With gridOrcamentos
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("CODIGO") & vbTab & IIf(IsNull(rs.Fields("DATA")), "", rs.Fields("DATA")) & vbTab & IIf(IsNull(rs.Fields("NOME_CLI")), "", rs.Fields("NOME_CLI")) & vbTab & rs.Fields("VALOR_GERAL") & vbTab & rs.Fields("STATUS"))
            If code = rs.Fields("CODIGO") Then
                selectedRow = .Rows - 1
            End If
            
''''''
                .row = .Rows - 1
                For col = .FixedCols To .Cols - 1
                    .col = col
                    
                    If rs.Fields("Status") = "Aprovado" Then
                        .CellBackColor = vbGreen
                    End If
                    
                    If rs.Fields("Status") = "Reprovado" Then
                        .CellBackColor = vbRed
                    End If
                    
                    If rs.Fields("Status") = "Enviado" Then
                        .CellBackColor = vbYellow
                    End If
                    
                Next
            
            
''''''
            
            
            rs.MoveNext
        Loop
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            If .Rows > 9 Then
                .TopRow = selectedRow
            End If
            .col = 0
            .ColSel = .Cols - 1
        Else
        End If
        
    End With
    rs.Close
    Set rs = Nothing
    
    strStatusOrcamento = ""
    strStatusDescricao = ""

    
End Sub

Private Sub btnRestaurar_Click()
    
    'If (cboStatus.Text <> "") And Not (strStatusOrcamento = "") Then
        Call carregaGridOrcamentos(Return_Recordset(Return_SqlString(10, txtCodigo.Text)))
    'End If
End Sub

Private Sub cboClientes_Click()

    Call carregaInfoCliente(Return_Recordset(Return_SqlString(2, cboClientes.Text)))
    Call carregaGridEnderecos(Return_Recordset(Return_SqlString(3, lblCodigo.Caption)))
        
    If pblc_strClienteTipo = "PROSPECT" Then
        btnConverterProspect.Visible = True
    Else
        btnConverterProspect.Visible = False
    End If
    
End Sub


Private Sub cboClientesProspects_GotFocus()
    optnCliente.value = 1
End Sub

Private Sub cboStatus_Change()
    If 1 = 2 Then
        MsgBox "asdf"
    End If
End Sub

Private Sub cboStatus_GotFocus()
    strStatusOrcamento = cboStatus.Text
End Sub

Private Sub cboStatus_LostFocus()
    If cboStatus.Text = "" Then
        cboStatus.BackColor = vbWhite
    End If
    
    If cboStatus.Text = "Enviado" Then
        cboStatus.BackColor = vbYellow
    End If
    
    If cboStatus.Text = "Aprovado" Then
        cboStatus.BackColor = vbGreen
    End If
    
    If cboStatus.Text = "Reprovado" Then
        cboStatus.BackColor = vbRed
    End If
    
    'houve alteração no status do orçamento, então salva automaticamente
    If cboStatus.Text <> strStatusOrcamento Then
        Dim rs As ADODB.Recordset
        
        Set rs = Return_Recordset(Return_SqlString(27, lblNumeroOrcamento.Text))
        rs.Fields("STATUS") = cboStatus.Text
        
        rs.Update
        rs.Close
        Set rs = Nothing
        'btnRestaurar_Click
    End If
    
End Sub

Private Sub chkCelular_Click()
    If chkCelular.value = 1 Then
        telCelular = chkCelular.Caption
        strRel1_lblTelefone = chkCelular.Caption
    Else
        telCelular = ""
    End If
End Sub

Private Sub chkComercial_Click()
    If chkComercial.value = 1 Then
        telComercial = chkComercial.Caption
        strRel1_lblTelefone = chkComercial.Caption
    Else
        telComercial = ""
    End If
End Sub

Private Sub chkResidencial_Click()
    If chkResidencial.value = 1 Then
        telResidencial = chkResidencial.Caption
        strRel1_lblTelefone = chkResidencial.Caption
    Else
        telResidencial = ""
    End If
End Sub

Private Sub Cmd_Edita_Click()

    lblNumeroOrcamento.enabled = False
    
    'If lblNumeroOrcamento.Text = "..." Then
    If lblNumeroOrcamento.Text = "..." Then
        MsgBox "Selecione um orçamento na lista."
        gridOrcamentos.SetFocus
        Exit Sub
    End If
    
    Cmd_Edita.enabled = False
    Cmd_Salva.enabled = True
    Cmd_Exclui.Caption = "Cancelar"
    If gridOrcamentos.Rows > 2 Then
        bEditando = True
        HabilitarComponentes (True)
        txtValidade.SetFocus
    Else
        Call limpaTela
        HabilitarComponentes (False)
    End If
    
    'para saber se inserio novas linhas de soluções, garantias, formas de pagto ou observacoes
    qtdSolucoesInicio = gridSolucoes.Rows - 2
    qtdGarantiasInicio = gridGarantias.Rows - 2
    qtdFormasPagtoInicio = gridFormasPagamento.Rows - 2
    qtdObsInicio = gridObservacoes.Rows - 2
    
    txtImagem.Locked = False
    txtImagem2.Locked = False
    
        gridOrcamentos.enabled = False
    
End Sub

Private Sub Cmd_Exclui_Click()
    
    lblNumeroOrcamento.enabled = True

    If Cmd_Exclui.Caption = "Cancelar" Then
        limpaTela
        HabilitarComponentes (False)
        Cmd_Inclui.enabled = True
        Cmd_Exclui.Caption = "Excluir"
        Cmd_Edita.enabled = True
        Cmd_Salva.enabled = False
        optnModelo1.enabled = True
        optnModelo2.enabled = True
        btnImprimir.enabled = True
        gridOrcamentos.enabled = True
        cboStatus.ListIndex = -1
        cboStatus.BackColor = vbWhite
    Else
        'If lblNumeroOrcamento.Text = "..." Then
        If lblNumeroOrcamento.Text = "..." Then
            MsgBox "Selecione um orçamento na lista"
            Exit Sub
        End If
    
        If gridOrcamentos.Rows > 2 Then
         If MsgBox("Confirma a Exclusão do orcamento?", vbYesNo, "Exclusão de Orçamento") = vbYes Then
            
            'Exclui o orçamento, validado
            Call Return_Recordset(Return_SqlString(16, gridOrcamentos.TextMatrix(gridOrcamentos.RowSel, 0))).Delete
            
            'Exclui as soluções, validado
            'Call excluirSolucoes(Return_SqlString(17, lblNumeroOrcamento.Text))
            Call excluirSolucoes(Return_SqlString(17, lblNumeroOrcamento.Text))
            
            'Exclui as garantias
            'Call excluirGarantias(Return_SqlString(18, lblNumeroOrcamento.Text))
            Call excluirGarantias(Return_SqlString(18, lblNumeroOrcamento.Text))

            
            'Exclui as formas de pagamento
            'Call excluirFormasPagto(Return_SqlString(19, lblNumeroOrcamento.Text))
            Call excluirFormasPagto(Return_SqlString(19, lblNumeroOrcamento.Text))

            
            'Exclui as observações
            'Call excluirObservacoes(Return_SqlString(20, lblNumeroOrcamento.Text))
            Call excluirObservacoes(Return_SqlString(20, lblNumeroOrcamento.Text))

            
            Call carregaGridOrcamentos(Return_Recordset(Return_SqlString(10)))
            Call HabilitarComponentes(False)
            
            Call limpaTela
         End If
        Else
            Call limpaTela
            HabilitarComponentes (False)
        End If
    End If
End Sub

Private Sub Cmd_Inclui_Click()
    Dim proximoNumeroOrcamento As Integer
    
    Cmd_Inclui.enabled = False
    Cmd_Edita.enabled = False
    optnModelo1.enabled = False
    optnModelo2.enabled = False
    Cmd_Salva.enabled = True
    btnImprimir.enabled = False
    
    Cmd_Exclui.Caption = "Cancelar"
    
    proximoNumeroOrcamento = retornaProximoNumeroOrcamento(Return_Recordset(Return_SqlString(4)))

    HabilitarComponentes (True)
    txtValidade.SetFocus
    
    Call limpaTela
    
    'lblNumeroOrcamento.Text = proximoNumeroOrcamento
    dtpData.value = Date
    'MsgBox dtpData.value
    lblNumeroOrcamento.Text = proximoNumeroOrcamento
    bEditando = False
    gridOrcamentos.enabled = False
    
    'Garantias padrão
    With gridGarantias
        .AddItem ("01 Ano para o produto pelo fabricante, defeito de fabricação")
        .AddItem ("03  Meses para mão de obra, tubos e conexões")
        .Refresh
    End With
    
    'Formas de pagamento padrão
    With gridFormasPagamento
        .AddItem ("3X sendo 50% - Deposito antecipado e 50% 2 vezes -30/60 dias da data da aprovação")
        .AddItem ("8% desconto - Deposito antecipado")
        .Refresh
    End With
    
    cboStatus.ListIndex = 0
    cboStatus.BackColor = vbYellow
    
    txtDesconto.Text = ""
        
End Sub

Private Sub clonagem()
    Dim proximoNumeroOrcamento As Integer
    globalProximoNumeroOrcamento = retornaProximoNumeroOrcamento(Return_Recordset(Return_SqlString(4)))
    
'--- Tabela Orcamento
    Dim rs As ADODB.Recordset
    
    Set rs = Return_Recordset(Return_SqlString(5))
    rs.AddNew
    
    
    rs.Fields("CODIGO") = globalProximoNumeroOrcamento
    rs.Fields("DATA") = Now
    
    If Trim(lblCodigo.Caption) <> "..." Then
        rs.Fields("COD_CLIENTE") = lblCodigo.Caption
    Else
        rs.Fields("COD_CLIENTE") = "0"
    End If
    
    rs.Fields("NECESSIDADE") = txtNecessidade.Text
    rs.Fields("ENTREGA") = txtEntrega.Text
    rs.Fields("TEMPO_SERVICO") = txtTempoServico.Text
    rs.Fields("VALIDADE") = txtValidade.Text
    If lblSolucoesTotalGeral.Caption <> "..." Then
        rs.Fields("VALOR_GERAL") = CDbl(lblSolucoesTotalGeral.Caption)
    End If
    
    If Trim(txtImagem.Text) <> "" Then
        rs.Fields("IMG_PATH") = txtImagem.Text
    Else
        rs.Fields("IMG_PATH") = ""
    End If
    
    If Trim(txtTextoImg1.Text) <> "" Then
        rs.Fields("IMG_TEXT") = txtTextoImg1.Text
    Else
        rs.Fields("IMG_TEXT") = ""
    End If
    
    If Trim(txtImagem2.Text) <> "" Then
        rs.Fields("IMG2_PATH") = txtImagem2.Text
    Else
        rs.Fields("IMG2_PATH") = ""
    End If
    If Trim(txtTextoImg2.Text) <> "" Then
        rs.Fields("IMG2_TEXT") = txtTextoImg2.Text
    Else
        rs.Fields("IMG2_TEXT") = ""
    End If
    
    If chkCelular.value = 1 Then
        rs.Fields("Fone_Padrao") = "celular"
    ElseIf chkComercial.value = 1 Then
        rs.Fields("Fone_Padrao") = "comercial"
    ElseIf chkResidencial.value = 1 Then
        rs.Fields("Fone_Padrao") = "residencial"
    End If
    
    If Trim(txtDesconto.Text) <> "" Then
        rs.Fields("Desconto") = txtDesconto.Text
    Else
        rs.Fields("Desconto") = "0"
    End If
        
   
    rs.Update
    rs.Close
    Set rs = Nothing

'---

    ClonaSolucoes
    clonaGarantias
    clonaFormasPagamento
    clonaObservacoes
    
    Call carregaGridOrcamentos(Return_Recordset(Return_SqlString(10)))
    'Call salvaOrcamento 'OK
    'Call salvaGarantias 'ok
    'Call salvaSolucoes 'ok
    'Call salvaFormasPagamento 'ok
    'Call salvaObservacoes
    
    
End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String, Optional param3 As String) As String
If queryType = 1 Then
        Return_SqlString = "Select Nome_Cli From Clientes Where Aquivo_morto = 'N' Order by 1"
    ElseIf queryType = 2 Then
        Return_SqlString = "Select Codigo, CPF, Contrato, e_mail, Fone_celular, Fone_res, Fone_com, Tipo From Clientes Where Nome_cli = '" & param1 & "'"
    ElseIf queryType = 3 Then
        Return_SqlString = "Select Endereco, Bairro, Cidade, Cod_orcamento From Endereco Where Cod_Cliente = " & lblCodigo
    ElseIf queryType = 4 Then
        Return_SqlString = "Select max(CODIGO) as MAXCODE from orcamento where CODIGO < 1000"
    ElseIf queryType = 5 Then
        Return_SqlString = "Select * from ORCAMENTO"
    ElseIf queryType = 6 Then
        Return_SqlString = "Select * from ORCAMENTO_SOLUCAO"
    ElseIf queryType = 7 Then
        Return_SqlString = "Select * from ORCAMENTO_GARANTIAS"
    ElseIf queryType = 8 Then
        Return_SqlString = "Select * from ORCAMENTO_FORMAS_PAGTO"
    ElseIf queryType = 9 Then
        Return_SqlString = "Select * from ORCAMENTO_OBS"
    ElseIf queryType = 10 Then
        Return_SqlString = "Select A.CODIGO, A.DATA, B.NOME_CLI, A.VALOR_GERAL, A.STATUS From ORCAMENTO A Inner Join Clientes B on B.codigo = A.COD_CLIENTE Order by A.CODIGO"
    ElseIf queryType = 11 Then
        Return_SqlString = "Select A.DATA, A.CODIGO as codOrcamento, A.VALIDADE, B.NOME_CLI, B.Codigo as codCli, B.CPF, B.Contrato, B.e_mail, B.Fone_celular, B.Fone_Res, B.Fone_Com, B.Tipo, A.Necessidade, A.Entrega, A.Tempo_servico, A.IMG_PATH, A.IMG_TEXT, A.IMG2_PATH, A.IMG2_TEXT, A.Desconto, A.STATUS, A.STATUS_MSG, A.descontotipo From ORCAMENTO A Inner Join Clientes B on B.codigo = A.COD_CLIENTE Where A.CODIGO = " & param1
    ElseIf queryType = 12 Then
        Return_SqlString = "Select Descricao, Quantidade, Valor_Unitario, Valor_Total, Codigo From ORCAMENTO_SOLUCAO Where COD_ORCAMENTO = " & param1
    ElseIf queryType = 13 Then
        Return_SqlString = "Select Garantia_Descricao From ORCAMENTO_GARANTIAS Where COD_ORCAMENTO = " & param1
    ElseIf queryType = 14 Then
        Return_SqlString = "Select FORMA_PAGTO_DESCRICAO From ORCAMENTO_FORMAS_PAGTO Where COD_ORCAMENTO = " & param1
    ElseIf queryType = 15 Then
        Return_SqlString = "Select OBSERVACAO_DESCRICAO From ORCAMENTO_OBS Where COD_ORCAMENTO = " & param1
    ElseIf queryType = 16 Then
        Return_SqlString = "Select * from ORCAMENTO Where CODIGO = " & param1
    ElseIf queryType = 17 Then
        Return_SqlString = "Select * from ORCAMENTO_SOLUCAO Where COD_ORCAMENTO = " & param1
    ElseIf queryType = 18 Then
        Return_SqlString = "Select * from ORCAMENTO_GARANTIAS Where COD_ORCAMENTO = " & param1
    ElseIf queryType = 19 Then
        Return_SqlString = "Select * from ORCAMENTO_FORMAS_PAGTO Where COD_ORCAMENTO = " & param1
    ElseIf queryType = 20 Then
        Return_SqlString = "Select * from ORCAMENTO_OBS Where COD_ORCAMENTO = " & param1
    ElseIf queryType = 21 Then
        Return_SqlString = "Select * From Clientes Where Codigo = " & param1
    ElseIf queryType = 22 Then
        Return_SqlString = "Select Nome_Cli From Clientes Where Aquivo_morto = 'N' And Tipo = 'PROSPECT' Order by 1"
    ElseIf queryType = 23 Then
        Return_SqlString = "Select A.CODIGO, A.DATA, B.NOME_CLI, A.VALOR_GERAL, A.STATUS From ORCAMENTO A Inner Join Clientes B on B.codigo = A.COD_CLIENTE Where B.TIPO = 'CLIENTE' AND B.NOME_CLI = '" & param1 & "'"
    ElseIf queryType = 24 Then
        Return_SqlString = "Select A.CODIGO, A.DATA, B.NOME_CLI, A.VALOR_GERAL From ORCAMENTO A Inner Join Clientes B on B.codigo = A.COD_CLIENTE Where B.TIPO = 'PROSPECT' AND B.NOME_CLI = '" & param1 & "'"
    ElseIf queryType = 25 Then
        Return_SqlString = "Select A.CODIGO, A.DATA, B.NOME_CLI, A.VALOR_GERAL, A.STATUS From ORCAMENTO A Inner Join Clientes B on B.codigo = A.COD_CLIENTE Where Format(A.DATA, 'Short Date') >= '" & dtpDataDe.value & "' AND Format(A.DATA,'Short Date') <= '" & dtpDataAte.value & "'"
    ElseIf queryType = 26 Then
        Return_SqlString = "Select A.CODIGO, A.DATA, B.NOME_CLI, A.VALOR_GERAL, A.STATUS From ORCAMENTO A Inner Join Clientes B on B.codigo = A.COD_CLIENTE Where A.CODIGO = " & param1
    ElseIf queryType = 27 Then
        Return_SqlString = "Select * from ORCAMENTO where codigo = " & param1
    ElseIf queryType = 28 Then
        Return_SqlString = "Select Fone_Padrao From ORCAMENTO Where CODIGO = " & param1
    End If
End Function

Private Sub Arquivos_Dados()
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
    dbname = Path + FileName
End Sub

Private Function excluirSolucoes(strSql As String) As ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim qtd As Integer
    Dim X As Integer
    
    Call Arquivos_Dados
    
    rs2.Open strSql, conn, adOpenStatic, adLockOptimistic
    qtd = rs2.RecordCount
    
    For X = 1 To qtd
        rs2.Delete
        If Not rs2.EOF Then
            rs2.MoveNext
        End If
    Next X
    
    Set excluirSolucoes = rs2
End Function

Private Function excluirGarantias(strSql As String) As ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim qtd As Integer
    Dim X As Integer
    
    Call Arquivos_Dados
    
    rs2.Open strSql, conn, adOpenStatic, adLockOptimistic
    qtd = rs2.RecordCount
    
    For X = 1 To qtd
        rs2.Delete
        If Not rs2.EOF Then
            rs2.MoveNext
        End If
    Next X
    
    Set excluirGarantias = rs2
End Function

Private Function excluirFormasPagto(strSql As String) As ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim qtd As Integer
    Dim X As Integer
    
    Call Arquivos_Dados
    
    rs2.Open strSql, conn, adOpenStatic, adLockOptimistic
    qtd = rs2.RecordCount
    
    For X = 1 To qtd
        rs2.Delete
        If Not rs2.EOF Then
            rs2.MoveNext
        End If
    Next X
    
    Set excluirFormasPagto = rs2
End Function

Private Function excluirObservacoes(strSql As String) As ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim qtd As Integer
    Dim X As Integer
    
    Call Arquivos_Dados
    
    rs2.Open strSql, conn, adOpenStatic, adLockOptimistic
    qtd = rs2.RecordCount
    
    For X = 1 To qtd
        rs2.Delete
        If Not rs2.EOF Then
            rs2.MoveNext
        End If
    Next X
    
    Set excluirObservacoes = rs2
End Function


Private Function Return_Recordset(strSql As String) As ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim qtd As Integer
    
    Call Arquivos_Dados
    
    rs2.Open strSql, conn, adOpenStatic, adLockOptimistic
    qtd = rs2.RecordCount
    
    Set Return_Recordset = rs2
End Function

Private Sub carregaCboClientes(rs As ADODB.Recordset, Optional code As String)
    
    With cboClientes
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("Nome_Cli"))
            rs.MoveNext
        Loop
    End With
    
    rs.Close
    Set rs = Nothing
    
End Sub

Private Sub carregaCboClientesProspects(rs As ADODB.Recordset, Optional code As String)
    
    With cboClientesProspects
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("Nome_Cli"))
            rs.MoveNext
        Loop
    End With
    
    rs.Close
    Set rs = Nothing
    
End Sub


Private Sub ConfigureGridEnderecos()
    With gridEnderecos
        .Cols = 4
        .ColWidth(0) = 4500  'Endreço
        .ColWidth(1) = 3300 'Bairro
        .ColWidth(2) = 3280  'Cidade
        .ColWidth(3) = 500  'Usar
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Endreço"
        .TextMatrix(0, 1) = "Bairro"
        .TextMatrix(0, 2) = "Cidade"
        .TextMatrix(0, 3) = "Usar"
        
        .RowHeight(1) = 0
    End With
End Sub


Private Sub ConfigureGridSolucoes()
    Dim nomepc As String
    nomepc = G_NomeComputador
    With gridSolucoes
        .Cols = 6
        
        If nomepc = "CLEITON-PC" Then
            .ColWidth(0) = 500  'Item
            .ColWidth(1) = 4400  'Descricao
            .ColWidth(2) = 700 'Qtde.
            .ColWidth(3) = 1000  'Val. unitário
            .ColWidth(4) = 1000  'Val. total
        Else
            .ColWidth(0) = 500   'Item
            .ColWidth(1) = 3950  'Descricao
            .ColWidth(2) = 700   'Qtde.
            .ColWidth(3) = 1000  'Val. unitário
            .ColWidth(4) = 1000  'Val. total
        End If
        
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Item"
        .TextMatrix(0, 1) = "Descrição"
        .TextMatrix(0, 2) = "Qtde."
        .TextMatrix(0, 3) = "Val. unitário"
        .TextMatrix(0, 4) = "Val. total"
        .TextMatrix(0, 5) = "Código"
        .RowHeight(1) = 0
    End With
End Sub

Private Sub ConfigureGridGarantias()
    With gridGarantias
        .Cols = 1
        .ColWidth(0) = 5650 'Descricao
        .ColAlignment(0) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Descrição"
        .RowHeight(1) = 0
    End With
End Sub

Private Sub ConfigureGridFormasPagamento()
    Dim nomepc As String
    nomepc = G_NomeComputador
    
    With gridFormasPagamento
        .Cols = 1
        If nomepc = "CLEITON-PC" Then
            .ColWidth(0) = 8370 'Descricao
        Else
            .ColWidth(0) = 5370 'Descricao
        End If
        
        .ColAlignment(0) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Descrição"
        .RowHeight(1) = 0
    End With
End Sub

Private Sub ConfigureGridObservacoes()
    With gridObservacoes
        .Cols = 1
        .ColWidth(0) = 10200 'Descricao
        .ColAlignment(0) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Descrição"
        .RowHeight(1) = 0
    End With
End Sub

Private Sub ConfigureGridOrcamentos()
    With gridOrcamentos
        .Cols = 5
        .ColWidth(0) = 800  'CODIGO
        .ColWidth(1) = 1200 'DATA
        .ColWidth(2) = 3000 'COD_CLIENTE
        .ColWidth(3) = 830 'VALOR_GERAL
        .ColWidth(4) = 1000 'STATUS
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Data"
        .TextMatrix(0, 2) = "Cliente"
        .TextMatrix(0, 3) = "Valor"
        .TextMatrix(0, 4) = "Status"
        .RowHeight(1) = 0
    End With
End Sub

Private Function consistirInformacoes() As Boolean
    
    consistirInformacoes = True
    
    If Trim(txtValidade.Text) = "" Then
        MsgBox ("Informe a validade do orçamento.")
        txtValidade.SetFocus
        consistirInformacoes = False
    ElseIf Trim(lblNumeroOrcamento.Text) = "" Then
        MsgBox ("Informe o número do orçamento")
        lblNumeroOrcamento.SetFocus
        consistirInformacoes = False
    ElseIf Trim(lblNumeroOrcamento.Text) = "0" Then
        MsgBox ("Informe o número do orçamento")
        lblNumeroOrcamento.SetFocus
        consistirInformacoes = False
    ElseIf Trim(cboClientes.Text) = "" Then
        MsgBox ("Selecione um cliente")
        cboClientes.SetFocus
        consistirInformacoes = False
    ElseIf chkCelular.value = False And chkResidencial.value = False And chkComercial.value = False Then
        MsgBox ("Selecione ao menos um telefone")
        chkCelular.SetFocus
        consistirInformacoes = False
    ElseIf Trim(txtNecessidade.Text) = "" Then
        MsgBox ("Informe a necessidade")
        txtNecessidade.SetFocus
        consistirInformacoes = False
    ElseIf gridSolucoes.Rows = 2 Then
        MsgBox ("Informe ao menos uma solução")
        btnAdicionarSolucao.SetFocus
        consistirInformacoes = False
    ElseIf Trim(txtEntrega.Text) = "" Then
        MsgBox ("Informe os dados de entrega")
        txtEntrega.SetFocus
        consistirInformacoes = False
    ElseIf Trim(txtTempoServico.Text) = "" Then
        MsgBox ("Informe o tempo de serviço")
        txtTempoServico.SetFocus
        consistirInformacoes = False
    ElseIf gridGarantias.Rows = 2 Then
        MsgBox ("Informe ao menos 1 garantia")
        btnAdicionarGarantia.SetFocus
        consistirInformacoes = False
    ElseIf gridFormasPagamento.Rows = 2 Then
        MsgBox ("Informe ao menos 1 forma de pagamento")
        btnAdicionarFormaPagamento.SetFocus
        consistirInformacoes = False
    ElseIf (jaExisteOrcamento(lblNumeroOrcamento.Text)) And (bEditando = False) Then
        MsgBox ("O número " & lblNumeroOrcamento.Text & " já foi utilizado, informe outro")
        lblNumeroOrcamento.SetFocus
        consistirInformacoes = False
    End If
    
    If consistirInformacoes = False Then
        Exit Function
    End If
    
    consistirInformacoes = True
End Function

Private Sub Read_Config_File()
    ConfigFile = App.Path & "\CONFIG.INI"
    
    Path = RTrim(GetINISetting("DATABASE", "PATH", ConfigFile))
    If Right(Path, 1) <> "\" Then
       Path = Path + "\"
    End If
    FileName = GetINISetting("DATABASE", "FILENAME", ConfigFile)
    
    ImageFolder = GetINISetting("IMAGE", "FOLDER", ConfigFile)
End Sub

Private Function jaExisteOrcamento(parNum As String) As Boolean
    
    Dim db As DAO.Database
    Dim rsNumero As DAO.Recordset
    
    Call Read_Config_File
    'Text_DatabaseFolder = Path
    'Text_DatabaseName = FileName
    'Text_ImageFolder = ImageFolder
    
    
    
    Set db = OpenDatabase(Path & "dados.mdb")
    Set rsNumero = db.OpenRecordset("Select * from orcamento where codigo = " & parNum)
    
    'MsgBox rsNumero.RecordCount
    
    'ver aqui se o endereço já existe para algum cliente
    'If Return_Recordset2(Return_SqlString(6, Text_End.Text)).RecordCount > 0 Then
    '     frmEnderecoExistente.Show (1)
    'End If
    
    If rsNumero.RecordCount > 0 Then
        jaExisteOrcamento = True
    Else
        jaExisteOrcamento = False
    End If
    
    
    
End Function


Private Sub salvaOrcamento()
    Dim rs As ADODB.Recordset
    
    
    If bEditando = False Then 'está inserindo
        Set rs = Return_Recordset(Return_SqlString(5))
        rs.AddNew
    Else 'está editando
        Set rs = Return_Recordset(Return_SqlString(27, lblNumeroOrcamento.Text))
    End If
    
    
    If Not bEditando Then
        rs.Fields("CODIGO") = Int(lblNumeroOrcamento.Text)
    End If
    
    
    rs.Fields("DATA") = dtpData.value
    
    If Trim(lblCodigo.Caption) <> "..." Then
        rs.Fields("COD_CLIENTE") = lblCodigo.Caption
    Else
        rs.Fields("COD_CLIENTE") = "0"
    End If
    
    rs.Fields("NECESSIDADE") = txtNecessidade.Text
    rs.Fields("ENTREGA") = txtEntrega.Text
    rs.Fields("TEMPO_SERVICO") = txtTempoServico.Text
    rs.Fields("VALIDADE") = txtValidade.Text
    If lblSolucoesTotalGeral.Caption <> "..." Then
        rs.Fields("VALOR_GERAL") = CDbl(lblSolucoesTotalGeral.Caption)
    End If
    
    If Trim(txtImagem.Text) <> "" Then
        rs.Fields("IMG_PATH") = txtImagem.Text
    Else
        rs.Fields("IMG_PATH") = ""
    End If
    
    If Trim(txtTextoImg1.Text) <> "" Then
        rs.Fields("IMG_TEXT") = txtTextoImg1.Text
    Else
        rs.Fields("IMG_TEXT") = ""
    End If
    
    If Trim(txtImagem2.Text) <> "" Then
        rs.Fields("IMG2_PATH") = txtImagem2.Text
    Else
        rs.Fields("IMG2_PATH") = ""
    End If
    If Trim(txtTextoImg2.Text) <> "" Then
        rs.Fields("IMG2_TEXT") = txtTextoImg2.Text
    Else
        rs.Fields("IMG2_TEXT") = ""
    End If
    
    If chkCelular.value = 1 Then
        rs.Fields("Fone_Padrao") = "celular"
    ElseIf chkComercial.value = 1 Then
        rs.Fields("Fone_Padrao") = "comercial"
    ElseIf chkResidencial.value = 1 Then
        rs.Fields("Fone_Padrao") = "residencial"
    End If
    
    If Trim(txtDesconto.Text) <> "" Then
        rs.Fields("Desconto") = txtDesconto.Text
        'MsgBox Mid(Frame13.Caption, Len(Frame13.Caption), 1)
        If Mid(Frame13.Caption, Len(Frame13.Caption), 1) = "%" Then
            rs.Fields("descontotipo") = "porcento"
        Else
            rs.Fields("descontotipo") = "valor"
        End If
    Else
        rs.Fields("Desconto") = "0"
    End If
    
    
    
    
    rs.Fields("STATUS") = cboStatus.Text
    rs.Fields("STATUS_MSG") = txtStatusMensagem.Text
    
    If descontoPorPorcentagem Then
    
    Else
    
    End If
        
    Call gravaOrcamento(rs)
    
    rs.Update
    rs.Close
    Set rs = Nothing
End Sub

Private Sub converteProspect()
    Dim rs As ADODB.Recordset
    Set rs = Return_Recordset(Return_SqlString(21, lblCodigo.Caption))
    
    rs.Fields("TIPO") = "CLIENTE"
    
    'Call gravaOrcamento(rs)
    Call gravaConverteProspect(rs)
    
    rs.Update
    rs.Close
    Set rs = Nothing
    
    MsgBox "O prospect foi convertido para cliente, vá até o seu cadastro para completar seus dados!"
End Sub

Private Sub salvaSolucoes()
    Dim contador As Integer
    Dim rsx0 As ADODB.Recordset
    
    If bEditando = False Then 'Está inserindo
        For contador = 2 To gridSolucoes.Rows - 1
            Set rsx0 = Return_Recordset(Return_SqlString(6))
            rsx0.AddNew
            'rsx0.Fields("COD_ORCAMENTO") = lblNumeroOrcamento.Text
            rsx0.Fields("COD_ORCAMENTO") = lblNumeroOrcamento.Text
            rsx0.Fields("DESCRICAO") = gridSolucoes.TextMatrix(contador, 1)
            rsx0.Fields("QUANTIDADE") = CDbl(gridSolucoes.TextMatrix(contador, 2))
            rsx0.Fields("VALOR_UNITARIO") = CDbl(gridSolucoes.TextMatrix(contador, 3))
            rsx0.Fields("VALOR_TOTAL") = CDbl(gridSolucoes.TextMatrix(contador, 4))

            rsx0.Update
            rsx0.Close
            Set rsx0 = Nothing
        Next contador
    End If
    
    If bEditando Then 'Está editando
        If (gridSolucoes.Rows - 2) < qtdSolucoesInicio Then 'se estiver removendo linhas do grid
            'Call excluirSolucoes(Return_SqlString(17, lblNumeroOrcamento.Text))    'exclui todas as solucoes do orcamento
            Call excluirSolucoes(Return_SqlString(17, lblNumeroOrcamento.Text))    'exclui todas as solucoes do orcamento
            Call salvaSolucoes_2                                                      'cadastra novamente as solucoes do orçamento corrente
            qtdSolucoesInicio = qtdSolucoesInicio - 1
            Exit Sub
        ElseIf (gridSolucoes.Rows - 2) > qtdSolucoesInicio Then 'se estiver incluindo linhas no grid
            'Call excluirSolucoes(Return_SqlString(17, lblNumeroOrcamento.Text))    'exclui todas as solucoes do orcamento
            Call excluirSolucoes(Return_SqlString(17, lblNumeroOrcamento.Text))    'exclui todas as solucoes do orcamento
            Call salvaSolucoes_2                                                      'cadastra novamente as solucoes do orçamento corrente
            Exit Sub
            qtdSolucoesInicio = qtdSolucoesInicio + 1
        ElseIf (gridSolucoes.Rows - 2) = qtdSolucoesInicio Then 'se as linhas iniciais e finais estiverem iguais
            'Call excluirSolucoes(Return_SqlString(17, lblNumeroOrcamento.Text))    'exclui todas as solucoes do orcamento
            Call excluirSolucoes(Return_SqlString(17, lblNumeroOrcamento.Text))    'exclui todas as solucoes do orcamento
            Call salvaSolucoes_2                                                      'cadastra novamente as solucoes do orçamento corrente
            Exit Sub
        End If
    End If
    
End Sub

Private Sub ClonaSolucoes()
    Dim contador As Integer
    Dim rsx0 As ADODB.Recordset
    
    For contador = 2 To gridSolucoes.Rows - 1
        Set rsx0 = Return_Recordset(Return_SqlString(6))
        rsx0.AddNew
        rsx0.Fields("COD_ORCAMENTO") = globalProximoNumeroOrcamento
        rsx0.Fields("DESCRICAO") = gridSolucoes.TextMatrix(contador, 1)
        rsx0.Fields("QUANTIDADE") = CDbl(gridSolucoes.TextMatrix(contador, 2))
        rsx0.Fields("VALOR_UNITARIO") = CDbl(gridSolucoes.TextMatrix(contador, 3))
        rsx0.Fields("VALOR_TOTAL") = CDbl(gridSolucoes.TextMatrix(contador, 4))
        rsx0.Update
        rsx0.Close
        Set rsx0 = Nothing
    Next contador
    
    
End Sub

Private Sub salvaSolucoes_2()
    Dim contador As Integer
    Dim rsx As ADODB.Recordset
    Set rsx = Return_Recordset(Return_SqlString(6))
    
    For contador = 2 To gridSolucoes.Rows - 1
        rsx.AddNew
        
        rsx.Fields("COD_ORCAMENTO") = lblNumeroOrcamento.Text
        rsx.Fields("DESCRICAO") = gridSolucoes.TextMatrix(contador, 0)
        rsx.Fields("QUANTIDADE") = CDbl(gridSolucoes.TextMatrix(contador, 2))
        rsx.Fields("VALOR_UNITARIO") = CDbl(gridSolucoes.TextMatrix(contador, 3))
        rsx.Fields("VALOR_TOTAL") = CDbl(gridSolucoes.TextMatrix(contador, 4))

        rsx.Fields("COD_ORCAMENTO").value = lblNumeroOrcamento.Text
        rsx.Fields("DESCRICAO").value = gridSolucoes.TextMatrix(contador, 1)
        rsx.Fields("QUANTIDADE").value = gridSolucoes.TextMatrix(contador, 2)
        rsx.Fields("VALOR_UNITARIO").value = gridSolucoes.TextMatrix(contador, 3)
        rsx.Fields("VALOR_TOTAL").value = gridSolucoes.TextMatrix(contador, 4)
        
        rsx.Update
    Next contador
    
    rsx.Close
    Set rsx = Nothing
End Sub

Private Sub salvaGarantias_2()
    Dim contador As Integer
    Dim rsx As ADODB.Recordset
    Set rsx = Return_Recordset(Return_SqlString(7))
    
    For contador = 2 To gridGarantias.Rows - 2
        rsx.AddNew
        
        rsx.Fields("COD_ORCAMENTO") = lblNumeroOrcamento.Text
        rsx.Fields("GARANTIA_DESCRICAO") = gridGarantias.TextMatrix(contador, 0)
        
        rsx.Fields("COD_ORCAMENTO").value = lblNumeroOrcamento.Text
        rsx.Fields("GARANTIA_DESCRICAO").value = gridGarantias.TextMatrix(contador, 0)
        
        rsx.Update
    Next contador
    
    rsx.Close
    Set rsx = Nothing
End Sub

Private Sub editaGarantias()
    Dim contador As Integer
    Dim rsx As ADODB.Recordset
    Set rsx = Return_Recordset(Return_SqlString(7))
    
    For contador = 2 To gridGarantias.Rows - 1
        rsx.AddNew
        
        rsx.Fields("COD_ORCAMENTO") = lblNumeroOrcamento.Text
        rsx.Fields("GARANTIA_DESCRICAO") = gridGarantias.TextMatrix(contador, 0)
        
        rsx.Fields("COD_ORCAMENTO").value = lblNumeroOrcamento.Text
        rsx.Fields("GARANTIA_DESCRICAO").value = gridGarantias.TextMatrix(contador, 0)
        
        rsx.Update
    Next contador
    
    rsx.Close
    Set rsx = Nothing
End Sub

Private Sub editaFormasPagto()
    Dim contador As Integer
    Dim rsx As ADODB.Recordset
    Set rsx = Return_Recordset(Return_SqlString(8))
    
    For contador = 2 To gridFormasPagamento.Rows - 1
        rsx.AddNew
        
        rsx.Fields("COD_ORCAMENTO") = lblNumeroOrcamento.Text
        rsx.Fields("FORMA_PAGTO_DESCRICAO") = gridFormasPagamento.TextMatrix(contador, 0)
        
        rsx.Fields("COD_ORCAMENTO").value = lblNumeroOrcamento.Text
        rsx.Fields("FORMA_PAGTO_DESCRICAO").value = gridFormasPagamento.TextMatrix(contador, 0)
        
        rsx.Update
    Next contador
    
    rsx.Close
    Set rsx = Nothing
End Sub

Private Sub editaObservacoes()
    Dim contador As Integer
    Dim rsx As ADODB.Recordset
    Set rsx = Return_Recordset(Return_SqlString(9))
    
    For contador = 2 To gridObservacoes.Rows - 1
        rsx.AddNew
        
        rsx.Fields("COD_ORCAMENTO") = lblNumeroOrcamento.Text
        rsx.Fields("OBSERVACAO_DESCRICAO") = gridObservacoes.TextMatrix(contador, 0)
        
        rsx.Fields("COD_ORCAMENTO").value = lblNumeroOrcamento.Text
        rsx.Fields("OBSERVACAO_DESCRICAO").value = gridObservacoes.TextMatrix(contador, 0)
        
        rsx.Update
    Next contador
    
    rsx.Close
    Set rsx = Nothing
End Sub


Private Sub salvaEnderecos()
    Dim contador As Integer
    Dim rs As ADODB.Recordset
    Set rs = Return_Recordset(Return_SqlString(3))
    
    For contador = 2 To rs.RecordCount + 2
          
        If contador = rs.RecordCount + 2 Then
            Exit For
        End If
      
        If gridEnderecos.TextMatrix(contador, 3) = "S" Then
            'rs.Fields("COD_ORCAMENTO") = lblNumeroOrcamento.Text
            rs.Fields("COD_ORCAMENTO") = lblNumeroOrcamento.Text
        Else
            rs.Fields("COD_ORCAMENTO") = Null
        End If

        rs.Update
        If Not rs.EOF Then
            rs.MoveNext
        Else
            Exit For
        End If
    Next contador
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub clonaGarantias()
    Dim contador As Integer
    Dim rsx0 As ADODB.Recordset
    
    For contador = 1 To gridGarantias.Rows - 2
        Set rsx0 = Return_Recordset(Return_SqlString(7))
        rsx0.AddNew
        
        rsx0.Fields("COD_ORCAMENTO") = globalProximoNumeroOrcamento
        rsx0.Fields("GARANTIA_DESCRICAO") = gridGarantias.TextMatrix(contador + 1, 0)

        rsx0.Update
        rsx0.Close
        Set rsx0 = Nothing
    Next contador
    
End Sub


Private Sub salvaGarantias()
    Dim contador As Integer
    Dim rsx0 As ADODB.Recordset
    
    If bEditando = False Then 'Está inserindo
        For contador = 2 To gridGarantias.Rows - 1
            Set rsx0 = Return_Recordset(Return_SqlString(7))
            rsx0.AddNew
            
            'rsx0.Fields("COD_ORCAMENTO") = lblNumeroOrcamento.Text
            rsx0.Fields("COD_ORCAMENTO") = lblNumeroOrcamento.Text
            rsx0.Fields("GARANTIA_DESCRICAO") = gridGarantias.TextMatrix(contador, 0)
    
            'rsx0.Fields("COD_ORCAMENTO").value = lblNumeroOrcamento.Text
            rsx0.Fields("COD_ORCAMENTO").value = lblNumeroOrcamento.Text
            rsx0.Fields("GARANTIA_DESCRICAO").value = gridGarantias.TextMatrix(contador, 0)
        
            rsx0.Update
            rsx0.Close
            Set rsx0 = Nothing
        Next contador
    End If
    
    If bEditando Then 'Está editando
        If (gridGarantias.Rows - 2) < qtdGarantiasInicio Then 'se estiver removendo linhas do grid
            'Call excluirGarantias(Return_SqlString(18, lblNumeroOrcamento.Text))    'exclui todas as solucoes do orcamento
            Call excluirGarantias(Return_SqlString(18, lblNumeroOrcamento.Text))    'exclui todas as solucoes do orcamento
            Call editaGarantias
            qtdGarantiasInicio = qtdGarantiasInicio - 1
            Exit Sub
        ElseIf (gridGarantias.Rows - 2) > qtdGarantiasInicio Then 'se estiver incluindo linhas no grid
            'Call excluirGarantias(Return_SqlString(18, lblNumeroOrcamento.Text))    'exclui todas as solucoes do orcamento
            Call excluirGarantias(Return_SqlString(18, lblNumeroOrcamento.Text))    'exclui todas as solucoes do orcamento
            Call editaGarantias 'cadastra novamente as solucoes do orçamento corrente
            qtdGarantiasInicio = qtdGarantiasInicio + 1
            Exit Sub
        ElseIf (gridGarantias.Rows - 2) = qtdGarantiasInicio Then 'se as linhas iniciais e finais estiverem iguais
            'Call excluirGarantias(Return_SqlString(18, lblNumeroOrcamento.Text))    'exclui todas as solucoes do orcamento
            Call excluirGarantias(Return_SqlString(18, lblNumeroOrcamento.Text))    'exclui todas as solucoes do orcamento
            Call editaGarantias 'cadastra novamente as solucoes do orçamento corrente
            Exit Sub
        End If
    End If
End Sub

Private Sub salvaFormasPagamento()
    Dim contador As Integer
    Dim rsx0 As ADODB.Recordset
    
    If bEditando = False Then 'Está inserindo
        For contador = 2 To gridFormasPagamento.Rows - 1
            Set rsx0 = Return_Recordset(Return_SqlString(8))
            rsx0.AddNew
            
            'rsx0.Fields("COD_ORCAMENTO") = lblNumeroOrcamento.Text
            rsx0.Fields("COD_ORCAMENTO") = lblNumeroOrcamento.Text
            rsx0.Fields("FORMA_PAGTO_DESCRICAO") = gridFormasPagamento.TextMatrix(contador, 0)
    
            'rsx0.Fields("COD_ORCAMENTO").value = lblNumeroOrcamento.Text
            rsx0.Fields("COD_ORCAMENTO").value = lblNumeroOrcamento.Text
            rsx0.Fields("FORMA_PAGTO_DESCRICAO").value = gridFormasPagamento.TextMatrix(contador, 0)
        
            rsx0.Update
            rsx0.Close
            Set rsx0 = Nothing
        Next contador
    End If
    
    If bEditando Then 'Está editando
        If (gridFormasPagamento.Rows - 2) < qtdFormasPagtoInicio Then 'se estiver removendo linhas do grid
            'Call excluirFormasPagto(Return_SqlString(19, lblNumeroOrcamento.Text))    'exclui todas as solucoes do orcamento
            Call excluirFormasPagto(Return_SqlString(19, lblNumeroOrcamento.Text))    'exclui todas as solucoes do orcamento
            Call editaFormasPagto
            qtdFormasPagtoInicio = qtdFormasPagtoInicio - 1
            Exit Sub
        ElseIf (gridFormasPagamento.Rows - 2) > qtdFormasPagtoInicio Then 'se estiver incluindo linhas no grid
            Call excluirFormasPagto(Return_SqlString(19, lblNumeroOrcamento.Text))    'exclui todas as solucoes do orcamento
            Call editaFormasPagto 'cadastra novamente as solucoes do orçamento corrente
            qtdFormasPagtoInicio = qtdFormasPagtoInicio + 1
            Exit Sub
        ElseIf (gridFormasPagamento.Rows - 2) = qtdFormasPagtoInicio Then 'se as linhas iniciais e finais estiverem iguais
            Call excluirFormasPagto(Return_SqlString(19, lblNumeroOrcamento.Text))    'exclui todas as solucoes do orcamento
            Call editaFormasPagto 'cadastra novamente as solucoes do orçamento corrente
            Exit Sub
        End If
    End If
    

End Sub

Private Sub clonaFormasPagamento()
    Dim contador As Integer
    Dim rsx0 As ADODB.Recordset
    

    For contador = 1 To gridFormasPagamento.Rows - 2
        Set rsx0 = Return_Recordset(Return_SqlString(8))
        rsx0.AddNew
        
        rsx0.Fields("COD_ORCAMENTO") = globalProximoNumeroOrcamento
        rsx0.Fields("FORMA_PAGTO_DESCRICAO") = gridFormasPagamento.TextMatrix(contador + 1, 0)

        rsx0.Update
        rsx0.Close
        Set rsx0 = Nothing
    Next contador
End Sub


Private Sub salvaObservacoes()
'-------------
    Dim contador As Integer
    Dim rsx3 As ADODB.Recordset
    
    If bEditando = False Then 'Está inserindo
        For contador = 2 To gridObservacoes.Rows - 1
            Set rsx3 = Return_Recordset(Return_SqlString(9))
            rsx3.AddNew
            
            rsx3.Fields("COD_ORCAMENTO") = lblNumeroOrcamento.Text
            rsx3.Fields("OBSERVACAO_DESCRICAO") = gridObservacoes.TextMatrix(contador, 0)
    
            rsx3.Fields("COD_ORCAMENTO").value = lblNumeroOrcamento.Text
            rsx3.Fields("OBSERVACAO_DESCRICAO").value = gridObservacoes.TextMatrix(contador, 0)
        
            rsx3.Update
            rsx3.Close
            Set rsx3 = Nothing
        Next contador
    End If
    
    If bEditando Then 'Está editando
        If (gridObservacoes.Rows - 2) < qtdObsInicio Then 'se estiver removendo linhas do grid
            Call excluirObservacoes(Return_SqlString(20, lblNumeroOrcamento.Text))    'exclui todas as observacoes do orcamento
            Call editaObservacoes
            qtdObsInicio = qtdObsInicio - 1
            Exit Sub
        ElseIf (gridObservacoes.Rows - 2) > qtdObsInicio Then 'se estiver incluindo linhas no grid
            Call excluirObservacoes(Return_SqlString(20, lblNumeroOrcamento.Text))
            Call editaObservacoes 'cadastra novamente as observacoes do orçamento corrente
            qtdObsInicio = qtdObsInicio + 1
            Exit Sub
        ElseIf (gridObservacoes.Rows - 2) = qtdObsInicio Then 'se as linhas iniciais e finais estiverem iguais
            Call excluirObservacoes(Return_SqlString(20, lblNumeroOrcamento.Text))
            Call editaObservacoes 'cadastra novamente as observacoes do orçamento corrente
            Exit Sub
        End If
    End If
End Sub

Private Sub clonaObservacoes()

    Dim contador As Integer
    Dim rsx3 As ADODB.Recordset

    For contador = 1 To gridObservacoes.Rows - 2
        Set rsx3 = Return_Recordset(Return_SqlString(9))
        rsx3.AddNew
        
        rsx3.Fields("COD_ORCAMENTO") = globalProximoNumeroOrcamento
        rsx3.Fields("OBSERVACAO_DESCRICAO") = gridObservacoes.TextMatrix(contador + 1, 0)

        rsx3.Update
        rsx3.Close
        Set rsx3 = Nothing
    Next contador

    
End Sub


Private Sub Cmd_Salva_Click()
    If Not consistirInformacoes Then
        Exit Sub
    End If
    
    
    Call salvaEnderecos 'ok
    Call salvaSolucoes 'ok
    Call salvaGarantias 'ok
    Call salvaFormasPagamento 'ok
    Call salvaObservacoes 'ok
    
    Call salvaOrcamento 'ok
    
    MsgBox "Orçamento gravado.", vbOKOnly, "xSol Aquecedores"
    
    Call limpaTela
    Call HabilitarComponentes(False)
    Cmd_Salva.enabled = False
    Cmd_Inclui.enabled = True
    Cmd_Edita.enabled = True
    optnModelo1.enabled = True
    optnModelo2.enabled = True
    btnImprimir.enabled = True
    Cmd_Exclui.Caption = "Excluir"
    gridOrcamentos.enabled = True
    Call carregaGridOrcamentos(Return_Recordset(Return_SqlString(10)))
    
    txtImagem.Locked = True
    txtImagem2.Locked = True
    
    
End Sub

Private Sub gravaOrcamento(rs As ADODB.Recordset)
   If Not bEditando Then
    rs.Fields("CODIGO").value = Int(lblNumeroOrcamento.Text)
   End If
   
   
   rs.Fields("DATA").value = dtpData.value
   rs.Fields("COD_CLIENTE").value = lblCodigo.Caption
   rs.Fields("NECESSIDADE").value = txtNecessidade.Text
   rs.Fields("ENTREGA").value = txtEntrega.Text
   rs.Fields("TEMPO_SERVICO").value = txtTempoServico.Text
   rs.Fields("VALIDADE").value = txtValidade.Text
   If lblSolucoesTotalGeral.Caption <> "..." Then
    rs.Fields("VALOR_GERAL").value = lblSolucoesTotalGeral.Caption
   End If
   
    If Trim(txtImagem.Text) <> "" Then
        rs.Fields("IMG_PATH").value = txtImagem.Text
    End If
    If Trim(txtTextoImg1.Text) <> "" Then
        rs.Fields("IMG_TEXT").value = txtTextoImg1.Text
    End If
   
   
    If Trim(txtImagem2.Text) <> "" Then
        rs.Fields("IMG2_PATH").value = txtImagem2.Text
    End If
    If Trim(txtTextoImg2.Text) <> "" Then
        rs.Fields("IMG2_TEXT").value = txtTextoImg2.Text
    End If
   
    If Trim(txtDesconto.Text) <> "" Then
        rs.Fields("Desconto") = txtDesconto.Text
    End If
   
End Sub


Private Sub gravaConverteProspect(rs As ADODB.Recordset)
    rs.Fields("TIPO").value = "CLIENTE"
End Sub


Private Sub Command1_Click()
    Dim strValor As String
    Dim strValorFim As String
    Dim dblValor As Double
    
    strValor = "-10"
    dblValor = CDbl(strValor)
    
    strValorFim = Str(dblValor)
    
    
End Sub

Private Sub dtpDataAte_GotFocus()
    pblc_strTipoPesquisa = "DATA"
End Sub

Private Sub dtpDataDe_GotFocus()
    pblc_strTipoPesquisa = "DATA"
End Sub

Private Sub Form_Activate()
    
ajustaDisplayPcCleiton

   
    
    If (MDIForm_Xsol.acessibilidade_TextoNegrito = True) Or (MDIForm_Xsol.acessibilidade_TextoNegrito = False) Then
txtCodigo.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtTextoImg2.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtTextoImg1.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtImagem2.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtImagem.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtValidade.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtTempoServico.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtEntrega.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtNecessidade.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
cboClientesProspects.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
cboClientes.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
gridObservacoes.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
gridFormasPagamento.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
gridGarantias.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
gridOrcamentos.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
gridSolucoes.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
gridEnderecos.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
dtpDataDe.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
dtpDataAte.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
dtpData.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    End If
End Sub

Private Sub ajustaDisplayPcCleiton()
    Dim nomepc As String
    nomepc = G_NomeComputador
    
    'tela do cleiton, monitor
    
    If nomepc = "CLEITON-PC" Then
        frmOrcamento.Height = 11820
        frmOrcamento.Width = 24225
                
        Frame1.Width = 23640
        Frame1.Height = 11415
        
        Frame2.Width = 15555
        
        Frame3.Width = 6510
        
        Frame4.Width = 9495
        Frame4.Left = 6855
                
        Frame5.Top = 10230
        
        Frame6.Top = 690
        Frame6.Left = 15945
        Frame6.Height = 7755
        Frame6.Width = 7680
        
        Frame7.Top = 10230
                
        Frame8.Top = 5940
        Frame8.Width = 6405
                
        Frame9.Top = 5940
        Frame9.Width = 8715
        Frame9.Left = 6750
        
        Frame10.Top = 8055
        Frame11.Top = 8055
                
        Frame12.Left = 13845
        Frame12.Width = 13845
                
        Frame13.Top = 8040
        Frame13.Left = 13845
        Frame13.Width = 2340
        
        frameStatus.Top = 8040
        frameStatus.Left = 15550
                
        optnModelo1.Top = 10350
        optnModelo2.Top = 10665
        
        panClone.Left = 22365
        panClone.Top = 7470
                
        btnConverterProspect.Left = 12405
                
        btnPesquisar.Left = 7905
                
        gridSolucoes.Width = 8460
        gridSolucoes.Left = 580
               
        gridGarantias.Width = 6195
                
        gridFormasPagamento.Left = 120
        gridFormasPagamento.Width = 8460
                
        txtDesconto.Left = 555
        txtEntrega.Top = 5400
        txtTempoServico.Top = 5400
        
        Label10.Top = 5505
        Label11.Top = 5505
        
        lblSolucoesTotalGeral.Left = 5880
        Label9.Left = 5190
        btnRemoverSolucao.Left = 2625
        btnEditarSolucao.Left = 1380
        btnAdicionarSolucao.Left = 225
        
        txtNecessidade.Width = 6135
        btnDesconto.Left = 1200
    End If
End Sub

Private Sub Form_Load()
   
    Me.Caption = Me.Caption + " - " + versaoAtual
   
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
    
    dtpData.value = DateTime.Now
    
    Call carregaCboClientes(Return_Recordset(Return_SqlString(1)))
    
    'Call carregaCboClientesProspects(Return_Recordset(Return_SqlString(1)))
    
    Call ConfigureGridEnderecos
    Call ConfigureGridSolucoes
    Call ConfigureGridGarantias
    Call ConfigureGridFormasPagamento
    Call ConfigureGridObservacoes
    
    Call ConfigureGridOrcamentos
    Call carregaGridOrcamentos(Return_Recordset(Return_SqlString(10)))
    
    Call HabilitarComponentes(False)
    
    OrcamentoGarantiaCodigo = 0
    OrcamentoFormaPagamentoCodigo = 0
    OrcamentoObservacaoCodigo = 0
    
    qtdSolucoesInicio = gridSolucoes.Rows - 2
    
    pblc_strClienteTipo = "PROSPECT"
    dtpData.value = DateTime.Now
    
End Sub

Private Sub gridEnderecos_DblClick()
    Dim X As Integer
    
    With gridEnderecos
        For X = 2 To .Rows - 1
            .TextMatrix(X, 3) = "N"
        Next X
    
        If .TextMatrix(.RowSel, 3) = "S" Then
            .TextMatrix(.RowSel, 3) = "N"
        Else
            .TextMatrix(.RowSel, 3) = "S"
        End If
    End With
    
End Sub

Private Sub gridOrcamentos_SelChange()
    
    If gridOrcamentos.Rows > 2 Then
        If gridOrcamentos.TextMatrix(2, 0) <> "" And gridOrcamentos.TextMatrix(gridOrcamentos.RowSel, 0) <> "" Then
            Call carregaInfoOrcamento(Return_Recordset(Return_SqlString(11, gridOrcamentos.TextMatrix(gridOrcamentos.RowSel, 0)))) 'validado
            Call carregaInfoSolucoes(Return_Recordset(Return_SqlString(12, gridOrcamentos.TextMatrix(gridOrcamentos.RowSel, 0)))) 'validado
            Call carregaInfoGarantias(Return_Recordset(Return_SqlString(13, gridOrcamentos.TextMatrix(gridOrcamentos.RowSel, 0)))) 'validado
            Call carregaInfoFormasPagto(Return_Recordset(Return_SqlString(14, gridOrcamentos.TextMatrix(gridOrcamentos.RowSel, 0)))) 'validado
            Call carregaInfoObservacoes(Return_Recordset(Return_SqlString(15, gridOrcamentos.TextMatrix(gridOrcamentos.RowSel, 0)))) 'validado
            Call carregaGridEnderecos(Return_Recordset(Return_SqlString(3, lblCodigo.Caption)))
            
            'novaFuncao para selecionar o fone padrão
            'Call carregaInfoOrcamento(Return_Recordset(Return_SqlString(11, gridOrcamentos.TextMatrix(gridOrcamentos.RowSel, 0)))) 'validado
            Call FonePadrao(Return_Recordset(Return_SqlString(28, gridOrcamentos.TextMatrix(gridOrcamentos.RowSel, 0))))
            
        End If
    End If
End Sub

Private Sub lblNumeroOrcamento_Change()
    If Not IsNumeric(lblNumeroOrcamento.Text) Then
    lblNumeroOrcamento.Text = ""
    End If
End Sub

Private Sub optnCliente_Click()
    txtCodigo.Text = ""
    pblc_strTipoPesquisa = "CLIENTE"
    Call carregaCboClientesProspects(Return_Recordset(Return_SqlString(1)))
End Sub

Private Sub optnProspect_Click()
    pblc_strTipoPesquisa = "PROSPECT"
    Call carregaCboClientesProspects(Return_Recordset(Return_SqlString(22)))
End Sub

Private Sub txtCodigo_GotFocus()
    pblc_strTipoPesquisa = "CODIGO"
    optnCliente.value = False
    cboClientesProspects.ListIndex = -1
End Sub

Private Sub txtDesconto_KeyPress(KeyAscii As Integer)

    'If Not IsNumeric(txtDesconto.Text) And KeyAscii = 8 Then
    '    MsgBox "Digite somente NÚMEROS para o CEP.", vbOKOnly, "xSol Aquecedores"
    '    Exit Sub
    'End If
    
End Sub

Private Sub txtDesconto_KeyUp(KeyCode As Integer, Shift As Integer)
    'If Not IsNumeric(txtDesconto.Text) And KeyCode <> 8 And KeyCode <> 46 Then
    '   MsgBox "Digite somente NÚMEROS para o valor do desconto.", vbOKOnly, "xSol Aquecedores"
    '    txtDesconto.Text = ""
    '    txtDesconto.SetFocus
    '    Exit Sub
    'End If
End Sub

Private Sub txtDesconto_LostFocus()
    'strCEP = Mid(txtDesconto.Text, 1, 5) & "-" & Mid(Text_Cep.Text, 6, 3)
    'Text_Cep.Text = strCEP
    If Trim(txtDesconto.Text) <> "" Then
        txtDesconto.Text = FormatNumber(txtDesconto.Text, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
End Sub

Private Sub txtNecessidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 180 Then
        KeyAscii = 0
    End If
End Sub


Private Sub editaSolucao(strNumeroOrcamento As String, strDescricao As String, strQtd As String, strValorUnitario As String, strValorTotal As String)
    frmOrcamentoSolucoes.blnEditando = True
    frmOrcamentoSolucoes.strNumeroOrcamento = strNumeroOrcamento
    frmOrcamentoSolucoes.strDescricaoSolucao = strDescricao

    frmOrcamentoSolucoes.txtDescricao = strDescricao
    frmOrcamentoSolucoes.txtQtd = strQtd
    
    'MsgBox (Mid(strValorUnitario, 3, Len(strValorUnitario)))
    
    frmOrcamentoSolucoes.txtValUnitario = Mid(strValorUnitario, 3, Len(strValorUnitario))
    frmOrcamentoSolucoes.txtValTotal = Mid(strValorTotal, 3, Len(strValorTotal))

    frmOrcamentoSolucoes.Show (1)
End Sub

Private Sub editaGarantia(strNumeroOrcamento As String, strDescricao As String)
    frmOrcamentoGarantias.blnEditando = True
    frmOrcamentoGarantias.strNumeroOrcamento = strNumeroOrcamento
    frmOrcamentoGarantias.strDescricaoGarantia = strDescricao
    frmOrcamentoGarantias.Show (1)
End Sub

Private Sub FonePadrao(rs As ADODB.Recordset, Optional code As String)
    

    
    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    If rs.Fields("Fone_Padrao") = "celular" Then
        chkCelular.value = 1
        chkResidencial.value = 0
        chkComercial.value = 0
    ElseIf rs.Fields("Fone_Padrao") = "comercial" Then
        chkComercial.value = 1
        chkResidencial.value = 0
        chkCelular.value = 0
    ElseIf rs.Fields("Fone_Padrao") = "residencial" Then
        chkResidencial.value = 1
        chkCelular.value = 0
        chkComercial.value = 0
    Else
        chkResidencial.value = 0
        chkCelular.value = 0
        chkComercial.value = 0
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Function G_NomeComputador() As String
On Error GoTo Err_G_NomeComputador
    Dim sBuffer As String
    Dim lBufSize As Long
    Dim lStatus As Long
   
    lBufSize = 255
    sBuffer = String$(lBufSize, " ")
    lStatus = GetComputerName(sBuffer, lBufSize)
    G_NomeComputador = ""
    If lStatus <> 0 Then
       G_NomeComputador = Trim(Left(sBuffer, lBufSize))
    End If
    If G_NomeComputador = "" Then G_NomeComputador = "Não Logado"
    Exit Function
Err_G_NomeComputador:
    G_NomeComputador = "Não Logado"
End Function

Private Sub txtStatusMensagem_GotFocus()
    strStatusDescricao = txtStatusMensagem.Text
End Sub

Private Sub txtStatusMensagem_LostFocus()

    'houve alteração na descrição do status do orçamento, então salva automaticamente
    If txtStatusMensagem.Text <> strStatusDescricao Then
        Dim rs As ADODB.Recordset
        
        Set rs = Return_Recordset(Return_SqlString(27, lblNumeroOrcamento.Text))
        
        rs.Fields("STATUS") = cboStatus.Text
        rs.Fields("STATUS_MSG") = txtStatusMensagem.Text
        
        rs.Update
        rs.Close
        Set rs = Nothing
        'btnRestaurar_Click
    End If


End Sub

