VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHospedagem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seminário São José - Hospedagem"
   ClientHeight    =   10110
   ClientLeft      =   3330
   ClientTop       =   525
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10110
   ScaleWidth      =   11460
   Begin VB.CommandButton btnInserirRefeicao 
      BackColor       =   &H0080FF80&
      Caption         =   "Inserir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10110
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5490
      Width           =   1215
   End
   Begin VB.CommandButton btnExcluirRefeicao 
      BackColor       =   &H008080FF&
      Caption         =   "Remover"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10110
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5910
      Width           =   1215
   End
   Begin VB.CommandButton btnFinalizarHospedagem 
      Caption         =   "FINALIZAR HOSPEDAGEM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9660
      Width           =   3105
   End
   Begin VB.CommandButton btnRemoverHospede 
      BackColor       =   &H008080FF&
      Caption         =   "Remover"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4365
      Width           =   1215
   End
   Begin VB.CommandButton btnInserirHospede 
      BackColor       =   &H0080FF80&
      Caption         =   "Inserir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3930
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grdHospedes 
      Height          =   1755
      Left            =   90
      TabIndex        =   18
      Top             =   3345
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   3096
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpDataEntrada 
      Height          =   315
      Left            =   2775
      TabIndex        =   12
      Top             =   990
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52756481
      CurrentDate     =   44787
   End
   Begin VB.ComboBox cboTipoHospedagem 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmHospedagem.frx":0000
      Left            =   7395
      List            =   "frmHospedagem.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1860
      Visible         =   0   'False
      Width           =   3915
   End
   Begin VB.TextBox txtObservacoes 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   285
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2385
      Width           =   11250
   End
   Begin VB.CommandButton btnSalvar 
      Caption         =   "SALVAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8820
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9660
      Width           =   1170
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "CANCELAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7365
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9660
      Width           =   1305
   End
   Begin MSComCtl2.DTPicker dtpDataSaida 
      Height          =   315
      Left            =   6630
      TabIndex        =   13
      Top             =   990
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52756481
      CurrentDate     =   44787
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4035
      Left            =   90
      TabIndex        =   25
      Top             =   5460
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   7117
      _Version        =   393216
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "REFEIÇÕES DA HOSPEDAGEM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   24
      Top             =   5115
      Width           =   11460
   End
   Begin VB.Label lblValorTotal 
      AutoSize        =   -1  'True
      Caption         =   "0.000,00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6450
      TabIndex        =   22
      Top             =   1530
      Width           =   960
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "R$ Total:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5415
      TabIndex        =   21
      Top             =   1530
      Width           =   945
   End
   Begin VB.Label lblQuartoClicado 
      AutoSize        =   -1  'True
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2595
      TabIndex        =   17
      Top             =   510
      Width           =   135
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Quarto:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1635
      TabIndex        =   16
      Top             =   510
      Width           =   825
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "INFORMAÇÕES DA HOSPEDAGEM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   15
      Top             =   45
      Width           =   11460
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "INFORMAÇÕES DO(S) HÓSPEDE(S)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   -15
      TabIndex        =   14
      Top             =   3060
      Width           =   11475
   End
   Begin VB.Label lblQtdHospedes 
      AutoSize        =   -1  'True
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2595
      TabIndex        =   10
      Top             =   1530
      Width           =   135
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Observações:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   300
      TabIndex        =   9
      Top             =   2100
      Width           =   1500
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Data de saída:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5025
      TabIndex        =   8
      Top             =   975
      Width           =   1560
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Data de entrada:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   855
      TabIndex        =   7
      Top             =   975
      Width           =   1830
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Qtd. de hóspedes:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   6
      Top             =   1530
      Width           =   1980
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de hospedagem:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4935
      TabIndex        =   5
      Top             =   1950
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Label lblNumeroHospedagem 
      AutoSize        =   -1  'True
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   10020
      TabIndex        =   4
      Top             =   405
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "NÚMERO:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8610
      TabIndex        =   3
      Top             =   450
      Width           =   1275
   End
End
Attribute VB_Name = "frmHospedagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnFinalizarHospedagem_Click()
    If MsgBox("Confirma a finalização da hospedagem?", vbYesNo, "Confirmação") = vbYes Then
        frmQuartos.statusQuarto = 2
        Unload Me
    End If
End Sub

Private Sub btnInserirHospede_Click()
    FrmCadHospedes.Show (1)
End Sub



Private Sub btnInserirRefeicao_Click()
    frmRefeicoes.dtDataEntrada = dtpDataEntrada.Value
    frmRefeicoes.dtDataSaida = dtpDataSaida.Value
    frmRefeicoes.Show (1)
End Sub

Private Sub btnSalvar_Click()
    frmQuartos.statusQuarto = 1
    Unload Me
End Sub

Private Sub Form_Load()
    lblQuartoClicado.Caption = frmQuartos.strQuartoClicado
End Sub
