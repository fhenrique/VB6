VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Begin VB.Form FrmCadHospedes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Hóspedes"
   ClientHeight    =   5505
   ClientLeft      =   3645
   ClientTop       =   2895
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   12510
   Begin VB.CommandButton btnExcluir 
      Caption         =   "EXCLUIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2415
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4905
      Width           =   1110
   End
   Begin VB.CommandButton btnAlterar 
      Caption         =   "ALTERAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1065
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4905
      Width           =   1215
   End
   Begin VB.CommandButton btnIncluir 
      Caption         =   "NOVO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4905
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de Hóspedes"
      Height          =   5415
      Left            =   6300
      TabIndex        =   18
      Top             =   0
      Width           =   6105
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5070
         Left            =   75
         TabIndex        =   19
         Top             =   285
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   8943
         _Version        =   393216
      End
   End
   Begin VB.CheckBox chkInativo 
      Caption         =   "INATIVO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4575
      TabIndex        =   17
      Top             =   4065
      Width           =   1245
   End
   Begin VB.CheckBox chkMenorIdade 
      Caption         =   "MENOR DE IDADE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   4050
      Width           =   2385
   End
   Begin VB.CommandButton btnCancelar 
      BackColor       =   &H00C0C0FF&
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
      Height          =   510
      Left            =   3690
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4905
      Width           =   1305
   End
   Begin VB.CommandButton btnSalvar 
      BackColor       =   &H00C0FFC0&
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
      Height          =   510
      Left            =   5190
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4905
      Width           =   1005
   End
   Begin VB.TextBox txtEmail 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1455
      TabIndex        =   13
      Top             =   3465
      Width           =   4710
   End
   Begin VB.TextBox txtTelefone 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1455
      TabIndex        =   12
      Top             =   2910
      Width           =   4710
   End
   Begin VB.TextBox txtCidade 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1455
      TabIndex        =   11
      Top             =   2355
      Width           =   4710
   End
   Begin VB.TextBox txtEndereco 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1455
      TabIndex        =   10
      Top             =   1800
      Width           =   4710
   End
   Begin VB.TextBox txtCPF 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1455
      TabIndex        =   9
      Top             =   1245
      Width           =   4710
   End
   Begin VB.TextBox txtRG 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1455
      TabIndex        =   8
      Top             =   690
      Width           =   4710
   End
   Begin VB.TextBox txtNome 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1455
      TabIndex        =   7
      Top             =   165
      Width           =   4710
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "EMAIL:"
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
      Left            =   600
      TabIndex        =   6
      Top             =   3510
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "TELEFONE:"
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
      Left            =   180
      TabIndex        =   5
      Top             =   2940
      Width           =   1170
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "CIDADE"
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
      Left            =   525
      TabIndex        =   4
      Top             =   2415
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ENDEREÇO:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1875
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CPF:"
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
      TabIndex        =   2
      Top             =   1335
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "RG:"
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
      Left            =   990
      TabIndex        =   1
      Top             =   780
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
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
      Left            =   645
      TabIndex        =   0
      Top             =   240
      Width           =   705
   End
End
Attribute VB_Name = "FrmCadHospedes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()

Set W_DB = New Connection
Set conn = New ADODB.Connection
conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
W_DB.Mode = adModeReadWrite
W_DB.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
    
End Sub
