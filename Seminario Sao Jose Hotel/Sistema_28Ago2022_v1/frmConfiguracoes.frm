VERSION 5.00
Begin VB.Form frmConfiguracoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurações"
   ClientHeight    =   4680
   ClientLeft      =   6000
   ClientTop       =   4305
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   4275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "R$ Valores das refeições"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   135
      TabIndex        =   7
      Top             =   1815
      Width           =   4020
      Begin VB.TextBox Text4 
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
         Left            =   1965
         TabIndex        =   15
         Top             =   1680
         Width           =   1185
      End
      Begin VB.TextBox Text3 
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
         Left            =   1965
         TabIndex        =   14
         Top             =   1245
         Width           =   1185
      End
      Begin VB.TextBox Text2 
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
         Left            =   1965
         TabIndex        =   13
         Top             =   765
         Width           =   1185
      End
      Begin VB.TextBox Text1 
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
         Left            =   1965
         TabIndex        =   11
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Jantar:"
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
         Left            =   1140
         TabIndex        =   12
         Top             =   1740
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Lanche:"
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
         Left            =   1050
         TabIndex        =   10
         Top             =   1305
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almoço:"
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
         Left            =   1020
         TabIndex        =   9
         Top             =   825
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Café da manhã:"
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
         Left            =   195
         TabIndex        =   8
         Top             =   390
         Width           =   1710
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "R$ Valores de diárias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   135
      TabIndex        =   2
      Top             =   105
      Width           =   4020
      Begin VB.TextBox txtDiariaBaixaTemporada 
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
         Left            =   2625
         TabIndex        =   6
         Top             =   420
         Width           =   1185
      End
      Begin VB.TextBox txtDiariaAltaTemporada 
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
         Left            =   2625
         TabIndex        =   5
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Na baixa temporada:"
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
         Left            =   210
         TabIndex        =   4
         Top             =   450
         Width           =   2280
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Na alta temporada:"
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
         Left            =   390
         TabIndex        =   3
         Top             =   975
         Width           =   2100
      End
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
      Height          =   435
      Left            =   2955
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4095
      Width           =   1170
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
      Height          =   435
      Left            =   1470
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4110
      Width           =   1305
   End
End
Attribute VB_Name = "frmConfiguracoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnSalvar_Click()
    Unload Me
End Sub
