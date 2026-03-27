VERSION 5.00
Begin VB.Form frmRefeicoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Refeições da Hospedagem"
   ClientHeight    =   2070
   ClientLeft      =   6015
   ClientTop       =   3270
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboDatas 
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
      ItemData        =   "frmRefeicoes.frx":0000
      Left            =   720
      List            =   "frmRefeicoes.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   765
      Width           =   1695
   End
   Begin VB.ComboBox cboRefeicao 
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
      ItemData        =   "frmRefeicoes.frx":003B
      Left            =   3900
      List            =   "frmRefeicoes.frx":004B
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   765
      Width           =   2250
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1575
      Width           =   1305
   End
   Begin VB.CommandButton btnSalvar 
      BackColor       =   &H0080FF80&
      Caption         =   "INSERIR"
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
      Left            =   5055
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1605
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Refeição:"
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
      Left            =   2805
      TabIndex        =   6
      Top             =   855
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
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
      Left            =   90
      TabIndex        =   4
      Top             =   855
      Width           =   585
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
      Left            =   3105
      TabIndex        =   3
      Top             =   60
      Width           =   1260
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "NÚMERO DA HOSPEDAGEM:"
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
      Left            =   105
      TabIndex        =   2
      Top             =   135
      Width           =   2910
   End
End
Attribute VB_Name = "frmRefeicoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dtDataEntrada As Date
Public dtDataSaida As Date

Private Sub btnCancelar_Click()
Unload Me
End Sub

Private Sub btnSalvar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Dim dataInicio As Date
    Dim dataFim As Date
    Dim qtdDias As Integer
    Dim x As Integer
    Dim proximaData As Date

    dataInicio = dtDataEntrada
    dataFim = dtDataSaida
    
    qtdDias = dataFim - dataInicio
    
    cboDatas.Clear
    
    cboDatas.AddItem Str(dataInicio)
    proximaData = dataInicio
    For x = 1 To qtdDias
        proximaData = proximaData + 1
        If qtdDias = 1 Then
            Exit For
        End If
        cboDatas.AddItem Str(proximaData)
    Next x
    
    
    
    

End Sub
