VERSION 5.00
Begin VB.Form frmQuartos 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quartos"
   ClientHeight    =   9450
   ClientLeft      =   2130
   ClientTop       =   690
   ClientWidth     =   16290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   16290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4950
      TabIndex        =   34
      Top             =   8760
      Width           =   1905
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Limpeza"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   315
         TabIndex        =   35
         Top             =   60
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2640
      TabIndex        =   32
      Top             =   8760
      Width           =   1905
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ocupado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   270
         TabIndex        =   33
         Top             =   60
         Width           =   1395
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   315
      TabIndex        =   30
      Top             =   8760
      Width           =   1905
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Livre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   570
         TabIndex        =   31
         Top             =   60
         Width           =   795
      End
   End
   Begin VB.CommandButton btnFechar 
      BackColor       =   &H00C0C0FF&
      Caption         =   "FECHAR"
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
      Left            =   14835
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8715
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      Height          =   8175
      Left            =   8415
      TabIndex        =   1
      Top             =   360
      Width           =   7755
      Begin VB.CheckBox chkQuarto211 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   4860
         TabIndex        =   60
         Top             =   975
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto209 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   4860
         TabIndex        =   59
         Top             =   2310
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto207 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   4860
         TabIndex        =   58
         Top             =   3630
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto205 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   4860
         TabIndex        =   57
         Top             =   4965
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto203 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   4860
         TabIndex        =   56
         Top             =   6300
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto201 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   4860
         TabIndex        =   55
         Top             =   7635
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto212 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   240
         TabIndex        =   54
         Top             =   990
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto210 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   225
         TabIndex        =   53
         Top             =   2325
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto208 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   240
         TabIndex        =   52
         Top             =   3645
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto206 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   240
         TabIndex        =   51
         Top             =   4995
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto204 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   240
         TabIndex        =   50
         Top             =   6315
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto202 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   240
         TabIndex        =   49
         Top             =   7650
         Width           =   2550
      End
      Begin VB.CommandButton btnQuarto205 
         BackColor       =   &H0080FF80&
         Caption         =   "205"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4845
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   4305
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto207 
         BackColor       =   &H0080FF80&
         Caption         =   "207"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4845
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2955
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto203 
         BackColor       =   &H0080FF80&
         Caption         =   "203"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4845
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   5625
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto209 
         BackColor       =   &H0080FF80&
         Caption         =   "209"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4845
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1650
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto201 
         BackColor       =   &H0080FF80&
         Caption         =   "201"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4845
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   6960
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto211 
         BackColor       =   &H0080FF80&
         Caption         =   "211"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4845
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   315
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto206 
         BackColor       =   &H0080FF80&
         Caption         =   "206"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4317
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto208 
         BackColor       =   &H0080FF80&
         Caption         =   "208"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2988
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto204 
         BackColor       =   &H0080FF80&
         Caption         =   "204"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5646
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto210 
         BackColor       =   &H0080FF80&
         Caption         =   "210"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1659
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto202 
         BackColor       =   &H0080FF80&
         Caption         =   "202"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   6975
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto212 
         BackColor       =   &H0080FF80&
         Caption         =   "212"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   330
         Width           =   2640
      End
      Begin VB.Label Label4 
         Caption         =   "C           O          R          R           E           D          O          R"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6585
         Left            =   3705
         TabIndex        =   28
         Top             =   795
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8175
      Left            =   330
      TabIndex        =   0
      Top             =   375
      Width           =   7755
      Begin VB.CheckBox chkQuarto104 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   225
         TabIndex        =   48
         Top             =   6285
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto106 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   225
         TabIndex        =   47
         Top             =   4980
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto108 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   240
         TabIndex        =   46
         Top             =   3660
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto110 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   240
         TabIndex        =   45
         Top             =   2340
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto112 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   255
         TabIndex        =   44
         Top             =   1035
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto111 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   4905
         TabIndex        =   43
         Top             =   1035
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto109 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   4920
         TabIndex        =   42
         Top             =   2340
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto107 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   4920
         TabIndex        =   41
         Top             =   3660
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto105 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   4920
         TabIndex        =   40
         Top             =   4965
         Width           =   2550
      End
      Begin VB.CheckBox chkQuarto103 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   4905
         TabIndex        =   39
         Top             =   6285
         Width           =   2550
      End
      Begin VB.CommandButton btnQuarto112 
         BackColor       =   &H0080FF80&
         Caption         =   "112"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   360
         Width           =   2640
      End
      Begin VB.CheckBox chkQuarto102 
         Caption         =   "Indisponível"
         Height          =   240
         Left            =   225
         TabIndex        =   37
         Top             =   7605
         Width           =   2565
      End
      Begin VB.CheckBox chkQuarto101 
         Caption         =   "Indisponível"
         Height          =   225
         Left            =   4935
         TabIndex        =   36
         Top             =   7590
         Width           =   2550
      End
      Begin VB.CommandButton btnQuarto105 
         BackColor       =   &H0080FF80&
         Caption         =   "105"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4308
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto107 
         BackColor       =   &H0080FF80&
         Caption         =   "107"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2997
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto103 
         BackColor       =   &H0080FF80&
         Caption         =   "103"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5610
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto109 
         BackColor       =   &H0080FF80&
         Caption         =   "109"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1665
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto101 
         BackColor       =   &H0080FF80&
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6930
         Width           =   2580
      End
      Begin VB.CommandButton btnQuarto111 
         BackColor       =   &H0080FF80&
         Caption         =   "111"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   375
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto104 
         BackColor       =   &H0080FF80&
         Caption         =   "104"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpeza iniciada em XX/XX/XXXX às X horas."
         Top             =   5619
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto108 
         BackColor       =   &H0080FF80&
         Caption         =   "108"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2985
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto106 
         BackColor       =   &H0080FF80&
         Caption         =   "106"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4320
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto102 
         BackColor       =   &H0080FF80&
         Caption         =   "102"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6945
         Width           =   2640
      End
      Begin VB.CommandButton btnQuarto110 
         BackColor       =   &H0080FF80&
         Caption         =   "110"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1686
         Width           =   2640
      End
      Begin VB.Label Label3 
         Caption         =   "C           O          R          R           E           D          O          R"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6585
         Left            =   3765
         TabIndex        =   15
         Top             =   780
         Width           =   870
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ALA II"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11655
      TabIndex        =   3
      Top             =   15
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ALA I"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   30
      Width           =   945
   End
End
Attribute VB_Name = "frmQuartos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public statusQuarto As Integer
Public strQuartoClicado As String

Private Sub Command2_Click()

End Sub

Private Sub btnFechar_Click()
    Unload Me
End Sub

Private Sub btnQuarto101_Click()
    strQuartoClicado = "101"
    
    With btnQuarto101
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto 101?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With
End Sub

Private Sub btnQuarto101_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With btnQuarto101
        If .BackColor = &H80FF80 Then     'Verde
            .ToolTipText = "QUARTO LIVRE!"
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            .ToolTipText = "QUARTO OCUPADO!"
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            .ToolTipText = "QUARTO PRECISA DE LIMPEZA!"
        End If
    End With
    
End Sub


Private Sub btnQuarto102_Click()
    strQuartoClicado = "102"
    
    With btnQuarto102
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto " & strQuartoClicado & " ?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto103_Click()
    strQuartoClicado = "103"
    
    With btnQuarto103
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto " & strQuartoClicado & " ?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With
End Sub

Private Sub btnQuarto104_Click()
    strQuartoClicado = "104"
    
    With btnQuarto104
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto " & strQuartoClicado & " ?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto105_Click()
    strQuartoClicado = "105"
    
    With btnQuarto105
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto " & strQuartoClicado & " ?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto106_Click()
    strQuartoClicado = "106"
    
    With btnQuarto106
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto " & strQuartoClicado & " ?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto107_Click()
    strQuartoClicado = "107"
    
    With btnQuarto107
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto " & strQuartoClicado & " ?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto108_Click()
    strQuartoClicado = "108"
    
    With btnQuarto108
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto " & strQuartoClicado & " ?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto109_Click()
    strQuartoClicado = "109"
    
    With btnQuarto109
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto " & strQuartoClicado & " ?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto110_Click()
    strQuartoClicado = "110"
    
    With btnQuarto110
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto " & strQuartoClicado & " ?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto111_Click()
    strQuartoClicado = "111"
    
    With btnQuarto111
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto " & strQuartoClicado & " ?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto112_Click()
    strQuartoClicado = "112"
    
    With btnQuarto112
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto " & strQuartoClicado & " ?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub Check5_Click()
    If chkQuarto101.Value = 1 Then
        btnQuarto101.Enabled = False
        btnQuarto101.BackColor = &H8000000F
    Else
        btnQuarto101.Enabled = True
        btnQuarto101.BackColor = &H80FF80
    End If

End Sub


Private Sub btnQuarto201_Click()
    strQuartoClicado = "201"
    
    With btnQuarto201
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto 101?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With
End Sub

Private Sub btnQuarto202_Click()
    strQuartoClicado = "202"
    With btnQuarto202
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto 101?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto203_Click()
    strQuartoClicado = "203"
    With btnQuarto203
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto 101?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With
End Sub

Private Sub btnQuarto204_Click()
    strQuartoClicado = "204"
    With btnQuarto204
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto 101?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto205_Click()
    strQuartoClicado = "205"
    With btnQuarto205
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto 101?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto206_Click()
    strQuartoClicado = "206"
    With btnQuarto206
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto 101?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto207_Click()
    strQuartoClicado = "207"
    With btnQuarto207
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto 101?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto208_Click()
    strQuartoClicado = "208"
    With btnQuarto208
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto 101?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto209_Click()
    strQuartoClicado = "209"
    With btnQuarto209
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto 101?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto210_Click()
    strQuartoClicado = "210"
    With btnQuarto210
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto 101?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto211_Click()
    strQuartoClicado = "211"
    With btnQuarto211
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto 101?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub btnQuarto212_Click()
    strQuartoClicado = "212"
    With btnQuarto212
        If .BackColor = &H80FF80 Then     'Verde
            frmHospedagem.Show (1)
        ElseIf .BackColor = &H8080FF Then 'Vermelho claro
            frmHospedagem.Show (1)
        ElseIf .BackColor = &HFFFF& Then  'Amarelo
            If MsgBox("Você confirma a limpeza do quarto 101?", vbYesNo, "Confirmação") = vbYes Then
                statusQuarto = 0
            End If
        End If
        
        If statusQuarto = 0 Then 'verde
            .BackColor = &H80FF80
        ElseIf statusQuarto = 1 Then 'vermelho claro
            .BackColor = &H8080FF
        ElseIf statusQuarto = 2 Then 'amarelo
            .BackColor = &HFFFF&
        End If
    End With

End Sub

Private Sub chkQuarto101_Click()
    If chkQuarto101.Value = 1 Then
        btnQuarto101.Enabled = False
        btnQuarto101.BackColor = &H8000000F
    Else
        btnQuarto101.Enabled = True
        btnQuarto101.BackColor = &H80FF80
    End If
End Sub

Private Sub chkQuarto102_Click()
    If chkQuarto102.Value = 1 Then
        btnQuarto102.Enabled = False
        btnQuarto102.BackColor = &H8000000F
    Else
        btnQuarto102.Enabled = True
        btnQuarto102.BackColor = &H80FF80
    End If

End Sub

Private Sub chkQuarto103_Click()
    If chkQuarto103.Value = 1 Then
        btnQuarto103.Enabled = False
        btnQuarto103.BackColor = &H8000000F
    Else
        btnQuarto103.Enabled = True
        btnQuarto103.BackColor = &H80FF80
    End If
End Sub

Private Sub chkQuarto104_Click()
    If chkQuarto104.Value = 1 Then
        btnQuarto104.Enabled = False
        btnQuarto104.BackColor = &H8000000F
    Else
        btnQuarto104.Enabled = True
        btnQuarto104.BackColor = &H80FF80
    End If

End Sub

Private Sub chkQuarto105_Click()
    If chkQuarto105.Value = 1 Then
        btnQuarto105.Enabled = False
        btnQuarto105.BackColor = &H8000000F
    Else
        btnQuarto105.Enabled = True
        btnQuarto105.BackColor = &H80FF80
    End If

End Sub

Private Sub chkQuarto106_Click()
    If chkQuarto106.Value = 1 Then
        btnQuarto106.Enabled = False
        btnQuarto106.BackColor = &H8000000F
    Else
        btnQuarto106.Enabled = True
        btnQuarto106.BackColor = &H80FF80
    End If

End Sub

Private Sub chkQuarto107_Click()
    If chkQuarto107.Value = 1 Then
        btnQuarto107.Enabled = False
        btnQuarto107.BackColor = &H8000000F
    Else
        btnQuarto107.Enabled = True
        btnQuarto107.BackColor = &H80FF80
    End If

End Sub

Private Sub chkQuarto108_Click()
    If chkQuarto108.Value = 1 Then
        btnQuarto108.Enabled = False
        btnQuarto108.BackColor = &H8000000F
    Else
        btnQuarto108.Enabled = True
        btnQuarto108.BackColor = &H80FF80
    End If

End Sub

Private Sub chkQuarto109_Click()
    If chkQuarto109.Value = 1 Then
        btnQuarto109.Enabled = False
        btnQuarto109.BackColor = &H8000000F
    Else
        btnQuarto109.Enabled = True
        btnQuarto109.BackColor = &H80FF80
    End If

End Sub

Private Sub chkQuarto110_Click()
    If chkQuarto110.Value = 1 Then
        btnQuarto110.Enabled = False
        btnQuarto110.BackColor = &H8000000F
    Else
        btnQuarto110.Enabled = True
        btnQuarto110.BackColor = &H80FF80
    End If

End Sub

Private Sub chkQuarto111_Click()
    If chkQuarto111.Value = 1 Then
        btnQuarto111.Enabled = False
        btnQuarto111.BackColor = &H8000000F
    Else
        btnQuarto111.Enabled = True
        btnQuarto111.BackColor = &H80FF80
    End If

End Sub

Private Sub chkQuarto112_Click()
    If chkQuarto112.Value = 1 Then
        btnQuarto112.Enabled = False
        btnQuarto112.BackColor = &H8000000F
    Else
        btnQuarto112.Enabled = True
        btnQuarto112.BackColor = &H80FF80
    End If

End Sub

Private Sub chkQuarto201_Click()
    If chkQuarto201.Value = 1 Then
        btnQuarto201.Enabled = False
        btnQuarto201.BackColor = &H8000000F
    Else
        btnQuarto201.Enabled = True
        btnQuarto201.BackColor = &H80FF80
    End If
End Sub

Private Sub chkQuarto202_Click()
    If chkQuarto202.Value = 1 Then
        btnQuarto202.Enabled = False
        btnQuarto202.BackColor = &H8000000F
    Else
        btnQuarto202.Enabled = True
        btnQuarto202.BackColor = &H80FF80
    End If
End Sub

Private Sub chkQuarto203_Click()
    If chkQuarto203.Value = 1 Then
        btnQuarto203.Enabled = False
        btnQuarto203.BackColor = &H8000000F
    Else
        btnQuarto203.Enabled = True
        btnQuarto203.BackColor = &H80FF80
    End If
End Sub

Private Sub chkQuarto204_Click()
    If chkQuarto204.Value = 1 Then
        btnQuarto204.Enabled = False
        btnQuarto204.BackColor = &H8000000F
    Else
        btnQuarto204.Enabled = True
        btnQuarto204.BackColor = &H80FF80
    End If
End Sub

Private Sub chkQuarto205_Click()
    If chkQuarto205.Value = 1 Then
        btnQuarto205.Enabled = False
        btnQuarto205.BackColor = &H8000000F
    Else
        btnQuarto205.Enabled = True
        btnQuarto205.BackColor = &H80FF80
    End If
End Sub

Private Sub chkQuarto206_Click()
    If chkQuarto206.Value = 1 Then
        btnQuarto206.Enabled = False
        btnQuarto206.BackColor = &H8000000F
    Else
        btnQuarto206.Enabled = True
        btnQuarto206.BackColor = &H80FF80
    End If
End Sub

Private Sub chkQuarto207_Click()
    If chkQuarto207.Value = 1 Then
        btnQuarto207.Enabled = False
        btnQuarto207.BackColor = &H8000000F
    Else
        btnQuarto207.Enabled = True
        btnQuarto207.BackColor = &H80FF80
    End If
End Sub

Private Sub chkQuarto208_Click()
    If chkQuarto208.Value = 1 Then
        btnQuarto208.Enabled = False
        btnQuarto208.BackColor = &H8000000F
    Else
        btnQuarto208.Enabled = True
        btnQuarto208.BackColor = &H80FF80
    End If
End Sub

Private Sub chkQuarto209_Click()
    If chkQuarto209.Value = 1 Then
        btnQuarto209.Enabled = False
        btnQuarto209.BackColor = &H8000000F
    Else
        btnQuarto209.Enabled = True
        btnQuarto209.BackColor = &H80FF80
    End If
End Sub

Private Sub chkQuarto210_Click()
    If chkQuarto210.Value = 1 Then
        btnQuarto210.Enabled = False
        btnQuarto210.BackColor = &H8000000F
    Else
        btnQuarto210.Enabled = True
        btnQuarto210.BackColor = &H80FF80
    End If
End Sub

Private Sub chkQuarto211_Click()
    If chkQuarto211.Value = 1 Then
        btnQuarto211.Enabled = False
        btnQuarto211.BackColor = &H8000000F
    Else
        btnQuarto211.Enabled = True
        btnQuarto211.BackColor = &H80FF80
    End If
End Sub

Private Sub chkQuarto212_Click()
    If chkQuarto212.Value = 1 Then
        btnQuarto212.Enabled = False
        btnQuarto212.BackColor = &H8000000F
    Else
        btnQuarto212.Enabled = True
        btnQuarto212.BackColor = &H80FF80
    End If
End Sub



