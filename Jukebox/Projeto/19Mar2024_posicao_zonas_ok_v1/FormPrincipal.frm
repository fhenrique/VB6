VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmPrincipal 
   BorderStyle     =   0  'None
   Caption         =   "Retro Jukebox"
   ClientHeight    =   10455
   ClientLeft      =   1455
   ClientTop       =   405
   ClientWidth     =   10005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   19545
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame panZonaPrincipal 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   14460
      Left            =   -5325
      TabIndex        =   10
      Top             =   930
      Width           =   14985
      Begin VB.Frame panPagina2 
         Appearance      =   0  'Flat
         BackColor       =   &H00202020&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   13875
         Left            =   7515
         TabIndex        =   12
         Top             =   225
         Width           =   195
         Begin VB.Label lblPagina2Movel_20 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   32.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   6390
            TabIndex        =   32
            Top             =   12510
            Width           =   435
         End
         Begin VB.Label lblPagina2Movel_19 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   32.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   6390
            TabIndex        =   31
            Top             =   11310
            Width           =   435
         End
         Begin VB.Label lblPagina2Movel_18 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   32.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   6390
            TabIndex        =   30
            Top             =   9840
            Width           =   435
         End
         Begin VB.Label lblPagina2Movel_17 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   32.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   6390
            TabIndex        =   29
            Top             =   8610
            Width           =   435
         End
         Begin VB.Label lblPagina2Movel_16 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   32.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   6390
            TabIndex        =   28
            Top             =   7050
            Width           =   435
         End
         Begin VB.Label lblPagina2Movel_15 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   32.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   6390
            TabIndex        =   27
            Top             =   5850
            Width           =   435
         End
         Begin VB.Label lblPagina2Movel_14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   32.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   6390
            TabIndex        =   26
            Top             =   4245
            Width           =   435
         End
         Begin VB.Label lblPagina2Movel_13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   32.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   6390
            TabIndex        =   25
            Top             =   3060
            Width           =   435
         End
         Begin VB.Label lblPagina2Movel_12 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   32.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   6390
            TabIndex        =   24
            Top             =   1500
            Width           =   435
         End
         Begin VB.Image Image43 
            Height          =   1185
            Left            =   6045
            Picture         =   "FormPrincipal.frx":0000
            Stretch         =   -1  'True
            Top             =   11160
            Width           =   1080
         End
         Begin VB.Image Image42 
            Height          =   1185
            Left            =   6030
            Picture         =   "FormPrincipal.frx":172D
            Stretch         =   -1  'True
            Top             =   12390
            Width           =   1080
         End
         Begin VB.Image Image41 
            Height          =   1185
            Left            =   6045
            Picture         =   "FormPrincipal.frx":2E89
            Stretch         =   -1  'True
            Top             =   8460
            Width           =   1080
         End
         Begin VB.Image Image40 
            Height          =   1185
            Left            =   6030
            Picture         =   "FormPrincipal.frx":45B6
            Stretch         =   -1  'True
            Top             =   9690
            Width           =   1080
         End
         Begin VB.Image Image39 
            Height          =   1185
            Left            =   6060
            Picture         =   "FormPrincipal.frx":5D12
            Stretch         =   -1  'True
            Top             =   5685
            Width           =   1080
         End
         Begin VB.Image Image38 
            Height          =   1185
            Left            =   6045
            Picture         =   "FormPrincipal.frx":743F
            Stretch         =   -1  'True
            Top             =   6915
            Width           =   1080
         End
         Begin VB.Image Image37 
            Height          =   1185
            Left            =   6060
            Picture         =   "FormPrincipal.frx":8B9B
            Stretch         =   -1  'True
            Top             =   2925
            Width           =   1080
         End
         Begin VB.Image Image36 
            Height          =   1185
            Left            =   6045
            Picture         =   "FormPrincipal.frx":A2C8
            Stretch         =   -1  'True
            Top             =   4155
            Width           =   1080
         End
         Begin VB.Label lblPagina2Movel_11 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   32.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   6390
            TabIndex        =   23
            Top             =   330
            Width           =   435
         End
         Begin VB.Image Image26 
            Height          =   1185
            Left            =   6045
            Picture         =   "FormPrincipal.frx":BA24
            Stretch         =   -1  'True
            Top             =   165
            Width           =   1080
         End
         Begin VB.Image Image25 
            Height          =   1185
            Left            =   6030
            Picture         =   "FormPrincipal.frx":D151
            Stretch         =   -1  'True
            Top             =   1395
            Width           =   1080
         End
         Begin VB.Image Image35 
            Height          =   1185
            Left            =   195
            Picture         =   "FormPrincipal.frx":E8AD
            Stretch         =   -1  'True
            Top             =   12495
            Width           =   5820
         End
         Begin VB.Image Image34 
            Height          =   1185
            Left            =   195
            Picture         =   "FormPrincipal.frx":F8E6
            Stretch         =   -1  'True
            Top             =   11100
            Width           =   5820
         End
         Begin VB.Image Image33 
            Height          =   1185
            Left            =   195
            Picture         =   "FormPrincipal.frx":1091F
            Stretch         =   -1  'True
            Top             =   9735
            Width           =   5820
         End
         Begin VB.Image Image32 
            Height          =   1185
            Left            =   195
            Picture         =   "FormPrincipal.frx":11958
            Stretch         =   -1  'True
            Top             =   8355
            Width           =   5820
         End
         Begin VB.Image Image31 
            Height          =   1185
            Left            =   195
            Picture         =   "FormPrincipal.frx":12991
            Stretch         =   -1  'True
            Top             =   6975
            Width           =   5820
         End
         Begin VB.Image Image30 
            Height          =   1185
            Left            =   195
            Picture         =   "FormPrincipal.frx":139CA
            Stretch         =   -1  'True
            Top             =   5610
            Width           =   5820
         End
         Begin VB.Image Image29 
            Height          =   1185
            Left            =   195
            Picture         =   "FormPrincipal.frx":14A03
            Stretch         =   -1  'True
            Top             =   4230
            Width           =   5820
         End
         Begin VB.Image Image28 
            Height          =   1185
            Left            =   195
            Picture         =   "FormPrincipal.frx":15A3C
            Stretch         =   -1  'True
            Top             =   2850
            Width           =   5820
         End
         Begin VB.Image Image27 
            Height          =   1185
            Left            =   195
            Picture         =   "FormPrincipal.frx":16A75
            Stretch         =   -1  'True
            Top             =   1485
            Width           =   5820
         End
         Begin VB.Image Image24 
            Height          =   1185
            Left            =   195
            Picture         =   "FormPrincipal.frx":17AAE
            Stretch         =   -1  'True
            Top             =   105
            Width           =   5820
         End
      End
      Begin VB.Timer timerLiberaPagina1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7410
         Top             =   12825
      End
      Begin VB.Timer timerAbrePagina2 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   13095
         Top             =   12420
      End
      Begin VB.Timer timerLiberaPagina2 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7395
         Top             =   13365
      End
      Begin VB.Timer timerFechaPagina2 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   13680
         Top             =   12375
      End
      Begin VB.Timer timerAbrePagina1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   300
         Top             =   14160
      End
      Begin VB.Timer timerFechaPagina1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   810
         Top             =   14145
      End
      Begin VB.Frame panPagina1 
         Appearance      =   0  'Flat
         BackColor       =   &H00202020&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   13875
         Left            =   255
         TabIndex        =   11
         Top             =   240
         Width           =   7140
         Begin VB.Label lblPagina1Movel_10 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   32.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   315
            TabIndex        =   22
            Top             =   12555
            Width           =   435
         End
         Begin VB.Label lblPagina1Movel_9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   300
            TabIndex        =   21
            Top             =   11370
            Width           =   495
         End
         Begin VB.Label lblPagina1Movel_8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   300
            TabIndex        =   20
            Top             =   9720
            Width           =   495
         End
         Begin VB.Label lblPagina1Movel_7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   285
            TabIndex        =   19
            Top             =   8595
            Width           =   495
         End
         Begin VB.Label lblPagina1Movel_6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   300
            TabIndex        =   18
            Top             =   6975
            Width           =   495
         End
         Begin VB.Label lblPagina1Movel_5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   300
            TabIndex        =   17
            Top             =   5805
            Width           =   495
         End
         Begin VB.Label lblPagina1Movel_4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   300
            TabIndex        =   16
            Top             =   4230
            Width           =   495
         End
         Begin VB.Label lblPagina1Movel_3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   300
            TabIndex        =   15
            Top             =   3045
            Width           =   495
         End
         Begin VB.Label lblPagina1Movel_2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   300
            TabIndex        =   14
            Top             =   1500
            Width           =   495
         End
         Begin VB.Label lblPagina1Movel_1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   300
            TabIndex        =   13
            Top             =   315
            Width           =   495
         End
         Begin VB.Image Image4 
            Height          =   1185
            Left            =   1155
            Picture         =   "FormPrincipal.frx":18AE7
            Stretch         =   -1  'True
            Top             =   105
            Width           =   5835
         End
         Begin VB.Image Image5 
            Height          =   1185
            Left            =   45
            Picture         =   "FormPrincipal.frx":19B20
            Stretch         =   -1  'True
            Top             =   150
            Width           =   1080
         End
         Begin VB.Image Image6 
            Height          =   1110
            Left            =   45
            Picture         =   "FormPrincipal.frx":1B33B
            Stretch         =   -1  'True
            Top             =   1425
            Width           =   1080
         End
         Begin VB.Image Image7 
            Height          =   1185
            Left            =   1155
            Picture         =   "FormPrincipal.frx":1CB43
            Stretch         =   -1  'True
            Top             =   1485
            Width           =   5835
         End
         Begin VB.Image Image8 
            Height          =   1185
            Left            =   1155
            Picture         =   "FormPrincipal.frx":1DB7C
            Stretch         =   -1  'True
            Top             =   2855
            Width           =   5835
         End
         Begin VB.Image Image9 
            Height          =   1185
            Left            =   1155
            Picture         =   "FormPrincipal.frx":1EBB5
            Stretch         =   -1  'True
            Top             =   4230
            Width           =   5835
         End
         Begin VB.Image Image10 
            Height          =   1185
            Left            =   1155
            Picture         =   "FormPrincipal.frx":1FBEE
            Stretch         =   -1  'True
            Top             =   5605
            Width           =   5835
         End
         Begin VB.Image Image11 
            Height          =   1185
            Left            =   1155
            Picture         =   "FormPrincipal.frx":20C27
            Stretch         =   -1  'True
            Top             =   6980
            Width           =   5835
         End
         Begin VB.Image Image12 
            Height          =   1185
            Left            =   1155
            Picture         =   "FormPrincipal.frx":21C60
            Stretch         =   -1  'True
            Top             =   8355
            Width           =   5835
         End
         Begin VB.Image Image13 
            Height          =   1185
            Left            =   1155
            Picture         =   "FormPrincipal.frx":22C99
            Stretch         =   -1  'True
            Top             =   9730
            Width           =   5835
         End
         Begin VB.Image Image14 
            Height          =   1185
            Left            =   1155
            Picture         =   "FormPrincipal.frx":23CD2
            Stretch         =   -1  'True
            Top             =   11100
            Width           =   5835
         End
         Begin VB.Image Image15 
            Height          =   1185
            Left            =   1155
            Picture         =   "FormPrincipal.frx":24D0B
            Stretch         =   -1  'True
            Top             =   12465
            Width           =   5835
         End
         Begin VB.Image Image16 
            Height          =   1185
            Left            =   45
            Picture         =   "FormPrincipal.frx":25D44
            Stretch         =   -1  'True
            Top             =   2880
            Width           =   1080
         End
         Begin VB.Image Image17 
            Height          =   1110
            Left            =   45
            Picture         =   "FormPrincipal.frx":2755F
            Stretch         =   -1  'True
            Top             =   4140
            Width           =   1080
         End
         Begin VB.Image Image18 
            Height          =   1185
            Left            =   45
            Picture         =   "FormPrincipal.frx":28D67
            Stretch         =   -1  'True
            Top             =   5625
            Width           =   1080
         End
         Begin VB.Image Image19 
            Height          =   1170
            Left            =   45
            Picture         =   "FormPrincipal.frx":2A582
            Stretch         =   -1  'True
            Top             =   6885
            Width           =   1080
         End
         Begin VB.Image Image20 
            Height          =   1185
            Left            =   45
            Picture         =   "FormPrincipal.frx":2BD8A
            Stretch         =   -1  'True
            Top             =   8415
            Width           =   1110
         End
         Begin VB.Image Image21 
            Height          =   1155
            Left            =   45
            Picture         =   "FormPrincipal.frx":2D5A5
            Stretch         =   -1  'True
            Top             =   9660
            Width           =   1080
         End
         Begin VB.Image Image22 
            Height          =   1185
            Left            =   45
            Picture         =   "FormPrincipal.frx":2EDAD
            Stretch         =   -1  'True
            Top             =   11160
            Width           =   1080
         End
         Begin VB.Image Image23 
            Height          =   1155
            Left            =   45
            Picture         =   "FormPrincipal.frx":305C8
            Stretch         =   -1  'True
            Top             =   12405
            Width           =   1080
         End
      End
      Begin VB.Label lblPagina2Fixa_11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   13755
         TabIndex        =   52
         Top             =   450
         Width           =   495
      End
      Begin VB.Label lblPagina2Fixa_12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   13755
         TabIndex        =   51
         Top             =   1635
         Width           =   495
      End
      Begin VB.Label lblPagina2Fixa_13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   13755
         TabIndex        =   50
         Top             =   3180
         Width           =   495
      End
      Begin VB.Label lblPagina2Fixa_14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   13755
         TabIndex        =   49
         Top             =   4365
         Width           =   495
      End
      Begin VB.Label lblPagina2Fixa_15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   13755
         TabIndex        =   48
         Top             =   5940
         Width           =   495
      End
      Begin VB.Label lblPagina2Fixa_16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   13755
         TabIndex        =   47
         Top             =   7095
         Width           =   495
      End
      Begin VB.Label lblPagina2Fixa_17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   13740
         TabIndex        =   46
         Top             =   8730
         Width           =   495
      End
      Begin VB.Label lblPagina2Fixa_18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   13755
         TabIndex        =   45
         Top             =   9840
         Width           =   495
      End
      Begin VB.Label lblPagina2Fixa_19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   13755
         TabIndex        =   44
         Top             =   11505
         Width           =   495
      End
      Begin VB.Label lblPagina2Fixa_20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   13770
         TabIndex        =   43
         Top             =   12690
         Width           =   435
      End
      Begin VB.Label lblPagina1Fixa_1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   570
         TabIndex        =   42
         Top             =   540
         Width           =   495
      End
      Begin VB.Label lblPagina1Fixa_2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   570
         TabIndex        =   41
         Top             =   1725
         Width           =   495
      End
      Begin VB.Label lblPagina1Fixa_3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   570
         TabIndex        =   40
         Top             =   3270
         Width           =   495
      End
      Begin VB.Label lblPagina1Fixa_4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   570
         TabIndex        =   39
         Top             =   4455
         Width           =   495
      End
      Begin VB.Label lblPagina1Fixa_5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   570
         TabIndex        =   38
         Top             =   6030
         Width           =   495
      End
      Begin VB.Label lblPagina1Fixa_6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   570
         TabIndex        =   37
         Top             =   7185
         Width           =   495
      End
      Begin VB.Label lblPagina1Fixa_7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   555
         TabIndex        =   36
         Top             =   8820
         Width           =   495
      End
      Begin VB.Label lblPagina1Fixa_8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   570
         TabIndex        =   35
         Top             =   9945
         Width           =   495
      End
      Begin VB.Label lblPagina1Fixa_9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   570
         TabIndex        =   34
         Top             =   11595
         Width           =   495
      End
      Begin VB.Label lblPagina1Fixa_10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   585
         TabIndex        =   33
         Top             =   12780
         Width           =   435
      End
      Begin VB.Image Image83 
         Height          =   1185
         Left            =   1455
         Picture         =   "FormPrincipal.frx":31DD0
         Stretch         =   -1  'True
         Top             =   12675
         Width           =   5730
      End
      Begin VB.Image Image82 
         Height          =   1185
         Left            =   1455
         Picture         =   "FormPrincipal.frx":32E09
         Stretch         =   -1  'True
         Top             =   11310
         Width           =   5730
      End
      Begin VB.Image Image81 
         Height          =   1185
         Left            =   1455
         Picture         =   "FormPrincipal.frx":33E42
         Stretch         =   -1  'True
         Top             =   9945
         Width           =   5730
      End
      Begin VB.Image Image80 
         Height          =   1185
         Left            =   1455
         Picture         =   "FormPrincipal.frx":34E7B
         Stretch         =   -1  'True
         Top             =   8565
         Width           =   5730
      End
      Begin VB.Image Image79 
         Height          =   1185
         Left            =   1455
         Picture         =   "FormPrincipal.frx":35EB4
         Stretch         =   -1  'True
         Top             =   7185
         Width           =   5730
      End
      Begin VB.Image Image78 
         Height          =   1185
         Left            =   1455
         Picture         =   "FormPrincipal.frx":36EED
         Stretch         =   -1  'True
         Top             =   5820
         Width           =   5730
      End
      Begin VB.Image Image77 
         Height          =   1185
         Left            =   1455
         Picture         =   "FormPrincipal.frx":37F26
         Stretch         =   -1  'True
         Top             =   4455
         Width           =   5730
      End
      Begin VB.Image Image76 
         Height          =   1185
         Left            =   1455
         Picture         =   "FormPrincipal.frx":38F5F
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   5730
      End
      Begin VB.Image Image75 
         Height          =   1185
         Left            =   1455
         Picture         =   "FormPrincipal.frx":39F98
         Stretch         =   -1  'True
         Top             =   1695
         Width           =   5730
      End
      Begin VB.Image Image74 
         Height          =   1185
         Left            =   1455
         Picture         =   "FormPrincipal.frx":3AFD1
         Stretch         =   -1  'True
         Top             =   315
         Width           =   5730
      End
      Begin VB.Image Image73 
         Height          =   1155
         Left            =   315
         Picture         =   "FormPrincipal.frx":3C00A
         Stretch         =   -1  'True
         Top             =   12615
         Width           =   1080
      End
      Begin VB.Image Image72 
         Height          =   1185
         Left            =   345
         Picture         =   "FormPrincipal.frx":3D812
         Stretch         =   -1  'True
         Top             =   11370
         Width           =   1080
      End
      Begin VB.Image Image71 
         Height          =   1155
         Left            =   345
         Picture         =   "FormPrincipal.frx":3F02D
         Stretch         =   -1  'True
         Top             =   9870
         Width           =   1080
      End
      Begin VB.Image Image70 
         Height          =   1185
         Left            =   345
         Picture         =   "FormPrincipal.frx":40835
         Stretch         =   -1  'True
         Top             =   8625
         Width           =   1110
      End
      Begin VB.Image Image69 
         Height          =   1170
         Left            =   345
         Picture         =   "FormPrincipal.frx":42050
         Stretch         =   -1  'True
         Top             =   7095
         Width           =   1080
      End
      Begin VB.Image Image68 
         Height          =   1185
         Left            =   345
         Picture         =   "FormPrincipal.frx":43858
         Stretch         =   -1  'True
         Top             =   5835
         Width           =   1080
      End
      Begin VB.Image Image67 
         Height          =   1110
         Left            =   345
         Picture         =   "FormPrincipal.frx":45073
         Stretch         =   -1  'True
         Top             =   4350
         Width           =   1080
      End
      Begin VB.Image Image66 
         Height          =   1185
         Left            =   345
         Picture         =   "FormPrincipal.frx":4687B
         Stretch         =   -1  'True
         Top             =   3090
         Width           =   1080
      End
      Begin VB.Image Image65 
         Height          =   1110
         Left            =   345
         Picture         =   "FormPrincipal.frx":48096
         Stretch         =   -1  'True
         Top             =   1635
         Width           =   1080
      End
      Begin VB.Image Image64 
         Height          =   1185
         Left            =   345
         Picture         =   "FormPrincipal.frx":4989E
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1080
      End
      Begin VB.Image Image63 
         Height          =   1125
         Left            =   13500
         Picture         =   "FormPrincipal.frx":4B0B9
         Stretch         =   -1  'True
         Top             =   12615
         Width           =   1080
      End
      Begin VB.Image Image62 
         Height          =   1230
         Left            =   13500
         Picture         =   "FormPrincipal.frx":4C8C1
         Stretch         =   -1  'True
         Top             =   11340
         Width           =   1080
      End
      Begin VB.Image Image61 
         Height          =   1140
         Left            =   13500
         Picture         =   "FormPrincipal.frx":4E0DC
         Stretch         =   -1  'True
         Top             =   9915
         Width           =   1080
      End
      Begin VB.Image Image60 
         Height          =   1230
         Left            =   13500
         Picture         =   "FormPrincipal.frx":4F8E4
         Stretch         =   -1  'True
         Top             =   8640
         Width           =   1110
      End
      Begin VB.Image Image59 
         Height          =   1125
         Left            =   13500
         Picture         =   "FormPrincipal.frx":510FF
         Stretch         =   -1  'True
         Top             =   7140
         Width           =   1080
      End
      Begin VB.Image Image58 
         Height          =   1230
         Left            =   13515
         Picture         =   "FormPrincipal.frx":52907
         Stretch         =   -1  'True
         Top             =   5865
         Width           =   1080
      End
      Begin VB.Image Image57 
         Height          =   1125
         Left            =   13500
         Picture         =   "FormPrincipal.frx":54122
         Stretch         =   -1  'True
         Top             =   4380
         Width           =   1080
      End
      Begin VB.Image Image56 
         Height          =   1230
         Left            =   13500
         Picture         =   "FormPrincipal.frx":5592A
         Stretch         =   -1  'True
         Top             =   3105
         Width           =   1080
      End
      Begin VB.Image Image55 
         Height          =   1125
         Left            =   13500
         Picture         =   "FormPrincipal.frx":57145
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   1080
      End
      Begin VB.Image Image54 
         Height          =   1230
         Left            =   13500
         Picture         =   "FormPrincipal.frx":5894D
         Stretch         =   -1  'True
         Top             =   345
         Width           =   1080
      End
      Begin VB.Image Image53 
         Height          =   1185
         Left            =   7650
         Picture         =   "FormPrincipal.frx":5A168
         Stretch         =   -1  'True
         Top             =   12690
         Width           =   5820
      End
      Begin VB.Image Image52 
         Height          =   1185
         Left            =   7650
         Picture         =   "FormPrincipal.frx":5B1A1
         Stretch         =   -1  'True
         Top             =   11325
         Width           =   5820
      End
      Begin VB.Image Image51 
         Height          =   1185
         Left            =   7650
         Picture         =   "FormPrincipal.frx":5C1DA
         Stretch         =   -1  'True
         Top             =   9960
         Width           =   5820
      End
      Begin VB.Image Image50 
         Height          =   1185
         Left            =   7650
         Picture         =   "FormPrincipal.frx":5D213
         Stretch         =   -1  'True
         Top             =   8580
         Width           =   5820
      End
      Begin VB.Image Image49 
         Height          =   1185
         Left            =   7650
         Picture         =   "FormPrincipal.frx":5E24C
         Stretch         =   -1  'True
         Top             =   7200
         Width           =   5820
      End
      Begin VB.Image Image48 
         Height          =   1185
         Left            =   7650
         Picture         =   "FormPrincipal.frx":5F285
         Stretch         =   -1  'True
         Top             =   5835
         Width           =   5820
      End
      Begin VB.Image Image47 
         Height          =   1185
         Left            =   7650
         Picture         =   "FormPrincipal.frx":602BE
         Stretch         =   -1  'True
         Top             =   4455
         Width           =   5820
      End
      Begin VB.Image Image46 
         Height          =   1185
         Left            =   7650
         Picture         =   "FormPrincipal.frx":612F7
         Stretch         =   -1  'True
         Top             =   3075
         Width           =   5820
      End
      Begin VB.Image Image45 
         Height          =   1185
         Left            =   7620
         Picture         =   "FormPrincipal.frx":62330
         Stretch         =   -1  'True
         Top             =   1710
         Width           =   5820
      End
      Begin VB.Image Image44 
         Height          =   1185
         Left            =   7650
         Picture         =   "FormPrincipal.frx":63369
         Stretch         =   -1  'True
         Top             =   330
         Width           =   5820
      End
      Begin VB.Image Image3 
         Height          =   13845
         Left            =   7305
         Picture         =   "FormPrincipal.frx":643A2
         Stretch         =   -1  'True
         Top             =   240
         Width           =   225
      End
      Begin VB.Image Image2 
         Height          =   14325
         Left            =   15
         Picture         =   "FormPrincipal.frx":6529E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   14985
      End
   End
   Begin VB.Timer timerPiscadas 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   825
      Top             =   18570
   End
   Begin VB.TextBox txtPathSound 
      Height          =   390
      Left            =   -1215
      TabIndex        =   9
      Text            =   "txtPathSound"
      Top             =   19185
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Timer timerInsertCoin 
      Interval        =   500
      Left            =   300
      Top             =   18570
   End
   Begin VB.Frame panZona3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4350
      Left            =   10320
      TabIndex        =   4
      Top             =   14670
      Width           =   5145
      Begin VB.Image Image1 
         Height          =   4575
         Left            =   0
         Picture         =   "FormPrincipal.frx":6906D
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5145
      End
   End
   Begin VB.Frame panZona2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4350
      Left            =   5085
      TabIndex        =   2
      Top             =   14670
      Width           =   5145
      Begin VB.ListBox lbxZona2PlayList 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BeginProperty Font 
            Name            =   "DS-Digital"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   3135
         ItemData        =   "FormPrincipal.frx":6E17F
         Left            =   180
         List            =   "FormPrincipal.frx":6E181
         TabIndex        =   3
         Top             =   315
         Width           =   4845
      End
      Begin VB.Image imgZona2 
         Height          =   4410
         Left            =   0
         Picture         =   "FormPrincipal.frx":6E183
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5145
      End
   End
   Begin VB.Frame panZona1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4350
      Left            =   -45
      TabIndex        =   1
      Top             =   14670
      Width           =   5145
      Begin VB.Label lblInsertCoin 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INSERT COIN"
         BeginProperty Font 
            Name            =   "DS-Digital"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   525
         Left            =   285
         TabIndex        =   7
         Top             =   1950
         Width           =   4710
      End
      Begin VB.Label lblPlease 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLEASE"
         BeginProperty Font 
            Name            =   "DS-Digital"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   435
         Left            =   2040
         TabIndex        =   6
         Top             =   1485
         Width           =   1260
      End
      Begin VB.Label lblQtdCoins 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "DS-Digital"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   705
         Left            =   315
         TabIndex        =   5
         Top             =   3180
         Width           =   1140
      End
      Begin VB.Image imgZona1 
         Height          =   4575
         Left            =   -15
         Picture         =   "FormPrincipal.frx":72943
         Stretch         =   -1  'True
         Top             =   -225
         Width           =   5145
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   2145
      Left            =   2610
      TabIndex        =   8
      Top             =   18780
      Visible         =   0   'False
      Width           =   5340
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   9419
      _cy             =   3784
   End
   Begin VB.Label lblTecla 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TECLA"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8655
      TabIndex        =   0
      Top             =   19050
      Width           =   975
   End
   Begin VB.Image imgBackGround 
      Height          =   20205
      Left            =   0
      Picture         =   "FormPrincipal.frx":78593
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   11565
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VARIÁVEIS USADAS PARA O PATH_DOS SONS
Private var_pathSound_InsertCoin As String
Private var_pathSound_virarPaginas As String
Private var_pathSound_noCoin As String
Private var_qtd_piscadas As Integer
Private var_qtd_musicasCarregadas As Integer

Private Sub insereCredito()
    
    
        playSound (var_pathSound_InsertCoin)
    
    
End Sub

Private Sub Form_Activate()
    imgBackGround.Left = 0
    imgBackGround.Top = 0
    imgBackGround.Width = frmPrincipal.Width
    imgBackGround.Height = frmPrincipal.Height
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Dim teclaPressionada As String
    
    If (KeyAscii = vbKeyF1) And (KeyAscii = vbKeyF2) And (KeyAscii = vbKeyF3) And (KeyAscii = vbKeyF4) And (KeyAscii = vbKeyF5) And (KeyAscii = vbKeyF6) And (KeyAscii = vbKeyF7) And (KeyAscii = vbKeyF8) And (KeyAscii = vbKeyF8) And (KeyAscii = vbKeyF9) And (KeyAscii = vbKeyF10) And (KeyAscii = vbKeyF11) And (KeyAscii = vbKeyF12) And (KeyAscii = 27) And (KeyAscii = 8) And (KeyAscii = 16) And (KeyAscii = 17) And (KeyAscii = 18) Then
        Exit Sub
    End If
    
    teclaPressionada = UCase(Chr$(KeyAscii))
    
    'entram no keyassci e não no keyup
    If KeyAscii = 8 Then
        teclaPressionada = "BACKSPACE"
    ElseIf KeyAscii = 9 Then
        teclaPressionada = "TAB"
    ElseIf KeyAscii = 13 Then
        teclaPressionada = "ENTER"
    ElseIf KeyAscii = 27 Then
        teclaPressionada = "ESC"
    ElseIf KeyAscii = 45 Then
        teclaPressionada = "-"
    ElseIf KeyAscii = 61 Then
        teclaPressionada = "="
    End If
    
    lerArquivoConfigIni
    lblTecla.Caption = teclaPressionada
    
    executaTeclaPressionada (teclaPressionada)
    
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim teclaPressionada As String

    If (KeyCode <> vbKeyF1) And (KeyCode <> vbKeyF2) And (KeyCode <> vbKeyF3) And (KeyCode <> vbKeyF4) And (KeyCode <> vbKeyF5) And (KeyCode <> vbKeyF6) And (KeyCode <> vbKeyF7) And (KeyCode <> vbKeyF8) And (KeyCode <> vbKeyF9) And (KeyCode <> vbKeyF10) And (KeyCode <> vbKeyF11) And (KeyCode <> vbKeyF12) And (KeyCode <> 16) And (KeyCode <> 17) And (KeyCode <> 18) And (KeyCode <> 33) And (KeyCode <> vbKey34) And (KeyCode <> 35) And (KeyCode <> vbKey36) And (KeyCode <> 45) And (KeyCode <> 46) And (KeyCode <> 37) And (KeyCode <> 38) And (KeyCode <> 39) And (KeyCode <> 40) Then
        Exit Sub
    End If

    If KeyCode = vbKeyF1 Then
        teclaPressionada = "F1"
        lblTecla.Caption = "F1"
    ElseIf KeyCode = vbKeyF2 Then
        teclaPressionada = "F2"
        lblTecla.Caption = "F2"
    ElseIf KeyCode = vbKeyF3 Then
        teclaPressionada = "F3"
        lblTecla.Caption = "F3"
    ElseIf KeyCode = vbKeyF4 Then
        teclaPressionada = "F4"
        lblTecla.Caption = "F4"
    ElseIf KeyCode = vbKeyF5 Then
        teclaPressionada = "F5"
        lblTecla.Caption = "F5"
    ElseIf KeyCode = vbKeyF6 Then
        teclaPressionada = "F6"
        lblTecla.Caption = "F6"
    ElseIf KeyCode = vbKeyF7 Then
        teclaPressionada = "F7"
        lblTecla.Caption = "F7"
    ElseIf KeyCode = vbKeyF8 Then
        teclaPressionada = "F8"
        lblTecla.Caption = "F8"
    ElseIf KeyCode = vbKeyF9 Then
        teclaPressionada = "F9"
        lblTecla.Caption = "F9"
    ElseIf KeyCode = vbKeyF10 Then
        teclaPressionada = "F10"
        lblTecla.Caption = "F10"
    ElseIf KeyCode = vbKeyF11 Then
        teclaPressionada = "F11"
        lblTecla.Caption = "F11"
    ElseIf KeyCode = vbKeyF12 Then
        teclaPressionada = "F12"
        lblTecla.Caption = "F12"
    ElseIf KeyCode = 16 Then
        teclaPressionada = "SHIFT"
        lblTecla.Caption = "SHIFT"
    ElseIf KeyCode = 17 Then
        teclaPressionada = "CTLR"
        lblTecla.Caption = "CTLR"
    ElseIf KeyCode = 18 Then
        teclaPressionada = "ALT"
        lblTecla.Caption = "ALT"
    ElseIf KeyCode = 33 Then
        teclaPressionada = "PAGEUP"
        lblTecla.Caption = "PAGEUP"
    ElseIf KeyCode = 34 Then
        teclaPressionada = "PAGEDOWN"
        lblTecla.Caption = "PAGEDOWN"
    ElseIf KeyCode = 35 Then
        teclaPressionada = "END"
        lblTecla.Caption = "END"
    ElseIf KeyCode = 36 Then
        teclaPressionada = "HOME"
        lblTecla.Caption = "HOME"
    
    ElseIf KeyCode = 37 Then
        teclaPressionada = "LEFT"
        lblTecla.Caption = "LEFT"
    ElseIf KeyCode = 38 Then
        teclaPressionada = "UP"
        lblTecla.Caption = "UP"
    ElseIf KeyCode = 39 Then
        teclaPressionada = "RIGHT"
        lblTecla.Caption = "RIGHT"
    ElseIf KeyCode = 40 Then
        teclaPressionada = "DOWN"
        lblTecla.Caption = "DOWN"
        
    ElseIf KeyCode = 45 Then
        teclaPressionada = "INSERT"
        lblTecla.Caption = "INSERT"
    ElseIf KeyCode = 46 Then
        teclaPressionada = "DELETE"
        lblTecla.Caption = "DELETE"
    End If

    lerArquivoConfigIni
    executaTeclaPressionada (teclaPressionada)
    
End Sub

Private Sub Form_Load()
    WindowsMediaPlayer1.Left = -5340
    panZonaPrincipal.Left = 160
    panZonaPrincipal.Top = 180
    panZonaPrincipal.Width = 14985
    
    lerArquivoConfigIni
    var_pathSound_InsertCoin = App.Path & "\sounds\insertcoin.mp3"
    var_pathSound_noCoin = App.Path & "\sounds\nocoin.mp3"
    var_pathSound_virarPaginas = App.Path & "\sounds\mudarpagina.mp3"
    
    If RTrim(LTrim(global_ini_zona1_qtdCoins)) <> "" Then
        lblQtdCoins.Caption = global_ini_zona1_qtdCoins
    Else
        lblQtdCoins.Caption = "00"
    End If
    
    carregaMusicas
    
    frmConfig.Show (1)
End Sub



Private Sub Form_Unload(Cancel As Integer)
    MsgBox "encerrando"
End Sub

Private Sub lbxZona2PlayList_KeyUp(KeyCode As Integer, Shift As Integer)
    'se quiser fazer algo na movimentação dentro do playlist, é aqui que faz
End Sub


Private Function retornaFuncaoTecla(tecla_pressionada As String) As String
    'retornaFuncaoTecla GetINISetting(
End Function

Private Sub timerAbrePagina1_Timer()
    
    'Pagina 1 Aberta: Width=7140, Left= 195
    'Pagina 1 Fechada: Width=90, Left= 7065
    
    
        If panPagina1.Left >= 110 Then
            If panPagina1.Left = 510 Then
                'MsgBox "chegou"
                panPagina1.Left = 210
                panPagina1.Width = 6950
                timerAbrePagina1.Enabled = False
                
                '------------------
                ' faz aqui o tratamento para deixar fazer o left de novo
                panPagina1.Visible = False
                panPagina2.Visible = False
                
                'Retorno original, quando pagina 1 aberta
                panPagina1.Left = 255
                panPagina1.Width = 7140
                panPagina1.Top = 195
                
                panPagina2.Left = 7515
                panPagina2.Width = 195
                panPagina2.Top = 210
                
                
                panPagina1.Visible = True
                panPagina2.Visible = True
                '-----------------------
                
                Exit Sub
            Else
                panPagina1.Left = panPagina1.Left - 400
                panPagina1.Width = panPagina1.Width + 400
            End If
        Else
            'MsgBox "para aqui"
            timerFechaPagina1.Enabled = False
        End If
    

End Sub

Private Sub timerAbrePagina2_Timer()
    
    'Pagina 2 Aberta: Width=7140, Left= 7515
    'Pagina 2 Fechada: Width=195, Left= 7515

    If panPagina2.Width <= 7350 Then
        If (panPagina2.Width = 7350) Or (panPagina2.Width = 6995) Or (panPagina2.Width = 6950) Then
            panPagina2.Width = 7060
            timerAbrePagina2.Enabled = False
            
            'faz aqui o tratamento para deixar fazer o left de novo
            panPagina1.Visible = False
            panPagina2.Visible = False
            
            'Retorno original, quando pagina 1 aberta
            panPagina1.Left = 255
            panPagina1.Width = 7140
            panPagina1.Top = 240
            
            panPagina2.Left = 7515
            panPagina2.Width = 195
            panPagina2.Top = 210
            
            
            panPagina1.Visible = True
            panPagina2.Visible = True
            
            Exit Sub
        Else
            panPagina2.Width = panPagina2.Width + 400
        End If
    End If
    
End Sub

Private Sub timerFechaPagina1_Timer()

    'Pagina 1 Aberta: Width=7140, Left= 195
    'Pagina 1 Fechada: Width=90, Left= 7065
    

    If panPagina1.Left < 7065 Then
        If (panPagina1.Left = 6995) Or (panPagina1.Left = 6910) Or (panPagina1.Left = 7055) Or (panPagina1.Left = 7010) Then
            'MsgBox "chegou"
            panPagina1.Width = 15
            panPagina1.Left = 7310
            timerFechaPagina1.Enabled = False
            timerLiberaPagina2.Enabled = True
            Exit Sub
        Else
            panPagina1.Width = panPagina1.Width - 400
            panPagina1.Left = panPagina1.Left + 400
        End If
    End If


End Sub

Private Sub timerFechaPagina2_Timer()
    'Pagina 2 Aberta: Width=7140, Left= 7515
    'Pagina 2 Fechada: Width=195, Left= 7515

    If panPagina2.Width >= 195 Then
        If (panPagina2.Width = 260) Then
            'MsgBox "chegou"
            panPagina2.Width = 150
            timerFechaPagina2.Enabled = False
            timerLiberaPagina1.Enabled = True
            Exit Sub
        Else
            panPagina2.Width = panPagina2.Width - 400
        End If
    Else
        timerFechaPagina2.Enabled = False
        Exit Sub
    End If

End Sub

Private Sub timerInsertCoin_Timer()
    
    'Enquanto não tiver crédito, pisca o INSERT COIN
    If (lblQtdCoins.Caption = "00") Or (RTrim(LTrim(lblQtdCoins.Caption)) = "") Then
        lblInsertCoin.Caption = "INSERT COIN"
        If lblInsertCoin.Visible = True Then
            lblInsertCoin.Visible = False
        Else
            lblInsertCoin.Visible = True
        End If
    Else
        lblInsertCoin.Caption = "SELECT A TRACK"
        
        If lblInsertCoin.Visible = True Then
            lblInsertCoin.Visible = False
        Else
            lblInsertCoin.Visible = True
        End If

        gravaCreditosNoIni

    End If
    
    
End Sub



Private Sub executaTeclaPressionada(teclaPressionada As String)

    'validação das teclas do KEYS CONFIG
    If global_ini_config_InsertCoin = teclaPressionada Then
        insereCredito

        If CDbl(lblQtdCoins.Caption) < 10 Then
            lblQtdCoins.Caption = "0" & lblQtdCoins.Caption + 1
        Else
            lblQtdCoins.Caption = lblQtdCoins.Caption + 1
        End If
        Exit Sub
    End If
    
    'Quit
    If global_ini_config_Quit = teclaPressionada Then
        End
    End If
    
    '3o. Left
    If global_ini_config_Left = teclaPressionada Then
            If panPagina1.Width = 15 Then
                Exit Sub
            End If
            playSound (var_pathSound_virarPaginas)
            fechaPagina1
    End If
    
    '4o. Right
    If global_ini_config_Right = teclaPressionada Then
            If panPagina2.Width = 150 Then
                Exit Sub
            End If
            playSound (var_pathSound_virarPaginas)
            fechaPagina2
    End If
    

    If global_ini_zona2_tecla1 = teclaPressionada Then
        If (lblQtdCoins.Caption = "00") Or (lblQtdCoins.Caption = "0") Then
            playSound (var_pathSound_noCoin)
            Call piscaLabel("lblQtdCoins", 3)
        Else
            'insere na playlist
            'toca a música
        End If
        Exit Sub
    End If
    
    If global_ini_config_Config = teclaPressionada Then
        If frmConfig.Visible = False Then
            frmConfig.Show (1)
        Else
            Unload frmConfig
        End If
    End If
    
    

        
End Sub



Private Sub gravaCreditosNoIni()
    gravaIni "FICHAS", "quantidade", lblQtdCoins.Caption
End Sub

Private Function playSound(strPath As String)
    txtPathSound.Text = strPath
    WindowsMediaPlayer1.URL = txtPathSound.Text
End Function


Private Sub piscaLabel(nomeLabel As String, numeroPiscadas As Integer)
    var_label_a_piscar = nomeLabel
    var_limite_piscadas = numeroPiscadas
    timerPiscadas.Enabled = True
End Sub


Private Sub timerPagina1_Timer()




End Sub

Private Sub timerLiberaPagina1_Timer()
    abrePagina1
End Sub

Private Sub timerLiberaPagina2_Timer()
    abrePagina_2
End Sub

Private Sub timerPiscadas_Timer()
    
    If var_qtd_piscadas <= 4 Then
        If lblQtdCoins.Visible = True Then
            lblQtdCoins.Visible = False
        Else
            lblQtdCoins.Visible = True
        End If
    Else
        timerPiscadas.Enabled = False
        var_qtd_piscadas = 0
        lblQtdCoins.Visible = True
        Exit Sub
    End If
    
    var_qtd_piscadas = var_qtd_piscadas + 1
    
End Sub


Private Sub fechaPagina1()
    If panPagina1.Left <> 7310 Then
        timerFechaPagina1.Enabled = True
    End If
End Sub

Private Sub abrePagina1()
    timerLiberaPagina1.Enabled = False
    If panPagina1.Left <> 195 Then
        timerAbrePagina1.Enabled = True
    End If
End Sub

Private Sub fechaPagina2()
    If panPagina2.Width <> 195 Then
        timerFechaPagina2.Enabled = True
    End If
End Sub

Private Sub abrePagina_2()
    timerLiberaPagina2.Enabled = False
    If panPagina2.Width <> 6995 Then
        timerAbrePagina2.Enabled = True
    End If
End Sub


Sub CountFiles()
    Dim xFolder As String
    Dim xPath As String
    Dim xCount As Long
    Dim xFile As String
        
    xFolder = App.Path & "\sounds"
    xPath = xFolder & "\*.mp3"
    xFile = Dir(xPath)
    
    Do While xFile <> ""
        xCount = xCount + 1
        xFile = Dir()
    Loop
    
    var_qtd_musicasCarregadas = xCount
    'MsgBox var_qtd_musicasCarregadas
End Sub

Private Sub carregaMusicas()
    Dim contador As Integer
    Dim myLabel As Label
    
    'var_qtd_musicasCarregadas = 0
    
    'Esta função conta a qtd de arquivos na pasta de musicas e grava o número na variavel privada var_qtd_musicasCarregadas
    'CountFiles
    
    'For contador = 1 To var_qtd_musicasCarregadas
    '        myLabel.Name = "Label" & Str(contador)
    '        myLabel.na
            'MsgBox myLabel.Name
    'Next
    
    
    
End Sub

Private Sub timerPrintEsquerdo_Timer()
    imgPrintEsquerdo.Width = imgPrintEsquerdo.Width - 400
    If imgPrintEsquerdo.Width = 100 Then
    
    End If
End Sub
