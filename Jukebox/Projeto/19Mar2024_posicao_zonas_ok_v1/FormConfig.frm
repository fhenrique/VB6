VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retro Jukebox Config"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   720
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame20 
      Caption         =   "Command Zone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   45
      TabIndex        =   162
      Top             =   8175
      Width           =   4515
      Begin VB.Frame Frame22 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   15
         TabIndex        =   167
         Top             =   585
         Width           =   4470
         Begin VB.OptionButton Option24 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO"
            Height          =   195
            Left            =   3285
            TabIndex        =   169
            Top             =   75
            Width           =   630
         End
         Begin VB.OptionButton Option23 
            BackColor       =   &H00C0C0C0&
            Caption         =   "YES"
            Height          =   195
            Left            =   2490
            TabIndex        =   168
            Top             =   75
            Width           =   630
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Active :"
            Height          =   195
            Left            =   1890
            TabIndex        =   170
            Top             =   60
            Width           =   540
         End
      End
      Begin VB.Frame Frame21 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   15
         TabIndex        =   163
         Top             =   255
         Width           =   4470
         Begin VB.OptionButton Option22 
            BackColor       =   &H00C0C0C0&
            Caption         =   "YES"
            Height          =   195
            Left            =   2490
            TabIndex        =   165
            Top             =   75
            Width           =   630
         End
         Begin VB.OptionButton Option21 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO"
            Height          =   195
            Left            =   3285
            TabIndex        =   164
            Top             =   75
            Width           =   630
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Visible :"
            Height          =   195
            Left            =   1890
            TabIndex        =   166
            Top             =   60
            Width           =   540
         End
      End
   End
   Begin VB.Frame Frame17 
      Caption         =   "Playlist Zone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   30
      TabIndex        =   153
      Top             =   7185
      Width           =   4515
      Begin VB.Frame Frame19 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   15
         TabIndex        =   158
         Top             =   255
         Width           =   4470
         Begin VB.OptionButton Option20 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO"
            Height          =   195
            Left            =   3285
            TabIndex        =   160
            Top             =   75
            Width           =   630
         End
         Begin VB.OptionButton Option19 
            BackColor       =   &H00C0C0C0&
            Caption         =   "YES"
            Height          =   195
            Left            =   2490
            TabIndex        =   159
            Top             =   75
            Width           =   630
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Visible :"
            Height          =   195
            Left            =   1890
            TabIndex        =   161
            Top             =   60
            Width           =   540
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   15
         TabIndex        =   154
         Top             =   585
         Width           =   4470
         Begin VB.OptionButton Option18 
            BackColor       =   &H00C0C0C0&
            Caption         =   "YES"
            Height          =   195
            Left            =   2490
            TabIndex        =   156
            Top             =   75
            Width           =   630
         End
         Begin VB.OptionButton Option17 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO"
            Height          =   195
            Left            =   3285
            TabIndex        =   155
            Top             =   75
            Width           =   630
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Active :"
            Height          =   195
            Left            =   1890
            TabIndex        =   157
            Top             =   60
            Width           =   540
         End
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "Main Zone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   30
      TabIndex        =   144
      Top             =   6195
      Width           =   4515
      Begin VB.Frame Frame16 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   15
         TabIndex        =   149
         Top             =   585
         Width           =   4470
         Begin VB.OptionButton Option16 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO"
            Height          =   195
            Left            =   3285
            TabIndex        =   151
            Top             =   75
            Width           =   630
         End
         Begin VB.OptionButton Option15 
            BackColor       =   &H00C0C0C0&
            Caption         =   "YES"
            Height          =   195
            Left            =   2490
            TabIndex        =   150
            Top             =   75
            Width           =   630
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Active :"
            Height          =   195
            Left            =   1890
            TabIndex        =   152
            Top             =   60
            Width           =   540
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   15
         TabIndex        =   145
         Top             =   255
         Width           =   4470
         Begin VB.OptionButton Option14 
            BackColor       =   &H00C0C0C0&
            Caption         =   "YES"
            Height          =   195
            Left            =   2490
            TabIndex        =   147
            Top             =   75
            Width           =   630
         End
         Begin VB.OptionButton Option13 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO"
            Height          =   195
            Left            =   3285
            TabIndex        =   146
            Top             =   75
            Width           =   630
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Visible :"
            Height          =   195
            Left            =   1890
            TabIndex        =   148
            Top             =   60
            Width           =   540
         End
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Theme"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   4590
      TabIndex        =   59
      Top             =   -15
      Width           =   3495
      Begin VB.ComboBox cboTheme 
         Height          =   315
         ItemData        =   "FormConfig.frx":0000
         Left            =   75
         List            =   "FormConfig.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   270
         Width           =   3360
      End
      Begin VB.Image Image22 
         Height          =   2430
         Left            =   90
         Picture         =   "FormConfig.frx":0011
         Stretch         =   -1  'True
         Top             =   600
         Width           =   3360
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "General Config"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3045
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   4560
      Begin VB.Frame Frame3 
         Caption         =   "Video Position"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   60
         TabIndex        =   40
         Top             =   2250
         Width           =   4440
         Begin VB.TextBox txtVideoPosition_Left 
            Height          =   300
            Left            =   435
            TabIndex        =   44
            Top             =   300
            Width           =   585
         End
         Begin VB.TextBox txtVideoPosition_Top 
            Height          =   300
            Left            =   1425
            TabIndex        =   43
            Top             =   300
            Width           =   585
         End
         Begin VB.TextBox txtVideoPosition_Width 
            Height          =   300
            Left            =   2520
            TabIndex        =   42
            Top             =   300
            Width           =   585
         End
         Begin VB.TextBox txtVideoPosition_Height 
            Height          =   300
            Left            =   3720
            TabIndex        =   41
            Top             =   300
            Width           =   585
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Left"
            Height          =   195
            Left            =   135
            TabIndex        =   48
            Top             =   345
            Width           =   270
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Top"
            Height          =   195
            Left            =   1110
            TabIndex        =   47
            Top             =   345
            Width           =   285
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Width"
            Height          =   195
            Left            =   2085
            TabIndex        =   46
            Top             =   345
            Width           =   420
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Height"
            Height          =   195
            Left            =   3225
            TabIndex        =   45
            Top             =   345
            Width           =   465
         End
      End
      Begin VB.OptionButton OptnManageCoin_No 
         Caption         =   "No"
         Height          =   195
         Left            =   3285
         TabIndex        =   39
         Top             =   2040
         Width           =   615
      End
      Begin VB.OptionButton OptnManageCoin_Yes 
         Caption         =   "Yes"
         Height          =   195
         Left            =   2355
         TabIndex        =   38
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton btnSelectPath 
         Caption         =   "Select"
         Height          =   360
         Left            =   3660
         TabIndex        =   36
         Top             =   1575
         Width           =   825
      End
      Begin VB.TextBox txtPath 
         Height          =   360
         Left            =   60
         TabIndex        =   35
         Top             =   1590
         Width           =   3555
      End
      Begin VB.Frame Frame8 
         Caption         =   "Jukebox Position"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   60
         TabIndex        =   25
         Top             =   600
         Width           =   4440
         Begin VB.TextBox txtJukeboxPosition_Height 
            Height          =   300
            Left            =   3720
            TabIndex        =   33
            Top             =   300
            Width           =   585
         End
         Begin VB.TextBox txtJukeboxPosition_Width 
            Height          =   300
            Left            =   2520
            TabIndex        =   32
            Top             =   300
            Width           =   585
         End
         Begin VB.TextBox txtJukeboxPosition_Top 
            Height          =   300
            Left            =   1425
            TabIndex        =   31
            Top             =   300
            Width           =   585
         End
         Begin VB.TextBox txtJukeboxPosition_Left 
            Height          =   300
            Left            =   435
            TabIndex        =   26
            Top             =   285
            Width           =   585
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Height"
            Height          =   195
            Left            =   3225
            TabIndex        =   30
            Top             =   345
            Width           =   465
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Width"
            Height          =   195
            Left            =   2085
            TabIndex        =   29
            Top             =   345
            Width           =   420
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Top"
            Height          =   195
            Left            =   1110
            TabIndex        =   28
            Top             =   345
            Width           =   285
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Left"
            Height          =   195
            Left            =   135
            TabIndex        =   27
            Top             =   345
            Width           =   270
         End
      End
      Begin VB.CheckBox chkAtivarMouse 
         Caption         =   "Activate Tactiler or Mouse Navigation"
         Height          =   315
         Left            =   150
         TabIndex        =   24
         Top             =   255
         Width           =   3150
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Manage Coin Payment:"
         Height          =   195
         Left            =   570
         TabIndex        =   37
         Top             =   2025
         Width           =   1650
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Path of Music Collection:"
         Height          =   195
         Left            =   75
         TabIndex        =   34
         Top             =   1365
         Width           =   1755
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "TrackList"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   15
      TabIndex        =   22
      Top             =   3105
      Width           =   4530
      Begin VB.OptionButton Option12 
         Caption         =   "NO"
         Height          =   195
         Left            =   3420
         TabIndex        =   143
         Top             =   2730
         Width           =   735
      End
      Begin VB.OptionButton Option11 
         Caption         =   "YES"
         Height          =   195
         Left            =   2640
         TabIndex        =   142
         Top             =   2730
         Width           =   735
      End
      Begin VB.Frame Frame6 
         Caption         =   "Show Cover"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   105
         TabIndex        =   124
         Top             =   975
         Width           =   4320
         Begin VB.Frame Frame13 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   15
            TabIndex        =   137
            Top             =   1260
            Width           =   4275
            Begin VB.OptionButton Option10 
               BackColor       =   &H00C0C0C0&
               Caption         =   "NO"
               Height          =   195
               Left            =   3285
               TabIndex        =   139
               Top             =   75
               Width           =   630
            End
            Begin VB.OptionButton Option9 
               BackColor       =   &H00C0C0C0&
               Caption         =   "YES"
               Height          =   195
               Left            =   2490
               TabIndex        =   138
               Top             =   75
               Width           =   630
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Show Big Cover :"
               Height          =   195
               Left            =   1185
               TabIndex        =   140
               Top             =   75
               Width           =   1230
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   15
            TabIndex        =   133
            Top             =   930
            Width           =   4275
            Begin VB.OptionButton Option8 
               BackColor       =   &H00C0C0C0&
               Caption         =   "NO"
               Height          =   195
               Left            =   3285
               TabIndex        =   135
               Top             =   75
               Width           =   630
            End
            Begin VB.OptionButton Option7 
               BackColor       =   &H00C0C0C0&
               Caption         =   "YES"
               Height          =   195
               Left            =   2490
               TabIndex        =   134
               Top             =   75
               Width           =   630
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Show MP3 Tag Cover :"
               Height          =   195
               Left            =   750
               TabIndex        =   136
               Top             =   60
               Width           =   1665
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   15
            TabIndex        =   129
            Top             =   600
            Width           =   4275
            Begin VB.OptionButton Option6 
               BackColor       =   &H00C0C0C0&
               Caption         =   "NO"
               Height          =   195
               Left            =   3285
               TabIndex        =   131
               Top             =   75
               Width           =   630
            End
            Begin VB.OptionButton Option5 
               BackColor       =   &H00C0C0C0&
               Caption         =   "YES"
               Height          =   195
               Left            =   2490
               TabIndex        =   130
               Top             =   75
               Width           =   630
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Show File Cover :"
               Height          =   195
               Left            =   1155
               TabIndex        =   132
               Top             =   60
               Width           =   1245
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   15
            TabIndex        =   125
            Top             =   270
            Width           =   4275
            Begin VB.OptionButton Option4 
               BackColor       =   &H00C0C0C0&
               Caption         =   "YES"
               Height          =   195
               Left            =   2490
               TabIndex        =   127
               Top             =   75
               Width           =   630
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H00C0C0C0&
               Caption         =   "NO"
               Height          =   195
               Left            =   3285
               TabIndex        =   126
               Top             =   75
               Width           =   630
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Show Folder Cover :"
               Height          =   195
               Left            =   960
               TabIndex        =   128
               Top             =   60
               Width           =   1440
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   105
         TabIndex        =   120
         Top             =   555
         Width           =   4350
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO"
            Height          =   195
            Left            =   3285
            TabIndex        =   123
            Top             =   75
            Width           =   630
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "YES"
            Height          =   195
            Left            =   2490
            TabIndex        =   122
            Top             =   75
            Width           =   630
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Play Transition Sound :"
            Height          =   195
            Left            =   630
            TabIndex        =   121
            Top             =   75
            Width           =   1635
         End
      End
      Begin VB.ComboBox cboDefaultOrder 
         Height          =   315
         ItemData        =   "FormConfig.frx":860D
         Left            =   2415
         List            =   "FormConfig.frx":8617
         Style           =   2  'Dropdown List
         TabIndex        =   119
         Top             =   210
         Width           =   1995
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Show Track Number :"
         Height          =   195
         Left            =   975
         TabIndex        =   141
         Top             =   2715
         Width           =   1560
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Default Order :"
         Height          =   195
         Left            =   1320
         TabIndex        =   118
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10005
      TabIndex        =   21
      Top             =   9240
      Width           =   1350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   510
      Left            =   8595
      TabIndex        =   20
      Top             =   9240
      Width           =   1350
   End
   Begin VB.Frame Frame2 
      Caption         =   "Theme keys config"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6060
      Left            =   4590
      TabIndex        =   19
      Top             =   3105
      Width           =   3510
      Begin VB.CommandButton btnThemeKeysConfig_20 
         Caption         =   "?"
         Height          =   300
         Left            =   1860
         TabIndex        =   95
         Top             =   5085
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_19 
         Caption         =   "?"
         Height          =   300
         Left            =   1860
         TabIndex        =   94
         Top             =   4575
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_18 
         Caption         =   "?"
         Height          =   300
         Left            =   1860
         TabIndex        =   93
         Top             =   4065
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_17 
         Caption         =   "?"
         Height          =   300
         Left            =   1860
         TabIndex        =   92
         Top             =   3555
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_16 
         Caption         =   "?"
         Height          =   300
         Left            =   1860
         TabIndex        =   91
         Top             =   3075
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_15 
         Caption         =   "?"
         Height          =   300
         Left            =   1860
         TabIndex        =   90
         Top             =   2550
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_14 
         Caption         =   "?"
         Height          =   300
         Left            =   1860
         TabIndex        =   89
         Top             =   2055
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_13 
         Caption         =   "?"
         Height          =   300
         Left            =   1860
         TabIndex        =   88
         Top             =   1545
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_12 
         Caption         =   "?"
         Height          =   300
         Left            =   1860
         TabIndex        =   87
         Top             =   1020
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_11 
         Caption         =   "?"
         Height          =   300
         Left            =   1860
         TabIndex        =   86
         Top             =   510
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_10 
         Caption         =   "?"
         Height          =   300
         Left            =   135
         TabIndex        =   85
         Top             =   5055
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_9 
         Caption         =   "?"
         Height          =   300
         Left            =   135
         TabIndex        =   84
         Top             =   4560
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_8 
         Caption         =   "?"
         Height          =   300
         Left            =   135
         TabIndex        =   83
         Top             =   4095
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_7 
         Caption         =   "?"
         Height          =   300
         Left            =   135
         TabIndex        =   82
         Top             =   3570
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_6 
         Caption         =   "?"
         Height          =   300
         Left            =   135
         TabIndex        =   81
         Top             =   3060
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_5 
         Caption         =   "?"
         Height          =   300
         Left            =   135
         TabIndex        =   80
         Top             =   2565
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_4 
         Caption         =   "?"
         Height          =   300
         Left            =   135
         TabIndex        =   79
         Top             =   2040
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_3 
         Caption         =   "?"
         Height          =   300
         Left            =   135
         TabIndex        =   78
         Top             =   1545
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_2 
         Caption         =   "?"
         Height          =   300
         Left            =   135
         TabIndex        =   77
         Top             =   1035
         Width           =   315
      End
      Begin VB.CommandButton btnThemeKeysConfig_1 
         Caption         =   "?"
         Height          =   300
         Left            =   135
         TabIndex        =   76
         Top             =   510
         Width           =   315
      End
      Begin VB.Label lblThemeKeysConfig_20 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   2295
         TabIndex        =   117
         Top             =   5145
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_19 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   2295
         TabIndex        =   116
         Top             =   4620
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_18 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   2295
         TabIndex        =   115
         Top             =   4110
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_17 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   2295
         TabIndex        =   114
         Top             =   3600
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_16 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   2295
         TabIndex        =   113
         Top             =   3105
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_15 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   2295
         TabIndex        =   112
         Top             =   2580
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_14 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   2295
         TabIndex        =   111
         Top             =   2085
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   2295
         TabIndex        =   110
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   2295
         TabIndex        =   109
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   2295
         TabIndex        =   108
         Top             =   555
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   510
         TabIndex        =   107
         Top             =   555
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   540
         TabIndex        =   106
         Top             =   5145
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   540
         TabIndex        =   105
         Top             =   4635
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   540
         TabIndex        =   104
         Top             =   4110
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   540
         TabIndex        =   103
         Top             =   3600
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   540
         TabIndex        =   102
         Top             =   3105
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   540
         TabIndex        =   101
         Top             =   2580
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   540
         TabIndex        =   100
         Top             =   2085
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   540
         TabIndex        =   99
         Top             =   1575
         Width           =   990
      End
      Begin VB.Label lblThemeKeysConfig_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   540
         TabIndex        =   98
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Double click on the label key to re-initialize"
         Height          =   195
         Left            =   165
         TabIndex        =   97
         Top             =   5775
         Width           =   3000
      End
      Begin VB.Image Image21 
         Height          =   390
         Left            =   2220
         Picture         =   "FormConfig.frx":8631
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   1125
      End
      Begin VB.Image Image20 
         Height          =   390
         Left            =   2220
         Picture         =   "FormConfig.frx":8F6E
         Stretch         =   -1  'True
         Top             =   4530
         Width           =   1125
      End
      Begin VB.Image Image19 
         Height          =   390
         Left            =   2220
         Picture         =   "FormConfig.frx":98AB
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   1125
      End
      Begin VB.Image Image18 
         Height          =   390
         Left            =   2220
         Picture         =   "FormConfig.frx":A1E8
         Stretch         =   -1  'True
         Top             =   3510
         Width           =   1125
      End
      Begin VB.Image Image17 
         Height          =   390
         Left            =   2220
         Picture         =   "FormConfig.frx":AB25
         Stretch         =   -1  'True
         Top             =   3015
         Width           =   1125
      End
      Begin VB.Image Image16 
         Height          =   390
         Left            =   2220
         Picture         =   "FormConfig.frx":B462
         Stretch         =   -1  'True
         Top             =   2505
         Width           =   1125
      End
      Begin VB.Image Image15 
         Height          =   390
         Left            =   2205
         Picture         =   "FormConfig.frx":BD9F
         Stretch         =   -1  'True
         Top             =   1995
         Width           =   1125
      End
      Begin VB.Image Image14 
         Height          =   390
         Left            =   2220
         Picture         =   "FormConfig.frx":C6DC
         Stretch         =   -1  'True
         Top             =   1485
         Width           =   1125
      End
      Begin VB.Image Image13 
         Height          =   390
         Left            =   2220
         Picture         =   "FormConfig.frx":D019
         Stretch         =   -1  'True
         Top             =   990
         Width           =   1125
      End
      Begin VB.Image Image12 
         Height          =   390
         Left            =   2220
         Picture         =   "FormConfig.frx":D956
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1125
      End
      Begin VB.Image Image11 
         Height          =   390
         Left            =   480
         Picture         =   "FormConfig.frx":E293
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   1125
      End
      Begin VB.Image Image10 
         Height          =   390
         Left            =   480
         Picture         =   "FormConfig.frx":EBD0
         Stretch         =   -1  'True
         Top             =   4530
         Width           =   1125
      End
      Begin VB.Image Image9 
         Height          =   390
         Left            =   480
         Picture         =   "FormConfig.frx":F50D
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   1125
      End
      Begin VB.Image image8 
         Height          =   390
         Left            =   480
         Picture         =   "FormConfig.frx":FE4A
         Stretch         =   -1  'True
         Top             =   3510
         Width           =   1125
      End
      Begin VB.Image Image7 
         Height          =   390
         Left            =   480
         Picture         =   "FormConfig.frx":10787
         Stretch         =   -1  'True
         Top             =   3015
         Width           =   1125
      End
      Begin VB.Image Image6 
         Height          =   390
         Left            =   480
         Picture         =   "FormConfig.frx":110C4
         Stretch         =   -1  'True
         Top             =   2505
         Width           =   1125
      End
      Begin VB.Image Image5 
         Height          =   390
         Left            =   480
         Picture         =   "FormConfig.frx":11A01
         Stretch         =   -1  'True
         Top             =   1995
         Width           =   1125
      End
      Begin VB.Image Image4 
         Height          =   390
         Left            =   480
         Picture         =   "FormConfig.frx":1233E
         Stretch         =   -1  'True
         Top             =   1485
         Width           =   1125
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   480
         Picture         =   "FormConfig.frx":12C7B
         Stretch         =   -1  'True
         Top             =   990
         Width           =   1125
      End
      Begin VB.Image Image2 
         Height          =   345
         Left            =   480
         Picture         =   "FormConfig.frx":135B8
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1125
      End
      Begin VB.Image Image1 
         Height          =   5430
         Left            =   30
         Picture         =   "FormConfig.frx":13EF5
         Stretch         =   -1  'True
         Top             =   255
         Width           =   3435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Keys config"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9165
      Left            =   8130
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txtKeysConfig_Config 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   1395
         TabIndex        =   96
         Top             =   3465
         Width           =   1365
      End
      Begin VB.CommandButton Command17 
         Caption         =   "?"
         Height          =   360
         Left            =   2790
         TabIndex        =   75
         Top             =   5880
         Width           =   390
      End
      Begin VB.CommandButton Command16 
         Caption         =   "?"
         Height          =   360
         Left            =   2790
         TabIndex        =   74
         Top             =   5490
         Width           =   390
      End
      Begin VB.CommandButton Command15 
         Caption         =   "?"
         Height          =   360
         Left            =   2790
         TabIndex        =   73
         Top             =   5070
         Width           =   390
      End
      Begin VB.CommandButton Command14 
         Caption         =   "?"
         Height          =   360
         Left            =   2790
         TabIndex        =   72
         Top             =   4680
         Width           =   390
      End
      Begin VB.CommandButton Command13 
         Caption         =   "?"
         Height          =   360
         Left            =   2790
         TabIndex        =   71
         Top             =   4275
         Width           =   390
      End
      Begin VB.CommandButton Command12 
         Caption         =   "?"
         Height          =   360
         Left            =   2790
         TabIndex        =   70
         Top             =   3885
         Width           =   390
      End
      Begin VB.CommandButton Command11 
         Caption         =   "?"
         Height          =   360
         Left            =   2790
         TabIndex        =   69
         Top             =   3495
         Width           =   390
      End
      Begin VB.CommandButton Command10 
         Caption         =   "?"
         Height          =   360
         Left            =   2790
         TabIndex        =   68
         Top             =   3090
         Width           =   390
      End
      Begin VB.CommandButton Command9 
         Caption         =   "?"
         Height          =   360
         Left            =   2790
         TabIndex        =   67
         Top             =   2700
         Width           =   390
      End
      Begin VB.CommandButton Command8 
         Caption         =   "?"
         Height          =   360
         Left            =   2790
         TabIndex        =   66
         Top             =   2295
         Width           =   390
      End
      Begin VB.CommandButton Command7 
         Caption         =   "?"
         Height          =   360
         Left            =   2790
         TabIndex        =   65
         Top             =   1905
         Width           =   390
      End
      Begin VB.CommandButton Command6 
         Caption         =   "?"
         Height          =   360
         Left            =   2790
         TabIndex        =   64
         Top             =   1515
         Width           =   390
      End
      Begin VB.CommandButton Command5 
         Caption         =   "?"
         Height          =   360
         Left            =   2790
         TabIndex        =   63
         Top             =   1125
         Width           =   390
      End
      Begin VB.CommandButton Command4 
         Caption         =   "?"
         Height          =   360
         Left            =   2790
         TabIndex        =   62
         Top             =   705
         Width           =   390
      End
      Begin VB.CommandButton Command3 
         Caption         =   "?"
         Height          =   360
         Left            =   2790
         TabIndex        =   61
         Top             =   315
         Width           =   390
      End
      Begin VB.TextBox txtKeysConfig_Next 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   1395
         TabIndex        =   53
         Top             =   5910
         Width           =   1365
      End
      Begin VB.TextBox txtKeysConfig_Pause 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   1395
         TabIndex        =   52
         Top             =   5472
         Width           =   1365
      End
      Begin VB.TextBox txtKeysConfig_Play 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   1395
         TabIndex        =   51
         Top             =   5073
         Width           =   1365
      End
      Begin VB.TextBox txtKeysConfig_Stop 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   1395
         TabIndex        =   50
         Top             =   4674
         Width           =   1365
      End
      Begin VB.TextBox txtKeysConfig_Preview 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   1395
         TabIndex        =   49
         Top             =   4275
         Width           =   1365
      End
      Begin VB.TextBox txtKeysConfig_Down 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   1395
         TabIndex        =   18
         Top             =   684
         Width           =   1365
      End
      Begin VB.TextBox txtKeysConfig_Left 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   1395
         TabIndex        =   15
         Top             =   1083
         Width           =   1365
      End
      Begin VB.TextBox txtKeysConfig_Right 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   1395
         TabIndex        =   14
         Top             =   1482
         Width           =   1365
      End
      Begin VB.TextBox txtKeysConfig_Validation 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   1395
         TabIndex        =   13
         Top             =   1881
         Width           =   1365
      End
      Begin VB.TextBox txtKeysConfig_BackCancel 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   1395
         TabIndex        =   12
         Top             =   2280
         Width           =   1365
      End
      Begin VB.TextBox txtKeysConfig_InsertCoin 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   1395
         TabIndex        =   11
         Top             =   2679
         Width           =   1365
      End
      Begin VB.TextBox txtKeysConfig_SwitchZone 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   1395
         TabIndex        =   5
         Top             =   3078
         Width           =   1365
      End
      Begin VB.TextBox txtKeysConfig_Quit 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   1395
         TabIndex        =   2
         Top             =   3876
         Width           =   1365
      End
      Begin VB.Label txtKeysConfig_UP 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1395
         TabIndex        =   171
         Top             =   345
         Width           =   1365
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Next Track :"
         Height          =   195
         Left            =   435
         TabIndex        =   58
         Top             =   5970
         Width           =   885
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Pause :"
         Height          =   195
         Left            =   780
         TabIndex        =   57
         Top             =   5550
         Width           =   540
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Play :"
         Height          =   195
         Left            =   930
         TabIndex        =   56
         Top             =   5145
         Width           =   390
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Stop :"
         Height          =   195
         Left            =   900
         TabIndex        =   55
         Top             =   4770
         Width           =   420
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Preview Track :"
         Height          =   195
         Left            =   195
         TabIndex        =   54
         Top             =   4365
         Width           =   1125
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Quit :"
         Height          =   195
         Left            =   945
         TabIndex        =   17
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Config :"
         Height          =   195
         Left            =   780
         TabIndex        =   16
         Top             =   3555
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Switch Zone :"
         Height          =   195
         Left            =   330
         TabIndex        =   10
         Top             =   3180
         Width           =   990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Insert  Coin :"
         Height          =   195
         Left            =   435
         TabIndex        =   9
         Top             =   2775
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Validation :"
         Height          =   195
         Left            =   540
         TabIndex        =   8
         Top             =   1965
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Back / Cancel :"
         Height          =   195
         Left            =   195
         TabIndex        =   7
         Top             =   2370
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Right :"
         Height          =   195
         Left            =   855
         TabIndex        =   6
         Top             =   1575
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Left :"
         Height          =   195
         Left            =   960
         TabIndex        =   4
         Top             =   1185
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Up :"
         Height          =   195
         Left            =   1020
         TabIndex        =   3
         Top             =   405
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Down :"
         Height          =   195
         Left            =   810
         TabIndex        =   1
         Top             =   765
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnThemeKeysConfig_1_Click()
    Call configTeclasTheme(1, lblThemeKeysConfig_1.Caption)
End Sub

Private Sub btnThemeKeysConfig_10_Click()
    Call configTeclasTheme(10, lblThemeKeysConfig_10.Caption)
End Sub

Private Sub btnThemeKeysConfig_11_Click()
    Call configTeclasTheme(11, lblThemeKeysConfig_11.Caption)
End Sub

Private Sub btnThemeKeysConfig_12_Click()
    Call configTeclasTheme(12, lblThemeKeysConfig_12.Caption)
End Sub

Private Sub btnThemeKeysConfig_13_Click()
    Call configTeclasTheme(13, lblThemeKeysConfig_13.Caption)
End Sub

Private Sub btnThemeKeysConfig_14_Click()
    Call configTeclasTheme(14, lblThemeKeysConfig_14.Caption)
End Sub

Private Sub btnThemeKeysConfig_15_Click()
    Call configTeclasTheme(15, lblThemeKeysConfig_15.Caption)
End Sub

Private Sub btnThemeKeysConfig_16_Click()
    Call configTeclasTheme(16, lblThemeKeysConfig_16.Caption)
End Sub

Private Sub btnThemeKeysConfig_17_Click()
    Call configTeclasTheme(17, lblThemeKeysConfig_17.Caption)
End Sub

Private Sub btnThemeKeysConfig_18_Click()
    Call configTeclasTheme(18, lblThemeKeysConfig_18.Caption)
End Sub

Private Sub btnThemeKeysConfig_19_Click()
    Call configTeclasTheme(19, lblThemeKeysConfig_19.Caption)
End Sub

Private Sub btnThemeKeysConfig_2_Click()
    Call configTeclasTheme(2, lblThemeKeysConfig_2.Caption)
End Sub

Private Sub btnThemeKeysConfig_20_Click()
    Call configTeclasTheme(20, lblThemeKeysConfig_20.Caption)
End Sub

Private Sub btnThemeKeysConfig_3_Click()
    Call configTeclasTheme(3, lblThemeKeysConfig_3.Caption)
End Sub

Private Sub btnThemeKeysConfig_4_Click()
    Call configTeclasTheme(4, lblThemeKeysConfig_4.Caption)
End Sub

Private Sub btnThemeKeysConfig_5_Click()
    Call configTeclasTheme(5, lblThemeKeysConfig_5.Caption)
End Sub

Private Sub btnThemeKeysConfig_6_Click()
    Call configTeclasTheme(6, lblThemeKeysConfig_6.Caption)
End Sub

Private Sub btnThemeKeysConfig_7_Click()
    Call configTeclasTheme(7, lblThemeKeysConfig_7.Caption)
End Sub

Private Sub btnThemeKeysConfig_8_Click()
    Call configTeclasTheme(8, lblThemeKeysConfig_8.Caption)
End Sub

Private Sub btnThemeKeysConfig_9_Click()
    Call configTeclasTheme(9, lblThemeKeysConfig_9.Caption)
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command10_Click()
    Call configTeclasKeysConfig("switchzone")
End Sub

Private Sub Command11_Click()
    Call configTeclasKeysConfig("config")
End Sub

Private Sub Command12_Click()
    Call configTeclasKeysConfig("quit")
End Sub

Private Sub Command13_Click()
    Call configTeclasKeysConfig("previewtrack")
End Sub

Private Sub Command14_Click()
    Call configTeclasKeysConfig("stop")
End Sub

Private Sub Command15_Click()
    Call configTeclasKeysConfig("play")
End Sub

Private Sub Command16_Click()
    Call configTeclasKeysConfig("pause")
End Sub

Private Sub Command17_Click()
    Call configTeclasKeysConfig("nexttrack")
End Sub

Private Sub Command2_Click()
    gravaIni "THEMEKEYSCONFIG", "track1", lblThemeKeysConfig_1.Caption
    gravaIni "THEMEKEYSCONFIG", "track2", lblThemeKeysConfig_2.Caption
    gravaIni "THEMEKEYSCONFIG", "track3", lblThemeKeysConfig_3.Caption
    gravaIni "THEMEKEYSCONFIG", "track4", lblThemeKeysConfig_4.Caption
    gravaIni "THEMEKEYSCONFIG", "track5", lblThemeKeysConfig_5.Caption
    gravaIni "THEMEKEYSCONFIG", "track6", lblThemeKeysConfig_6.Caption
    gravaIni "THEMEKEYSCONFIG", "track7", lblThemeKeysConfig_7.Caption
    gravaIni "THEMEKEYSCONFIG", "track8", lblThemeKeysConfig_8.Caption
    gravaIni "THEMEKEYSCONFIG", "track9", lblThemeKeysConfig_9.Caption
    gravaIni "THEMEKEYSCONFIG", "track10", lblThemeKeysConfig_10.Caption
    gravaIni "THEMEKEYSCONFIG", "track11", lblThemeKeysConfig_11.Caption
    gravaIni "THEMEKEYSCONFIG", "track12", lblThemeKeysConfig_12.Caption
    gravaIni "THEMEKEYSCONFIG", "track13", lblThemeKeysConfig_13.Caption
    gravaIni "THEMEKEYSCONFIG", "track14", lblThemeKeysConfig_14.Caption
    gravaIni "THEMEKEYSCONFIG", "track15", lblThemeKeysConfig_15.Caption
    gravaIni "THEMEKEYSCONFIG", "track16", lblThemeKeysConfig_16.Caption
    gravaIni "THEMEKEYSCONFIG", "track17", lblThemeKeysConfig_17.Caption
    gravaIni "THEMEKEYSCONFIG", "track18", lblThemeKeysConfig_18.Caption
    gravaIni "THEMEKEYSCONFIG", "track19", lblThemeKeysConfig_19.Caption
    gravaIni "THEMEKEYSCONFIG", "track20", lblThemeKeysConfig_20.Caption
    
    gravaIni "KEYSCONFIG", "Up", txtKeysConfig_UP.Caption
    gravaIni "KEYSCONFIG", "Down", txtKeysConfig_Down.Text
    gravaIni "KEYSCONFIG", "Left", txtKeysConfig_Left.Text
    gravaIni "KEYSCONFIG", "Right", txtKeysConfig_Right.Text
    gravaIni "KEYSCONFIG", "Validation", txtKeysConfig_Validation.Text
    gravaIni "KEYSCONFIG", "BackCancel", txtKeysConfig_BackCancel.Text
    gravaIni "KEYSCONFIG", "InsertCoin", txtKeysConfig_InsertCoin.Text
    gravaIni "KEYSCONFIG", "SwitchZone", txtKeysConfig_SwitchZone.Text
    gravaIni "KEYSCONFIG", "Config", txtKeysConfig_Config.Text
    gravaIni "KEYSCONFIG", "Quit", txtKeysConfig_Quit.Text
    gravaIni "KEYSCONFIG", "PreviewTrack", txtKeysConfig_Preview.Text
    gravaIni "KEYSCONFIG", "Stop", txtKeysConfig_Stop.Text
    gravaIni "KEYSCONFIG", "Play", txtKeysConfig_Play.Text
    gravaIni "KEYSCONFIG", "Pause", txtKeysConfig_Pause.Text
    gravaIni "KEYSCONFIG", "NextTrack", txtKeysConfig_Next.Text
    
    
    

    
    Unload Me
End Sub

Private Sub configTeclasTheme(numeroTecla As Integer, strTecla As String)
    teclaAConfigurar = numeroTecla
    frmPressAnyKey.lblTeclaAtalho.Caption = strTecla
    frmPressAnyKey.Show (1)
End Sub

Private Sub configTeclasKeysConfig(strTecla As String)
    teclaAConfigurar_str = strTecla
    frmKeysConfig.Show (1)
End Sub

Private Sub Command3_Click()
    Call configTeclasKeysConfig("up")
End Sub

Private Sub Command4_Click()
    Call configTeclasKeysConfig("down")
End Sub

Private Sub Command5_Click()
    Call configTeclasKeysConfig("left")
End Sub

Private Sub Command6_Click()
    Call configTeclasKeysConfig("right")
End Sub

Private Sub Command7_Click()
    Call configTeclasKeysConfig("validation")
End Sub

Private Sub Command8_Click()
    Call configTeclasKeysConfig("backcancel")
End Sub

Private Sub Command9_Click()
    Call configTeclasKeysConfig("insertcoin")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    cboTheme.ListIndex = 0
    Call Arquivo_Dados
End Sub

Public Sub Arquivo_Dados()
    Dim ConfigFile As String
    Dim atalho_track1 As String
    Dim strTemp As String

    ConfigFile = App.Path & "\config.INI"
    
    'Theme Keys Config
    config_track1_tecla = GetINISetting("THEMEKEYSCONFIG", "track1", ConfigFile)
    config_track2_tecla = GetINISetting("THEMEKEYSCONFIG", "track2", ConfigFile)
    config_track3_tecla = GetINISetting("THEMEKEYSCONFIG", "track3", ConfigFile)
    config_track4_tecla = GetINISetting("THEMEKEYSCONFIG", "track4", ConfigFile)
    config_track5_tecla = GetINISetting("THEMEKEYSCONFIG", "track5", ConfigFile)
    config_track6_tecla = GetINISetting("THEMEKEYSCONFIG", "track6", ConfigFile)
    config_track7_tecla = GetINISetting("THEMEKEYSCONFIG", "track7", ConfigFile)
    config_track8_tecla = GetINISetting("THEMEKEYSCONFIG", "track8", ConfigFile)
    config_track9_tecla = GetINISetting("THEMEKEYSCONFIG", "track9", ConfigFile)
    config_track10_tecla = GetINISetting("THEMEKEYSCONFIG", "track10", ConfigFile)
    config_track11_tecla = GetINISetting("THEMEKEYSCONFIG", "track11", ConfigFile)
    config_track12_tecla = GetINISetting("THEMEKEYSCONFIG", "track12", ConfigFile)
    config_track13_tecla = GetINISetting("THEMEKEYSCONFIG", "track13", ConfigFile)
    config_track14_tecla = GetINISetting("THEMEKEYSCONFIG", "track14", ConfigFile)
    config_track15_tecla = GetINISetting("THEMEKEYSCONFIG", "track15", ConfigFile)
    config_track16_tecla = GetINISetting("THEMEKEYSCONFIG", "track16", ConfigFile)
    config_track17_tecla = GetINISetting("THEMEKEYSCONFIG", "track17", ConfigFile)
    config_track18_tecla = GetINISetting("THEMEKEYSCONFIG", "track18", ConfigFile)
    config_track19_tecla = GetINISetting("THEMEKEYSCONFIG", "track19", ConfigFile)
    config_track20_tecla = GetINISetting("THEMEKEYSCONFIG", "track20", ConfigFile)
    
    'txtThemeKeysConfig_1.Text = config_track1_tecla
    lblThemeKeysConfig_1.Caption = config_track1_tecla
    lblThemeKeysConfig_2.Caption = config_track2_tecla
    lblThemeKeysConfig_3.Caption = config_track3_tecla
    lblThemeKeysConfig_4.Caption = config_track4_tecla
    lblThemeKeysConfig_5.Caption = config_track5_tecla
    lblThemeKeysConfig_6.Caption = config_track6_tecla
    lblThemeKeysConfig_7.Caption = config_track7_tecla
    lblThemeKeysConfig_8.Caption = config_track8_tecla
    lblThemeKeysConfig_9.Caption = config_track9_tecla
    lblThemeKeysConfig_10.Caption = config_track10_tecla
    lblThemeKeysConfig_11.Caption = config_track11_tecla
    lblThemeKeysConfig_12.Caption = config_track12_tecla
    lblThemeKeysConfig_13.Caption = config_track13_tecla
    lblThemeKeysConfig_14.Caption = config_track14_tecla
    lblThemeKeysConfig_15.Caption = config_track15_tecla
    lblThemeKeysConfig_16.Caption = config_track16_tecla
    lblThemeKeysConfig_17.Caption = config_track17_tecla
    lblThemeKeysConfig_18.Caption = config_track18_tecla
    lblThemeKeysConfig_19.Caption = config_track19_tecla
    lblThemeKeysConfig_20.Caption = config_track20_tecla
    
    
    
    'Keys Config
    txtKeysConfig_UP.Caption = GetINISetting("KEYSCONFIG", "Up", ConfigFile)
    txtKeysConfig_Down.Text = GetINISetting("KEYSCONFIG", "Down", ConfigFile)
    txtKeysConfig_Left.Text = GetINISetting("KEYSCONFIG", "Left", ConfigFile)
    txtKeysConfig_Right.Text = GetINISetting("KEYSCONFIG", "Right", ConfigFile)
    txtKeysConfig_Validation.Text = GetINISetting("KEYSCONFIG", "Validation", ConfigFile)
    txtKeysConfig_BackCancel.Text = GetINISetting("KEYSCONFIG", "BackCancel", ConfigFile)
    txtKeysConfig_InsertCoin.Text = GetINISetting("KEYSCONFIG", "InsertCoin", ConfigFile)
    txtKeysConfig_SwitchZone.Text = GetINISetting("KEYSCONFIG", "Switchzone", ConfigFile)
    txtKeysConfig_Config.Text = GetINISetting("KEYSCONFIG", "Config", ConfigFile)
    txtKeysConfig_Quit.Text = GetINISetting("KEYSCONFIG", "Quit", ConfigFile)
    txtKeysConfig_Preview.Text = GetINISetting("KEYSCONFIG", "PreviewTrack", ConfigFile)
    txtKeysConfig_Stop.Text = GetINISetting("KEYSCONFIG", "Stop", ConfigFile)
    txtKeysConfig_Play.Text = GetINISetting("KEYSCONFIG", "Play", ConfigFile)
    txtKeysConfig_Pause.Text = GetINISetting("KEYSCONFIG", "Pause", ConfigFile)
    txtKeysConfig_Next.Text = GetINISetting("KEYSCONFIG", "NextTrack", ConfigFile)
    
    
    
    
    
    
    
    
    
End Sub


Private Sub lblThemeKeysConfig_1_DblClick()
    lblThemeKeysConfig_1.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_10_DblClick()
    lblThemeKeysConfig_10.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_11_DblClick()
    lblThemeKeysConfig_11.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_12_DblClick()
    lblThemeKeysConfig_12.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_13_DblClick()
    lblThemeKeysConfig_13.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_14_DblClick()
    lblThemeKeysConfig_14.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_15_DblClick()
    lblThemeKeysConfig_15.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_16_DblClick()
    lblThemeKeysConfig_16.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_17_DblClick()
    lblThemeKeysConfig_17.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_18_DblClick()
    lblThemeKeysConfig_18.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_19_DblClick()
    lblThemeKeysConfig_19.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_2_DblClick()
    lblThemeKeysConfig_2.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_20_DblClick()
    lblThemeKeysConfig_20.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_3_DblClick()
    lblThemeKeysConfig_3.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_4_DblClick()
    lblThemeKeysConfig_4.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_5_DblClick()
    lblThemeKeysConfig_5.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_6_DblClick()
    lblThemeKeysConfig_6.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_7_DblClick()
    lblThemeKeysConfig_7.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_8_DblClick()
    lblThemeKeysConfig_8.Caption = ""
End Sub

Private Sub lblThemeKeysConfig_9_DblClick()
    lblThemeKeysConfig_9.Caption = ""
End Sub

Private Sub txtKeysConfig_UP_DblClick()
    txtKeysConfig_UP.Caption = ""
End Sub
