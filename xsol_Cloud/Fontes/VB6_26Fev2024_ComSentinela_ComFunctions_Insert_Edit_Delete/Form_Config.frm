VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form_Config 
   Caption         =   "Configurações"
   ClientHeight    =   5055
   ClientLeft      =   2220
   ClientTop       =   2115
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Impressora PDF"
      Height          =   2175
      Left            =   120
      TabIndex        =   11
      Top             =   2730
      Width           =   3060
      Begin VB.ListBox lstImpressoras 
         Height          =   1620
         Left            =   180
         TabIndex        =   12
         Top             =   345
         Width           =   2670
      End
   End
   Begin VB.CommandButton But_Save 
      Caption         =   "&Salvar"
      Height          =   495
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4365
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Imagens"
      Height          =   1245
      Left            =   120
      TabIndex        =   6
      Top             =   1410
      Width           =   9375
      Begin VB.CommandButton But_ChooseImageFolder 
         Caption         =   "..."
         Height          =   315
         Left            =   8880
         TabIndex        =   8
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text_ImageFolder 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   990
         TabIndex        =   7
         ToolTipText     =   "Cnpj"
         Top             =   720
         Width           =   7815
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   6480
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Diretório:"
         Height          =   195
         Left            =   225
         TabIndex        =   9
         Top             =   720
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Database"
      Height          =   1380
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.TextBox Text_DatabaseName 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   975
         TabIndex        =   5
         ToolTipText     =   "Cnpj"
         Top             =   870
         Width           =   3855
      End
      Begin VB.TextBox Text_DatabaseFolder 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   975
         TabIndex        =   3
         ToolTipText     =   "Cnpj"
         Top             =   390
         Width           =   8265
      End
      Begin VB.CommandButton But_ChooseDatabase 
         Caption         =   "..."
         Height          =   315
         Left            =   4935
         TabIndex        =   2
         Top             =   870
         Width           =   375
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7170
         Top             =   690
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Arquivo:"
         Height          =   195
         Left            =   285
         TabIndex        =   4
         Top             =   870
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Diretório:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   390
         Width           =   630
      End
   End
End
Attribute VB_Name = "Form_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConfigFile As String
Dim Path As String
Dim FileName As String
Dim ImageFolder As String

Private Sub Form_Activate()
    lstImpressoras.Clear
    listaImpressoras
    
    
Text_ImageFolder.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_DatabaseName.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_DatabaseFolder.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito

End Sub

Private Sub Form_Load()
    Call Read_Config_File
    Text_DatabaseFolder = Path
    Text_DatabaseName = FileName
    Text_ImageFolder = ImageFolder
    
End Sub

Private Sub listaImpressoras()
    Dim strCount As String
        Dim strMsg As String
        Dim prtLoop As Printer
        Dim contador As Integer
         
        On Error GoTo ShowPrinters_Err
     
        If Printers.count > 0 Then
            ' Get count of installed printers.
            strMsg = "Impressoras instaladas: " & Printers.count & vbCrLf & vbCrLf
         
            ' Enumerate printer system properties.
            
            
            For contador = 0 To Printers.count - 1
                If Printers(contador).DeviceName Like "*PDF*" Then
                    lstImpressoras.AddItem (Printers(contador).DeviceName)
                End If
            Next contador
         
        Else
            strMsg = "Não há impressoras instaladas"
        End If
         
        ' Display printer information.
        'MsgBox Prompt:=strMsg, Buttons:=vbOKOnly, Title:="Impressoras instaladas"
         
ShowPrinters_End:
        Exit Sub
         
ShowPrinters_Err:
        MsgBox Prompt:=Err.Description, Buttons:=vbCritical & vbOKOnly, _
            Title:="Error Number " & Err.Number & " Occurred"
        Resume ShowPrinters_End
End Sub

Private Sub Read_Config_File()
    ConfigFile = App.Path & "\CONFIG.INI"
    
    Path = RTrim(GetINISetting("DATABASE", "PATH", ConfigFile))
    If Right(Path, 1) <> "\" Then
       Path = Path + "\"
    End If
    FileName = GetINISetting("DATABASE", "FILENAME", ConfigFile)
    
    ImageFolder = GetINISetting("IMAGE", "FOLDER", ConfigFile)
End Sub

Private Sub But_ChooseDatabase_Click()
   With CommonDialog1
      .Filter = "Image files (*.mdb) | *.mdb;"
      .DefaultExt = "mdb"
      .DialogTitle = "Select File"
      .ShowOpen
    
      Text_DatabaseFolder = Left(.FileName, InStrRev(.FileName, "\"))
      Text_DatabaseName = Right(.FileName, Len(.FileName) - InStrRev(.FileName, "\"))
   End With
   Call Enable_SaveButton
End Sub

Private Sub But_ChooseImageFolder_Click()
   On Error Resume Next
   With CommonDialog1
      .DialogTitle = "Select a directory" 'titlebar
      .InitDir = App.Path 'start dir, might be "C:\" or so also
      .FileName = "Selecione um Diretório"  'Something in filenamebox
      .Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
      .Filter = "Directories|*.~#~" 'set files-filter to show dirs only
      .CancelError = True 'allow escape key/cancel
      .ShowSave   'show the dialog screen
   
      If Err <> 32755 Then    ' User didn't chose Cancel.
         Text_ImageFolder = Left(.FileName, InStrRev(.FileName, "\"))
      End If
   End With
   Call Enable_SaveButton
End Sub

Private Sub Enable_SaveButton()
   'But_Save.enabled = Text_DatabaseFolder <> Path Or Text_DatabaseName <> FileName Or Text_ImageFolder <> ImageFolder
   'But_Save.BackColor = IIf(But_Save.enabled, &H80FF80, &H8000000F)
End Sub

Private Sub But_Save_Click()
   Dim ConfigFile As String
    
   ConfigFile = App.Path & "\CONFIG.INI"
   
   If Path <> Text_DatabaseFolder Then
      Path = Text_DatabaseFolder
      PutINISetting "DATABASE", "PATH", Path, ConfigFile
   End If
   If FileName <> Text_DatabaseName Then
      FileName = Text_DatabaseName
      PutINISetting "DATABASE", "FILENAME", FileName, ConfigFile
   End If
   If ImageFolder <> Text_ImageFolder Then
      ImageFolder = Text_ImageFolder
      Image_Folder = ImageFolder
      PutINISetting "IMAGE", "FOLDER", ImageFolder, ConfigFile
   End If
   W_DBname = Path + FileName
   
   Call Enable_SaveButton
   Unload Me
End Sub

Private Sub lstImpressoras_Click()
    PDF_PRINTER_NAME = lstImpressoras.List(lstImpressoras.ListIndex)
End Sub
