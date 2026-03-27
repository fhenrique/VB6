VERSION 5.00
Begin VB.Form frmKeysConfig 
   Caption         =   "Form1"
   ClientHeight    =   2070
   ClientLeft      =   5505
   ClientTop       =   3045
   ClientWidth     =   4590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   4590
   Begin VB.Label Label1 
      Caption         =   "to manage the alphabetical shortcuts depending on the number of page, click on the ""page up"" button"
      Height          =   540
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   4545
   End
   Begin VB.Label lblTeclaAtalho 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   1860
      TabIndex        =   0
      Top             =   795
      Width           =   825
   End
End
Attribute VB_Name = "frmKeysConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
Dim teclaPressionada As String
    
    If (KeyAscii <> 8) And (KeyAscii <> 32) Then
        teclaPressionada = UCase(Chr$(KeyAscii))
    End If
    
    If KeyAscii = 32 Then
        teclaPressionada = "SPACEBAR"
    End If
    
    'Valida se a tecla já está sendo usada no Theme Keys config:
    If teclaUsadaNoThemeKeysConfig(teclaPressionada) Then
        MsgBox "A tecla " & teclaPressionada & " já foi atribuida no Theme keys config"
        Exit Sub
    End If
    
    'Valida se a tecla já está sendo usada no Keys config:
    If teclaUsadaNoKeysConfig(teclaPressionada) Then
        MsgBox "A tecla " & teclaPressionada & " já foi atribuiída no Keys config"
        Exit Sub
    End If
    
    
    If KeyAscii = 8 Then
        config_tecla_config = "BACKSPACE"
        lblTeclaAtalho.Caption = config_tecla_config
        GoTo configuraTecla
    End If
    
    If KeyAscii = 9 Then
        config_tecla_config = "TAB"
        lblTeclaAtalho.Caption = config_tecla_config
        GoTo configuraTecla
    End If
    
    If KeyAscii = 13 Then
        config_tecla_config = "ENTER"
        lblTeclaAtalho.Caption = config_tecla_config
        GoTo configuraTecla
    End If
    
    If KeyAscii = 27 Then
        config_tecla_config = "ESC"
        lblTeclaAtalho.Caption = config_tecla_config
        GoTo configuraTecla
    End If
    
    If KeyAscii = 32 Then
        config_tecla_config = "SPACEBAR"
        lblTeclaAtalho.Caption = config_tecla_config
        GoTo configuraTecla
    End If
    
    
    
    
    lblTeclaAtalho.Caption = UCase(Chr$(KeyAscii))
    
configuraTecla:
        If teclaAConfigurar = 1 Then
            config_track1_tecla = KeyAscii
        Else
            If KeyAscii <> 32 Then
                config_tecla_config = UCase(Chr$(KeyAscii))
            End If
            
            If KeyAscii = 32 Then
                config_tecla_config = "SPACEBAR"
            End If
            
        End If
        
        If teclaAConfigurar_str = "up" Then
            frmConfig.txtKeysConfig_UP.Caption = config_tecla_config
        ElseIf teclaAConfigurar_str = "down" Then
            frmConfig.txtKeysConfig_Down.Text = config_tecla_config
        ElseIf teclaAConfigurar_str = "left" Then
            frmConfig.txtKeysConfig_Left.Text = config_tecla_config
        ElseIf teclaAConfigurar_str = "right" Then
            frmConfig.txtKeysConfig_Right.Text = config_tecla_config
        ElseIf teclaAConfigurar_str = "validation" Then
            frmConfig.txtKeysConfig_Validation.Text = config_tecla_config
        ElseIf teclaAConfigurar_str = "backcancel" Then
            frmConfig.txtKeysConfig_BackCancel.Text = config_tecla_config
        ElseIf teclaAConfigurar_str = "insertcoin" Then
            frmConfig.txtKeysConfig_InsertCoin.Text = config_tecla_config
        ElseIf teclaAConfigurar_str = "switchzone" Then
            frmConfig.txtKeysConfig_SwitchZone.Text = config_tecla_config
        ElseIf teclaAConfigurar_str = "config" Then
            frmConfig.txtKeysConfig_Config.Text = config_tecla_config
        ElseIf teclaAConfigurar_str = "quit" Then
            frmConfig.txtKeysConfig_Quit.Text = config_tecla_config
        ElseIf teclaAConfigurar_str = "previewtrack" Then
            frmConfig.txtKeysConfig_Preview.Text = config_tecla_config
        ElseIf teclaAConfigurar_str = "stop" Then
            frmConfig.txtKeysConfig_Stop.Text = config_tecla_config
        ElseIf teclaAConfigurar_str = "play" Then
            frmConfig.txtKeysConfig_Play.Text = config_tecla_config
        ElseIf teclaAConfigurar_str = "pause" Then
            frmConfig.txtKeysConfig_Pause.Text = config_tecla_config
        ElseIf teclaAConfigurar_str = "nexttrack" Then
            frmConfig.txtKeysConfig_Next.Text = config_tecla_config
        End If
        
        Unload Me

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim teclaPressionada As String
    
    If (KeyCode <> vbKeyF1) And (KeyCode <> vbKeyF2) And (KeyCode <> vbKeyF3) And (KeyCode <> vbKeyF4) And (KeyCode <> vbKeyF5) And (KeyCode <> vbKeyF6) And (KeyCode <> vbKeyF7) And (KeyCode <> vbKeyF8) And (KeyCode <> vbKeyF9) And (KeyCode <> vbKeyF10) And (KeyCode <> vbKeyF11) And (KeyCode <> vbKeyF12) And (KeyCode <> 16) And (KeyCode <> 17) And (KeyCode <> 18) And (KeyCode <> 20) And (KeyCode <> 37) And (KeyCode <> 38) And (KeyCode <> 39) And (KeyCode <> 40) Then
        teclaPressionada = Chr$(KeyCode)
    End If
    
    
    If KeyCode = vbKeyF1 Then
        teclaPressionada = "F1"
    End If
    
    If KeyCode = vbKeyF2 Then
        teclaPressionada = "F2"
    End If
    
    If KeyCode = vbKeyF3 Then
        teclaPressionada = "F3"
    End If
    
    If KeyCode = vbKeyF4 Then
        teclaPressionada = "F4"
    End If
    
    If KeyCode = vbKeyF5 Then
        teclaPressionada = "F5"
    End If
    
    If KeyCode = vbKeyF6 Then
        teclaPressionada = "F6"
    End If
    
    If KeyCode = vbKeyF7 Then
        teclaPressionada = "F7"
    End If
    
    If KeyCode = vbKeyF8 Then
        teclaPressionada = "F8"
    End If
    
    If KeyCode = vbKeyF9 Then
        teclaPressionada = "F9"
    End If
    
    If KeyCode = vbKeyF10 Then
        teclaPressionada = "F10"
    End If
    
    If KeyCode = vbKeyF11 Then
        teclaPressionada = "F11"
    End If
    
    If KeyCode = vbKeyF12 Then
        teclaPressionada = "F12"
    End If
    
    '16,17,18,20
    
    If KeyCode = 16 Then
        teclaPressionada = "SHIFT"
    End If
    
    If KeyCode = 17 Then
        teclaPressionada = "CTRL"
    End If
    
    If KeyCode = 18 Then
        teclaPressionada = "ALT"
    End If
    
    If KeyCode = 20 Then
        teclaPressionada = "CAPSLOCK"
    End If
    
    If KeyCode = 33 Then
        teclaPressionada = "PAGEUP"
    End If
    
    If KeyCode = 34 Then
        teclaPressionada = "PAGEDOWN"
    End If
    
    If KeyCode = 35 Then
        teclaPressionada = "END"
    End If
    
    If KeyCode = 36 Then
        teclaPressionada = "HOME"
    End If
    
    If KeyCode = 37 Then
        teclaPressionada = "LEFT"
    End If
    
    If KeyCode = 38 Then
        teclaPressionada = "UP"
    End If
    
    If KeyCode = 39 Then
        teclaPressionada = "RIGHT"
    End If
    
    If KeyCode = 40 Then
        teclaPressionada = "DOWN"
    End If
    
    If KeyCode = 45 Then
        teclaPressionada = "INSERT"
    End If
    
    If KeyCode = 46 Then
        teclaPressionada = "DELETE"
    End If
    
    
    'Valida se a tecla já está sendo usada no Keys config:
    If teclaUsadaNoKeysConfig(teclaPressionada) Then
        MsgBox "A tecla " & teclaPressionada & " já foi atribuiída no Keys config"
        Exit Sub
    End If



    If KeyCode = vbKeyF1 Then
        lblTeclaAtalho.Caption = "F1"
        GoTo configuraTecla
    ElseIf KeyCode = vbKeyF2 Then
        lblTeclaAtalho.Caption = "F2"
        GoTo configuraTecla
    ElseIf KeyCode = vbKeyF3 Then
        lblTeclaAtalho.Caption = "F3"
        GoTo configuraTecla
    ElseIf KeyCode = vbKeyF4 Then
        lblTeclaAtalho.Caption = "F4"
        GoTo configuraTecla
    ElseIf KeyCode = vbKeyF5 Then
        lblTeclaAtalho.Caption = "F5"
        GoTo configuraTecla
    ElseIf KeyCode = vbKeyF6 Then
        lblTeclaAtalho.Caption = "F6"
        GoTo configuraTecla
    ElseIf KeyCode = vbKeyF7 Then
        lblTeclaAtalho.Caption = "F7"
        GoTo configuraTecla
    ElseIf KeyCode = vbKeyF8 Then
        lblTeclaAtalho.Caption = "F8"
        GoTo configuraTecla
    ElseIf KeyCode = vbKeyF9 Then
        lblTeclaAtalho.Caption = "F9"
        GoTo configuraTecla
    ElseIf KeyCode = vbKeyF10 Then
        lblTeclaAtalho.Caption = "F10"
        GoTo configuraTecla
    ElseIf KeyCode = vbKeyF11 Then
        lblTeclaAtalho.Caption = "F11"
        GoTo configuraTecla
    ElseIf KeyCode = vbKeyF12 Then
        lblTeclaAtalho.Caption = "F12"
        GoTo configuraTecla
    ElseIf KeyCode = 20 Then
        lblTeclaAtalho.Caption = "CAPSLOCK"
        GoTo configuraTecla
    ElseIf KeyCode = 16 Then
        lblTeclaAtalho.Caption = "SHIFT"
        GoTo configuraTecla
    ElseIf KeyCode = 17 Then
        lblTeclaAtalho.Caption = "CONTROL"
        GoTo configuraTecla
    ElseIf KeyCode = 18 Then
        lblTeclaAtalho.Caption = "ALT"
        
    ElseIf KeyCode = 37 Then
        lblTeclaAtalho.Caption = "LEFT"
    ElseIf KeyCode = 38 Then
        lblTeclaAtalho.Caption = "UP"
    ElseIf KeyCode = 39 Then
        lblTeclaAtalho.Caption = "RIGHT"
    ElseIf KeyCode = 40 Then
        lblTeclaAtalho.Caption = "DOWN"
        
    ElseIf KeyCode = 45 Then
        lblTeclaAtalho.Caption = "INSERT"
    ElseIf KeyCode = 46 Then
        lblTeclaAtalho.Caption = "DELETE"
    ElseIf KeyCode = 36 Then
        lblTeclaAtalho.Caption = "HOME"
    ElseIf KeyCode = 35 Then
        lblTeclaAtalho.Caption = "END"
    ElseIf KeyCode = 33 Then
        lblTeclaAtalho.Caption = "PAGEUP"
    ElseIf KeyCode = 34 Then
        lblTeclaAtalho.Caption = "PAGEDOWN"
        
    End If
    
    'lblTeclaAtalho.Caption = UCase(Chr$(KeyAscii))
    
configuraTecla:
            
            If KeyCode = vbKeyF1 Then
                config_tecla1 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla1 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla1 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla1 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla1 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla1 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla1 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla1 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla1 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla1 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla1 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla1 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla1 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla1 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla1 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla1 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla1 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla1 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla1 = "DOWN"
            ElseIf KeyCode = 45 Then
               config_tecla1 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla1 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla1 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla1 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla1 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla1 = "PAGEDOWN"
            End If
        
        If teclaAConfigurar_str = "up" Then
            frmConfig.txtKeysConfig_UP.Caption = config_tecla1
        ElseIf teclaAConfigurar_str = "down" Then
            frmConfig.txtKeysConfig_Down.Text = config_tecla1
        ElseIf teclaAConfigurar_str = "left" Then
            frmConfig.txtKeysConfig_Left.Text = config_tecla1
        ElseIf teclaAConfigurar_str = "right" Then
            frmConfig.txtKeysConfig_Right.Text = config_tecla1
        ElseIf teclaAConfigurar_str = "validation" Then
            frmConfig.txtKeysConfig_Validation.Text = config_tecla1
        ElseIf teclaAConfigurar_str = "backcancel" Then
            frmConfig.txtKeysConfig_BackCancel.Text = config_tecla1
        ElseIf teclaAConfigurar_str = "insertcoin" Then
            frmConfig.txtKeysConfig_InsertCoin.Text = config_tecla1
        ElseIf teclaAConfigurar_str = "switchzone" Then
            frmConfig.txtKeysConfig_SwitchZone.Text = config_tecla1
        ElseIf teclaAConfigurar_str = "config" Then
            frmConfig.txtKeysConfig_Config.Text = config_tecla1
        ElseIf teclaAConfigurar_str = "quit" Then
            frmConfig.txtKeysConfig_Quit.Text = config_tecla1
        ElseIf teclaAConfigurar_str = "previewtrack" Then
            frmConfig.txtKeysConfig_Preview.Text = config_tecla1
        ElseIf teclaAConfigurar_str = "stop" Then
            frmConfig.txtKeysConfig_Stop.Text = config_tecla1
        ElseIf teclaAConfigurar_str = "play" Then
            frmConfig.txtKeysConfig_Play.Text = config_tecla1
        ElseIf teclaAConfigurar_str = "pause" Then
            frmConfig.txtKeysConfig_Pause.Text = config_tecla1
        ElseIf teclaAConfigurar_str = "nexttrack" Then
            frmConfig.txtKeysConfig_Next.Text = config_tecla1
        End If
        
        Unload Me
End Sub

