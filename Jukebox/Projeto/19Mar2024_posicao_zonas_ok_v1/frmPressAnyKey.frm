VERSION 5.00
Begin VB.Form frmPressAnyKey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Press any Key"
   ClientHeight    =   2070
   ClientLeft      =   4185
   ClientTop       =   4275
   ClientWidth     =   4590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1875
      TabIndex        =   1
      Top             =   825
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "to manage the alphabetical shortcuts depending on the number of page, click on the ""page up"" button"
      Height          =   540
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   4545
   End
End
Attribute VB_Name = "frmPressAnyKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim teclaPressionada As String
    
    If (KeyAscii <> 27) And (KeyAscii <> 16) And (KeyAscii <> 17) And (KeyAscii <> 18) And (KeyAscii <> 20) And (KeyAscii <> 37) And (KeyAscii <> 38) And (KeyAscii <> 39) And (KeyAscii <> 40) And (KeyAscii <> 32) Then
        teclaPressionada = UCase(Chr$(KeyAscii))
    ElseIf KeyAscii = 32 Then 'SPACEBAR
        teclaPressionada = "SPACEBAR"
    End If
    
'    If KeyAscii = vbKeyF1 Then
'        teclaPressionada = "F1"
'    End If
'
'    If KeyAscii = vbKeyF2 Then
'        teclaPressionada = "F2"
'    End If
'
'    If KeyAscii = vbKeyF3 Then
'        teclaPressionada = "F3"
'    End If
'
'    If KeyAscii = vbKeyF4 Then
'        teclaPressionada = "F4"
'    End If
'
'    If KeyAscii = vbKeyF5 Then
'        teclaPressionada = "F5"
'    End If
'
'    If KeyAscii = vbKeyF6 Then
'        teclaPressionada = "F6"
'    End If
'
'    If KeyAscii = vbKeyF7 Then
'        teclaPressionada = "F7"
'    End If
'
'    If KeyAscii = vbKeyF8 Then
'        teclaPressionada = "F8"
'    End If
'
'    If KeyAscii = vbKeyF9 Then
'        teclaPressionada = "F9"
'    End If
'
'    If KeyAscii = vbKeyF10 Then
'        teclaPressionada = "F10"
'    End If
'
'    If KeyAscii = vbKeyF11 Then
'        teclaPressionada = "F11"
'    End If
'
'    If KeyAscii = vbKeyF12 Then
'        teclaPressionada = "F12"
'    End If
    
    If KeyAscii = 27 Then
        teclaPressionada = "ESC"
    End If
    
    'Valida se a tecla já está sendo usada no Keys config:
    If teclaUsadaNoKeysConfig(teclaPressionada) Then
        MsgBox "Tecla " & teclaPressionada & " já foi atribuiída no Keys config"
        Exit Sub
    End If
    
    'Valida se a tecla já está sendo usada no próprio Theme Keys config
    If teclaAConfigurar = 1 Then
        If teclaUsadaSelf(1, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 2 Then
        If teclaUsadaSelf(2, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 3 Then
        If teclaUsadaSelf(3, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 4 Then
        If teclaUsadaSelf(4, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 5 Then
        If teclaUsadaSelf(5, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 6 Then
        If teclaUsadaSelf(6, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 7 Then
        If teclaUsadaSelf(7, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 8 Then
        If teclaUsadaSelf(8, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 9 Then
        If teclaUsadaSelf(9, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 10 Then
        If teclaUsadaSelf(10, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 11 Then
        If teclaUsadaSelf(11, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 12 Then
        If teclaUsadaSelf(12, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 13 Then
        If teclaUsadaSelf(13, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 14 Then
        If teclaUsadaSelf(14, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 15 Then
        If teclaUsadaSelf(15, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 16 Then
        If teclaUsadaSelf(16, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 17 Then
        If teclaUsadaSelf(17, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 18 Then
        If teclaUsadaSelf(18, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 19 Then
        If teclaUsadaSelf(19, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 20 Then
        If teclaUsadaSelf(20, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuida aqui mesmo no Theme keys config"
            Exit Sub
        End If
    End If
    
    
    If KeyAscii = 8 Then
        lblTeclaAtalho.Caption = "BACKSPACE"
        GoTo configuraTecla
    End If
    
    
    If KeyAscii = 9 Then
        lblTeclaAtalho.Caption = "TAB"
        GoTo configuraTecla
    End If
    
    If KeyAscii = 13 Then
        lblTeclaAtalho.Caption = "ENTER"
        GoTo configuraTecla
    End If
    
    If KeyAscii = 27 Then
        lblTeclaAtalho.Caption = "ESC"
        GoTo configuraTecla
    End If
    
    If KeyAscii = 32 Then
        lblTeclaAtalho.Caption = "SPACEBAR"
        GoTo configuraTecla
    End If
    
    
    
    
    lblTeclaAtalho.Caption = UCase(Chr$(KeyAscii))
    
configuraTecla:
        
        If teclaAConfigurar = 1 Then
            frmConfig.lblThemeKeysConfig_1.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 2 Then
            frmConfig.lblThemeKeysConfig_2.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 3 Then
            frmConfig.lblThemeKeysConfig_3.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 4 Then
            frmConfig.lblThemeKeysConfig_4.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 5 Then
            frmConfig.lblThemeKeysConfig_5.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 6 Then
            frmConfig.lblThemeKeysConfig_6.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 7 Then
            frmConfig.lblThemeKeysConfig_7.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 8 Then
            frmConfig.lblThemeKeysConfig_8.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 9 Then
            frmConfig.lblThemeKeysConfig_9.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 10 Then
            frmConfig.lblThemeKeysConfig_10.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 11 Then
            frmConfig.lblThemeKeysConfig_11.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 12 Then
            frmConfig.lblThemeKeysConfig_12.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 13 Then
            frmConfig.lblThemeKeysConfig_13.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 14 Then
            frmConfig.lblThemeKeysConfig_14.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 15 Then
            frmConfig.lblThemeKeysConfig_15.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 16 Then
            frmConfig.lblThemeKeysConfig_16.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 17 Then
            frmConfig.lblThemeKeysConfig_17.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 18 Then
            frmConfig.lblThemeKeysConfig_18.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 19 Then
            frmConfig.lblThemeKeysConfig_19.Caption = lblTeclaAtalho.Caption
        End If
        
        If teclaAConfigurar = 20 Then
            frmConfig.lblThemeKeysConfig_20.Caption = lblTeclaAtalho.Caption
        End If
        
        Unload Me
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim teclaPressionada As String
    
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
    
    If KeyCode = 16 Then 'SHIFT
        teclaPressionada = "SHIFT"
    End If
    
    If KeyCode = 17 Then 'CTRL
        teclaPressionada = "CTRL"
    End If
    
    If KeyCode = 18 Then 'ALT
        teclaPressionada = "ALT"
    End If
    
    If KeyCode = 33 Then 'PAGEUP
        teclaPressionada = "PAGEUP"
    End If
    
    If KeyCode = 34 Then 'PAGEDOWN
        teclaPressionada = "PAGEDOWN"
    End If
    
    
    If KeyCode = 35 Then 'END
        teclaPressionada = "END"
    End If
    
    If KeyCode = 36 Then 'HOME
        teclaPressionada = "HOME"
    End If
    
    If KeyCode = 37 Then 'LEFT
        teclaPressionada = "LEFT"
    End If
    
    If KeyCode = 38 Then 'Up
        teclaPressionada = "UP"
    End If
    
    If KeyCode = 39 Then 'Right
        teclaPressionada = "RIGHT"
    End If
    
    If KeyCode = 40 Then 'DOWN
        teclaPressionada = "DOWN"
    End If
    
    If KeyCode = 45 Then 'INSERT
        teclaPressionada = "INSERT"
    End If
    
    If KeyCode = 46 Then 'DELETE
        teclaPressionada = "DELETE"
    End If
    
    
    'Valida se a tecla já está sendo usada no Keys config:
    If teclaUsadaNoKeysConfig(teclaPressionada) Then
        MsgBox "A tecla " & teclaPressionada & " já foi atribuiída no Keys config"
        Exit Sub
    End If
    
    'Valida se a tecla já está sendo usada no próprio Theme Keys config
    If teclaAConfigurar = 1 Then
        If teclaUsadaSelf(1, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 2 Then
        If teclaUsadaSelf(2, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 3 Then
        If teclaUsadaSelf(3, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 4 Then
        If teclaUsadaSelf(4, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 5 Then
        If teclaUsadaSelf(5, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 6 Then
        If teclaUsadaSelf(6, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 7 Then
        If teclaUsadaSelf(7, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 8 Then
        If teclaUsadaSelf(8, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 9 Then
        If teclaUsadaSelf(9, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 10 Then
        If teclaUsadaSelf(10, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 11 Then
        If teclaUsadaSelf(11, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 12 Then
        If teclaUsadaSelf(12, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 13 Then
        If teclaUsadaSelf(13, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 14 Then
        If teclaUsadaSelf(14, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 15 Then
        If teclaUsadaSelf(15, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 16 Then
        If teclaUsadaSelf(16, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 17 Then
        If teclaUsadaSelf(17, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 18 Then
        If teclaUsadaSelf(18, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 19 Then
        If teclaUsadaSelf(19, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
    ElseIf teclaAConfigurar = 20 Then
        If teclaUsadaSelf(20, teclaPressionada) Then
            MsgBox "A tecla " & teclaPressionada & " já foi atribuída aqui mesmo no Theme keys config"
            Exit Sub
        End If
        
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
        GoTo configuraTecla
    
    ElseIf KeyCode = 37 Then
        lblTeclaAtalho.Caption = "LEFT"
        GoTo configuraTecla
    ElseIf KeyCode = 38 Then
        lblTeclaAtalho.Caption = "UP"
        GoTo configuraTecla
    ElseIf KeyCode = 39 Then
        lblTeclaAtalho.Caption = "RIGHT"
        GoTo configuraTecla
    ElseIf KeyCode = 40 Then
        lblTeclaAtalho.Caption = "DOWN"
        GoTo configuraTecla
        
    ElseIf KeyCode = 45 Then
        lblTeclaAtalho.Caption = "INSERT"
        GoTo configuraTecla
    ElseIf KeyCode = 46 Then
        lblTeclaAtalho.Caption = "DELETE"
        GoTo configuraTecla
    ElseIf KeyCode = 36 Then
        lblTeclaAtalho.Caption = "HOME"
        GoTo configuraTecla
    ElseIf KeyCode = 35 Then
        lblTeclaAtalho.Caption = "END"
        GoTo configuraTecla
    ElseIf KeyCode = 33 Then
        lblTeclaAtalho.Caption = "PAGEUP"
        GoTo configuraTecla
    ElseIf KeyCode = 34 Then
        lblTeclaAtalho.Caption = "PAGEDOWN"
        GoTo configuraTecla
        
    End If
    
    lblTeclaAtalho.Caption = UCase(Chr$(KeyAscii))
    
configuraTecla:
        If teclaAConfigurar = 1 Then
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
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 2 Then
            If KeyCode = vbKeyF1 Then
               config_tecla2 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla2 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla2 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla2 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla2 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla2 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla2 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla2 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla2 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla2 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla2 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla2 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla2 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla2 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla2 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla2 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla2 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla2 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla2 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla2 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla2 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla2 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla2 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla2 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla2 = "PAGEDOWN"
               
            End If
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 3 Then
            If KeyCode = vbKeyF1 Then
               config_tecla3 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla3 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla3 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla3 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla3 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla3 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla3 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla3 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla3 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla3 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla3 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla3 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla3 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla3 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla3 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla3 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla3 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla3 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla3 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla3 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla3 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla3 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla3 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla3 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla3 = "PAGEDOWN"
               
            End If
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 4 Then
            If KeyCode = vbKeyF1 Then
               config_tecla4 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla4 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla4 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla4 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla4 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla4 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla4 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla4 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla4 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla4 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla4 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla4 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla4 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla4 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla4 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla4 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla4 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla4 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla4 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla4 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla4 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla4 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla4 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla4 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla4 = "PAGEDOWN"
               
            End If
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 5 Then
            If KeyCode = vbKeyF1 Then
               config_tecla5 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla5 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla5 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla5 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla5 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla5 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla5 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla5 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla5 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla5 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla5 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla5 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla5 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla5 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla5 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla5 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla5 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla5 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla5 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla5 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla5 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla5 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla5 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla5 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla5 = "PAGEDOWN"
            End If
            
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 6 Then
            If KeyCode = vbKeyF1 Then
               config_tecla6 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla6 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla6 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla6 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla6 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla6 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla6 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla6 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla6 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla6 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla6 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla6 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla6 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla6 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla6 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla6 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla6 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla6 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla6 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla6 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla6 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla6 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla6 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla6 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla6 = "PAGEDOWN"
            End If
            
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 7 Then
            If KeyCode = vbKeyF1 Then
               config_tecla7 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla7 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla7 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla7 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla7 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla7 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla7 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla7 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla7 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla7 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla7 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla7 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla7 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla7 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla7 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla7 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla7 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla7 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla7 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla7 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla7 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla7 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla7 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla7 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla7 = "PAGEDOWN"
               
            End If
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 8 Then
            If KeyCode = vbKeyF1 Then
               config_tecla8 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla8 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla8 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla8 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla8 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla8 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla8 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla8 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla8 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla8 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla8 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla8 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla8 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla8 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla8 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla8 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla8 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla8 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla8 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla8 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla8 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla8 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla8 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla8 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla8 = "PAGEDOWN"
               
            End If
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 9 Then
            If KeyCode = vbKeyF1 Then
               config_tecla9 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla9 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla9 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla9 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla9 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla9 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla9 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla9 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla9 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla9 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla9 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla9 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla9 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla9 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla9 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla9 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla9 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla9 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla9 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla9 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla9 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla9 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla9 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla9 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla9 = "PAGEDOWN"
            End If
            
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 10 Then
            If KeyCode = vbKeyF1 Then
               config_tecla10 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla10 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla10 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla10 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla10 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla10 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla10 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla10 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla10 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla10 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla10 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla10 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla10 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla10 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla10 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla10 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla10 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla10 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla10 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla10 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla10 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla10 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla10 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla10 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla10 = "PAGEDOWN"
            End If
            
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 11 Then
            If KeyCode = vbKeyF1 Then
               config_tecla11 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla11 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla11 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla11 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla11 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla11 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla11 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla11 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla11 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla11 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla11 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla11 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla11 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla11 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla11 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla11 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla11 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla11 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla11 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla11 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla11 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla11 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla11 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla11 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla11 = "PAGEDOWN"
            End If
            
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 12 Then
            If KeyCode = vbKeyF1 Then
               config_tecla12 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla12 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla12 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla12 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla12 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla12 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla12 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla12 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla12 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla12 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla12 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla12 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla12 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla12 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla12 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla12 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla12 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla12 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla12 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla12 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla12 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla12 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla12 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla12 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla12 = "PAGEDOWN"
            End If
            
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 13 Then
            If KeyCode = vbKeyF1 Then
               config_tecla13 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla13 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla13 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla13 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla13 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla13 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla13 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla13 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla13 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla13 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla13 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla13 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla13 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla13 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla13 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla13 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla13 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla13 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla13 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla13 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla13 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla13 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla13 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla13 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla13 = "PAGEDOWN"
            End If
            
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 14 Then
            If KeyCode = vbKeyF1 Then
               config_tecla14 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla14 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla14 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla14 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla14 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla14 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla14 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla14 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla14 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla14 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla14 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla14 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla14 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla14 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla14 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla14 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla14 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla14 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla14 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla14 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla14 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla14 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla14 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla14 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla14 = "PAGEDOWN"
               
            End If
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 15 Then
            If KeyCode = vbKeyF1 Then
               config_tecla15 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla15 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla15 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla15 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla15 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla15 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla15 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla15 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla15 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla15 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla15 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla15 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla15 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla15 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla15 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla15 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla15 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla15 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla15 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla15 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla15 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla15 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla15 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla15 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla15 = "PAGEDOWN"
               
            End If
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 16 Then
            If KeyCode = vbKeyF1 Then
               config_tecla16 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla16 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla16 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla16 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla16 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla16 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla16 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla16 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla16 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla16 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla16 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla16 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla16 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla16 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla16 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla16 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla16 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla16 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla16 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla16 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla16 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla16 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla16 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla16 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla16 = "PAGEDOWN"
               
            End If
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 17 Then
            If KeyCode = vbKeyF1 Then
               config_tecla17 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla17 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla17 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla17 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla17 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla17 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla17 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla17 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla17 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla17 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla17 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla17 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla17 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla17 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla17 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla17 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla17 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla17 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla17 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla17 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla17 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla17 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla17 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla17 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla17 = "PAGEDOWN"
               
            End If
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 18 Then
            If KeyCode = vbKeyF1 Then
               config_tecla18 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla18 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla18 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla18 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla18 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla18 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla18 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla18 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla18 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla18 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla18 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla18 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla18 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla18 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla18 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla18 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla18 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla18 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla18 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla18 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla18 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla18 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla18 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla18 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla18 = "PAGEDOWN"
               
            End If
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 19 Then
            If KeyCode = vbKeyF1 Then
               config_tecla19 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla19 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla19 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla19 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla19 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla19 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla19 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla19 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla19 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla19 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla19 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla19 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla19 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla19 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla19 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla19 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla19 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla19 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla19 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla19 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla19 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla19 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla19 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla19 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla19 = "PAGEDOWN"
            End If
            
            GoTo finaliza
        End If
        
        If teclaAConfigurar = 20 Then
            If KeyCode = vbKeyF1 Then
               config_tecla20 = "F1"
            ElseIf KeyCode = vbKeyF2 Then
               config_tecla20 = "F2"
            ElseIf KeyCode = vbKeyF3 Then
               config_tecla20 = "F3"
            ElseIf KeyCode = vbKeyF4 Then
               config_tecla20 = "F4"
            ElseIf KeyCode = vbKeyF5 Then
               config_tecla20 = "F5"
            ElseIf KeyCode = vbKeyF6 Then
               config_tecla20 = "F6"
            ElseIf KeyCode = vbKeyF7 Then
               config_tecla20 = "F7"
            ElseIf KeyCode = vbKeyF8 Then
               config_tecla20 = "F8"
            ElseIf KeyCode = vbKeyF9 Then
               config_tecla20 = "F9"
            ElseIf KeyCode = vbKeyF10 Then
               config_tecla20 = "F10"
            ElseIf KeyCode = vbKeyF11 Then
               config_tecla20 = "F11"
            ElseIf KeyCode = vbKeyF12 Then
               config_tecla20 = "F12"
            ElseIf KeyCode = 16 Then
               config_tecla20 = "SHIFT"
            ElseIf KeyCode = 17 Then
               config_tecla20 = "CTRL"
            ElseIf KeyCode = 18 Then
               config_tecla20 = "ALT"
               
            ElseIf KeyCode = 37 Then
               config_tecla20 = "LEFT"
            ElseIf KeyCode = 38 Then
               config_tecla20 = "UP"
            ElseIf KeyCode = 39 Then
               config_tecla20 = "RIGHT"
            ElseIf KeyCode = 40 Then
               config_tecla20 = "DOWN"
               
            ElseIf KeyCode = 45 Then
               config_tecla20 = "INSERT"
            ElseIf KeyCode = 46 Then
               config_tecla20 = "DELETE"
            ElseIf KeyCode = 36 Then
               config_tecla20 = "HOME"
            ElseIf KeyCode = 35 Then
               config_tecla20 = "END"
            ElseIf KeyCode = 33 Then
               config_tecla20 = "PAGEUP"
            ElseIf KeyCode = 34 Then
               config_tecla20 = "PAGEDOWN"
            End If
            
        End If
        
finaliza:
        If teclaAConfigurar = 1 Then
            frmConfig.lblThemeKeysConfig_1.Caption = config_tecla1
        ElseIf teclaAConfigurar = 2 Then
            frmConfig.lblThemeKeysConfig_2.Caption = config_tecla2
        ElseIf teclaAConfigurar = 3 Then
            frmConfig.lblThemeKeysConfig_3.Caption = config_tecla3
        ElseIf teclaAConfigurar = 4 Then
            frmConfig.lblThemeKeysConfig_4.Caption = config_tecla4
        ElseIf teclaAConfigurar = 5 Then
            frmConfig.lblThemeKeysConfig_5.Caption = config_tecla5
        ElseIf teclaAConfigurar = 6 Then
            frmConfig.lblThemeKeysConfig_6.Caption = config_tecla6
        ElseIf teclaAConfigurar = 7 Then
            frmConfig.lblThemeKeysConfig_7.Caption = config_tecla7
        ElseIf teclaAConfigurar = 8 Then
            frmConfig.lblThemeKeysConfig_8.Caption = config_tecla8
        ElseIf teclaAConfigurar = 9 Then
            frmConfig.lblThemeKeysConfig_9.Caption = config_tecla9
        ElseIf teclaAConfigurar = 10 Then
            frmConfig.lblThemeKeysConfig_10.Caption = config_tecla10
        ElseIf teclaAConfigurar = 11 Then
            frmConfig.lblThemeKeysConfig_11.Caption = config_tecla11
        ElseIf teclaAConfigurar = 12 Then
            frmConfig.lblThemeKeysConfig_12.Caption = config_tecla12
        ElseIf teclaAConfigurar = 13 Then
            frmConfig.lblThemeKeysConfig_13.Caption = config_tecla13
        ElseIf teclaAConfigurar = 14 Then
            frmConfig.lblThemeKeysConfig_14.Caption = config_tecla14
        ElseIf teclaAConfigurar = 15 Then
            frmConfig.lblThemeKeysConfig_15.Caption = config_tecla15
        ElseIf teclaAConfigurar = 16 Then
            frmConfig.lblThemeKeysConfig_16.Caption = config_tecla16
        ElseIf teclaAConfigurar = 17 Then
            frmConfig.lblThemeKeysConfig_17.Caption = config_tecla17
        ElseIf teclaAConfigurar = 18 Then
            frmConfig.lblThemeKeysConfig_18.Caption = config_tecla18
        ElseIf teclaAConfigurar = 19 Then
            frmConfig.lblThemeKeysConfig_19.Caption = config_tecla19
        ElseIf teclaAConfigurar = 20 Then
            frmConfig.lblThemeKeysConfig_20.Caption = config_tecla20
        End If
        
        Unload Me
    
End Sub



