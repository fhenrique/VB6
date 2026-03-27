Attribute VB_Name = "Global"
Option Explicit

'VARIAVEIS USADAS NO FRMCONFIG
Public teclaAConfigurar As Integer
Public teclaAConfigurar_str As String

Public config_track1_tecla As String
Public config_track2_tecla As String
Public config_track3_tecla As String
Public config_track4_tecla As String
Public config_track5_tecla As String
Public config_track6_tecla As String
Public config_track7_tecla As String
Public config_track8_tecla As String
Public config_track9_tecla As String
Public config_track10_tecla As String
Public config_track11_tecla As String
Public config_track12_tecla As String
Public config_track13_tecla As String
Public config_track14_tecla As String
Public config_track15_tecla As String
Public config_track16_tecla As String
Public config_track17_tecla As String
Public config_track18_tecla As String
Public config_track19_tecla As String
Public config_track20_tecla As String
Public config_tecla_config As String

'VARIAVEIS USADAS NA FUNCAO DE LEITURA DO CONFIG.INI
Public global_ini_zona1_qtdCoins As String

Public global_ini_config_SwitchZone As String
Public global_ini_config_Config As String
Public global_ini_config_Up As String
Public global_ini_config_Down As String
Public global_ini_config_Left As String
Public global_ini_config_Right As String
Public global_ini_config_Validation As String
Public global_ini_config_BackCancel As String
Public global_ini_config_InsertCoin As String
Public global_ini_config_Quit As String
Public global_ini_config_PreviewTrack As String
Public global_ini_config_Stop As String
Public global_ini_config_Play As String
Public global_ini_config_Pause As String
Public global_ini_config_NextTrack As String

Public global_ini_zona2_tecla1 As String
Public global_ini_zona2_tecla2 As String
Public global_ini_zona2_tecla3 As String
Public global_ini_zona2_tecla4 As String
Public global_ini_zona2_tecla5 As String
Public global_ini_zona2_tecla6 As String
Public global_ini_zona2_tecla7 As String
Public global_ini_zona2_tecla8 As String
Public global_ini_zona2_tecla9 As String
Public global_ini_zona2_tecla10 As String
Public global_ini_zona2_tecla11 As String
Public global_ini_zona2_tecla12 As String
Public global_ini_zona2_tecla13 As String
Public global_ini_zona2_tecla14 As String
Public global_ini_zona2_tecla15 As String
Public global_ini_zona2_tecla16 As String
Public global_ini_zona2_tecla17 As String
Public global_ini_zona2_tecla18 As String
Public global_ini_zona2_tecla19 As String
Public global_ini_zona2_tecla20 As String

'VARIÁVEIS USADAS PARA AS MÚSICAS
Public global_MusicaAtual_1 As String
Public global_MusicaAtual_2 As String
Public global_MusicaAtual_3 As String
Public global_MusicaAtual_4 As String
Public global_MusicaAtual_5 As String
Public global_MusicaAtual_6 As String
Public global_MusicaAtual_7 As String
Public global_MusicaAtual_8 As String
Public global_MusicaAtual_9 As String
Public global_MusicaAtual_10 As String
Public global_MusicaAtual_11 As String
Public global_MusicaAtual_12 As String
Public global_MusicaAtual_13 As String
Public global_MusicaAtual_14 As String
Public global_MusicaAtual_15 As String
Public global_MusicaAtual_16 As String
Public global_MusicaAtual_17 As String
Public global_MusicaAtual_18 As String
Public global_MusicaAtual_19 As String
Public global_MusicaAtual_20 As String

'Função API para leitura das informações do arquivo INI
Public Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
 
'Função API para escrever informações no arquivo INI
Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long
 
'Buscar informações no arquivo INI
Public Function GetINISetting(ByVal sHeading As String, _
    ByVal sKey As String, _
    sINIFileName) As String
     
    Const cparmLen = 50
    Dim sReturn As String * cparmLen
    Dim sDefault As String * cparmLen
    Dim lLength As Long
    lLength = GetPrivateProfileString(sHeading, _
            sKey, _
            sDefault, _
            sReturn, _
            cparmLen, _
            sINIFileName)
             
    GetINISetting = Mid(sReturn, 1, lLength)
End Function
 
'Salvar informações no arquivo INI
Public Function PutINISetting(ByVal sHeading As String, _
    ByVal sKey As String, _
    ByVal sSetting As String, _
    sINIFileName) As Boolean
     
    Const cparmLen = 50
    Dim sReturn As String * cparmLen
    Dim sDefault As String * cparmLen
    Dim aLength As Long
    aLength = WritePrivateProfileString(sHeading, _
        sKey, _
        sSetting, _
        sINIFileName)
         
    PutINISetting = True
End Function

Public Sub gravaIni(cabecalho As String, chave As String, valor As String)
   Dim ConfigFile As String
   ConfigFile = App.Path & "\config.INI"
   
   PutINISetting cabecalho, chave, valor, ConfigFile
   
   
End Sub


Public Function teclaUsadaNoKeysConfig(strTecla As String) As Boolean
    teclaUsadaNoKeysConfig = False

    If frmConfig.txtKeysConfig_UP.Caption = strTecla Then
        teclaUsadaNoKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.txtKeysConfig_Down.Text = strTecla Then
        teclaUsadaNoKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.txtKeysConfig_Left.Text = strTecla Then
        teclaUsadaNoKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.txtKeysConfig_Right.Text = strTecla Then
        teclaUsadaNoKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.txtKeysConfig_Validation.Text = strTecla Then
        teclaUsadaNoKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.txtKeysConfig_BackCancel.Text = strTecla Then
        teclaUsadaNoKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.txtKeysConfig_InsertCoin.Text = strTecla Then
        teclaUsadaNoKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.txtKeysConfig_SwitchZone.Text = strTecla Then
        teclaUsadaNoKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.txtKeysConfig_Config.Text = strTecla Then
        teclaUsadaNoKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.txtKeysConfig_Quit.Text = strTecla Then
        teclaUsadaNoKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.txtKeysConfig_Preview.Text = strTecla Then
        teclaUsadaNoKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.txtKeysConfig_Stop.Text = strTecla Then
        teclaUsadaNoKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.txtKeysConfig_Play.Text = strTecla Then
        teclaUsadaNoKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.txtKeysConfig_Pause.Text = strTecla Then
        teclaUsadaNoKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.txtKeysConfig_Next.Text = strTecla Then
        teclaUsadaNoKeysConfig = True
        Exit Function
    End If
    
    

    
End Function

Public Function teclaUsadaNoThemeKeysConfig(strTecla As String) As Boolean
    teclaUsadaNoThemeKeysConfig = False

    If frmConfig.lblThemeKeysConfig_1.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_2.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_3.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_4.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_5.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_6.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_7.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_8.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_9.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_10.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_11.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_12.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_13.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_14.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_15.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_16.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_17.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_18.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_19.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If
    
    If frmConfig.lblThemeKeysConfig_20.Caption = strTecla Then
        teclaUsadaNoThemeKeysConfig = True
        Exit Function
    End If

    
End Function


Public Function teclaUsadaSelf(numTeclaExcessao As Integer, teclaValor As String) As Boolean
    
    teclaUsadaSelf = False
        
    If numTeclaExcessao <> 1 Then
        If frmConfig.lblThemeKeysConfig_1.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 2 Then
        If frmConfig.lblThemeKeysConfig_2.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 3 Then
        If frmConfig.lblThemeKeysConfig_3.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 4 Then
        If frmConfig.lblThemeKeysConfig_4.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 5 Then
        If frmConfig.lblThemeKeysConfig_5.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 6 Then
        If frmConfig.lblThemeKeysConfig_6.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 7 Then
        If frmConfig.lblThemeKeysConfig_7.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 8 Then
        If frmConfig.lblThemeKeysConfig_8.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 9 Then
        If frmConfig.lblThemeKeysConfig_9.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 10 Then
        If frmConfig.lblThemeKeysConfig_10.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 11 Then
        If frmConfig.lblThemeKeysConfig_11.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 12 Then
        If frmConfig.lblThemeKeysConfig_12.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 13 Then
        If frmConfig.lblThemeKeysConfig_13.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 14 Then
        If frmConfig.lblThemeKeysConfig_14.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 15 Then
        If frmConfig.lblThemeKeysConfig_15.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 16 Then
        If frmConfig.lblThemeKeysConfig_16.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 17 Then
        If frmConfig.lblThemeKeysConfig_17.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 18 Then
        If frmConfig.lblThemeKeysConfig_18.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 19 Then
        If frmConfig.lblThemeKeysConfig_19.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
    If numTeclaExcessao <> 20 Then
        If frmConfig.lblThemeKeysConfig_20.Caption = teclaValor Then
            teclaUsadaSelf = True
            Exit Function
        End If
    End If
    
End Function

Public Sub lerArquivoConfigIni()
    Dim ConfigFile As String

    ConfigFile = App.Path & "\config.INI"
    
    'Leitura dos atalhos do KEYS CONFIG
    global_ini_config_SwitchZone = GetINISetting("KEYSCONFIG", "SwitchZone", ConfigFile)
    global_ini_config_Config = GetINISetting("KEYSCONFIG", "Config", ConfigFile)
    global_ini_config_Up = GetINISetting("KEYSCONFIG", "Up", ConfigFile)
    global_ini_config_Down = GetINISetting("KEYSCONFIG", "Down", ConfigFile)
    global_ini_config_Left = GetINISetting("KEYSCONFIG", "Left", ConfigFile)
    global_ini_config_Right = GetINISetting("KEYSCONFIG", "Right", ConfigFile)
    global_ini_config_Validation = GetINISetting("KEYSCONFIG", "Validation", ConfigFile)
    global_ini_config_BackCancel = GetINISetting("KEYSCONFIG", "BackCancel", ConfigFile)
    global_ini_config_InsertCoin = GetINISetting("KEYSCONFIG", "InsertCoin", ConfigFile)
    global_ini_config_Quit = GetINISetting("KEYSCONFIG", "Quit", ConfigFile)
    global_ini_config_PreviewTrack = GetINISetting("KEYSCONFIG", "PreviewTrack", ConfigFile)
    global_ini_config_Stop = GetINISetting("KEYSCONFIG", "Stop", ConfigFile)
    global_ini_config_Play = GetINISetting("KEYSCONFIG", "Play", ConfigFile)
    global_ini_config_Pause = GetINISetting("KEYSCONFIG", "Pause", ConfigFile)
    global_ini_config_NextTrack = GetINISetting("KEYSCONFIG", "NextTrack", ConfigFile)
    
    'Leitura da quantidade de fichas
    global_ini_zona1_qtdCoins = GetINISetting("FICHAS", "quantidade", ConfigFile)
    
    'Leitura dos atalhos do THEME KEYS CONFIG
    global_ini_zona2_tecla1 = GetINISetting("THEMEKEYSCONFIG", "track1", ConfigFile)
    global_ini_zona2_tecla2 = GetINISetting("THEMEKEYSCONFIG", "track2", ConfigFile)
    global_ini_zona2_tecla3 = GetINISetting("THEMEKEYSCONFIG", "track3", ConfigFile)
    global_ini_zona2_tecla4 = GetINISetting("THEMEKEYSCONFIG", "track4", ConfigFile)
    global_ini_zona2_tecla5 = GetINISetting("THEMEKEYSCONFIG", "track5", ConfigFile)
    global_ini_zona2_tecla6 = GetINISetting("THEMEKEYSCONFIG", "track6", ConfigFile)
    global_ini_zona2_tecla7 = GetINISetting("THEMEKEYSCONFIG", "track7", ConfigFile)
    global_ini_zona2_tecla8 = GetINISetting("THEMEKEYSCONFIG", "track8", ConfigFile)
    global_ini_zona2_tecla9 = GetINISetting("THEMEKEYSCONFIG", "track9", ConfigFile)
    global_ini_zona2_tecla10 = GetINISetting("THEMEKEYSCONFIG", "track10", ConfigFile)
    global_ini_zona2_tecla11 = GetINISetting("THEMEKEYSCONFIG", "track11", ConfigFile)
    global_ini_zona2_tecla12 = GetINISetting("THEMEKEYSCONFIG", "track12", ConfigFile)
    global_ini_zona2_tecla13 = GetINISetting("THEMEKEYSCONFIG", "track13", ConfigFile)
    global_ini_zona2_tecla14 = GetINISetting("THEMEKEYSCONFIG", "track14", ConfigFile)
    global_ini_zona2_tecla15 = GetINISetting("THEMEKEYSCONFIG", "track15", ConfigFile)
    global_ini_zona2_tecla16 = GetINISetting("THEMEKEYSCONFIG", "track16", ConfigFile)
    global_ini_zona2_tecla17 = GetINISetting("THEMEKEYSCONFIG", "track17", ConfigFile)
    global_ini_zona2_tecla18 = GetINISetting("THEMEKEYSCONFIG", "track18", ConfigFile)
    global_ini_zona2_tecla19 = GetINISetting("THEMEKEYSCONFIG", "track19", ConfigFile)
    global_ini_zona2_tecla20 = GetINISetting("THEMEKEYSCONFIG", "track20", ConfigFile)
    
    
    
    
End Sub




















