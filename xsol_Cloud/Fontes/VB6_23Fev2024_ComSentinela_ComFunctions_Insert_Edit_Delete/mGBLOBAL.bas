Attribute VB_Name = "GBLOBAL"
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
As String, ByVal lpKeyName As String, ByVal lpDefault As _
String, ByVal lpReturnedString As String, ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As _
  Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu _
  As Long) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As _
  Long, ByVal bRevert As Long) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, _
  ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Default_Printer As String

Public Const MF_BYPOSITION = &H400
Public Const MF_REMOVE = &H1000

Option Explicit
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Const COLOR_BTNFACE = &H8000000F  'Button
Public Const ADMIN_PROFILE = "ADMIN"

Private conn As ADODB.Connection
Public Nome_Cli_Atend As String
Public W_DBname As String
Public Image_Folder As String
Public Cod_Cli As Integer
Public Cod_Address As String
Public Cod_Equipment As String

Public Type User
    Login As String
    Profile As String
End Type

Public usuario As User

Public cloud_conexao As New ADODB.Connection
Public cloud_Servidor As String
Public cloud_Database As String
Public cloud_usuario As String
Public cloud_senha As String
Public cloud_ConfigFile As String
Public cloud_Path As String
Public cloud_FileName As String
Public cloud_Image_Folder As String
Public cloud_status_conexao As Integer
Public cloud_img_online As String
Public cloud_img_offline As String
Public cloud_registro As New ADODB.Recordset
Public cloud_select As String
Public cloud_sql As String
Public cloud_campos As String
Public cloud_values As String
Public cloud_qtd As Integer

Public Sub Set_Connection()
   Set conn = New ADODB.Connection
   conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
End Sub

Public Sub ConnectionRefresh()
    'conn.Close
    'Set conn = Nothing
    'Set conn = New ADODB.Connection
    'conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
End Sub

Public Function Get_Field_Val(ByVal val As Object) As String
    Get_Field_Val = UCase(IIf(IsNull(val), "", val))
End Function

Public Function Set_Field_Val(ByVal val As Object) As String
    Set_Field_Val = IIf(val = "", vbNullString, UCase(val))
End Function

Public Function Return_Recordset(strSql As String) As ADODB.Recordset
    
On Error GoTo conecta
    
retenta:
    Set cloud_registro = cloud_conexao.Execute(strSql)
    Set Return_Recordset = cloud_registro
    Exit Function
    
conecta:
    If ConectaMySQL Then
        GoTo retenta
    Else
        Unload MDIForm_Xsol
    End If
    
End Function

Public Function isAdmin() As Boolean
    isAdmin = usuario.Profile = ADMIN_PROFILE
End Function

Public Sub setGridSelRow(grid As MSFlexGrid, cod As String)
    Dim row As Integer
        
    row = 1
    With grid
        While row <= .Rows
            If .TextMatrix(row, 0) = cod Then
                .row = row
                .RowSel = row
                .col = 0
                .ColSel = .Cols - 1
                Exit Sub
            End If
            row = row + 1
        Wend
    End With
End Sub

Public Function Generate_Next_Visit(tipoServ As String) As Boolean
    
    'Verifica se o tipoServ (codigo do serviço) está como autoagendamento na tabela servico
    
    'Se o retorno for true então
        'Generate_Next_Visit = True
    'senao
        'Generate_Next_Visit = false
    'fimse
    
    'Vai no cadastro de tipo de serviços para saber se o serviço está setado para gerar agendamento automático após 1 ano
    
    Dim rsServico As ADODB.Recordset
    Set rsServico = Return_Recordset("Select autoagendamento From servico Where codigo = " & tipoServ)
    
    If rsServico.Fields("autoagendamento") = "S" Then
        Generate_Next_Visit = True
    Else
        Generate_Next_Visit = False
    End If
    
    
    
End Function




Public Function GET_BACKCOLOR_BY_STATUS(status As String) As Long
    If status = "BAIXADO" Then
        GET_BACKCOLOR_BY_STATUS = RGB(102, 255, 153)
    ElseIf status = "CANCELADO" Then
        GET_BACKCOLOR_BY_STATUS = RGB(255, 204, 204)
    Else
        GET_BACKCOLOR_BY_STATUS = 0
    End If
End Function

Public Sub setGridDefault(grid As MSFlexGrid)
    With grid
        If .Rows = 3 Then
            .row = 2
            .RowSel = 2
            .col = 0
            .ColSel = .Cols - 1
        End If
    End With
End Sub

Public Function cvInt(ByVal Value_Int As Variant) As Integer
    
    On Error Resume Next
    cvInt = CInt(Value_Int)
    If Err.Number <> 0 Then cvInt = 0
    
End Function

Public Function cvDbl(ByVal Value_Dbl As Variant) As Double
    
    On Error Resume Next
    cvDbl = Round(CDbl(Value_Dbl), 2)
    If Err.Number <> 0 Then cvDbl = 0
    
End Function

Public Function vData(ByVal Value_Data As Variant) As Date
   
   On Error Resume Next
   vData = CDate(Value_Data)
   If Err.Number <> 0 Then vData = Date
   
End Function

Public Function CPFValido(CPF As String) As Boolean
  
  Dim Soma As Integer
  Dim Resto As Integer
  Dim I As Integer
    
  If val(CPF) = 0 Then
   CPFValido = False
   Exit Function
  End If
  
  If Len(CPF) < 11 Then
   CPF = Format(val(CPF), "00000000000")
  End If

  Soma = 0
  For I = 1 To 9
   Soma = Soma + val(Mid$(CPF, I, 1)) * (11 - I)
  Next I
  Resto = 11 - (Soma - (Int(Soma / 11) * 11))
  If Resto = 10 Or Resto = 11 Then Resto = 0
  If Resto <> val(Mid(CPF, 10, 1)) Then
   CPFValido = False
   Exit Function
  End If
        
  Soma = 0
  For I = 1 To 10
   Soma = Soma + val(Mid$(CPF, I, 1)) * (12 - I)
  Next I
  Resto = 11 - (Soma - (Int(Soma / 11) * 11))
  If Resto = 10 Or Resto = 11 Then Resto = 0
  If Resto <> val(Mid(CPF, 11, 1)) Then
   CPFValido = False
   Exit Function
  End If
  CPFValido = True
  
End Function

Public Function CGCValido(CGC As String) As Boolean
 
 Dim a As Integer
 Dim J As Integer
 Dim I As Integer
 Dim D1 As Integer
 Dim D2 As Integer
 
  If Len(CGC) = 8 And val(CGC) > 0 Then
     a = 0
     J = 0
     D1 = 0
     For I = 1 To 7
      a = val(Mid(CGC, I, 1))
      If (I Mod 2) <> 0 Then a = a * 2
      If a > 9 Then
         J = J + Int(a / 10) + (a Mod 10)
      Else
         J = J + a
      End If
     Next I
     D1 = IIf((J Mod 10) <> 0, 10 - (J Mod 10), 0)
     If D1 = val(Mid(CGC, 8, 1)) Then
        CGCValido = True
     Else
        CGCValido = False
     End If
  Else
     If Len(CGC) = 14 And val(CGC) > 0 Then
        a = 0
        I = 0
        D1 = 0
        D2 = 0
        J = 5
        For I = 1 To 12 Step 1
            a = a + (val(Mid(CGC, I, 1)) * J)
            J = IIf(J > 2, J - 1, 9)
        Next I
        a = a Mod 11
        D1 = IIf(a > 1, 11 - a, 0)
        a = 0
        I = 0
        J = 6
        For I = 1 To 13 Step 1
            a = a + (val(Mid(CGC, I, 1)) * J)
            J = IIf(J > 2, J - 1, 9)
        Next I
        a = a Mod 11
        D2 = IIf(a > 1, 11 - a, 0)
        If (D1 = val(Mid(CGC, 13, 1)) And D2 = val(Mid(CGC, 14, 1))) Then
           CGCValido = True
        Else
           CGCValido = False
        End If
     Else
        CGCValido = False
     End If
  End If
  
End Function

Public Function Ajusta_Form(frm As Form)

    frm.Move 0, 0, MDIForm_Xsol.ScaleWidth, MDIForm_Xsol.ScaleHeight

End Function

Sub Arquivos_Dados()
    Dim ConfigFile As String
    Dim Path As String
    Dim FileName As String
    
    ConfigFile = App.Path & "\CONFIG.INI"
    
    Path = RTrim(GetINISetting("DATABASE", "PATH", ConfigFile))
    If Right(Path, 1) <> "\" Then
       Path = Path + "\"
    End If
    FileName = GetINISetting("DATABASE", "FILENAME", ConfigFile)
    
    Image_Folder = GetINISetting("IMAGE", "FOLDER", ConfigFile)
    W_DBname = Path + FileName
End Sub

Public Sub SetComboItem(c As ComboBox, s As String)
    Dim qtt As Integer
    Dim I As Integer
    
    qtt = c.ListCount
    
    For I = 0 To qtt - 1
        If StrComp(c.List(I), s, vbTextCompare) = 0 Then
            c.ListIndex = I
        End If
    Next
End Sub

Public Function MandatoryFieldFilled(objects() As Variant) As Boolean
    Dim v As Variant
    
    MandatoryFieldFilled = True
    For Each v In objects
       If Not IsEmpty(v) Then
          If TypeName(v) = "MaskEdBox" Then
             If v.Mask = "##/##/####" Then
                If Trim(v) = "__/__/____" Or Not IsDate(v) Then
                   MsgBox "É necessário digitar uma " & v.ToolTipText & " válida", vbOKOnly, "XSol Aquecedores"
                   v.SetFocus
                   MandatoryFieldFilled = False
                   Exit Function
                End If
             End If
          ElseIf Trim(v) = "" Then
             MsgBox "Campo " & v.ToolTipText & " deve ser preenchido", vbOKOnly, "XSol Aquecedores"
             v.SetFocus
             MandatoryFieldFilled = False
             Exit Function
          End If
       End If
    Next
End Function

Public Sub PaintMandatoryField(objects() As Variant, edit As Boolean)
    Dim v As Variant
    Dim color As ColorConstants
    
    color = IIf(edit, vbYellow, vbWhite)
    For Each v In objects
       If Not IsEmpty(v) Then
          v.BackColor = color
       End If
    Next
End Sub

Public Function GetFileExtension(file As String) As String
    Dim pos As Integer
    Dim fileLen As Integer
    Dim cont As Integer
    
    fileLen = Len(file)
    cont = 1
    While cont < fileLen
        pos = InStr(fileLen - cont, file, ".")
        If pos > 0 Then
            GetFileExtension = Mid(file, pos + 1)
            cont = fileLen
        Else
            cont = cont + 1
        End If
    Wend
    
End Function

Public Function versaoAtual() As String
    versaoAtual = "Versão de 20/02/2024.v2"
End Function

Public Function ConectaMySQL() As Boolean
Dim tentativas As Integer
On Error GoTo erro

cloud_conexao.CursorLocation = adUseClient

inicio:
    If cloud_conexao.State = 1 Then
        cloud_conexao.Close
    End If
    
    tentativas = tentativas + 1
    
    'FAZENDO A CONEXÃO COM MYSQL remoto
    cloud_conexao.Open "Driver={MySQL ODBC 3.51 Driver};" & _
    "Server= " & cloud_Servidor & ";" & _
    "Port=3306;" & _
    "Database=" & cloud_Database & ";" & _
    "User=" & cloud_usuario & ";" & _
    "Password=" & cloud_senha & ";" & _
    "Option=3;"
    
    cloud_status_conexao = 1
    MDIForm_Xsol.StatusBar.Panels(1).Text = "Conectado com " & cloud_Servidor
    cloud_img_online = cloud_FileName & "cloud_online_v2.jpg"
    MDIForm_Xsol.StatusBar.Panels(2).Picture = LoadPicture(cloud_img_online)
    ConectaMySQL = True
    
    Exit Function
erro:

    If cloud_status_conexao = 0 Then
        
        If cloud_conexao.State = 1 Then
        cloud_conexao.Close
        End If
        
        GoTo inicio
    End If
    
    cloud_status_conexao = 0
    MDIForm_Xsol.StatusBar.Panels(1).Text = "Banco de dados desconectado"
    cloud_img_offline = cloud_FileName & "cloud_offline_v2.jpg"
    MDIForm_Xsol.StatusBar.Panels(2).Picture = LoadPicture(cloud_img_offline)
    ConectaMySQL = False
End Function


Public Function textoComVirgulaEsquerda(texto As String) As String
    Dim textoFormatado
    
    textoFormatado = "," & texto
    
    textoComVirgulaEsquerda = textoFormatado
End Function


Public Sub cloud_Insert(tabela As String, camposSepPorVirgula As String, valuesSepPorVirgula As String)
    Dim contador As Integer
    Dim posicao As String
    
    Dim strCampos As String
    Dim strValues As String
    Dim strValuesFim As String
    
    strCampos = camposSepPorVirgula
    
    
    For contador = 1 To Len(valuesSepPorVirgula)
        posicao = Mid(valuesSepPorVirgula, contador, 1)
        If posicao <> "," Then
            strValues = strValues & posicao
        Else
            strValuesFim = strValuesFim & "'" & strValues & "',"
            strValues = ""
        End If
    Next contador
    strValuesFim = strValuesFim & "'" & strValues & "',"
    
    'Retira a vírbula do último
    strValuesFim = Mid(strValuesFim, 1, Len(strValuesFim) - 1)
    
    
    cloud_sql = "Insert into " & tabela & "(" & strCampos & ") Values(" & strValuesFim & ")"
    Set cloud_registro = cloud_conexao.Execute(cloud_sql)
    
End Sub


Public Sub cloud_Edit(tabela As String, camposSepPorVirgula As String, valuesSepPorVirgula As String, strWhere As String)
On Error GoTo erro
    Dim contador As Integer
    Dim contador2 As Integer
    Dim posicao As String
    Dim posicao2 As String
    
    Dim strCampos As String
    Dim strCampoDaVez As String
    Dim strcamposFim
    
    Dim intQtdCampos As Integer
    Dim strValues As String
    Dim strValuesFim As String
    Dim strvaloresUpdate As String
    
    Dim campoDaVez As String
    
    '------------------------- TRATAMENTO DOS CAMPOS ----------------------------------------------
    strCampos = camposSepPorVirgula & "."
    valuesSepPorVirgula = valuesSepPorVirgula & "."
    'StrCampos: Bairro,CEP,Cidade,Cod_cliente,Endereco,fonecel,fonecom,foneres,Ref,UF,Zona
    
    'conta quantos campos estarão no UPDATE

    For contador = 1 To Len(strCampos) + 1
proximoContador:
        posicao = Mid(strCampos, contador, 1)
        If posicao <> "," And posicao <> "." Then
            strcamposFim = strcamposFim & posicao
        Else
            
            'acabou a lista
            'If posicao = "." Then
            '    GoTo continua
            'End If
            'intCampos = intCampos + 1
            
            strCampoDaVez = strcamposFim
            '-----
                '------------------------- TRATAMENTO DOS VALUES-----------------------------------------------
                'Separa os values
                If contador2 > 0 Then
                    GoTo proximoContador2
                End If
                
                For contador2 = 1 To Len(valuesSepPorVirgula) + 1
proximoContador2:
                    posicao2 = Mid(valuesSepPorVirgula, contador2, 1)
                    If (posicao2 <> ",") And (posicao2 <> ".") Then
                        strValues = strValues & posicao2
                    Else
                        
                    
                        strvaloresUpdate = strvaloresUpdate & strCampoDaVez & "=" & Aspas(strValues) & ","
                        strValues = ""
                        strCampoDaVez = ""
                        strcamposFim = ""
                        contador = contador + 1
                        contador2 = contador2 + 1
                        
                        If posicao = "." Then
                            GoTo continua
                        End If
                        
                        GoTo proximoContador
                        'Next contador
                    End If
                Next contador2
                
                
                'String com os values:
                'strvaloresUpdate = Mid(strvaloresUpdate, 1, Len(strvaloresUpdate) - 1) 'Retira a vírbula do último value
                '-----------------------------------
            
            '-----

            'strcamposFim = ""
        End If
    Next contador
    
continua:

    strvaloresUpdate = Mid(strvaloresUpdate, 1, Len(strvaloresUpdate) - 1)
    cloud_sql = "UPDATE `" & tabela & "` set " & strvaloresUpdate & " Where " & strWhere
    
retenta:
    cloud_conexao.Execute (cloud_sql)
    Exit Sub
erro:
    If ConectaMySQL Then
        GoTo retenta
    Else
        MsgBox "Desconectado do cloud", vbCritical, "XSOL Aquecedores"
        End
    End If
    
End Sub


Public Function Aspas(texto As String) As String
    Dim textoFinal As String
    
    textoFinal = "'" & texto & "'"
    Aspas = textoFinal
End Function

Public Function cloud_delete(tabela As String, strWhere As String) As Boolean
    On Error GoTo erro
    
    cloud_sql = "DELETE from `" & tabela & "` Where " & strWhere
retenta:
    cloud_conexao.Execute (cloud_sql)
    cloud_delete = True
    Exit Function
    
erro:
    
    If cloud_status_conexao = 0 Then
        cloud_conexao.Close
    End If
    
    If ConectaMySQL Then
        GoTo retenta
    Else
        MsgBox "Erro ao tentar conexão com o cloud, o sistema será finalizado!", vbCritical, "XSOL Aquecedores"
        End
    End If
    
    
End Function

Public Sub reconectaMySQL()
    cloud_conexao.Close
    ConectaMySQL
End Sub

Public Function cloud_TemReferencia(strQuery As String) As Boolean
On Error GoTo erro

retenta:
    Set cloud_registro = cloud_conexao.Execute(strQuery)

    If cloud_registro.EOF Then
        cloud_TemReferencia = False
    Else
        cloud_TemReferencia = True
    End If
    Exit Function
    
erro:
    If cloud_status_conexao = 0 Then
        cloud_conexao.Close
        If ConectaMySQL Then
            GoTo retenta
        Else
            MsgBox "Erro ao estabelecer conexão com o cloud", vbCritical, "XSOL Aquecedores"
            End
        End If
    End If
    
End Function

