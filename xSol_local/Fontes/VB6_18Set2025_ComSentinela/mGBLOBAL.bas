Attribute VB_Name = "GBLOBAL"
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

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

Public global_Ficha_telefone1 As String
Public global_Ficha_telefone2 As String
Public global_Ficha_telefone3 As String
Public global_abrindoForm As Boolean

Public Type User
    Login As String
    Profile As String
End Type

Public usuario As User

Public Sub Set_Connection()
   Set conn = New ADODB.Connection
   conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
End Sub

Public Sub ConnectionRefresh()
    conn.Close
    Set conn = Nothing
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
End Sub

Public Function Get_Field_Val(ByVal val As Object) As String
    Get_Field_Val = UCase(IIf(IsNull(val), "", val))
End Function

Public Function Set_Field_Val(ByVal val As Object) As String
    Set_Field_Val = IIf(val = "", vbNullString, UCase(val))
End Function

Public Function Return_Recordset(strSql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    rs.Open strSql, conn, adOpenStatic, adLockOptimistic
    
    Set Return_Recordset = rs
End Function

Public Function isAdmin() As Boolean
    isAdmin = usuario.Profile = ADMIN_PROFILE
End Function

Public Sub setGridSelRow(grid As MSFlexGrid, cod As String)
    Dim row As Integer
        
    row = 1
    With grid
        If .Rows = 2 Then
            Exit Sub
        End If
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
    versaoAtual = "Versão de 30/09/2025.v5"
End Function

Public Function validaCaractere(intKeyAscii As String) As Boolean
    If IsNumeric(intKeyAscii) Then
        intKeyAscii = Int(intKeyAscii)
    End If
    
    Select Case intKeyAscii
        Case 8 'back
            validaCaractere = True
        Case 44 'vírgula
            validaCaractere = True
        Case 46 'ponto
            validaCaractere = True
        Case 48 To 57 'Números de 0 a 9
            validaCaractere = True
        Case Else
            validaCaractere = False 'Não deixa nenhuma outra caractere ser escrito
    End Select
End Function

Public Function validaCaractere2(intKeyAscii As String) As Boolean
    If IsNumeric(intKeyAscii) Then
        intKeyAscii = Int(intKeyAscii)
    End If
    
    Select Case intKeyAscii
        Case 8 'back
            validaCaractere2 = True
        Case 47 '/
            validaCaractere2 = True
        Case 48 To 57 'Números de 0 a 9
            validaCaractere2 = True
        Case Else
            validaCaractere2 = False 'Não deixa nenhuma outra caractere ser escrito
    End Select
End Function

