Attribute VB_Name = "GBLOBAL"
Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
As String, ByVal lpKeyName As String, ByVal lpDefault As _
String, ByVal lpReturnedString As String, ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Declare Function DrawMenuBar Lib "User32" (ByVal hWnd As _
  Long) As Long
Declare Function GetMenuItemCount Lib "User32" (ByVal hMenu _
  As Long) As Long
Declare Function GetSystemMenu Lib "User32" (ByVal hWnd As _
  Long, ByVal bRevert As Long) As Long
Declare Function RemoveMenu Lib "User32" (ByVal hMenu As Long, _
  ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Const MF_BYPOSITION = &H400
Public Const MF_REMOVE = &H1000
Public Nome_Cli_Atend As String
Public W_DBname As String
Public Cod_Cli As Integer

Public Function cvInt(ByVal Value_Int As Variant) As Integer
    
    On Error Resume Next
    cvInt = CInt(Value_Int)
    If Err.Number <> 0 Then cvInt = 0
    
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
    
  If Val(CPF) = 0 Then
   CPFValido = False
   Exit Function
  End If
  
  If Len(CPF) < 11 Then
   CPF = Format(Val(CPF), "00000000000")
  End If

  Soma = 0
  For I = 1 To 9
   Soma = Soma + Val(Mid$(CPF, I, 1)) * (11 - I)
  Next I
  Resto = 11 - (Soma - (Int(Soma / 11) * 11))
  If Resto = 10 Or Resto = 11 Then Resto = 0
  If Resto <> Val(Mid(CPF, 10, 1)) Then
   CPFValido = False
   Exit Function
  End If
        
  Soma = 0
  For I = 1 To 10
   Soma = Soma + Val(Mid$(CPF, I, 1)) * (12 - I)
  Next I
  Resto = 11 - (Soma - (Int(Soma / 11) * 11))
  If Resto = 10 Or Resto = 11 Then Resto = 0
  If Resto <> Val(Mid(CPF, 11, 1)) Then
   CPFValido = False
   Exit Function
  End If
  CPFValido = True
  
End Function

Public Function CGCValido(CGC As String) As Boolean
 
 Dim A As Integer
 Dim J As Integer
 Dim I As Integer
 Dim D1 As Integer
 Dim D2 As Integer
 
  If Len(CGC) = 8 And Val(CGC) > 0 Then
     A = 0
     J = 0
     D1 = 0
     For I = 1 To 7
      A = Val(Mid(CGC, I, 1))
      If (I Mod 2) <> 0 Then A = A * 2
      If A > 9 Then
         J = J + Int(A / 10) + (A Mod 10)
      Else
         J = J + A
      End If
     Next I
     D1 = IIf((J Mod 10) <> 0, 10 - (J Mod 10), 0)
     If D1 = Val(Mid(CGC, 8, 1)) Then
        CGCValido = True
     Else
        CGCValido = False
     End If
  Else
     If Len(CGC) = 14 And Val(CGC) > 0 Then
        A = 0
        I = 0
        D1 = 0
        D2 = 0
        J = 5
        For I = 1 To 12 Step 1
            A = A + (Val(Mid(CGC, I, 1)) * J)
            J = IIf(J > 2, J - 1, 9)
        Next I
        A = A Mod 11
        D1 = IIf(A > 1, 11 - A, 0)
        A = 0
        I = 0
        J = 6
        For I = 1 To 13 Step 1
            A = A + (Val(Mid(CGC, I, 1)) * J)
            J = IIf(J > 2, J - 1, 9)
        Next I
        A = A Mod 11
        D2 = IIf(A > 1, 11 - A, 0)
        If (D1 = Val(Mid(CGC, 13, 1)) And D2 = Val(Mid(CGC, 14, 1))) Then
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

'If Right(App.Path, 1) = "\" Then
 ' W_DBname = App.Path & "C:\dados\xsol_novo\dados.mdb"
' Else
'  W_DBname = App.Path & "C:\dados\xsol_novo\dados.mdb"
'End If

W_DBname = "\\Servidor\Micro_Servidor\dados\xsol_novo\dados.mdb"

End Sub
