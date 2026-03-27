Attribute VB_Name = "Module_Db"
Public W_DB As Connection
Public Tabela_Clientes_TB As Recordset
Public W_DBname As String

Sub Main()

If Right(App.Path, 1) = "\" Then
  W_DBname = App.Path & "dados.mdb"
 Else
  W_DBname = App.Path & "\" & "dados.mdb"
End If

Set W_DB = New Connection

Set Tabela_Clientes_TB = New Recordset

W_DB.Mode = adModeReadWrite

W_DB.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"

Tabela_Clientes_TB.Open "select * from Clientes Order By Nome_Cli ", W_DB, adOpenStatic, adLockOptimistic

Form_Xsol_01.Show




End Sub


