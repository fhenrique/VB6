VERSION 5.00
Begin VB.Form frmFiltroRelatorio_Cliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecione a cidade"
   ClientHeight    =   1665
   ClientLeft      =   2970
   ClientTop       =   2655
   ClientWidth     =   4110
   Icon            =   "frmFiltroRelatorio_Cliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   30
      TabIndex        =   0
      Top             =   -45
      Width           =   4065
      Begin VB.ComboBox cboCidades 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   435
         Width           =   3855
      End
      Begin VB.CommandButton btnOk 
         Caption         =   "&OK"
         Height          =   345
         Left            =   3330
         TabIndex        =   2
         Top             =   1215
         Width           =   615
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   2370
         TabIndex        =   1
         Top             =   1215
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmFiltroRelatorio_Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub EstoqueArquivos_Dados()
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
    Estoque_DBname = Path + FileName
End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String) As String
    Dim whereClause As String
    
    If queryType = 1 Then
        Return_SqlString = "Select Distinct(cidade) from Endereco Where cidade not in ('*','**') order by 1"
    End If
    
End Function


Public Function Return_Recordset(strSql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    Call EstoqueArquivos_Dados
    
    Set conn = Nothing
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
    
    rs.Open strSql, conn, adOpenStatic, adLockOptimistic
    
    Set Return_Recordset = rs
End Function

Private Sub carregaCidades(rs As ADODB.Recordset, Optional code As String)
    
    cboCidades.Clear
    
    If rs.RecordCount > 0 Then
        With cboCidades
            If rs.EOF Or rs.BOF Then
               Exit Sub
            End If
    
            Do Until rs.EOF
                .AddItem (rs.Fields("cidade"))
                rs.MoveNext
            Loop
        End With
    Else
        cboCidades.Clear
    End If
    
    rs.Close
    Set rs = Nothing
End Sub


Private Sub executaRelatorio(dbname As String, titulo1Rel As String)

    relClientes.DAODataControl1.DatabaseName = dbname

    relClientes.lblTipoRelatorio.Caption = titulo1Rel
    relClientes.DAODataControl1.RecordSource = "Select cli.Codigo, cli.Data_Emissao, cli.Nome_Cli, cli.CPF, cli.CNPJ, cli.Fone_Celular, cli.Fone_Res, cli.Fone_Com, cli.E_Mail, cli.Aquivo_Morto,end.Endereco, end.Bairro, end.Ref, end.CEP From Clientes cli Inner join Endereco end on end.Cod_Cliente = cli.Codigo Where ucase(end.Cidade) = '" & UCase(cboCidades.Text) & "' Order by 3"
    relClientes.Show (1)
    
End Sub

Private Sub executaRelatorioArquivoMorto(dbname As String, titulo1Rel As String)
    relClientes.DAODataControl1.DatabaseName = dbname

    relClientes.lblTipoRelatorio.Caption = titulo1Rel
    relClientes.DAODataControl1.RecordSource = "Select cli.Codigo, cli.Data_Emissao, cli.Nome_Cli, cli.CPF, cli.CNPJ, cli.Fone_Celular, cli.Fone_Res, cli.Fone_Com, cli.E_Mail, cli.Aquivo_Morto,end.Endereco, end.Bairro, end.Ref, end.CEP From Clientes cli Inner join Endereco end on end.Cod_Cliente = cli.Codigo Where ucase(end.Cidade) = '" & UCase(cboCidades.Text) & "' Order by 3"
    relClientes.Show (1)
End Sub


Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    executaRelatorio W_DBname, cboCidades.Text
End Sub

Private Sub Form_Activate()
If (MDIForm_Xsol.acessibilidade_TextoNegrito = True) Or (MDIForm_Xsol.acessibilidade_TextoNegrito = False) Then
        cboCidades.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    End If
End Sub

Private Sub Form_Load()
    Call carregaCidades(Return_Recordset(Return_SqlString(1)))
End Sub

