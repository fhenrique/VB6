VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEnderecoExistente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Endereços existentes"
   ClientHeight    =   2640
   ClientLeft      =   8265
   ClientTop       =   3045
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSair 
      BackColor       =   &H80000000&
      Caption         =   "&Sair"
      Height          =   495
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2055
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid_Address 
      Height          =   1890
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   3334
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmEnderecoExistente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSair_Click()
    'MsgBox "au"
    'gridEnderecos.TextMatrix(gridEnderecos.RowSel, 0)
    frmOrcamento.lblEndereco.Caption = Grid_Address.TextMatrix(Grid_Address.RowSel, 0)
    frmOrcamento.lblBairro.Caption = Grid_Address.TextMatrix(Grid_Address.RowSel, 1)
    frmOrcamento.lblCidade.Caption = Grid_Address.TextMatrix(Grid_Address.RowSel, 2)
    frmOrcamento.vStrEnderecoContato = Grid_Address.TextMatrix(Grid_Address.RowSel, 6)
    frmOrcamento.vStrEnderecoUF = Grid_Address.TextMatrix(Grid_Address.RowSel, 7)
    frmOrcamento.vStrEnderecoCep = Grid_Address.TextMatrix(Grid_Address.RowSel, 8)
    frmOrcamento.vStrEnderecoRef = Grid_Address.TextMatrix(Grid_Address.RowSel, 9)
    frmOrcamento.vStrEnderecoZona = Grid_Address.TextMatrix(Grid_Address.RowSel, 10)
    frmOrcamento.vStrEnderecoCodigo = Grid_Address.TextMatrix(Grid_Address.RowSel, 11)
    
    frmOrcamento.btnMudarEndereco.BackColor = &HE0E0E0
    frmOrcamento.timerAtualizarEndereco.enabled = False
    Unload Me
    'foneres 3,fonecom 4,fonecel 5, contato 6, uf 7, cep 8, referencia 9, zona 10, codigo 11
    
End Sub

Private Sub Form_Activate()
    Grid_Address.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
End Sub

Private Sub Form_Load()
    
    If frmOrcamento.Visible Then
        Call ConfigureGridEnderecos
        Call carregaGrid(Return_Recordset(Return_SqlString(1, frmOrcamento.lblCodigo.Caption)))
    Else
        frmEnderecoExistente.Height = 6645
    End If
    
End Sub

Private Sub carregaGrid(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    selectedRow = 2
    With Grid_Address
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("Endereco") & vbTab & IIf(IsNull(rs.Fields("Bairro")), "", rs.Fields("Bairro")) & vbTab & IIf(IsNull(rs.Fields("Cidade")), "", rs.Fields("Cidade")) & vbTab & IIf(IsNull(rs.Fields("foneres")), "", rs.Fields("foneres")) & vbTab & IIf(IsNull(rs.Fields("fonecom")), "", rs.Fields("fonecom")) & vbTab & IIf(IsNull(rs.Fields("fonecel")), "", rs.Fields("fonecel")) & vbTab & IIf(IsNull(rs.Fields("contato")), "", rs.Fields("contato")) & vbTab & IIf(IsNull(rs.Fields("uf")), "", rs.Fields("uf")) & vbTab & IIf(IsNull(rs.Fields("cep")), "", rs.Fields("cep")) & vbTab & IIf(IsNull(rs.Fields("ref")), "", rs.Fields("ref")) & vbTab & IIf(IsNull(rs.Fields("zona")), "", rs.Fields("zona")) & vbTab & IIf(IsNull(rs.Fields("codigo")), "", rs.Fields("codigo")))
            If code = rs.Fields("Endereco") Then
                selectedRow = .Rows - 1
            End If
            
            rs.MoveNext
        Loop
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            If .Rows > 9 Then
                .TopRow = selectedRow
            End If
            .col = 0
            .ColSel = .Cols - 1
        Else
        End If
        
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub ConfigureGridEnderecos()
    With Grid_Address
        .Cols = 12
        .ColWidth(0) = 4500  'Endreço
        .ColWidth(1) = 3300 'Bairro
        .ColWidth(2) = 3280  'Cidade
        .ColWidth(3) = 2000  'Foneres
        .ColWidth(4) = 2000  'Fonecom
        .ColWidth(5) = 2000  'Fonecel
        .ColWidth(6) = 2000  'contato
        .ColWidth(7) = 500  'uf
        .ColWidth(8) = 1000  'cep
        .ColWidth(9) = 2000  'ref
        .ColWidth(10) = 2000  'zona
        .ColWidth(10) = 1000  'codigo do endereço
        
        
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter

        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Endereço"
        .TextMatrix(0, 1) = "Bairro"
        .TextMatrix(0, 2) = "Cidade"
        .TextMatrix(0, 3) = "Fone Res."
        .TextMatrix(0, 4) = "Fone Com."
        .TextMatrix(0, 5) = "Fone Cel."
        .TextMatrix(0, 6) = "Contato"
        .TextMatrix(0, 7) = "UF"
        .TextMatrix(0, 8) = "CEP"
        .TextMatrix(0, 9) = "Referencia"
        .TextMatrix(0, 10) = "Zona"
        .TextMatrix(0, 11) = "Código"
        
        
        .RowHeight(1) = 0
    End With
End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String, Optional param3 As String) As String
    If queryType = 1 Then
        Return_SqlString = "Select endereco, bairro, cidade, uf, cep, contato, foneres, fonecom, fonecel, ref, zona, codigo From Endereco Where Cod_Cliente = " & param1
    End If
End Function

Private Function Return_Recordset(strSql As String) As ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim qtd As Integer
    
    Call Arquivos_Dados
    
    Set_Connection
    
    Set conn = Nothing
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & MDIForm_Xsol.W_DBname & ";"

    rs2.Open strSql, conn, adOpenStatic, adLockOptimistic
    qtd = rs2.RecordCount
    
    Set Return_Recordset = rs2
End Function
