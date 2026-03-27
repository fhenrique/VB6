VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Begin VB.Form Form_Fornecedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Fornecedores"
   ClientHeight    =   7305
   ClientLeft      =   1230
   ClientTop       =   1830
   ClientWidth     =   19095
   Icon            =   "Form_Fornecedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   19095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame panEquipamentos 
      Caption         =   "Equipamentos"
      Height          =   2655
      Left            =   6195
      TabIndex        =   32
      Top             =   4575
      Width           =   12855
      Begin VB.CommandButton btnAtualizaValorCompraEquipamento 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   420
         Left            =   9150
         Picture         =   "Form_Fornecedores.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Atualizar o valor do Equipamento."
         Top             =   210
         Width           =   375
      End
      Begin VB.ComboBox txtEquipFornDescricao 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1065
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   255
         Width           =   5880
      End
      Begin VB.CommandButton btnEquipFornRemove 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   390
         Left            =   9675
         Picture         =   "Form_Fornecedores.frx":0C3D
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Remover o equipamento selecionado, deste fornecedor."
         Top             =   2115
         Width           =   390
      End
      Begin VB.CommandButton btnEquipFornAdd 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   9675
         Picture         =   "Form_Fornecedores.frx":133F
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Inserir Equipamento a este fornecedor"
         Top             =   1650
         Width           =   345
      End
      Begin VB.TextBox txtEquipFornValor 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7860
         TabIndex        =   34
         Top             =   300
         Width           =   1065
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_FornecedorEquipamento 
         Height          =   1860
         Left            =   195
         TabIndex        =   36
         Top             =   690
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor:"
         Height          =   255
         Left            =   7140
         TabIndex        =   35
         Top             =   345
         Width           =   615
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   195
         TabIndex        =   33
         Top             =   345
         Width           =   765
      End
   End
   Begin VB.Frame panFornecedores 
      Caption         =   "Listagem de Fornecedores"
      Height          =   4470
      Left            =   6195
      TabIndex        =   30
      Top             =   45
      Width           =   12825
      Begin MSFlexGridLib.MSFlexGrid Grid_Fornecedores 
         Height          =   4155
         Left            =   120
         TabIndex        =   31
         Top             =   225
         Width           =   12600
         _ExtentX        =   22225
         _ExtentY        =   7329
         _Version        =   393216
         Rows            =   1
         Cols            =   12
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Edição"
      Height          =   1095
      Left            =   90
      TabIndex        =   25
      Top             =   6120
      Width           =   6015
      Begin VB.CommandButton Cmd_Exclui 
         BackColor       =   &H008080FF&
         Caption         =   "E&xcluir"
         Height          =   495
         Left            =   4650
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Edita 
         Caption         =   "&Editar"
         Height          =   495
         Left            =   1650
         TabIndex        =   28
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Salva 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3180
         TabIndex        =   27
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Inclui 
         BackColor       =   &H0080FF80&
         Caption         =   "&Inserir"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   375
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fornecedor"
      Height          =   6030
      Left            =   105
      TabIndex        =   0
      Top             =   30
      Width           =   6015
      Begin VB.TextBox txtMaterial 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   1185
         Left            =   1215
         MaxLength       =   120
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   4245
         Width           =   4590
      End
      Begin VB.TextBox txtEmail 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1215
         MaxLength       =   120
         TabIndex        =   21
         Top             =   3825
         Width           =   4560
      End
      Begin VB.TextBox txtContato 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1215
         MaxLength       =   24
         TabIndex        =   19
         Top             =   3405
         Width           =   4560
      End
      Begin VB.TextBox txtFax 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1215
         MaxLength       =   49
         TabIndex        =   17
         Top             =   5580
         Visible         =   0   'False
         Width           =   4560
      End
      Begin VB.TextBox txtFone 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1215
         MaxLength       =   49
         TabIndex        =   15
         Top             =   3000
         Width           =   4545
      End
      Begin VB.TextBox txtCnpj 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3585
         MaxLength       =   20
         TabIndex        =   13
         Top             =   2580
         Width           =   2220
      End
      Begin VB.TextBox txtCep 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1215
         MaxLength       =   11
         TabIndex        =   11
         Top             =   2580
         Width           =   1140
      End
      Begin VB.TextBox txtCidade 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1215
         MaxLength       =   25
         TabIndex        =   9
         Top             =   2175
         Width           =   4545
      End
      Begin VB.TextBox txtBairro 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1215
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1725
         Width           =   4560
      End
      Begin VB.TextBox txtEndereco 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1215
         MaxLength       =   70
         TabIndex        =   5
         Top             =   1200
         Width           =   4560
      End
      Begin VB.TextBox txtNome 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1215
         MaxLength       =   69
         TabIndex        =   3
         Top             =   780
         Width           =   4560
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1215
         TabIndex        =   1
         Top             =   330
         Width           =   1230
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Material:"
         Height          =   255
         Left            =   135
         TabIndex        =   24
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Email:"
         Height          =   255
         Left            =   135
         TabIndex        =   22
         Top             =   3885
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Contato:"
         Height          =   255
         Left            =   135
         TabIndex        =   20
         Top             =   3465
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Fax:"
         Height          =   255
         Left            =   135
         TabIndex        =   18
         Top             =   5640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Fone:"
         Height          =   255
         Left            =   135
         TabIndex        =   16
         Top             =   3060
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "CNPJ:"
         Height          =   255
         Left            =   2475
         TabIndex        =   14
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "CEP:"
         Height          =   255
         Left            =   135
         TabIndex        =   12
         Top             =   2580
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cidade:"
         Height          =   255
         Left            =   135
         TabIndex        =   10
         Top             =   2235
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Bairro:"
         Height          =   255
         Left            =   135
         TabIndex        =   8
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "End.:"
         Height          =   255
         Left            =   135
         TabIndex        =   6
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Lbl_Cliente 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome:"
         Height          =   255
         Left            =   135
         TabIndex        =   4
         Top             =   825
         Width           =   975
      End
      Begin VB.Label Lbl_Registro 
         Alignment       =   1  'Right Justify
         Caption         =   "Código:"
         Height          =   255
         Left            =   495
         TabIndex        =   2
         Top             =   375
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form_Fornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Fornecedor_DBname As String


Private Sub Configure_GridFornEquip()
    With Grid_FornecedorEquipamento
        .Cols = 3
        .ColWidth(0) = 800 'Codigo
        .ColWidth(1) = 7080 'Descricao
        .ColWidth(2) = 1500 'Valor
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Descrição"
        .TextMatrix(0, 2) = "Valor"
        .RowHeight(1) = 0
    End With
End Sub

Private Sub Configure_Grid()
    With Grid_Fornecedores
        .Cols = 12
        .ColWidth(0) = 800
        .ColWidth(1) = 4080
        .ColWidth(2) = 3100 'endereço
        .ColWidth(3) = 2100 'bairro
        .ColWidth(4) = 2100 'cidade
        .ColWidth(5) = 1100 'cep
        .ColWidth(6) = 2100 'cnpj
        .ColWidth(7) = 2100 'fone
        .ColWidth(8) = 2100 'fax
        .ColWidth(9) = 2100 'contato
        .ColWidth(10) = 2100 'email
        .ColWidth(11) = 5000 'tipoMaterial
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignLeftCenter
        .ColAlignment(8) = flexAlignLeftCenter
        .ColAlignment(9) = flexAlignLeftCenter
        .ColAlignment(10) = flexAlignLeftCenter
        .ColAlignment(11) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Nome Fornecedor"
        .TextMatrix(0, 2) = "Endereço"
        .TextMatrix(0, 3) = "Bairro"
        .TextMatrix(0, 4) = "Cidade"
        .TextMatrix(0, 5) = "CEP"
        .TextMatrix(0, 6) = "CNPJ"
        .TextMatrix(0, 7) = "Fone"
        .TextMatrix(0, 8) = "Fax"
        .TextMatrix(0, 9) = "Contato"
        .TextMatrix(0, 10) = "Email"
        .TextMatrix(0, 11) = "Tipo Material"
        .RowHeight(1) = 0
    End With
End Sub

Private Sub Fill_GridFornEquuip(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    
    selectedRow = 2
    With Grid_FornecedorEquipamento
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("CodMaterial") & vbTab & rs.Fields("Descricao") & vbTab & rs.Fields("Valor"))
            If code = rs.Fields("Codigo") Then
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
            
            'Call Grid_Fornecedores_Click
        Else
            'Call Object_Controller(1)
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub


Private Sub Fill_Grid(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    
    selectedRow = 2
    With Grid_Fornecedores
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("codigo") & vbTab & IIf(IsNull(rs.Fields("NomeCli")), "", rs.Fields("NomeCli")) & vbTab & IIf(IsNull(rs.Fields("End")), "", rs.Fields("End")) & vbTab & IIf(IsNull(rs.Fields("Bairro")), "", rs.Fields("Bairro")) & vbTab & IIf(IsNull(rs.Fields("Cidade")), "", rs.Fields("Cidade")) & vbTab & IIf(IsNull(rs.Fields("Cep")), "", rs.Fields("Cep")) & vbTab & IIf(IsNull(rs.Fields("CNPJ")), "", rs.Fields("CNPJ")) & vbTab & IIf(IsNull(rs.Fields("Fone")), "", rs.Fields("Fone")) & vbTab & IIf(IsNull(rs.Fields("Fax")), "", rs.Fields("Fax")) & vbTab & IIf(IsNull(rs.Fields("Contato")), "", rs.Fields("Contato")) & vbTab & IIf(IsNull(rs.Fields("Email")), "", rs.Fields("Email")) & vbTab & IIf(IsNull(rs.Fields("Tipo_Mat")), "", rs.Fields("Tipo_Mat")))
            If code = rs.Fields("codigo") Then
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
            
            'Call Object_Controller(2)
            Call Grid_Fornecedores_Click
        Else
            'Call Object_Controller(1)
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Sub FornecedorArquivos_Dados()
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
    Fornecedor_DBname = Path + FileName
End Sub


Public Function Return_Recordset(strSql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    Call FornecedorArquivos_Dados
    
    Set conn = Nothing
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & Fornecedor_DBname & ";"
    
    'MsgBox strSql
    rs.Open strSql, conn, adOpenStatic, adLockOptimistic
    
    Set Return_Recordset = rs
End Function

Private Function carregaCodMaterial(rs As ADODB.Recordset, Optional code As String) As String
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    If rs.EOF Or rs.BOF Then
       carregaCodMaterial = 0
    Else
       carregaCodMaterial = rs.Fields("Codigo")
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String) As String
    Dim whereClause As String
    
    If queryType = 1 Then
        Return_SqlString = "SELECT * FROM Fornecedor"
    ElseIf queryType = 2 Then
        Return_SqlString = "SELECT * FROM Fornecedor WHERE Codigo = " & param1
    ElseIf queryType = 3 Then
        If param1 <> "" Then
            Return_SqlString = "SELECT tb1.Codigo, tb1.codMaterial as codMaterial, tb2.Descricao as Descricao, tb1.ValorCompra as Valor FROM MaterialFornecedor as tb1 Inner Join Material as tb2 on tb2.Codigo = tb1.codMaterial WHERE CodFornecedor = " & param1
        Else
            Return_SqlString = "SELECT * FROM FornecedorEquipamento WHERE CodigoFornecedor = 0"
        End If
    ElseIf queryType = 4 Then
        Return_SqlString = "SELECT * FROM FornecedorEquipamento"
    ElseIf queryType = 5 Then
        Return_SqlString = "SELECT * FROM FornecedorEquipamento WHERE Codigo = " & param1
    ElseIf queryType = 6 Then
        Return_SqlString = "SELECT * FROM Material Order by Descricao"
    ElseIf queryType = 7 Then
        Return_SqlString = "Select ValorCompra From Material tb1 Inner Join MaterialFornecedor tb2 on tb2.codMaterial = tb1.Codigo where tb1.Descricao = '" & param1 & "'"
    ElseIf queryType = 8 Then
        Return_SqlString = "Select * From MaterialFornecedor Where CodFornecedor = " & param1 & " And codMaterial = " & param2
    ElseIf queryType = 9 Then
        Return_SqlString = "SELECT * FROM Material WHERE Descricao= """ & param1 & """"
    ElseIf queryType = 10 Then
        Return_SqlString = "Select tb1.Descricao,  tb2.ValorCompra From Material tb1 Inner Join MaterialFornecedor tb2 on tb2.codMaterial = tb1.Codigo where tb1.Descricao = '" & param1 & "'"
    End If
End Function

Private Sub btnAtualizaValorCompraEquipamento_Click()
    
    Dim rs As ADODB.Recordset
    Dim codMaterial As String
   
    If txtEquipFornDescricao.Text = "" Then
        MsgBox "Selecione um equipamento e informe o novo valor antes de clicar neste botão."
        txtEquipFornDescricao.SetFocus
        Exit Sub
    End If
    
    If Trim(txtEquipFornValor.Text) = "" Then
        MsgBox "Informe o novo valor do equipamento " & txtEquipFornDescricao.Text & " antes de clicar neste botão"
        txtEquipFornDescricao.SetFocus
        Exit Sub
    End If
        
    codMaterial = carregaCodMaterial(Return_Recordset(Return_SqlString(9, txtEquipFornDescricao.Text)))
        
    Set rs = Return_Recordset(Return_SqlString(8, txtCodigo.Text, codMaterial))

    If rs.RecordCount = 0 Then
        MsgBox "O produto " & txtEquipFornDescricao.Text & " não está na lista de equipametos fornecidos pelo fornecedor " & txtNome.Text
        rs.Clone
        Set rs = Nothing
        Exit Sub
    Else
        Call Record_Equipamento(rs)
        rs.Update
    
        rs.Close
        Set rs = Nothing
        Call Fill_GridFornEquuip(Return_Recordset(Return_SqlString(3, txtCodigo.Text)))
    End If
        
End Sub

Private Sub btnEquipFornAdd_Click()
    Dim rs As ADODB.Recordset
    
    If Trim(txtEquipFornDescricao.Text) = "" Then
        MsgBox "Informe a descrição do Equipamento a ser inserido.", vbOKOnly, "xSol Aquecedores"
        Exit Sub
    End If
    
    If Trim(txtEquipFornValor.Text) = "" Then
        MsgBox "Informe o valor do Equipamento a ser inserido.", vbOKOnly, "xSol Aquecedores"
        Exit Sub
    End If

    If Not IsNumeric(txtEquipFornValor.Text) Then
        MsgBox "Informe um valor válido.", vbOKOnly, "xSol Aquecedores"
        Exit Sub
    End If
    
    Set rs = Return_Recordset(Return_SqlString(4))

    rs.AddNew
    
    rs.Fields("Descricao") = txtEquipFornDescricao.Text
    rs.Fields("Valor") = txtEquipFornValor.Text
    rs.Fields("CodigoFornecedor") = txtCodigo.Text
    
    rs.Update
   
    rs.Close
    Set rs = Nothing
    
    txtEquipFornDescricao.ListIndex = -1
    txtEquipFornValor.Text = ""
    txtEquipFornDescricao.SetFocus
    
    Call Fill_GridFornEquuip(Return_Recordset(Return_SqlString(3, txtCodigo.Text)))
End Sub

Private Sub btnEquipFornRemove_Click()
    If Grid_FornecedorEquipamento.Rows > 2 Then
        Call Return_Recordset(Return_SqlString(5, Grid_FornecedorEquipamento.TextMatrix(Grid_FornecedorEquipamento.RowSel, 0))).Delete
        Call Fill_GridFornEquuip(Return_Recordset(Return_SqlString(3, txtCodigo.Text)))
    End If
End Sub

Private Sub Cmd_Edita_Click()
    If Grid_Fornecedores.Rows = 2 Then
        Exit Sub
    End If

    If Cmd_Edita.Caption = "&Editar" Then 'HABILITA
        Cmd_Edita.Caption = "&Cancelar"
        
        Cmd_Inclui.enabled = False
        Cmd_Salva.enabled = True
        Cmd_Exclui.enabled = False
        
        btnAtualizaValorCompraEquipamento.enabled = True
        btnEquipFornAdd.enabled = True
        btnEquipFornRemove.enabled = True
        
        panFornecedores.enabled = False
        
        txtNome.BackColor = &HFFFFFF
        txtNome.enabled = True
        
        txtEndereco.BackColor = &HFFFFFF
        txtEndereco.enabled = True
        
        txtBairro.BackColor = &HFFFFFF
        txtBairro.enabled = True
        
        txtCidade.BackColor = &HFFFFFF
        txtCidade.enabled = True
        
        txtCep.BackColor = &HFFFFFF
        txtCep.enabled = True
        
        txtCnpj.BackColor = &HFFFFFF
        txtCnpj.enabled = True
        
        txtFone.BackColor = &HFFFFFF
        txtFone.enabled = True
        
        'txtFax.BackColor = &HFFFFFF
        'txtFax.enabled = True
        
        txtContato.BackColor = &HFFFFFF
        txtContato.enabled = True
        
        txtEmail.BackColor = &HFFFFFF
        txtEmail.enabled = True
        
        txtMaterial.BackColor = &HFFFFFF
        txtMaterial.enabled = True
        
        txtEquipFornDescricao.BackColor = &HFFFFFF
        txtEquipFornDescricao.enabled = True
        
        panEquipamentos.enabled = True
        
        txtEquipFornValor.BackColor = &HFFFFFF
        txtEquipFornValor.enabled = True
        
    Else    'DESABILITA
        Cmd_Edita.Caption = "&Editar"
        
        Cmd_Inclui.enabled = True
        Cmd_Salva.enabled = False
        Cmd_Exclui.enabled = True
        
        btnAtualizaValorCompraEquipamento.enabled = False
        btnEquipFornAdd.enabled = False
        btnEquipFornRemove.enabled = False
        
        panFornecedores.enabled = True
        
        txtNome.BackColor = &HC0C0C0
        txtNome.enabled = False
        
        txtEndereco.BackColor = &HC0C0C0
        txtEndereco.enabled = False
        
        txtBairro.BackColor = &HC0C0C0
        txtBairro.enabled = False
        
        txtCidade.BackColor = &HC0C0C0
        txtCidade.enabled = False
        
        txtCep.BackColor = &HC0C0C0
        txtCep.enabled = False
        
        txtCnpj.BackColor = &HC0C0C0
        txtCnpj.enabled = False
        
        txtFone.BackColor = &HC0C0C0
        txtFone.enabled = False
        
        'txtFax.BackColor = &HC0C0C0
        'txtFax.enabled = False
        
        txtContato.BackColor = &HC0C0C0
        txtContato.enabled = False
        
        txtEmail.BackColor = &HC0C0C0
        txtEmail.enabled = False
        
        txtMaterial.BackColor = &HC0C0C0
        txtMaterial.enabled = False
        
        'panEquipamentos.enabled = False
        
        txtEquipFornDescricao.BackColor = &HC0C0C0
        txtEquipFornDescricao.enabled = False
        
        txtEquipFornValor.BackColor = &HC0C0C0
        txtEquipFornValor.enabled = False
    End If
End Sub

Private Sub Cmd_Exclui_Click()
   If Grid_Fornecedores.Rows > 2 Then
    If MsgBox("Confirma a Exclusão do fornecedor " & txtNome.Text & " ? ", vbYesNo, "Excluir Registros") = vbYes Then
       Call Return_Recordset(Return_SqlString(2, Grid_Fornecedores.TextMatrix(Grid_Fornecedores.RowSel, 0))).Delete
       Call Fill_Grid(Return_Recordset(Return_SqlString(1)))
    End If
   Else
    
    txtNome.Text = ""
    txtEndereco.Text = ""
    txtBairro.Text = ""
    txtCidade.Text = ""
    txtCep.Text = ""
    txtCnpj.Text = ""
    txtFone.Text = ""
    'txtFax.Text = ""
    txtContato.Text = ""
    txtEmail.Text = ""
    txtMaterial.Text = ""
    
   End If
End Sub

Private Sub Cmd_Inclui_Click()

    If Cmd_Inclui.Caption = "&Inserir" Then 'HABILITA, está Inserir
    
        Dim qtdLinhas As Integer
        Dim ultimoCodigo As Integer
        
        Cmd_Inclui.Caption = "&Cancelar"
        
        qtdLinhas = Grid_Fornecedores.Rows
        qtdLinhas = qtdLinhas - 1
        
        If qtdLinhas >= 2 Then
            ultimoCodigo = Grid_Fornecedores.TextMatrix(qtdLinhas, 0)
        Else
            ultimoCodigo = 0
        End If
        '
    
        txtCodigo.Text = Str(ultimoCodigo + 1)
        
        txtNome.Text = ""
        txtNome.BackColor = &HFFFFFF
        txtNome.enabled = True
        
        txtEndereco.Text = ""
        txtEndereco.BackColor = &HFFFFFF
        txtEndereco.enabled = True
        
        txtBairro.Text = ""
        txtBairro.BackColor = &HFFFFFF
        txtBairro.enabled = True
        
        txtCidade.Text = ""
        txtCidade.BackColor = &HFFFFFF
        txtCidade.enabled = True
        
        txtCep.Text = ""
        txtCep.BackColor = &HFFFFFF
        txtCep.enabled = True
        
        txtCnpj.Text = ""
        txtCnpj.BackColor = &HFFFFFF
        txtCnpj.enabled = True
        
        txtFone.Text = ""
        txtFone.BackColor = &HFFFFFF
        txtFone.enabled = True
        
        'txtFax.Text = ""
        'txtFax.BackColor = &HFFFFFF
        'txtFax.enabled = True
        
        txtContato.Text = ""
        txtContato.BackColor = &HFFFFFF
        txtContato.enabled = True
        
        txtEmail.Text = ""
        txtEmail.BackColor = &HFFFFFF
        txtEmail.enabled = True
        
        txtMaterial.Text = ""
        txtMaterial.BackColor = &HFFFFFF
        txtMaterial.enabled = True
    
        txtNome.SetFocus
    
        Cmd_Inclui.Caption = "&Cancelar"
        
        Cmd_Inclui.enabled = True
        Cmd_Edita.enabled = False
        Cmd_Salva.enabled = True
        Cmd_Exclui.enabled = False
        
        Grid_Fornecedores.Visible = False
        Grid_FornecedorEquipamento.Visible = False
        
    Else    'DESABILITA, está CANCELAR
        Cmd_Inclui.Caption = "&Inserir"
        
        Cmd_Inclui.enabled = True
        Cmd_Salva.enabled = False
        Cmd_Edita.enabled = True
        Cmd_Exclui.enabled = True
        
        'txtProdutoDescricao.enabled = False
        'txtProdutoDescricao.BackColor = &HC0C0C0
        
        txtNome.BackColor = &HC0C0C0
        txtNome.enabled = False
        
        txtEndereco.BackColor = &HC0C0C0
        txtEndereco.enabled = False
        
        txtBairro.BackColor = &HC0C0C0
        txtBairro.enabled = False
        
        txtCidade.BackColor = &HC0C0C0
        txtCidade.enabled = False
        
        txtCep.BackColor = &HC0C0C0
        txtCep.enabled = False
        
        txtCnpj.BackColor = &HC0C0C0
        txtCnpj.enabled = False
        
        txtFone.BackColor = &HC0C0C0
        txtFone.enabled = False
        
        'txtFax.BackColor = &HC0C0C0
        'txtFax.enabled = False
        
        txtContato.BackColor = &HC0C0C0
        txtContato.enabled = False
        
        txtEmail.BackColor = &HC0C0C0
        txtEmail.enabled = False
        
        txtMaterial.BackColor = &HC0C0C0
        txtMaterial.enabled = False
        
        Grid_Fornecedores.Visible = True
        Grid_FornecedorEquipamento.Visible = True
    End If


End Sub

Private Sub Record_Record(rs As ADODB.Recordset)
   rs.Fields("NomeCli").value = txtNome.Text
   rs.Fields("End").value = txtEndereco.Text
   rs.Fields("Bairro").value = txtBairro.Text
   rs.Fields("Cidade").value = txtCidade.Text
   rs.Fields("Cep").value = txtCep.Text
   rs.Fields("CNPJ").value = txtCnpj.Text
   rs.Fields("Fone").value = txtFone.Text
   'rs.Fields("Fax").value = txtFax.Text
   rs.Fields("Contato").value = txtContato.Text
   rs.Fields("Email").value = txtEmail.Text
   rs.Fields("Tipo_Mat").value = txtMaterial.Text
End Sub

Private Sub Record_Equipamento(rs As ADODB.Recordset)
   rs.Fields("ValorCompra").value = txtEquipFornValor.Text
End Sub


Private Sub Cmd_Salva_Click()

   If Trim(txtNome.Text) = "" Then
        MsgBox "O campo Nome é de preenchimento obrigatório."
        txtNome.SetFocus
        Exit Sub
   End If
   
   If Trim(txtEndereco.Text) = "" Then
        MsgBox "O campo Endereço é de preenchimento obrigatório."
        txtEndereco.SetFocus
        Exit Sub
   End If

   If Trim(txtBairro.Text) = "" Then
        MsgBox "O campo Bairro é de preenchimento obrigatório."
        txtBairro.SetFocus
        Exit Sub
   End If

   If Trim(txtCep.Text) = "" Then
        MsgBox "O campo CEP é de preenchimento obrigatório."
        txtCep.SetFocus
        Exit Sub
   End If
   
   If Trim(txtCnpj.Text) = "" Then
        MsgBox "O campo CNPJ é de preenchimento obrigatório."
        txtCnpj.SetFocus
        Exit Sub
   End If
   
   If Trim(txtFone.Text) = "" Then
        MsgBox "O campo FONE é de preenchimento obrigatório."
        txtFone.SetFocus
        Exit Sub
   End If
   
   If Trim(txtContato.Text) = "" Then
        MsgBox "O campo CONTATO é de preenchimento obrigatório."
        txtContato.SetFocus
        Exit Sub
   End If
   
   If Trim(txtEmail.Text) = "" Then
        MsgBox "O campo CNPJ é de preenchimento obrigatório."
        txtEmail.SetFocus
        Exit Sub
   End If
   
   If Trim(txtMaterial.Text) = "" Then
        MsgBox "O campo TIPO MATERIAL é de preenchimento obrigatório."
        txtMaterial.SetFocus
        Exit Sub
   End If

   Dim rs As ADODB.Recordset
   
      If Trim(txtNome.Text) = "" Then
        MsgBox "Preencha o nome do Fornecedor.", vbOKOnly, "Xsol Aquecedores"
        txtNome.SetFocus
        Exit Sub
      End If
      
      If Trim(txtEndereco.Text) = "" Then
        MsgBox "Preencha o endereço.", vbOKOnly, "Xsol Aquecedores"
        txtEndereco.SetFocus
        Exit Sub
      End If
      
      If Trim(txtBairro.Text) = "" Then
        MsgBox "Preencha o bairro.", vbOKOnly, "Xsol Aquecedores"
        txtBairro.SetFocus
        Exit Sub
      End If
      
      If Trim(txtCidade.Text) = "" Then
        MsgBox "Preencha o nome da Cidade.", vbOKOnly, "Xsol Aquecedores"
        txtCidade.SetFocus
        Exit Sub
      End If
      
      If Trim(txtCep.Text) = "" Then
        MsgBox "Preencha o CEP.", vbOKOnly, "Xsol Aquecedores"
        txtCep.SetFocus
        Exit Sub
      End If
      
      If Trim(txtFone.Text) = "" Then
        MsgBox "Preencha o campo Fone.", vbOKOnly, "Xsol Aquecedores"
        txtFone.SetFocus
        Exit Sub
      End If
      
      If Trim(txtContato.Text) = "" Then
        MsgBox "Preencha o nome do Contato", vbOKOnly, "Xsol Aquecedores"
        txtContato.SetFocus
        Exit Sub
      End If
      
      If Trim(txtMaterial.Text) = "" Then
        MsgBox "Preencha o tipo de material.", vbOKOnly, "Xsol Aquecedores"
        txtMaterial.SetFocus
        Exit Sub
      End If
      
'
   If Cmd_Inclui.Caption = "&Cancelar" Then 'Faz Insert
      
    Set rs = Return_Recordset(Return_SqlString(1))
    rs.AddNew
    
    rs.Fields("NomeCli") = txtNome.Text
    rs.Fields("End") = txtEndereco.Text
    rs.Fields("Bairro") = txtBairro.Text
    rs.Fields("Cidade") = txtCidade.Text
    rs.Fields("Cep") = txtCep.Text
    rs.Fields("CNPJ") = txtCnpj.Text
    rs.Fields("Fone") = txtFone.Text
    'rs.Fields("Fax") = txtFax.Text
    rs.Fields("Contato") = txtContato.Text
    rs.Fields("Email") = txtEmail.Text
    rs.Fields("Tipo_Mat") = txtMaterial.Text
      
    Cmd_Inclui.Caption = "&Inserir"
    Cmd_Inclui.enabled = True
    Cmd_Edita.enabled = True
    Cmd_Exclui.enabled = True
    Cmd_Salva.enabled = False
    
    
   Else 'Faz update
    Set rs = Return_Recordset(Return_SqlString(2, Grid_Fornecedores.TextMatrix(Grid_Fornecedores.RowSel, 0)))
   End If

    'txtProdutoDescricao.enabled = False
    'txtProdutoDescricao.BackColor = &HC0C0C0
    
    txtNome.BackColor = &HC0C0C0
    txtNome.enabled = False
    
    txtEndereco.BackColor = &HC0C0C0
    txtEndereco.enabled = False
    
    txtBairro.BackColor = &HC0C0C0
    txtBairro.enabled = False
    
    txtCidade.BackColor = &HC0C0C0
    txtCidade.enabled = False
    
    txtCep.BackColor = &HC0C0C0
    txtCep.enabled = False
    
    txtCnpj.BackColor = &HC0C0C0
    txtCnpj.enabled = False
    
    txtFone.BackColor = &HC0C0C0
    txtFone.enabled = False
    
    'txtFax.BackColor = &HC0C0C0
    'txtFax.enabled = False
    
    txtContato.BackColor = &HC0C0C0
    txtContato.enabled = False
    
    txtEmail.BackColor = &HC0C0C0
    txtEmail.enabled = False
    
    txtMaterial.BackColor = &HC0C0C0
    txtMaterial.enabled = False
    
    Cmd_Inclui.enabled = True
    
    Cmd_Edita.Caption = "&Editar"
    Cmd_Edita.enabled = True
    
    Cmd_Exclui.enabled = True
    Cmd_Salva.enabled = False

    panFornecedores.enabled = True
    
    'panEquipamentos.enabled = False
    btnEquipFornAdd.enabled = False
    btnEquipFornRemove.enabled = False
    txtEquipFornDescricao.BackColor = &HC0C0C0
    txtEquipFornDescricao.enabled = False
    
    txtEquipFornValor.BackColor = &HC0C0C0
    txtEquipFornValor.enabled = False


   Call Record_Record(rs)
   rs.Update

   rs.Close
   Set rs = Nothing
   
   Grid_Fornecedores.Visible = True
   Grid_FornecedorEquipamento.Visible = True

   
   Call Fill_Grid(Return_Recordset(Return_SqlString(1)), txtNome.Text)
End Sub

Private Sub carregaMaterial(rs As ADODB.Recordset, Optional code As String)
    
    txtEquipFornDescricao.Clear

    If rs.RecordCount > 0 Then
        With txtEquipFornDescricao
            If rs.EOF Or rs.BOF Then
               Exit Sub
            End If

            Do Until rs.EOF
                .AddItem (rs.Fields("Descricao"))
                rs.MoveNext
            Loop
        End With
    Else
        txtProdutoDescricao.Clear
    End If

    rs.Close
    Set rs = Nothing
End Sub


Private Sub Form_Activate()

txtEquipFornValor.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Grid_Fornecedores.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Grid_FornecedorEquipamento.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtMaterial.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtEmail.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtContato.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtFax.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtFone.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtCnpj.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtCep.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtCidade.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtBairro.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtEndereco.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtNome.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtCodigo.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
End Sub

Private Sub Form_Load()
   
   Call Configure_Grid
   Call Fill_Grid(Return_Recordset(Return_SqlString(1)))
   
   Call Configure_GridFornEquip
   Call Fill_GridFornEquuip(Return_Recordset(Return_SqlString(3, txtCodigo.Text)))
   
   Call carregaMaterial(Return_Recordset(Return_SqlString(6)))
   
End Sub

Sub Show_Record_Detail(rs As ADODB.Recordset)

    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    txtCodigo.Text = rs.Fields("Codigo")
    txtNome.Text = rs.Fields("NomeCli")
    txtEndereco.Text = rs.Fields("End")
    txtBairro.Text = rs.Fields("Bairro")
    txtCidade.Text = rs.Fields("Cidade")
    txtCep.Text = rs.Fields("Cep")
    txtCnpj.Text = rs.Fields("CNPJ")
    txtFone.Text = rs.Fields("Fone")
    'txtFax.Text = rs.Fields("Fax")
    txtContato.Text = rs.Fields("Contato")
    txtEmail.Text = rs.Fields("Email")
    txtMaterial.Text = rs.Fields("Tipo_Mat")
    
    rs.Close
    Set rs = Nothing
End Sub

Sub Show_Detalhes_Equipamento(rs As ADODB.Recordset)

    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    txtEquipFornDescricao.Text = rs.Fields("Descricao")
    txtEquipFornValor.Text = rs.Fields("ValorCompra")
    
    rs.Close
    Set rs = Nothing
End Sub


Private Sub Grid_FornecedorEquipamento_Click()
    Call Show_Detalhes_Equipamento(Return_Recordset(Return_SqlString(10, Grid_FornecedorEquipamento.TextMatrix(Grid_FornecedorEquipamento.RowSel, 1))))
End Sub

Private Sub Grid_Fornecedores_Click()
    
    With Grid_Fornecedores
        If .RowSel = 1 Then
            Exit Sub
        End If
        Call Show_Record_Detail(Return_Recordset(Return_SqlString(2, .TextMatrix(.RowSel, 0))))
        Call Fill_GridFornEquuip(Return_Recordset(Return_SqlString(3, txtCodigo.Text)))
    End With

End Sub

Private Sub carregaValor(rs As ADODB.Recordset, Optional code As String)
    
    'txtEquipFornValor.Text =
        
    If Not rs.BOF Then
        txtEquipFornValor.Text = Format(Str(rs.Fields("ValorCompra")), "###,###,##0.00")
    Else
        txtEquipFornValor.Text = ""
    End If
    
    rs.Close
    Set rs = Nothing
End Sub


Private Sub txtEquipFornDescricao_Click()
    Call carregaValor(Return_Recordset(Return_SqlString(7, txtEquipFornDescricao.Text)))
    btnAtualizaValorCompraEquipamento.ToolTipText = "Atualizar o valor do equipamento " & txtEquipFornDescricao.Text & " para R$" & txtEquipFornValor.Text
    
End Sub

Private Sub txtEquipFornValor_Change()
    btnAtualizaValorCompraEquipamento.ToolTipText = "Atualizar o valor do equipamento " & txtEquipFornDescricao.Text & " para R$" & txtEquipFornValor.Text
End Sub

Private Sub txtEquipFornValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btnEquipFornAdd_Click
    End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 180 Then
        KeyAscii = 0
    End If
End Sub
