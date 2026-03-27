VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Produtos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Equipamentos"
   ClientHeight    =   6270
   ClientLeft      =   2505
   ClientTop       =   2805
   ClientWidth     =   13425
   Icon            =   "frm_Produtos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   13425
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame panFornecedor 
      Enabled         =   0   'False
      Height          =   2775
      Left            =   135
      TabIndex        =   9
      Top             =   3300
      Width           =   13140
      Begin VB.Frame panCego 
         Height          =   2565
         Left            =   105
         TabIndex        =   26
         Top             =   120
         Width           =   5760
      End
      Begin VB.CommandButton cmdAtualizar 
         Caption         =   "Atualizar"
         Enabled         =   0   'False
         Height          =   270
         Left            =   2670
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2355
         Width           =   780
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "<< Remover"
         Height          =   315
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2295
         Width           =   1020
      End
      Begin VB.TextBox txtValorVenda 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2295
         Width           =   1065
      End
      Begin VB.TextBox txtPorcentagemLucro 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1470
         TabIndex        =   16
         Top             =   1590
         Width           =   1065
      End
      Begin VB.TextBox txtValorCompra 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1470
         TabIndex        =   13
         Top             =   990
         Width           =   1065
      End
      Begin VB.CommandButton btnAdicionar 
         BackColor       =   &H0080FF80&
         Caption         =   "Adicionar >>"
         Height          =   315
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1890
         Width           =   1020
      End
      Begin VB.ComboBox cboFornecedor 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   255
         Width           =   4005
      End
      Begin MSFlexGridLib.MSFlexGrid gridFornecedores 
         Height          =   2085
         Left            =   5940
         TabIndex        =   10
         Top             =   375
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   3678
         _Version        =   393216
         Rows            =   1
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
      End
      Begin VB.Label lblInfoEditFornecedor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Selecione um fornececedor acima para editar os valores de compra e lucro."
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   5940
         TabIndex        =   29
         Top             =   2490
         Visible         =   0   'False
         Width           =   5340
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   1365
         TabIndex        =   27
         Top             =   630
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblCodFornecedor 
         AutoSize        =   -1  'True
         Caption         =   "Codigo do fornecedor"
         Height          =   195
         Left            =   1950
         TabIndex        =   23
         Top             =   630
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor(es) do equipamento acima selecionado:"
         Height          =   195
         Left            =   5985
         TabIndex        =   22
         Top             =   135
         Width           =   3660
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   750
         TabIndex        =   20
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor de venda:"
         Height          =   195
         Left            =   270
         TabIndex        =   18
         Top             =   2370
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lucro (%):"
         Height          =   195
         Left            =   690
         TabIndex        =   15
         Top             =   1665
         Width           =   705
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor de compra:"
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   1080
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Material"
      Height          =   2625
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   13140
      Begin VB.CommandButton btnIncluir 
         BackColor       =   &H0080FF80&
         Caption         =   "Inserir"
         Height          =   420
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1500
         Width           =   1005
      End
      Begin VB.CommandButton btnSalvar 
         Caption         =   "Salvar"
         Enabled         =   0   'False
         Height          =   420
         Left            =   1425
         TabIndex        =   7
         Top             =   1515
         Width           =   1050
      End
      Begin VB.CommandButton btnEditar 
         Caption         =   "Editar"
         Height          =   420
         Left            =   4080
         TabIndex        =   6
         Top             =   1515
         Width           =   1065
      End
      Begin VB.CommandButton btnExcluir 
         BackColor       =   &H008080FF&
         Caption         =   "Excluir"
         Height          =   420
         Left            =   2715
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1500
         Width           =   1065
      End
      Begin VB.TextBox txtDescricaoMaterial 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1095
         TabIndex        =   4
         Top             =   750
         Width           =   4275
      End
      Begin MSFlexGridLib.MSFlexGrid gridMateriais 
         Height          =   2130
         Left            =   5985
         TabIndex        =   19
         Top             =   375
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   3757
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         SelectionMode   =   1
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lista de Materiais:"
         Height          =   195
         Left            =   6015
         TabIndex        =   21
         Top             =   165
         Width           =   1275
      End
      Begin VB.Label lblCodigoMaterial 
         AutoSize        =   -1  'True
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1155
         TabIndex        =   3
         Top             =   360
         Width           =   195
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   270
         TabIndex        =   2
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Lbl_Cliente 
         Alignment       =   1  'Right Justify
         Caption         =   "Código:"
         Height          =   255
         Left            =   405
         TabIndex        =   1
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "FORNECEDOR (ES) DO EQUIPAMENTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   150
      TabIndex        =   25
      Top             =   2880
      Width           =   13125
   End
End
Attribute VB_Name = "frm_Produtos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Estoque_DBname As String
Private conn As ADODB.Connection
Public movEditar As Boolean
Public nomeAntes As String
Public nomeDepois As String


Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String) As String
    Dim whereClause As String
    
    If queryType = 1 Then
        Return_SqlString = "SELECT * FROM Fornecedor"
    ElseIf queryType = 2 Then
        Return_SqlString = "SELECT * FROM Material order by descricao"
    ElseIf queryType = 3 Then
        Return_SqlString = "SELECT * FROM Material order by descricao"
    ElseIf queryType = 4 Then
        Return_SqlString = "SELECT * FROM MaterialFornecedor"
    ElseIf queryType = 5 Then
        Return_SqlString = "SELECT a.codFornecedor, b.nomecli, a.valorCompra, a.porcentagemLucro, a.valorVenda FROM MaterialFornecedor a Inner join Fornecedor b on  a.codFornecedor = b.codigo WHERE a.CodMaterial = " & param1
    ElseIf queryType = 6 Then
        Return_SqlString = "SELECT * FROM Material WHERE Codigo = " & param1 & " order by descricao"
    ElseIf queryType = 7 Then
        Return_SqlString = "SELECT * FROM Fornecedor WHERE NomeCli = """ & param1 & """"
    ElseIf queryType = 8 Then
        Return_SqlString = "SELECT * FROM MaterialFornecedor WHERE CodMaterial = " & param1 & " And CodFornecedor = " & param2
    ElseIf queryType = 9 Then
        Return_SqlString = "SELECT * FROM MaterialFornecedor WHERE CodMaterial = " & param1
    ElseIf queryType = 10 Then
        Return_SqlString = "UPDATE Estoque Set Produto = '" & param1 & "' WHERE CodMaterial = " & CDbl(param2)  '//parou aqui
        
    End If
End Function

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


Public Function Return_Recordset(strSql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    Call EstoqueArquivos_Dados
    
    Set conn = Nothing
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & Estoque_DBname & ";"
    
    rs.Open strSql, conn, adOpenStatic, adLockOptimistic
    
    Set Return_Recordset = rs
End Function

Private Sub Record_Record(rs As ADODB.Recordset)
   rs.Fields("Descricao").value = txtDescricaoMaterial.Text
End Sub

Private Sub Record_MaterialFornecedor(rs As ADODB.Recordset)
    rs.Fields("CodMaterial") = lblCodigoMaterial.Caption
    rs.Fields("CodFornecedor") = lblCodFornecedor.Caption
    rs.Fields("CodMaterial") = lblCodigoMaterial.Caption
    rs.Fields("PorcentagemLucro") = txtPorcentagemLucro.Text
    rs.Fields("ValorCompra") = txtValorCompra.Text
    rs.Fields("ValorVenda") = txtValorVenda.Text
End Sub




Private Sub carregaFornecedores(rs As ADODB.Recordset, Optional code As String)
    
    With cboFornecedor
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("NomeCli"))
            rs.MoveNext
        Loop
    End With
    
    rs.Close
    Set rs = Nothing
End Sub


Public Function consultaFornecedor(rs As ADODB.Recordset) As Boolean

    If rs.EOF Or rs.BOF Then
        consultaFornecedor = False
        Exit Function
    Else
        consultaFornecedor = True
    End If
    
'    If rs.Fields("Saldo") <> Null Or rs.Fields("Saldo") >= 0 Or rs.Fields("Saldo") < 0 Then
'        txtVisaoEstoque_Estoque.Text = rs.Fields("Saldo")
'    Else
'        txtVisaoEstoque_Estoque.Text = "0"
'    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub Fill_Grid_FornecedorMaterial(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    lblCodigoMaterial.Caption = gridMateriais.TextMatrix(gridMateriais.RowSel, 0)
    txtDescricaoMaterial.Text = gridMateriais.TextMatrix(gridMateriais.RowSel, 1)
    
    nomeAntes = txtDescricaoMaterial.Text
    
    selectedRow = 2
    With gridFornecedores
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           GoTo sai
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("codFornecedor") & vbTab & rs.Fields("nomeCli") & vbTab & rs.Fields("valorCompra")) & vbTab & rs.Fields("porcentagemLucro") & vbTab & rs.Fields("valorVenda")
            If code = rs.Fields("codFornecedor") Then
                selectedRow = .Rows - 1
            End If
            
            
            If bkColor <> 0 Then
                .row = .Rows - 1
                For col = .FixedCols To .Cols - 1
                    .col = col
                    .CellBackColor = bkColor
                Next
            End If
            
            rs.MoveNext
        Loop
        
sai:
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
            'Call Object_Controller(1)
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub carregaCodFornecedor(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    If rs.EOF Or rs.BOF Then
       Exit Sub
    Else
        lblCodFornecedor.Caption = rs.Fields("Codigo")
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub btnAdicionar_Click()
On Error GoTo sai
        
    If cboFornecedor.Text = "" Then
        MsgBox "Selecione um fornecedor"
        cboFornecedor.SetFocus
        Exit Sub
    End If
    
    If Trim(txtValorCompra.Text) = "" Then
        MsgBox "Informe o valor de compra do produto!"
        txtValorCompra.SetFocus
        Exit Sub
    End If
    
    If Trim(txtPorcentagemLucro.Text) = "" Then
        MsgBox "Informe a porcentagem de lucro do produto!"
        txtPorcentagemLucro.SetFocus
        Exit Sub
    End If
    
    If Trim(txtValorVenda.Text) = "" Then
        MsgBox "Informe o valor de venda do produto!"
        txtValorVenda.SetFocus
        Exit Sub
    End If
    
    'If consultaFornecedor(Return_Recordset(Return_SqlString(7, lblCodFornecedor.Caption))) Then
    'End If
    
Dim rs As ADODB.Recordset
    
    Set rs = Return_Recordset(Return_SqlString(4))
    rs.AddNew
    
    rs.Fields("CodMaterial") = lblCodigoMaterial.Caption
    rs.Fields("CodFornecedor") = lblCodFornecedor.Caption
    rs.Fields("CodMaterial") = lblCodigoMaterial.Caption
    rs.Fields("PorcentagemLucro") = txtPorcentagemLucro.Text
    rs.Fields("ValorCompra") = txtValorCompra.Text
    rs.Fields("ValorVenda") = txtValorVenda.Text

   Call Record_MaterialFornecedor(rs)
   rs.Update

   rs.Close
   Set rs = Nothing
    
sai:
    cboFornecedor.ListIndex = -1
    Label9.Visible = False
    lblCodFornecedor.Caption = ""
    txtValorCompra.Text = ""
    txtPorcentagemLucro.Text = ""
    txtValorVenda.Text = ""
    
    Call Fill_Grid_FornecedorMaterial(Return_Recordset(Return_SqlString(5, lblCodigoMaterial.Caption)))
    
End Sub

Private Sub btnEditar_Click()
    
    If gridMateriais.Rows = 1 Then
        Exit Sub
    End If
    
    If lblCodigoMaterial.Caption = "..." Then
        MsgBox "Selecione um material", vbOKOnly, "xSOL Aquecedores"
        Exit Sub
    End If

    movEditar = True
    
    panCego.Visible = False
    panFornecedor.enabled = True


    If btnEditar.Caption = "Editar" Then
        btnEditar.Caption = "Cancelar"
        txtDescricaoMaterial.enabled = True
        btnIncluir.enabled = False
        btnSalvar.enabled = True
        btnExcluir.enabled = False
        panFornecedor.enabled = True
        lblInfoEditFornecedor.Visible = True
    Else
        btnEditar.Caption = "Editar"
        txtDescricaoMaterial.enabled = False
        btnIncluir.enabled = True
        btnSalvar.enabled = False
        btnExcluir.enabled = True
        cboFornecedor.ListIndex = -1
        panCego.Visible = True
        panFornecedor.enabled = False
        lblInfoEditFornecedor.Visible = False
    End If
    
    

End Sub

Private Sub btnExcluir_Click()
        
    If gridMateriais.Rows = 1 Then
        Exit Sub
    End If
    
    If lblCodigoMaterial.Caption = "..." Then
        MsgBox "Selecione um material", vbOKOnly, "xSOL Aquecedores"
        Exit Sub
    End If
   
   If gridMateriais.Rows > 1 Then
    If MsgBox("Confirma a Exclusão do produto " & gridMateriais.TextMatrix(gridMateriais.RowSel, 1) & " ? ", vbYesNo, "Excluir Registro") = vbYes Then
       Call Return_Recordset(Return_SqlString(6, lblCodigoMaterial.Caption)).Delete
       
       
       'Call Return_Recordset(Return_SqlString(9, lblCodigoMaterial.Caption)).Delete
       Dim X As Integer
       For X = 0 To Return_Recordset(Return_SqlString(9, lblCodigoMaterial.Caption)).RecordCount
            If Not Return_Recordset(Return_SqlString(9, lblCodigoMaterial.Caption)).EOF Then
                Return_Recordset(Return_SqlString(9, lblCodigoMaterial.Caption)).Delete
            End If
       Next X
       
       Call Fill_Grid_Material(Return_Recordset(Return_SqlString(2)))
       Call Fill_Grid_FornecedorMaterial(Return_Recordset(Return_SqlString(5, lblCodigoMaterial.Caption)))
       lblCodigoMaterial.Caption = ""
       txtDescricaoMaterial.Text = ""
    End If
   Else
    lblCodigoMaterial.Caption = ""
   End If

End Sub

Private Sub btnIncluir_Click()

    movEditar = False

    lblCodigoMaterial.Caption = ""
    txtDescricaoMaterial.Text = ""

    If btnIncluir.Caption = "Inserir" Then
        btnIncluir.Caption = "Cancelar"
        txtDescricaoMaterial.enabled = True
        txtDescricaoMaterial.SetFocus
        cboFornecedor.enabled = True
        btnSalvar.enabled = True
        
        btnExcluir.enabled = False
        btnEditar.enabled = False
        panFornecedor.enabled = False
        
        btnAdicionar.enabled = True
    Else
        btnIncluir.Caption = "Inserir"
        txtDescricaoMaterial.enabled = False
        txtDescricaoMaterial.Text = ""
        
        btnExcluir.enabled = True
        btnEditar.enabled = True
        
        cboFornecedor.enabled = False
        cboFornecedor.ListIndex = -1
        btnSalvar.enabled = False
        btnAdicionar.enabled = False
    End If
    

End Sub

Private Sub Fill_Grid_Material(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    selectedRow = 1
    With gridMateriais
        .Rows = 1
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("Codigo") & vbTab & rs.Fields("Descricao"))
            If code = rs.Fields("Codigo") Then
                selectedRow = .Rows - 1
            End If
            
            If bkColor <> 0 Then
                .row = .Rows - 1
                For col = .FixedCols To .Cols - 1
                    .col = col
                    .CellBackColor = bkColor
                Next
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
            'Call Grid_Estoque_Click
        Else
            'Call Object_Controller(1)
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub


Private Sub btnSalvar_Click()
    'On Error GoTo sai
    If Trim(txtDescricaoMaterial.Text) = "" Then
        MsgBox "Informe a descrição do produto!"
        Exit Sub
    End If
    
    'If consultaFornecedor(Return_Recordset(Return_SqlString(7, lblCodFornecedor.Caption))) Then
    'End If
    
    Dim rs As ADODB.Recordset
    
    If movEditar Then
        Set rs = Return_Recordset(Return_SqlString(6, lblCodigoMaterial.Caption))
    Else
        Set rs = Return_Recordset(Return_SqlString(2))
    End If
    
    If movEditar = False Then
        rs.AddNew
    End If
    
    rs.Fields("Descricao") = txtDescricaoMaterial.Text
    
    'ver aqui se houve mudanca no nome do equipamento e atualiza a tabela de estoque com o novo nome
    If (movEditar) And (txtDescricaoMaterial.Text <> nomeAntes) Then
        atualizaTabelaEstoque
    End If

   Call Record_Record(rs)
   rs.Update

    If movEditar = False Then
        lblCodigoMaterial.Caption = rs.Fields("Codigo")
    End If

   rs.Close
   Set rs = Nothing
    
    If movEditar Then
        MsgBox "Material EDITADO com sucesso", vbOKOnly, "xSOL Aquecedores"
    Else
        MsgBox "Material SALVO com sucesso", vbOKOnly, "xSOL Aquecedores"
    End If



Call Fill_Grid_Material(Return_Recordset(Return_SqlString(2)))
panFornecedor.Visible = True

sai:
    
    btnSalvar.enabled = False
    
    btnIncluir.Caption = "Inserir"
    btnIncluir.enabled = True
    
    btnExcluir.enabled = True
    btnEditar.enabled = True
    
    btnEditar.Caption = "Editar"
    
    txtDescricaoMaterial.Text = ""
    panCego.Visible = True
    
    If movEditar Then
        movEditar = False
        panFornecedor.enabled = False
    End If
    
    lblInfoEditFornecedor.Visible = False
    
End Sub

Private Sub atualizaTabelaEstoque()
    Dim rs As ADODB.Recordset
    
    'On Error GoTo erro
    Set rs = Return_Recordset(Return_SqlString(10, txtDescricaoMaterial.Text, lblCodigoMaterial.Caption))
    Set rs = Nothing
    
    

'(Return_Recordset(Return_SqlString(10, , )))
    
End Sub

Private Sub Configure_Grid_Material()
    With gridMateriais
        .Cols = 2
        .FixedCols = 0
        .ColWidth(0) = 600
        .ColWidth(1) = 6000
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Descrição"
        .RowHeight(1) = 0
    End With
End Sub

Private Sub Configure_Grid_Fornecedores()
    With gridFornecedores
        .Cols = 5
        .FixedCols = 0
        .ColWidth(0) = 900
        .ColWidth(1) = 3000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Cod. Forn."
        .TextMatrix(0, 1) = "Fornecedor"
        .TextMatrix(0, 2) = "R$ Compra"
        .TextMatrix(0, 3) = "% Lucro"
        .TextMatrix(0, 4) = "R$ Venda"
        .RowHeight(1) = 0
    End With
End Sub


Private Sub cboFornecedor_Click()
    Call carregaCodFornecedor(Return_Recordset(Return_SqlString(7, cboFornecedor.Text)))
    Label9.Visible = True
    lblCodFornecedor.Visible = True
End Sub

Private Sub cmdAtualizar_Click()
Dim rs As ADODB.Recordset
    
    Set rs = Return_Recordset(Return_SqlString(9, lblCodigoMaterial.Caption))
    'rs.AddNew
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        cmdAtualizar.enabled = False
        Exit Sub
    End If
    
    rs.Fields("ValorCompra") = txtValorCompra.Text
    rs.Fields("PorcentagemLucro") = txtPorcentagemLucro.Text
    rs.Fields("ValorVenda") = txtValorVenda.Text

   Call Record_MaterialFornecedor(rs)
   rs.Update

   rs.Close
   Set rs = Nothing
    
sai:
    
    lblCodFornecedor.Caption = ""
    txtValorCompra.Text = ""
    txtPorcentagemLucro.Text = ""
    txtValorVenda.Text = ""
    
    Call Fill_Grid_FornecedorMaterial(Return_Recordset(Return_SqlString(5, lblCodigoMaterial.Caption)))
    
    cmdAtualizar.enabled = False

End Sub

Private Sub Command1_Click()
   If gridFornecedores.Rows > 2 Then
    If MsgBox("Confirma a exclusão do fornecedor " & gridFornecedores.TextMatrix(gridFornecedores.RowSel, 0) & " para o produto " & gridMateriais.TextMatrix(gridMateriais.RowSel, 0) & " ?", vbYesNo, "Excluir Registro") = vbYes Then
       Call Return_Recordset(Return_SqlString(8, lblCodigoMaterial.Caption, gridFornecedores.TextMatrix(gridFornecedores.RowSel, 0))).Delete
       Call Fill_Grid_FornecedorMaterial(Return_Recordset(Return_SqlString(5, gridMateriais.TextMatrix(gridMateriais.RowSel, 0))))
    End If
   Else
    lblCodFornecedor.Caption = ""
   End If

    cboFornecedor.ListIndex = -1

    lblCodFornecedor.Caption = ""
    txtValorCompra.Text = ""
    txtPorcentagemLucro.Text = ""
    txtValorVenda.Text = ""

End Sub

Private Sub Form_Activate()
txtValorVenda.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtPorcentagemLucro.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtValorCompra.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
txtDescricaoMaterial.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
cboFornecedor.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
gridFornecedores.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
gridMateriais.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
End Sub

Private Sub Form_Load()
    Call Configure_Grid_Material
    Call Fill_Grid_Material(Return_Recordset(Return_SqlString(2)))
    
    Call Configure_Grid_Fornecedores
    Call carregaFornecedores(Return_Recordset(Return_SqlString(1)))
End Sub


Private Sub gridFornecedores_Click()

    If panCego.Visible = False And gridFornecedores.Rows > 2 Then
        cboFornecedor.Text = gridFornecedores.TextMatrix(gridFornecedores.RowSel, 1)
        txtValorCompra.Text = gridFornecedores.TextMatrix(gridFornecedores.RowSel, 2)
        txtPorcentagemLucro.Text = gridFornecedores.TextMatrix(gridFornecedores.RowSel, 3)
    End If
    
End Sub

Private Sub gridMateriais_Click()
    If gridMateriais.Rows = 1 Then
        Exit Sub
    End If
    
    Call Fill_Grid_FornecedorMaterial(Return_Recordset(Return_SqlString(5, gridMateriais.TextMatrix(gridMateriais.RowSel, 0))))
End Sub

Private Sub txtDescricaoMaterial_LostFocus()
    
    txtDescricaoMaterial.Text = UCase(txtDescricaoMaterial.Text)
    txtDescricaoMaterial.Text = Replace(txtDescricaoMaterial.Text, """", "'")
    txtDescricaoMaterial.Text = UCase(txtDescricaoMaterial.Text)
    
End Sub

Private Sub txtPorcentagemLucro_Change()

    If Trim(txtPorcentagemLucro.Text) = "" Then
        Exit Sub
    End If

    If Not IsNumeric(txtPorcentagemLucro.Text) Then
        MsgBox "Digite somente números no campo porcentagem.", vbOKOnly, "xSol Aquecedores"
        txtPorcentagemLucro.Text = ""
        txtPorcentagemLucro.SetFocus
        Exit Sub
    End If
    
    Dim valCompra As Double
    Dim valPorcentagem As Double
    Dim valVenda As Double
    
    valCompra = txtValorCompra.Text
    valPorcentagem = txtPorcentagemLucro.Text
    valVenda = valCompra + (valCompra * valPorcentagem / 100)
    
    txtValorVenda.Text = valVenda
    
    If txtPorcentagemLucro.Text <> gridFornecedores.TextMatrix(gridFornecedores.RowSel, 3) Then
        cmdAtualizar.enabled = True
    Else
        cmdAtualizar.enabled = False
    End If
    
End Sub

Private Sub txtValorCompra_Change()
    If Trim(txtValorCompra.Text) <> "" Then
        If txtValorCompra.Text <> gridFornecedores.TextMatrix(gridFornecedores.RowSel, 2) Then
            cmdAtualizar.enabled = True
        Else
            cmdAtualizar.enabled = False
        End If
    End If
End Sub

Private Sub txtValorCompra_LostFocus()

    If Trim(txtValorCompra.Text) = "" Then
        Exit Sub
    End If

    If Not IsNumeric(txtValorCompra.Text) Then
        MsgBox "Informe um valor válido para o valor de compra.", vbOKOnly, "xSol Aquecedores"
        txtValorCompra.SetFocus
        Exit Sub
    End If
End Sub



Private Sub txtValorVenda_LostFocus()
    If Not IsNumeric(txtValorVenda.Text) And Trim(txtValorCompra.Text) <> "" Then
        MsgBox "Informe um valor válido para o valor de venda.", vbOKOnly, "xSol Aquecedores"
        txtValorVenda.SetFocus
        Exit Sub
    End If
End Sub
