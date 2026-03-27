VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVeiculos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Veículos"
   ClientHeight    =   2970
   ClientLeft      =   6045
   ClientTop       =   3165
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   9495
   Begin VB.Frame Frame3 
      Height          =   2010
      Left            =   60
      TabIndex        =   7
      Top             =   30
      Width           =   4590
      Begin VB.TextBox txtPlaca 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1560
         Width           =   1470
      End
      Begin VB.TextBox txtAno 
         Enabled         =   0   'False
         Height          =   315
         Left            =   105
         MaxLength       =   4
         TabIndex        =   12
         Top             =   975
         Width           =   1470
      End
      Begin VB.TextBox txtMarcaModelo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   105
         MaxLength       =   50
         TabIndex        =   8
         Top             =   390
         Width           =   2715
      End
      Begin VB.Label lblCodigoVeiculo 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4110
         TabIndex        =   15
         Top             =   165
         Width           =   360
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   3525
         TabIndex        =   14
         Top             =   180
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Placa"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1350
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ano:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   765
         Width           =   330
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Marca/Modelo:"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.Frame panVeiculos 
      Caption         =   "Veículos já cadastrados"
      Height          =   2820
      Left            =   4770
      TabIndex        =   5
      Top             =   75
      Width           =   4620
      Begin MSFlexGridLib.MSFlexGrid gridVeiculos 
         Height          =   2505
         Left            =   45
         TabIndex        =   6
         Top             =   225
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   4419
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Edição"
      Height          =   795
      Left            =   105
      TabIndex        =   0
      Top             =   2085
      Width           =   4575
      Begin VB.CommandButton btnExcluir 
         BackColor       =   &H008080FF&
         Caption         =   "E&xcluir"
         Height          =   360
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   255
         Width           =   750
      End
      Begin VB.CommandButton btnEditar 
         Caption         =   "&Editar"
         Height          =   360
         Left            =   1350
         TabIndex        =   3
         Top             =   270
         Width           =   780
      End
      Begin VB.CommandButton btnSalvar 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   360
         Left            =   2520
         TabIndex        =   2
         Top             =   255
         Width           =   765
      End
      Begin VB.CommandButton btnIncluir 
         BackColor       =   &H0080FF80&
         Caption         =   "&Incluir"
         Height          =   360
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   270
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmVeiculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public movEditar As Boolean
Public W_DB As ADODB.Connection
Public Configuracao_DBname As String
Public globalProximaPKVeiculo As Integer
Public conn As ADODB.Connection
Public dbname As String

Private Sub btnEditar_Click()
    
    If gridVeiculos.Rows = 1 Then
        MsgBox "Selecione um material", vbOKOnly, "xSOL Aquecedores"
        Exit Sub
    End If
    
    If lblCodigoVeiculo.Caption = "" Then
        MsgBox "Selecione um material", vbOKOnly, "xSOL Aquecedores"
        Exit Sub
    End If
    
    movEditar = True
        
    txtMarcaModelo.enabled = True
    txtAno.enabled = True
    txtPlaca.enabled = True
    
    
    btnIncluir.Caption = "Cancelar"
    btnEditar.enabled = False
    btnSalvar.enabled = True
    btnExcluir.enabled = False
    
    panVeiculos.enabled = False

End Sub

Private Sub btnExcluir_Click()

    If gridVeiculos.Rows = 1 Then
        Exit Sub
    End If
    
    If lblCodigoVeiculo.Caption = "0" Then
        MsgBox "Selecione um veículo", vbOKOnly, "xSOL Aquecedores"
        Exit Sub
    End If
   
   If gridVeiculos.Rows > 1 Then
    If MsgBox("Confirma a Exclusão do veículo " & gridVeiculos.TextMatrix(gridVeiculos.RowSel, 1) & " ? ", vbYesNo, "Excluir Registro") = vbYes Then
       Call Return_Recordset(Return_SqlString(5, lblCodigoVeiculo.Caption)).Delete
       
       Call Fill_Grid_Veiculos(Return_Recordset(Return_SqlString(4)))
       lblCodigoVeiculo.Caption = ""
       
       txtMarcaModelo.Text = ""
       txtAno.Text = ""
       txtPlaca.Text = ""
    End If
   Else
    lblCodigoVeiculo.Caption = ""
   End If

End Sub

Private Sub btnIncluir_Click()
    lblCodigoVeiculo.Caption = "0"
    
    movEditar = False

    txtMarcaModelo.Text = ""
    txtAno.Text = ""
    txtPlaca.Text = ""

    If btnIncluir.Caption = "&Incluir" Then
        btnIncluir.Caption = "Cancelar"
        txtMarcaModelo.enabled = True
        txtAno.enabled = True
        txtPlaca.enabled = True
        btnSalvar.enabled = True
        
        btnExcluir.enabled = False
        btnEditar.enabled = False
        panVeiculos.enabled = False
    Else
        btnIncluir.Caption = "&Incluir"
        txtMarcaModelo.enabled = False
        txtAno.enabled = False
        txtPlaca.enabled = False
        
        btnExcluir.enabled = False
        btnEditar.enabled = True
        btnSalvar.enabled = True
        
        panVeiculos.enabled = True
        
    End If
    
    If btnIncluir.Caption = "Cancelar" Then
        'lblCodigoVeiculo.Caption = retornaProximaPKVeiculo(Return_Recordset(Return_SqlString(3)))
        txtMarcaModelo.SetFocus
    Else
        lblCodigoVeiculo.Caption = ""
    End If



End Sub

Private Sub btnSalvar_Click()
        
    'On Error GoTo sai
        If (Trim(txtMarcaModelo.Text) = "") Or (Trim(txtAno.Text) = "") Or (Trim(txtPlaca.Text) = "") Then
            MsgBox "Informe todos os campos!"
            Exit Sub
        End If
        
        Dim rs As ADODB.Recordset
        
        If movEditar Then
            Set rs = Return_Recordset(Return_SqlString(2, lblCodigoVeiculo.Caption))
        Else
            Set rs = Return_Recordset(Return_SqlString(1))
        End If
        
        If movEditar = False Then
            rs.AddNew
        Else
            rs.Update
        End If
        
        rs.Fields("MARCAMODELO") = UCase(txtMarcaModelo.Text)
        rs.Fields("ANO") = UCase(txtAno.Text)
        rs.Fields("PLACA") = UCase(txtPlaca.Text)
    
       'Call Record_Record(rs)
       rs.Update
    
        If movEditar = False Then
            lblCodigoVeiculo.Caption = rs.Fields("Codigo")
        End If
        
        lblCodigoVeiculo.Caption = rs.Fields("CODIGO")
    
       rs.Close
       Set rs = Nothing
        
        'If movEditar Then
        '    MsgBox "Registro EDITADO com sucesso", vbOKOnly, "xSOL Aquecedores"
        'Else
        '    MsgBox "Registro SALVO com sucesso", vbOKOnly, "xSOL Aquecedores"
        'End If
    
        'Call Fill_Grid_Veiculos(Return_Recordset(Return_SqlString(1)))
        'panFornecedor.Visible = True
    
sai:
        
        txtMarcaModelo.enabled = False
        txtAno.enabled = False
        txtPlaca.enabled = False
        
        btnSalvar.enabled = False
        btnExcluir.enabled = True
        btnEditar.enabled = True
        panVeiculos.enabled = True
        
        btnIncluir.Caption = "&Incluir"

        Call Fill_Grid_Veiculos(Return_Recordset(Return_SqlString(4)))

End Sub


Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String) As String
    Dim whereClause As String
    
    If queryType = 1 Then
        Return_SqlString = "SELECT * FROM Veiculos"
    ElseIf queryType = 2 Then
        Return_SqlString = "Select * From VEICULOS Where Codigo = " & param1
    ElseIf queryType = 3 Then
        Return_SqlString = "SELECT max(codigo) + 1 as cod FROM Veiculos"
    ElseIf queryType = 4 Then
        Return_SqlString = "Select * From VEICULOS"
    ElseIf queryType = 5 Then
        Return_SqlString = "Select * From VEICULOS Where Codigo = " & param1
    End If
End Function

Public Function Return_Recordset(strSql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    'Call EstoqueArquivos_Dados
    Call Arquivos_Dados
    
    Set conn = Nothing
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & dbname & ";"
    
    rs.Open strSql, conn, adOpenStatic, adLockOptimistic
    
    Set Return_Recordset = rs
End Function

Private Sub Fill_Grid_Veiculos(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    selectedRow = 1
    With gridVeiculos
        .Rows = 1
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("Codigo") & vbTab & rs.Fields("MARCAMODELO") & vbTab & rs.Fields("ANO") & vbTab & rs.Fields("PLACA"))
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

Private Function Return_Recordsetveiculo(strSql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    rs.Open strSql, W_DB, adOpenStatic, adLockOptimistic
    
    Set Return_Recordsetveiculo = rs
End Function

Function retornaProximaPKVeiculo(rs As ADODB.Recordset) As Integer
    If IsNull(retornaProximaPKVeiculo = rs.Fields("cod")) Then
        retornaProximaPKVeiculo = 1
    Else
        retornaProximaPKVeiculo = rs.Fields("cod")
    End If
    
    rs.Close
    Set rs = Nothing
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
    
    'Image_Folder = GetINISetting("IMAGE", "FOLDER", ConfigFile)
    dbname = Path + FileName
End Sub

Private Sub Form_Load()
    Call ConfigureGridVeiculos
    Call carregaGridVeiculos(Return_Recordset(Return_SqlString(4)))
End Sub

Private Sub ConfigureGridVeiculos()
    With gridVeiculos
        .Cols = 4
        .ColWidth(1) = 300  'Codigo
        .ColWidth(1) = 1500  'Marca/modelo
        .ColWidth(2) = 800 'Ano
        .ColWidth(3) = 1100 'Placa
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Marca/Modelo"
        .TextMatrix(0, 2) = "Ano"
        .TextMatrix(0, 3) = "Placa"
        .RowHeight(1) = 0
    End With
End Sub

Private Sub carregaGridVeiculos(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    'Dim bkColor As Long
    Dim col As Integer
    
    selectedRow = 1
    With gridVeiculos
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("CODIGO") & vbTab & rs.Fields("MARCAMODELO") & vbTab & IIf(IsNull(rs.Fields("ANO")), "", rs.Fields("ANO")) & vbTab & IIf(IsNull(rs.Fields("PLACA")), "", rs.Fields("PLACA")))
            If code = rs.Fields("MARCAMODELO") Then
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

Private Sub gridVeiculos_Click()
    If gridVeiculos.Rows > 2 Then
        If gridVeiculos.TextMatrix(2, 0) <> "" And gridVeiculos.TextMatrix(gridVeiculos.RowSel, 0) <> "" Then
            Call carregaInfoVeiculo(Return_Recordset(Return_SqlString(5, gridVeiculos.TextMatrix(gridVeiculos.RowSel, 0)))) 'validado
            'Return_SqlString = "Select A.DATA, A.CODIGO as codOrcamento, A.VALIDADE, B.NOME_CLI, B.Codigo as codCli, B.CPF, B.Contrato, B.e_mail, B.Fone_celular, B.Fone_Res, B.Fone_Com, B.Tipo, A.Necessidade, A.Entrega, A.Tempo_servico, A.IMG_PATH, A.IMG_TEXT, A.IMG2_PATH, A.IMG2_TEXT, A.Desconto From ORCAMENTO A Inner Join Clientes B on B.codigo = A.COD_CLIENTE Where A.CODIGO = " & param1
        End If
    End If
End Sub

Private Sub carregaInfoVeiculo(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        
        If IsNull(rs.Fields("CODIGO")) Then
            lblCodigoVeiculo.Caption = ""
        Else
            lblCodigoVeiculo.Caption = rs.Fields("CODIGO")
        End If
        
        If IsNull(rs.Fields("MARCAMODELO")) Then
            txtMarcaModelo.Text = ""
        Else
            txtMarcaModelo.Text = rs.Fields("MARCAMODELO")
        End If
        
        If IsNull(rs.Fields("ANO")) Then
            txtAno.Text = ""
        Else
            txtAno.Text = rs.Fields("ANO")
        End If
        
        If IsNull(rs.Fields("PLACA")) Then
            txtPlaca.Text = ""
        Else
            txtPlaca.Text = rs.Fields("PLACA")
        End If
        
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

