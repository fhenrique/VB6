VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRecibos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibos"
   ClientHeight    =   7875
   ClientLeft      =   3360
   ClientTop       =   2385
   ClientWidth     =   12510
   Icon            =   "frmRecibos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   12510
   Begin VB.Frame Frame12 
      Height          =   1320
      Left            =   5925
      TabIndex        =   18
      Top             =   6480
      Width           =   6510
      Begin VB.CommandButton btnVerTodos 
         Caption         =   "Ver todos"
         Height          =   360
         Left            =   4440
         TabIndex        =   30
         Top             =   840
         Width           =   825
      End
      Begin VB.TextBox txtPesquisaClienteCodigo 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1005
         TabIndex        =   29
         Top             =   360
         Width           =   765
      End
      Begin VB.CommandButton btnPesquisaCliente2 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5910
         TabIndex        =   28
         Top             =   345
         Width           =   420
      End
      Begin VB.TextBox txtPesquisaClienteNome 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1830
         TabIndex        =   27
         Top             =   360
         Width           =   3990
      End
      Begin VB.OptionButton optnData 
         Caption         =   "Data"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   915
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton optnCliente 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   420
         Width           =   900
      End
      Begin VB.CommandButton btnPesquisar 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Pesquisar"
         Height          =   360
         Left            =   5385
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   840
         Width           =   945
      End
      Begin MSComCtl2.DTPicker dtpDataDe 
         Height          =   270
         Left            =   1320
         TabIndex        =   21
         Top             =   885
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   476
         _Version        =   393216
         Format          =   101318657
         CurrentDate     =   44349
      End
      Begin MSComCtl2.DTPicker dtpDataAte 
         Height          =   270
         Left            =   2985
         TabIndex        =   22
         Top             =   870
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   476
         _Version        =   393216
         Format          =   101318657
         CurrentDate     =   44349
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   1020
         TabIndex        =   24
         Top             =   930
         Width           =   255
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Até"
         Height          =   195
         Left            =   2655
         TabIndex        =   23
         Top             =   945
         Width           =   240
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Edição"
      Height          =   795
      Left            =   120
      TabIndex        =   11
      Top             =   6990
      Width           =   5715
      Begin VB.CommandButton btnGerarRecibo 
         Caption         =   "Gerar recibo"
         Height          =   360
         Left            =   4455
         TabIndex        =   26
         Top             =   270
         Width           =   1125
      End
      Begin VB.CommandButton btnIncluir 
         BackColor       =   &H0080FF80&
         Caption         =   "Incluir"
         Height          =   360
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   270
         Width           =   675
      End
      Begin VB.CommandButton btnSalvar 
         Caption         =   "Salvar"
         Enabled         =   0   'False
         Height          =   360
         Left            =   2100
         TabIndex        =   14
         Top             =   270
         Width           =   660
      End
      Begin VB.CommandButton btnEditar 
         Caption         =   "Editar"
         Height          =   360
         Left            =   1095
         TabIndex        =   13
         Top             =   270
         Width           =   615
      End
      Begin VB.CommandButton btnExcluir 
         BackColor       =   &H008080FF&
         Caption         =   "Excluir"
         Height          =   360
         Left            =   3135
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   255
         Width           =   750
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6915
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   5760
      Begin VB.TextBox txtReferencia 
         Enabled         =   0   'False
         Height          =   915
         Left            =   165
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   2595
         Width           =   5400
      End
      Begin VB.CommandButton btnPesquisaCliente 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   945
         TabIndex        =   17
         Top             =   1680
         Width           =   420
      End
      Begin VB.TextBox txtClienteNome 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1395
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1695
         Width           =   4155
      End
      Begin VB.TextBox txtData 
         Enabled         =   0   'False
         Height          =   315
         Left            =   165
         MaxLength       =   10
         TabIndex        =   5
         Top             =   960
         Width           =   1230
      End
      Begin VB.TextBox txtClienteCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   165
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1710
         Width           =   765
      End
      Begin VB.TextBox txtValor 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1725
         MaxLength       =   15
         TabIndex        =   3
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label lblCPF 
         AutoSize        =   -1  'True
         Caption         =   "CPF"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1410
         TabIndex        =   33
         Top             =   2040
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Referente a"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   2355
         Width           =   840
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   750
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   165
         TabIndex        =   9
         Top             =   1485
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   1710
         TabIndex        =   8
         Top             =   765
         Width           =   360
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lblCodRecibo 
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
         Left            =   735
         TabIndex        =   6
         Top             =   315
         Width           =   360
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Lista de recibos já emitidos"
      Height          =   6405
      Left            =   5925
      TabIndex        =   0
      Top             =   45
      Width           =   6495
      Begin MSFlexGridLib.MSFlexGrid gridRecibos 
         Height          =   6150
         Left            =   60
         TabIndex        =   1
         Top             =   210
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   10848
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         SelectionMode   =   1
      End
   End
End
Attribute VB_Name = "frmRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public conn As ADODB.Connection
Public strReciboClienteCodigo As String
Public strReciboClienteNome As String
Public bEditando As Boolean
Public strTextoDoRecibo As String

Private Sub btnEditar_Click()
    
    If lblCodRecibo.Caption = "" Then
        MsgBox "Selecine um recibo na listagem!", vbExclamation, "Validação"
        Exit Sub
    End If
    
    txtData.enabled = True
    txtValor.enabled = True
    btnPesquisaCliente.enabled = True
    txtReferencia.enabled = True
    
    btnIncluir.enabled = False
    btnEditar.enabled = False
    btnSalvar.enabled = True
    
    btnExcluir.Caption = "Cancelar"
    btnGerarRecibo.enabled = False
    
    
    bEditando = True
End Sub

Private Sub btnExcluir_Click()
    If (gridRecibos.Rows <= 2) Or (txtData.Text = "") Then
        MsgBox "Não há recibo selecionado!", vbExclamation, "Validação"
        Exit Sub
    End If
    
    
    If btnExcluir.Caption = "Excluir" Then
        
        If MsgBox("Confirma a exclusão deste recibo?", vbYesNo, "Confirmação") = vbYes Then
            'frmRecibos.strReciboClienteCodigo = gridClientes.TextMatrix(gridClientes.RowSel, 0)
            
            Call Return_Recordset(Return_SqlString(3, gridRecibos.TextMatrix(gridRecibos.RowSel, 0))).Delete
            
            'atualiza o grid
            Call carregaGridRecibos(Return_Recordset(Return_SqlString(4)))
            
            MsgBox "Recibo excluído!"
            Call limpaTela
        End If
        

    ElseIf btnExcluir.Caption = "Cancelar" Then
        Call limpaTela
        gridRecibos.enabled = True
        
        btnIncluir.enabled = True
        btnEditar.enabled = True
        btnSalvar.enabled = False
        btnGerarRecibo.enabled = True
        
        btnExcluir.Caption = "Excluir"
        
        btnPesquisaCliente.enabled = False
    End If

End Sub

Private Sub btnGerarRecibo_Click()
    If (gridRecibos.Rows <= 2) Or (txtData.Text = "") Then
        MsgBox "Não há recibo selecionado!", vbExclamation, "Validação"
        Exit Sub
    End If
    
    If txtData.Text = "" Then
        MsgBox "Selecione um recibo na lista", vbExclamation, "Validação"
        Exit Sub
    End If
    
    'relRecibo.lblDataHora.Caption = Date & " " & Time
    relRecibo.lblNumeroRecibo.Caption = lblCodRecibo.Caption
    relRecibo.lblNumeroRecibo2.Caption = lblCodRecibo.Caption
    
    relRecibo.lblValorRecibo.Caption = FormatCurrency(txtValor.Text)
    relRecibo.lblValorRecibo2.Caption = FormatCurrency(txtValor.Text)
    
    strTextoDoRecibo = "Recebi(emos) de " & txtClienteNome.Text & ", CPF/CNPJ n° " & lblCPF.Caption & ", a importância de " & Extenso(txtValor.Text) & " referente a " & txtReferencia.Text & ". "
    strTextoDoRecibo = strTextoDoRecibo & "E, para maior clareza firmo o presente recibo para que produza os seus efeitos, dando plena, rasa e irrevogável quitação do valor recebido."
    relRecibo.lblTexto.Caption = strTextoDoRecibo
    relRecibo.lblTexto2.Caption = strTextoDoRecibo
    
    relRecibo.lblDataSP.Caption = "São Paulo, " & Day(Date) & " de " & nomeMes(Date) & " de " & Year(Date)
    relRecibo.lblDataSP2.Caption = "São Paulo, " & Day(Date) & " de " & nomeMes(Date) & " de " & Year(Date)
    
    relRecibo.lblClienteAssinatura.Caption = txtClienteNome.Text
    relRecibo.lblClienteAssinatura2.Caption = txtClienteNome.Text
    
    
    
    relRecibo.Show
End Sub

Private Sub btnIncluir_Click()
    Dim proximoNumeroRecibo As Integer
    
    bEditando = False
        
    btnIncluir.enabled = False
    btnEditar.enabled = False
    btnSalvar.enabled = True
    btnExcluir.Caption = "Cancelar"
    
    proximoNumeroRecibo = retornaProximoNumeroRecibo(Return_Recordset(Return_SqlString(1)))

    Call limpaTela
    
    txtData.Text = Date
    
    lblCodRecibo.Caption = proximoNumeroRecibo
    bEditando = False
    gridRecibos.enabled = False
    
    txtData.enabled = True
    txtValor.enabled = True
    
    txtReferencia.enabled = True
    
    btnPesquisaCliente.enabled = True
    
    txtValor.SetFocus
    
End Sub


Private Sub btnPesquisaCliente_Click()
    
    frmConsultaCliente.Show (1)
    txtClienteCodigo.Text = strReciboClienteCodigo
    txtClienteNome.Text = strReciboClienteNome
End Sub


Private Sub btnPesquisaCliente2_Click()
    frmConsultaCliente.Show (1)
    txtPesquisaClienteNome.Text = strReciboClienteNome
    txtPesquisaClienteCodigo.Text = strReciboClienteCodigo
    
    btnPesquisar_Click
End Sub

Private Sub btnPesquisar_Click()
    If optnCliente = True Then
        Call carregaGridRecibos(Return_Recordset(Return_SqlString(6, txtPesquisaClienteCodigo.Text)), "pesquisa")
    Else
    
        If dtpDataDe.value > dtpDataAte.value Then
            MsgBox "A data inicial não pode ser maior que a final!", vbExclamation, "Validação"
            Exit Sub
        End If
    
        Call carregaGridRecibos(Return_Recordset(Return_SqlString(7, dtpDataDe.value, dtpDataAte.value)), "pesquisa")
    End If
End Sub

Private Sub btnSalvar_Click()
    If RTrim(LTrim(txtValor.Text)) = "" Then
        MsgBox "Informe o valor do recibo!", vbInformation, "Validação"
        Exit Sub
    End If
    
    If txtClienteCodigo.Text = "" Then
        MsgBox "Informe o código do cliente!", vbInformation, "Validação"
        Exit Sub
    End If
    
    If RTrim(LTrim(txtReferencia.Text)) = "" Then
        MsgBox "Informe a referência do recibo!", vbInformation, "Validação"
        Exit Sub
    End If
    
    Call gravaRecibo
    
    limpaTela
    
    btnIncluir.enabled = True
    btnEditar.enabled = True
    btnSalvar.enabled = False
    btnExcluir.enabled = True
    
    gridRecibos.enabled = True
    btnGerarRecibo.enabled = True
    
    
    
End Sub

Private Sub btnVerTodos_Click()
    Call carregaGridRecibos(Return_Recordset(Return_SqlString(4)))
End Sub


Private Sub Form_Load()
    Call ConfigureGridRecibos
    
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
    
    Call carregaGridRecibos(Return_Recordset(Return_SqlString(4)))
    
    dtpDataDe.value = Date - 30
    dtpDataAte.value = Date
    
End Sub

Private Sub ConfigureGridRecibos()
    With gridRecibos
        .Cols = 4
        .ColWidth(0) = 800  'CODIGO
        .ColWidth(1) = 1000 'DATA
        .ColWidth(2) = 3400 'COD_CLIENTE
        .ColWidth(3) = 1050 'VALOR_GERAL
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Data"
        .TextMatrix(0, 2) = "Cliente"
        .TextMatrix(0, 3) = "Valor"
        .RowHeight(1) = 0
    End With
End Sub

Private Function retornaProximoNumeroRecibo(rs As ADODB.Recordset) As Integer
    
    If IsNull((rs.Fields("MAXCODE"))) Then
        retornaProximoNumeroRecibo = 1
    Else
        retornaProximoNumeroRecibo = rs.Fields("MAXCODE") + 1
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function Return_Recordset(strSql As String) As ADODB.Recordset

    If strSql = "" Then
        Exit Function
    End If
    
    Dim rs2 As New ADODB.Recordset
    Dim qtd As Integer
    
    Call Arquivos_Dados
    
    rs2.Open strSql, conn, adOpenStatic, adLockOptimistic
    qtd = rs2.RecordCount
    
    Set Return_Recordset = rs2
End Function

Private Sub Arquivos_Dados()
    Dim ConfigFile As String
    Dim Path As String
    Dim FileName As String
    
    ConfigFile = App.Path & "\CONFIG.INI"
    
    Path = RTrim(GetINISetting("DATABASE", "PATH", ConfigFile))
    If Right(Path, 1) <> "\" Then
       Path = Path + "\"
    End If
    FileName = GetINISetting("DATABASE", "FILENAME", ConfigFile)

    dbname = Path + FileName
End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String, Optional param3 As String) As String
    If queryType = 1 Then
        Return_SqlString = "Select max(CODIGO) as MAXCODE from Recibos"
    ElseIf queryType = 2 Then
        Return_SqlString = "Select * from Recibos"
    ElseIf queryType = 3 Then
        Return_SqlString = "Select * from Recibos Where CODIGO = " & param1
    ElseIf queryType = 4 Then
        Return_SqlString = "Select r.codigo, r.data , c.nome_cli as cliente, r.valor, r.referencia from recibos as r inner join clientes as c ON (c.codigo = r.cod_cliente)"
    ElseIf queryType = 5 Then
        If param1 <> "" Then
            Return_SqlString = "Select r.codigo, r.data , r.cod_cliente, c.nome_cli as nome, r.valor, r.referencia, c.cpf  from recibos as r inner join clientes as c ON (c.codigo = r.cod_cliente) where r.codigo = " & param1
        End If
    ElseIf queryType = 6 Then
        Return_SqlString = "Select r.codigo, r.data , r.cod_cliente, c.nome_cli as nome, r.valor from recibos as r inner join clientes as c ON (c.codigo = r.cod_cliente) where r.cod_cliente = " & param1
    ElseIf queryType = 7 Then
        Return_SqlString = "Select r.codigo, r.data , r.cod_cliente, c.nome_cli as nome, r.valor from recibos as r inner join clientes as c ON (c.codigo = r.cod_cliente) where r.data between CDate('" & param1 & "') And CDate('" & param2 & "')"
    End If
    
End Function

Private Sub limpaTela()
    lblCodRecibo.Caption = ""
    txtData.Text = ""
    txtValor.Text = ""
    txtClienteCodigo.Text = ""
    txtClienteNome.Text = ""
    txtReferencia.Text = ""
    lblCPF.Caption = ""
    
    txtData.enabled = False
    txtValor.enabled = False
    txtReferencia.enabled = False
End Sub


Private Sub gridRecibos_SelChange()
    If gridRecibos.Rows > 2 Then
    'If gridRecibos.TextMatrix(2, 0) <> "" And gridOrcamentos.TextMatrix(gridOrcamentos.RowSel, 0) <> "" Then
        Call carregaInfoRecibo(Return_Recordset(Return_SqlString(5, gridRecibos.TextMatrix(gridRecibos.RowSel, 0))))
    'End If
    End If
End Sub

Private Sub Option1_Click()
    btnPesquisaCliente2.enabled = True
    
    dtpDataDe.enabled = True
    dtpDataAte.enabled = True
    
    txtPesquisaClienteCodigo.Text = ""
    txtPesquisaClienteNome.Text = ""
    
    dtpDataDe.value = Date - 30
    dtpDataAte.value = Date

End Sub

Private Sub optnCliente_Click()
    btnPesquisaCliente2.enabled = True
    
    dtpDataDe.enabled = False
    dtpDataAte.enabled = False
End Sub

Private Sub optnData_Click()
    dtpDataDe.enabled = True
    dtpDataAte.enabled = True
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    If validaCaractere2(Str(KeyAscii)) = False Then
        KeyAscii = 0 'Não deixa nenhuma outra caractere ser escrito
    End If
End Sub


Private Sub txtData_LostFocus()
    If (Len(txtData.Text) < 10) Or (Not IsDate(txtData.Text)) Then
        MsgBox "A data informada não está em um formato válido!" & Chr(10) & "Exemplo de formato válido: " & Date, vbExclamation, "Validação"
        txtData.SetFocus
    End If
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If validaCaractere(Str(KeyAscii)) = False Then
        KeyAscii = 0 'Não deixa nenhuma outra caractere ser escrito
    End If
End Sub

Private Sub gravaRecibo()
    Dim rs As ADODB.Recordset
    
    If bEditando = False Then 'está inserindo
        Set rs = Return_Recordset(Return_SqlString(2))
        rs.AddNew
    Else 'está editando
        Set rs = Return_Recordset(Return_SqlString(3, lblCodRecibo.Caption))
    End If
    
    'CODIGO, DATA, COD_CLIENTE, VALOR
    rs.Fields("CODIGO") = lblCodRecibo.Caption
    rs.Fields("DATA") = txtData.Text
    rs.Fields("COD_CLIENTE") = txtClienteCodigo.Text
    rs.Fields("VALOR") = txtValor.Text
    rs.Fields("REFERENCIA") = txtReferencia.Text
    
    rs.Update
    rs.Close
    Set rs = Nothing

    Call carregaGridRecibos(Return_Recordset(Return_SqlString(4)))

    MsgBox "Recibo gravado", vbInformation, "Informação"
    

End Sub

Private Sub carregaGridRecibos(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    'Dim bkColor As Long
    Dim col As Integer
    
    dblValorVerde = 0
    
    selectedRow = 2
    With gridRecibos
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            If (IsNull(code)) Or (code = "") Then
                .AddItem (rs.Fields("codigo") & vbTab & rs.Fields("data") & vbTab & rs.Fields("cliente") & vbTab & rs.Fields("valor"))
            Else
                .AddItem (rs.Fields("codigo") & vbTab & rs.Fields("data") & vbTab & rs.Fields("nome") & vbTab & rs.Fields("valor"))
            End If
            
            If code = rs.Fields("CODIGO") Then
                selectedRow = .Rows - 1
            End If
            
                .row = .Rows - 1
                For col = .FixedCols To .Cols - 1
                    .col = col
                    
                    'If rs.Fields("Status") = "Aprovado" Then
                    '    .CellBackColor = vbGreen
                    'End If
                    
                    'If rs.Fields("Status") = "Reprovado" Then
                    '    .CellBackColor = vbRed
                    'End If
                    
                    'If rs.Fields("Status") = "Enviado" Then
                    '    .CellBackColor = vbYellow
                    'End If
                Next
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

Private Sub carregaInfoRecibo(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    If gridRecibos.TextMatrix(gridRecibos.RowSel, 0) = "" Then
        Exit Sub
    End If
    
    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        lblCodRecibo.Caption = rs.Fields("codigo")
        txtData.Text = rs.Fields("data")
        txtValor.Text = rs.Fields("valor")
        txtClienteCodigo.Text = rs.Fields("cod_cliente")
        txtClienteNome.Text = rs.Fields("nome")
        txtReferencia.Text = rs.Fields("referencia")
        
        If Not IsNull(rs.Fields("cpf")) Then
            lblCPF.Caption = "CPF: " & rs.Fields("cpf")
        Else
            lblCPF.Caption = "CPF: Não informado no cadastro do cliente"
        End If
    End If
        
    rs.Close
    Set rs = Nothing
    
    
End Sub

Public Function Extenso(nvalor)
    'Valida Argumento
    If IsNull(nvalor) Or nvalor <= 0 Or nvalor > 9999999.99 Then
    Exit Function
    End If
    
    'Variáveis
    Dim nContador, nTamanho As Integer
    Dim cValor, cParte, cFinal As String
    ReDim aGrupo(4), aTexto(4) As String
    
    'Matrizes de extensos (Parciais)
    ReDim aUnid(19) As String
    aUnid(1) = "um ": aUnid(2) = "dois ": aUnid(3) = "tres "
    aUnid(4) = "quatro ": aUnid(5) = "cinco ": aUnid(6) = "seis "
    aUnid(7) = "sete ": aUnid(8) = "oito ": aUnid(9) = "nove "
    aUnid(10) = "dez ": aUnid(11) = "onze ": aUnid(12) = "doze "
    aUnid(13) = "treze ": aUnid(14) = "quatorze ": aUnid(15) = "quinze "
    aUnid(16) = "dezesseis ": aUnid(17) = "dezessete ": aUnid(18) = "dezoito "
    aUnid(19) = "dezenove "
    
    ReDim aDezena(9) As String
    aDezena(1) = "dez ": aDezena(2) = "vinte ": aDezena(3) = "trinta "
    aDezena(4) = "quarenta ": aDezena(5) = "cinquenta "
    aDezena(6) = "sessenta ": aDezena(7) = "setenta ": aDezena(8) = "oitenta "
    aDezena(9) = "noventa "
    
    ReDim aCentena(9) As String
    aCentena(1) = "cento ": aCentena(2) = "duzentos "
    aCentena(3) = "trezentos ": aCentena(4) = "quatrocentos "
    aCentena(5) = "quinhentos ": aCentena(6) = "seiscentos "
    aCentena(7) = "setecentos ": aCentena(8) = "oitocentos "
    aCentena(9) = "novecentos "
    
    'Separa valor em grupos
    cValor = Format$(nvalor, "0000000000.00")
    aGrupo(1) = Mid$(cValor, 2, 3)
    aGrupo(2) = Mid$(cValor, 5, 3)
    aGrupo(3) = Mid$(cValor, 8, 3)
    aGrupo(4) = "0" + Mid$(cValor, 12, 2)
    
    'Calcula cada grupo
    For nContador = 1 To 4
      cParte = aGrupo(nContador)
      nTamanho = Switch(val(cParte) < 10, 1, val(cParte) < 100, 2, val(cParte) < 1000, 3)
      If nTamanho = 3 Then
        If Right$(cParte, 2) <> "00" Then
          aTexto(nContador) = aTexto(nContador) + aCentena(Left(cParte, 1)) + "e "
          nTamanho = 2
        Else
          aTexto(nContador) = aTexto(nContador) + IIf(Left$(cParte, 1) = "1", "cem ", aCentena(Left(cParte, 1)))
        End If
      End If
      If nTamanho = 2 Then
        If val(Right(cParte, 2)) < 20 Then
          aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 2))
        Else
          aTexto(nContador) = aTexto(nContador) + aDezena(Mid(cParte, 2, 1))
          If Right$(cParte, 1) <> "0" Then
            aTexto(nContador) = aTexto(nContador) + "e "
            nTamanho = 1
          End If
        End If
      End If
      If nTamanho = 1 Then
        aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 1))
      End If
    Next
    
    'Final
    If val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And val(aGrupo(4)) <> 0 Then
      cFinal = aTexto(4) + IIf(val(aGrupo(4)) = 1, "centavo", "centavos")
    Else
      cFinal = ""
      cFinal = cFinal + IIf(val(aGrupo(1)) <> 0, aTexto(1) + IIf(val(aGrupo(1)) > 1, "milhões ", "milhão "), "")
      If val(aGrupo(2) + aGrupo(3)) = 0 Then
        cFinal = cFinal + "de "
      Else
        cFinal = cFinal + IIf(val(aGrupo(2)) <> 0, aTexto(2) + "mil ", "")
      End If
      cFinal = cFinal + aTexto(3) + IIf(val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1, "real ", "reais ")
      cFinal = cFinal + IIf(val(aGrupo(4)) <> 0, "E " + aTexto(4) + IIf(val(aGrupo(4)) = 1, "centavo", "centavos"), "")
    End If
    Extenso = UCase$(cFinal)
End Function

Function nomeMes(dData) As String
    
    If Month(dData) = 1 Then
        nomeMes = "janeiro"
    ElseIf Month(dData) = 2 Then
        nomeMes = "fevereiro"
    ElseIf Month(dData) = 3 Then
        nomeMes = "março"
    ElseIf Month(dData) = 4 Then
        nomeMes = "abril"
    ElseIf Month(dData) = 5 Then
        nomeMes = "maio"
    ElseIf Month(dData) = 6 Then
        nomeMes = "junho"
    ElseIf Month(dData) = 7 Then
        nomeMes = "julho"
    ElseIf Month(dData) = 8 Then
        nomeMes = "agosto"
    ElseIf Month(dData) = 9 Then
        nomeMes = "setembro"
    ElseIf Month(dData) = 10 Then
        nomeMes = "outubro"
    ElseIf Month(dData) = 11 Then
        nomeMes = "novembro"
    ElseIf Month(dData) = 12 Then
        nomeMes = "dezembro"
    End If
    
    
    
End Function
