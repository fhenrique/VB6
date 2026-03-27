VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFerramentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XSOL Aquecedores"
   ClientHeight    =   10395
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   20295
   Icon            =   "frmFerramentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   20295
   Begin VB.Frame Frame4 
      Caption         =   "Histórico de empréstimos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9570
      Left            =   10260
      TabIndex        =   11
      Top             =   735
      Width           =   9930
      Begin MSFlexGridLib.MSFlexGrid gridHistorico 
         Height          =   9150
         Left            =   120
         TabIndex        =   17
         Top             =   270
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   16140
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Lista de ferramentas cadastradas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6315
      Left            =   75
      TabIndex        =   3
      Top             =   735
      Width           =   10005
      Begin MSFlexGridLib.MSFlexGrid gridFerramentas 
         Height          =   5955
         Left            =   90
         TabIndex        =   16
         Top             =   270
         Width           =   9840
         _ExtentX        =   17357
         _ExtentY        =   10504
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados da ferramenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3150
      Left            =   90
      TabIndex        =   2
      Top             =   7170
      Width           =   10035
      Begin MSComCtl2.DTPicker dtpDataAquisicao 
         Height          =   330
         Left            =   1560
         TabIndex        =   19
         Top             =   870
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   104136705
         CurrentDate     =   45930
      End
      Begin VB.CommandButton btnEmprestimo 
         BackColor       =   &H0080FFFF&
         Caption         =   "&Empréstimo"
         Height          =   360
         Left            =   8565
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2565
         Width           =   1140
      End
      Begin VB.CommandButton btnExcluir 
         BackColor       =   &H008080FF&
         Caption         =   "E&xcluir"
         Height          =   360
         Left            =   4260
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2565
         Width           =   1100
      End
      Begin VB.CommandButton btnEditar 
         Caption         =   "&Editar"
         Height          =   360
         Left            =   1620
         TabIndex        =   14
         Top             =   2565
         Width           =   1100
      End
      Begin VB.CommandButton btnSalvar 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   360
         Left            =   2940
         TabIndex        =   13
         Top             =   2565
         Width           =   1100
      End
      Begin VB.CommandButton btnIncluir 
         BackColor       =   &H0080FF80&
         Caption         =   "&Incluir"
         Height          =   360
         Left            =   315
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2565
         Width           =   1100
      End
      Begin VB.TextBox txtVendedor 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1815
         Width           =   7635
      End
      Begin VB.TextBox txtComprador 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1335
         Width           =   7635
      End
      Begin VB.TextBox txtDescricao 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   70
         TabIndex        =   5
         Top             =   420
         Width           =   7635
      End
      Begin VB.Label lblCodFerramenta 
         Caption         =   "codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   1905
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comprador:"
         Height          =   195
         Left            =   645
         TabIndex        =   7
         Top             =   1425
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data de aquisição:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   990
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   690
         TabIndex        =   4
         Top             =   525
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   20190
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ferramentas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8985
         TabIndex        =   1
         Top             =   195
         Width           =   2235
      End
   End
End
Attribute VB_Name = "frmFerramentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEditando As Boolean


Private Sub btnEditar_Click()

    bEditando = True

    If gridFerramentas.Rows = 2 Then
        MsgBox "Não há ferramenta(s) há editar!"
        Exit Sub
    End If

    If RTrim(LTrim(txtDescricao.Text)) = "" Then
        MsgBox "Selecione uma ferramenta na Lista de ferramentas cadastradas!"
        Exit Sub
    End If

    bEditando = True

    btnIncluir.enabled = False
    btnEditar.enabled = False
    btnSalvar.enabled = True
    
    btnExcluir.Caption = "Cancelar"
    btnExcluir.enabled = True
    
    btnEmprestimo.enabled = False
    
    gridFerramentas.enabled = False
    gridHistorico.enabled = False
     
    habilitaCampos (True)
    



End Sub

Private Sub btnEmprestimo_Click()
    
    If gridFerramentas.Rows = 2 Then
        MsgBox "É necessário cadastrar ao menos 1 ferramenta!"
        Exit Sub
    End If
    
    
    frmEmprestimo.lblFerramenta.Caption = txtDescricao.Text
    frmEmprestimo.Show (1)
    
End Sub

Private Sub Cmd_Inclui_Click()

    

End Sub

Private Sub btnExcluir_Click()
    
    
    If btnExcluir.Caption = "Excluir" Then
        If RTrim(LTrim(txtDescricao.Text)) = "" Then
            MsgBox "É necessário selecionar uma ferramenta!"
            Exit Sub
            End If

        'faz a exclusão
        
        btnIncluir.enabled = True
        btnEditar.enabled = True
        btnSalvar.enabled = False
        
        btnExcluir.enabled = False
        
        btnEmprestimo.enabled = True
    Else
        btnIncluir.enabled = True
        btnEditar.enabled = True
        btnSalvar.enabled = False
        
        btnExcluir.Caption = "Excluir"
        btnExcluir.enabled = True
        btnEmprestimo.enabled = True
    End If
    
    
End Sub

Private Sub btnIncluir_Click()
    bEditando = False
    
    btnIncluir.enabled = False
    btnEditar.enabled = False
    btnSalvar.enabled = True
    
    btnExcluir.Caption = "Cancelar"
    btnExcluir.enabled = True
    
    btnEmprestimo.enabled = False
    
    habilitaCampos (True)
    
    txtDescricao.SetFocus
    
End Sub

Private Sub btnSalvar_Click()
    btnIncluir.enabled = True
    btnEditar.enabled = True
    btnSalvar.enabled = False
    
    btnExcluir.Caption = "Excluir"
    btnExcluir.enabled = True
    
    btnEmprestimo.enabled = True
    
'-----------
    
    habilitaCampos (False)
 
    If RTrim(LTrim(txtDescricao.Text)) = "" Then
        MsgBox "Informe o nome da ferramenta!", vbInformation, "Validação"
        txtDescricao.SetFocus
        Exit Sub
    End If
    
    Call gravaFerramenta
    
    telaVazia
    
    gridFerramentas.enabled = True
    gridHistorico.enabled = True
    
    'gridFerramentas_sellchange
    
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption + " - " + versaoAtual
       
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
    
    'Call carregaCboClientes(Return_Recordset(Return_SqlString(1)))
    'Call carregaCboClientesProspects(Return_Recordset(Return_SqlString(1)))
    
    Call ConfigureGridFerramentas
    Call ConfigureGridHistorico
    
    'Call carregaFerramentas(Return_Recordset(Return_SqlString(10)))
    
    'Call HabilitarComponentes(False)
End Sub

Private Sub ConfigureGridHistorico()
    With gridHistorico
        .Cols = 7
        .ColWidth(0) = 0 'Codigo (invisível")
        .ColWidth(1) = 0 'codigo ferramenta (invisivel)
        .ColWidth(2) = 1200 'Dt. emprestimo
        .ColWidth(3) = 2300 'Beneficiário(a)
        .ColWidth(4) = 3800 'Finalidade
        .ColWidth(5) = 1300 'Dt. devolucao
        .ColWidth(6) = 1000 'Devolvido?


        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter

        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Cód. Ferramenta"
        .TextMatrix(0, 2) = "Dt. Empréstimo"
        .TextMatrix(0, 3) = "Beneficiário"
        .TextMatrix(0, 4) = "Finalidade"
        .TextMatrix(0, 5) = "Dt. Devolução"
        .TextMatrix(0, 6) = "Devolvido?"
        
        .RowHeight(1) = 0
    End With
End Sub

Private Sub ConfigureGridFerramentas()
    With gridFerramentas
        .Cols = 5
        .ColWidth(0) = 0 'Codigo Invisível
        .ColWidth(1) = 3900 'descricao
        .ColWidth(2) = 1150 'dt. aquisicao
        .ColWidth(3) = 2300  'Comprador
        .ColWidth(4) = 2300 'Vendedor
        '.ColWidth(5) = 900 'EmUso

        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        '.ColAlignment(5) = flexAlignLeftCenter

        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Descrição"
        .TextMatrix(0, 2) = "Dt. Aquisição"
        .TextMatrix(0, 3) = "Comprador"
        .TextMatrix(0, 4) = "Vendedor"
        '.TextMatrix(0, 5) = "Em uso?"
        
        .RowHeight(1) = 0
    End With
End Sub

Private Sub limpaTela()
    
    txtDescricao.Text = ""
    dtpDataAquisicao.value = Date
    txtComprador.Text = ""
    txtVendedor.Text = ""
    
    txtDescricao.SetFocus
End Sub


Private Sub habilitaCampos(habilitar As Boolean)
    
    txtDescricao.Locked = Not (habilitar)
    dtpDataAquisicao.enabled = habilitar
    txtComprador.Locked = Not (habilitar)
    txtVendedor.Locked = Not (habilitar)
    
End Sub

Private Sub gravaFerramenta()
    Dim rs As ADODB.Recordset
    
    If bEditando = False Then 'está inserindo
        Set rs = Return_Recordset(Return_SqlString(1))
        rs.AddNew
    Else 'está editando
        Set rs = Return_Recordset(Return_SqlString(2, lblCodFerramenta.Caption))
    End If
    
    'CODIGO, DATA, COD_CLIENTE, VALOR
    'rs.Fields("CODIGO") = lblCodRecibo.Caption
    rs.Fields("DESCRICAO") = txtDescricao.Text
    rs.Fields("DATA_AQUISICAO") = dtpDataAquisicao.value
    rs.Fields("COMPRADOR") = txtComprador.Text
    rs.Fields("VENDEDOR") = txtVendedor.Text
    
    rs.Update
    rs.Close
    Set rs = Nothing

    Call carregaGridFerramentas(Return_Recordset(Return_SqlString(1)))

    MsgBox "Cadastro salvo!", vbInformation, "Informação"
    

End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String, Optional param3 As String) As String
    If queryType = 1 Then
        Return_SqlString = "Select * from FERRAMENTAS"
    ElseIf queryType = 2 Then
        Return_SqlString = "Select * from FERRAMENTAS where codigo = " & param1
    End If
End Function

Private Sub telaVazia()
    
    txtDescricao.Text = ""
    dtpDataAquisicao.value = Date
    txtComprador.Text = ""
    txtVendedor.Text = ""
    
End Sub

Private Sub gridFerramentas_SelChange()
    If gridFerramentas.Rows > 2 Then
    'If gridRecibos.TextMatrix(2, 0) <> "" And gridOrcamentos.TextMatrix(gridOrcamentos.RowSel, 0) <> "" Then
        Call carregaInfoFerramenta(Return_Recordset(Return_SqlString(2, gridFerramentas.TextMatrix(gridFerramentas.RowSel, 0))))
    'End If
    End If

End Sub

Private Sub carregaInfoFerramenta(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    'If gridFerramentas.TextMatrix(gridFerramentas.RowSel, 0) = "" Then
    '    Exit Sub
    'End If
    
    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        lblCodFerramenta.Caption = rs.Fields("CODIGO")
        txtDescricao.Text = rs.Fields("DESCRICAO")
        dtpDataAquisicao.value = rs.Fields("DATA_AQUISICAO")
        txtComprador.Text = rs.Fields("COMPRADOR")
        txtVendedor.Text = rs.Fields("VENDEDOR")
    End If
        
    rs.Close
    Set rs = Nothing
    
    
End Sub

Private Sub carregaGridFerramentas(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    
    'Dim bkColor As Long
    Dim col As Integer
    
    Dim bDevolucao As Boolean
    Dim strDevolvido As String
    
    
    
    dblValorVerde = 0
    
    selectedRow = 2
    With gridFerramentas
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            'If (IsNull(code)) Or (code = "") Then
                '.AddItem (rs.Fields("codigo") & vbTab & rs.Fields("data") & vbTab & rs.Fields("cliente") & vbTab & rs.Fields("valor"))
            'Else
                
                'If rs.Fields("EMPRESTIMO_DEVOLVIDO") = "0" Then
                '    bDevolucao = False
                'Else
                '    bDevolucao = True
                'End If
                
                
                'If bDevolucao = False Then
                '    strDevolvido = "N"
                'Else
                '    strDevolvido = "S"
                'End If
                
                .AddItem (rs.Fields("CODIGO") & vbTab & rs.Fields("DESCRICAO") & vbTab & rs.Fields("DATA_AQUISICAO") & vbTab & rs.Fields("COMPRADOR") & vbTab & rs.Fields("VENDEDOR"))
            'End If
            
            'If code = rs.Fields("CODIGO") Then
            '    selectedRow = .Rows - 1
            'End If
            
                .row = .Rows - 1
                For col = .FixedCols To .Cols - 1
                    .col = col
                    
                    'If rs.Fields("EMPRESTIMO_DEVOLVIDO") = "1" Then
                    '    .CellBackColor = vbGreen
                    'Else
                    '    .CellBackColor = vbRed
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
