VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_Equipment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Equipamentos"
   ClientHeight    =   4020
   ClientLeft      =   3075
   ClientTop       =   3075
   ClientWidth     =   15780
   Icon            =   "Form_Equipment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   15780
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   30
      TabIndex        =   18
      Top             =   2175
      Width           =   6585
      Begin VB.TextBox txtTipoTelhado 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1545
         TabIndex        =   19
         Top             =   240
         Width           =   4545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo de telhado:"
         Height          =   195
         Left            =   285
         TabIndex        =   20
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Lista de Equipamentos"
      Height          =   3990
      Left            =   6720
      TabIndex        =   13
      Top             =   0
      Width           =   9015
      Begin MSFlexGridLib.MSFlexGrid Grid_Equipment 
         Height          =   3690
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   6509
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Edição"
      Height          =   1095
      Left            =   0
      TabIndex        =   8
      Top             =   2895
      Width           =   6615
      Begin VB.CommandButton Cmd_Exclui 
         BackColor       =   &H008080FF&
         Caption         =   "E&xcluir"
         Height          =   495
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Edita 
         Caption         =   "&Editar"
         Height          =   495
         Left            =   1920
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Salva 
         Caption         =   "&Salvar"
         Height          =   495
         Left            =   3480
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Inclui 
         BackColor       =   &H0080FF80&
         Caption         =   "&Incluir"
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Equipamentos"
      Height          =   2145
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox chkFicha 
         Caption         =   "Imprimir na Ficha de Assistência Técnica"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2985
         TabIndex        =   17
         Top             =   1755
         Width           =   3255
      End
      Begin VB.TextBox txtCodEnderecos 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1575
         MaxLength       =   200
         TabIndex        =   16
         Top             =   1365
         Width           =   4575
      End
      Begin VB.TextBox Text_Equip 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   69
         TabIndex        =   4
         Top             =   1005
         Width           =   4590
      End
      Begin VB.TextBox Text_Cod 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1575
         TabIndex        =   3
         Top             =   270
         Width           =   855
      End
      Begin VB.TextBox Text_Nome_Cliente 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   2
         Top             =   630
         Width           =   3615
      End
      Begin VB.TextBox Text_Cod_Cliente 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1575
         TabIndex        =   1
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cod. Endereço(s):"
         Height          =   255
         Left            =   90
         TabIndex        =   15
         Top             =   1410
         Width           =   1380
      End
      Begin VB.Label Lbl_End 
         Alignment       =   1  'Right Justify
         Caption         =   "Equipamento:"
         Height          =   255
         Left            =   495
         TabIndex        =   7
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Lbl_Registro 
         Alignment       =   1  'Right Justify
         Caption         =   "Código:"
         Height          =   255
         Left            =   855
         TabIndex        =   6
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Lbl_Cliente 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   495
         TabIndex        =   5
         Top             =   630
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form_Equipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conn As ADODB.Connection

Private Sub Form_Activate()
txtCodEnderecos.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Equip.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Cod.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Nome_Cliente.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Cod_Cliente.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Grid_Equipment.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
End Sub

Private Sub Form_Load()

    

   Text_Cod = Cod_Equipment
   Text_Cod_Cliente.Text = Cod_Cli
   Text_Nome_Cliente.Text = Nome_Cli_Atend
   
   Call Object_Controller(1)
   Call Configure_Grid
   
   'cloud fazendo
   Call carregaGridEquipamentos(Return_Recordset(Return_SqlString(1, str(Cod_Cli))))
   
   
   
End Sub

Private Sub carregaGridEquipamentos(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim col As Integer
    
    If cloud_status_conexao = 0 Then
        Exit Sub
    End If
    
    selectedRow = 2
    With Grid_Equipment
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("CODIGO") & vbTab & rs.Fields("equipamento") & vbTab & rs.Fields("cod_endereco"))
            If code = rs.Fields("CODIGO") Then
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
    
    strStatusOrcamento = ""
    strStatusDescricao = ""

    
End Sub



Private Sub Cmd_Edita_Click()
    If Cmd_Edita.Caption = "&Editar" Then
        Call Object_Controller(4)
    Else
        Call Object_Controller(5)
    End If
End Sub

Private Sub Cmd_Inclui_Click()
    If Cmd_Inclui.Caption = "&Incluir" Then
        Call Object_Controller(3)
    Else
        Call Clean_Boxes
        Call Object_Controller(5)
    End If
End Sub

Private Sub Cmd_Salva_Click()
    Dim imprimeNaFicha As String
    
    If Not Mandatory_Fields_Filled Then
        Exit Sub
    End If
    
    If chkFicha.value = 0 Then
        imprimeNaFicha = "N"
    Else
        imprimeNaFicha = "S"
    End If
    
    
    If Cmd_Inclui.Caption = "&Cancela" Then
        'insere
   
        'parou aqui o cloud fazendo
        Call cloud_Insert("Equipamento", "Cod_Cliente,Cod_Endereco,Equipamento,ExibeNaFicha", Text_Cod_Cliente.Text & textoComVirgulaEsquerda(txtCodEnderecos.Text) & textoComVirgulaEsquerda(Text_Equip.Text) & textoComVirgulaEsquerda(imprimeNaFicha))
    Else
        'edita
        Call cloud_Edit("Equipamento", "Cod_Endereco,Equipamento,ExibeNaFicha", txtCodEnderecos.Text & textoComVirgulaEsquerda(Text_Equip.Text) & textoComVirgulaEsquerda(imprimeNaFicha) & ".", "Codigo=" & Text_Cod)
    End If
    
    Call Object_Controller(5)
    Call carregaGridEquipamentos(Return_Recordset(Return_SqlString(1, str(Cod_Cli))))
    Form_Customer.Refresh_Grid_Equipment
End Sub

Private Sub Cmd_Exclui_Click()
    If MsgBox("Confirma a Exclusão do Equipamento " & Text_Equip.Text & " do Cliente " & Text_Nome_Cliente.Text & " ? ", vbYesNo, "Excluir Registros") = vbYes Then
        If Have_Reference() Then
            MsgBox "Não é possível sua remoção pois há uma OS que referencia a este endereço", vbInformation, "Informação"
            Exit Sub
        End If
        Call Return_Recordset(Return_SqlString(2, Text_Cod.Text)).Delete
        Call Object_Controller(1)
        Call Show_Customer_Equipment(Return_Recordset(Return_SqlString(1, Text_Cod_Cliente.Text)))
        Form_Customer.Refresh_Grid_Equipment
    End If
End Sub

Private Function Have_Reference() As Boolean
    Dim rs As ADODB.Recordset

    Set rs = Return_Recordset(Return_SqlString(4, Text_Cod.Text))
    
    If rs.RecordCount > 0 Then
        Have_Reference = True
    Else
        Have_Reference = False
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub Record_Record(rs As ADODB.Recordset)
    rs.Fields("Equipamento").value = UCase(Text_Equip.Text)
    rs.Fields("Cod_Endereco").value = txtCodEnderecos.Text
    rs.Fields("Tipo_Telhado").value = txtTipoTelhado.Text
   
    If chkFicha.value = 1 Then
        rs.Fields("ExibeNaFicha") = "S"
    Else
        rs.Fields("ExibeNaFicha") = "N"
    End If
End Sub

Sub Clean_Boxes()
    Text_Cod.Text = ""
    Text_Equip.Text = ""
End Sub

Private Function Mandatory_Fields_Filled() As Boolean
    
    Mandatory_Fields_Filled = True

    If Trim(Text_Equip.Text) = "" Then
       MsgBox "É necessário digitar o Equipamento! ", vbOKOnly, "XSol Aquecedores"
       Text_Equip.SetFocus
       Mandatory_Fields_Filled = False
       Exit Function
    End If
    
    If Trim(txtCodEnderecos.Text) = "" Then
       MsgBox "É necessário informar o código do endereço! ", vbOKOnly, "XSol Aquecedores"
       txtCodEnderecos.SetFocus
       Mandatory_Fields_Filled = False
       Exit Function
    End If
    
End Function

Private Sub Object_Controller(cType As Integer)
    If cType = 1 Then
        Cmd_Inclui.enabled = True
        Cmd_Salva.enabled = False
        Cmd_Edita.enabled = False
        Cmd_Exclui.enabled = False
        Grid_Equipment.enabled = True
        Call Clean_Boxes
    ElseIf cType = 2 Then  'record selected
        Cmd_Edita.enabled = True
        Cmd_Exclui.enabled = True
        Grid_Equipment.enabled = True
    ElseIf cType = 3 Then  'insert a record
        Cmd_Inclui.Caption = "&Cancela"
        Text_Equip.enabled = True
        txtTipoTelhado.enabled = True
        Text_Equip.BackColor = &HFFFFFF
        txtTipoTelhado.BackColor = &HFFFFFF
        txtCodEnderecos.Text = ""
        txtCodEnderecos.enabled = True
        txtCodEnderecos.BackColor = &HFFFFFF
        Cmd_Edita.enabled = False
        chkFicha.enabled = True
        Call insEdit_Controller
        Call Clean_Boxes
    ElseIf cType = 4 Then  'Edit a record
        Text_Equip.enabled = True
        txtTipoTelhado.enabled = True
        
        Text_Equip.BackColor = &HFFFFFF
        txtTipoTelhado.BackColor = &HFFFFFF
        
        txtCodEnderecos.enabled = True
        txtCodEnderecos.BackColor = &HFFFFFF
        chkFicha.enabled = True
        Cmd_Inclui.enabled = False
        Cmd_Edita.Caption = "&Cancela"
        Call insEdit_Controller
    ElseIf cType = 5 Then  'Cancel a insertion or edition or after saving
        Cmd_Inclui.Caption = "&Incluir"
        Cmd_Inclui.enabled = True
        Cmd_Salva.enabled = False
        Cmd_Edita.Caption = "&Editar"
        Cmd_Edita.enabled = True
        Cmd_Exclui.enabled = True
        Grid_Equipment.enabled = True
        
        Text_Equip.BackColor = &H8000000F
        txtTipoTelhado.BackColor = &H8000000F
        Text_Equip.enabled = False
        txtTipoTelhado.enabled = False
        
        txtCodEnderecos.BackColor = &H8000000F
        txtCodEnderecos.enabled = False
        chkFicha.enabled = False
        
        Call Grid_Equipment_Click
    End If
End Sub

Private Sub insEdit_Controller()
    Cmd_Salva.enabled = True
    Cmd_Exclui.enabled = False
    Grid_Equipment.enabled = False
    
    Text_Equip.BackColor = vbYellow
    Text_Equip.SetFocus
End Sub

Private Sub Configure_Grid()
    With Grid_Equipment
        .Cols = 3
        .ColWidth(0) = 630
        .ColWidth(1) = 7280
        .ColWidth(2) = 800
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Equipamento"
        .TextMatrix(0, 2) = "Cód. End."
        .RowHeight(1) = 0
    End With
End Sub
'
'Private Function Return_Recordset(strSql As String) As ADODB.Recordset
'    Dim rs As New ADODB.Recordset
'
'    rs.Open strSql, conn, adOpenStatic, adLockOptimistic
'
'    Set Return_Recordset = rs
'End Function

Private Sub Show_Customer_Equipment(rs As ADODB.Recordset, Optional code As Integer)
    Dim selectedRow As Integer
    
    selectedRow = 2
    
    'MsgBox rs.RecordCount - (rs.RecordCount / 2)
    
    With Grid_Equipment
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("codigo") & vbTab & IIf(IsNull(rs.Fields("Equipamento")), "", rs.Fields("Equipamento")) & vbTab & IIf(IsNull(rs.Fields("Cod_Endereco")), "", rs.Fields("Cod_Endereco")))
            If code = rs.Fields("codigo") Then
                selectedRow = .Rows - 1
            End If
            rs.MoveNext
        Loop
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            '.TopRow = selectedRow
            .col = 0
            .ColSel = .Cols - 1
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Grid_Equipment_Click()
    With Grid_Equipment
        If .RowSel = 1 Then
            Exit Sub
        End If
        Call Show_Record(Return_Recordset(Return_SqlString(7, .TextMatrix(.RowSel, 0))))
    End With
End Sub

Private Sub Grid_Equipment_SelChange()
    Call Grid_Equipment_Click
End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String, Optional param3 As String) As String
    Dim whereClause As String
    
    
    If queryType = 1 Then
        Return_SqlString = "SELECT distinct(a.codigo), a.cod_cliente, a.equipamento, a.cod_endereco, a.exibenaficha FROM `Equipamento` a Where a.cod_cliente = " & param1 & " ORDER BY a.Codigo"
    ElseIf queryType = 2 Then
        Return_SqlString = "SELECT a.*, b.tipo_telhado as tipo_telhado FROM `Equipamento` a INNER JOIN `Endereco` b on b.cod_cliente = a.cod_cliente WHERE a.cod_cliente = " & Text_Cod_Cliente.Text & " ORDER BY a.Codigo"
    ElseIf queryType = 3 Then
        Return_SqlString = "SELECT a.*, b.tipo_telhado FROM `Equipamento` a INNER JOIN `Endereco` b on b.cod_cliente = a.cod_cliente ORDER BY a.Codigo"
    ElseIf queryType = 4 Then
        Return_SqlString = "SELECT * FROM `Atendimento` WHERE Cod_Equip = " & param1
    ElseIf queryType = 5 Then
        Return_SqlString = "SELECT * FROM `Estoque` "
    ElseIf queryType = 6 Then
        Return_SqlString = "SELECT * FROM `Endereco` WHERE cod_cliente = " & param1
    ElseIf queryType = 7 Then
        Return_SqlString = "SELECT a.*, b.tipo_telhado FROM `Equipamento` a INNER JOIN `Endereco` b on b.cod_cliente = a.cod_cliente Where a.codigo = " & param1
    End If
    
End Function

Sub Show_Record(rs As ADODB.Recordset)

    If cloud_status_conexao = 0 Then
        Exit Sub
    End If
    

    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    Text_Cod.Text = rs.Fields("Codigo")
    Text_Equip.Text = UCase(IIf(IsNull(rs.Fields("Equipamento")), "", rs.Fields("Equipamento")))
    txtCodEnderecos.Text = UCase(IIf(IsNull(rs.Fields("Cod_Endereco")), "", rs.Fields("Cod_Endereco")))
    txtTipoTelhado.Text = UCase(IIf(IsNull(rs.Fields("tipo_telhado")), "", rs.Fields("tipo_telhado")))
    
    If rs.Fields("ExibeNaFicha") = "S" Then
        chkFicha.value = 1
    Else
        chkFicha.value = 0
    End If
        
    rs.Close
    Set rs = Nothing
    
    Call Object_Controller(2)
End Sub

Private Sub Text_Equip_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 180 Then
        KeyAscii = 0
    End If
End Sub
