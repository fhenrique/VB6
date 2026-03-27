VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_Address 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Endereços"
   ClientHeight    =   5880
   ClientLeft      =   2760
   ClientTop       =   3255
   ClientWidth     =   15930
   Icon            =   "Form_Address.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   15930
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      Caption         =   "Lista de Endereços"
      Height          =   5685
      Left            =   6840
      TabIndex        =   25
      Top             =   105
      Width           =   9015
      Begin MSFlexGridLib.MSFlexGrid Grid_Address 
         Height          =   5355
         Left            =   135
         TabIndex        =   26
         Top             =   225
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   9446
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
      Left            =   120
      TabIndex        =   20
      Top             =   4680
      Width           =   6615
      Begin VB.CommandButton Cmd_Inclui 
         BackColor       =   &H0080FF80&
         Caption         =   "&Incluir"
         Height          =   495
         Left            =   375
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Salva 
         Caption         =   "&Salvar"
         Height          =   495
         Left            =   3480
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Edita 
         Caption         =   "&Editar"
         Height          =   495
         Left            =   1920
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Exclui 
         BackColor       =   &H008080FF&
         Caption         =   "E&xcluir"
         Height          =   495
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Endereços"
      Height          =   4545
      Left            =   135
      TabIndex        =   0
      Top             =   105
      Width           =   6615
      Begin VB.TextBox txtTelCelular 
         Height          =   285
         Left            =   1425
         TabIndex        =   32
         Top             =   4110
         Width           =   2745
      End
      Begin VB.TextBox txtTelCom 
         Height          =   285
         Left            =   1425
         TabIndex        =   31
         Top             =   3690
         Width           =   2745
      End
      Begin VB.TextBox txtTelResidencial 
         Height          =   285
         Left            =   1425
         TabIndex        =   30
         Top             =   3270
         Width           =   2745
      End
      Begin VB.TextBox Text_Cod_Cliente 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1425
         TabIndex        =   18
         Top             =   660
         Width           =   855
      End
      Begin VB.TextBox Text_Nome_Cliente 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2385
         TabIndex        =   17
         Top             =   660
         Width           =   3615
      End
      Begin VB.TextBox Text_Cod 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1425
         TabIndex        =   15
         Top             =   285
         Width           =   855
      End
      Begin VB.TextBox Text_Ref 
         Height          =   285
         Left            =   1425
         TabIndex        =   6
         Top             =   2460
         Width           =   4590
      End
      Begin VB.TextBox Text_Cep 
         Height          =   285
         Left            =   2625
         TabIndex        =   5
         Top             =   2100
         Width           =   1455
      End
      Begin VB.TextBox Text_Cidade 
         Height          =   285
         Left            =   1425
         TabIndex        =   4
         Top             =   1740
         Width           =   4575
      End
      Begin VB.TextBox Text_Bairro 
         Height          =   285
         Left            =   1425
         TabIndex        =   3
         Top             =   1380
         Width           =   4575
      End
      Begin VB.TextBox Text_End 
         Height          =   285
         Left            =   1425
         TabIndex        =   2
         Top             =   1020
         Width           =   4575
      End
      Begin VB.ComboBox Combo_Uf 
         Height          =   315
         Left            =   1425
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2100
         Width           =   615
      End
      Begin VB.ComboBox Combo_Zona 
         Height          =   315
         Left            =   1425
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2820
         Width           =   2775
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tel. Celular:"
         Height          =   195
         Left            =   480
         TabIndex        =   29
         Top             =   4155
         Width           =   840
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tel. Com.:"
         Height          =   195
         Left            =   600
         TabIndex        =   28
         Top             =   3750
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tel. Resid.:"
         Height          =   195
         Left            =   510
         TabIndex        =   27
         Top             =   3345
         Width           =   810
      End
      Begin VB.Label Lbl_Cliente 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   795
         TabIndex        =   19
         Top             =   720
         Width           =   525
      End
      Begin VB.Label Lbl_Registro 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   780
         TabIndex        =   16
         Top             =   345
         Width           =   540
      End
      Begin VB.Label Lbl_Zona 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Zona:"
         Height          =   195
         Left            =   900
         TabIndex        =   14
         Top             =   2895
         Width           =   420
      End
      Begin VB.Label Lbl_Ref 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Referência:"
         Height          =   195
         Left            =   495
         TabIndex        =   13
         Top             =   2520
         Width           =   825
      End
      Begin VB.Label Lbl_Cep 
         Alignment       =   1  'Right Justify
         Caption         =   "Cep:"
         Height          =   255
         Left            =   2145
         TabIndex        =   12
         Top             =   2100
         Width           =   375
      End
      Begin VB.Label Lbl_Uf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Uf:"
         Height          =   195
         Left            =   1110
         TabIndex        =   11
         Top             =   2145
         Width           =   210
      End
      Begin VB.Label Lbl_Cidade 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   780
         TabIndex        =   10
         Top             =   1770
         Width           =   540
      End
      Begin VB.Label Lbl_Bairro 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         Height          =   195
         Left            =   870
         TabIndex        =   9
         Top             =   1425
         Width           =   450
      End
      Begin VB.Label Lbl_End 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   585
         TabIndex        =   7
         Top             =   1065
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form_Address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conn As ADODB.Connection
Private Path As String
Private FileName As String
Private ImageFolder As String

Private Sub habilitaCampos(bAcao As Boolean)
    Text_End.enabled = bAcao
    Text_Bairro.enabled = bAcao
    Text_Cidade.enabled = bAcao
    Combo_Uf.enabled = bAcao
    Text_Cep.enabled = bAcao
    Text_Ref.enabled = bAcao
    Combo_Zona.enabled = bAcao
    'txtTipoTelhado.enabled = bAcao
    txtTelResidencial.enabled = bAcao
    txtTelCom.enabled = bAcao
    txtTelCelular.enabled = bAcao
End Sub


Private Sub Form_Activate()
    
    If (MDIForm_Xsol.acessibilidade_TextoNegrito = True) Or (MDIForm_Xsol.acessibilidade_TextoNegrito = False) Then
        Combo_Uf.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Combo_Zona.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        'cboContato.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Cod_Cliente.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Nome_Cliente.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Cod.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Ref.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Cep.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Cidade.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Bairro.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_End.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        
        Grid_Address.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        
        'labels
        Label1.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Cliente.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Registro.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Zona.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Ref.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Cep.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_UF.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Cidade.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Bairro.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_End.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame1.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame4.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Frame6.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Inclui.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Edita.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Salva.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Exclui.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        
    End If

End Sub

Private Sub Form_Load()
   Text_Cod = Cod_Address
   Text_Cod_Cliente.Text = Cod_Cli
   Text_Nome_Cliente.Text = Nome_Cli_Atend
   
   Call Object_Controller(1)
   Call Configure_Grid
   Call Show_Customer_Address(Return_Recordset(Return_SqlString(1, str(Cod_Cli))), str(IIf(Cod_Address = "", "-1", Cod_Address)))
   
   Call Grid_Address_Click
   habilitaCampos (False)
   
End Sub

Private Sub Cmd_Edita_Click()
    If Cmd_Edita.Caption = "&Editar" Then
        Call Object_Controller(4)
        habilitaCampos (True)
    Else
        Call Object_Controller(5)
        habilitaCampos (False)
    End If
End Sub

Private Sub Cmd_Inclui_Click()
    If Cmd_Inclui.Caption = "&Incluir" Then
        Call Object_Controller(3)
        habilitaCampos (True)
    Else
        Call Clean_Boxes
        Call Object_Controller(5)
        habilitaCampos (False)
    End If
    
    
End Sub





'2222222222222
Private Sub enderecoExistente(rs As ADODB.Recordset, Optional code As Integer)
    
    If rs.RecordCount = 0 Then
       Exit Sub
    End If
    
    
    selectedRow = 2
    With frmEnderecoExistente.Grid_Address
        .Rows = 2
        
        Do Until rs.EOF
            .AddItem (IIf(IsNull(rs.Fields("endereco")), "", rs.Fields("endereco")) & vbTab & IIf(IsNull(rs.Fields("Nome_cli")), "", rs.Fields("Nome_cli")) & vbTab & IIf(IsNull(rs.Fields("Fone_celular")), "", rs.Fields("Fone_celular")))
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
    
    frmEnderecoExistente.Show (1)
    
End Sub
'22222222222


Private Sub Read_Config_File()
    ConfigFile = App.Path & "\CONFIG.INI"
    
    Path = RTrim(GetINISetting("DATABASE", "PATH", ConfigFile))
    If Right(Path, 1) <> "\" Then
       Path = Path + "\"
    End If
    FileName = GetINISetting("DATABASE", "FILENAME", ConfigFile)
    
    ImageFolder = GetINISetting("IMAGE", "FOLDER", ConfigFile)
End Sub


Private Sub Cmd_Salva_Click()
    
    Dim orcamento As String
    
    Text_DatabaseFolder = Path
    Text_DatabaseName = FileName
    Text_ImageFolder = ImageFolder
    
    If Return_Recordset2(Return_SqlString(6, Text_End.Text)).RecordCount > 0 Then
         frmEnderecoExistente.Show (1)
    End If
    
    If Not Mandatory_Fields_Filled Then
        Exit Sub
    End If
    
    cloud_campos = "Bairro,CEP,Cidade,Cod_cliente,Endereco,fonecel,fonecom,foneres,Ref,UF,Zona."
    cloud_values = Text_Bairro.Text & textoComVirgulaEsquerda(Text_Cep.Text) & textoComVirgulaEsquerda(Text_Cidade.Text) & textoComVirgulaEsquerda(Text_Cod_Cliente.Text) & textoComVirgulaEsquerda(Text_End.Text) & textoComVirgulaEsquerda(txtTelCelular.Text) & textoComVirgulaEsquerda(txtTelCom.Text) & textoComVirgulaEsquerda(txtTelResidencial.Text) & textoComVirgulaEsquerda(Text_Ref.Text) & textoComVirgulaEsquerda(Combo_Uf.Text) & textoComVirgulaEsquerda(Combo_Zona.Text)
    
    If Cmd_Inclui.Caption = "&Cancela" Then
        'Insere
        Call cloud_Insert("Endereco", cloud_campos, cloud_values)
    Else
        'Edita
        Call cloud_Edit("Endereco", cloud_campos, cloud_values, "Codigo = " & Text_Cod.Text)
    End If
    
    Call Object_Controller(5)
    Call Show_Customer_Address(Return_Recordset(Return_SqlString(1, Text_Cod_Cliente.Text)))
    Form_Customer.Refresh_Grid_Address
    
    habilitaCampos (False)
End Sub

Private Sub Cmd_Exclui_Click()
    If MsgBox("Confirma a Exclusão do Endereço " & Text_End.Text & " do Cliente " & Text_Nome_Cliente.Text & " ? ", vbYesNo, "Excluir Registros") = vbYes Then
        If Have_Reference() Then
            MsgBox "Não é possível sua remoção pois há uma OS que referencia a este endereço", vbInformation, "Informação"
            Exit Sub
        End If
        Call Return_Recordset(Return_SqlString(2, Text_Cod.Text)).Delete
        Call Object_Controller(1)
        Call Show_Customer_Address(Return_Recordset(Return_SqlString(1, Text_Cod_Cliente.Text)))
        Form_Customer.Refresh_Grid_Address
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
   rs.Fields("Endereco").value = UCase(Text_End.Text)
   rs.Fields("Bairro").value = UCase(Text_Bairro.Text)
   rs.Fields("Cidade").value = UCase(Text_Cidade.Text)
   rs.Fields("Uf").value = Combo_Uf.Text
   If Trim(Text_Cep.Text) = "" Then
      rs.Fields("Cep").value = Null
   Else
      rs.Fields("Cep").value = Text_Cep.Text
   End If
   If Trim(Text_Ref.Text) = "" Then
      rs.Fields("Ref").value = Null
   Else
      rs.Fields("Ref").value = UCase(Text_Ref.Text)
   End If
   rs.Fields("Zona").value = Combo_Zona.Text
   
   rs.Fields("FONERES").value = txtTelResidencial.Text
   rs.Fields("FONECOM").value = txtTelCom.Text
   rs.Fields("FONECEL").value = txtTelCelular.Text
End Sub
Sub Clean_Boxes()
    Text_Cod.Text = ""
    Text_End.Text = ""
    Text_Bairro.Text = ""
    Text_Cidade.Text = ""
    Combo_Uf.Text = "SP"
    Text_Cep.Text = ""
    Text_Ref.Text = ""
    Combo_Zona.ListIndex = -1
End Sub

Private Function Mandatory_Fields_Filled() As Boolean
    
    Mandatory_Fields_Filled = True

    If Trim(Text_End.Text) = "" Then
       MsgBox "É necessário digitar o Endereço! ", vbOKOnly, "XSol Aquecedores"
       Text_End.SetFocus
       Mandatory_Fields_Filled = False
       Exit Function
    End If
    
    If Trim(Text_Bairro.Text) = "" Then
       MsgBox "É necessário digitar o Bairro! ", vbOKOnly, "XSol Aquecedores"
       Text_Bairro.SetFocus
       Mandatory_Fields_Filled = False
       Exit Function
    End If
    
    If Trim(Text_Cidade.Text) = "" Then
       MsgBox "É necessário digitar Cidade! ", vbOKOnly, "XSol Aquecedores"
       Text_Cidade.SetFocus
       Mandatory_Fields_Filled = False
       Exit Function
    End If
    
    If Combo_Uf.Text = "" Then
       MsgBox "É necessário selecionar a Unidade Federativa! ", vbOKOnly, "XSol Aquecedores"
       Combo_Uf.SetFocus
       Mandatory_Fields_Filled = False
       Exit Function
    End If
    
    If Combo_Zona.Text = "" Then
       MsgBox "É necessário selecionar a ZONA! ", vbOKOnly, "XSol Aquecedores"
       Combo_Zona.SetFocus
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
        Combo_Uf.enabled = False
        Combo_Zona.enabled = False
        Grid_Address.enabled = True
        Call Combo_Initialize
        Call Clean_Boxes
    ElseIf cType = 2 Then  'record selected
        Cmd_Edita.enabled = True
        Cmd_Exclui.enabled = True
        Combo_Uf.enabled = False
        Combo_Zona.enabled = False
        Grid_Address.enabled = True
    ElseIf cType = 3 Then  'insert a record
        Cmd_Inclui.Caption = "&Cancela"
        Cmd_Edita.enabled = False
        Combo_Uf.enabled = True
        Combo_Zona.enabled = True
        Call insEdit_Controller
        Call Combo_Initialize(True)
        Call Clean_Boxes
    ElseIf cType = 4 Then  'Edit a record
        Cmd_Inclui.enabled = False
        Cmd_Edita.Caption = "&Cancela"
        Combo_Uf.enabled = True
        Combo_Zona.enabled = True
        Call insEdit_Controller
    ElseIf cType = 5 Then  'Cancel a insertion or edition or after saving
        Cmd_Inclui.Caption = "&Incluir"
        Cmd_Inclui.enabled = True
        Cmd_Salva.enabled = False
        Cmd_Edita.Caption = "&Editar"
        Cmd_Edita.enabled = False
        Cmd_Exclui.enabled = False
        Combo_Uf.enabled = False
        Combo_Zona.enabled = False
        Grid_Address.enabled = True
        
        Text_End.BackColor = vbWhite
        Text_Bairro.BackColor = vbWhite
        Text_Cidade.BackColor = vbWhite
        Combo_Uf.BackColor = vbWhite
        Combo_Zona.BackColor = vbWhite

        
        Call Combo_Initialize
        Call Grid_Address_Click
    End If
End Sub

Private Sub insEdit_Controller()
    Cmd_Salva.enabled = True
    Cmd_Exclui.enabled = False
    Grid_Address.enabled = False
    
    Text_End.BackColor = vbYellow
    Text_Bairro.BackColor = vbYellow
    Text_Cidade.BackColor = vbYellow
    Combo_Uf.BackColor = vbYellow
    Combo_Zona.BackColor = vbYellow
    'Text_End.SetFocus
End Sub

Private Sub Configure_Grid()
    With Grid_Address
        .Cols = 4
        .ColWidth(0) = 630
        .ColWidth(1) = 3870
        .ColWidth(2) = 2100
        .ColWidth(3) = 2100
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Endereço"
        .TextMatrix(0, 2) = "Bairro"
        .TextMatrix(0, 3) = "Cidade"
        .RowHeight(1) = 0
    End With
End Sub

Private Function Return_Recordset(strSql As String) As ADODB.Recordset
    
    
    Set cloud_registro = cloud_conexao.Execute(strSql)
    Set Return_Recordset = cloud_registro
End Function

Private Function Return_Recordset2(strSql As String) As ADODB.Recordset
On Error GoTo reconecta
    
inicia:
    
    Set cloud_registro = cloud_conexao.Execute(strSql)

    Set Return_Recordset2 = cloud_registro
    Exit Function
reconecta:
        If ConectaMySQL Then
            GoTo inicia:
        Else
            MsgBox "O sistema perdeu a conexão com  banco de dados e será encerrado!", vbCritical, "Falha na conexão com mysql em KingHost"
            Unload MDIForm_Xsol
        End If

End Function



Private Sub Show_Customer_Address(rs As ADODB.Recordset, Optional code As Integer)
    Dim selectedRow As Integer
    
    selectedRow = 2
    With Grid_Address
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("codigo") & vbTab & IIf(IsNull(rs.Fields("Endereco")), "", rs.Fields("Endereco")) & vbTab & IIf(IsNull(rs.Fields("Bairro")), "", rs.Fields("Bairro")) & vbTab & IIf(IsNull(rs.Fields("Cidade")), "", rs.Fields("Cidade")))
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

Private Sub Grid_Address_Click()
    With Grid_Address
        If .RowSel = 1 Then
            Exit Sub
        End If
        Call Show_Record(Return_Recordset(Return_SqlString(2, .TextMatrix(.RowSel, 0))))
    End With
End Sub

Private Sub Grid_Address_SelChange()
    Call Grid_Address_Click
End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String, Optional param3 As String) As String
    Dim whereClause As String
    
    If queryType = 1 Then
        Return_SqlString = "SELECT * FROM `Endereco` WHERE cod_cliente = " & param1 & " ORDER BY Codigo"
    ElseIf queryType = 2 Then
        Return_SqlString = "SELECT * FROM `Endereco` WHERE codigo = " & param1
    ElseIf queryType = 3 Then
        Return_SqlString = "SELECT * FROM `Endereco`"
    ElseIf queryType = 4 Then
        Return_SqlString = "SELECT * FROM `Atendimento` WHERE Cod_End = " & param1
    ElseIf queryType = 5 Then
        Return_SqlString = "SELECT * FROM `Zona`"
    ElseIf queryType = 6 Then
        Return_SqlString = "Select  p2.endereco, a.Nome_Cli, a.Fone_celular, a.Fone_res, a.Fone_Com, a.e_mail From `Clientes` a INNER JOIN `Endereco` p2 on  p2.cod_cliente = a.Codigo where endereco like '*" & param1 & "*'"
    End If
End Function

Sub Show_Record(rs As ADODB.Recordset)

    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    Text_Cod.Text = rs.Fields("Codigo")

    Text_End.Text = UCase(IIf(IsNull(rs.Fields("Endereco")), "", rs.Fields("Endereco")))
    
    Text_Bairro.Text = UCase(IIf(IsNull(rs.Fields("Bairro")), "", rs.Fields("Bairro")))
    
    Text_Cidade.Text = UCase(IIf(IsNull(rs.Fields("Cidade")), "", rs.Fields("Cidade")))
    
    If Not IsNull(rs.Fields("Uf")) Then
        Combo_Uf.Text = UCase(rs.Fields("Uf"))
    End If
    
    Text_Cep.Text = IIf(IsNull(rs.Fields("Cep")), "", rs.Fields("Cep"))
    
    Text_Ref.Text = UCase(IIf(IsNull(rs.Fields("Ref")), "", rs.Fields("Ref")))
    
    
    If Not IsNull(rs.Fields("Zona")) Then
        'Combo_Zona.Text = UCase(rs.Fields("Zona"))
        Call SetComboItem(Combo_Zona, rs.Fields("Zona"))
    End If
    
    If Trim(rs.Fields("Tipo_Telhado")) <> "" Then
        txtTipoTelhado.Text = rs.Fields("Tipo_Telhado")
    End If
    
    If Trim(rs.Fields("FONERES")) <> "" Then
        txtTelResidencial.Text = rs.Fields("FONERES")
    Else
        txtTelResidencial.Text = ""
    End If
    
    If Trim(rs.Fields("FONECOM")) <> "" Then
        txtTelCom.Text = rs.Fields("FONECOM")
    Else
        txtTelCom.Text = ""
    End If
    
    If Trim(rs.Fields("FONECEL")) <> "" Then
        txtTelCelular.Text = rs.Fields("FONECEL")
    Else
        txtTelCelular.Text = ""
    End If
    
    
    rs.Close
    Set rs = Nothing
    
    Call Object_Controller(2)
End Sub

Sub Combo_Initialize(Optional toInsert As Boolean = False)
    Combo_Uf.Clear
    Combo_Uf.AddItem "SP"
    If Not toInsert Then
        Combo_Uf.AddItem "PR"
        Combo_Uf.AddItem "RJ"
        Combo_Uf.AddItem "SC"
        Combo_Uf.AddItem "SE"
    End If
    
    Call Initialize_Zone
    
'    Combo_Zona.Clear
'    Combo_Zona.AddItem "INTERIOR"
'    Combo_Zona.AddItem "LITORAL"
'    Combo_Zona.AddItem "ABC"
'    Combo_Zona.AddItem "ZONA NORTE"
'    Combo_Zona.AddItem "ZONA LESTE"
'    Combo_Zona.AddItem "ZONA SUL"
'    Combo_Zona.AddItem "ZONA OESTE"
End Sub

Sub Initialize_Zone()
   Dim rs As ADODB.Recordset
   
   Combo_Zona.Clear
   
   Set rs = Return_Recordset(Return_SqlString(5))

   
   Do Until rs.EOF Or rs.BOF
      Combo_Zona.AddItem rs.Fields("Zona")
      rs.MoveNext
   Loop
   
   rs.Close
   Set rs = Nothing
End Sub

