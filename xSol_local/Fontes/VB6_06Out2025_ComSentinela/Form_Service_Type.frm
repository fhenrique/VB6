VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_Service_Type 
   Caption         =   "Cadastro de Tipos de Serviços"
   ClientHeight    =   5280
   ClientLeft      =   3120
   ClientTop       =   3375
   ClientWidth     =   13230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   13230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Lista de Tipos de Serviços"
      Height          =   5235
      Left            =   6120
      TabIndex        =   10
      Top             =   0
      Width           =   7080
      Begin VB.Frame Frame2 
         Caption         =   "Pesquisa"
         Height          =   1185
         Left            =   90
         TabIndex        =   15
         Top             =   3960
         Width           =   6885
         Begin VB.CommandButton btnAtualizar 
            Caption         =   "Atualizar"
            Height          =   390
            Left            =   5790
            TabIndex        =   21
            Top             =   705
            Width           =   990
         End
         Begin VB.TextBox txtPorDescricao 
            Height          =   330
            Left            =   1050
            MaxLength       =   20
            TabIndex        =   20
            Top             =   750
            Width           =   3345
         End
         Begin VB.TextBox txtPorCodigo 
            Height          =   330
            Left            =   1035
            MaxLength       =   4
            TabIndex        =   19
            Top             =   300
            Width           =   1095
         End
         Begin VB.CommandButton btnPesquisar 
            Caption         =   "Pesquisar"
            Height          =   360
            Left            =   4620
            TabIndex        =   18
            Top             =   720
            Width           =   990
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Descrição:"
            Height          =   255
            Left            =   165
            TabIndex        =   17
            Top             =   885
            Width           =   765
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Código:"
            Height          =   255
            Left            =   330
            TabIndex        =   16
            Top             =   405
            Width           =   615
         End
      End
      Begin VB.OptionButton optnPorDescricao 
         Caption         =   "Ordernar por descrição"
         Height          =   285
         Left            =   3540
         TabIndex        =   14
         Top             =   3720
         Width           =   2385
      End
      Begin VB.OptionButton optnPorCodigo 
         Caption         =   "Ordernar por código"
         Height          =   285
         Left            =   1575
         TabIndex        =   13
         Top             =   3690
         Value           =   -1  'True
         Width           =   2010
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Main 
         Height          =   3510
         Left            =   75
         TabIndex        =   11
         Top             =   210
         Width           =   6960
         _ExtentX        =   12277
         _ExtentY        =   6191
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Edição"
      Height          =   1485
      Left            =   15
      TabIndex        =   5
      Top             =   3705
      Width           =   6030
      Begin VB.CommandButton Cmd_Exclui 
         BackColor       =   &H008080FF&
         Caption         =   "E&xcluir"
         Height          =   495
         Left            =   4665
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   525
         Width           =   945
      End
      Begin VB.CommandButton Cmd_Edita 
         Caption         =   "Editar"
         Height          =   495
         Left            =   1785
         TabIndex        =   8
         Top             =   525
         Width           =   990
      End
      Begin VB.CommandButton Cmd_Salva 
         Caption         =   "&Salvar"
         Height          =   495
         Left            =   3180
         TabIndex        =   7
         Top             =   525
         Width           =   1005
      End
      Begin VB.CommandButton Cmd_Inclui 
         BackColor       =   &H0080FF80&
         Caption         =   "&Incluir"
         Height          =   495
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   525
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Tipo de Serviço"
      Height          =   3675
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CheckBox chkAgenda1Ano 
         Caption         =   "Agendamento automático de próxima visita após 1 ano."
         Enabled         =   0   'False
         Height          =   375
         Left            =   1530
         TabIndex        =   12
         Top             =   2055
         Width           =   4215
      End
      Begin VB.TextBox Text_Code 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1545
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text_Service_Type 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1530
         TabIndex        =   1
         Top             =   1560
         Width           =   4215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Código:"
         Height          =   255
         Left            =   810
         TabIndex        =   4
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo de Serviço:"
         Height          =   255
         Left            =   210
         TabIndex        =   3
         Top             =   1575
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form_Service_Type"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const GRID_NUMBER_OF_ROWS = 10  'Number of rows without use scroll
Public strEditando As Boolean

Private Sub btnAtualizar_Click()
   Dim tipoOrderBy As Integer
   
   If optnPorCodigo.value = 1 Then
    tipoOrderBy = 1
   Else
    tipoOrderBy = 2
   End If
   
   
   Call Fill_Grid(Return_Recordset(Return_SqlString(1, tipoOrderBy)))

End Sub

Private Sub btnPesquisar_Click()

    'Dim tipoOrderBy As Integer

    If (Trim(txtPorCodigo.Text) = "") And (Trim(txtPorDescricao.Text) = "") Then
        MsgBox ("Informe um código ou descrição para busca.")
        Exit Sub
    End If
    
    
    If Trim(txtPorCodigo.Text) <> "" Then
        Call Fill_Grid(Return_Recordset(Return_SqlString(2, txtPorCodigo.Text)))
    Else
        Call Fill_Grid(Return_Recordset(Return_SqlString(6, , txtPorDescricao.Text)))
    End If

End Sub



Private Sub Form_Activate()
 
 Text_Code.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
 Text_Service_Type.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
 Grid_Main.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
End Sub

Private Sub Form_Load()
   Call Configure_Grid
   Dim tipoOrderBy As Integer
   
   If optnPorCodigo.value = 1 Then
    tipoOrderBy = 1
   Else
    tipoOrderBy = 2
   End If
   
   
   Call Fill_Grid(Return_Recordset(Return_SqlString(1, tipoOrderBy)))
   

End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As Integer, Optional param2 As String) As String
    Dim whereClause As String
    Dim orderbyparam As String
    
    If param1 = 1 Then
        orderbyparam = "Codigo"
    Else
        orderbyparam = "Tipo_Serv"
    End If
    
    
    If queryType = 1 Then
        Return_SqlString = "SELECT * FROM Servico ORDER BY " & orderbyparam
    ElseIf queryType = 2 Then
        If param2 = "" Then
            param2 = "Tipo_Serv"
        End If
        
        Return_SqlString = "SELECT * FROM Servico WHERE Codigo = " & param1
    ElseIf queryType = 3 Then
        Return_SqlString = "SELECT COUNT(*) AS qtt FROM Atendimento WHERE Tipo_Serv = '" & param2 & "'"
    ElseIf queryType = 4 Then
        Return_SqlString = "SELECT * FROM Servico"
    ElseIf queryType = 6 Then
        Return_SqlString = "SELECT * FROM Servico Where Tipo_Serv like '%" & param2 & "%'"
    End If
End Function

Private Sub Fill_Grid(rs As ADODB.Recordset, Optional code As Integer)
    Dim selectedRow As Integer
    
    selectedRow = 2
    With Grid_Main
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("codigo") & vbTab & rs.Fields("Tipo_Serv") & vbTab & rs.Fields("autoagendamento"))
            If code = rs.Fields("codigo") Then
                selectedRow = .Rows - 1
            End If
            rs.MoveNext
        Loop
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            If .Rows > GRID_NUMBER_OF_ROWS Then
                .TopRow = selectedRow
            End If
            .col = 0
            .ColSel = .Cols - 1
            
            Call Object_Controller(2)
            Call Grid_Main_Click
        Else
            Call Object_Controller(1)
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Sub Show_Record_Detail(rs As ADODB.Recordset)
    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    Text_Code.Text = rs.Fields("Codigo")
    Text_Service_Type.Text = rs.Fields("Tipo_Serv")
    
    If rs.Fields("autoagendamento") = "S" Then
        chkAgenda1Ano.value = 1
    Else
        chkAgenda1Ano.value = 0
    End If
        
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Cmd_Inclui_Click()
    If Cmd_Inclui.Caption = "&Incluir" Then
        Call Object_Controller(3)
    Else
        Call Clean_Boxes
        Call Object_Controller(2)
        Call Grid_Main_Click
    End If
End Sub

Private Sub Cmd_Edita_Click()
   'If Have_Reference() Then
   '   MsgBox "Não é possível editar " & Text_Service_Type & " pois há atendimento(s) que referencia(m) este tipo de serviço", vbInformation, "Informação"
   '   Exit Sub
   'End If
   
   strEditando = True
   
   If Cmd_Edita.Caption = "&Editar" Then
      chkAgenda1Ano.enabled = True
      Call Object_Controller(4)
   Else
      chkAgenda1Ano.enabled = False
      If Not strEditando Then
        Call Clean_Boxes
      End If
        Call Object_Controller(2)
      Call Grid_Main_Click
   End If
   
'9-10-51-13
   
   
   
   If (Text_Code.Text = 9) Or (Text_Code.Text = 10) Or (Text_Code.Text = 13) Or (Text_Code.Text = 51) Then
        chkAgenda1Ano.enabled = False
   End If
   
End Sub

Private Sub Cmd_Salva_Click()
   Dim rs As ADODB.Recordset
    
   If Not Mandatory_Fields_Filled Then
      Exit Sub
   End If
    
   If Cmd_Inclui.Caption = "&Cancela" Then
      Set rs = Return_Recordset(Return_SqlString(4))
      rs.AddNew
      Text_Code.Text = rs.Fields("Codigo")
   Else
      Set rs = Return_Recordset(Return_SqlString(2, Text_Code))
   End If
    
   Call Record_Record(rs)

   rs.Update
   
   rs.Close
   Set rs = Nothing
   
   chkAgenda1Ano.enabled = False
    
   'Call Fill_Grid(Return_Recordset(Return_SqlString(1)), Text_Code)
   Call Fill_Grid(Return_Recordset(Return_SqlString(1, 1)))
End Sub

Private Sub Record_Record(rs As ADODB.Recordset)
    rs.Fields("Tipo_Serv").value = Text_Service_Type
   
    If chkAgenda1Ano.value = 1 Then
        rs.Fields("autoagendamento").value = "S"
    Else
        rs.Fields("autoagendamento").value = "N"
    End If

End Sub

Private Sub Cmd_Exclui_Click()
   If Have_Reference Then
      MsgBox "Não é possível remover " & Text_Service_Type & " pois há atendimento(s) que referencia(m) este tipo de serviço", vbInformation, "Informação"
      Exit Sub
   End If
   If MsgBox("Confirma a Exclusão de " & Text_Service_Type & " ? ", vbYesNo, "Excluir Registros") = vbYes Then
      Call Return_Recordset(Return_SqlString(2, Text_Code)).Delete
      Call Object_Controller(1)
      Call Fill_Grid(Return_Recordset(Return_SqlString(1)))
   End If
End Sub

Private Function Have_Reference() As Boolean
    Dim rs As ADODB.Recordset

    Set rs = Return_Recordset(Return_SqlString(3, , Text_Service_Type.Text))
    
    If rs.Fields("qtt") > 0 Then
        Have_Reference = True
    Else
        Have_Reference = False
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function Mandatory_Fields_Filled() As Boolean
    
    Mandatory_Fields_Filled = True
    
    If Trim(Text_Service_Type) = "" Then
       MsgBox "É necessário digitar o Tipo de Serviço! ", vbOKOnly, "XSol Aquecedores"
       Text_Service_Type.SetFocus
       Mandatory_Fields_Filled = False
       Exit Function
    End If
    
End Function

Private Sub Object_Controller(cType As Integer)

    Text_Service_Type.enabled = cType >= 3
    'chkAgenda1Ano.enabled = cType >= 3
    
    Text_Service_Type.BackColor = IIf(cType >= 3, vbYellow, vbWhite)
    
    Cmd_Inclui.enabled = cType <> 4
    'chkAgenda1Ano.enabled = cType <> 4
    
    Cmd_Inclui.Caption = IIf(cType = 3, "&Cancela", "&Incluir")
    Cmd_Edita.enabled = cType = 2 Or cType = 4
    Cmd_Edita.Caption = IIf(cType = 4, "&Cancela", "&Editar")
    Cmd_Salva.enabled = cType >= 3
    Cmd_Exclui.enabled = cType = 2
    Grid_Main.enabled = cType <= 2
    
    If cType = 1 Then  'without any selection
        Call Clean_Boxes
    ElseIf cType = 3 Then  'insert a record
        Text_Service_Type.SetFocus
        Call Clean_Boxes
        chkAgenda1Ano.value = 0
    End If
    
    
    
End Sub

Sub Clean_Boxes()
    
    
    Text_Code = ""
    Text_Service_Type.Text = ""
End Sub

Private Sub Configure_Grid()
    With Grid_Main
        .Cols = 3
        .ColWidth(0) = 800
        .ColWidth(1) = 4700
        .ColWidth(2) = 1000
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Tipo de Serviço"
        .TextMatrix(0, 2) = "Auto Agend."
        .RowHeight(1) = 0
    End With
End Sub

Private Sub Grid_Main_Click()
'    With Grid_Main
'        If .RowSel = 1 Then
'            Exit Sub
'        End If
'        Call Show_Record_Detail(Return_Recordset(Return_SqlString(2, .TextMatrix(.RowSel, 0))))
'    End With
End Sub


Private Sub Grid_Main_EnterCell()
    'MsgBox "troca"
End Sub

Private Sub Grid_Main_SelChange()
    With Grid_Main
        If .RowSel = 1 Then
            Exit Sub
        End If
        Call Show_Record_Detail(Return_Recordset(Return_SqlString(2, .TextMatrix(.RowSel, 0))))
        
        
    End With
End Sub

Private Sub optnPorCodigo_Click()
    Call Fill_Grid(Return_Recordset(Return_SqlString(1, 1)))
End Sub

Private Sub optnPorDescricao_Click()
    Call Fill_Grid(Return_Recordset(Return_SqlString(1, 2)))
End Sub

Private Sub Text_Service_Type_LostFocus()
    Text_Service_Type.Text = UCase(Text_Service_Type.Text)
End Sub

Private Sub txtPorCodigo_Change()
    If Not (IsNumeric(txtPorCodigo.Text)) Then
        txtPorCodigo.Text = ""
    End If
End Sub

Private Sub txtPorCodigo_GotFocus()
    txtPorDescricao.Text = ""
End Sub

Private Sub txtPorDescricao_GotFocus()
    txtPorCodigo.Text = ""
End Sub
