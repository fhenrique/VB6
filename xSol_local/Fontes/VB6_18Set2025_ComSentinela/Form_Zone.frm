VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Begin VB.Form Form_Zone 
   Caption         =   "Cadastro de Zonas"
   ClientHeight    =   2865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   12915
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Lista de Zonas"
      Height          =   2775
      Left            =   6720
      TabIndex        =   8
      Top             =   0
      Width           =   6135
      Begin MSFlexGridLib.MSFlexGrid Grid_Main 
         Height          =   2415
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   4260
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
      TabIndex        =   3
      Top             =   1680
      Width           =   6615
      Begin VB.CommandButton Cmd_Inclui 
         BackColor       =   &H0080FF80&
         Caption         =   "&Incluir"
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Salva 
         Caption         =   "&Salvar"
         Height          =   495
         Left            =   3480
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Edita 
         Caption         =   "&Editar"
         Height          =   495
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Exclui 
         BackColor       =   &H008080FF&
         Caption         =   "E&xcluir"
         Height          =   495
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados da Zona"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
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
         Left            =   1440
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text_Zone 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         TabIndex        =   1
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Código:"
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Zona:"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   960
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form_Zone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const GRID_NUMBER_OF_ROWS = 9  'Number of rows without use scroll

Private Sub Form_Activate()
Text_Code.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Zone.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Grid_Main.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
End Sub

Private Sub Form_Load()
   Call Configure_Grid
    
   Call Fill_Grid(Return_Recordset(Return_SqlString(1)))
End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As Integer, Optional param2 As String) As String
    Dim whereClause As String
    
    If queryType = 1 Then
        Return_SqlString = "SELECT * FROM Zona ORDER BY Zona"
    ElseIf queryType = 2 Then
        Return_SqlString = "SELECT * FROM Zona WHERE Codigo = " & param1
    ElseIf queryType = 3 Then
        Return_SqlString = "SELECT COUNT(*) AS qtt FROM Endereco WHERE Zona = '" & param2 & "'"
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
            .AddItem (rs.Fields("codigo") & vbTab & rs.Fields("Zona"))
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
    Text_Zone.Text = rs.Fields("Zona")
            
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
   If Have_Reference() Then
      MsgBox "Não é possível editar " & Text_Zone & " pois há endereço(s) que referencia(m) esta zona", vbInformation, "Informação"
      Exit Sub
   End If
   If Cmd_Edita.Caption = "&Editar" Then
      Call Object_Controller(4)
   Else
      Call Clean_Boxes
      Call Object_Controller(2)
      Call Grid_Main_Click
   End If
End Sub

Private Sub Cmd_Salva_Click()
   Dim rs As ADODB.Recordset
    
   If Not Mandatory_Fields_Filled Then
      Exit Sub
   End If
    
   If Cmd_Inclui.Caption = "&Cancela" Then
      Set rs = Return_Recordset(Return_SqlString(1))
      rs.AddNew
      Text_Code.Text = rs.Fields("Codigo")
   Else
      Set rs = Return_Recordset(Return_SqlString(2, Text_Code))
   End If
    
   Call Record_Record(rs)

   rs.Update
   
   rs.Close
   Set rs = Nothing
    
   Call Fill_Grid(Return_Recordset(Return_SqlString(1)), Text_Code)
End Sub

Private Sub Record_Record(rs As ADODB.Recordset)
   rs.Fields("Zona").value = Text_Zone
End Sub

Private Sub Cmd_Exclui_Click()
   If Have_Reference() Then
      MsgBox "Não é possível remover " & Text_Zone & " pois há endereço(s) que referencia(m) esta zona", vbInformation, "Informação"
      Exit Sub
   End If
   If MsgBox("Confirma a Exclusão de " & Text_Zone & " ? ", vbYesNo, "Excluir Registros") = vbYes Then
      Call Return_Recordset(Return_SqlString(2, Text_Code)).Delete
      Call Object_Controller(1)
      Call Fill_Grid(Return_Recordset(Return_SqlString(1)))
   End If
End Sub

Private Function Have_Reference() As Boolean
    Dim rs As ADODB.Recordset

    Set rs = Return_Recordset(Return_SqlString(3, param2:=Text_Zone))
    
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
    
    If Trim(Text_Zone) = "" Then
       MsgBox "É necessário digitar a Zona! ", vbOKOnly, "XSol Aquecedores"
       Text_Zone.SetFocus
       Mandatory_Fields_Filled = False
       Exit Function
    End If
    
End Function

Private Sub Object_Controller(cType As Integer)

    Text_Zone.enabled = cType >= 3
    
    Text_Zone.BackColor = IIf(cType >= 3, vbYellow, vbWhite)
    
    Cmd_Inclui.enabled = cType <> 4
    Cmd_Inclui.Caption = IIf(cType = 3, "&Cancela", "&Incluir")
    Cmd_Edita.enabled = cType = 2 Or cType = 4
    Cmd_Edita.Caption = IIf(cType = 4, "&Cancela", "&Editar")
    Cmd_Salva.enabled = cType >= 3
    Cmd_Exclui.enabled = cType = 2
    Grid_Main.enabled = cType <= 2
    
    If cType = 1 Then  'without any selection
        Call Clean_Boxes
    ElseIf cType = 3 Then  'insert a record
        Text_Zone.SetFocus
        Call Clean_Boxes
    End If
End Sub

Sub Clean_Boxes()
    Text_Code = ""
    Text_Zone = ""
End Sub

Private Sub Configure_Grid()
    With Grid_Main
        .Cols = 2
        .ColWidth(0) = 800
        .ColWidth(1) = 4700
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Zona"
        .RowHeight(1) = 0
    End With
End Sub

Private Sub Grid_Main_Click()
    With Grid_Main
        If .RowSel = 1 Then
            Exit Sub
        End If
        Call Show_Record_Detail(Return_Recordset(Return_SqlString(2, .TextMatrix(.RowSel, 0))))
    End With
End Sub



