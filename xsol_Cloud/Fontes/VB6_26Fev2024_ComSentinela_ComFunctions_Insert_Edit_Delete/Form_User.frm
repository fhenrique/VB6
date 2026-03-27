VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Begin VB.Form Form_User 
   Caption         =   "Cadastro de Usuário"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   11610
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Usuário"
      Height          =   1815
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6615
      Begin VB.TextBox Text_Passwd 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   840
         Width           =   3615
      End
      Begin VB.ComboBox Combo_Profile 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Text_User 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Senha:"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Lbl_Cliente 
         Alignment       =   1  'Right Justify
         Caption         =   "Perfil:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label U 
         Alignment       =   1  'Right Justify
         Caption         =   "Usuário:"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Edição"
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   6615
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
      Begin VB.CommandButton Cmd_Salva 
         Caption         =   "&Salvar"
         Height          =   495
         Left            =   3480
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Edita 
         Caption         =   "&Editar"
         Height          =   495
         Left            =   1920
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Exclui 
         BackColor       =   &H008080FF&
         Caption         =   "E&xcluir"
         Height          =   495
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Lista de Usuários"
      Height          =   2895
      Left            =   6720
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin MSFlexGridLib.MSFlexGrid Grid_Main 
         Height          =   2535
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         SelectionMode   =   1
      End
   End
End
Attribute VB_Name = "Form_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const GRID_NUMBER_OF_ROWS = 9  'Number of rows without use scroll

Private Sub Form_Activate()
    If (MDIForm_Xsol.acessibilidade_TextoNegrito = True) Or (MDIForm_Xsol.acessibilidade_TextoNegrito = False) Then
        Text_User.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Passwd.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Combo_Profile.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Grid_Main.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    End If
End Sub

Private Sub Form_Load()
   Call Combo_Set_Initialize
   
   Call Configure_Grid
    
   Call Fill_Grid(Return_Recordset(Return_SqlString(1)))
End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As String) As String
    Dim whereClause As String
    
    If queryType = 1 Then
        Return_SqlString = "SELECT * FROM Usuarios ORDER BY Usuario"
    ElseIf queryType = 2 Then
        Return_SqlString = "SELECT * FROM Usuarios WHERE Usuario = '" & param1 & "'"
    End If
End Function

Private Sub Fill_Grid(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    
    selectedRow = 2
    With Grid_Main
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("usuario") & vbTab & IIf(IsNull(rs.Fields("Perfil")), "", rs.Fields("Perfil")))
            If code = rs.Fields("usuario") Then
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
    
    Text_User.Text = rs.Fields("Usuario")
    
    Text_Passwd.Text = rs.Fields("Senha")

    Combo_Profile.Text = Get_Field_Val(rs.Fields("Perfil"))
        
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
      rs.Fields("Usuario") = Text_User
   Else
      Set rs = Return_Recordset(Return_SqlString(2, Text_User))
   End If
    
   Call Record_Record(rs)

   rs.Update
   
   rs.Close
   Set rs = Nothing
    
   Call Fill_Grid(Return_Recordset(Return_SqlString(1)), Text_User)
End Sub

Private Sub Record_Record(rs As ADODB.Recordset)
   rs.Fields("Senha").value = Text_Passwd
   rs.Fields("Perfil").value = Combo_Profile
End Sub

Private Sub Cmd_Exclui_Click()
   If MsgBox("Confirma a Exclusão do Usuário " & Text_User & " ? ", vbYesNo, "Excluir Registros") = vbYes Then
      Call Return_Recordset(Return_SqlString(2, Text_User)).Delete
      Call Object_Controller(1)
      Call Fill_Grid(Return_Recordset(Return_SqlString(1)))
   End If
End Sub

Private Function Mandatory_Fields_Filled() As Boolean
    
    Mandatory_Fields_Filled = True

    If Trim(Text_User.Text) = "" Then
       MsgBox "É necessário digitar o Usuário! ", vbOKOnly, "XSol Aquecedores"
       Text_User.SetFocus
       Mandatory_Fields_Filled = False
       Exit Function
    End If
    If Trim(Text_Passwd.Text) = "" Then
       MsgBox "É necessário digitar a Senha! ", vbOKOnly, "XSol Aquecedores"
       Text_Passwd.SetFocus
       Mandatory_Fields_Filled = False
       Exit Function
    End If
    If Trim(Combo_Profile.Text) = "" Then
       MsgBox "É necessário escolher um Perfil! ", vbOKOnly, "XSol Aquecedores"
       Combo_Profile.SetFocus
       Mandatory_Fields_Filled = False
       Exit Function
    End If
    
End Function

Private Sub Object_Controller(cType As Integer)

    Text_User.enabled = cType = 3
    Text_Passwd.enabled = cType >= 3
    Combo_Profile.enabled = Text_Passwd.enabled
    
    Text_User.BackColor = IIf(Text_User.enabled, vbYellow, vbWhite)
    Text_Passwd.BackColor = IIf(cType >= 3, vbYellow, vbWhite)
    Combo_Profile.BackColor = IIf(cType >= 3, vbYellow, vbWhite)
    
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
        Text_User.SetFocus
        Call Clean_Boxes
    End If
End Sub

Sub Clean_Boxes()
    Text_User.Text = ""
    Text_Passwd.Text = ""
    Combo_Profile.ListIndex = -1
End Sub

Sub Combo_Set_Initialize()
    Combo_Profile.Clear
    Combo_Profile.AddItem "Admin"
    Combo_Profile.AddItem "N1"
End Sub

Private Sub Configure_Grid()
    With Grid_Main
        .Cols = 2
        .ColWidth(0) = 2080
        .ColWidth(1) = 2080
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Usuário"
        .TextMatrix(0, 1) = "Perfil"
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

