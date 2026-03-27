VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Begin VB.Form frmGarantiaCobranca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Garantias e Cobranças"
   ClientHeight    =   6570
   ClientLeft      =   4245
   ClientTop       =   2355
   ClientWidth     =   9480
   Icon            =   "frmGarantiaCobranca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Cobranças"
      Height          =   3195
      Left            =   90
      TabIndex        =   10
      Top             =   3315
      Width           =   9300
      Begin VB.TextBox txtCobrancaDescricao 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1305
         TabIndex        =   16
         Top             =   975
         Width           =   2745
      End
      Begin VB.TextBox txtCobrancaCodigo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1305
         TabIndex        =   15
         Top             =   450
         Width           =   855
      End
      Begin VB.CommandButton btnCobrancasExcluir 
         BackColor       =   &H008080FF&
         Caption         =   "E&xcluir"
         Height          =   360
         Left            =   3255
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2640
         Width           =   765
      End
      Begin VB.CommandButton btnCobrancasEditar 
         Caption         =   "&Editar"
         Height          =   345
         Left            =   1215
         TabIndex        =   13
         Top             =   2655
         Width           =   885
      End
      Begin VB.CommandButton btnCobrancasSalvar 
         Caption         =   "&Salvar"
         Height          =   345
         Left            =   2295
         TabIndex        =   12
         Top             =   2655
         Width           =   735
      End
      Begin VB.CommandButton btnCobrancasIncluir 
         BackColor       =   &H0080FF80&
         Caption         =   "&Incluir"
         Height          =   345
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2655
         Width           =   870
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Cobrancas 
         Height          =   2955
         Left            =   4230
         TabIndex        =   17
         Top             =   195
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   5212
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
         Height          =   255
         Left            =   210
         TabIndex        =   19
         Top             =   975
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Código:"
         Height          =   255
         Left            =   555
         TabIndex        =   18
         Top             =   450
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Garantias"
      Height          =   3240
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   9315
      Begin VB.CommandButton btnGarantiasIncluir 
         BackColor       =   &H0080FF80&
         Caption         =   "&Incluir"
         Height          =   345
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2610
         Width           =   810
      End
      Begin VB.CommandButton btnGarantiasSalvar 
         Caption         =   "&Salvar"
         Height          =   345
         Left            =   2295
         TabIndex        =   7
         Top             =   2610
         Width           =   735
      End
      Begin VB.CommandButton btnGarantiasEditar 
         Caption         =   "&Editar"
         Height          =   345
         Left            =   1215
         TabIndex        =   6
         Top             =   2610
         Width           =   870
      End
      Begin VB.CommandButton btnGarantiasExcluir 
         BackColor       =   &H008080FF&
         Caption         =   "E&xcluir"
         Height          =   360
         Left            =   3255
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2595
         Width           =   765
      End
      Begin VB.TextBox txtGarantiaCodigo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1305
         TabIndex        =   2
         Top             =   375
         Width           =   855
      End
      Begin VB.TextBox txtGarantiaDescricao 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1290
         TabIndex        =   1
         Top             =   870
         Width           =   2745
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Garantias 
         Height          =   2955
         Left            =   4215
         TabIndex        =   9
         Top             =   195
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5212
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin VB.Label Lbl_Registro 
         Alignment       =   1  'Right Justify
         Caption         =   "Código:"
         Height          =   255
         Left            =   555
         TabIndex        =   4
         Top             =   375
         Width           =   615
      End
      Begin VB.Label Lbl_End 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição:"
         Height          =   255
         Left            =   210
         TabIndex        =   3
         Top             =   870
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmGarantiaCobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub TabStrip1_Click()

End Sub

Private Sub insEdit_Controller_Garantias()
    btnGarantiasSalvar.enabled = True
    btnGarantiasExcluir.enabled = False
    Grid_Garantias.enabled = False
    
    txtGarantiaDescricao.BackColor = vbYellow
    txtGarantiaDescricao.SetFocus
End Sub

Private Sub insEdit_Controller_Cobrancas()
    btnCobrancasSalvar.enabled = True
    btnCobrancasExcluir.enabled = False
    Grid_Cobrancas.enabled = False
    
    txtCobrancaDescricao.BackColor = vbYellow
    txtCobrancaDescricao.SetFocus
End Sub


Private Sub Object_Controller_Garantias(cType As Integer)
    If cType = 1 Then
        btnGarantiasIncluir.enabled = True
        btnGarantiasSalvar.enabled = False
        btnGarantiasEditar.enabled = False
        btnGarantiasExcluir.enabled = False
        
        Grid_Garantias.enabled = True
        
        Call Clean_Boxes_Garantia
    ElseIf cType = 2 Then  '
        btnGarantiasEditar.enabled = True
        btnGarantiasExcluir.enabled = True
        Grid_Garantias.enabled = True
    ElseIf cType = 3 Then 'Inserindo registro
        btnGarantiasIncluir.Caption = "&Cancela"
        txtGarantiaDescricao.enabled = True
        txtGarantiaDescricao.BackColor = &HFFFFFF
        btnGarantiasEditar.enabled = False

        Call insEdit_Controller_Garantias
        Call Clean_Boxes_Garantia
    ElseIf cType = 4 Then  'Edit a record
        txtGarantiaDescricao.enabled = True
        txtGarantiaDescricao.BackColor = &HFFFFFF

        btnGarantiasIncluir.enabled = False
        btnGarantiasEditar.Caption = "&Cancela"
        Call insEdit_Controller_Garantias
    ElseIf cType = 5 Then  'cancelando a inserção ou editção
        btnGarantiasIncluir.Caption = "&Incluir"
        btnGarantiasIncluir.enabled = True
        btnGarantiasSalvar.enabled = False
        btnGarantiasEditar.Caption = "&Editar"
        btnGarantiasEditar.enabled = True
        btnGarantiasExcluir.enabled = True
        
        Grid_Garantias.enabled = True
        
        txtGarantiaDescricao.BackColor = &H8000000F
        txtGarantiaDescricao.enabled = False
        
        Call Grid_Garantias_Click
    End If
End Sub

Private Sub Object_Controller_Cobrancas(cType As Integer)
    If cType = 1 Then
        btnCobrancasIncluir.enabled = True
        btnCobrancasSalvar.enabled = False
        btnCobrancasEditar.enabled = False
        btnCobrancasExcluir.enabled = False
        
        Grid_Cobrancas.enabled = True
        
        Call Clean_Boxes_Cobranca
    ElseIf cType = 2 Then  '
        btnCobrancasEditar.enabled = True
        btnCobrancasExcluir.enabled = True
        Grid_Cobrancas.enabled = True
    ElseIf cType = 3 Then 'Inserindo registro
        btnCobrancasIncluir.Caption = "&Cancela"
        txtCobrancaDescricao.enabled = True
        txtCobrancaDescricao.BackColor = &HFFFFFF
        btnCobrancasEditar.enabled = False

        Call insEdit_Controller_Cobrancas
        Call Clean_Boxes_Cobranca
    ElseIf cType = 4 Then  'Edit a record
        txtCobrancaDescricao.enabled = True
        txtCobrancaDescricao.BackColor = &HFFFFFF

        btnCobrancasIncluir.enabled = False
        btnCobrancasEditar.Caption = "&Cancela"
        Call insEdit_Controller_Cobrancas
        
    ElseIf cType = 5 Then  'cancelando a inserção ou editção
        btnCobrancasIncluir.Caption = "&Incluir"
        btnCobrancasIncluir.enabled = True
        btnCobrancasSalvar.enabled = False
        btnCobrancasEditar.Caption = "&Editar"
        btnCobrancasEditar.enabled = True
        btnCobrancasExcluir.enabled = True
        
        Grid_Cobrancas.enabled = True
        
        txtCobrancaDescricao.BackColor = &H8000000F
        txtCobrancaDescricao.enabled = False
        
        Call Grid_Cobrancas_Click
    ElseIf cType = 6 Then  'cancelando a inserção ou editção
        btnCobrancasExcluir.enabled = False
        btnCobrancasSalvar.enabled = False
        btnCobrancasEditar.enabled = False
        btnCobrancasIncluir.enabled = True
        
        Call Clean_Boxes_Cobranca
    End If
End Sub


Private Sub Configure_Grid_Garantias()
    With Grid_Garantias
        .Cols = 2
        .ColWidth(0) = 630
        .ColWidth(1) = 4280
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Descrição"
        .RowHeight(1) = 0
    End With
End Sub

Private Sub Configure_Grid_Cobrancas()
    With Grid_Cobrancas
        .Cols = 2
        .ColWidth(0) = 630
        .ColWidth(1) = 4280
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Descrição"
        .RowHeight(1) = 0
    End With
End Sub


Private Sub Show_Garantias(rs As ADODB.Recordset, Optional code As Integer)
    Dim selectedRow As Integer
    
    selectedRow = 2
    With Grid_Garantias
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("Codigo") & vbTab & IIf(IsNull(rs.Fields("Descricao")), "", rs.Fields("Descricao")))
            If code = rs.Fields("codigo") Then
                selectedRow = .Rows - 1
            End If
            rs.MoveNext
        Loop
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            .col = 0
            .ColSel = .Cols - 1
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Show_Cobrancas(rs As ADODB.Recordset, Optional code As Integer)
    Dim selectedRow As Integer
    
    selectedRow = 2
    With Grid_Cobrancas
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("Codigo") & vbTab & IIf(IsNull(rs.Fields("Descricao")), "", rs.Fields("Descricao")))
            If code = rs.Fields("codigo") Then
                selectedRow = .Rows - 1
            End If
            rs.MoveNext
        Loop
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            .col = 0
            .ColSel = .Cols - 1
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Sub Clean_Boxes_Garantia()
    txtGarantiaCodigo.Text = ""
    txtGarantiaDescricao.Text = ""
End Sub

Sub Clean_Boxes_Cobranca()
    txtCobrancaCodigo.Text = ""
    txtCobrancaDescricao.Text = ""
End Sub


Private Sub btnCobrancasEditar_Click()
    If btnCobrancasEditar.Caption = "&Editar" Then
        Call Object_Controller_Cobrancas(4)
    Else
        Call Object_Controller_Cobrancas(5)
    End If
End Sub

Private Sub btnCobrancasExcluir_Click()
    If MsgBox("Confirma a Exclusão da Cobrança " & txtCobrancaDescricao.Text & " ? ", vbYesNo, "Exclui Cobrança") = vbYes Then
        Call Return_Recordset(Return_SqlString(4, txtCobrancaCodigo.Text)).Delete
        Call Object_Controller_Cobrancas(6) 'aqui
        Call Show_Cobrancas(Return_Recordset(Return_SqlString(3)))
    End If
End Sub

Private Sub btnCobrancasIncluir_Click()
    If btnCobrancasIncluir.Caption = "&Incluir" Then
        Call Object_Controller_Cobrancas(3)
    Else
        Call Clean_Boxes_Cobranca
        Call Object_Controller_Cobrancas(5)
    End If

End Sub

Private Sub btnCobrancasSalvar_Click()
    Dim rs As ADODB.Recordset
    
    If txtCobrancaDescricao.Text = "" Then
        MsgBox "Digite a descrição da cobrança!"
        txtCobrancaDescricao.SetFocus
        Exit Sub
    End If
    
    If btnCobrancasIncluir.Caption = "&Cancela" Then
        Set rs = Return_Recordset(Return_SqlString(3))
        rs.AddNew
        txtCobrancaCodigo.Text = rs.Fields("codigo")
        
        rs.Fields("Descricao") = txtCobrancaDescricao.Text
    Else
        Set rs = Return_Recordset(Return_SqlString(4, txtCobrancaCodigo.Text))
    End If
    
    Call Record_Record_Cobrancas(rs)

    rs.Update

    rs.Close
    Set rs = Nothing
    
    Call Object_Controller_Cobrancas(5)
    Call Show_Cobrancas(Return_Recordset(Return_SqlString(3)))

    btnCobrancasEditar.enabled = False
'    Form_Customer.Refresh_Grid_Equipment

End Sub

Private Sub btnGarantiasEditar_Click()
    If btnGarantiasEditar.Caption = "&Editar" Then
        Call Object_Controller_Garantias(4)
    Else
        Call Object_Controller_Garantias(5)
    End If
End Sub

Private Sub btnGarantiasExcluir_Click()
    If MsgBox("Confirma a Exclusão da Garantia " & txtGarantiaDescricao.Text & " ? ", vbYesNo, "Exclui Garantia") = vbYes Then
        Call Return_Recordset(Return_SqlString(2, txtGarantiaCodigo.Text)).Delete
        Call Object_Controller_Garantias(1)
        Call Show_Garantias(Return_Recordset(Return_SqlString(1)))
    End If
End Sub

Private Sub btnGarantiasIncluir_Click()
    If btnGarantiasIncluir.Caption = "&Incluir" Then
        Call Object_Controller_Garantias(3)
    Else
        Call Clean_Boxes_Garantia
        Call Object_Controller_Garantias(5)
    End If
End Sub

Private Sub Record_Record_Garantias(rs As ADODB.Recordset)
    rs.Fields("Descricao").value = UCase(txtGarantiaDescricao.Text)
End Sub

Private Sub Record_Record_Cobrancas(rs As ADODB.Recordset)
    rs.Fields("Descricao").value = UCase(txtCobrancaDescricao.Text)
End Sub

Private Sub btnGarantiasSalvar_Click()
    Dim rs As ADODB.Recordset
    
    If txtGarantiaDescricao.Text = "" Then
        MsgBox "Digite a descrição da garantia!"
        txtGarantiaDescricao.SetFocus
        Exit Sub
    End If
    
    If btnGarantiasIncluir.Caption = "&Cancela" Then
        Set rs = Return_Recordset(Return_SqlString(1))
        rs.AddNew
        txtGarantiaCodigo.Text = rs.Fields("codigo")
        
        rs.Fields("Descricao") = txtGarantiaDescricao.Text
    Else
        Set rs = Return_Recordset(Return_SqlString(2, txtGarantiaCodigo.Text))
    End If
    
    Call Record_Record_Garantias(rs)

    rs.Update

    rs.Close
    Set rs = Nothing
    
    Call Object_Controller_Garantias(5)
    Call Show_Garantias(Return_Recordset(Return_SqlString(1)))
    
    btnGarantiasEditar.enabled = False
'    Form_Customer.Refresh_Grid_Equipment
End Sub

Private Sub Form_Activate()
    txtGarantiaCodigo.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    txtGarantiaDescricao.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    txtCobrancaCodigo.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    txtCobrancaDescricao.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    
    Grid_Garantias.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    Grid_Cobrancas.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
End Sub

Private Sub Form_Load()
   Set conn = New ADODB.Connection
   conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"

   Call Object_Controller_Garantias(1)
   Call Object_Controller_Cobrancas(1)
   Call Configure_Grid_Garantias
   Call Configure_Grid_Cobrancas
   Call Show_Garantias(Return_Recordset(Return_SqlString(1)))
   Call Show_Cobrancas(Return_Recordset(Return_SqlString(3)))
   
End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String, Optional param3 As String) As String
    
    If queryType = 1 Then
        Return_SqlString = "SELECT * FROM Garantias"
    ElseIf queryType = 2 Then
        Return_SqlString = "SELECT * FROM Garantias Where Codigo = " & param1
    ElseIf queryType = 3 Then
        Return_SqlString = "SELECT * FROM Cobrancas"
    ElseIf queryType = 4 Then
        Return_SqlString = "SELECT * FROM Cobrancas Where Codigo = " & param1
    ElseIf queryType = 5 Then
        Return_SqlString = "SELECT * FROM Cobrancas Where Codigo = " & param1
    End If
    
End Function

Sub Show_Record(rs As ADODB.Recordset)

    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    txtGarantiaCodigo.Text = rs.Fields("Codigo")
    txtGarantiaDescricao.Text = UCase(rs.Fields("Descricao"))
    
    rs.Close
    Set rs = Nothing
    
    Call Object_Controller_Garantias(2)
End Sub

Sub Show_Record_Cobrancas(rs As ADODB.Recordset)

    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    txtCobrancaCodigo.Text = rs.Fields("Codigo")
    txtCobrancaDescricao.Text = UCase(rs.Fields("Descricao"))
    
    rs.Close
    Set rs = Nothing
    
    Call Object_Controller_Cobrancas(2)
End Sub

Private Sub Grid_Cobrancas_Click()
    With Grid_Cobrancas
        If .RowSel = 1 Then
            Exit Sub
        End If
        Call Show_Record_Cobrancas(Return_Recordset(Return_SqlString(4, .TextMatrix(.RowSel, 0))))
    End With

End Sub

Private Sub Grid_Garantias_Click()
    
    With Grid_Garantias
        If .RowSel = 1 Then
            Exit Sub
        End If
        Call Show_Record(Return_Recordset(Return_SqlString(2, .TextMatrix(.RowSel, 0))))
    End With

End Sub
