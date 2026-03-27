VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Begin VB.Form frmEscolhaValorFornecedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "xSOL Aquecedores - Material com mais de um fornecedor"
   ClientHeight    =   4815
   ClientLeft      =   7710
   ClientTop       =   1140
   ClientWidth     =   7335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmEscolhaValorFornecedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   330
      Left            =   6705
      TabIndex        =   2
      Top             =   4350
      Width           =   510
   End
   Begin MSFlexGridLib.MSFlexGrid gridFornecedores 
      Height          =   3090
      Left            =   165
      TabIndex        =   1
      Top             =   1140
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   5450
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin VB.Label lblCodMaterial 
      Caption         =   "..."
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   4305
      Width           =   330
   End
   Begin VB.Label lblSobre 
      Alignment       =   2  'Center
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
      Height          =   855
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   7065
   End
End
Attribute VB_Name = "frmEscolhaValorFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOk_Click()
    Form_Estoque.entradaValor = gridFornecedores.TextMatrix(gridFornecedores.RowSel, 2)
    Form_Estoque.saidaValor = gridFornecedores.TextMatrix(gridFornecedores.RowSel, 4)
    
    Unload Me
End Sub

Private Sub Configure_Grid_Fornecedores()
    With gridFornecedores
        .Cols = 5
        .FixedCols = 0
        .ColWidth(0) = 900
        .ColWidth(1) = 3160
        .ColWidth(2) = 1000
        .ColWidth(3) = 800
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
        .TextMatrix(0, 0) = "Cod.Forn."
        .TextMatrix(0, 1) = "Nome Fornecedor"
        .TextMatrix(0, 2) = "R$ Compra"
        .TextMatrix(0, 3) = "% Lucro"
        .TextMatrix(0, 4) = "R$ Venda"
        .RowHeight(1) = 0
    End With
End Sub

Private Sub Fill_Grid_Fornecedores(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    selectedRow = 2
    With gridFornecedores
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("CodFornecedor") & vbTab & rs.Fields("NomeCli") & vbTab & rs.Fields("ValorCompra") & vbTab & rs.Fields("PorcentagemLucro") & vbTab & rs.Fields("ValorVenda"))
            'If code = rs.Fields("codFornecedor") Then
            '    selectedRow = .Rows - 1
            'End If
            
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
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String) As String
    Dim whereClause As String
    
    If queryType = 1 Then
        If param1 = "..." Then
            param1 = "0"
        End If
        Return_SqlString = "Select main.CodFornecedor, f.NomeCli, main.ValorCompra, main.PorcentagemLucro, main.ValorVenda From MaterialFornecedor main Inner Join Fornecedor f ON f.codigo = main.codFornecedor Where main.codMaterial = " & param1
    ElseIf queryType = 2 Then
        If param1 = "" Then
            param1 = 0
        End If
    
        Return_SqlString = "Select * From MaterialFornecedor Where CodFornecedor = " & param1
    End If
    
End Function


Private Sub Form_Activate()
    Call Configure_Grid_Fornecedores
    Call Fill_Grid_Fornecedores(Return_Recordset(Return_SqlString(1, Form_Estoque.lblCodMaterial.Caption)))
    
    gridFornecedores.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito

End Sub

Sub Show_Record_Detail(rs As ADODB.Recordset)

    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    lblCodMaterial.Caption = rs.Fields("codFornecedor")
        
    rs.Close
    Set rs = Nothing
End Sub


Private Sub gridFornecedores_Click()
    Call Show_Record_Detail(Return_Recordset(Return_SqlString(2, gridFornecedores.TextMatrix(gridFornecedores.RowSel, 0))))
End Sub
