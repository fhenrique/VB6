VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsultaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa de cliente"
   ClientHeight    =   4335
   ClientLeft      =   7890
   ClientTop       =   2370
   ClientWidth     =   6525
   Icon            =   "frmConsultaCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNome 
      Height          =   360
      Left            =   1380
      TabIndex        =   2
      Top             =   165
      Width           =   4950
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   420
      Left            =   5850
      TabIndex        =   0
      Top             =   3840
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid gridClientes 
      Height          =   3120
      Left            =   120
      TabIndex        =   3
      Top             =   630
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   5503
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      SelectionMode   =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Selecione o cliente no grid acima, depois clique no botão OK"
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   3780
      Width           =   4305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Busca pr nome:"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   285
      Width           =   1110
   End
End
Attribute VB_Name = "frmConsultaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strTpConsulta As String

Public intqtdLetras As Integer
Public conn As ADODB.Connection
Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String, Optional param3 As String) As String
    If queryType = 1 Then
        Return_SqlString = "Select top 20 codigo, nome_cli From Clientes Where Nome_Cli like '" & param1 & "%'"
    End If
End Function

Private Sub carregaGridClientes(rs As ADODB.Recordset)
    Dim selectedRow As Integer
    Dim col As Integer
    
    selectedRow = 2
    With gridClientes
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("CODIGO") & vbTab & rs.Fields("nome_cli"))
                .row = .Rows - 1
                For col = .FixedCols To .Cols - 1
                    .col = col
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

Private Sub btnOk_Click()
    If gridClientes.Rows > 2 Then
        frmRecibos.strReciboClienteCodigo = gridClientes.TextMatrix(gridClientes.RowSel, 0)
        frmRecibos.strReciboClienteNome = gridClientes.TextMatrix(gridClientes.RowSel, 1)
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    txtNome.SetFocus
End Sub

Private Sub Form_Load()
    Call ConfigureGridClientes
    
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
    
End Sub

Private Sub ConfigureGridClientes()
    With gridClientes
        .Cols = 2
        .ColWidth(0) = 800  'CODIGO
        .ColWidth(1) = 5000 'NOME
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Nome"
        .RowHeight(1) = 0
    End With
End Sub

Private Sub txtNome_KeyUp(KeyCode As Integer, Shift As Integer)
    intqtdLetras = Len(txtNome.Text)
    
    If intqtdLetras >= 4 Then
        If (KeyCode <> 8) And (KeyCode <> 192) And (KeyCode <> 32) And (KeyCode <> 56) And (KeyCode <> 53) And (KeyCode <> 46) And (KeyCode <> 27) Then
            'chama a função para buscar o cliente
            Call carregaGridClientes(Return_Recordset(Return_SqlString(1, txtNome.Text)))
        End If
    End If

End Sub

Private Function Return_Recordset(strSql As String) As ADODB.Recordset
    Dim rsx As New ADODB.Recordset
    Dim qtd As Integer
    
    Call Arquivos_Dados
    
    rsx.Open strSql, conn, adOpenStatic, adLockOptimistic
    qtd = rsx.RecordCount
    
    Set Return_Recordset = rsx
End Function

