VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFiltroRelatorio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtros para o Relatório"
   ClientHeight    =   1620
   ClientLeft      =   8535
   ClientTop       =   4395
   ClientWidth     =   4110
   Icon            =   "frmFiltroRelatorio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   2490
      TabIndex        =   2
      Top             =   1215
      Width           =   825
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "&OK"
      Height          =   345
      Left            =   3450
      TabIndex        =   1
      Top             =   1215
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   1200
      Left            =   15
      TabIndex        =   0
      Top             =   -60
      Width           =   4065
      Begin VB.ComboBox cboTecnicos 
         Height          =   315
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   825
         Width           =   3885
      End
      Begin MSComCtl2.DTPicker dtpDataDe 
         DataField       =   "'"
         Height          =   315
         Left            =   990
         TabIndex        =   5
         Top             =   255
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   195756033
         CurrentDate     =   44068
      End
      Begin MSComCtl2.DTPicker dtpDataAte 
         Height          =   315
         Left            =   2685
         TabIndex        =   6
         Top             =   255
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   233570305
         CurrentDate     =   44068
      End
      Begin VB.Label lblTecnico 
         AutoSize        =   -1  'True
         Caption         =   "Técnico:"
         Height          =   195
         Left            =   60
         TabIndex        =   8
         Top             =   630
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "até"
         Height          =   195
         Left            =   2340
         TabIndex        =   4
         Top             =   315
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Período de:"
         DataField       =   "'"
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   315
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmFiltroRelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tipoRelatorio As Integer
Public pathDataBase As String
Public Estoque_DBname As String
Private conn As ADODB.Connection

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    relEntradaEquipamento.DAODataControl1.DatabaseName = pathDataBase
    relSaidaEquipamento.DAODataControl1.DatabaseName = pathDataBase
    relRetiradaEquipamento.DAODataControl1.DatabaseName = pathDataBase
    
    If tipoRelatorio = 1 Then
        relEntradaEquipamento.lblTitulo.Caption = "ENTRADA DE EQUIPAMENTOS"
        relEntradaEquipamento.DAODataControl1.RecordSource = "Select Produto, Dt_Ent, Qtd_Ent, Valor_Ent, Qtd_Ent * Valor_Ent as Total From Estoque Where Dt_Ent >= #" & dtpDataDe.value & "# And Dt_Ent <= #" & dtpDataAte.value & "# "
        relEntradaEquipamento.lblDataDe.Caption = dtpDataDe.value
        relEntradaEquipamento.lblDataAte.Caption = dtpDataAte.value
        
        relEntradaEquipamento.Show (1)
    ElseIf tipoRelatorio = 2 Then
        relSaidaEquipamento.DAODataControl1.RecordSource = "Select Produto, Dt_Saida, Qtd_Saida, Valor_Saida, Qtd_Saida * Valor_Saida as Total From Estoque Where Dt_Saida >= #" & dtpDataDe.value & "# And Dt_Saida <= #" & dtpDataAte.value & "# "
        relSaidaEquipamento.lblDataDe.Caption = dtpDataDe.value
        relSaidaEquipamento.lblDataAte.Caption = dtpDataAte.value
        
        relSaidaEquipamento.Show (1)
    ElseIf tipoRelatorio = 3 Then
        If cboTecnicos.Text = "" Then
            MsgBox "Selecione um técnico."
            Exit Sub
        End If
        relRetiradaEquipamento.lblTecnico.Caption = "Técnico: " & cboTecnicos.Text
        relRetiradaEquipamento.DAODataControl1.RecordSource = "Select Produto, Dt_Saida, Qtd_Saida, Valor_Saida, Qtd_Saida * Valor_Saida as Total From Estoque Where Dt_Saida >= #" & dtpDataDe.value & "# And Dt_Saida <= #" & dtpDataAte.value & "# And Nome_Tecnico = '" & cboTecnicos.Text & "'"
        relRetiradaEquipamento.Show (1)
    End If
End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As String) As String
    Dim whereClause As String
    
    If queryType = 1 Then
        Return_SqlString = "SELECT Distinct(Nome) FROM Tecnico"
    End If
End Function

Public Function Return_Recordset(strSql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    Call EstoqueArquivos_Dados
    
    Set conn = Nothing
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & Estoque_DBname & ";"
    
    rs.Open strSql, conn, adOpenStatic, adLockOptimistic
    
    Set Return_Recordset = rs
End Function

Sub EstoqueArquivos_Dados()
    Dim ConfigFile As String
    Dim Path As String
    Dim FileName As String
    
    ConfigFile = App.Path & "\CONFIG.INI"
    
    Path = RTrim(GetINISetting("DATABASE", "PATH", ConfigFile))
    If Right(Path, 1) <> "\" Then
       Path = Path + "\"
    End If
    FileName = GetINISetting("DATABASE", "FILENAME", ConfigFile)
    
    Image_Folder = GetINISetting("IMAGE", "FOLDER", ConfigFile)
    Estoque_DBname = Path + FileName
End Sub

Private Sub carregaTecnicos(rs As ADODB.Recordset, Optional code As String)
    With cboTecnicos
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("Nome"))
            rs.MoveNext
        Loop
    End With
    
    rs.Close
    Set rs = Nothing
    
    cboTecnicos.ListIndex = 0
    
End Sub

Private Sub Form_Activate()
    If (MDIForm_Xsol.acessibilidade_TextoNegrito = True) Or (MDIForm_Xsol.acessibilidade_TextoNegrito = False) Then
        cboTecnicos.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        dtpDataDe.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        dtpDataAte.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    End If

End Sub

Private Sub Form_Load()
    dtpDataDe.value = Format(DateTime.Now, "Short Date")
    dtpDataAte.value = Format(DateTime.Now + 30, "Short Date")
    relEntradaEquipamento.DAODataControl1.DatabaseName = pathDataBase
    relSaidaEquipamento.DAODataControl1.DatabaseName = pathDataBase
    relRetiradaEquipamento.DAODataControl1.DatabaseName = pathDataBase
    
    cboTecnicos.Visible = tipoRelatorio = 3
    lblTecnico.Visible = tipoRelatorio = 3
    
    Call carregaTecnicos(Return_Recordset(Return_SqlString(1)))
End Sub
