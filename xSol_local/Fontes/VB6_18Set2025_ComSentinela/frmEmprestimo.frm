VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmprestimo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XSOL Aquecedores - Empréstimo de Ferramenta"
   ClientHeight    =   6135
   ClientLeft      =   1545
   ClientTop       =   3000
   ClientWidth     =   16155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   16155
   Begin VB.Frame Frame4 
      Caption         =   "Empréstimos já realizados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6030
      Left            =   7110
      TabIndex        =   11
      Top             =   0
      Width           =   8985
      Begin MSFlexGridLib.MSFlexGrid gridEmprestimos 
         Height          =   5685
         Left            =   75
         TabIndex        =   12
         Top             =   255
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   10028
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6030
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   6945
      Begin VB.CheckBox chkDevolvido 
         Caption         =   "Empréstimo devolvido"
         Enabled         =   0   'False
         Height          =   405
         Left            =   345
         TabIndex        =   14
         Top             =   4035
         Width           =   2775
      End
      Begin VB.CommandButton btnExcluir 
         BackColor       =   &H008080FF&
         Caption         =   "Excluir"
         Height          =   360
         Left            =   4260
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5280
         Width           =   1100
      End
      Begin VB.TextBox txtBeneficiario 
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1845
         Width           =   5025
      End
      Begin VB.TextBox txtFinalidade 
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2580
         Width           =   5010
      End
      Begin VB.CommandButton btnIncluir 
         BackColor       =   &H0080FF80&
         Caption         =   "Incluir"
         Height          =   360
         Left            =   315
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5280
         Width           =   1100
      End
      Begin VB.CommandButton btnSalvar 
         Caption         =   "Salvar"
         Enabled         =   0   'False
         Height          =   360
         Left            =   2940
         TabIndex        =   2
         Top             =   5295
         Width           =   1100
      End
      Begin VB.CommandButton btnEditar 
         Caption         =   "Editar"
         Height          =   360
         Left            =   1620
         TabIndex        =   1
         Top             =   5295
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dtpDataPrevista 
         Height          =   315
         Left            =   1725
         TabIndex        =   13
         Top             =   3300
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   275382273
         CurrentDate     =   45930
      End
      Begin MSComCtl2.DTPicker dtpEmprestimo 
         Height          =   315
         Left            =   1740
         TabIndex        =   15
         Top             =   1065
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   275382273
         CurrentDate     =   45930
      End
      Begin VB.Label lblFerramenta 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1740
         TabIndex        =   17
         Top             =   255
         Width           =   4830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ferramenta:"
         Height          =   195
         Left            =   750
         TabIndex        =   16
         Top             =   255
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data do empréstimo:"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beneficiário:"
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   1965
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finalidade:"
         Height          =   195
         Left            =   825
         TabIndex        =   7
         Top             =   2670
         Width           =   765
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Data prevista para devolução:"
         Height          =   465
         Left            =   315
         TabIndex        =   6
         Top             =   3240
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmEmprestimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEditando As Boolean

Private Sub btnEditar_Click()
 
    If gridEmprestimos.Rows = 2 Then
        MsgBox "Não há empréstimo há editar!"
        Exit Sub
    End If
 
 
 bEditando = True
 
 btnIncluir.enabled = False
 btnEditar.enabled = False
 btnSalvar.enabled = True
 
 btnExcluir.enabled = True 'vai ser cancelar
 btnExcluir.Caption = "Cancelar"
 
gridEmprestimos.enabled = False
 
habilitaCampos (True)
 
End Sub

Private Sub btnExcluir_Click()
 
    If gridEmprestimos.Rows = 2 Then
        MsgBox "Não há empréstimo há excluir!"
        Exit Sub
    End If
 
 If btnExcluir.Caption = "Excluir" Then
    If MsgBox("Confirma exclusão?", vbYesNo, "Confirmação") = vbYes Then
       'faz a exclusao ...
       
       btnIncluir.enabled = True
       btnEditar.enabled = True
       btnSalvar.enabled = False
       btnExcluir.enabled = True
       
       habilitaCampos (False)
       
    Else
       btnIncluir.enabled = False
       btnEditar.enabled = False
       btnSalvar.enabled = True
       btnExcluir.enabled = False
    End If
 ElseIf btnExcluir.Caption = "Cancelar" Then
       btnIncluir.enabled = True
       btnEditar.enabled = True
       btnSalvar.enabled = False
       
       btnExcluir.enabled = True
       btnExcluir.Caption = "Excluir"
       
       habilitaCampos (False)
       
 End If
 
 
 

 
 
End Sub

Private Sub btnIncluir_Click()
  
 bEditando = False
 
 btnIncluir.enabled = False
 btnEditar.enabled = False
 btnSalvar.enabled = True
 btnExcluir.enabled = True 'pq vai ser cancelar
 
 btnExcluir.Caption = "Cancelar"
 
    Call limpaTela

    dtpEmprestimo.value = Date
    gridEmprestimos.enabled = False
    
habilitaCampos (True)
 
End Sub

Private Sub btnSalvar_Click()
 btnIncluir.enabled = True
 btnEditar.enabled = True
 btnSalvar.enabled = False
 btnExcluir.enabled = True
 
 btnExcluir.Caption = "Excluir"
 
    habilitaCampos (False)
 
    If RTrim(LTrim(txtBeneficiario.Text)) = "" Then
        MsgBox "Informe o nome do beneficiário!", vbInformation, "Validação"
        txtBeneficiario.SetFocus
        Exit Sub
    End If
    
    Call gravaEmprestimo
    
    telaVazia
    
    gridEmprestimos.enabled = True
    
    gridEmprestimos_SelChange

End Sub



Private Sub Form_Load()
    
    Call ConfigureGrid
    Call carregaGridEmprestimos(Return_Recordset(Return_SqlString(1)))
    
    

End Sub

Private Sub ConfigureGrid()
    With gridEmprestimos
        .Cols = 7
        .ColWidth(0) = 0 'Codigo (invisível")
        .ColWidth(1) = 0 'codigo ferramenta (invisivel)
        .ColWidth(2) = 1050 'Data
        .ColWidth(3) = 2300 'Beneficiário(a)
        .ColWidth(4) = 3100 'Finalidade
        .ColWidth(5) = 1200 'Dt. devolucao
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
        .TextMatrix(0, 2) = "Data"
        .TextMatrix(0, 3) = "Beneficiário"
        .TextMatrix(0, 4) = "Finalidade"
        .TextMatrix(0, 5) = "Dt. Devolução"
        .TextMatrix(0, 6) = "Devolvido?"
        
        .RowHeight(1) = 0
    End With
End Sub

Private Sub limpaTela()
    
    txtBeneficiario.Text = ""
    txtFinalidade.Text = ""
    dtpDataPrevista.value = Date + 3
    
    chkDevolvido.value = 0
    
    txtBeneficiario.SetFocus
End Sub

Private Sub telaVazia()
    
    txtBeneficiario.Text = ""
    txtFinalidade.Text = ""
    dtpDataPrevista.value = Date
    
    chkDevolvido.value = 0
    
End Sub

Private Sub habilitaCampos(habilitar As Boolean)

    dtpEmprestimo.enabled = habilitar
    txtBeneficiario.Locked = Not (habilitar)
    txtFinalidade.Locked = Not (habilitar)
    dtpDataPrevista.enabled = habilitar
    chkDevolvido.enabled = habilitar
    
    
End Sub

Private Sub carregaGridEmprestimos(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    
    'Dim bkColor As Long
    Dim col As Integer
    
    Dim bDevolucao As Boolean
    Dim strDevolvido As String
    
    
    
    dblValorVerde = 0
    
    selectedRow = 2
    With gridEmprestimos
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            'If (IsNull(code)) Or (code = "") Then
                '.AddItem (rs.Fields("codigo") & vbTab & rs.Fields("data") & vbTab & rs.Fields("cliente") & vbTab & rs.Fields("valor"))
            'Else
                
                If rs.Fields("EMPRESTIMO_DEVOLVIDO") = "0" Then
                    bDevolucao = False
                Else
                    bDevolucao = True
                End If
                
                
                If bDevolucao = False Then
                    strDevolvido = "N"
                Else
                    strDevolvido = "S"
                End If
                
                .AddItem (rs.Fields("CODIGO") & vbTab & rs.Fields("CODIGO") & vbTab & rs.Fields("EMPRESTIMO_DTRETIRADA") & vbTab & rs.Fields("EMPRESTIMO_BENEFICIARIO") & vbTab & rs.Fields("EMPRESTIMO_FINALIDADE") & vbTab & rs.Fields("EMPRESTIMO_DTDEVOLUCAO") & vbTab & strDevolvido)
            'End If
            
            'If code = rs.Fields("CODIGO") Then
            '    selectedRow = .Rows - 1
            'End If
            
                .row = .Rows - 1
                For col = .FixedCols To .Cols - 1
                    .col = col
                    
                    If rs.Fields("EMPRESTIMO_DEVOLVIDO") = "1" Then
                        .CellBackColor = vbGreen
                    Else
                        .CellBackColor = vbRed
                    End If
                    
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


Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String, Optional param3 As String) As String
    If queryType = 1 Then
        Return_SqlString = "Select * from FERRAMENTAS_EMPRESTIMO"
    ElseIf queryType = 2 Then
        Return_SqlString = "Select * from FERRAMENTAS_EMPRESTIMO where codigo = " & param1
    End If
End Function

Private Sub gridEmprestimos_SelChange()
    
    If gridEmprestimos.Rows > 2 Then
    'If gridRecibos.TextMatrix(2, 0) <> "" And gridOrcamentos.TextMatrix(gridOrcamentos.RowSel, 0) <> "" Then
        Call carregaInfoEmprestimo(Return_Recordset(Return_SqlString(2, gridEmprestimos.TextMatrix(gridEmprestimos.RowSel, 0))))
    'End If
    End If

End Sub

Private Sub carregaInfoEmprestimo(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    If gridEmprestimos.TextMatrix(gridEmprestimos.RowSel, 0) = "" Then
        Exit Sub
    End If
    
    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        dtpEmprestimo.value = rs.Fields("EMPRESTIMO_DTRETIRADA")
        txtBeneficiario.Text = rs.Fields("EMPRESTIMO_BENEFICIARIO")
        txtFinalidade.Text = rs.Fields("EMPRESTIMO_FINALIDADE")
        
        dtpDataPrevista.value = rs.Fields("EMPRESTIMO_DTDEVOLUCAO")
        
        chkDevolvido.value = rs.Fields("EMPRESTIMO_DEVOLVIDO")
        
        
    End If
        
    rs.Close
    Set rs = Nothing
    
    
End Sub

Private Sub gravaEmprestimo()
    Dim rs As ADODB.Recordset
    
    If bEditando = False Then 'está inserindo
        Set rs = Return_Recordset(Return_SqlString(1))
        rs.AddNew
    Else 'está editando
        Set rs = Return_Recordset(Return_SqlString(2, lblCodRecibo.Caption))
    End If
    
    'CODIGO, DATA, COD_CLIENTE, VALOR
    'rs.Fields("CODIGO") = lblCodRecibo.Caption
    rs.Fields("EMPRESTIMO_DTRETIRADA") = dtpEmprestimo.value
    rs.Fields("EMPRESTIMO_BENEFICIARIO") = txtBeneficiario.Text
    rs.Fields("EMPRESTIMO_FINALIDADE") = txtFinalidade.Text
    rs.Fields("EMPRESTIMO_DTDEVOLUCAO") = dtpDataPrevista.value
    rs.Fields("EMPRESTIMO_DEVOLVIDO") = chkDevolvido.value
    
    rs.Update
    rs.Close
    Set rs = Nothing

    Call carregaGridEmprestimos(Return_Recordset(Return_SqlString(1)))

    MsgBox "Empréstimo gravado", vbInformation, "Informação"
    

End Sub

