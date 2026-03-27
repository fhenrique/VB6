VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDiasIndisponiveis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XSol Aquecedores - Dias que o técnico estará indisponível"
   ClientHeight    =   3705
   ClientLeft      =   4485
   ClientTop       =   3240
   ClientWidth     =   11070
   Icon            =   "frmDiasIndisponiveis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMes 
      Height          =   360
      Left            =   960
      Locked          =   -1  'True
      MaxLength       =   240
      TabIndex        =   9
      ToolTipText     =   "Separe os dias por vírgula, ex: 01,10,25  .No caso de indisponível no mês inteiro, informe apenas a letra M"
      Top             =   840
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   165
      TabIndex        =   7
      Top             =   2850
      Width           =   5535
      Begin VB.CommandButton btnSalvar 
         Caption         =   "&Salvar"
         Height          =   330
         Left            =   255
         TabIndex        =   8
         Top             =   240
         Width           =   4995
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3555
      Left            =   5835
      TabIndex        =   3
      Top             =   90
      Width           =   5130
      Begin MSFlexGridLib.MSFlexGrid gridDiasIndisponiveis 
         Height          =   3345
         Left            =   105
         TabIndex        =   4
         Top             =   165
         Width           =   4950
         _ExtentX        =   8731
         _ExtentY        =   5900
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         SelectionMode   =   1
      End
   End
   Begin VB.TextBox txtDias 
      Height          =   360
      Left            =   945
      MaxLength       =   240
      TabIndex        =   0
      ToolTipText     =   "Separe os dias por vírgula, ex: 01,10,25  .No caso de indisponível no mês inteiro, informe apenas a letra M"
      Top             =   1470
      Width           =   4740
   End
   Begin VB.Label Label3 
      Caption         =   "Caso necessário, separe os dias por vírgula, ex: 01,02,03"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   915
      TabIndex        =   10
      Top             =   1905
      Width           =   4380
   End
   Begin VB.Label Label2 
      Caption         =   "Dias:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   420
      TabIndex        =   6
      Top             =   1590
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "Mês:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   465
      TabIndex        =   5
      Top             =   960
      Width           =   450
   End
   Begin VB.Label lblNomeTecnico 
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   975
      TabIndex        =   2
      Top             =   315
      Width           =   4620
   End
   Begin VB.Label Label8 
      Caption         =   "Técnico:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   135
      TabIndex        =   1
      Top             =   300
      Width           =   945
   End
End
Attribute VB_Name = "frmDiasIndisponiveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nomeTecnico As String
Public codTecnico As String
Private bEditando As Boolean


Private Sub btnSalvar_Click()
    
    Dim rs As ADODB.Recordset

    Set rs = Return_Recordset(Return_SqlString(3, codTecnico, gridDiasIndisponiveis.TextMatrix(gridDiasIndisponiveis.RowSel, 0)))
    
    
    rs.Fields("dias") = txtDias.Text
    
    
    
    rs.Update
    rs.Close
    Set rs = Nothing
       
    MsgBox "A agenda do técnico " & nomeTecnico & " foi atualizada!", vbOKOnly, "XSol Aquecedores"

    Call carregaGridDiasIndisponiveis(Return_Recordset(Return_SqlString(1, codTecnico)))
End Sub





Private Sub Form_Load()
    
    Me.Caption = "XSol Aquecedores - Dias que o técnico " & nomeTecnico & " estará indisponível"
    lblNomeTecnico.Caption = nomeTecnico
    
    Call ConfigureGridDiasIndisponiveis
    Call carregaGridDiasIndisponiveis(Return_Recordset(Return_SqlString(1, codTecnico)))
    
End Sub

Private Sub ConfigureGridDiasIndisponiveis()
    With gridDiasIndisponiveis
        .Cols = 2
        .ColWidth(0) = 700  'Mes
        .ColWidth(1) = 4000 'Dias
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Mes"
        .TextMatrix(0, 1) = "Dia(s)"
        .RowHeight(1) = 0
    End With
End Sub

Private Sub carregaGridDiasIndisponiveis(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim col As Integer
    
    selectedRow = 2
    With gridDiasIndisponiveis
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (IIf(IsNull(rs.Fields("MES")), "", rs.Fields("MES")) & vbTab & IIf(IsNull(rs.Fields("DIAS")), "", rs.Fields("DIAS")))
            'If code = rs.Fields("CODIGO") Then
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
        Else
        End If
        
    End With
    rs.Close
    Set rs = Nothing
End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String, Optional param3 As String, Optional param4 As String) As String
    Dim whereClause As String
    
    If queryType = 1 Then
        Return_SqlString = "Select Codigo, Mes, Dias from tecnicos_agenda where codigo = '" & param1 & "'"
    ElseIf queryType = 2 Then
        Return_SqlString = "Select Codigo, Mes, Dias from tecnicos_agenda where codigo = '" & param1 & "' and mes = '" & param2 & "'"
    ElseIf queryType = 3 Then
        Return_SqlString = "Select Dias from tecnicos_agenda where codigo = '" & param1 & "' and mes = '" & param2 & "'"
    End If


End Function

Private Sub Form_Terminate()
    Call carregaGridDiasIndisponiveis(Return_Recordset(Return_SqlString(1, codTecnico)))
End Sub

Private Sub gridDiasIndisponiveis_Click()
    Call carregaInfo(Return_Recordset(Return_SqlString(2, codTecnico, gridDiasIndisponiveis.TextMatrix(gridDiasIndisponiveis.RowSel, 0))))
End Sub

Private Sub carregaInfo(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If
        
    txtMes.Text = retornaNomeMes(rs.Fields("MES"))
    
    If IsNull(rs.Fields("DIAS")) Then
        txtDias.Text = ""
    Else
        txtDias.Text = rs.Fields("DIAS")
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Function retornaNomeMes(sigla As String) As String
    If sigla = "JAN" Then
        retornaNomeMes = "JANEIRO"
    ElseIf sigla = "FEV" Then
        retornaNomeMes = "FEVEREIRO"
    ElseIf sigla = "MAR" Then
        retornaNomeMes = "MARÇO"
    ElseIf sigla = "ABR" Then
        retornaNomeMes = "ABRIL"
    ElseIf sigla = "MAI" Then
        retornaNomeMes = "MAIO"
    ElseIf sigla = "JUN" Then
        retornaNomeMes = "JUNHO"
    ElseIf sigla = "JUL" Then
        retornaNomeMes = "JULHO"
    ElseIf sigla = "AGO" Then
        retornaNomeMes = "AGOSTO"
    ElseIf sigla = "SET" Then
        retornaNomeMes = "SETEMBRO"
    ElseIf sigla = "OUT" Then
        retornaNomeMes = "OUTUBRO"
    ElseIf sigla = "NOV" Then
        retornaNomeMes = "NOVEMBRO"
    ElseIf sigla = "DEZ" Then
        retornaNomeMes = "DEZEMBRO"
    End If
End Function

Private Sub txtDias_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim conteudo As String
    Dim doublecoma As String
        
    If txtDias.Text <> "" And (Len(txtDias.Text) > 2) Then
        doublecoma = Mid(txtDias.Text, Len(txtDias.Text) - 1, 2)
    End If
    
    If txtDias.Text = "," Then
        txtDias.Text = ""
    End If
    
    If txtDias.Text <> "" Then
    
        conteudo = Mid(txtDias.Text, 1, Len(txtDias.Text) - 1)
    
        If doublecoma = ",," Then
            txtDias.Text = conteudo
            Exit Sub
        End If
        
        
        If (KeyCode <> 8) And (KeyCode <> 188) And (KeyCode <> 35) And (KeyCode <> 36) Then
            If Not IsNumeric(Mid(txtDias.Text, Len(txtDias.Text), 1)) And (KeyCode <> 188) And (KeyCode <> 35) And (KeyCode <> 36) Then
                'MsgBox Mid(txtDias.Text, Len(txtDias.Text), 1)
                'MsgBox "não é número"
                txtDias.Text = conteudo
            End If
        End If
        
    End If
End Sub
