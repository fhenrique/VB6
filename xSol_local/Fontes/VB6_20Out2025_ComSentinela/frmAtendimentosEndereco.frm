VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAtendimentosEndereco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XSOL Aquecedores - Atendimentos do Endereço"
   ClientHeight    =   10275
   ClientLeft      =   6870
   ClientTop       =   480
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Atendimentos já realizados no endereço acima:"
      Height          =   9090
      Left            =   90
      TabIndex        =   11
      Top             =   1110
      Width           =   9300
      Begin MSFlexGridLib.MSFlexGrid gridAtendimentosDoEndereco 
         Height          =   8730
         Left            =   105
         TabIndex        =   12
         Top             =   285
         Width           =   9090
         _ExtentX        =   16034
         _ExtentY        =   15399
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   9300
      Begin VB.Label Lbl_Registro 
         Caption         =   "End.:"
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
         Left            =   2025
         TabIndex        =   10
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
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
         Left            =   90
         TabIndex        =   9
         Top             =   255
         Width           =   690
      End
      Begin VB.Label lblcodigoendereco 
         Caption         =   "..."
         Height          =   210
         Left            =   825
         TabIndex        =   8
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Bairro:"
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
         Left            =   105
         TabIndex        =   7
         Top             =   660
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Cidade:"
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
         Left            =   3945
         TabIndex        =   6
         Top             =   660
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "UF:"
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
         Left            =   8550
         TabIndex        =   5
         Top             =   660
         Width           =   345
      End
      Begin VB.Label lblEndereco 
         Caption         =   "..."
         Height          =   210
         Left            =   2565
         TabIndex        =   4
         Top             =   240
         Width           =   6525
      End
      Begin VB.Label lblbairro 
         Caption         =   "..."
         Height          =   210
         Left            =   825
         TabIndex        =   3
         Top             =   630
         Width           =   2985
      End
      Begin VB.Label lblcidade 
         Caption         =   "..."
         Height          =   210
         Left            =   4665
         TabIndex        =   2
         Top             =   630
         Width           =   3060
      End
      Begin VB.Label lbluf 
         Caption         =   "..."
         Height          =   210
         Left            =   8910
         TabIndex        =   1
         Top             =   660
         Width           =   330
      End
   End
End
Attribute VB_Name = "frmAtendimentosEndereco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblcodigoendereco.Caption = Form_Service.strCodEndereco
    lblEndereco.Caption = Form_Service.strEndereco
    lblbairro.Caption = Form_Service.strBairro
    lblcidade.Caption = Form_Service.strCidade
    lbluf.Caption = Form_Service.strUF
    
    Call configuregrid
    
    Call Show_Records(Return_Recordset(Return_SqlString(1, lblcodigoendereco.Caption)))
    
End Sub

Private Sub configuregrid()
    With gridAtendimentosDoEndereco
        .Cols = 6
        .ColWidth(0) = 690
        .ColWidth(1) = 1200
        .ColWidth(2) = 3900
        .ColWidth(3) = 850
        .ColWidth(4) = 1075
        .ColWidth(5) = 1100
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "OS"
        .TextMatrix(0, 1) = "Dt Atendimento"
        .TextMatrix(0, 2) = "Tipo Serv"
        .TextMatrix(0, 3) = "Valor"
        .TextMatrix(0, 4) = "Status"
        .TextMatrix(0, 5) = "Dt Status"
        
        .RowHeight(1) = 0
    End With
End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String, Optional param3 As String) As String
    Dim whereClause As String
    
    If queryType = 1 Then
        Return_SqlString = "SELECT * FROM Atendimento WHERE cod_end = " & param1 & " ORDER BY Codigo"
    End If
End Function

Private Function Return_Recordset(strSql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Call Form_Fornecedores.FornecedorArquivos_Dados
    
    Set conn = Nothing
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
    
    'MsgBox strSql
    rs.Open strSql, conn, adOpenStatic, adLockOptimistic
    
    Set Return_Recordset = rs
End Function

Private Sub Show_Records(rs As ADODB.Recordset, Optional code As Integer)
    Dim selectedRow As Integer
    Dim col As Integer
    Dim bkColor As Long
    Dim qtdRows As Integer
        
    selectedRow = 2
    With gridAtendimentosDoEndereco
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("codigo") & vbTab & IIf(IsNull(rs.Fields("Dt_Atend_Serv")), "", rs.Fields("Dt_Atend_Serv")) & vbTab & IIf(IsNull(rs.Fields("Tipo_Serv")), "", rs.Fields("Tipo_Serv")) & vbTab & IIf(IsNull(rs.Fields("Valor_Serv")), "", rs.Fields("Valor_Serv")) & vbTab & IIf(IsNull(rs.Fields("Status")), "", rs.Fields("Status"))) & vbTab & rs.Fields("Dt_Status") & vbTab & rs.Fields("Equip1_Codigo") & vbTab & IIf(IsNull(rs.Fields("Equip2_Codigo")), "", rs.Fields("Equip2_Codigo")) & vbTab & IIf(IsNull(rs.Fields("Equip3_Codigo")), "", rs.Fields("Equip3_Codigo")) & vbTab & IIf(IsNull(rs.Fields("Equip4_Codigo")), "", rs.Fields("Equip4_Codigo")) & vbTab & IIf(IsNull(rs.Fields("Equip5_Codigo")), "", rs.Fields("Equip5_Codigo")) & vbTab & IIf(IsNull(rs.Fields("Equip6_Codigo")), "", rs.Fields("Equip6_Codigo")) & vbTab & IIf(IsNull(rs.Fields("Equip7_Codigo")), "", rs.Fields("Equip7_Codigo")) _
            & vbTab & IIf(IsNull(rs.Fields("Equip8_Codigo")), "", rs.Fields("Equip8_Codigo")) & vbTab & IIf(IsNull(rs.Fields("Equip9_Codigo")), "", rs.Fields("Equip9_Codigo")) & vbTab & IIf(IsNull(rs.Fields("Equip10_Codigo")), "", rs.Fields("Equip10_Codigo"))
            
            
            '.AddItem (rs.Fields("Equip1_Codigo"))
            'MsgBox .TextMatrix(22, 6)
            
            If code = rs.Fields("codigo") Then
                selectedRow = .Rows - 1
            End If
            'bkColor = GET_BACKCOLOR_BY_STATUS(rs.Fields("Status"))
            
            'If bkColor <> 0 Then
            '    .row = .Rows - 1
            '    For col = .FixedCols To .Cols - 1
            '        .col = col
            '        .CellBackColor = bkColor
            '    Next
            'End If
            rs.MoveNext
        Loop
        
        qtdRows = .Rows '
        
        .Refresh
        

        If rs.RecordCount > 0 Then
            .row = qtdRows - 1 'Para posicionar o cursor no último registro do grid
            .RowSel = selectedRow
            .TopRow = 2 'selectedRow
            .col = 0
            .ColSel = .Cols - 1
            'comentado
            'Call Grid_Service_Click
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

