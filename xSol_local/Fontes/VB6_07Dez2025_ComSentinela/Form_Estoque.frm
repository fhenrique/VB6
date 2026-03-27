VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Estoque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Estoque"
   ClientHeight    =   7065
   ClientLeft      =   3240
   ClientTop       =   1905
   ClientWidth     =   15105
   Icon            =   "Form_Estoque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   15105
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Visão do Estoque"
      Height          =   930
      Left            =   5235
      TabIndex        =   36
      Top             =   15
      Width           =   9795
      Begin VB.ComboBox cboVisaoEstoqueMaterial 
         Height          =   315
         Left            =   1470
         TabIndex        =   37
         Text            =   "cboVisaoEstoqueMaterial"
         ToolTipText     =   "Verifique se a mercadoria já existe antes de salvar uma nova."
         Top             =   375
         Width           =   4155
      End
      Begin VB.Label lblCodigoMaterial 
         AutoSize        =   -1  'True
         Caption         =   "Label13"
         Height          =   195
         Left            =   5760
         TabIndex        =   50
         Top             =   390
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lblCodMaterial 
         AutoSize        =   -1  'True
         Caption         =   "..."
         Height          =   195
         Left            =   5475
         TabIndex        =   45
         Top             =   465
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Material.........:"
         Height          =   195
         Left            =   90
         TabIndex        =   38
         Top             =   465
         Width           =   1005
      End
   End
   Begin VB.CommandButton btnFechar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14415
      Picture         =   "Form_Estoque.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6375
      UseMaskColor    =   -1  'True
      Width           =   525
   End
   Begin VB.Frame Frame1 
      Caption         =   "Relatórios"
      Height          =   1110
      Left            =   5250
      TabIndex        =   30
      Top             =   5925
      Width           =   8115
      Begin VB.OptionButton optnEstoqueGeral 
         Caption         =   "Estoque geral"
         Height          =   195
         Left            =   2925
         TabIndex        =   51
         Top             =   750
         Width           =   2835
      End
      Begin VB.CommandButton btnRelGerar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Gerar relatório"
         Height          =   765
         Left            =   6585
         Picture         =   "Form_Estoque.frx":03DA
         Style           =   1  'Graphical
         TabIndex        =   35
         Tag             =   "1"
         Top             =   210
         Width           =   1380
      End
      Begin VB.OptionButton optnRelRetiradaEquipamentos 
         Caption         =   "&Retirada de equipamentos"
         Height          =   195
         Left            =   2925
         TabIndex        =   34
         Tag             =   "2"
         Top             =   375
         Width           =   2550
      End
      Begin VB.OptionButton optnRelSaidaEquipamentos 
         Caption         =   "&Saída de equipamentos"
         Height          =   195
         Left            =   225
         TabIndex        =   33
         Tag             =   "3"
         Top             =   750
         Width           =   2595
      End
      Begin VB.OptionButton optnRelEntradaEquipamentos 
         Caption         =   "&Entrada de equipamentos"
         Height          =   195
         Left            =   225
         TabIndex        =   32
         Tag             =   "1"
         Top             =   375
         Value           =   -1  'True
         Width           =   2445
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Histórico de Movimentações"
      Height          =   4830
      Left            =   5235
      TabIndex        =   23
      Top             =   1065
      Width           =   9795
      Begin VB.Frame Frame5 
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   2490
         TabIndex        =   44
         Top             =   4335
         Width           =   1245
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtd. Saída:"
            Height          =   195
            Left            =   30
            TabIndex        =   49
            Top             =   75
            Width           =   825
         End
         Begin VB.Label lblQtdSaida 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            Height          =   195
            Left            =   870
            TabIndex        =   48
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   915
         TabIndex        =   43
         Top             =   4335
         Width           =   1350
         Begin VB.Label lblQtdEntrada 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            Height          =   195
            Left            =   990
            TabIndex        =   47
            Top             =   75
            Width           =   135
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtd. Entrada:"
            Height          =   195
            Left            =   15
            TabIndex        =   46
            Top             =   75
            Width           =   945
         End
      End
      Begin VB.TextBox txtVisaoEstoque_Estoque 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   9060
         TabIndex        =   40
         Top             =   4365
         Width           =   600
      End
      Begin MSFlexGridLib.MSFlexGrid gridVisaoEstoqueProduto 
         Height          =   3975
         Left            =   120
         TabIndex        =   39
         Top             =   285
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   7011
         _Version        =   393216
         Rows            =   1
         Cols            =   9
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         SelectionMode   =   1
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Legenda:"
         Height          =   195
         Left            =   165
         TabIndex        =   42
         Top             =   4410
         Width           =   675
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Estoque........:"
         Height          =   195
         Left            =   7920
         TabIndex        =   41
         Top             =   4485
         Width           =   990
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Edição"
      Height          =   1095
      Left            =   30
      TabIndex        =   22
      Top             =   5925
      Width           =   5115
      Begin VB.CommandButton Cmd_Inclui 
         BackColor       =   &H0080FF80&
         Caption         =   "&Inserir"
         Height          =   495
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   390
         Width           =   1005
      End
      Begin VB.CommandButton Cmd_Salva 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1380
         TabIndex        =   8
         Top             =   360
         Width           =   1050
      End
      Begin VB.CommandButton Cmd_Edita 
         Caption         =   "&Editar"
         Height          =   495
         Left            =   3900
         TabIndex        =   7
         Top             =   375
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton Cmd_Exclui 
         BackColor       =   &H008080FF&
         Caption         =   "E&xcluir"
         Height          =   495
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   375
         Width           =   1065
      End
   End
   Begin VB.Frame Form_Service 
      Caption         =   "Movimentação"
      Height          =   5895
      Left            =   30
      TabIndex        =   10
      Top             =   15
      Width           =   5115
      Begin VB.CheckBox ChkLiberado 
         BackColor       =   &H8000000A&
         Caption         =   "Liberado"
         Enabled         =   0   'False
         Height          =   360
         Left            =   4035
         MaskColor       =   &H00000000&
         TabIndex        =   29
         Top             =   2745
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4320
         TabIndex        =   27
         Top             =   330
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.ComboBox txtProdutoDescricao 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   26
         ToolTipText     =   "Verifique se a mercadoria já existe antes de salvar uma nova."
         Top             =   960
         Visible         =   0   'False
         Width           =   3750
      End
      Begin VB.Frame panSaida 
         Caption         =   "Saída"
         Height          =   2520
         Left            =   120
         TabIndex        =   18
         Top             =   3240
         Width           =   4860
         Begin VB.ComboBox Combo_Tecnicos 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Form_Estoque.frx":0D1C
            Left            =   195
            List            =   "Form_Estoque.frx":0D1E
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1710
            Width           =   3450
         End
         Begin VB.TextBox txtSaidaValor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   885
            TabIndex        =   4
            Top             =   690
            Width           =   765
         End
         Begin VB.TextBox txtSaidaQtd 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2670
            TabIndex        =   5
            Top             =   690
            Width           =   600
         End
         Begin MSComCtl2.DTPicker dtpSaidaData 
            Height          =   315
            Left            =   885
            TabIndex        =   3
            Top             =   240
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CalendarBackColor=   16777215
            Format          =   131465217
            CurrentDate     =   44025
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Técnico:"
            Height          =   195
            Left            =   195
            TabIndex        =   24
            Top             =   1455
            Width           =   630
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   195
            TabIndex        =   21
            Top             =   330
            Width           =   390
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            Height          =   195
            Left            =   210
            TabIndex        =   20
            Top             =   780
            Width           =   405
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtd:"
            Height          =   195
            Left            =   2100
            TabIndex        =   19
            Top             =   810
            Width           =   300
         End
      End
      Begin VB.Frame panEntrada 
         Caption         =   "Entrada"
         Height          =   1290
         Left            =   120
         TabIndex        =   14
         Top             =   1470
         Width           =   4860
         Begin MSComCtl2.DTPicker dtpEntradaData 
            Height          =   315
            Left            =   945
            TabIndex        =   0
            Top             =   345
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CalendarBackColor=   16777215
            Format          =   131465217
            CurrentDate     =   44025
         End
         Begin VB.TextBox txtEntradaQtd 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2700
            TabIndex        =   2
            Top             =   795
            Width           =   600
         End
         Begin VB.TextBox txtEntradaValor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   930
            TabIndex        =   1
            Top             =   780
            Width           =   765
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtd:"
            Height          =   195
            Left            =   2100
            TabIndex        =   17
            Top             =   900
            Width           =   300
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            Height          =   195
            Left            =   195
            TabIndex        =   16
            Top             =   870
            Width           =   405
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   210
            TabIndex        =   15
            Top             =   435
            Width           =   390
         End
      End
      Begin VB.TextBox txtEstoqueCodigo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   345
         Width           =   960
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
         Height          =   195
         Left            =   3555
         TabIndex        =   28
         Top             =   435
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Lbl_Cliente 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Material:"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   990
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Lbl_Registro 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   405
         Width           =   540
      End
   End
End
Attribute VB_Name = "Form_Estoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit


'Public Const COLOR_BTNFACE = &H8000000F  'Button
'Public Const ADMIN_PROFILE = "ADMIN"

Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public saldoEstoque As Double
Public movimentoEntrada As Boolean
Public movimentoSaida As Boolean
Public movEditar As Boolean

Private conn As ADODB.Connection
Public Nome_Cli_Atend As String
Public Estoque_DBname As String
Public Image_Folder As String
Public Cod_Cli As Integer
Public Cod_Address As String
Public Cod_Equipment As String
Public strPorcemtagemLucro As String
Public strValorSaida As String
Public continuaMovimento As Boolean
Public cancelaMovimento As Boolean

Public entradaValor As String
Public saidaValor As String

Public entradaQtd As Integer
Public saidaQtd As Integer

Public Sub Set_Connection()
   Set conn = New ADODB.Connection
   conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
End Sub

Public Sub ConnectionRefresh()
    conn.Close
    Set conn = Nothing
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"
End Sub

Public Function Get_Field_Val(ByVal val As Object) As String
    Get_Field_Val = UCase(IIf(IsNull(val), "", val))
End Function

Public Function Set_Field_Val(ByVal val As Object) As String
    Set_Field_Val = IIf(val = "", vbNullString, UCase(val))
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


Public Function Return_Recordset(strSql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    Call EstoqueArquivos_Dados
    
    Set conn = Nothing
    Set conn = New ADODB.Connection
    conn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & Estoque_DBname & ";"
    
    rs.Open strSql, conn, adOpenStatic, adLockOptimistic
    
    Set Return_Recordset = rs
End Function

Public Function isAdmin() As Boolean
    isAdmin = usuario.Profile = ADMIN_PROFILE
End Function


Private Sub Configure_Grid()
    With gridVisaoEstoqueProduto
        .Cols = 10
        .ColWidth(0) = 1000 'Data
        .ColWidth(1) = 900 'Val. Entrada
        .ColWidth(2) = 1100 'Qtd. Entrada
        .ColWidth(3) = 700 '% Lucro
        .ColWidth(4) = 800 'Val. Saída
        .ColWidth(5) = 870 'Qtd saída
        .ColWidth(6) = 800 'Liberado
        .ColWidth(7) = 2370  'Tecnico
        .ColWidth(8) = 900  'Saldo
        .ColWidth(9) = 900  'Codigo
        
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignLeftCenter
        .ColAlignment(8) = flexAlignLeftCenter
        
        .Rows = 2
        
        .FixedRows = 1
        .TextMatrix(0, 0) = "Data"
        .TextMatrix(0, 1) = "Vl. Entrada"
        .TextMatrix(0, 2) = "Qtd. Entrada"
        .TextMatrix(0, 3) = "% Lucro"
        .TextMatrix(0, 4) = "Vl. Saída"
        .TextMatrix(0, 5) = "Qtd. Saída"
        .TextMatrix(0, 6) = "Liberado"
        .TextMatrix(0, 7) = "Técnico"
        .TextMatrix(0, 8) = "Saldo"
        
        .RowHeight(1) = 0
    End With
End Sub

Private Sub Object_Controller(cType As Integer)

    txtEstoqueCodigo.enabled = cType = 3
    'txtProdutoDescricao.enabled = cType >= 3
    'Combo_Profile.enabled = Text_Passwd.enabled
    
    'txtEstoqueCodigo.BackColor = IIf(txtEstoqueCodigo.enabled, vbYellow, vbWhite)
    'txtProdutoDescricao.BackColor = IIf(cType >= 3, vbYellow, vbWhite)
    'Combo_Profile.BackColor = IIf(cType >= 3, vbYellow, vbWhite)
    
    'Cmd_Inclui.enabled = cType <> 4
    'Cmd_Inclui.Caption = IIf(cType = 3, "&Cancela", "&Incluir")
    'Cmd_Edita.enabled = cType = 2 Or cType = 4
    'Cmd_Edita.Caption = IIf(cType = 4, "&Cancela", "&Editar")
    'Cmd_Salva.enabled = cType >= 3
    'Cmd_Exclui.enabled = cType = 2
    'Grid_Estoque.enabled = cType <= 2
    
    If cType = 1 Then  'without any selection
        'Call Clean_Boxes
    ElseIf cType = 3 Then  'insert a record
        'txtEstoqueCodigo.SetFocus
        'Call Clean_Boxes
    End If
End Sub


Private Function GET_BACKCOLOR(d As Date) As Long
    If d <> #1/1/1900# Then
        GET_BACKCOLOR = RGB(250, 250, 250)
    ElseIf d < Date Then
        GET_BACKCOLOR = RGB(250, 160, 160)
    ElseIf d = Date Then
        GET_BACKCOLOR = RGB(255, 255, 71)
    ElseIf d < DateAdd("d", 7, Date) Then
        GET_BACKCOLOR = RGB(250, 192, 144)
    ElseIf d < DateAdd("m", 1, Date) Then
        GET_BACKCOLOR = RGB(0, 255, 0)
    Else
        GET_BACKCOLOR = 0
    End If
End Function


Private Sub Fill_Grid_VisaoEstoque(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    entradaQtd = 0
    saidaQtd = 0
    
    selectedRow = 2
    
    With gridVisaoEstoqueProduto
        '.Clear
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
            lblQtdEntrada.Caption = "0"
            lblQtdSaida.Caption = "0"
            txtVisaoEstoque_Estoque.Text = "0"
           Exit Sub
        End If

        Do Until rs.EOF
            
            .AddItem (rs.Fields("data") & vbTab & rs.Fields("valor_ent") & vbTab & rs.Fields("Qtd_Ent") & vbTab & rs.Fields("Porcent_Lucro") & vbTab & rs.Fields("Valor_Saida") & vbTab & rs.Fields("Qtd_Saida") & vbTab & rs.Fields("Liberado") & vbTab & rs.Fields("Nome_Tecnico") & vbTab & rs.Fields("Saldo") & vbTab & rs.Fields("Codigo"))
            
            If (rs.Fields("Qtd_Ent") <> "") And (rs.Fields("Qtd_Ent") > 0) Then
                entradaQtd = entradaQtd + rs.Fields("Qtd_Ent")
            Else
                saidaQtd = saidaQtd + rs.Fields("Qtd_Saida")
            End If
            
            
            
            
            'If code = rs.Fields("codigo") Then
            '    selectedRow = .Rows - 1
            'End If
            
            'faz a pintura das linhas aqui
            If (IsNull(rs.Fields("Nome_Tecnico"))) Or (rs.Fields("Nome_Tecnico") = "") Then
                bkColor = RGB(153, 255, 153)
                'bkColor = vbgreen
            Else
                bkColor = RGB(255, 153, 153)
                'bkColor = vbRed
            End If
            'bkColor = GET_BACKCOLOR(rs.Fields("Dt_Ent"))
            
            If bkColor <> 0 Then
                .row = .Rows - 1
                For col = .FixedCols To .Cols - 1
                    .col = col
                    .CellBackColor = bkColor
                Next
            End If
            
            
'            If rs.Fields("Qtd_Ent") <> "" Then
'                bkcolor = RGB(250, 250, 250)
'                .CellBackColor = bkcolor
'            Else
'                .CellBackColor = vbRed
'            End If
            
            rs.MoveNext
        Loop
        
'        MsgBox entradaQtd - saidaQtd
        lblQtdEntrada.Caption = Str(entradaQtd)
        lblQtdSaida.Caption = Str(saidaQtd)
        
        txtVisaoEstoque_Estoque.Text = Str(entradaQtd - saidaQtd)
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            If .Rows > 9 Then
                .TopRow = selectedRow
            End If
            .col = 0
            .ColSel = .Cols - 1
            
            Call Object_Controller(2)
            Call Grid_Estoque_Click

        Else
            Call Object_Controller(1)
        End If
    End With
    
    
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Fill_Grid(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    
    selectedRow = 2
    With gridVisaoEstoqueProduto
        .Rows = 2
        
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If

        Do Until rs.EOF
            .AddItem (rs.Fields("codigo") & vbTab & IIf(IsNull(rs.Fields("produto")), "", rs.Fields("produto")) & vbTab & IIf(IsNull(rs.Fields("Dt_Ent")), "", rs.Fields("Dt_Ent")) & vbTab & IIf(IsNull(rs.Fields("Valor_Ent")), "", rs.Fields("Valor_Ent")) & vbTab & IIf(IsNull(rs.Fields("Qtd_Ent")), "", rs.Fields("Qtd_Ent")) & vbTab & IIf(IsNull(rs.Fields("Porcent_Lucro")), "", rs.Fields("Porcent_Lucro")) & vbTab & IIf(IsNull(rs.Fields("Dt_Saida")), "", rs.Fields("Dt_Saida")) & vbTab & IIf(IsNull(rs.Fields("Valor_Saida")), "", rs.Fields("Valor_Saida")) & vbTab & IIf(IsNull(rs.Fields("Qtd_Saida")), "", rs.Fields("Qtd_Saida")) & vbTab & IIf(IsNull(rs.Fields("Liberado")), "", rs.Fields("Liberado")) & vbTab & IIf(IsNull(rs.Fields("Nome_Tecnico")), "", rs.Fields("Nome_Tecnico")) & vbTab & IIf(IsNull(rs.Fields("Saldo")), "", rs.Fields("Saldo")))
            If code = rs.Fields("codigo") Then
                selectedRow = .Rows - 1
            End If
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
            
            Call Object_Controller(2)
            Call Grid_Estoque_Click
        Else
            Call Object_Controller(1)
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Public Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String) As String
    Dim whereClause As String
    
    If queryType = 1 Then
        Return_SqlString = "SELECT * FROM Estoque"
    ElseIf queryType = 2 Then
        Return_SqlString = "SELECT * FROM Estoque WHERE Codigo = " & param1
    ElseIf queryType = 3 Then
        Return_SqlString = "SELECT Descricao, Codigo FROM Material order by Descricao"
    ElseIf queryType = 4 Then
        Return_SqlString = "SELECT Distinct(Nome) FROM Tecnico Where Status = 'A'"
    ElseIf queryType = 5 Then
        Return_SqlString = "SELECT Sum(Saldo) As Saldo FROM Estoque WHERE Produto = """ & param1 & """"
    ElseIf queryType = 6 Then
        Return_SqlString = "SELECT Min(Saldo) As Saldo FROM Estoque WHERE Produto = """ & param1 & """"
    ElseIf queryType = 7 Then
        Return_SqlString = "SELECT Sum(Saldo) As Saldo FROM Estoque WHERE Produto = """ & param1 & """"
        'MsgBox Return_SqlString
    ElseIf queryType = 8 Then
        'Traz o último valor de entrada do produto  x
        Return_SqlString = "SELECT Top 1 Max(Codigo), Valor_Ent, Porcent_Lucro From Estoque WHERE Produto = '" & param1 & "' AND Qtd_Saida = 0 GROUP BY Valor_Ent, Porcent_Lucro ORDER by 1 DESC"
    ElseIf queryType = 9 Then
        If param1 = "..." Then
           param1 = "0"
        End If
        Return_SqlString = "SELECT Codigo, Dt_Ent & Dt_Saida as data , Valor_Ent, Qtd_Ent, Porcent_Lucro, Valor_Saida, Qtd_Saida, Liberado, Nome_Tecnico, Saldo, Codigo FROM Estoque WHERE CodMaterial = " & param1
    ElseIf queryType = 10 Then
        Return_SqlString = "SELECT * FROM Material WHERE Descricao= """ & param1 & """"
    ElseIf queryType = 11 Then
        Return_SqlString = "SELECT * FROM MaterialFornecedor WHERE CodMaterial = " & param1
    ElseIf queryType = 12 Then
        Return_SqlString = "Select top 1 valor_saida from estoque WHERE produto = '" & param1 & "' and cdbl(IIf(IsNull(valor_saida), '0', valor_saida)) > 0 order by dt_saida desc"
    ElseIf queryType = 13 Then
        Return_SqlString = "Select top 1 valor_ent from estoque WHERE produto = '" & param1 & "' and cdbl(IIf(IsNull(valor_ent), '0', valor_ent)) > 0 order by dt_ent desc"
    End If
    
End Function

Private Sub btnFechar_Click()
    Unload Me
End Sub

Private Sub executaRelatorio(nTag As Integer, pathDB As String)
    frmFiltroRelatorio.tipoRelatorio = nTag
    frmFiltroRelatorio.pathDataBase = pathDB
    
    frmFiltroRelatorio.Show (1)
End Sub

Private Sub btnRelGerar_Click()
    If btnRelGerar.Tag <> 4 Then
        Call executaRelatorio(btnRelGerar.Tag, Estoque_DBname)
    Else
        'Gera o relatório geral
        relEstoqueGeral.DAODataControl1.DatabaseName = Estoque_DBname
        relEstoqueGeral.DAODataControl1.RecordSource = "Select codmaterial as codigo, produto, sum(saldo) as qtd From Estoque Group by codmaterial, produto order by 2"
        relEstoqueGeral.Show (1)
    End If
End Sub

Private Sub carregaCodMaterial(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    If rs.EOF Or rs.BOF Then
       Exit Sub
    Else
        lblCodMaterial.Caption = rs.Fields("Codigo")
        lblCodigoMaterial.Caption = rs.Fields("Codigo")
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Function vlrFornecedorCompra(rs As ADODB.Recordset, Optional code As String) As Integer
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    If rs.EOF Or rs.BOF Then
       Exit Function
    Else
        vlrFornecedorCompra = rs.Fields("ValorCompra")
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function vlrFornecedorVenda(rs As ADODB.Recordset, Optional code As String) As Integer
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    If rs.EOF Or rs.BOF Then
       Exit Function
    Else
        vlrFornecedorVenda = rs.Fields("ValorVenda")
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function qtdFornecedores(rs As ADODB.Recordset, Optional code As String) As Integer
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    If rs.EOF Or rs.BOF Then
       Exit Function
    Else
        qtdFornecedores = rs.RecordCount
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function carregaValorMaterial(rs As ADODB.Recordset, Optional code As String) As Integer
    Dim selectedRow As Integer
    Dim bkColor As Long
    Dim col As Integer
    
    If rs.EOF Or rs.BOF Then
       Exit Function
    Else
        lblCodMaterial.Caption = rs.Fields("Codigo")
        carregaValorMaterial = rs.RecordCount
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub cboVisaoEstoqueMaterial_Click()
'On Error GoTo saida
    If cboVisaoEstoqueMaterial.Text <> "" Then
        Call carregaCodMaterial(Return_Recordset(Return_SqlString(10, cboVisaoEstoqueMaterial.Text)))
        Call Fill_Grid_VisaoEstoque(Return_Recordset(Return_SqlString(9, lblCodMaterial.Caption)))
        
        
        Call Show_Record_Detail(Return_Recordset(Return_SqlString(9, lblCodMaterial.Caption)))
        
        
        Call VisaoEstoque_Show_Record_Detail_Saldo(Return_Recordset(Return_SqlString(7, cboVisaoEstoqueMaterial.Text)))
        
    End If
    
'saida:
'    Exit Sub
End Sub

Private Sub Cmd_Edita_Click()
   If Combo_Tecnicos.Text = "" Then
    movimentoEntrada = True
    movimentoSaida = False
   Else
    movimentoEntrada = False
    movimentoSaida = True
   End If
      
   movEditar = True
   
   gridVisaoEstoqueProduto.enabled = False
    
   'Dim teste As String
   'teste = Grid_Estoque.TextMatrix(Grid_Estoque.RowSel, 11)
   
   'If Grid_Estoque.TextMatrix(Grid_Estoque.RowSel, 0) = "1" Then
    
   'End If
    
    If txtProdutoDescricao.Text = "" Then
        MsgBox "Selecione um movimento na grade ao lado ->", vbOKOnly, "Xsol Aquecedores"
        gridVisaoEstoqueProduto.enabled = True
        Exit Sub
    End If

    If gridVisaoEstoqueProduto.Rows = 2 Then
        Exit Sub
    End If
    
    If Combo_Tecnicos.Text = "" Then
        movimentoEntrada = True
        movimentoSaida = False
        panEntrada.enabled = True
        panSaida.enabled = False
    Else
        movimentoEntrada = False
        movimentoSaida = True
        panEntrada.enabled = False
        panSaida.enabled = True
    End If

    If movimentoEntrada Then
        txtEntradaValor.enabled = True
        txtEntradaValor.BackColor = &HFFFFFF
        txtEntradaQtd.enabled = True
        txtEntradaQtd.BackColor = &HFFFFFF
        
        txtSaidaValor.enabled = False
        txtSaidaValor.BackColor = &HC0C0C0
        txtSaidaQtd.enabled = False
        txtSaidaQtd.BackColor = &HC0C0C0
    Else
        txtEntradaValor.enabled = False
        txtEntradaValor.BackColor = &HC0C0C0
        txtEntradaQtd.enabled = False
        txtEntradaQtd.BackColor = &HC0C0C0
        
        txtSaidaValor.enabled = True
        txtSaidaValor.BackColor = &HFFFFFF
        txtSaidaQtd.enabled = True
        txtSaidaQtd.BackColor = &HFFFFFF
    End If


    If Cmd_Edita.Caption = "&Editar" Then 'HABILITA
    
        Cmd_Edita.Caption = "&Cancelar"
        
        Cmd_Inclui.enabled = False
        Cmd_Salva.enabled = True
        Cmd_Exclui.enabled = False
        
        txtProdutoDescricao.enabled = True
        txtProdutoDescricao.BackColor = &HFFFFFF
        
        
'        If Cmd_Edita.Caption <> "&Cancelar" Then
'            Call carregaMaterial(Return_Recordset(Return_SqlString(3)))
'        End If
'
'        If Cmd_Edita.Caption <> "&Cancelar" Then
'            Call carregaTecnicos(Return_Recordset(Return_SqlString(4)))
'        End If
        
    Else    'DESABILITA
        Cmd_Edita.Caption = "&Editar"
        gridVisaoEstoqueProduto.enabled = True
        
        Cmd_Inclui.enabled = True
        Cmd_Salva.enabled = False
        Cmd_Exclui.enabled = True
        
        ' desabilita todos os campos
        txtProdutoDescricao.enabled = False
        txtProdutoDescricao.BackColor = &HC0C0C0
        
        txtEntradaValor.enabled = False
        txtEntradaValor.BackColor = &HC0C0C0
        txtEntradaQtd.enabled = False
        txtEntradaQtd.BackColor = &HC0C0C0
        
        
        txtSaidaValor.enabled = False
        txtSaidaValor.BackColor = &HC0C0C0
        txtSaidaQtd.enabled = False
        txtSaidaQtd.BackColor = &HC0C0C0
        
        
        ChkLiberado.enabled = False
        Combo_Tecnicos.enabled = False
        
    End If

End Sub

Private Sub Record_Record(rs As ADODB.Recordset)
   
   
   rs.Fields("Produto").value = cboVisaoEstoqueMaterial.Text
   rs.Fields("codMaterial").value = lblCodMaterial.Caption
   
   If (movimentoEntrada) Then
    rs.Fields("Dt_Ent").value = dtpEntradaData.value
    rs.Fields("Saldo").value = CDec(txtEntradaQtd.Text)
    rs.Fields("Valor_Ent").value = CDec(txtEntradaValor.Text)
    rs.Fields("Qtd_Ent").value = CDec(txtEntradaQtd.Text)
   Else
    rs.Fields("Dt_Saida").value = dtpEntradaData.value
    rs.Fields("Saldo").value = CDec(txtSaidaQtd.Text)
    rs.Fields("Valor_Saida").value = CDec(txtSaidaValor.Text)
    rs.Fields("Qtd_Saida").value = CDec(txtSaidaQtd.Text)
   
   End If
   
   If (movimentoSaida) Then
    rs.Fields("Dt_Saida").value = dtpSaidaData.value
    rs.Fields("Saldo").value = CDec(txtSaidaQtd.Text) * -1
    rs.Fields("Valor_Saida").value = CDec(txtSaidaValor.Text)
    rs.Fields("Qtd_Saida").value = CDec(txtSaidaQtd.Text)
   End If
   
   If Trim(strPorcemtagemLucro) <> "" Then
    rs.Fields("Porcent_Lucro").value = strPorcemtagemLucro
   End If
   
   If (movimentoSaida) And Trim(strValorSaida) <> "" Then
    rs.Fields("Valor_Saida").value = strValorSaida
   End If
   
   If Trim(txtSaidaQtd.Text) <> "" And Trim(txtSaidaQtd.Text) <> "0" Then
    rs.Fields("Dt_Saida").value = dtpSaidaData.value
    
    If Trim(txtSaidaValor.Text) <> "" Then
     rs.Fields("Valor_Saida").value = txtSaidaValor.Text
    Else
     rs.Fields("Valor_Saida").value = "0"
    End If
    
    rs.Fields("Qtd_Saida").value = txtSaidaQtd.Text
    rs.Fields("Nome_Tecnico").value = Combo_Tecnicos.Text
   End If
   
   If ChkLiberado.value = 1 Then
    rs.Fields("Liberado").value = "S"
   Else
    rs.Fields("Liberado").value = "N"
   End If
   
End Sub

Private Sub Cmd_Exclui_Click()
   If gridVisaoEstoqueProduto.Rows > 2 Then
    If MsgBox("Confirma a Exclusão de estoque do produto " & cboVisaoEstoqueMaterial.Text & " ? ", vbYesNo, "Excluir Registros") = vbYes Then
       Call Return_Recordset(Return_SqlString(2, txtEstoqueCodigo.Text)).Delete
       'Call Fill_Grid(Return_Recordset(Return_SqlString(1)))
       cboVisaoEstoqueMaterial_Click
    End If
   Else
    txtEstoqueCodigo.Text = ""
    
    
    txtProdutoDescricao.ListIndex = -1
    
    txtEntradaValor.Text = ""
    txtEntradaQtd.Text = ""
    txtSaidaValor.Text = ""
    txtSaidaQtd.Text = ""
    ChkLiberado.value = 0
   End If
   
End Sub



Private Sub Cmd_Inclui_Click()
Dim ultimoValorDeSaida As Double
Dim ultimoValorDeEntrada As Double
    
    If cboVisaoEstoqueMaterial.Text = "" Then
        MsgBox "Selecione o material"
        cboVisaoEstoqueMaterial.SetFocus
        Exit Sub
    End If

'On Error GoTo saida
    
    movEditar = False
    
    If Cmd_Inclui.Caption = "&Inserir" Then 'HABILITA, está Inserir
        gridVisaoEstoqueProduto.enabled = False
        cboVisaoEstoqueMaterial.enabled = False
        Dim qtdLinhas As Integer
        Dim ultimoCodigo As Integer
        
        Cmd_Inclui.Caption = "&Cancelar"
        
        'qtdLinhas = gridVisaoEstoqueProduto.Rows
        'qtdLinhas = qtdLinhas - 1
        
        'If qtdLinhas >= 2 Then
            'ultimoCodigo = gridVisaoEstoqueProduto.TextMatrix(qtdLinhas, 0)
        'Else
        '    ultimoCodigo = 0
        'End If
        
        'ultimoCodigo = retornaUltimoCodigoEstoque()
        '
    
        txtEstoqueCodigo.Text = Str(ultimoCodigo + 1)
        
        
        'Pergunta se é movimento de ENTRADA ou SAÍDA
        frmMovimentoEstoque.Show (1)
        
        If Not continuaMovimento Then
            Cmd_Inclui_Click
            Exit Sub
        End If
        
        If cancelaMovimento Then
            gridVisaoEstoqueProduto.enabled = True
            cboVisaoEstoqueMaterial.enabled = True
            
            Cmd_Inclui.Caption = "&Inserir"
            
            Cmd_Inclui.enabled = True
            Cmd_Salva.enabled = False
            Cmd_Edita.enabled = True
            Cmd_Exclui.enabled = True
            
            'txtProdutoDescricao.enabled = False
            txtProdutoDescricao.BackColor = &HC0C0C0
            
            dtpEntradaData.enabled = False
            
            'txtEntradaValor.enabled = False
            'txtEntradaValor.BackColor = &HC0C0C0
            
            txtEntradaQtd.enabled = False
            txtEntradaQtd.BackColor = &HC0C0C0
            
            dtpSaidaData.enabled = False
            
            txtSaidaValor.enabled = False
            txtSaidaValor.BackColor = &HC0C0C0
            
            txtSaidaQtd.enabled = False
            txtSaidaQtd.BackColor = &HC0C0C0
            
            ChkLiberado.enabled = False
            
            Combo_Tecnicos.enabled = False
        
            Exit Sub
        
        End If
        
        'Calcula o SALDO ATUAL do material
        Call carregaMaterial(Return_Recordset(Return_SqlString(3)))
        
        If movimentoEntrada Then
            txtProdutoDescricao.enabled = True
            panEntrada.enabled = True
            panSaida.enabled = False
            dtpEntradaData.value = Format(Date, "Short Date")
            dtpEntradaData.enabled = True
        
            
            ultimoValorDeEntrada = retornaUltimoValorDeEntrada(Return_Recordset(Return_SqlString(13, cboVisaoEstoqueMaterial.Text)))
            
            
            txtEntradaValor.Text = ultimoValorDeEntrada
            
            txtEntradaValor.BackColor = &HFFFFFF
            txtEntradaValor.enabled = True
            
            txtEntradaQtd.BackColor = &HFFFFFF
            txtEntradaQtd.enabled = True
            
            Combo_Tecnicos.ListIndex = -1
        Else
            txtProdutoDescricao.enabled = True
            
            txtEntradaValor.Text = ""
            txtEntradaQtd.Text = ""
            txtEntradaValor.enabled = False
            txtEntradaQtd.enabled = False
            
            panEntrada.enabled = False
            panSaida.enabled = True
            dtpSaidaData.value = Format(Date, "Short Date")
            dtpSaidaData.enabled = True
            
            
            
            'jacaré
            ultimoValorDeSaida = retornaUltimoValorDeSaida(Return_Recordset(Return_SqlString(12, cboVisaoEstoqueMaterial.Text)))
            
            txtSaidaValor.Text = Str(ultimoValorDeSaida)
            
            
            txtSaidaValor.BackColor = &HFFFFFF
            txtSaidaValor.enabled = True
            
            txtSaidaQtd.Text = ""
            txtSaidaQtd.BackColor = &HFFFFFF
            txtSaidaQtd.enabled = True
            
            Combo_Tecnicos.enabled = True
            
        End If
    
        'If txtProdutoDescricao.enabled Then
            'txtProdutoDescricao.Text = cboVisaoEstoqueMaterial.Text
            'txtProdutoDescricao.SetFocus
        'End If
        Cmd_Inclui.Caption = "&Cancelar"
        
        Cmd_Inclui.enabled = True
        Cmd_Edita.enabled = False
        Cmd_Salva.enabled = True
        Cmd_Exclui.enabled = False
        
        'Call carregaTecnicos(Return_Recordset(Return_SqlString(4)))
    Else    'DESABILITA, está CANCELAR
saida:
        gridVisaoEstoqueProduto.enabled = True
        cboVisaoEstoqueMaterial.enabled = True
        
        Cmd_Inclui.Caption = "&Inserir"
        
        Cmd_Inclui.enabled = True
        Cmd_Salva.enabled = False
        Cmd_Edita.enabled = True
        Cmd_Exclui.enabled = True
        
        txtProdutoDescricao.ListIndex = -1
        txtProdutoDescricao.enabled = False
        txtProdutoDescricao.BackColor = &HC0C0C0
        
        dtpEntradaData.enabled = False
        
        txtEntradaValor.Text = ""
                
        txtEntradaQtd.Text = ""
        txtEntradaQtd.enabled = False
        txtEntradaQtd.BackColor = &HC0C0C0
        
        
        dtpSaidaData.enabled = False
        
        txtSaidaValor.enabled = False
        txtSaidaValor.BackColor = &HC0C0C0
        
        txtSaidaQtd.enabled = False
        txtSaidaQtd.BackColor = &HC0C0C0
        
        ChkLiberado.enabled = False
        
        Combo_Tecnicos.enabled = False
    End If



End Sub

Private Sub Cmd_Salva_Click()
    Dim dblTmp As Double
    
    Dim rs As ADODB.Recordset
    Dim nomeDoProduto As String
    
    'If Trim(txtProdutoDescricao.Text) = "" Then
    '  MsgBox "Preencha o nome do produto.", vbOKOnly, "Xsol Aquecedores"
    '  txtProdutoDescricao.SetFocus
    '  Exit Sub
    'End If
      
    If movimentoEntrada Then
        If (Trim(txtEntradaQtd.Text) = "0") Or (Trim(txtEntradaQtd.Text) = "") Then
          MsgBox "Preencha a Qtd.", vbOKOnly, "Xsol Aquecedores"
          txtEntradaQtd.SetFocus
          Exit Sub
        End If
        
        If Trim(txtEntradaValor.Text) = "" Then
          If MsgBox("O valor do produto não foi informado, deseja que o sistema traga automaticamente o valor da última compra?", vbYesNo) = vbNo Then
            Exit Sub
          End If
        End If
        
        If Trim(txtEntradaQtd.Text) = "" Then
          MsgBox "Preencha a Qtd do produto na Entrada", vbOKOnly, "Xsol Aquecedores"
          txtEntradaQtd.SetFocus
          Exit Sub
        End If
        
        If Not IsNumeric(txtEntradaQtd.Text) Then
          MsgBox "Informe um valor numérico no campo Qtd de Entrada.", vbOKOnly, "Xsol Aquecedores"
          txtEntradaQtd.SetFocus
          Exit Sub
        End If
        
        validaQtdFornecedor
        
    Else
        If (movEditar) And (txtSaidaQtd.Text = "0") Then
            MsgBox "Informe um valor numérico no campo Qtd.", vbOKOnly, "Xsol Aquecedores"
            txtSaidaQtd.SetFocus
            Exit Sub
        End If
      
        If Trim(txtSaidaQtd.Text) = "" Then
            MsgBox "Informe um valor numérico no campo Qtd.", vbOKOnly, "Xsol Aquecedores"
            txtSaidaQtd.SetFocus
            Exit Sub
        End If
      
        If Trim(txtSaidaValor.Text) <> "" Then
          If Not IsNumeric(txtSaidaValor.Text) Then
            MsgBox "Informe um valor numérico no campo Valor de Saída.", vbOKOnly, "Xsol Aquecedores"
            txtSaidaValor.SetFocus
            Exit Sub
          End If
        End If
        
        If Trim(txtSaidaQtd.Text) <> "" Then
          If Not IsNumeric(txtSaidaQtd.Text) Then
            MsgBox "Informe um valor numérico no campo Qtd de Saída.", vbOKOnly, "Xsol Aquecedores"
            txtSaidaQtd.SetFocus
            Exit Sub
          End If
        End If
        
        If (CDbl(txtVisaoEstoque_Estoque.Text) < 1) And (movEditar = False) Then
          MsgBox "O estoque do produto " & txtProdutoDescricao.Text & " é de " & txtVisaoEstoque_Estoque.Text, vbOKOnly, "Xsol Aquecedores"
          Exit Sub
        End If
        
        If Trim(txtSaidaQtd.Text) <> "" Then
          If (CDbl(txtSaidaQtd.Text) > CDbl(txtVisaoEstoque_Estoque.Text)) And (movEditar = False) Then
            MsgBox "O estoque possui apenas " & txtVisaoEstoque_Estoque.Text & " itens do material " & txtProdutoDescricao.Text, vbOKOnly, "Xsol Aquecedores"
            txtSaidaQtd.SetFocus
            Exit Sub
          End If
        End If
      
        If movEditar = True Then
          dblTmp = gridVisaoEstoqueProduto.TextMatrix(gridVisaoEstoqueProduto.RowSel, 11)
          dblTmp = dblTmp * -1
          
          If CDbl(txtEntradaQtd.Text) > dblTmp Then
            MsgBox "A quantidade não pode ser superior a " & dblTmp, vbOKOnly, "Xsol Aquecedores"
            Exit Sub
          End If
          
          If CDbl(txtVisaoEstoque_Estoque.Text) < txtSaidaQtd.Text Then
            MsgBox "A qtd máxima para alteração é " & txtVisaoEstoque_Estoque.Text, vbOKOnly, "Xsol Aquecedores"
            Exit Sub
          End If
        End If
        
        If Trim(txtSaidaQtd.Text) <> "0" And Trim(txtSaidaQtd.Text) <> "" Then
          If Trim(Combo_Tecnicos.Text) = "" Then
            MsgBox "Selecione um técnico responsável pela saída do Material.", vbOKOnly, "Xsol Aquecedores"
            Combo_Tecnicos.SetFocus
            Exit Sub
          End If
        End If
    End If
      
'
    
    If Trim(txtSaidaValor.Text) <> "" Then
        strValorSaida = txtSaidaValor.Text
    End If

   If Cmd_Inclui.Caption = "&Cancelar" Then 'Faz Insert
      
    Set rs = Return_Recordset(Return_SqlString(1))
    rs.AddNew
    
    'Pega o nome do produto pelo codigo
    
    rs.Fields("Produto") = cboVisaoEstoqueMaterial.Text
    
    rs.Fields("codMaterial") = lblCodMaterial.Caption
    
    If movimentoEntrada Then
        rs.Fields("Dt_Ent") = dtpEntradaData.value
        rs.Fields("Valor_Ent") = txtEntradaValor.Text
        rs.Fields("Qtd_Ent") = txtEntradaQtd.Text

    End If
    
    'Recupera o Saldo do produto
    'Call retornaQtdEntrada(Return_Recordset(Return_SqlString(6)))
    
    If movimentoEntrada Then
        rs.Fields("Saldo") = txtEntradaQtd.Text
    Else
        If CDbl(txtSaldo.Text) - CDbl(txtSaidaQtd.Text) >= 0 Then
            rs.Fields("Saldo") = CDbl(txtSaidaQtd.Text) * -1
        End If
        If Trim(txtSaidaQtd.Text) <> "0" And Trim(txtSaidaQtd.Text) <> "" Then
            rs.Fields("Dt_Saida") = dtpSaidaData.value
            rs.Fields("Valor_Saida") = txtSaidaValor.Text
        End If
        
        If Trim(txtSaidaQtd.Text) <> "" And Trim(txtSaidaQtd.Text) <> "0" Then
            rs.Fields("Qtd_Saida") = txtSaidaQtd.Text
        End If
    End If
    
    If Trim(Combo_Tecnicos.Text) <> "" Then
        rs.Fields("Nome_Tecnico") = Combo_Tecnicos.Text
    End If
    
    If ChkLiberado.value = 1 Then
        rs.Fields("Liberado") = "S"
    Else
        rs.Fields("Liberado") = "N"
    End If
    
    
      
    Cmd_Inclui.Caption = "&Inserir"
    Cmd_Inclui.enabled = True
    Cmd_Edita.enabled = True
    Cmd_Exclui.enabled = True
    Cmd_Salva.enabled = False
   'Else 'Faz update
    'Set rs = Return_Recordset(Return_SqlString(2, gridVisaoEstoqueProduto.TextMatrix(gridVisaoEstoqueProduto.RowSel, 0)))
   End If

'    txtProdutoDescricao.enabled = False
    'txtProdutoDescricao.BackColor = &HC0C0C0
    
    dtpEntradaData.enabled = False
    
    'txtEntradaValor.enabled = False
    'txtEntradaValor.BackColor = &HC0C0C0
    
    txtEntradaQtd.enabled = False
    txtEntradaQtd.BackColor = &HC0C0C0
    
    
    dtpSaidaData.enabled = False
    
    txtSaidaValor.enabled = False
    txtSaidaValor.BackColor = &HC0C0C0
    
    txtSaidaQtd.enabled = False
    txtSaidaQtd.BackColor = &HC0C0C0
    
    Combo_Tecnicos.enabled = False
    
    ChkLiberado.enabled = False
    
    Cmd_Inclui.enabled = True
    
    Cmd_Edita.Caption = "&Editar"
    Cmd_Edita.enabled = True
    
    Cmd_Exclui.enabled = True
    Cmd_Salva.enabled = False

   Call Record_Record(rs)
   rs.Update
   rs.Close
   Set rs = Nothing
   
   'Call Fill_Grid(Return_Recordset(Return_SqlString(9, txtProdutoDescricao.Text)), txtProdutoDescricao.Text)
   'Call VisaoEstoque_Show_Record_Detail_Saldo(Return_Recordset(Return_SqlString(9, txtProdutoDescricao.Text)))
   cboVisaoEstoqueMaterial_Click
    
    'txtProdutoDescricao.ListIndex = -1
    
    txtEntradaValor.Text = ""
    txtEntradaQtd.Text = ""
    txtSaidaValor.Text = ""
    txtSaidaQtd.Text = ""
    
    If Combo_Tecnicos.ListCount > 0 Then
        Combo_Tecnicos.ListIndex = -1
    Else
        Combo_Tecnicos.Text = ""
    End If
   
    txtEntradaValor.enabled = False
    txtEntradaQtd.enabled = False
    
    txtSaidaValor.Text = ""
    txtSaidaQtd.Text = ""
    txtSaidaValor.enabled = False
    txtSaidaQtd.enabled = False
   
    cboVisaoEstoqueMaterial.enabled = True
    gridVisaoEstoqueProduto.enabled = True
End Sub

Private Sub carregaMaterial(rs As ADODB.Recordset, Optional code As String)
    
    'tudo isto tinha sido comentado porque eu troquei o cbo de materiais por um lbl
    'txtProdutoDescricao.Clear

    If rs.RecordCount > 0 Then
        With txtProdutoDescricao
            If rs.EOF Or rs.BOF Then
               Exit Sub
            End If

            Do Until rs.EOF
                .AddItem (rs.Fields("Descricao") & "-" & rs.Fields("Codigo"))
                rs.MoveNext
            Loop
        End With
    'Else
        'txtProdutoDescricao.Clear
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub visaoEstoqueCarregaCboMaterial(rs As ADODB.Recordset, Optional code As String)
    
    cboVisaoEstoqueMaterial.Clear
    
    If rs.RecordCount > 0 Then
        With cboVisaoEstoqueMaterial
            If rs.EOF Or rs.BOF Then
               Exit Sub
            End If
    
            Do Until rs.EOF
                .AddItem (rs.Fields("Descricao"))
                rs.MoveNext
            Loop
        End With
    Else
        cboVisaoEstoqueMaterial.Clear
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Public Sub carregaTecnicos(rs As ADODB.Recordset, Optional code As String)
    'Dim selectedRow As Integer
    
    'selectedRow = 2
    
    With Combo_Tecnicos
        '.Rows = 2
        
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
End Sub

Private Sub Form_Activate()
    If (MDIForm_Xsol.acessibilidade_TextoNegrito = True) Or (MDIForm_Xsol.acessibilidade_TextoNegrito = False) Then
        txtProdutoDescricao.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        cboVisaoEstoqueMaterial.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Combo_Tecnicos.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    
    
        txtVisaoEstoque_Estoque.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        txtSaldo.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        txtProdutoDescricao.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        txtSaidaValor.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        txtSaidaQtd.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        txtEntradaQtd.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        txtEntradaValor.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        txtEstoqueCodigo.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        dtpSaidaData.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        dtpEntradaData.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        
        gridVisaoEstoqueProduto.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        lblCodMaterial.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label11.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        'Label14.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        'Label13.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label10.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label12.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label8.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label7.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label6.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label2.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label5.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Label9.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Cliente.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Lbl_Registro.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        btnFechar.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        btnRelGerar.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Inclui.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Salva.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Edita.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Cmd_Exclui.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        btnFechar.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        btnRelGerar.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        gridVisaoEstoqueProduto.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        ChkLiberado.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        optnRelRetiradaEquipamentos.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        optnRelSaidaEquipamentos.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        optnRelEntradaEquipamentos.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    End If

End Sub

Private Sub Form_Load()
   
    Call Configure_Grid
    
    Call carregaMaterial(Return_Recordset(Return_SqlString(3)))
    Call carregaTecnicos(Return_Recordset(Return_SqlString(4)))
    
    'Nova tela
    Call visaoEstoqueCarregaCboMaterial(Return_Recordset(Return_SqlString(3)))
    
    Call Fill_Grid_VisaoEstoque(Return_Recordset(Return_SqlString(9, lblCodMaterial.Caption)))
   
End Sub

Sub Show_Record_Detail(rs As ADODB.Recordset)

    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    txtEstoqueCodigo.Text = rs.Fields("Codigo")
    txtSaldo.Text = rs.Fields("Saldo")
    
    If txtProdutoDescricao.enabled = False Then
        txtProdutoDescricao.enabled = True
    End If
    
    'txtProdutoDescricao.Text = rs.Fields("Produto")
    
    'If rs.Fields("Dt_Ent") <> Null Then
    '    dtpEntradaData.value = rs.Fields("Dt_Ent")
    'End If
    
    txtEntradaValor.Text = rs.Fields("Valor_Ent")
    txtEntradaQtd.Text = rs.Fields("Qtd_Ent")
    
    
    
    'If rs.Fields("Dt_Saida") <> Null Then
    '    dtpSaidaData.value = rs.Fields("Dt_Saida")
    'End If
    
    If rs.Fields("Valor_Saida") <> Null Then
        txtSaidaValor.Text = rs.Fields("Valor_Saida")
    End If
    
    If rs.Fields("Qtd_Saida") <> Null Then
        txtSaidaQtd.Text = rs.Fields("Qtd_Saida")
    End If
    
    If Trim(rs.Fields("Nome_Tecnico")) <> "" Then
        Combo_Tecnicos.Text = rs.Fields("Nome_Tecnico")
    Else
        Combo_Tecnicos.ListIndex = -1
    End If

    If rs.Fields("Liberado") <> "" Then
        If UCase(rs.Fields("Liberado")) = "S" Then
            ChkLiberado.value = 1
        Else
            ChkLiberado.value = 0
        End If
    End If
    
    If rs.Fields("Valor_Saida") <> "" Then
        txtSaidaValor.Text = rs.Fields("Valor_Saida")
    Else
        txtSaidaValor.Text = ""
    End If
    
    If rs.Fields("Qtd_Saida") <> "" Then
        txtSaidaQtd.Text = rs.Fields("Qtd_Saida")
    Else
        txtSaidaQtd.Text = ""
    End If
        
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Grid_Estoque_Click()
    
'    With Grid_Estoque
'        If .RowSel = 1 Then
'            Exit Sub
'        End If
'        Call Show_Record_Detail(Return_Recordset(Return_SqlString(2, .TextMatrix(.RowSel, 0))))
'        Call Show_Record_Detail_Saldo(Return_Recordset(Return_SqlString(7, .TextMatrix(.RowSel, 1))))
'    End With

End Sub


'
Private Function retornaQtdEntrada(rs As ADODB.Recordset) As Double
    
    If rs.RecordCount > 0 Then
        With txtProdutoDescricao
            If rs.EOF Or rs.BOF Then
               Exit Function
            End If
    
            Do Until rs.EOF
                .AddItem (rs.Fields("Produto"))
                rs.MoveNext
            Loop
        End With
    Else
        txtProdutoDescricao.Clear
    End If
    
    rs.Close
    Set rs = Nothing
End Function
'

Private Sub gridVisaoEstoqueProduto_SelChange()
    txtEstoqueCodigo.Text = gridVisaoEstoqueProduto.TextMatrix(gridVisaoEstoqueProduto.RowSel, 9)
End Sub

Private Sub optnEstoqueGeral_Click()
    btnRelGerar.Tag = 4
End Sub

'Private Sub calculaPorcentagem(decValor As Double)
'
'    Dim valorEntrada As Double
'    Dim percentual As Double
'    Dim valorSaida As Double
'
'    If IsNumeric(decValor) Then
'            txtPorcentagemLucro.Text = Replace(txtPorcentagemLucro.Text, ".", ",")
'            If movimentoEntrada Then
'                If (Trim(txtEntradaValor.Text) <> "") And (IsNumeric(txtEntradaValor.Text)) Then
'                    valorEntrada = CDbl(txtEntradaValor.Text)
'                    percentual = CDbl(txtPorcentagemLucro.Text)
'                    valorSaida = valorEntrada + (valorEntrada * percentual / 100)
'                    txtSaidaValor.Text = Str(valorSaida)
'                Else
'                    MsgBox "Informe um valor válido para Entrada", vbOKOnly, "Xsol Aquecedores"
'                    Exit Sub
'                End If
'            Else
'                If (Trim(txtEntradaValor.Text) <> "") And (IsNumeric(txtEntradaValor.Text)) Then
'                    valorEntrada = CDbl(txtEntradaValor.Text)
'                    percentual = CDbl(txtPorcentagemLucro.Text)
'                    valorSaida = valorEntrada + (valorEntrada * percentual / 100)
'                    txtSaidaValor.Text = Str(valorSaida)
'                End If
'            End If
'    End If
'
'End Sub


Private Sub optnRelEntradaEquipamentos_Click()
    btnRelGerar.Tag = 1
End Sub

Private Sub optnRelRetiradaEquipamentos_Click()
    btnRelGerar.Tag = 3
End Sub

Private Sub optnRelSaidaEquipamentos_Click()
    btnRelGerar.Tag = 2
End Sub


Sub Show_Record_Detail_Saldo(rs As ADODB.Recordset)

'    If rs.EOF Or rs.BOF Then
'       txtSaldo.Text = "0"
'       Exit Sub
'    End If
'
'    If txtProdutoDescricao.enabled Then
'        If rs.Fields("Saldo") <> Null Or rs.Fields("Saldo") >= 0 Or rs.Fields("Saldo") < 0 Then
'            txtSaldo.Text = rs.Fields("Saldo")
'        End If
'    Else
'        If rs.Fields("Saldo") <> Null Or rs.Fields("Saldo") >= 0 Or rs.Fields("Saldo") < 0 Then
'            txtVisaoEstoque_Estoque.Text = rs.Fields("Saldo")
'        End If
'    End If
'
'    rs.Close
'    Set rs = Nothing
End Sub

Sub VisaoEstoque_Show_Record_Detail_Saldo(rs As ADODB.Recordset)

'    If rs.EOF Or rs.BOF Then
'       txtSaldo.Text = "0"
'       Exit Sub
'    End If
'
'    If rs.Fields("Saldo") <> Null Or rs.Fields("Saldo") >= 0 Or rs.Fields("Saldo") < 0 Then
'        txtVisaoEstoque_Estoque.Text = rs.Fields("Saldo")
'    Else
'        txtVisaoEstoque_Estoque.Text = "0"
'    End If
'
'    rs.Close
'    Set rs = Nothing
End Sub


Sub Show_Record_Detail_ValPorcent(rs As ADODB.Recordset)
    'verificar aqui se é entrada
    'só entra aqui se for entrada
    
    If Trim(txtEntradaValor.Text) = "" Then
       txtEntradaValor.Text = "0"
    End If
    
    
    If Trim(txtSaidaValor.Text) = "" Then
       txtSaidaValor.Text = "0"
    End If
    
    
    If (rs.EOF Or rs.BOF) Then
        Exit Sub
    End If
    
    
    If rs.Fields("Valor_Ent") <> Null Or rs.Fields("Valor_Ent") >= 0 Or rs.Fields("Valor_Ent") < 0 Then
        txtEntradaValor.Text = rs.Fields("Valor_Ent")
        
        
    End If
    
    
    rs.Close
    Set rs = Nothing
End Sub

Public Sub escolhaFornecedor(nomeProduto As String, codProduto As Integer)
                                                
    lblCodMaterial.Caption = Str(codProduto)
    frmEscolhaValorFornecedor.lblSobre.Caption = "O material " & nomeProduto & Chr(13) _
                                                & " possui mais de um fornecedor, escolha abaixo" _
                                                & " por qual deles será feito o movimento."
    frmEscolhaValorFornecedor.Show (1)

    txtEntradaValor.Text = entradaValor
    txtSaidaValor.Text = saidaValor
    
End Sub


Private Sub validaQtdFornecedor()
    Dim qtdFornecedor As Integer
    Dim vlrEntrada As String
    Dim vlrSaida As String
    
    'If Cmd_Inclui.Caption = "&Inserir" Then
        'Exit Sub
    'End If

    Call carregaCodMaterial(Return_Recordset(Return_SqlString(10, cboVisaoEstoqueMaterial.Text)))
    
    
    qtdFornecedor = qtdFornecedores(Return_Recordset(Return_SqlString(11, lblCodMaterial.Caption)))
    
    If qtdFornecedor > 0 Then
        If qtdFornecedor > 1 Then
            Call escolhaFornecedor(cboVisaoEstoqueMaterial.Text, Int(lblCodMaterial.Caption))
        Else
            If movimentoEntrada Then
                vlrEntrada = vlrFornecedorCompra(Return_Recordset(Return_SqlString(11, lblCodMaterial.Caption)))
                
                If (txtEntradaValor.Text = "") Or (txtEntradaValor.Text = "0") Then
                    txtEntradaValor.Text = vlrEntrada
                End If
            End If
            
            If movimentoSaida Then
                vlrSaida = vlrFornecedorVenda(Return_Recordset(Return_SqlString(11, lblCodMaterial.Caption)))
                txtSaidaValor.Text = vlrSaida
            End If
        End If
    End If

End Sub

Private Sub txtProdutoDescricao_old_Change()
    Call Show_Record_Detail_Saldo(Return_Recordset(Return_SqlString(5, txtProdutoDescricao.Text)))
    Call Show_Record_Detail_ValPorcent(Return_Recordset(Return_SqlString(8, txtProdutoDescricao.Text)))
End Sub


Private Function retornaUltimoValorDeSaida(rs As ADODB.Recordset) As Double
    
    If IsNull((rs.Fields("valor_saida"))) Then
        retornaUltimoValorDeSaida = 0
    Else
        retornaUltimoValorDeSaida = rs.Fields("valor_saida")
    End If
    
    rs.Close
    Set rs = Nothing
End Function

    
Private Function retornaUltimoValorDeEntrada(rs As ADODB.Recordset) As Double
    
    If IsNull((rs.Fields("valor_ent"))) Then
        retornaUltimoValorDeEntrada = 0
    Else
        retornaUltimoValorDeEntrada = rs.Fields("valor_ent")
    End If
    
    rs.Close
    Set rs = Nothing
End Function

