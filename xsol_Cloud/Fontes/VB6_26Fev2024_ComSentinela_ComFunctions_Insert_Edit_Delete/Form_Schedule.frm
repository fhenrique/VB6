VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Schedule 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agenda"
   ClientHeight    =   10455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10455
   ScaleWidth      =   19530
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Informações do Cliente"
      Height          =   1620
      Left            =   120
      TabIndex        =   27
      Top             =   2520
      Width           =   19275
      Begin VB.Label Lbl_Cidade 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   11880
         TabIndex        =   60
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Cidade:"
         Height          =   255
         Left            =   11160
         TabIndex        =   59
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Lbl_Ref 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   450
         Left            =   6705
         TabIndex        =   58
         Top             =   1140
         Width           =   7695
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Ref.:"
         Height          =   255
         Left            =   6240
         TabIndex        =   57
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label Lbl_Zona 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   11880
         TabIndex        =   56
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Zona:"
         Height          =   255
         Left            =   11280
         TabIndex        =   55
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Lbl_Bairro 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   6720
         TabIndex        =   54
         Top             =   690
         Width           =   4455
      End
      Begin VB.Label Lbl_Equip 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   1200
         TabIndex        =   53
         Top             =   1140
         Width           =   4815
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Equipamento:"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Bairro:"
         Height          =   255
         Left            =   6120
         TabIndex        =   51
         Top             =   690
         Width           =   495
      End
      Begin VB.Label Lbl_Cliente_Cel 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   15105
         TabIndex        =   50
         Top             =   1140
         Width           =   3840
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Celular:"
         Height          =   255
         Left            =   14400
         TabIndex        =   49
         Top             =   1140
         Width           =   615
      End
      Begin VB.Label Lbl_Cliente_Fone_Res 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   15120
         TabIndex        =   48
         Top             =   720
         Width           =   3675
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Fone Res.:"
         Height          =   255
         Left            =   14160
         TabIndex        =   47
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Lbl_Cliente_Fone_Com 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   15120
         TabIndex        =   46
         Top             =   360
         Width           =   3435
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Fone Com.:"
         Height          =   255
         Left            =   14160
         TabIndex        =   45
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_Cliente 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1200
         TabIndex        =   44
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome:"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Lbl_Cliente_Cod 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1200
         TabIndex        =   42
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Código:"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lbl_End 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   465
         Left            =   6705
         TabIndex        =   29
         Top             =   210
         Width           =   4455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "End.:"
         Height          =   255
         Left            =   6300
         TabIndex        =   28
         Top             =   315
         Width           =   375
      End
   End
   Begin VB.Frame Frame_Serv_Info 
      Caption         =   "Informações do Atendimento"
      Height          =   2505
      Left            =   6240
      TabIndex        =   19
      Top             =   0
      Width           =   11535
      Begin VB.Label Lbl_Dt_Status 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4440
         TabIndex        =   63
         Top             =   375
         Width           =   975
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Status:"
         Height          =   255
         Left            =   3360
         TabIndex        =   62
         Top             =   375
         Width           =   975
      End
      Begin VB.Label Lbl_Status 
         Caption         =   "(Status)"
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
         Left            =   1800
         TabIndex        =   61
         Top             =   375
         Width           =   1455
      End
      Begin VB.Label Lbl_Tecnico_Cel 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   7080
         TabIndex        =   40
         Top             =   960
         Width           =   4230
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Celular Técnico:"
         Height          =   255
         Left            =   5760
         TabIndex        =   39
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Lbl_Tecnico_Fone 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   7080
         TabIndex        =   38
         Top             =   675
         Width           =   4230
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Fone Técnico:"
         Height          =   255
         Left            =   5880
         TabIndex        =   37
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label Lbl_Obs 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   750
         Left            =   1080
         TabIndex        =   36
         Top             =   1575
         Width           =   10050
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Obs.:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Lbl_Valor 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1080
         TabIndex        =   34
         Top             =   1275
         Width           =   2415
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor:"
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   1275
         Width           =   495
      End
      Begin VB.Label Lbl_Tipo_Serv 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3000
         TabIndex        =   32
         Top             =   675
         Width           =   2895
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Serv.:"
         Height          =   255
         Left            =   2040
         TabIndex        =   31
         Top             =   675
         Width           =   855
      End
      Begin VB.Label Lbl_Tecnico 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   7080
         TabIndex        =   30
         Top             =   390
         Width           =   4230
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Técnico:"
         Height          =   255
         Left            =   6360
         TabIndex        =   26
         Top             =   375
         Width           =   615
      End
      Begin VB.Label Lbl_Periodo 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Período:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Lbl_Dt_Serv 
         Caption         =   "End:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   675
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Serv.:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   675
         Width           =   855
      End
      Begin VB.Label Lbl_OS 
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
         Left            =   1080
         TabIndex        =   21
         Top             =   375
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "OS:"
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
         Left            =   360
         TabIndex        =   20
         Top             =   375
         Width           =   615
      End
   End
   Begin VB.Frame Frame_Result 
      Caption         =   "Resultado"
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   9690
      Width           =   19320
      Begin VB.Label Lbl_Registros_Result 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   4200
         TabIndex        =   18
         Top             =   165
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Registros:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   17
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Lbl_Total_Result 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   14160
         TabIndex        =   16
         Top             =   165
         Width           =   3015
      End
      Begin VB.Label Lbl_Total 
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12960
         TabIndex        =   15
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Atendimentos"
      Height          =   5475
      Left            =   120
      TabIndex        =   12
      Top             =   4170
      Width           =   19305
      Begin MSFlexGridLib.MSFlexGrid Grid_Schedule 
         Height          =   5160
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   19050
         _ExtentX        =   33602
         _ExtentY        =   9102
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame_Print 
      Caption         =   "Impressão"
      Height          =   2520
      Left            =   17880
      TabIndex        =   10
      Top             =   0
      Width           =   1500
      Begin VB.CommandButton Cmd_Print 
         Caption         =   "&Imprimir"
         Height          =   390
         Left            =   150
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pesquisa"
      Height          =   2505
      Left            =   135
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton btnTelaDiasIndisponiveis 
         Caption         =   "Tabela de dias indisponíveis"
         Height          =   450
         Left            =   285
         TabIndex        =   66
         Top             =   1755
         Width           =   5415
      End
      Begin MSComCtl2.DTPicker Mask_Dt_Inicial 
         Height          =   285
         Left            =   1170
         TabIndex        =   64
         Top             =   750
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Format          =   117374977
         CurrentDate     =   44035
      End
      Begin VB.CommandButton But_Search 
         Height          =   420
         Left            =   5280
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Form_Schedule.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Pesquisar"
         Top             =   825
         Width           =   435
      End
      Begin VB.ComboBox Combo_Tecnico 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   315
         Width           =   3495
      End
      Begin VB.CommandButton But_Search_Clean 
         Height          =   435
         Left            =   5280
         Picture         =   "Form_Schedule.frx":0942
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Limpar"
         Top             =   300
         Width           =   435
      End
      Begin VB.TextBox Text_Tecnico_Cod 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   315
         Width           =   415
      End
      Begin MSMask.MaskEdBox Mask_Dt_Inicial2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   3570
         TabIndex        =   5
         Top             =   765
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   " "
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mask_Dt_Final2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   3570
         TabIndex        =   6
         Top             =   1155
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker Mask_Dt_Final 
         Height          =   285
         Left            =   1185
         TabIndex        =   65
         Top             =   1140
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Format          =   117374977
         CurrentDate     =   44035
      End
      Begin VB.Label Lbl_DtFinal 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Final:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label Lbl_DtInicial 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Inicial:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   750
         Width           =   855
      End
      Begin VB.Label Lbl_Tec 
         Alignment       =   1  'Right Justify
         Caption         =   "Técnico:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   315
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form_Schedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private W_DB As ADODB.Connection
Private bEditando As Boolean


'Private Function Return_SqlString(queryType As Integer, Optional param1 As String) As String
'    Dim whereClause As String
'
'    If queryType = 1 Then
'        Return_SqlString = "SELECT * FROM Tecnico"
'    ElseIf queryType = 2 Then
'        Return_SqlString = "SELECT * FROM Tecnico WHERE Codigo = " & param1
'    End If
'End Function

Private Sub Record_Record(rs As ADODB.Recordset)
   'rs.Fields("DiasNaoDisponivel").value = UCase(txtCalendarioDias.Text)
   'rs.Fields("MesNaoDisponivel").value = UCase(cboMeses.Text)
End Sub

Private Sub btnDiasIndisponiveis_Click()
    
End Sub

Private Sub btnTelaDiasIndisponiveis_Click()
    
    If Combo_Tecnico.Text <> "" Then
        'Verifica se existe agenda dos 12 meses para o técnico
        gestaoDeAgenda (Text_Tecnico_Cod.Text)
        
        Sleep (1500)
        
        frmDiasIndisponiveis.nomeTecnico = Combo_Tecnico.Text
        frmDiasIndisponiveis.codTecnico = Text_Tecnico_Cod.Text
        frmDiasIndisponiveis.Show (1)
    Else
        MsgBox "Selecione 1 técnico", vbOKCancel, "XSol Aquecedores"
    End If
    
End Sub



Private Sub gestaoDeAgenda(codTecnico As String)
On Error GoTo erro
    Dim rs As ADODB.Recordset
    Dim contador As Integer
    
    'Set rs = Return_Recordset(Return_SqlString(8, Text_Tecnico_Cod.Text))
    cloud_sql = "SELECT * FROM `tecnicos_agenda` WHERE codigo = '" & Text_Tecnico_Cod.Text & "'"
    
retenta:
    Set cloud_registro = cloud_conexao.Execute("SELECT * FROM `tecnicos_agenda` WHERE codigo = '" & Text_Tecnico_Cod.Text & "'")
    
    Dim aMes(12) As String
    
    aMes(1) = "JAN"
    aMes(2) = "FEV"
    aMes(3) = "MAR"
    aMes(4) = "ABR"
    aMes(5) = "MAI"
    aMes(6) = "JUN"
    aMes(7) = "JUL"
    aMes(8) = "AGO"
    aMes(9) = "SET"
    aMes(10) = "OUT"
    aMes(11) = "NOV"
    aMes(12) = "DEZ"
    
    If cloud_registro.RecordCount = 0 Then 'Se for vazio o técnico não tem agenda ainda, então crio até dezembro
        For contador = 1 To 12
            cloud_sql = "insert into `tecnicos_agenda` (codigo, MES) Values(" & codTecnico & ",'" & aMes(contador) & "')"
            cloud_conexao.Execute (cloud_sql)
        Next contador
    End If
    cloud_registro.Close
    Exit Sub
erro:
    If cloud_conexao.State = 1 Then
        cloud_conexao.Close
        If ConectaMySQL Then
            GoTo retenta
        Else
            MsgBox "Erro conectando ao Cloud, favor reiniciar o sistema!", vbCritical, "XSOL Aquecedores"
            End
        End If
    End If
End Sub






Private Sub Form_Activate()
    If (MDIForm_Xsol.acessibilidade_TextoNegrito = True) Or (MDIForm_Xsol.acessibilidade_TextoNegrito = False) Then
        Combo_Tecnico.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        'cboMeses.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Mask_Dt_Inicial.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Mask_Dt_Final.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        'txtCalendarioDias.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
        Text_Tecnico_Cod.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    End If

End Sub

Private Sub Form_Load()
    
    
    Call grant_access
    Call Combo_Initialize
    Call Date_Initialize
    Call Configure_Grid
    Call Clean_Controls
    
    Mask_Dt_Inicial.value = Format(DateTime.Now, "Short Date")
    Mask_Dt_Final.value = Format(DateTime.Now + 30, "Short Date")
    
    Me.Height = 11055
    
End Sub

Private Sub grant_access()
    Frame_Print.Visible = isAdmin
    Lbl_Total.Visible = isAdmin
    Lbl_Total_Result.Visible = isAdmin
    'Frame_Serv_Info.Width = IIf(isAdmin, 9735, 11535)
End Sub

Private Sub Fill_Grid(rs As ADODB.Recordset, Optional code As String)

    Dim count As Integer
    Dim valSum As Double
    Dim bkColor As Long
    
    count = 0
    With Grid_Schedule
        .Rows = 2
        Do Until rs.EOF
            count = count + 1
            .AddItem (rs.Fields("codigo") & vbTab & rs.Fields("Dt_Atend_Serv") & vbTab & rs.Fields("Hr_Atend_Serv") & vbTab & rs.Fields("Nome_Cli") & vbTab & rs.Fields("Endereco") & " - " & rs.Fields("Bairro") & vbTab & rs.Fields("Equipamento") & vbTab & rs.Fields("Tipo_Serv") & vbTab & rs.Fields("Nome") & vbTab & rs.Fields("Valor_Serv"))
            If rs.Fields("Status") <> "CANCELADO" Then
                valSum = valSum + rs.Fields("Valor_Serv")
            End If
            bkColor = GET_BACKCOLOR_BY_STATUS(rs.Fields("Status"))
            If bkColor <> 0 Then
                .row = .Rows - 1
                For count = .FixedCols To .Cols - 1
                    .col = count
                    .CellBackColor = bkColor
                Next
            End If
            rs.MoveNext
        Loop
        
        .Refresh
        
        cloud_qtd = cloud_registro.RecordCount
        
        If cloud_registro.RecordCount > 0 Then
            cloud_registro.MoveFirst
        End If
        If Not cloud_registro.EOF Then
            .col = 0
            .ColSel = .Cols - 1
        End If
    End With
    
    Lbl_Registros_Result.Caption = Str(cloud_qtd)
    
    
    Lbl_Total_Result.Caption = "R$ " + Format(valSum, "###,###,##0.00")
    Exit Sub
    

    
End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String, Optional param3 As String) As String
    Dim whereClause As String
    
    
    If queryType = 1 Then
        Return_SqlString = "SELECT a.Codigo, a.Dt_Atend_Serv, a.Hr_Atend_Serv, c.Nome_Cli, e.Endereco, e.Bairro, eq.Equipamento, a.Tipo_Serv,a.Tipo_Serv2, t.Nome, a.Valor_Serv, a.Status"
        Return_SqlString = Return_SqlString & " FROM (((`Atendimento` a INNER JOIN `Clientes` c ON a.Cod_Cliente = c.Codigo) INNER JOIN `Tecnico` AS t ON  a.Cod_Tecnico = t.Codigo) INNER JOIN `Endereco` e ON a.Cod_End = e.Codigo) LEFT OUTER JOIN `Equipamento` eq ON a.Cod_Equip = eq.Codigo"
        
        If param1 <> "" Then
            whereClause = " WHERE a.Cod_Tecnico = " & param1
        End If
        If param2 <> "__/__/____" Then
            If whereClause = "" Then
                whereClause = " WHERE"
            Else
                whereClause = whereClause & " AND"
            End If
            'MsgBox Format(param2, "yyyy-MM-dd")
            whereClause = whereClause & " a.Dt_Atend_Serv >= '" & Format(param2, "yyyy-MM-dd") & "'"
        End If
        If param3 <> "__/__/____" Then
            If whereClause = "" Then
                whereClause = " WHERE"
            Else
                whereClause = whereClause & " AND"
            End If
            whereClause = whereClause & " a.Dt_Atend_Serv <= '" & Format(param3, "yyyy-MM-dd") & "'"
        End If
        whereClause = whereClause & " ORDER BY a.Dt_Atend_Serv"
        Return_SqlString = Return_SqlString & whereClause
        'Debug.Print (Return_SqlString)
    ElseIf queryType = 2 Then
        Return_SqlString = "SELECT Nome AS Val FROM `Tecnico` where status = 'A' "
    ElseIf queryType = 3 Then
        Return_SqlString = "SELECT * FROM `Tecnico` WHERE Nome = '" & param1 & "'"
    ElseIf queryType = 4 Then
        Return_SqlString = "SELECT a.Codigo AS OS, a.Dt_Atend_Serv, a.Hr_Atend_Serv, a.Tipo_Serv,a.Tipo_Serv2, a.Valor_Serv AS Valor, a.Status, a.Dt_Status, a.Obs, t.Nome AS Tecnico, t.Fone AS Tecnico_Fone, t.Celular AS Tecnico_Cel, c.Codigo, c.Nome_Cli, c.Fone_Com, c.Fone_Res, c.Fone_Celular, e.Endereco, e.Cidade, e.Bairro, e.Zona, e.Ref, eq.Equipamento"
        Return_SqlString = Return_SqlString & " FROM (((`Atendimento` a INNER JOIN `Clientes` c ON a.Cod_Cliente = c.Codigo) INNER JOIN `Tecnico` AS t ON  a.Cod_Tecnico = t.Codigo) INNER JOIN `Endereco` e ON a.Cod_End = e.Codigo) LEFT OUTER JOIN `Equipamento` eq ON a.Cod_Equip = eq.Codigo"
        Return_SqlString = Return_SqlString & " WHERE a.Codigo = " & param1
        'Debug.Print (Return_SqlString)
    ElseIf queryType = 5 Then
        Return_SqlString = "SELECT SUM(a.Valor_Serv)"
        Return_SqlString = Return_SqlString & " FROM `Atendimento` a"
        whereClause = " WHERE a.Status <> ""CANCELADO"" "
        
        If param1 <> "" Then
            whereClause = whereClause & " AND a.Cod_Tecnico = " & param1
        End If
        If param2 <> "__/__/____" Then
            whereClause = whereClause & " AND a.Dt_Atend_Serv >= #" & CDate(Format(param2, "mm\/dd\/yyyy")) & "#"
        End If
        If param3 <> "__/__/____" Then
            whereClause = whereClause & " AND a.Dt_Atend_Serv <= #" & CDate(Format(param3, "mm\/dd\/yyyy")) & "#"
        End If
        Return_SqlString = Return_SqlString & whereClause
    ElseIf queryType = 6 Then
        Return_SqlString = "SELECT DiasNaoDisponivel, MesNaoDisponivel FROM `Tecnico` WHERE Codigo = " & param1
    ElseIf queryType = 7 Then
        Return_SqlString = "SELECT * FROM `Tecnico` WHERE Nome = '" & param1 & "' And mesnaodisponivel = '" & param2 & "'"
    ElseIf queryType = 8 Then
        Return_SqlString = "SELECT * FROM `tecnicos_agenda` WHERE codigo = '" & param1 & "'"
    End If
End Function

Sub Show_Record_Detail(rs As ADODB.Recordset)

    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If
    
    Lbl_OS = rs.Fields("OS")
    Lbl_Status = "(" & rs.Fields("Status") & ")"
    Lbl_Dt_Status = Get_Field_Val(rs.Fields("Dt_Status"))
    Lbl_Dt_Serv = Get_Field_Val(rs.Fields("Dt_Atend_Serv"))
    Lbl_Periodo = Get_Field_Val(rs.Fields("Hr_Atend_Serv"))
    Lbl_Tipo_Serv = Get_Field_Val(rs.Fields("Tipo_Serv"))
    Lbl_Valor = Get_Field_Val(rs.Fields("Valor"))
    Lbl_Obs = Get_Field_Val(rs.Fields("Obs"))
    Lbl_Tecnico = Get_Field_Val(rs.Fields("Tecnico"))
    Lbl_Tecnico_Fone = Get_Field_Val(rs.Fields("Tecnico_Fone"))
    Lbl_Tecnico_Cel = Get_Field_Val(rs.Fields("Tecnico_Cel"))
    Lbl_Cliente_Cod = rs.Fields("Codigo")
    Lbl_Cliente = Get_Field_Val(rs.Fields("Nome_Cli"))
    Lbl_Equip = Get_Field_Val(rs.Fields("Equipamento"))
    Lbl_End = Get_Field_Val(rs.Fields("Endereco"))
    Lbl_Bairro = Get_Field_Val(rs.Fields("Bairro"))
    Lbl_Ref = Get_Field_Val(rs.Fields("Ref"))
    Lbl_Cidade = Get_Field_Val(rs.Fields("Cidade"))
    Lbl_Zona = Get_Field_Val(rs.Fields("Zona"))
    Lbl_Cliente_Fone_Com = Get_Field_Val(rs.Fields("Fone_Com"))
    Lbl_Cliente_Fone_Res = Get_Field_Val(rs.Fields("Fone_Res"))
    Lbl_Cliente_Cel = Get_Field_Val(rs.Fields("Fone_Celular"))
            
    rs.Close
    Set rs = Nothing
End Sub

Private Function Return_Recordset(strSql As String) As ADODB.Recordset
On Error GoTo erro
    'Dim rs As New ADODB.Recordset
    'rs.Open strSql, W_DB, adOpenStatic, adLockOptimistic
retenta:
    Set cloud_registro = cloud_conexao.Execute(strSql)
    Set Return_Recordset = cloud_registro
    Exit Function
    
erro:
    If cloud_conexao.State = 1 Then
        cloud_conexao.Close
    End If
    
    If ConectaMySQL Then
        GoTo retenta
    Else
        MsgBox "Desconectado do banco na nuvem!", vbCritical, "XSOL Aquecedores"
        End
    End If
    
End Function

Sub Clean_Controls()
    Lbl_OS = ""
    Lbl_Status = ""
    Lbl_Dt_Status = ""
    Lbl_Dt_Serv = ""
    Lbl_Periodo = ""
    Lbl_Tipo_Serv = ""
    Lbl_Valor = ""
    Lbl_Obs = ""
    Lbl_Tecnico = ""
    Lbl_Tecnico_Fone = ""
    Lbl_Tecnico_Cel = ""
    Lbl_Cliente_Cod = ""
    Lbl_Cliente = ""
    Lbl_Equip = ""
    Lbl_End = ""
    Lbl_Bairro = ""
    Lbl_Ref = ""
    Lbl_Cidade = ""
    Lbl_Zona = ""
    Lbl_Cliente_Fone_Com = ""
    Lbl_Cliente_Fone_Res = ""
    Lbl_Cliente_Cel = ""
End Sub

Private Sub Configure_Grid()
    With Grid_Schedule
        .Cols = 9
        .FixedCols = 0
        .ColWidth(0) = 0
        .ColWidth(1) = 1000
        .ColWidth(2) = 1100
        .ColWidth(3) = 3100
        .ColWidth(4) = 4800
        .ColWidth(5) = 2600
        .ColWidth(6) = 2000
        .ColWidth(7) = IIf(isAdmin, 1650, 2450)
        .ColWidth(8) = IIf(isAdmin, 800, 0)
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignLeftCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "OS"
        .TextMatrix(0, 1) = "Data"
        .TextMatrix(0, 2) = "Período"
        .TextMatrix(0, 3) = "Cliente"
        .TextMatrix(0, 4) = "Endereço"
        .TextMatrix(0, 5) = "Equipamento"
        .TextMatrix(0, 6) = "Tipo de Serviço"
        .TextMatrix(0, 7) = "Técnico"
        .TextMatrix(0, 8) = "Valor"
        .RowHeight(1) = 0
    End With
End Sub

Sub Combo_Initialize()
    'Dim rs As ADODB.Recordset
    'Set rs = Return_Recordset(Return_SqlString(2))
    
    cloud_sql = "SELECT Nome AS Val FROM Tecnico where status = 'A'"
    Set cloud_registro = cloud_conexao.Execute(cloud_sql)
    
    
    If cloud_registro.EOF Then
       Exit Sub
    End If
    
    Combo_Tecnico.Clear
    Do While Not cloud_registro.EOF
       Combo_Tecnico.AddItem cloud_registro.Fields("Val")
       cloud_registro.MoveNext
    Loop
End Sub

Sub Date_Initialize()
    'Mask_Dt_Inicial.Text = Format(Date, "dd\/mm\/yyyy")
    'Mask_Dt_Final.Text = Format(DateAdd("m", 1, Date), "dd\/mm\/yyyy")
End Sub

Private Sub Cmd_Print_Click()
    Dim rs As ADODB.Recordset
    Dim I As Integer
    
    Set rs = Return_Recordset(Return_SqlString(1, Text_Tecnico_Cod, Mask_Dt_Inicial, Mask_Dt_Final))
    With DataReport1
        Set .DataSource = rs
        .DataMember = ""
        .Show
    End With
    
    With DataReport1.Sections(5).Controls
        For I = 1 To .count
            If UCase(.Item(I).Name) = "LBL_REG_QTT" Then
                .Item(I).Caption = rs.RecordCount
                Exit For
            End If
        Next I
    End With
    
    Set rs = Return_Recordset(Return_SqlString(5, Text_Tecnico_Cod, Mask_Dt_Inicial, Mask_Dt_Final))
    
    With DataReport1.Sections(5).Controls
        For I = 1 To .count
            If UCase(.Item(I).Name) = "LBL_TOT" Then
                .Item(I).Caption = "R$ " + Format(rs.Fields(0), "###,###,##0.00")
                Exit For
            End If
        Next I
    End With
    
    With DataReport1.Sections(1).Controls
        For I = 1 To .count
            If UCase(.Item(I).Name) = "LBL_PERIOD" Then
                .Item(I).Caption = Str(Mask_Dt_Inicial.value) + " - " + Str(Mask_Dt_Final.value)
                Exit For
            End If
        Next I
    End With
End Sub

Private Sub Combo_Tecnico_Click()
    retornaDadosDoTecnico (Combo_Tecnico.Text)
End Sub

Private Sub retornaDadosDoTecnico(nomeDoTecnico As String)
On Error GoTo erro
    
    If nomeDoTecnico = "" Then
        Exit Sub
    End If
    
    'Set rs = Return_Recordset(Return_SqlString(3, nomeDoTecnico))
    cloud_sql = "SELECT * FROM Tecnico WHERE Nome = '" & nomeDoTecnico & "'"
    
retenta:
    Set cloud_registro = cloud_conexao.Execute(cloud_sql)
    
    Text_Tecnico_Cod.Text = cloud_registro.Fields("Codigo")
    Exit Sub
    
erro:

    If cloud_conexao.State = 1 Then
        cloud_conexao.Close
    End If
    
    If ConectaMySQL Then
        GoTo retenta
    Else
        MsgBox "Desconectado do banco na nuvem!", vbCritical, "XSOL Aquecedores"
        End
    End If


End Sub


Private Sub But_Search_Click()
    With Grid_Schedule
        Call Fill_Grid(Return_Recordset(Return_SqlString(1, Text_Tecnico_Cod, Mask_Dt_Inicial, Mask_Dt_Final)))
    End With
    Call Clean_Controls
End Sub

Private Sub Grid_Schedule_Click()
    With Grid_Schedule
        If .RowSel = 1 Then
            Exit Sub
        End If
        Call Show_Record_Detail(Return_Recordset(Return_SqlString(4, .TextMatrix(.RowSel, 0))))
    End With
End Sub

Private Sub But_Search_Clean_Click()
    Text_Tecnico_Cod = ""
    Combo_Tecnico.ListIndex = -1
    'Mask_Dt_Inicial.Text = "__/__/____"
    'Mask_Dt_Final.Text = "__/__/____"
End Sub

Private Sub Grid_Schedule_SelChange()
   Call Grid_Schedule_Click
End Sub



Private Sub Mask_Dt_Final_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call But_Search_Click
    End If
End Sub

Private Sub Mask_Dt_Inicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call But_Search_Click
    End If
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)

End Sub

