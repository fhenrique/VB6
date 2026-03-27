VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form_Technician 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Técnicos"
   ClientHeight    =   8295
   ClientLeft      =   2520
   ClientTop       =   1365
   ClientWidth     =   16905
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   16905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Controles"
      Height          =   975
      Left            =   9600
      TabIndex        =   31
      Top             =   7230
      Width           =   7245
      Begin VB.CommandButton But_Delete 
         BackColor       =   &H008080FF&
         Caption         =   "E&xcluir"
         Height          =   495
         Left            =   5700
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton But_Insert 
         BackColor       =   &H0080FF80&
         Caption         =   "&Incluir"
         Height          =   495
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton But_Save 
         Caption         =   "&Salvar"
         Height          =   495
         Left            =   3825
         TabIndex        =   33
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton But_Edit 
         Caption         =   "&Editar"
         Height          =   495
         Left            =   2070
         TabIndex        =   32
         Top             =   315
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lista de Técnicos"
      Height          =   7065
      Left            =   9615
      TabIndex        =   25
      Top             =   120
      Width           =   7215
      Begin MSFlexGridLib.MSFlexGrid Grid_Technician 
         Height          =   6705
         Left            =   135
         TabIndex        =   26
         Top             =   240
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   11827
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin VB.Frame But_ChooseTecPhotoBut_ChooseTecPhoto 
      Caption         =   "Dados do Técnico"
      Height          =   8070
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   9345
      Begin VB.Frame Frame3 
         Caption         =   "Veículo"
         Height          =   915
         Left            =   5370
         TabIndex        =   57
         Top             =   6360
         Width           =   3795
         Begin VB.ComboBox cboVeiculos 
            Height          =   315
            ItemData        =   "Form_Tecnicos.frx":0000
            Left            =   165
            List            =   "Form_Tecnicos.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   360
            Width           =   3420
         End
      End
      Begin VB.TextBox Text_Email 
         Height          =   285
         Left            =   1425
         TabIndex        =   56
         ToolTipText     =   "Nome"
         Top             =   855
         Width           =   6195
      End
      Begin VB.Frame Frame1 
         Caption         =   "Fotos"
         Height          =   3195
         Left            =   4365
         TabIndex        =   44
         Top             =   1170
         Width           =   4785
         Begin VB.TextBox Text_AuxTec_Photo 
            BackColor       =   &H80000004&
            CausesValidation=   0   'False
            Enabled         =   0   'False
            Height          =   285
            Left            =   780
            TabIndex        =   53
            ToolTipText     =   "Cnpj"
            Top             =   705
            Width           =   2655
         End
         Begin VB.CommandButton But_ChooseAuxTecPhoto 
            Caption         =   "..."
            Height          =   315
            Left            =   3540
            TabIndex        =   52
            Top             =   705
            Width           =   375
         End
         Begin VB.CommandButton But_AuxTecPhoto_Clean 
            Height          =   315
            Left            =   4020
            Picture         =   "Form_Tecnicos.frx":0004
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   705
            Width           =   375
         End
         Begin VB.TextBox Text_Tec_Photo 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   285
            Left            =   780
            TabIndex        =   49
            ToolTipText     =   "Cnpj"
            Top             =   315
            Width           =   2655
         End
         Begin VB.CommandButton But_ChooseTecPhoto 
            Caption         =   "..."
            Height          =   315
            Left            =   3540
            TabIndex        =   48
            Top             =   315
            Width           =   375
         End
         Begin VB.CommandButton But_TecPhoto_Clean 
            Height          =   315
            Left            =   4020
            Picture         =   "Form_Tecnicos.frx":03E8
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   315
            Width           =   375
         End
         Begin VB.Image Pic_AuxTec 
            BorderStyle     =   1  'Fixed Single
            Height          =   1695
            Left            =   2655
            Stretch         =   -1  'True
            Top             =   1140
            Width           =   1455
         End
         Begin VB.Image Pic_Tec 
            BorderStyle     =   1  'Fixed Single
            Height          =   1695
            Left            =   645
            Stretch         =   -1  'True
            Top             =   1155
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Auxiliar"
            Height          =   195
            Left            =   195
            TabIndex        =   54
            Top             =   765
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Técnico"
            Height          =   195
            Left            =   105
            TabIndex        =   50
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Aux. Técnico"
            Height          =   195
            Left            =   2925
            TabIndex        =   46
            Top             =   2880
            Width           =   945
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Técnico"
            Height          =   195
            Left            =   1080
            TabIndex        =   45
            Top             =   2880
            Width           =   585
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8085
         Top             =   690
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox Comb_Status 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Status"
         Top             =   3795
         Width           =   2655
      End
      Begin VB.Frame Frame_Aux 
         Caption         =   "Auxiliar Técnico"
         Height          =   1515
         Left            =   255
         TabIndex        =   39
         Top             =   6360
         Width           =   4935
         Begin VB.TextBox Text_Nome_Aux 
            Height          =   285
            Left            =   1170
            TabIndex        =   15
            ToolTipText     =   "Nome do Auxiliar"
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox Text_Rg_Aux 
            Height          =   285
            Left            =   1170
            TabIndex        =   16
            ToolTipText     =   "Rg do Auxiliar"
            Top             =   720
            Width           =   2655
         End
         Begin VB.TextBox Text_Cpf_Aux 
            Height          =   285
            Left            =   1170
            TabIndex        =   17
            ToolTipText     =   "Cpf do Auxiliar"
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label Lbl_Nome_Aux 
            Alignment       =   1  'Right Justify
            Caption         =   "Nome:"
            Height          =   255
            Left            =   195
            TabIndex        =   42
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Rg:"
            Height          =   255
            Left            =   555
            TabIndex        =   41
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Lbl_Cpf_Aux 
            Alignment       =   1  'Right Justify
            Caption         =   "Cpf:"
            Height          =   255
            Left            =   435
            TabIndex        =   40
            Top             =   1080
            Width           =   615
         End
      End
      Begin VB.Frame Frame_End 
         Caption         =   "Endereço"
         Height          =   1935
         Left            =   255
         TabIndex        =   29
         Top             =   4380
         Width           =   8940
         Begin VB.TextBox Text_Cep 
            Height          =   285
            Left            =   1185
            MaxLength       =   9
            TabIndex        =   14
            ToolTipText     =   "Cep"
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox Text_Cidade 
            Height          =   285
            Left            =   1185
            TabIndex        =   13
            ToolTipText     =   "Cidade"
            Top             =   1080
            Width           =   4575
         End
         Begin VB.TextBox Text_Bairro 
            Height          =   285
            Left            =   1185
            TabIndex        =   12
            ToolTipText     =   "Bairro"
            Top             =   720
            Width           =   4575
         End
         Begin VB.TextBox Text_End 
            Height          =   285
            Left            =   1200
            TabIndex        =   11
            ToolTipText     =   "Logradouro"
            Top             =   360
            Width           =   6270
         End
         Begin VB.Label Lbl_Cep 
            Alignment       =   1  'Right Justify
            Caption         =   "Cep:"
            Height          =   255
            Left            =   705
            TabIndex        =   38
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Lbl_Cidade 
            Alignment       =   1  'Right Justify
            Caption         =   "Cidade:"
            Height          =   255
            Left            =   465
            TabIndex        =   37
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Lbl_Bairro 
            Alignment       =   1  'Right Justify
            Caption         =   "Bairro:"
            Height          =   255
            Left            =   585
            TabIndex        =   36
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Lbl_End 
            Alignment       =   1  'Right Justify
            Caption         =   "Logradouro:"
            Height          =   255
            Left            =   225
            TabIndex        =   30
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.TextBox Text_Rg 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         ToolTipText     =   "Rg"
         Top             =   2355
         Width           =   2655
      End
      Begin VB.TextBox Text_Razao 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         ToolTipText     =   "Razão Social"
         Top             =   3075
         Width           =   2655
      End
      Begin VB.TextBox Text_Celular 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         ToolTipText     =   "Celular"
         Top             =   1635
         Width           =   2655
      End
      Begin VB.TextBox Text_Fone_Resid 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         ToolTipText     =   "Fone Residencial"
         Top             =   1275
         Width           =   2655
      End
      Begin VB.TextBox Text_Cpf 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         ToolTipText     =   "Cpf"
         Top             =   2715
         Width           =   2655
      End
      Begin VB.TextBox Text_Nome 
         Height          =   285
         Left            =   4335
         TabIndex        =   2
         ToolTipText     =   "Nome"
         Top             =   390
         Width           =   4575
      End
      Begin VB.TextBox Text_Cnpj 
         Height          =   285
         Left            =   1440
         MaxLength       =   18
         TabIndex        =   9
         ToolTipText     =   "Cnpj"
         Top             =   3450
         Width           =   2700
      End
      Begin VB.TextBox Text_Cod 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   420
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Mask_Dt_Nasc 
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
         Left            =   1440
         TabIndex        =   5
         ToolTipText     =   "Data de Nascimento"
         Top             =   1995
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   -2147483634
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   8595
         Top             =   630
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Email:"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   870
         Width           =   1095
      End
      Begin VB.Label Lbl_Status 
         Alignment       =   1  'Right Justify
         Caption         =   "Status:"
         Height          =   255
         Left            =   480
         TabIndex        =   43
         Top             =   3795
         Width           =   855
      End
      Begin VB.Label Lbl_Rg 
         Alignment       =   1  'Right Justify
         Caption         =   "Rg:"
         Height          =   255
         Left            =   960
         TabIndex        =   28
         Top             =   2355
         Width           =   375
      End
      Begin VB.Label Lbl_Dt_Prox_Visita 
         Alignment       =   1  'Right Justify
         Caption         =   "Data de Nasc.:"
         Height          =   255
         Left            =   195
         TabIndex        =   27
         Top             =   1995
         Width           =   1140
      End
      Begin VB.Label Lbl_Razao 
         Alignment       =   1  'Right Justify
         Caption         =   "Razão Social:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   3075
         Width           =   1215
      End
      Begin VB.Label Lbl_Celular 
         Alignment       =   1  'Right Justify
         Caption         =   "Celular:"
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   1635
         Width           =   615
      End
      Begin VB.Label Lbl_Fone_Resid 
         Alignment       =   1  'Right Justify
         Caption         =   "Fone Resid:"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   1275
         Width           =   855
      End
      Begin VB.Label Lbl_Cpf 
         Alignment       =   1  'Right Justify
         Caption         =   "Cpf:"
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   2715
         Width           =   375
      End
      Begin VB.Label Lbl_Nome 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome:"
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   435
         Width           =   1095
      End
      Begin VB.Label Lbl_Cnpj 
         Alignment       =   1  'Right Justify
         Caption         =   "Cnpj:"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   3435
         Width           =   495
      End
      Begin VB.Label Lbl_Cod 
         Alignment       =   1  'Right Justify
         Caption         =   "Código:"
         Height          =   255
         Left            =   660
         TabIndex        =   18
         Top             =   480
         Width           =   675
      End
   End
End
Attribute VB_Name = "Form_Technician"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private W_DB As ADODB.Connection
Private mandatoryObjs(9) As Variant
Dim fso As New FileSystemObject

Private Sub But_AuxTecPhoto_Clean_Click()
    Text_AuxTec_Photo = ""
End Sub

Private Sub But_ChooseAuxTecPhoto_Click()
    CommonDialog2.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png;|All files (*.*)|*.*"
    CommonDialog2.DefaultExt = "jpg"
    CommonDialog2.DialogTitle = "Select File"
    CommonDialog2.ShowOpen

    Text_AuxTec_Photo.Text = CommonDialog2.FileName
    Pic_AuxTec.Picture = LoadPicture(Text_AuxTec_Photo.Text)
End Sub

Private Sub But_ChooseTecPhoto_Click()
    CommonDialog1.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png;|All files (*.*)|*.*"
    CommonDialog1.DefaultExt = "jpg"
    CommonDialog1.DialogTitle = "Select File"
    CommonDialog1.ShowOpen

    Text_Tec_Photo.Text = CommonDialog1.FileName
    Pic_Tec.Picture = LoadPicture(Text_Tec_Photo.Text)
    
End Sub

Private Sub But_TecPhoto_Clean_Click()
    Text_Tec_Photo = ""
End Sub

Private Sub Command1_Click()
    Comb_Status.Font.Bold = True
    
End Sub

Private Sub Form_Activate()
Text_AuxTec_Photo.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Tec_Photo.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Nome_Aux.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Rg_Aux.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Cpf_Aux.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Cep.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Cidade.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Bairro.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_End.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Rg.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Razao.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Celular.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Fone_Resid.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Cpf.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Nome.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Cnpj.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Text_Cod.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Comb_Status.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Grid_Technician.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
Mask_Dt_Nasc.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
End Sub

Sub Form_Load()
    
    If (MDIForm_Xsol.acessibilidade_TextoNegrito = True) Or (MDIForm_Xsol.acessibilidade_TextoNegrito = False) Then
        Comb_Status.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
    End If
    
    Call InitCombos
    Call SetFields
    Call SetMandatoryFields
    
    Set W_DB = New ADODB.Connection
    W_DB.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & W_DBname & ";"

    Call Configure_Grid
    Call Fill_Grid(Return_Recordset(Return_SqlString(1)))
    Call carregaCboVeiculos(Return_Recordset(Return_SqlString(5)))
    
End Sub

Sub carregaCboVeiculos(rs As ADODB.Recordset, Optional code As String)
    
    With cboVeiculos
        cboVeiculos.Clear
    
        If rs.EOF Or rs.BOF Then
           Exit Sub
        End If
        
        .AddItem (" ")
        
        Do Until rs.EOF
            .AddItem (rs.Fields("codigo") & " - " & rs.Fields("MARCAMODELO") & " - " & rs.Fields("ANO") & " - " + rs.Fields("PLACA"))
            rs.MoveNext
        Loop
    End With
    
    rs.Close
    Set rs = Nothing
    
End Sub

Private Sub Grid_Technician_Click()
    With Grid_Technician
        If .RowSel = 1 Then
            Exit Sub
        End If
        Call Show_Record(Return_Recordset(Return_SqlString(2, .TextMatrix(.RowSel, 0))))
    End With
    Call Configure_Controls
End Sub

'status = 0 - normal
'       = 1 - insert
'       = 2 - edit

Private Sub Configure_Controls(Optional status As Integer)
    Dim color As ColorConstants

    But_Insert.enabled = status <= 1
    But_Edit.enabled = Grid_Technician.RowSel > 0 And (status = 0 Or status = 2)
    But_Save.enabled = status > 0
    But_Delete.enabled = Grid_Technician.RowSel > 0 And status <= 0
    Grid_Technician.enabled = status <= 0
    
    But_Insert.Caption = IIf(status = 1, "&Cancelar", "&Incluir")
    But_Edit.Caption = IIf(status = 2, "&Cancelar", "&Editar")
    
    Text_Nome.enabled = status > 0
    Text_Fone_Resid.enabled = status > 0
    Text_Celular.enabled = status > 0
    Mask_Dt_Nasc.enabled = status > 0
    Text_Rg.enabled = status > 0
    Text_Cpf.enabled = status > 0
    Text_Razao.enabled = status > 0
    Text_Cnpj.enabled = status > 0
    Comb_Status.enabled = status > 0
    Text_End.enabled = status > 0
    Text_Bairro.enabled = status > 0
    Text_Cidade.enabled = status > 0
    Text_Cep.enabled = status > 0
    Text_Nome_Aux.enabled = status > 0
    Text_Rg_Aux.enabled = status > 0
    Text_Cpf_Aux.enabled = status > 0
    Text_Email.enabled = status > 0
    But_ChooseTecPhoto.enabled = status > 0
    But_TecPhoto_Clean.enabled = status > 0
    But_ChooseAuxTecPhoto.enabled = status > 0
    But_AuxTecPhoto_Clean.enabled = status > 0
    
    Call PaintMandatoryField(mandatoryObjs, status > 0)
    Call Clean_Controls(status)
    
End Sub

Private Sub But_Insert_Click()
    Dim status As Integer
    
    status = IIf(But_Insert.Caption = "&Incluir", 1, 0)
    Call Configure_Controls(status)
    If status = 0 Then
        Call Grid_Technician_Click
    End If
End Sub

Private Sub But_Edit_Click()
    Dim status As Integer
    
    status = IIf(But_Edit.Caption = "&Editar", 2, 0)
    Call Configure_Controls(status)
    If status = 0 Then
        Call Grid_Technician_Click
    End If
End Sub

Private Sub But_Delete_Click()
    If MsgBox("Confirma a Exclusão do Técnico " & Text_Nome & " ?", vbYesNo, "Excluir Registro") = vbYes Then
        If HaveReference() Then
            MsgBox "Não é possível sua remoção pois há um Atendimento vinculado a este técnico", vbInformation, "Informação"
            Exit Sub
        End If
        Call Return_Recordset(Return_SqlString(2, Text_Cod)).Delete
        Call Delete_Photo
        Call Configure_Controls
        Call Fill_Grid(Return_Recordset(Return_SqlString(1)))
    End If
End Sub

Private Function HaveReference() As Boolean
    Dim rs As ADODB.Recordset

    Set rs = Return_Recordset(Return_SqlString(4, Text_Cod.Text))
    
    If rs.RecordCount > 0 Then
        HaveReference = True
    Else
        HaveReference = False
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub But_Save_Click()
    Dim rs As ADODB.Recordset

    
    If Trim(Text_Cep.Text) = "" Then
            MsgBox ("O campo CEP é de preenchimento obrigatório!")
            'Text_Cep.SetFocus
            Exit Sub
    End If
    
    If (Len(Text_Cep.Text) <> 8) And (Len(Text_Cep.Text) <> 9) Then
        MsgBox ("O CEP deve conter exatamente 8 dígitos, verifique!")
        'Text_Cep.SetFocus
        Exit Sub
    End If
    
    
    If Not MandatoryFieldFilled(mandatoryObjs) Then
        Exit Sub
    End If
    
    If But_Insert.Caption = "&Cancelar" Then
        Set rs = Return_Recordset(Return_SqlString(3))
        rs.AddNew
        Text_Cod.Text = rs.Fields("codigo")
    Else
        Set rs = Return_Recordset(Return_SqlString(2, Text_Cod.Text))
    End If
    
    Call Record_Record(rs)

    Call Save_Photo
    
    rs.Update

    rs.Close
    Set rs = Nothing
    
    Call Fill_Grid(Return_Recordset(Return_SqlString(1)), Text_Cod)
    Call Configure_Controls
End Sub

Private Sub Record_Record(rs As ADODB.Recordset)
   rs.Fields("Nome").value = UCase(Text_Nome)
   If Trim(Text_Fone_Resid) = "" Then
      rs.Fields("Fone").value = Null
   Else
      rs.Fields("Fone").value = Text_Fone_Resid
   End If
   rs.Fields("Celular").value = Text_Celular
   rs.Fields("Dt_Nasc").value = Mask_Dt_Nasc
   If Trim(Text_Rg) = "" Then
      rs.Fields("Rg").value = Null
   Else
      rs.Fields("Rg").value = Text_Rg
   End If
   rs.Fields("Cpf").value = Text_Cpf
   If Trim(Text_Razao) = "" Then
      rs.Fields("Razao").value = Null
   Else
      rs.Fields("Razao").value = Text_Razao
   End If
   If Trim(Text_Cnpj) = "" Then
      rs.Fields("Cnpj").value = Null
   Else
      rs.Fields("Cnpj").value = Text_Cnpj
   End If
   rs.Fields("Status").value = UCase(Mid(Comb_Status, 1, 1))
   rs.Fields("End").value = UCase(Text_End)
   rs.Fields("Bairro").value = UCase(Text_Bairro)
   rs.Fields("Cidade").value = UCase(Text_Cidade)
   
   If Trim(Text_Cep.Text) = "" Then
      rs.Fields("Cep").value = Null
   Else
      rs.Fields("Cep").value = Text_Cep
   End If
   
   If Trim(Text_Nome_Aux) = "" Then
      rs.Fields("Aux_Tecnico").value = Null
   Else
      rs.Fields("Aux_Tecnico").value = Text_Nome_Aux
   End If
   
'   If Trim(Text_Tec_Photo.Text) <> "" Then
'    rs.Fields("Foto_Tecnico_Path").value = Text_Tec_Photo.Text
'   End If
'
'   If Trim(Text_AuxTec_Photo.Text) <> "" Then
'    rs.Fields("Foto_Tecnico_Path").value = Text_AuxTec_Photo.Text
'   End If
   
   If Trim(Text_Rg_Aux) = "" Then
      rs.Fields("Aux_Tec_Rg").value = Null
   Else
      rs.Fields("Aux_Tec_Rg").value = Text_Rg_Aux
   End If
   If Trim(Text_Cpf_Aux) = "" Then
      rs.Fields("Aux_Tec_Cpf").value = Null
   Else
      rs.Fields("Aux_Tec_Cpf").value = Text_Cpf_Aux
   End If
   
   If Trim(Text_Email) = "" Then
      rs.Fields("email").value = ""
   Else
      rs.Fields("email").value = Text_Email.Text
   End If
   
   If (cboVeiculos.ListIndex > -1) And (cboVeiculos.Text <> " ") Then
    rs.Fields("CODVEICULO").value = Mid(cboVeiculos.Text, 1, 2)
   End If
   
   If (cboVeiculos.ListIndex > -1) And (cboVeiculos.Text = " ") Then
    rs.Fields("CODVEICULO").value = "0"
   End If
   
End Sub

Private Sub Save_Photo()
    On Error GoTo ErrHandler
    
    If Not fso.FolderExists(Image_Folder) Then
        fso.CreateFolder (Image_Folder)
    End If
    If Text_Tec_Photo = "" Then
        If fso.FileExists(Image_Folder & "\foto1_" & Text_Cod) Then
            fso.DeleteFile Image_Folder & "\foto1_" & Text_Cod
        End If
    ElseIf Text_Tec_Photo <> "(foto)" Then
        fso.CopyFile Text_Tec_Photo, Image_Folder & "\foto1_" & Text_Cod ' & "." & GetFileExtension(Text_Tec_Photo)
    End If
    If Text_AuxTec_Photo = "" Then
        If fso.FileExists(Image_Folder & "\foto2_" & Text_Cod) Then
            fso.DeleteFile Image_Folder & "\foto2_" & Text_Cod
        End If
    ElseIf Text_AuxTec_Photo <> "(foto)" Then
        fso.CopyFile Text_AuxTec_Photo, Image_Folder & "\foto2_" & Text_Cod ' & "." & GetFileExtension(Text_Tec_Photo)
    End If
    Exit Sub
    
ErrHandler:
    'MsgBox "Não foi possível criar o diretório [" & Image_Folder & "] definido em CONFIG.INI", vbCritical, "Erro"
End Sub

Private Sub Delete_Photo()
    If fso.FileExists(Image_Folder & "\foto1_" & Text_Cod) Then
        fso.DeleteFile Image_Folder & "\foto1_" & Text_Cod
    End If
    If fso.FileExists(Image_Folder & "\foto2_" & Text_Cod) Then
        fso.DeleteFile Image_Folder & "\foto2_" & Text_Cod
    End If
End Sub

Private Sub Show_Photo()
    If fso.FileExists(Image_Folder & "\foto1_" & Text_Cod) Then
        Pic_Tec.Picture = LoadPicture(Image_Folder & "\foto1_" & Text_Cod)
        'Text_Tec_Photo = "(foto)"
    Else
        Pic_Tec.Picture = Nothing
    End If
    If fso.FileExists(Image_Folder & "\foto2_" & Text_Cod) Then
        Pic_AuxTec.Picture = LoadPicture(Image_Folder & "\foto2_" & Text_Cod)
        'Text_AuxTec_Photo = "(foto)"
    Else
        Pic_AuxTec.Picture = Nothing
    End If
End Sub

Private Function Return_SqlString(queryType As Integer, Optional param As String) As String
    Dim whereClause As String
    
    If queryType = 1 Then
        Return_SqlString = "SELECT * FROM Tecnico AS t"
    ElseIf queryType = 2 Then
        Return_SqlString = "SELECT * FROM Tecnico t WHERE codigo = " & param
    ElseIf queryType = 3 Then
        Return_SqlString = "SELECT * FROM Tecnico"
    ElseIf queryType = 4 Then
        Return_SqlString = "SELECT * FROM Atendimento WHERE cod_tecnico = " & param
    ElseIf queryType = 5 Then
        Return_SqlString = "SELECT * FROM Veiculos"
    End If

End Function

Private Function Return_Recordset(strSql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    rs.Open strSql, W_DB, adOpenStatic, adLockOptimistic
    
    Set Return_Recordset = rs
End Function

Private Sub Fill_Grid(rs As ADODB.Recordset, Optional code As String)
    Dim selectedRow As Integer
    Dim lastCode As Integer
    Dim count As Integer
    
    lastCode = -1
    selectedRow = 2
    count = 0
    With Grid_Technician
        .Rows = 2
        Do Until rs.EOF
            If lastCode <> rs.Fields("codigo") Then
               lastCode = rs.Fields("codigo")
               count = count + 1
               .AddItem (rs.Fields("codigo") & vbTab & IIf(IsNull(rs.Fields("nome")), "", rs.Fields("nome")) & vbTab & IIf(IsNull(rs.Fields("Celular")), "", rs.Fields("Celular")) & vbTab & IIf(IsNull(rs.Fields("cpf")), "", rs.Fields("cpf")) & vbTab & IIf(IsNull(rs.Fields("Status")), "", IIf(rs.Fields("Status") = "A", "Ativo", "Suspenso")))
               If code <> "" Then
                  If code = rs.Fields("codigo") Then
                     'selectedRow = count
                     selectedRow = .Rows - 1
                  End If
               End If
            End If
            rs.MoveNext
        Loop
        
        .Refresh
        If rs.RecordCount > 0 Then
            .row = selectedRow
            .RowSel = selectedRow
            .TopRow = selectedRow
            .col = 0
            .ColSel = .Cols - 1
            But_Edit.enabled = True
            But_Delete.enabled = True
            Call Grid_Technician_Click
        Else
            But_Edit.enabled = False
            But_Delete.enabled = False
        End If
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Sub Show_Record(rs As ADODB.Recordset)
On Error GoTo finaliza
    If rs.EOF Or rs.BOF Then
       Exit Sub
    End If

    Text_Cod.Text = UCase(IIf(IsNull(rs.Fields("Codigo")), "", rs.Fields("Codigo")))
    
    Text_Nome.Text = IIf(IsNull(rs.Fields("Nome")), "", rs.Fields("Nome"))
    
    Text_Fone_Resid.Text = IIf(IsNull(rs.Fields("Fone")), "", rs.Fields("Fone"))
        
    Text_Celular.Text = IIf(IsNull(rs.Fields("Celular")), "", rs.Fields("Celular"))
    
    Mask_Dt_Nasc.Text = IIf(IsNull(rs.Fields("Dt_Nasc")), "__/__/____", rs.Fields("Dt_Nasc"))
    
    Text_Rg.Text = IIf(IsNull(rs.Fields("RG")), "", rs.Fields("RG"))
    
    Text_Cpf.Text = IIf(IsNull(rs.Fields("CPF")), "", rs.Fields("CPF"))
    
    Text_Razao.Text = IIf(IsNull(rs.Fields("Razao")), "", rs.Fields("Razao"))
    
    Text_Cnpj.Text = IIf(IsNull(rs.Fields("CNPJ")), "", rs.Fields("CNPJ"))
    
    
    If StrComp(rs.Fields("Status"), "A", vbTextCompare) = 0 Then
        Comb_Status.ListIndex = 0
    ElseIf StrComp(rs.Fields("Status"), "S", vbTextCompare) = 0 Then
        Comb_Status.ListIndex = 1
    End If
    
    Text_End.Text = UCase(IIf(IsNull(rs.Fields("End")), "", rs.Fields("End")))

    Text_Bairro.Text = UCase(IIf(IsNull(rs.Fields("Bairro")), "", rs.Fields("Bairro")))

    Text_Cidade.Text = UCase(IIf(IsNull(rs.Fields("Cidade")), "", rs.Fields("Cidade")))

    'MsgBox Mid(rs.Fields("Cep"), 1, 5) & "-" & Mid(rs.Fields("Cep"), 6, 3)
    
    
    Text_Cep.Text = IIf(IsNull(rs.Fields("Cep")), "", rs.Fields("Cep"))
    
    Text_Nome_Aux.Text = IIf(IsNull(rs.Fields("Aux_Tecnico")), "", rs.Fields("Aux_Tecnico"))
    
    Text_Rg_Aux.Text = IIf(IsNull(rs.Fields("Aux_Tec_RG")), "", rs.Fields("Aux_Tec_RG"))
    
    Text_Cpf_Aux.Text = IIf(IsNull(rs.Fields("Aux_Tec_CPF")), "", rs.Fields("Aux_Tec_CPF"))
    
    'Text_Tec_Photo = ""
    
    'Text_AuxTec_Photo = ""
    
    If IsNull(rs.Fields("email")) Then
        Text_Email.Text = ""
    Else
        Text_Email.Text = rs.Fields("email")
    End If
    
    'veiculo
    If Form_Technician.Visible Then
        If IsNull(rs.Fields("CODVEICULO")) Then
            cboVeiculos.ListIndex = -1
        Else
            cboVeiculos.ListIndex = retornaIndiceCboVeiculo(rs.Fields("CODVEICULO"))
        End If
    End If
    
finaliza:
    Show_Photo
    
    rs.Close
    Set rs = Nothing
End Sub

Function retornaIndiceCboVeiculo(nCodigoVeiculo As Integer) As Integer

If cboVeiculos.ListCount = 0 Then
    Exit Function
End If

Dim xindice As Integer

For xindice = 0 To 20
    cboVeiculos.ListIndex = xindice
    If Mid(cboVeiculos.Text, 1, 2) = nCodigoVeiculo Then
        'cboVeiculos.ListIndex = xindice
        Exit For
    Else
        If nCodigoVeiculo = 0 Then
            cboVeiculos.ListIndex = -1
            Exit For
        Else
            cboVeiculos.ListIndex = cboVeiculos.ListIndex + 1
        End If
    End If
Next xindice

retornaIndiceCboVeiculo = xindice

End Function


Sub Clean_Controls(Optional status As Integer)
    If status = 1 Then
        Text_Cod.Text = ""
        Text_Nome.Text = ""
        Text_Fone_Resid.Text = ""
        Text_Celular.Text = ""
        Mask_Dt_Nasc.Text = "__/__/____"
        Text_Rg.Text = ""
        Text_Cpf.Text = ""
        Text_Razao.Text = ""
        Text_Cnpj.Text = ""
        Comb_Status.ListIndex = -1
        Text_End.Text = ""
        Text_Bairro.Text = ""
        Text_Cidade.Text = ""
        Text_Cep.Text = ""
        Text_Nome_Aux.Text = ""
        Text_Rg_Aux.Text = ""
        Text_Cpf_Aux.Text = ""
        Text_Tec_Photo = ""
        Text_AuxTec_Photo = ""
    End If
End Sub
Private Sub InitCombos()
    Comb_Status.AddItem "Ativo"
    Comb_Status.AddItem "Suspenso"
End Sub

Private Sub SetFields()
   Text_Nome.MaxLength = 80
   Text_Fone_Resid.MaxLength = 25
   Text_Celular.MaxLength = 25
   Text_Rg.MaxLength = 15
   Text_Cpf.MaxLength = 15
   Text_Razao.MaxLength = 80
   'Text_Cnpj.MaxLength = 15
   Text_End.MaxLength = 130
   Text_Bairro.MaxLength = 80
   Text_Cidade.MaxLength = 80
   Text_Cep.MaxLength = 9
   Text_Nome_Aux.MaxLength = 80
   Text_Rg_Aux.MaxLength = 15
   Text_Cpf_Aux.MaxLength = 15
End Sub

Private Sub SetMandatoryFields()
   Set mandatoryObjs(1) = Text_Nome
   Set mandatoryObjs(2) = Text_Celular
   Set mandatoryObjs(3) = Mask_Dt_Nasc
   Set mandatoryObjs(4) = Text_Cpf
   Set mandatoryObjs(5) = Comb_Status
   Set mandatoryObjs(6) = Text_End
   Set mandatoryObjs(7) = Text_Bairro
   Set mandatoryObjs(8) = Text_Cidade
   Set mandatoryObjs(9) = Text_Cep
End Sub

Private Sub Configure_Grid()
    With Grid_Technician
        .Cols = 5
        .FixedCols = 0
        .ColWidth(0) = 650
        .ColWidth(1) = 3490
        .ColWidth(2) = 1400
        .ColWidth(3) = 1500
        .ColWidth(4) = 900
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Nome"
        .TextMatrix(0, 2) = "Celular"
        .TextMatrix(0, 3) = "CPF"
        .TextMatrix(0, 4) = "Status"
        .RowHeight(1) = 0
    End With
End Sub

Private Sub Grid_Technician_LostFocus()
'Valida o tamanho do CNPJ

End Sub

Private Sub Grid_Technician_SelChange()
   Call Grid_Technician_Click
End Sub



Private Sub Text_Cep_KeyPress(KeyAscii As Integer)
    'MsgBox KeyAscii
    

End Sub

Private Sub Text_Cep_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim strCEP As String
    
    If (Text_Cep.Text <> "") And (Not IsNumeric(Text_Cep.Text)) And (KeyAscii <> 8) Then
        MsgBox "Digite somente NÚMEROS para o CEP.", vbOKOnly, "xSol Aquecedores"
        Text_Cep.Text = ""
        'Text_Cep.SetFocus
        Exit Sub
    End If
    
    If Len(Text_Cep.Text) = 8 Then
        strCEP = Mid(Text_Cep.Text, 1, 5) & "-" & Mid(Text_Cep.Text, 6, 3)
        Text_Cep.Text = strCEP
    End If
    
    
End Sub

Private Sub Text_Cnpj_KeyUp(KeyCode As Integer, Shift As Integer)
    
    
    If (KeyCode <> 8) And (KeyCode <> 13) And (KeyCode <> 48) And (KeyCode <> 96) And (KeyCode <> 97) And (KeyCode <> 98) And (KeyCode <> 99) And (KeyCode <> 100) And (KeyCode <> 101) And (KeyCode <> 102) And (KeyCode <> 103) And (KeyCode <> 104) And (KeyCode <> 105) Then
        MsgBox "Informe somente números para o CNPJ"
        Text_Cnpj.Text = ""
        
    End If
    

End Sub

Private Sub Text_Cnpj_LostFocus()
    Dim strCnpj As String

    If Trim(Text_Cnpj.Text) <> "" Then
        
        '46 927 393 0001 84
        '46.927.393/0001-84

        If (Len(Text_Cnpj.Text) = 14) Then
            strCnpj = Mid(Text_Cnpj.Text, 1, 2) & "." & Mid(Text_Cnpj.Text, 3, 3) & "." & Mid(Text_Cnpj.Text, 6, 3) & "/" & Mid(Text_Cnpj.Text, 9, 4) & "-" & Mid(Text_Cnpj, 13, 2)
            Text_Cnpj.Text = strCnpj
        End If

    End If
    

    
    
End Sub


Private Function ValidateNumeric(strText As String) _
    As Boolean
    ValidateNumeric = CBool(strText = "" _
    Or strText = "-" _
    Or strText = "-." _
    Or strText = "." _
    Or IsNumeric(strText))
End Function

