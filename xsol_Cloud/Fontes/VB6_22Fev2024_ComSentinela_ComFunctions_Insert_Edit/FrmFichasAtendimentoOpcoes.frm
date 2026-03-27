VERSION 5.00
Begin VB.Form FrmFichasAtendimentoOpcoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opções de impressão"
   ClientHeight    =   1275
   ClientLeft      =   6915
   ClientTop       =   3735
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1200
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   3285
      Begin VB.CheckBox chkNegrito 
         Caption         =   "TEXTO EM NEGRITO"
         Height          =   270
         Left            =   105
         TabIndex        =   4
         Top             =   795
         Width           =   2055
      End
      Begin VB.ComboBox cboCorTexto 
         Height          =   315
         ItemData        =   "FrmFichasAtendimentoOpcoes.frx":0000
         Left            =   1395
         List            =   "FrmFichasAtendimentoOpcoes.frx":0010
         TabIndex        =   2
         Top             =   180
         Width           =   1785
      End
      Begin VB.CommandButton btnAplicar 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   675
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "COR DO TEXTO"
         Height          =   195
         Left            =   105
         TabIndex        =   3
         Top             =   300
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmFichasAtendimentoOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub FichaSetaCorTexto()
   Dim ConfigFile As String

   ConfigFile = App.Path & "\CONFIG.INI"

   PutINISetting "FICHA_DESTAQUE", "FORECOLOR", cboCorTexto.Text, ConfigFile
   
   If chkNegrito.value = 1 Then
    PutINISetting "FICHA_DESTAQUE", "NEGRITO", "True", ConfigFile
   Else
    PutINISetting "FICHA_DESTAQUE", "NEGRITO", "False", ConfigFile
   End If
   
   'Unload Me

End Sub


Private Sub btnAplicar_Click()
    Dim strNegrito As String
    
    
    If chkNegrito.value = 1 Then
        strNegrito = "ligada"
    Else
        strNegrito = "desligada"
    End If
        
    If (cboCorTexto.Text = "") Then
        MsgBox "Selecione as cores do texto e fundo!"
        Exit Sub
    End If
    
    Call FichaSetaCorTexto
    Call MDIForm_Xsol.Arquivo_Dados
    MsgBox "A cor do texto em destaque será " & cboCorTexto.Text & Chr(10) & _
    "A opção de negrito está " & strNegrito & "!"
    
    Unload Me

End Sub

Private Sub Form_Activate()
    Call MDIForm_Xsol.Arquivo_Dados
    
    cboCorTexto.Text = MDIForm_Xsol.fichaDestaqueForecolor
    
    If MDIForm_Xsol.fichaDestaqueNegrigo = True Then
        chkNegrito.value = 1
    Else
        chkNegrito.value = 0
    End If
    
    cboCorTexto.Font.Bold = MDIForm_Xsol.acessibilidade_TextoNegrito
End Sub

