VERSION 5.00
Begin VB.Form frmGerenciadorDeDados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerenciador de dados"
   ClientHeight    =   2340
   ClientLeft      =   5580
   ClientTop       =   3180
   ClientWidth     =   6840
   Icon            =   "frmGerenciadorDeDados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2265
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   6705
      Begin VB.CommandButton btnExecutarSql 
         Caption         =   "Executar"
         Height          =   375
         Left            =   5415
         TabIndex        =   2
         Top             =   1785
         Width           =   1140
      End
      Begin VB.TextBox txtHotFix 
         Height          =   1485
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   195
         Width           =   6480
      End
   End
End
Attribute VB_Name = "frmGerenciadorDeDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnExecutarSql_Click()
    Dim rs As ADODB.Recordset
    Dim strComando As String
    
    'comandoDescriptografado = retornaComandoNatural(txtHotFix.Text)
    strComando = txtHotFix.Text
    
    On Error GoTo erro
    
    Set rs = Return_Recordset(strComando)
    Set rs = Nothing
    
    MsgBox "Comando executado com sucesso!", vbOKOnly, "XSol Aquecedores"
    txtHotFix.Text = ""
    Exit Sub
erro:
    MsgBox "Não foi possível executar o comando!", vbOKOnly, "XSol Aquecedores"
    txtHotFix.Text = ""
End Sub

Private Function retornaComandoNatural(comandoCriptografado As String) As String
    Dim stringDeRetorno As String
    
    'Vogal a
    stringDeRetorno = Replace(comandoCriptografado, "1", "A")
    
    'Vogal e
    stringDeRetorno = Replace(stringDeRetorno, "2", "E")
    
    'Vogal i
    stringDeRetorno = Replace(stringDeRetorno, "3", "I")
    
    'Vogal o
    stringDeRetorno = Replace(stringDeRetorno, "4", "O")
    
    'Vogal u
    stringDeRetorno = Replace(stringDeRetorno, "5", "U")

    
    retornaComandoNatural = stringDeRetorno
    
    
End Function
