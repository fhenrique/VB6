VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7980
   ClientLeft      =   9840
   ClientTop       =   2295
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   6585
   Begin VB.CommandButton Command3 
      Caption         =   "Inserir registro"
      Height          =   810
      Left            =   570
      TabIndex        =   2
      Top             =   1920
      Width           =   1410
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pegar registro"
      Height          =   810
      Left            =   480
      TabIndex        =   1
      Top             =   915
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conetar"
      Height          =   615
      Left            =   525
      TabIndex        =   0
      Top             =   165
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public conexao As New ADODB.Connection
Public registro As New ADODB.Recordset

Private Sub Command1_Click()
    
    'FAZENDO A CONEXÃO COM MYSQL remoto
    conexao.Open "Driver={MySQL ODBC 3.51 Driver};" & _
    "Server=mysql.xsolaquecedores.com.br;" & _
    "Port=3306;" & _
    "Database=xsolaquecedore01;" & _
    "User=xsolaquecedore01;" & _
    "Password=xsol303080;" & _
    "Option=3;"
    
    MsgBox "conectado"
    

End Sub

Private Sub Command2_Click()
    
    If conexao.State = 0 Then 'conexao fechada
        MsgBox "Conexão com o banco está fechada"
        Exit Sub
    End If
    
    'TRAZENDO UM REGISTRO
    Set registro = conexao.Execute("Select * from clientes")
    
    Do While Not registro.EOF
        MsgBox registro("nome")
        registro.MoveNext
    Loop


End Sub

Private Sub Command3_Click()
    If conexao.State = 0 Then 'conexao fechada
        MsgBox "Conexão com o banco está fechada"
        Exit Sub
    End If
    
    'TRAZENDO UM REGISTRO
    Set registro = conexao.Execute("Insert into `Zona`(Zona) Value('Teste')")
    

End Sub
