VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAtualizador 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atualizador XSol Aquecedores"
   ClientHeight    =   3135
   ClientLeft      =   6795
   ClientTop       =   3825
   ClientWidth     =   4455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAtualizador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4455
   Begin VB.Timer timerProgress 
      Interval        =   500
      Left            =   1305
      Top             =   1185
   End
   Begin MSComctlLib.ProgressBar pbAtualizacao 
      Height          =   480
      Left            =   135
      TabIndex        =   0
      Top             =   1770
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   847
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Verificando"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1500
      TabIndex        =   2
      Top             =   2325
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Atualização do Sistema..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   735
      TabIndex        =   1
      Top             =   2670
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   1515
      Left            =   1515
      Picture         =   "frmAtualizador.frx":08CA
      Top             =   60
      Width           =   1425
   End
End
Attribute VB_Name = "frmAtualizador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Record_Record(rs As ADODB.Recordset)
    If rs.RecordCount = 0 Then
        rs.AddNew
    End If
    rs.Fields("NumeroAtualizacao").value = 3
End Sub

Private Function Return_SqlString(queryType As Integer, Optional param1 As String, Optional param2 As String, Optional param3 As String) As String
'    If queryType = 1 Then 'Cria o campo MesNaoDisponivel na tabela Tecnico
'        Return_SqlString = "ALTER TABLE Tecnico ADD COLUMN  MesNaoDisponivel VARCHAR(3)"
'    ElseIf queryType = 2 Then 'Cria o campo que guarda o número de atualizacoes na tabela Configuracao
'        Return_SqlString = "ALTER TABLE Configuracao ADD COLUMN  NumeroAtualizacao INTEGER"
'    ElseIf queryType = 3 Then 'Cria o campo que guarda o número de atualizacoes na tabela Configuracao
'        Return_SqlString = "UPDATE Configuracao SET NumeroAtualizacao = " & param1
'    ElseIf queryType = 4 Then 'Cria o campo que guarda o número de atualizacoes na tabela Configuracao
'        Return_SqlString = "Update Estoque SET  Produto = Replace(Produto, """", ""'"") "
'    ElseIf queryType = 5 Then 'Cria o campo codMaterial na tabela Estoque
'        Return_SqlString = "ALTER TABLE Estoque ADD COLUMN  codMaterial INTEGER"
'
'    ElseIf queryType = 6 Then 'Cria a tabela ORCAMENTO
'        Return_SqlString = "CREATE TABLE ORCAMENTO(CODIGO INTEGER,DATA DATETIME,COD_CLIENTE INTEGER,NECESSIDADE MEMO,ENTREGA TEXT(200),TEMPO_SERVICO TEXT(100),VALIDADE TEXT(50),VALOR_GERAL DECIMAL(15,2))"
'
'    ElseIf queryType = 7 Then 'Cria a tabela ORCAMENTO_GARANTIAS
'        Return_SqlString = "CREATE TABLE ORCAMENTO_GARANTIAS (CODIGO AUTOINCREMENT PRIMARY KEY, COD_ORCAMENTO INTEGER, GARANTIA_DESCRICAO TEXT(200))"
'
'    ElseIf queryType = 8 Then 'Cria a tabela ORCAMENTO_SOLUCAO
'        Return_SqlString = "CREATE TABLE ORCAMENTO_SOLUCAO(CODIGO AUTOINCREMENT PRIMARY KEY,COD_ORCAMENTO INTEGER,DESCRICAO TEXT(200),QUANTIDADE INTEGER,VALOR_UNITARIO DECIMAL(15,2),VALOR_TOTAL DECIMAL(15,2))"
'
'    ElseIf queryType = 9 Then 'Cria a tabela ORCAMENTO_FORMAS_PAGTO
'        Return_SqlString = "CREATE TABLE ORCAMENTO_FORMAS_PAGTO(CODIGO AUTOINCREMENT PRIMARY KEY,COD_ORCAMENTO INTEGER,FORMA_PAGTO_DESCRICAO TEXT(200))"
'
'    ElseIf queryType = 10 Then 'Cria a tabela ORCAMENTO_OBS
'        Return_SqlString = "CREATE TABLE ORCAMENTO_OBS(CODIGO AUTOINCREMENT PRIMARY KEY,COD_ORCAMENTO INTEGER,OBSERVACAO_DESCRICAO TEXT(200))"
'
'    ElseIf queryType = 11 Then 'Atualiza a tabela Endereco, incluindo a coluna COD_ORCAMENTO
'        Return_SqlString = "ALTER TABLE Endereco ADD COLUMN  COD_ORCAMENTO INTEGER"
'
'    ElseIf queryType = 12 Then 'Atualiza a tabela Clientes, incluindo as colunas FoneCelular_orcamento, FoneResidencial_orcamento, FoneComercial_orcamento (S ou N)
'        Return_SqlString = "ALTER TABLE Clientes ADD COLUMN  FoneCelular_orcamento VARCHAR(1), FoneResidencial_orcamento VARCHAR(1), FoneComercial_orcamento VARCHAR(1)"
'
'    'query 1 da atualização 7
'    ElseIf queryType = 13 And param1 = 7 Then 'Altera a tabela Clientes, criando o campo Tipo, varchar(20)
'        Return_SqlString = "ALTER TABLE Clientes ADD COLUMN  Tipo VARCHAR(20)"
'
'    ElseIf queryType = 14 And param1 = 7 Then 'Altera a tabela ORCAMENTO, criando o campo IMG_PATH, varchar(200)
'        Return_SqlString = "ALTER TABLE ORCAMENTO ADD COLUMN  IMG_PATH TEXT(250)"
'
'    ElseIf queryType = 15 And param1 = 7 Then 'Atualiza a tabela clientes, setando todos os campo TIPO = "CLIENTE"
'        Return_SqlString = "UPDATE Clientes SET Tipo = 'CLIENTE'"
    
'    'ATUALIZAÇÃO 8
'    ElseIf queryType = 16 And param1 = 8 Then 'Altera a tabela ORCAMENTO, criando o campo IMG_TEXT
'        Return_SqlString = "ALTER TABLE ORCAMENTO ADD COLUMN  IMG_TEXT TEXT(250)"
'    ElseIf queryType = 17 And param1 = 8 Then 'Altera a tabela ORCAMENTO, criando o campo IMG2_PATH
'        Return_SqlString = "ALTER TABLE ORCAMENTO ADD COLUMN  IMG2_PATH TEXT(250)"
'    ElseIf queryType = 18 And param1 = 8 Then 'Altera a tabela ORCAMENTO, criando o campo IMG2_TEXT
'        Return_SqlString = "ALTER TABLE ORCAMENTO ADD COLUMN  IMG2_TEXT TEXT(250)"
'    ElseIf queryType = 19 And param1 = 8 Then 'Altera a tabela de OS, incluindo os campos de equipamentos atendidos
'        'Return_SqlString = "CREATE TABLE EquipamentosDaOS (CODIGO AUTOINCREMENT PRIMARY KEY, COD_OS INTEGER, COD_CLIENTE INTEGER, COD_EQUIPAMENTO INTEGER)"
'        Return_SqlString = "ALTER TABLE Atendimento ADD COLUMN  Equip1_Codigo INTEGER, Equip2_Codigo INTEGER, Equip3_Codigo INTEGER, Equip4_Codigo INTEGER, Equip5_Codigo INTEGER, Equip6_Codigo INTEGER, Equip7_Codigo INTEGER, Equip8_Codigo INTEGER, Equip9_Codigo INTEGER, Equip10_Codigo INTEGER"
'    ElseIf queryType = 20 And param1 = 9 Then 'Altera a tabela de OS, incluindo os campos de equipamentos atendidos
'        Return_SqlString = "ALTER TABLE Endereco ADD COLUMN  CONTATO TEXT(250)"
'    ElseIf queryType = 21 And param1 = 10 Then 'Altera a tabela de OS, incluindo os campos de equipamentos atendidos
'        Return_SqlString = "ALTER TABLE Atendimento ADD COLUMN  COD_OS INTEGER"
        
    'ATUALIZAÇÃO 9
    'If (queryType = 1 And param1 = 9) Then 'Cria a tabela Atendimento_Pecas
    '    Return_SqlString = "CREATE TABLE Atendimento_Pecas(CODIGO AUTOINCREMENT PRIMARY KEY,CODIGO_ATENDIMENTO INTEGER, CODIGO_MATERIAL INTEGER, QTD INTEGER, VALOR DECIMAL(15,2))"
    'End If
    
    'ATUALIZAÇÃO 10
    'If (queryType = 1 And param1 = 10) Then 'Cria o campo Fone_Padrao na tabela Clientes
    '    Return_SqlString = "ALTER TABLE Orcamento ADD COLUMN Fone_Padrao TEXT(11)"
    'End If
    
    'ATUALIZAÇÃO 11
    'If (queryType = 1 And param1 = 11) Then 'Cria o campo Tipo_Telhado na tabela Endereco
    '    Return_SqlString = "ALTER TABLE Endereco ADD COLUMN Tipo_Telhado TEXT(50)"
    'End If
    
    'ATUALIZAÇÃO 12
    'If (queryType = 1 And param1 = 12) Then 'Altera a tabela ORCAMENTO, incluindo o campo desconto DECIMAL(15,2)
    '    Return_SqlString = "ALTER TABLE ORCAMENTO ADD COLUMN desconto DECIMAL(15,2)"
    'End If
    
    'ATUALIZAÇÃO 13
    'If (queryType = 1 And param1 = 13) Then 'Altera a tabela TECNICO, incluindo o campo email TEXT(30)
    '    Return_SqlString = "ALTER TABLE Tecnico ADD COLUMN email TEXT(30)"
    'End If
    
    'ATUALIZAÇÃO 14
    'If (queryType = 1 And param1 = 14) Then 'Cria a tabela de veiculos
    '    Return_SqlString = "CREATE TABLE VEICULOS(CODIGO AUTOINCREMENT PRIMARY KEY,MARCAMODELO TEXT(200),ANO TEXT(4), PLACA TEXT(10))"
    'End If
    
    'ATUALIZAÇÃO 15
    'If (queryType = 1 And param1 = 15) Then 'Cria campo codveiculo na tabela tecnicos
    '    Return_SqlString = "ALTER TABLE tecnico ADD COLUMN  CODVEICULO INTEGER"
    'End If
    
    'ATUALIZAÇÃO 16
    'If (queryType = 1 And param1 = 16) Then 'Cria campos STATUS e STATUS_MSG na tabela orçamento
    '    Return_SqlString = "ALTER TABLE orcamento ADD COLUMN  STATUS TEXT(10), STATUS_MSG TEXT(50)"
    'End If
    
    'ATUALIZAÇÃO 17
    'If (queryType = 1 And param1 = 17) Then 'Cria tabela tecnicos_agenda
    '    Return_SqlString = "CREATE TABLE tecnicos_agenda(ID AUTOINCREMENT PRIMARY KEY, codigo TEXT(5), DIAS TEXT(100), MES TEXT(3))"
    'End If
    
    'ATUALIZAÇÃO 18
    'If (queryType = 1 And param1 = 18) Then 'Alera a tabela orcamentos, inserindo o tipodesconto text(10)
    '    Return_SqlString = "ALTER TABLE orcamento ADD COLUMN  descontotipo TEXT(10)"
    'End If
    
    'ATUALIZAÇÃO 19
    'If (queryType = 1 And param1 = 19) Then 'altera a tabela enderecos, inserindo 3 novos campos para telefone
    '    Return_SqlString = "ALTER TABLE endereco ADD COLUMN  foneres TEXT(30), fonecom TEXT(30), fonecel TEXT(30)"
    'End If
    
    'ATUALIZAÇÃO 20
    'If (queryType = 1 And param1 = 20) Then 'Roda UPDATE entre as tabelas endereco e clientes para atualizar os campos de telefone na tabela endereco
    '    Return_SqlString = "UPDATE endereco a INNER JOIN clientes b ON a.cod_cliente = b.codigo SET a.foneres = b.fone_res, a.fonecom = b.fone_com, a.fonecel = b.fone_celular"
    'End If
    
    'ATUALIZAÇÃO 21: 04/12/2023
    If (queryType = 1 And param1 = 21) Then 'Altera a tabela Servico, inserindo o campo autoagendamento
        Return_SqlString = "ALTER TABLE servico ADD COLUMN autoagendamento TEXT(1)"
    End If
    
    'ATUALIZAÇÃO 22: 04/12/2023
    If (queryType = 1 And param1 = 22) Then 'Altera a tabela Servico, inserindo o campo autoagendamento
        Return_SqlString = "UPDATE servico set autoagendamento = 'N'"
    End If
    
    'ATUALIZAÇÃO 23: 04/12/2023
    If (queryType = 1 And param1 = 23) Then 'Altera a tabela Servico, inserindo o campo autoagendamento
        Return_SqlString = "UPDATE servico set autoagendamento = 'S' Where codigo in (9,10,13,51)"
    End If

    
    
End Function

Private Function atualizar(queryType As Integer, numeroAtualizacao As Integer) As Boolean
    Dim rs As ADODB.Recordset
    
    On Error GoTo erro
    Set rs = Return_Recordset(Return_SqlString(queryType, Str(numeroAtualizacao)))
    Set rs = Nothing
    
    atualizar = True
    Exit Function
    
erro:
    atualizar = False
End Function


Private Sub timerProgress_Timer()
    
    'Gerencia a atualização do banco de dados
    If pbAtualizacao.value < 100 Then
        pbAtualizacao.value = pbAtualizacao.value + 10
    Else
        'If (atualizar(6, 6) Or atualizar(7, 6) Or atualizar(8, 6) Or atualizar(9, 6) Or atualizar(10, 6) Or atualizar(11, 6) Or atualizar(12, 6) Or atualizar(13, 7) Or atualizar(14, 7) Or atualizar(15, 7)) Then
        'If (atualizar(16, 8) Or atualizar(17, 8) Or atualizar(18, 8) Or atualizar(19, 8)) Then
        'If (atualizar(1, 10)) Then
        
'        If (atualizar(1, 12)) Then
'            MsgBox "O banco de dados foi atualizado com sucesso!", vbOKOnly, "Xsol Aquecedores"
'            MDIForm_Xsol.sistemaAtualizado = True
'        End If
'
'        If (atualizar(1, 13)) Then
'            MsgBox "O banco de dados foi atualizado com sucesso!", vbOKOnly, "Xsol Aquecedores"
'            MDIForm_Xsol.sistemaAtualizado = True
'        End If
'
'        If (atualizar(1, 14)) Then
'            MsgBox "O banco de dados foi atualizado com sucesso!", vbOKOnly, "Xsol Aquecedores"
'            MDIForm_Xsol.sistemaAtualizado = True
'        Else
'            MsgBox "O banco de dados não precisa ser atualizado!", vbOKOnly, "Xsol Aquecedores"
'            Unload Me
'        End If
    
'        If (atualizar(1, 17)) Then
'            MsgBox "O banco de dados foi atualizado com sucesso!", vbOKOnly, "Xsol Aquecedores"
'            MDIForm_Xsol.sistemaAtualizado = True
'        End If
    
    
        If (atualizar(1, 21) And atualizar(1, 22) And atualizar(1, 23)) Then
            MsgBox "O banco de dados foi atualizado com sucesso!", vbOKOnly, "Xsol Aquecedores"
            MDIForm_Xsol.sistemaAtualizado = True
        Else
            MsgBox "O banco de dados não precisa ser atualizado!", vbOKOnly, "Xsol Aquecedores"
            Unload Me
        End If
    
    
        Unload Me
    End If
    
End Sub
