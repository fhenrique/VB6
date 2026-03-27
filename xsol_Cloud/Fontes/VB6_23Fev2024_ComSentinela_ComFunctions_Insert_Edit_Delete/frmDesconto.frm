VERSION 5.00
Begin VB.Form frmDesconto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desconto dor orçamento"
   ClientHeight    =   1545
   ClientLeft      =   6990
   ClientTop       =   4710
   ClientWidth     =   4305
   Icon            =   "frmDesconto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   45
      TabIndex        =   0
      Top             =   -30
      Width           =   4185
      Begin VB.OptionButton optnPorPorcentagem 
         Caption         =   "Por porcentagem %"
         Height          =   225
         Left            =   2235
         TabIndex        =   4
         Top             =   600
         Width           =   1890
      End
      Begin VB.TextBox txtValorDesconto 
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Top             =   1080
         Width           =   915
      End
      Begin VB.OptionButton optnPorValor 
         Caption         =   "Por valor R$"
         Height          =   210
         Left            =   195
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton btnAplicarDesconto 
         Caption         =   "Aplicar"
         Height          =   345
         Left            =   2325
         TabIndex        =   1
         Top             =   1050
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Conceder desconto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   975
         TabIndex        =   5
         Top             =   165
         Width           =   2040
      End
   End
End
Attribute VB_Name = "frmDesconto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public valorQueveio As Double
Public tipoDesconto As String

Private Sub btnAplicarDesconto_Click()
    If optnPorValor = True Then
        frmOrcamento.descontoPorValor = True
        frmOrcamento.descontoPorPorcentagem = False
    End If
    
    If optnPorPorcentagem = True Then
        frmOrcamento.descontoPorValor = False
        frmOrcamento.descontoPorPorcentagem = True
    End If
    
    
    frmOrcamento.descontoValorConcedido = txtValorDesconto.Text
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    If valorQueveio > 0 Then
        txtValorDesconto.Text = valorQueveio
    Else
        txtValorDesconto.Text = ""
    End If
    
    If tipoDesconto <> "" Then
        If tipoDesconto = "valor" Then
            optnPorValor = True
        Else
            optnPorPorcentagem = True
        End If
    Else
        optnPorValor = True
    End If
    
End Sub

