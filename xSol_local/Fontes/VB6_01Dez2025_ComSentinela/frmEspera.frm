VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmEspera 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   ClientHeight    =   600
   ClientLeft      =   7650
   ClientTop       =   3735
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   150
      Top             =   540
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmEspera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Timer1.enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.enabled = False
End Sub

Private Sub Timer1_Timer()
 If ProgressBar1.value < 100 Then
    ProgressBar1.value = ProgressBar1.value + 20
 Else
    ProgressBar1.value = 0
 End If
 
End Sub
