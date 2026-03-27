VERSION 5.00
Begin VB.Form Form_Logo 
   Caption         =   "XSol Aquecedores"
   ClientHeight    =   2655
   ClientLeft      =   2085
   ClientTop       =   3345
   ClientWidth     =   5130
   Icon            =   "Form_Logo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Logo.frx":08CA
   ScaleHeight     =   2655
   ScaleWidth      =   5130
End
Attribute VB_Name = "Form_Logo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'Parametros de inicialização

Dim hMenu As Long
Dim lItemCount As Long

hMenu = GetSystemMenu(Me.hWnd, 0)
If hMenu Then
   lItemCount = GetMenuItemCount(hMenu)
   Call RemoveMenu(hMenu, lItemCount - 1, MF_REMOVE Or _
   MF_BYPOSITION)
   Call RemoveMenu(hMenu, lItemCount - 2, MF_REMOVE Or _
   MF_BYPOSITION)
   Call DrawMenuBar(Me.hWnd)
End If

Form_Logo.Top = (Screen.Height - Form_Logo.Height) / 2
Form_Logo.Left = (Screen.Width - Form_Logo.Width) / 2

Form_Logo.Show
Form_Logo.Refresh
Form_Logo.Caption = "Loading XSol..."

W_Form_Time = Timer
While Timer - W_Form_Time < 3
   DoEvents
Wend

MDIForm_Xsol.Show
Form_Login.Show

Unload Me

End Sub
