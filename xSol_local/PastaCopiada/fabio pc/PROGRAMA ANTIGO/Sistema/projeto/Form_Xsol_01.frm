VERSION 5.00
Begin VB.Form Form_Xsol_01 
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   2085
   ClientTop       =   3345
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Xsol_01.frx":0000
   ScaleHeight     =   2655
   ScaleWidth      =   5130
End
Attribute VB_Name = "Form_Xsol_01"
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

Form_Xsol_01.Top = (Screen.Height - Form_Xsol_01.Height) / 2
Form_Xsol_01.Left = (Screen.Width - Form_Xsol_01.Width) / 2

Form_Xsol_01.Show
Form_Xsol_01.Refresh
Form_Xsol_01.Caption = "Loading XSol..."

W_Form_Time = Timer
While Timer - W_Form_Time < 3
   DoEvents
Wend

MDIForm_Xsol.Show
Form_Xsol_03.Show

Unload Me

End Sub
