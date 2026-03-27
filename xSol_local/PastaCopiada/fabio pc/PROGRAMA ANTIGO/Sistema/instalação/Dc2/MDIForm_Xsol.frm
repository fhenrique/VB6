VERSION 5.00
Begin VB.MDIForm MDIForm_Xsol 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "MDIForm_Xsol.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu nmenu 
      Caption         =   "Menu"
      Begin VB.Menu ntecnico 
         Caption         =   "Cadastro de Tecnico"
         Shortcut        =   ^T
      End
      Begin VB.Menu nclientes 
         Caption         =   "Cadastro de Clientes"
         Shortcut        =   ^C
      End
      Begin VB.Menu nlogon 
         Caption         =   "Logon On"
         Shortcut        =   ^L
      End
      Begin VB.Menu nSair 
         Caption         =   "Sair"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "MDIForm_Xsol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()

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

End Sub

Private Sub nclientes_Click()

Form_Xsol_02.Show

End Sub

Private Sub nlogon_Click()

Form_Xsol_03.Show

End Sub


Private Sub nSair_Click()

End

End Sub

Private Sub ntecnico_Click()

Form_Xsol_04.Show

End Sub
