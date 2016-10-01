VERSION 5.00
Begin VB.MDIForm MDIForm1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "WinIV-Billing Software"
   ClientHeight    =   8295
   ClientLeft      =   1785
   ClientTop       =   1995
   ClientWidth     =   12435
   LinkTopic       =   "MDIForm1"
   Picture         =   "Mainmenu.frx":0000
   Begin VB.Menu mnuMst 
      Caption         =   "Master"
      Begin VB.Menu mnuItem 
         Caption         =   "Item Master"
      End
   End
   Begin VB.Menu mnuSal 
      Caption         =   "Sales"
      Begin VB.Menu mnuBill 
         Caption         =   "Bill"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu mnuuti 
      Caption         =   "Utility"
      Begin VB.Menu mnuAcpanel 
         Caption         =   "Access Control Panel"
      End
   End
   Begin VB.Menu mnuHlp 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Me.Width = 12555
Me.Height = 9450
End Sub

Private Sub mnuAcpanel_Click()

frmAccCp.Show
End Sub


Private Sub mnuBill_Click()
frmBill.Height = 8115
frmBill.Width = 10515
frmBill.Show
End Sub

Private Sub Picture1_Click()

End Sub


