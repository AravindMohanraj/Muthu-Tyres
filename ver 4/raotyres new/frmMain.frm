VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "WinIv"
   ClientHeight    =   8190
   ClientLeft      =   945
   ClientTop       =   1965
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar MainSb 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   7860
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Date:"
            TextSave        =   "Date:"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu0 
      Caption         =   "Master"
      Begin VB.Menu mnu1 
         Caption         =   "Account Master"
      End
      Begin VB.Menu mnu2 
         Caption         =   "Product Master"
      End
      Begin VB.Menu mnu3 
         Caption         =   "Quick Product Edit"
      End
      Begin VB.Menu mnu4 
         Caption         =   "Branch Creation"
      End
      Begin VB.Menu mnu5 
         Caption         =   "Godown Creation"
      End
      Begin VB.Menu mnu6 
         Caption         =   "Bill of Materials"
      End
      Begin VB.Menu mnuXit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnu7 
      Caption         =   "Sales"
      Begin VB.Menu mnu8 
         Caption         =   "Bill"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnu9 
         Caption         =   "Delivery Challan"
      End
      Begin VB.Menu mnu10 
         Caption         =   "Order"
      End
      Begin VB.Menu mnu11 
         Caption         =   "Sales Return"
      End
   End
   Begin VB.Menu mnu12 
      Caption         =   "Purchase"
      Begin VB.Menu mnu13 
         Caption         =   "Purchase Entry"
      End
      Begin VB.Menu mnu14 
         Caption         =   "Purchase Order"
      End
      Begin VB.Menu mnu15 
         Caption         =   "Purchase Return"
      End
   End
   Begin VB.Menu mnu16 
      Caption         =   "Inventory"
      Begin VB.Menu mnu17 
         Caption         =   "Goods Receipt\Issue"
      End
      Begin VB.Menu mnu18 
         Caption         =   "Godown Transfer"
      End
      Begin VB.Menu mnu19 
         Caption         =   "Branch Transfer"
      End
      Begin VB.Menu mnu20 
         Caption         =   "Dummy"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu21 
         Caption         =   "Dummy"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu22 
         Caption         =   "Dummy"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu23 
      Caption         =   "Accounts"
      Begin VB.Menu mnu24 
         Caption         =   "Bill Receipt"
      End
      Begin VB.Menu mnu25 
         Caption         =   "Receipt"
      End
      Begin VB.Menu mnu26 
         Caption         =   "Bill Payment"
      End
      Begin VB.Menu mnu27 
         Caption         =   "Payment"
      End
      Begin VB.Menu mnu28 
         Caption         =   "Debit Note"
      End
      Begin VB.Menu mnu29 
         Caption         =   "Credit Note"
      End
   End
   Begin VB.Menu mnu30 
      Caption         =   "Report"
      Begin VB.Menu mnu31 
         Caption         =   "Sales Report"
      End
      Begin VB.Menu mnu32 
         Caption         =   "Day Book"
      End
      Begin VB.Menu mnu33 
         Caption         =   "Cash Book"
      End
      Begin VB.Menu mnu34 
         Caption         =   "Bank Book"
      End
      Begin VB.Menu mnu35 
         Caption         =   "Stock Statement"
      End
      Begin VB.Menu mnu36 
         Caption         =   "General ledger"
      End
      Begin VB.Menu mnu37 
         Caption         =   "Item Stock Card"
      End
      Begin VB.Menu mnu38 
         Caption         =   "Statement of Account"
      End
      Begin VB.Menu mnu39 
         Caption         =   "Purchase"
      End
      Begin VB.Menu mnu40 
         Caption         =   "dummy"
      End
      Begin VB.Menu mnu41 
         Caption         =   "dummy"
      End
      Begin VB.Menu mnu42 
         Caption         =   "dummy"
      End
      Begin VB.Menu mnu43 
         Caption         =   "dummy"
      End
      Begin VB.Menu mnu44 
         Caption         =   "dummy"
      End
   End
   Begin VB.Menu mnu45 
      Caption         =   "Utility"
      Begin VB.Menu mnu46 
         Caption         =   "Create New User"
      End
      Begin VB.Menu mnu47 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnu48 
         Caption         =   "Access Control Panel"
      End
      Begin VB.Menu mnu49 
         Caption         =   "Data Backup"
      End
      Begin VB.Menu mnu50 
         Caption         =   "Data Repair and Compression"
      End
      Begin VB.Menu mnuUIFB 
         Caption         =   "Update Item for Branch"
      End
   End
   Begin VB.Menu mnu60 
      Caption         =   "Company"
      Begin VB.Menu mnu61 
         Caption         =   "Company List"
      End
      Begin VB.Menu mnu62 
         Caption         =   "New Company"
      End
   End
   Begin VB.Menu mnuwindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Main

nUser = 2
cUser = "ganesh"

IntMenu nUser

Me.Width = 12555
Me.Height = 9450

StartSet
End Sub


Private Sub Form_Terminate()
'MsgBox "This will close the software"
End Sub

Private Sub mnu1_Click()
FrmAcc.Show
End Sub

Private Sub mnu13_Click()
frmPurchase.Show
End Sub

Private Sub mnu17_Click()
frmStores.Show
Limitforform "mnu17", frmStores, nUser
End Sub

Private Sub mnu2_Click()
frmIteMaster.Show
End Sub

Private Sub mnu24_Click()
frmBillRec.Show
End Sub

Private Sub mnu25_Click()
frmVoucher.Show
End Sub

Private Sub mnu31_Click()
frmRptSal.Show
End Sub

Private Sub mnu35_Click()
frmStkStmnt.Show
End Sub

Private Sub mnu37_Click()
frmIsCrd.Show
End Sub

Private Sub mnu39_Click()
frmRptPur.Show
End Sub

Private Sub mnu4_Click()
frmBranch.Show
End Sub

Private Sub mnu46_Click()
frmUserCreation.Show
End Sub

Private Sub mnu48_Click()
FrmAccCtrl.Show
End Sub

Private Sub mnu5_Click()
frmGodown.Show
End Sub

Private Sub mnu6_Click()
frmBom.Show
End Sub

Private Sub mnu8_Click()
'frmBill.Height = 8115
'frmBill.Width = 12000
frmBill.Show
Limitforform "mnu8", frmBill, nUser

End Sub


Private Sub mnuAcp_Click()
frmAccCp.Show
End Sub




Private Sub mnuUIFB_Click()
frmRunBranch.Show
End Sub

Private Sub mnuXit_Click()
Unload Me
End
End Sub


