VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBillRec 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Bill Receipt"
   ClientHeight    =   6510
   ClientLeft      =   3150
   ClientTop       =   2415
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   7845
   Begin VB.Frame fremCust 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6315
      Left            =   2775
      TabIndex        =   14
      Top             =   1905
      Visible         =   0   'False
      Width           =   4560
      Begin MSFlexGridLib.MSFlexGrid CustLst 
         Height          =   6045
         Left            =   60
         TabIndex        =   15
         Top             =   105
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   10663
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   2701557
         ForeColorFixed  =   0
         BackColorSel    =   -2147483636
         BackColorBkg    =   14737632
         GridColor       =   0
         FocusRect       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "               Customer  Name                              "
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6150
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   7545
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   10
         Top             =   1875
         Width           =   5355
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   420
         Left            =   6195
         TabIndex        =   9
         Top             =   5565
         Width           =   1260
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save"
         Height          =   420
         Left            =   4875
         TabIndex        =   8
         Top             =   5580
         Width           =   1260
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1365
         Width           =   1125
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2550
         Left            =   1410
         TabIndex        =   5
         Top             =   2730
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   4498
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         Appearance      =   0
         FormatString    =   "Bill Number    | Bill Date  |         Amount    |         Paid    |      Balance"
      End
      Begin VB.TextBox TxtName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1455
         MaxLength       =   50
         TabIndex        =   2
         Top             =   795
         Width           =   5370
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   1
         Top             =   255
         Width           =   1125
      End
      Begin MSMask.MaskEdBox mskDt 
         Height          =   285
         Left            =   5490
         TabIndex        =   12
         Top             =   225
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Date"
         Height          =   195
         Left            =   4470
         TabIndex        =   13
         Top             =   225
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Narration"
         Height          =   195
         Left            =   225
         TabIndex        =   11
         Top             =   1920
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   1410
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
         Height          =   195
         Left            =   225
         TabIndex        =   4
         Top             =   855
         Width           =   1065
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No."
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Top             =   300
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmBillRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lFnd As Boolean
Private Sub Form_Load()
mskDt.Text = datecon(Date)
FillAcc CustLst, -1, " "
End Sub

Private Sub TxtName_Change()
fremCust.Visible = True
End Sub

Private Sub TxtName_GotFocus()
   SendKeys "{HOME}"
    SendKeys "+{END}"
    fremCust.Visible = True
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And TxtName <> "" And lFnd Then
lANew = False
   'Find GrpGrid, Format(txtPartyName.Text, ">"), 0  for change option we have check this
    TxtName.Text = CustLst.TextMatrix(CustLst.Row, 0)
    cCustCode = CustLst.TextMatrix(CustLst.Row, 1)
   ' cGroup = GrpGrid.TextMatrix(GrpGrid.Row, 2)
'    FillAdd cCustCode
    fremCust.Visible = False


ElseIf KeyAscii = 13 And txtPartyName <> "" And Not lFnd Then

    GrpGrid.Visible = False
    txtAdd1.SetFocus
    lANew = True
ElseIf KeyAscii = 13 And txtPartyName = "" Then
ElseIf KeyAscii = 27 Then
    Frame4.Visible = False
    flxgrd.SetFocus
End If

If KeyAscii = 13 And Not lFnd Then
fremCust.Visible = False
End If
End Sub


Private Sub TxtName_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyEscape And KeyCode <> vbKeyRight And KeyCode <> vbKeyLeft _
            And TxtName.Text <> "" And KeyCode <> vbKeyReturn Then
   lFnd = Find(CustLst, Format(TxtName.Text, ">"), 0, frmBill.fSp)
   'FillAdd GrpGrid.TextMatrix(GrpGrid.Row, 1)
End If
End Sub


