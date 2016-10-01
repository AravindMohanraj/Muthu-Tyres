VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIteMaster 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Item Master"
   ClientHeight    =   6465
   ClientLeft      =   1965
   ClientTop       =   2760
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   11865
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6270
      Left            =   75
      TabIndex        =   12
      Top             =   105
      Width           =   7545
      Begin VB.CheckBox chkStkReq 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Do not Check for Stock availablity"
         Height          =   345
         Left            =   210
         TabIndex        =   30
         Tag             =   "For Service Bill Stock not required so Check this box"
         Top             =   5745
         Width           =   2790
      End
      Begin VB.TextBox txtTaxAccountCode 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7305
         TabIndex        =   29
         Top             =   5595
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid fgTax 
         Height          =   1875
         Left            =   4140
         TabIndex        =   28
         Top             =   2835
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3307
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   16777215
         FocusRect       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "Tax                                 | Code   |  Tax P "
      End
      Begin VB.Frame frmSearch 
         Caption         =   "Search"
         Height          =   1095
         Left            =   2400
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   4935
         Begin VB.TextBox txtSearch 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   360
            MaxLength       =   50
            TabIndex        =   27
            Top             =   480
            Width           =   4050
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grplist 
         Height          =   3390
         Left            =   2745
         TabIndex        =   24
         Top             =   1890
         Visible         =   0   'False
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   5980
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   16777215
         FocusRect       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "Product  Name                                  | Code   |  Stock       | MRP "
      End
      Begin VB.CheckBox chkITax 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Selling Rate Inclusive Tax"
         Height          =   345
         Left            =   210
         TabIndex        =   25
         Top             =   5205
         Width           =   2790
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1290
         MaxLength       =   5
         TabIndex        =   0
         Top             =   255
         Width           =   1125
      End
      Begin VB.TextBox TxtBar 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   3
         Top             =   2205
         Width           =   3495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save"
         Height          =   420
         Left            =   4155
         TabIndex        =   10
         Top             =   5415
         Width           =   1260
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   420
         Left            =   5460
         TabIndex        =   11
         Top             =   5385
         Width           =   1260
      End
      Begin VB.TextBox txtPrt 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5400
         TabIndex        =   9
         Top             =   4515
         Width           =   1335
      End
      Begin VB.TextBox txtSr 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1395
         TabIndex        =   8
         Top             =   4545
         Width           =   1335
      End
      Begin VB.TextBox txtTax 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5415
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3765
         Width           =   1335
      End
      Begin VB.TextBox txtRL 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1395
         TabIndex        =   6
         Top             =   3660
         Width           =   1335
      End
      Begin VB.TextBox txtPack 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5415
         TabIndex        =   5
         Top             =   2955
         Width           =   1335
      End
      Begin VB.ComboBox cboUof 
         Height          =   315
         ItemData        =   "frmIteMaster.frx":0000
         Left            =   1860
         List            =   "frmIteMaster.frx":0010
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   2940
         Width           =   1455
      End
      Begin VB.TextBox txtGrp 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1320
         TabIndex        =   2
         Top             =   1515
         Width           =   5370
      End
      Begin VB.TextBox TxtName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   1
         Top             =   810
         Width           =   5370
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   195
         Left            =   225
         TabIndex        =   22
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   2295
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Rate"
         Height          =   195
         Left            =   4260
         TabIndex        =   20
         Top             =   4590
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Rate"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   4620
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax %"
         Height          =   195
         Left            =   4305
         TabIndex        =   18
         Top             =   3765
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reorder Level"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   3690
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Package"
         Height          =   195
         Left            =   4290
         TabIndex        =   16
         Top             =   3000
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit of Measurement"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   3000
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Group"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   855
         Width           =   765
      End
   End
   Begin MSFlexGridLib.MSFlexGrid ProductList 
      Height          =   6165
      Left            =   7635
      TabIndex        =   23
      Top             =   180
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   10874
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorBkg    =   16777215
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   "Product  Name                                  | Code   |  Stock       | MRP "
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuclose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmIteMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOpt As Integer, cGroup As String
Dim cOldname As String
Private Sub ClearData()
txtCode = ""
TxtName = ""
txtSr = ""
txtGrp = ""
txtTax = ""
txtTaxAccountCode = ""
chkStkReq.Value = 0
End Sub

Private Sub EditCode()
Dim cNum As String
Set rsItem = New ADODB.Recordset
rsItem.Open "select * from ite0203d where faccode='" & txtCode & "'", Con, adOpenDynamic, adLockPessimistic

rsItem!facname = TxtName
rsItem!facparent = cGroup + txtCode
rsItem!faclevel = (Len(cGroup + txtCode) / 5) * -1
'rsItem!fopbal = Val(txtOpbal)
rsItem!fSp = Val(txtSr)
rsItem!ftax = Val(txtTax.Text)
rsItem!ftaxacccode = txtTaxAccountCode.Text
rsItem!FstockRequired = chkStkReq.Value
rsItem.Update
rsItem.Close
End Sub

Private Sub StoreData()
Dim cNum As String
Set rsItem = New ADODB.Recordset
rsItem.Open "select * from ite0203d", Con, adOpenDynamic, adLockPessimistic
rsItem.AddNew
Set rsNum = New ADODB.Recordset
If txtCode = "AUTO" Then
    rsNum.Open "select * from num0203d", Con, adOpenDynamic, adLockPessimistic
    cNum = Right(String(5, "0") + Trim(Str(Val(rsNum!fivnum) + 1)), 5)
    rsNum!fivnum = cNum
    rsNum.Update
Else
cNum = txtCode
End If
rsItem!faccode = cNum
rsItem!facname = TxtName
rsItem!facparent = cGroup + cNum
rsItem!faclevel = (Len(cGroup + cNum) / 5) * -1
'rsItem!fopbal = Val(txtOpbal)
rsItem!fSp = Val(txtSr)
rsItem!ftax = Val(txtTax.Text)
rsItem!ftaxacccode = txtTaxAccountCode.Text
rsItem!FstockRequired = chkStkReq.Value

rsItem.Update
rsItem.Close

If lSuBra Then
Set rsBranch = New ADODB.Recordset
Set rsBrStk = New ADODB.Recordset
rsBranch.Open "select * from branch", Con, adOpenDynamic, adLockPessimistic
If Not rsBranch.EOF Then
rsBrStk.Open "select * from brstk", Con, adOpenDynamic, adLockPessimistic

If Not rsBranch.BOF Then rsBranch.MoveFirst
Do While Not rsBranch.EOF
rsBrStk.AddNew
rsBrStk!fbranch = rsBranch!fbranch
rsBrStk!faccode = cNum
rsBrStk!fbal = 0
rsBrStk.Update

rsBranch.MoveNext
Loop
rsBrStk.Close
End If
rsBranch.Close
Set rsBrStk = Nothing
Set rsBranch = Nothing
End If


End Sub




Private Sub stuffData()
Dim cNum As String
Set rsItem = New ADODB.Recordset
Set rsAcc = New ADODB.Recordset

rsItem.Open "select * from ite0203d where faccode='" & ProductList.TextMatrix(ProductList.Row, 1) & "'", Con, adOpenDynamic, adLockPessimistic

txtCode = rsItem!faccode
TxtName = rsItem!facname
cOldname = rsItem!facname
rsAcc.Open "select * from ite0203d where faccode='" & Left(rsItem!facparent, 5) & "'", Con, adOpenStatic
If Not rsAcc.EOF Then
If Not rsAcc.BOF Then rsAcc.MoveFirst
txtGrp.Text = rsAcc!facname
Else
txtGrp.Text = ""
End If
rsAcc.Close

If Not IsNull(rsItem!fSp) Then txtSr = rsItem!fSp
If Not IsNull(rsItem!fbarcode) Then TxtBar = rsItem!fbarcode
If Not IsNull(rsItem!funit) Then cboUof.Text = rsItem!funit
If Not IsNull(rsItem!fweight) Then txtPack = rsItem!fweight
If Not IsNull(rsItem!freqty) Then txtRL = rsItem!freqty
txtTax = rsItem!ftax
txtPrt = rsItem!fcp
 If Not IsNull(rsItem!ftaxacccode) Then txtTaxAccountCode = rsItem!ftaxacccode
chkStkReq.Value = IIf(rsItem!FstockRequired, 1, 0)
chkITax.Value = rsItem!ftaxir


rsItem.Close




End Sub

Private Sub cboUof_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPack.SetFocus
End If
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
nOpt = 1
ClearData
txtCode.SetFocus

End Sub

Private Sub Command3_Click()
If nOpt = 1 Then
StoreData
ElseIf nOpt = 2 Then
EditCode
End If

nOpt = 1
ClearData
FillItem Grplist, 1, "Product  Name                                  | Code   |  Stock       | MRP "
FillItem ProductList, -1, "Product  Name                                  | Code   |  Stock       | MRP "

txtCode.SetFocus
End Sub

Private Sub fgTax_Click()
txtTax.Text = fgTax.TextMatrix(fgTax.Row, 2)
txtTaxAccountCode.Text = fgTax.TextMatrix(fgTax.Row, 1)
fgTax.Visible = False
txtSr.SetFocus

End Sub

Private Sub fgTax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtTax.Text = fgTax.TextMatrix(fgTax.Row, 2)
txtTaxAccountCode.Text = fgTax.TextMatrix(fgTax.Row, 1)
fgTax.Visible = False
txtSr.SetFocus

End If
End Sub


Private Sub Form_Load()
FillItem Grplist, 1, "Product  Name                                  | Code   |  Stock       | MRP "
FillItem ProductList, -1, "Product  Name                                  | Code   |  Stock       | MRP "
FillTax fgTax
nOpt = 1

End Sub



Private Sub Grplist_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtGrp.Text = Grplist.TextMatrix(Grplist.Row, 0)
    txtGrp_KeyPress 13
End If
End Sub


Private Sub mnuclose_Click()
Unload Me
End Sub

Private Sub mnuSearch_Click()
If frmSearch.Visible = False Then
frmSearch.Visible = True
txtSearch.SetFocus
ElseIf frmSearch.Visible = True Then
frmSearch.Visible = True
txtSearch.SetFocus
End If
End Sub

Private Sub ProductList_EnterCell()
stuffData
nOpt = 2
End Sub


Private Sub Text1_Change()

End Sub

Private Sub productlist_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If frmSearch.Visible = True Then frmSearch.Visible = False
stuffData
nOpt = 2
TxtName.SetFocus
End If
End Sub

Private Sub TxtBar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cboUof.SetFocus
End If
End Sub


Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtCode = "" Then
    txtCode = "AUTO"
    TxtName.Enabled = True
    TxtName.SetFocus
ElseIf KeyAscii = 13 And txtCode <> "" Then
    txtCode = Left(UCase(txtCode) + String(5, "0"), 5)
    Set rsItem = New ADODB.Recordset
    rsItem.Open "select * from ite0203d where faccode='" & txtCode & "'", Con, adOpenStatic
        If Not rsItem.EOF Then
            MsgBox "Code Already exists", vbCritical
            txtCode = ""
        Else
            TxtName.Enabled = True
            TxtName.SetFocus
        End If
    rsItem.Close
    Set rsItem = Nothing
End If

End Sub


Private Sub txtGrp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
Grplist.SetFocus
End If

End Sub

Private Sub txtGrp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtGrp.Text <> "" Then
    Find Grplist, Format(txtGrp.Text, ">"), 0
    txtGrp.Text = Grplist.TextMatrix(Grplist.Row, 0)
    cGroup = Grplist.TextMatrix(Grplist.Row, 1)
    Grplist.Visible = False
    TxtBar.Enabled = True
    TxtBar.SetFocus
End If

End Sub


Private Sub txtGrp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyEscape And KeyCode <> vbKeyRight And KeyCode <> vbKeyLeft _
            And txtGrp.Text <> "" And KeyCode <> vbKeyReturn Then
    Find Grplist, Format(txtGrp.Text, ">"), 0
End If
   
End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And TxtName <> "" And nOpt = 1 Then
TxtName = UCase(TxtName)
Set rsItem = New ADODB.Recordset
    rsItem.Open "select * from ite0203d where facname='" & TxtName & "'", Con, adOpenStatic
        If Not rsItem.EOF Then
            MsgBox "Name  Already exists", vbCritical
            TxtName = ""
        Else
            txtGrp.Enabled = True
            txtGrp.SetFocus
        End If
    rsItem.Close
    Set rsItem = Nothing
ElseIf KeyAscii = 13 And TxtName <> "" And nOpt = 2 Then
TxtName = UCase(TxtName)
If UCase(cOldname) <> TxtName.Text Then

Set rsItem = New ADODB.Recordset
    rsItem.Open "select * from ite0203d where facname='" & TxtName & "'", Con, adOpenStatic
        If Not rsItem.EOF Then
            MsgBox "Name  Already exists", vbCritical
            TxtName = ""
        Else
            txtGrp.Enabled = True
            txtGrp.SetFocus
        End If
    rsItem.Close
    Set rsItem = Nothing
Else
            txtGrp.Enabled = True
            txtGrp.SetFocus

End If
End If
End Sub

Private Sub txtPack_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
txtRL.SetFocus
End If
End Sub


Private Sub txtPrt_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
Command3.SetFocus
End If
End Sub


Private Sub txtRL_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
txtTax.SetFocus
End If
End Sub


Private Sub txtSearch_GotFocus()
    SendKeys "{HOME}"
    SendKeys "+{END}"

End Sub


Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
ProductList.SetFocus
ElseIf KeyCode = vbKeyUp Then
ProductList.SetFocus
End If

End Sub


Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtSearch.Text <> "" Then
stuffData
nOpt = 2
frmSearch.Visible = False
TxtName.SetFocus
End If
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyEscape And KeyCode <> vbKeyRight And KeyCode <> vbKeyLeft _
            And txtSearch.Text <> "" And KeyCode <> vbKeyReturn Then
   lFnd = Find(ProductList, Format(txtSearch.Text, ">"), 0, frmBill.fSp)
End If

End Sub

Private Sub txtSr_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
txtPrt.SetFocus
End If
End Sub


Private Sub txtTax_GotFocus()
fgTax.Visible = True
fgTax.SetFocus
End Sub

Private Sub txtTax_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
txtSr.SetFocus
End If
End Sub


