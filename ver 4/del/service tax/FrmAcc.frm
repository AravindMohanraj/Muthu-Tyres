VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmAcc 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Accounts Master"
   ClientHeight    =   6615
   ClientLeft      =   240
   ClientTop       =   945
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   11400
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6150
      Left            =   105
      TabIndex        =   12
      Top             =   120
      Width           =   7545
      Begin MSFlexGridLib.MSFlexGrid AreaGrid 
         Height          =   2250
         Left            =   7410
         TabIndex        =   22
         Top             =   5940
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3969
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "                                                        Area |Transport"
      End
      Begin MSFlexGridLib.MSFlexGrid GrpGrid 
         Height          =   2490
         Left            =   1830
         TabIndex        =   21
         Top             =   1995
         Visible         =   0   'False
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   4392
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         FocusRect       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "                                                          Account Name                                | Code"
      End
      Begin VB.TextBox TxtLocalRef 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1500
         MaxLength       =   50
         TabIndex        =   4
         Top             =   3360
         Width           =   3990
      End
      Begin VB.TextBox txtGrp 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1485
         MaxLength       =   50
         TabIndex        =   20
         Top             =   1455
         Width           =   5415
      End
      Begin VB.TextBox txtOpbal 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1500
         TabIndex        =   8
         Top             =   5550
         Width           =   1365
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1500
         MaxLength       =   50
         TabIndex        =   7
         Top             =   4965
         Width           =   2175
      End
      Begin VB.TextBox txtCity 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1485
         MaxLength       =   50
         TabIndex        =   6
         Top             =   4365
         Width           =   3990
      End
      Begin VB.TextBox TxtName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   1
         Top             =   810
         Width           =   5115
      End
      Begin VB.TextBox txtStreet 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1500
         MaxLength       =   50
         TabIndex        =   3
         Top             =   2865
         Width           =   4005
      End
      Begin VB.TextBox txtArea 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1485
         MaxLength       =   50
         TabIndex        =   5
         Top             =   3885
         Width           =   3990
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   420
         Left            =   6120
         TabIndex        =   10
         Top             =   5565
         Width           =   1260
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Apply"
         Height          =   420
         Left            =   4770
         TabIndex        =   9
         Top             =   5580
         Width           =   1260
      End
      Begin VB.TextBox txtContact 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1500
         MaxLength       =   20
         TabIndex        =   2
         Top             =   2235
         Width           =   3495
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1455
         MaxLength       =   5
         TabIndex        =   0
         Top             =   225
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
         Height          =   195
         Left            =   270
         TabIndex        =   19
         Top             =   5010
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Balance"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   5655
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   855
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Group"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   2925
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tin No"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2295
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   195
         Left            =   225
         TabIndex        =   13
         Top             =   300
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid AccLst 
      Height          =   6045
      Left            =   7695
      TabIndex        =   11
      Top             =   195
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   10663
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   "Account Name                                | Code"
   End
End
Attribute VB_Name = "FrmAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOpt As Integer, cGroup As String
Private Sub ClearData()
txtCode = ""
txtName = ""
txtOpbal = ""
txtStreet = ""
TxtLocalRef = ""
txtArea = ""
txtCity = ""
txtPhone = ""
txtContact = ""

End Sub

Private Sub StoreData()
Dim cNum As String
Set rsAcc = New ADODB.Recordset
rsAcc.Open "select * from acc0203d", Con, adOpenDynamic, adLockPessimistic
rsAcc.AddNew
Set rsNum = New ADODB.Recordset
If txtCode = "AUTO" Then
    rsNum.Open "select * from num0203d", Con, adOpenDynamic, adLockPessimistic
    cNum = Right(String(5, "0") + Trim(Str(Val(rsNum!facnum) + 1)), 5)
    rsNum!facnum = cNum
    rsNum.Update
Else
cNum = txtCode
End If
rsAcc!faccode = cNum
rsAcc!facname = txtName
rsAcc!facparent = cGroup + cNum
rsAcc!faclevel = (Len(cGroup + cNum) / 5) * -1
rsAcc!fopbal = Val(txtOpbal)
rsAcc.Update
rsAcc.Close
Set rsAdd = New ADODB.Recordset
rsAdd.Open "select * from add0203d", Con, adOpenDynamic, adLockPessimistic
rsAdd.AddNew
rsAdd!faccode = cNum
rsAdd!add1 = txtStreet
rsAdd!add2 = TxtLocalRef
rsAdd!add3 = txtArea
rsAdd!add4 = txtCity
rsAdd!PHone_no1 = txtPhone
rsAdd!cp = txtContact
rsAdd.Update
rsAdd.Close
End Sub

Private Sub stuffData(cCode As String)
Dim rsTacc As New ADODB.Recordset, cParent As String
TxtEnable False
    Set rsAcc = New ADODB.Recordset
    rsAcc.Open "select * from lstacname where faccode='" & cCode & "'", Con, adOpenStatic
        If Not rsAcc.EOF Then
        txtCode = cCode
        txtName = rsAcc!facname
           cParent = Left(Right(rsAcc!facparent, 10), 5)
           rsTacc.Open "select * from acc0203d where faccode='" & cParent & "'", Con, adOpenStatic
           If Not rsTacc.EOF Then
                txtGrp.Text = rsTacc!facname
                cGroup = rsTacc!facparent
           End If
           rsTacc.Close
           Set rsTacc = Nothing
           If Not IsNull(rsAcc!cp) Then txtContact = rsAcc!cp
           If Not IsNull(rsAcc!add1) Then txtStreet = rsAcc!add1
           If Not IsNull(rsAcc!add2) Then TxtLocalRef = rsAcc!add2
           If Not IsNull(rsAcc!add3) Then txtArea = rsAcc!add3
           If Not IsNull(rsAcc!add4) Then txtCity = rsAcc!add4
           If Not IsNull(rsAcc!PHone_no1) Then txtPhone = rsAcc!PHone_no1
           
        End If
    rsAcc.Close
    Set rsAcc = Nothing
End Sub

Private Sub TxtEnable(lTF As Boolean)
txtCode.Enabled = lTF
txtName.Enabled = lTF
txtGrp.Enabled = lTF
txtContact.Enabled = lTF
txtStreet.Enabled = lTF
TxtLocalRef.Enabled = lTF
txtArea.Enabled = lTF
txtCity.Enabled = lTF
txtPhone.Enabled = lTF
txtOpbal.Enabled = lTF
End Sub

Private Sub updatedata()
Dim cNum As String
Set rsAcc = New ADODB.Recordset
Set rsAdd = New ADODB.Recordset

rsAcc.Open "select * from acc0203d where faccode='" & txtCode & "'", Con, adOpenDynamic, adLockPessimistic
If Not rsAcc.EOF Then
rsAcc!facname = txtName
rsAcc!facparent = cGroup + cNum
rsAcc!fopbal = Val(txtOpbal)
rsAcc.Update
End If
rsAcc.Close




rsAdd.Open "select * from add0203d where faccode='" & txtCode & "'", Con, adOpenDynamic, adLockPessimistic
If Not rsAdd.EOF Then
rsAdd!add1 = txtStreet
rsAdd!add2 = TxtLocalRef
rsAdd!add3 = txtArea
rsAdd!add4 = txtCity
rsAdd!PHone_no1 = txtPhone
rsAdd!cp = txtContact
rsAdd.Update
End If
rsAdd.Close
End Sub

Private Sub AccLst_EnterCell()
stuffData AccLst.TextMatrix(AccLst.Row, 1)
txtName.Enabled = True
txtName.SetFocus

nOpt = 2
End Sub


Private Sub AccLst_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtName.Enabled = True
txtName.SetFocus
End If
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
ClearData
nOpt = 1
End Sub

Private Sub Command3_Click()
If nOpt = 1 Then
StoreData
ElseIf nOpt = 2 Then
updatedata
End If
nOpt = 1
ClearData
TxtEnable False
txtCode.Enabled = True
txtCode.SetFocus
End Sub

Private Sub Form_Load()
nOpt = 1
TxtEnable False
txtCode.Enabled = True
FillAcc Grpgrid, 1, "                                                          Account Name                                | Code"
FillAcc AccLst, -1, "Account Name                                | Code"
'FillArea AreaGrid, "                                                        Area |Transport"
End Sub



Private Sub GrpGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtGrp = Grpgrid.TextMatrix(Grpgrid.Row, 0)
txtGrp_KeyPress 13
End If
End Sub


Private Sub ItemLst_Click()
stuffData txtCode
End Sub

Private Sub txtArea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtArea.Text <> "" Then
   ' Find AreaGrid, Format(txtArea.Text, ">"), 0
    'txtArea.Text = AreaGrid.TextMatrix(AreaGrid.Row, 0)
    'AreaGrid.Visible = False
    txtCity.Enabled = True
    txtCity.SetFocus
End If

End Sub


Private Sub txtArea_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode <> vbKeyEscape And KeyCode <> vbKeyRight And KeyCode <> vbKeyLeft _
 '           And txtArea.Text <> "" And KeyCode <> vbKeyReturn Then
  '  Find AreaGrid, Format(txtArea.Text, ">"), 0
'End If
End Sub


Private Sub txtCity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPhone.Enabled = True
txtPhone.SetFocus
End If
End Sub


Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtCode = "" Then
    txtCode = "AUTO"
    txtName.Enabled = True
    txtName.SetFocus
ElseIf KeyAscii = 13 And txtCode <> "" Then
    txtCode = UCase(txtCode)
    Set rsAcc = New ADODB.Recordset
    rsAcc.Open "select * from acc0203d where faccode='" & txtCode & "'", Con, adOpenStatic
        If Not rsAcc.EOF Then
            MsgBox "Code Already exists", vbCritical
            txtCode = ""
        Else
            txtName.Enabled = True
            txtName.SetFocus
        End If
    rsAcc.Close
    Set rsAcc = Nothing
End If
End Sub


Private Sub txtContact_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtStreet.Enabled = True
txtStreet.SetFocus
End If
End Sub


Private Sub txtGrp_GotFocus()
    SendKeys "{HOME}"
    SendKeys "+{END}"

End Sub

Private Sub txtGrp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
 Grpgrid.SetFocus
ElseIf KeyCode = vbKeyUp Then
 Grpgrid.SetFocus
End If
 
End Sub

Private Sub txtGrp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtGrp.Text <> "" Then
    Find Grpgrid, Format(txtGrp.Text, ">"), 0
    txtGrp.Text = Grpgrid.TextMatrix(Grpgrid.Row, 0)
    cGroup = Grpgrid.TextMatrix(Grpgrid.Row, 2)
    Grpgrid.Visible = False
    txtContact.Enabled = True
    txtContact.SetFocus
End If

End Sub


Private Sub txtGrp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyEscape And KeyCode <> vbKeyRight And KeyCode <> vbKeyLeft _
            And txtGrp.Text <> "" And KeyCode <> vbKeyReturn Then
    Find Grpgrid, Format(txtGrp.Text, ">"), 0
End If
   
End Sub


Private Sub TxtLocalRef_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtArea.Enabled = True
txtArea.SetFocus
End If
End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtName <> "" Then
txtName = UCase(txtName)
    If nOpt = 1 Then
        Set rsAcc = New ADODB.Recordset
        rsAcc.Open "select * from acc0203d where facname='" & txtName & "'", Con, adOpenStatic
            If Not rsAcc.EOF Then
                MsgBox "Name  Already exists", vbCritical
                txtName = ""
            Else
                txtGrp.Enabled = True
                txtGrp.SetFocus
            End If
        rsAcc.Close
        Set rsAcc = Nothing
    ElseIf nOpt = 2 Then
        txtGrp.Enabled = True
        txtGrp.SetFocus
    End If
End If
End Sub





Private Sub txtOpbal_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
Command3.Enabled = True
Command3.SetFocus
End If
End Sub


Private Sub txtPhone_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
txtOpbal.Enabled = True
txtOpbal.SetFocus
End If
End Sub


Private Sub txtStreet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtStreet <> "" Then
TxtLocalRef.Enabled = True
TxtLocalRef.SetFocus
ElseIf KeyAscii = 13 And txtStreet = "" Then
txtPhone.Enabled = True
txtPhone.SetFocus

End If
End Sub


