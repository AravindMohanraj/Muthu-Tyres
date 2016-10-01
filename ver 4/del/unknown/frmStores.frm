VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmStores 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Stores Voucher"
   ClientHeight    =   6420
   ClientLeft      =   2085
   ClientTop       =   1530
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   7695
   Begin VB.Frame fremIte 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3825
      Left            =   7935
      TabIndex        =   12
      Top             =   2445
      Visible         =   0   'False
      Width           =   3795
      Begin VB.TextBox txtItemLst 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   45
         TabIndex        =   13
         Top             =   75
         Width           =   3615
      End
      Begin MSFlexGridLib.MSFlexGrid productlist 
         Height          =   3090
         Left            =   60
         TabIndex        =   14
         Top             =   615
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   5450
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   2701557
         ForeColorFixed  =   0
         BackColorSel    =   8388608
         BackColorBkg    =   14737632
         GridColor       =   0
         FocusRect       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "Product  Name                                  | Stock       | MRP"
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6315
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   7545
      Begin VB.Frame fremGdwn 
         Caption         =   "Godown"
         Height          =   2730
         Left            =   4965
         TabIndex        =   17
         Top             =   2895
         Visible         =   0   'False
         Width           =   2115
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1065
            MaxLength       =   5
            TabIndex        =   22
            Top             =   2400
            Width           =   960
         End
         Begin VB.TextBox txtGtqty 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   825
            MaxLength       =   5
            TabIndex        =   19
            Top             =   210
            Width           =   960
         End
         Begin MSFlexGridLib.MSFlexGrid GdwnGrid 
            Height          =   1725
            Left            =   90
            TabIndex        =   18
            Top             =   615
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   3043
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            BackColorFixed  =   -2147483648
            BackColorBkg    =   16777215
            GridLinesFixed  =   1
            ScrollBars      =   2
            Appearance      =   0
            FormatString    =   "Godown               | Qty  |GC"
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Qty"
            Height          =   195
            Left            =   105
            TabIndex        =   20
            Top             =   255
            Width           =   645
         End
      End
      Begin VB.ComboBox cboBranch2 
         Height          =   315
         ItemData        =   "frmStores.frx":0000
         Left            =   4320
         List            =   "frmStores.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2100
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.ComboBox cboBranch 
         Height          =   315
         ItemData        =   "frmStores.frx":0004
         Left            =   1440
         List            =   "frmStores.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2100
         Width           =   1830
      End
      Begin VB.TextBox txtStock 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5475
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   300
         Width           =   1380
      End
      Begin VB.TextBox txtRmk 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2670
         Width           =   5355
      End
      Begin VB.ComboBox cboTag 
         Height          =   315
         ItemData        =   "frmStores.frx":0008
         Left            =   1470
         List            =   "frmStores.frx":0024
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1500
         Width           =   2700
      End
      Begin VB.TextBox txtVrno 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1455
         MaxLength       =   5
         TabIndex        =   2
         Top             =   915
         Width           =   1125
      End
      Begin VB.ComboBox cboVrType 
         Height          =   315
         ItemData        =   "frmStores.frx":0083
         Left            =   1425
         List            =   "frmStores.frx":008D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   315
         Width           =   1185
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   285
         Left            =   5475
         TabIndex        =   3
         Top             =   765
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid flxgrd 
         Height          =   2820
         Left            =   270
         TabIndex        =   11
         Top             =   3165
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   4974
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   -2147483648
         BackColorBkg    =   16777215
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         Appearance      =   0
         FormatString    =   "SlNo. |                  Product Name              | Quantity     | Rate |  Amount | Code           "
      End
      Begin VB.Label lblBranch2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch To"
         Height          =   195
         Left            =   3450
         TabIndex        =   26
         Top             =   2145
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label lblBranch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch "
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   2145
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock In hand"
         Height          =   195
         Left            =   4350
         TabIndex        =   16
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Left            =   225
         TabIndex        =   10
         Top             =   2715
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Tag"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1575
         Width           =   930
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No."
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Date"
         Height          =   195
         Left            =   4320
         TabIndex        =   5
         Top             =   810
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Type"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   390
         Width           =   1005
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gd 
      Height          =   6450
      Left            =   7710
      TabIndex        =   21
      Top             =   1485
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   11377
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   -2147483648
      BackColorBkg    =   16777215
      GridLinesFixed  =   1
      ScrollBars      =   2
      Appearance      =   0
      FormatString    =   "Godown               | Qty  |  code"
   End
   Begin VB.Menu mnu1 
      Caption         =   "File"
      Begin VB.Menu mnu2 
         Caption         =   "New"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu3 
         Caption         =   "Edit"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnu4 
         Caption         =   "Cancel"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnu5 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnu6 
         Caption         =   "Print"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnu7 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu8 
         Caption         =   "Save"
         Shortcut        =   {F2}
         Visible         =   0   'False
      End
      Begin VB.Menu mnu9 
         Caption         =   "-"
      End
      Begin VB.Menu mnu10 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnu11 
      Caption         =   "Edit"
      Begin VB.Menu mnu12 
         Caption         =   "Remove Line"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "frmStores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim nOpt As Single




Private Sub ClearData()

FlxGrd.Rows = 2
FlxGrd.Clear
FlxGrd.FormatString = "SlNo. |                  Product Name              | Quantity     | Rate |  Amount | Code           "
FillGdwn

gd.Rows = 2
gd.Clear
gd.FormatString = "Godown               | Qty  |  code"
txtVrno = ""
txtNote = ""
        Set rsNum = New ADODB.Recordset
        rsNum.Open "select * from num0203d", Con, adOpenDynamic, adLockPessimistic
            If Not rsNum.EOF Then
                If cboVrType.ListIndex = 0 Then
                   nNum = rsNum!fireceipt + 1
                ElseIf cboVrType.ListIndex = 1 Then
                   nNum = rsNum!fiissue + 1
                End If
            End If
        rsNum.Close
        txtVrno = nNum

End Sub


Private Function DeleteData() As Boolean
Dim nR As Integer, nR1 As Integer, nVt As Single
Set rsIled = New ADODB.Recordset
Set rsGdTrn = New ADODB.Recordset


nVt = IIf(cboVrType.ListIndex = 0, VrType("GREC"), VrType("GISS"))

rsIled.Open "select * from lstiled where val(fvrtype)='" & nVt & "' and val(fvrno)='" & Val(txtVrno) & "'", Con, adOpenStatic
If Not rsIled.EOF Then
DeleteData = True

If Not rsIled.BOF Then rsIled.MoveFirst

    Do While Not rsIled.EOF
MinusStock rsIled!faccode, rsIled!fqty
            
                    If lSuGd Then
                        rsGdTrn.Open "select * from gdtrans where val(fvrtype)='" & nVt & "' and faccode= '" & rsIled!faccode & "' and val(fvrno)='" & Val(txtVrno) & "'", Con, adOpenStatic
                        If Not rsGdTrn.EOF Then
                            If Not rsGdTrn.BOF Then rsGdTrn.MoveFirst
                                 MinusGdStk rsIled!faccode, rsGdTrn!fqty, rsGdTrn!fgodown
                        End If
                        rsGdTrn.Close
                    End If
            
            
            rsIled.MoveNext
    Loop
Else
    DeleteData = False
End If
rsIled.Close

rsIled.Open "delete from ile0203d where val(fvrtype)='" & nVt & "' and val(fvrno)='" & Val(txtVrno) & "'", Con, adOpenDynamic, adLockPessimistic
rsGdTrn.Open "delete from gdtrans where val(fvrtype)='" & nVt & "' and  val(fvrno)='" & Val(txtVrno) & "'", Con, adOpenDynamic, adLockPessimistic



End Function

Private Function fillData() As Boolean
Dim nR As Integer, nR1 As Integer, nVt As Single
Set rsIled = New ADODB.Recordset
Set rsGdTrn = New ADODB.Recordset


nVt = IIf(cboVrType.ListIndex = 0, VrType("GREC"), VrType("GISS"))

rsIled.Open "select * from lstiled where val(fvrtype)='" & nVt & "' and val(fvrno)='" & Val(txtVrno) & "'", Con, adOpenStatic
If Not rsIled.EOF Then
fillData = True

If Not rsIled.BOF Then rsIled.MoveFirst
gd.Rows = 2
gd.Clear
gd.FormatString = "Godown               | Qty  |  code"
FlxGrd.Rows = 2
FlxGrd.Clear
FlxGrd.FormatString = "SlNo. |                  Product Name              | Quantity     | Rate |  Amount | Code           "
txtDate = datecon(rsIled!fvrdate)
cboTag.ListIndex = rsIled!ftag

CboBranch.ListIndex = GetCboDataIndex(CboBranch, rsIled!fbranch)
cboBranch2.ListIndex = GetCboDataIndex(cboBranch2, rsIled!FBRANCHT)



txtRmk = rsIled!fnote
nR = 1
                            nR1 = 1
    
    Do While Not rsIled.EOF
            FlxGrd.TextMatrix(nR, 0) = nR
            FlxGrd.TextMatrix(nR, 1) = rsIled!facname
            FlxGrd.TextMatrix(nR, 2) = rsIled!fqty
            FlxGrd.TextMatrix(nR, 3) = rsIled!frate
            FlxGrd.TextMatrix(nR, 4) = rsIled!fval
            FlxGrd.TextMatrix(nR, 5) = rsIled!faccode
            
                    If lSuGd Then
                        rsGdTrn.Open "select * from gdtrans where val(fvrtype)='" & nVt & "' and faccode= '" & rsIled!faccode & "' and val(fvrno)='" & Val(txtVrno) & "'", Con, adOpenStatic
                        If Not rsGdTrn.EOF Then
                            If Not rsGdTrn.BOF Then rsGdTrn.MoveFirst
                            Do While Not rsGdTrn.EOF
                                gd.TextMatrix(nR1, 0) = rsGdTrn!fgodown
                                gd.TextMatrix(nR1, 1) = rsGdTrn!fqty
                                gd.TextMatrix(nR1, 2) = rsGdTrn!faccode
                                gd.AddItem ""
                                nR1 = nR1 + 1
                            rsGdTrn.MoveNext
                            Loop
                        End If
                        rsGdTrn.Close
                    End If
            
            
            FlxGrd.AddItem ""
            nR = nR + 1
            rsIled.MoveNext
    Loop
Else
fillData = False
End If
rsIled.Close
End Function


Private Sub FillGdwn()
Dim nR As Integer

Set rsGodown = New ADODB.Recordset
rsGodown.Open "select * from godown", Con, adOpenStatic
If Not rsGodown.EOF Then
GdwnGrid.Rows = 2
GdwnGrid.Clear
GdwnGrid.FormatString = "Godown               | Qty  |GC"
If Not rsGodown.BOF Then rsGodown.MoveFirst
nR = 1

Do While Not rsGodown.EOF
GdwnGrid.TextMatrix(nR, 0) = rsGodown!fgodown
nR = nR + 1
GdwnGrid.AddItem ""
rsGodown.MoveNext
Loop
End If
rsGodown.Close

End Sub



Private Sub SaveData()
Dim nNum As Long, nVrtype As Single


If nOpt = 1 Then
        Set rsNum = New ADODB.Recordset
        rsNum.Open "select * from num0203d", Con, adOpenDynamic, adLockPessimistic
            If Not rsNum.EOF Then
                If cboVrType.ListIndex = 0 Then
                   rsNum!fireceipt = rsNum!fireceipt + 1
                   rsNum.Update
                   nNum = rsNum!fireceipt
                ElseIf cboVrType.ListIndex = 1 Then
                   rsNum!fiissue = rsNum!fiissue + 1
                   rsNum.Update
                   nNum = rsNum!fiissue
                End If
            End If
        rsNum.Close
        txtVrno = nNum
Else

    nNum = Val(txtVrno)
End If

Set rsIled = New ADODB.Recordset
Set rsItem = New ADODB.Recordset
Set rsBrStk = New ADODB.Recordset

rsIled.Open "select * from ile0203d ", Con, adOpenDynamic, adLockPessimistic
            If cboVrType.ListIndex = 0 Then
                nVrtype = VrType("GREC")
            ElseIf cboVrType.ListIndex = 1 Then
                nVrtype = VrType("GISS")
            End If

    For I = 1 To FlxGrd.Rows - 1
        If FlxGrd.TextMatrix(I, 5) <> "" Then
            rsIled.AddNew
                rsIled!fvrtype = nVrtype
                rsIled!fvrno = nNum
                rsIled!fvrdate = txtDate.FormattedText
                rsIled!faccode = FlxGrd.TextMatrix(I, 5)
                rsIled!fqty = Val(FlxGrd.TextMatrix(I, 2))
                rsIled!frate = Val(FlxGrd.TextMatrix(I, 3))
                rsIled!fval = Val(FlxGrd.TextMatrix(I, 4))
                rsIled!ftag = cboTag.ListIndex
                rsIled!fnote = txtRmk
                rsIled!fbranch = CboBranch.Text
                rsIled!FBRANCHT = cboBranch2.Text
            rsIled.Update
            
            If lSuBra Then
                     If cboTag.ListIndex = 5 Then ' goods transfer
                        rsBrStk.Open "update brstk set fbal=fbal- '" & Val(FlxGrd.TextMatrix(I, 2)) & "' where faccode='" & FlxGrd.TextMatrix(I, 5) & "' and fbranch='" & CboBranch.Text & "' ", Con, adOpenDynamic, adLockPessimistic
                        rsBrStk.Open "update brstk set fbal=fbal+ '" & Val(FlxGrd.TextMatrix(I, 2)) & "' where faccode='" & FlxGrd.TextMatrix(I, 5) & "' and fbranch='" & cboBranch2.Text & "' ", Con, adOpenDynamic, adLockPessimistic
                     Else
                        If cboVrType.ListIndex = 0 Then
                        rsBrStk.Open "update brstk set fbal=fbal+ '" & Val(FlxGrd.TextMatrix(I, 2)) & "' where faccode='" & FlxGrd.TextMatrix(I, 5) & "' and fbranch='" & CboBranch.Text & "' ", Con, adOpenDynamic, adLockPessimistic
                        ElseIf cboVrType.ListIndex = 1 Then
                        rsBrStk.Open "update brstk set fbal=fbal- '" & Val(FlxGrd.TextMatrix(I, 2)) & "' where faccode='" & FlxGrd.TextMatrix(I, 5) & "' and fbranch='" & CboBranch.Text & "' ", Con, adOpenDynamic, adLockPessimistic
                        End If
                     End If
            End If
                     
                     
                     
'Opening
'Return
'Production
'Damaged
'Scrab
'Transfer
'Purchase without Bill
'adjustment
                   
                     
                     
                     
           
            
            If Val(FlxGrd.TextMatrix(I, 2)) > 0 Then
                rsItem.Open "update ite0203d set fcp= '" & Val(FlxGrd.TextMatrix(I, 2)) & "' where faccode='" & FlxGrd.TextMatrix(I, 5) & "'", Con, adOpenDynamic, adLockPessimistic
            End If
            
'            * this line to be activated for main branch
            
           If cboTag.ListIndex <> 5 Then  ' this line to be deactivated for branches
                If cboVrType.ListIndex = 0 Then
                    AddStock FlxGrd.TextMatrix(I, 5), Val(FlxGrd.TextMatrix(I, 2))
                ElseIf cboVrType.ListIndex = 1 Then
                    MinusStock FlxGrd.TextMatrix(I, 5), Val(FlxGrd.TextMatrix(I, 2))
                End If
            End If ' this line to be deactivated for branches
                   If lSuGd Then
                        Set rsGdTrn = New ADODB.Recordset
                        rsGdTrn.Open "select * from gdtrans", Con, adOpenDynamic, adLockPessimistic
                        For j = 1 To gd.Rows - 1
                            If gd.TextMatrix(j, 2) = FlxGrd.TextMatrix(I, 5) And Val(gd.TextMatrix(j, 1)) > 0 Then
                              rsGdTrn.AddNew
                              rsGdTrn!fvrno = nNum
                              rsGdTrn!fgodown = gd.TextMatrix(j, 0)
                              rsGdTrn!faccode = FlxGrd.TextMatrix(I, 5)
                              rsGdTrn!fqty = Val(gd.TextMatrix(j, 1))
                              rsGdTrn!fvrtype = nVrtype
                              rsGdTrn.Update
                              If cboVrType.ListIndex = 0 Then
                              AddGdStk FlxGrd.TextMatrix(I, 5), Val(gd.TextMatrix(j, 1)), gd.TextMatrix(j, 0)
                              ElseIf cboVrType.ListIndex = 1 Then
                              MinusGdStk FlxGrd.TextMatrix(I, 5), Val(gd.TextMatrix(j, 1)), gd.TextMatrix(j, 0)
                              End If
                              
                            End If
                        Next
                        rsGdTrn.Close
                End If
        End If
    Next
rsIled.Close
End Sub



Private Sub cboBranch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cboTag.ListIndex = 5 Then
        cboBranch2.SetFocus
    Else
        txtRmk.SetFocus
    End If
End If
End Sub


Private Sub cboBranch2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtRmk.SetFocus
End If
End Sub


Private Sub cboTag_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lSuBra Then
        If cboTag.ListIndex = 5 Then
          ' lblBranch.Visible = True
           lblBranch2.Visible = True
         '  cboBranch.Visible = True
           cboBranch2.Visible = True
           CboBranch.SetFocus
        Else
           'lblBranch.Visible = False
           lblBranch2.Visible = False
          ' cboBranch.Visible = False
           cboBranch2.Visible = False
           CboBranch.SetFocus

        End If
    Else

        txtRmk.SetFocus
    End If
End If
End Sub


Private Sub cboVrType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Set rsNum = New ADODB.Recordset
        rsNum.Open "select * from num0203d", Con, adOpenDynamic, adLockPessimistic
            If Not rsNum.EOF Then
                If cboVrType.ListIndex = 0 Then
                   nNum = rsNum!fireceipt + 1
                ElseIf cboVrType.ListIndex = 1 Then
                   nNum = rsNum!fiissue + 1
                End If
            End If
        rsNum.Close
        txtVrno = nNum
        
    If nOpt = 1 Then
        cboTag.SetFocus
    ElseIf nOpt = 2 Then
        txtVrno.SetFocus
    ElseIf nOpt = 3 Then
        txtVrno.SetFocus
    End If
End If
End Sub


Private Sub FlxGrd_KeyPress(KeyAscii As Integer)
Dim lFnd As Boolean, nR As Integer

If KeyAscii <> 13 Then
            If FlxGrd.Col = 1 Or FlxGrd.Col = 0 Then
                If KeyAscii = 8 Then      ' for backspace
                    If (IsEmpty(FlxGrd.Text)) Then
                        No = FlxGrd.Row
                        If (No > 1) Then
                            FlxGrd.RemoveItem (FlxGrd.Row)
                            FlxGrd.Row = No - 1
                            FlxGrd.Col = 2
                        End If
                    Else
                        If FlxGrd.Text <> "" Then
                            Stg = FlxGrd.Text
                            Strg = Mid(Stg, 1, (Len(Stg) - 1))
                            FlxGrd.Text = Strg
                        Else
                            If vbYes = MsgBox("Do you want to remove this row", vbYesNo, "Alert") Then
                                If FlxGrd.Row = 1 Then
                                    FlxGrd.AddItem ""
                                    FlxGrd.RemoveItem (1)
                                    FlxGrd.Col = 1
                                Else
                                    n = FlxGrd.Row
                                    FlxGrd.RemoveItem (FlxGrd.Row)
                                    FlxGrd.Row = n - 1
                                    FlxGrd.Col = 6
                                End If
                            End If
                        End If
                    End If
                ElseIf KeyAscii = 32 Then
                
                    Frame4.Top = 240
                    Frame4.Left = 3285
                    Frame4.Visible = True
                    cboPaytype.SetFocus
                Else                             'for normal
                    
                    FlxGrd.TextMatrix(FlxGrd.Row, 1) = FlxGrd.TextMatrix(FlxGrd.Row, 1) + Chr(KeyAscii)
                    If nLst = 2 Then
                        txtItemLst.Text = ""
                        txtItemLst.Text = FlxGrd.TextMatrix(FlxGrd.Row, 1)
                        fremIte.Top = 600
                        fremIte.Left = 2500
                        fremIte.Visible = True
                        txtItemLst.SetFocus
                        SendKeys "{END}"
                        Find ProductList, UCase(txtItemLst.Text), 0
                        
                    End If
                End If
           ElseIf FlxGrd.Col = 2 Then
                If KeyAscii = 8 Then      ' for backspace
                    If (IsEmpty(FlxGrd.Text)) Then
                        No = FlxGrd.Row
                        If (No > 1) Then
                            FlxGrd.RemoveItem (FlxGrd.Row)
                            FlxGrd.Row = No - 1
                            FlxGrd.Col = 2
                        End If
                    Else
                        If FlxGrd.Text <> "" Then
                            Stg = FlxGrd.Text
                            Strg = Mid(Stg, 1, (Len(Stg) - 1))
                            FlxGrd.Text = Strg
                        End If
                    End If
                   
                 Else
                   FlxGrd.TextMatrix(FlxGrd.Row, 2) = FlxGrd.TextMatrix(FlxGrd.Row, 2) + Chr(KeyAscii)
                 End If
           ElseIf FlxGrd.Col = 3 Then
                    If KeyAscii = 8 Then      ' for backspace
                    If (IsEmpty(FlxGrd.Text)) Then
                        No = FlxGrd.Row
                        If (No > 1) Then
                            FlxGrd.RemoveItem (FlxGrd.Row)
                            FlxGrd.Row = No - 1
                            FlxGrd.Col = 2
                        End If
                    Else
                        If FlxGrd.Text <> "" Then
                            Stg = FlxGrd.Text
                            Strg = Mid(Stg, 1, (Len(Stg) - 1))
                            FlxGrd.Text = Strg
                        End If
                    End If
                 Else
                     FlxGrd.TextMatrix(FlxGrd.Row, 3) = FlxGrd.TextMatrix(FlxGrd.Row, 3) + Chr(KeyAscii)
                 End If
           End If
ElseIf KeyAscii = 13 Then
      If FlxGrd.TextMatrix(FlxGrd.Row, 0) = "" Then
          If MsgBox("Save Data", vbYesNo + vbDefaultButton1) = vbYes Then
            If nOpt = 1 Then
             SaveData
            ElseIf nOpt = 2 Then
            If DeleteData Then SaveData
            ElseIf nOpt = 3 Then
             DeleteData
            End If
            ClearData
          Else
             Exit Sub
          End If
      End If
     If FlxGrd.Col = 1 Then
        If nLst = 0 Then
           rsItem.Open "select * from ite0203d where faccode='" & FlxGrd.TextMatrix(FlxGrd.Row, 1) & "'", Con, adOpenStatic
           If Not rsItem.EOF Then
            If Not rsItem.BOF Then rsItem.MoveFirst
              FlxGrd.TextMatrix(FlxGrd.Row, 1) = rsItem!facname
              FlxGrd.TextMatrix(FlxGrd.Row, 5) = rsItem!faccode
              FlxGrd.Col = 6
              FlxGrd.CellFontName = "Mylaiplain"
              FlxGrd.TextMatrix(FlxGrd.Row, 6) = rsItem!facname
            Else
              MsgBox "Code Not Found"
            End If
           rsItem.Close
        End If
        txtStock = GetStock(FlxGrd.TextMatrix(FlxGrd.Row, 5))
        FlxGrd.Col = 2
     
     ElseIf FlxGrd.Col = 2 Then
       FlxGrd.TextMatrix(FlxGrd.Row, 4) = Val(FlxGrd.TextMatrix(FlxGrd.Row, 2)) * Val(FlxGrd.TextMatrix(FlxGrd.Row, 3))
       If lSuGd Then  ' Godown for this company yes no
           lFnd = False
           For I = 1 To gd.Rows - 1
                If gd.TextMatrix(I, 2) = FlxGrd.TextMatrix(FlxGrd.Row, 5) Then
                  lFnd = True
                  Exit For
                End If
           Next
           For I = 1 To GdwnGrid.Rows - 1
               GdwnGrid.TextMatrix(I, 1) = ""
           Next
           If lFnd Then
                For I = 1 To gd.Rows - 1
                    If FlxGrd.TextMatrix(FlxGrd.Row, 5) = gd.TextMatrix(I, 2) Then
                        For j = 1 To GdwnGrid.Rows - 1
                              If GdwnGrid.TextMatrix(j, 0) = gd.TextMatrix(I, 0) Then
                                 GdwnGrid.TextMatrix(j, 1) = gd.TextMatrix(I, 1)
                              End If
                        Next
                    End If
                Next
           End If
           fremGdwn.Visible = True
           txtGtqty = FlxGrd.TextMatrix(FlxGrd.Row, 2)
           Text1 = FlxGrd.TextMatrix(FlxGrd.Row, 2)
           GdwnGrid.SetFocus
           GdwnGrid.Col = 1
       Else
           FlxGrd.Col = 3
       End If
     ElseIf FlxGrd.Col = 3 Then
            FlxGrd.TextMatrix(FlxGrd.Row, 4) = Val(FlxGrd.TextMatrix(FlxGrd.Row, 2)) * Val(FlxGrd.TextMatrix(FlxGrd.Row, 3))
            FlxGrd.AddItem ""
            FlxGrd.Row = FlxGrd.Row + 1
            FlxGrd.Col = 0
     End If
End If

End Sub


Private Sub Form_Load()
nOpt = 1

txtDate.Text = datecon(Date)
FillItem ProductList, -1, "Product  Name                                  | Code   |  Stock       | MRP "

StartSet
FillGdwn
If lSuBra Then
   ' lblBranch.Visible = True
   ' cboBranch.Visible = True
    lblBranch2.Visible = True
    cboBranch2.Visible = True
    
    BranchLoad CboBranch
    BranchLoad cboBranch2
    CboBranch.ListIndex = 1
    CboBranch.Locked = False
Else
  '  lblBranch.Visible = False
  '  cboBranch.Visible = False
    lblBranch2.Visible = False
    cboBranch2.Visible = False
End If

cboVrType.ListIndex = 0
cboTag.ListIndex = 0
cboTag.Locked = False
End Sub


Private Sub fillGrdflst()
FlxGrd.TextMatrix(FlxGrd.Row, 0) = FlxGrd.Row
FlxGrd.TextMatrix(FlxGrd.Row, 1) = ProductList.TextMatrix(ProductList.Row, 0)
FlxGrd.TextMatrix(FlxGrd.Row, 3) = ShowRate(ProductList.TextMatrix(ProductList.Row, 1))

FlxGrd.TextMatrix(FlxGrd.Row, 5) = ProductList.TextMatrix(ProductList.Row, 1)
        txtStock = GetStock(ProductList.TextMatrix(ProductList.Row, 1))

fremIte.Visible = False
FlxGrd.Col = 2
FlxGrd.SetFocus
End Sub

Private Sub GdwnGrid_KeyPress(KeyAscii As Integer)
Dim nGr As Integer
If KeyAscii <> 13 Then
            If KeyAscii = 27 Then
            fremGdwn.Visible = False
            FlxGrd.SetFocus
            End If
            If GdwnGrid.Col = 1 Then
                If KeyAscii = 8 Then      ' for backspace
                    If (IsEmpty(GdwnGrid.Text)) Then
                        No = GdwnGrid.Row
                        If (No > 1) Then
                            GdwnGrid.RemoveItem (GdwnGrid.Row)
                            GdwnGrid.Row = No - 1
                            GdwnGrid.Col = 2
                        End If
                    Else
                        If GdwnGrid.Text <> "" Then
                            Stg = GdwnGrid.Text
                            Strg = Mid(Stg, 1, (Len(Stg) - 1))
                            GdwnGrid.Text = Strg
                        Else
                            If vbYes = MsgBox("Do you want to remove this row", vbYesNo, "Alert") Then
                                If GdwnGrid.Row = 1 Then
                                    GdwnGrid.AddItem ""
                                    GdwnGrid.RemoveItem (1)
                                    GdwnGrid.Col = 1
                                Else
                                    n = GdwnGrid.Row
                                    GdwnGrid.RemoveItem (GdwnGrid.Row)
                                    GdwnGrid.Row = n - 1
                                    GdwnGrid.Col = 6
                                End If
                            End If
                        End If
                    End If
                ElseIf KeyAscii = 32 Then
                
                    Frame4.Top = 240
                    Frame4.Left = 3285
                    Frame4.Visible = True
                    cboPaytype.SetFocus
                Else                             'for normal
                    
                    GdwnGrid.TextMatrix(GdwnGrid.Row, 1) = GdwnGrid.TextMatrix(GdwnGrid.Row, 1) + Chr(KeyAscii)
                End If
            

            End If
            
ElseIf KeyAscii = 13 Then
            Text1 = ""
            For I = 1 To GdwnGrid.Rows - 1
                If GdwnGrid.TextMatrix(I, 0) <> "" Then
                   Text1 = Val(Text1) + Val(GdwnGrid.TextMatrix(I, 1))
                End If
            Next
            If Val(Text1) <= Val(txtGtqty) Then
               If Val(Text1) = Val(txtGtqty) Then
                  nGr = GdwnGrid.Rows - 1
                    
                    For I = 1 To nGr
                     For j = 1 To gd.Rows - 1
                            If gd.TextMatrix(j, 2) = FlxGrd.TextMatrix(FlxGrd.Row, 5) And gd.TextMatrix(j, 0) = GdwnGrid.TextMatrix(I, 0) Then
                                          
                                   gd.RemoveItem (j)
                                   Exit For
                            End If
                     Next
                    Next
                    
                    For I = 1 To GdwnGrid.Rows - 1
                        If GdwnGrid.TextMatrix(I, 1) <> "" Then
                           gd.AddItem ""
                           gd.TextMatrix(gd.Rows - 1, 0) = GdwnGrid.TextMatrix(I, 0)
                           gd.TextMatrix(gd.Rows - 1, 1) = GdwnGrid.TextMatrix(I, 1)
                           gd.TextMatrix(gd.Rows - 1, 2) = FlxGrd.TextMatrix(FlxGrd.Row, 5)
                        End If
                    Next
                
                  fremGdwn.Visible = False
                  FlxGrd.SetFocus
                  FlxGrd.Col = 3
               Else
                  GdwnGrid.Row = GdwnGrid.Row + 1
               End If
            Else
            End If

End If
End Sub


Private Sub mnu10_Click()
Unload Me
End Sub


Private Sub mnu2_Click()
nOpt = 1
cboVrType.SetFocus
End Sub

Private Sub mnu3_Click()
nOpt = 2
cboVrType.SetFocus
End Sub


Private Sub mnu4_Click()
nOpt = 3
End Sub


Private Sub mnu5_Click()
nOpt = 4
End Sub


Private Sub mnu6_Click()
nOpt = 5
End Sub


Private Sub mnu8_Click()
SaveData
End Sub

Private Sub productlist_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
fillGrdflst

ElseIf KeyAscii = 27 Then
fremIte.Visible = False
ElseIf KeyAscii = 13 Then


End If

End Sub


Private Sub txtItemLst_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
ProductList.SetFocus
ElseIf KeyCode = vbKeyUp Then
ProductList.SetFocus
End If

End Sub

Private Sub txtItemLst_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
fillGrdflst

ElseIf KeyAscii = 27 Then
fremIte.Visible = False
ElseIf KeyAscii = 13 Then


End If
End Sub

Private Sub txtItemLst_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 13 Then
   Find ProductList, UCase(txtItemLst.Text), 0
End If

End Sub


Private Sub txtRmk_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then FlxGrd.SetFocus

End Sub


Private Sub txtVrno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtVrno <> "" Then
If Not fillData() Then
    MsgBox "Not Found"
Else
    cboTag.SetFocus
End If
End If
End Sub


