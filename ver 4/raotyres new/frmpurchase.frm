VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPurchase 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Entry"
   ClientHeight    =   7815
   ClientLeft      =   2355
   ClientTop       =   3960
   ClientWidth     =   10635
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   10635
   WindowState     =   2  'Maximized
   Begin VB.Frame fremIte 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6315
      Left            =   10035
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   4845
      Begin VB.TextBox txtItemLst 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   30
         TabIndex        =   3
         Top             =   60
         Width           =   4770
      End
      Begin MSFlexGridLib.MSFlexGrid ItemLst 
         Height          =   5595
         Left            =   45
         TabIndex        =   2
         Top             =   600
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   9869
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
         FormatString    =   "Product  Name                                                           | Stock| MRP"
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grpgrid 
      Height          =   3030
      Left            =   1110
      TabIndex        =   6
      Top             =   4515
      Visible         =   0   'False
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   5345
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
   Begin VB.Frame fSp 
      Appearance      =   0  'Flat
      BackColor       =   &H006F472B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   7215
      TabIndex        =   4
      Top             =   5235
      Visible         =   0   'False
      Width           =   2490
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spelling Mistake"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   255
         TabIndex        =   5
         Top             =   825
         Width           =   1980
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   7335
      Left            =   75
      TabIndex        =   7
      Top             =   75
      Width           =   10440
      Begin VB.TextBox txtRound 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8685
         TabIndex        =   39
         Top             =   4800
         Width           =   1665
      End
      Begin VB.TextBox txtTaxR 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   8670
         TabIndex        =   38
         Top             =   4395
         Width           =   1665
      End
      Begin VB.TextBox txtTaxP 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7890
         TabIndex        =   37
         Top             =   4410
         Width           =   705
      End
      Begin VB.TextBox txtDedR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8670
         TabIndex        =   36
         Top             =   3540
         Width           =   1665
      End
      Begin VB.TextBox txtFreight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8670
         TabIndex        =   35
         Top             =   3990
         Width           =   1665
      End
      Begin VB.TextBox txtadd4 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1770
         TabIndex        =   32
         Top             =   3105
         Width           =   4935
      End
      Begin VB.TextBox txtAdd3 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1770
         TabIndex        =   31
         Top             =   2745
         Width           =   4935
      End
      Begin VB.TextBox txtAdd2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1770
         TabIndex        =   30
         Top             =   2385
         Width           =   4935
      End
      Begin VB.TextBox txtAdd1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1770
         TabIndex        =   29
         Top             =   2025
         Width           =   4935
      End
      Begin VB.TextBox txtPartyName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1770
         TabIndex        =   28
         Top             =   1650
         Width           =   4935
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   1440
         Left            =   4950
         TabIndex        =   21
         Top             =   105
         Width           =   5280
         Begin VB.TextBox txtStock 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   390
            Width           =   1380
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1275
            Left            =   1545
            TabIndex        =   23
            Text            =   "0.00"
            Top             =   135
            Width           =   3735
         End
         Begin VB.TextBox txtBstk 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   990
            Width           =   1380
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stock In hand"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   135
            Width           =   1005
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Branch Stock"
            Height          =   195
            Left            =   135
            TabIndex        =   25
            Top             =   735
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   1440
         Left            =   1770
         TabIndex        =   16
         Top             =   120
         Width           =   1380
         Begin VB.TextBox txtBno 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   75
            TabIndex        =   17
            Top             =   420
            Width           =   1170
         End
         Begin MSMask.MaskEdBox dBilldt 
            Height          =   285
            Left            =   90
            TabIndex        =   18
            Top             =   1065
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   8
            Format          =   "dd-mmm-yy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voucher Date"
            Height          =   195
            Left            =   75
            TabIndex        =   20
            Top             =   780
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voucher No"
            Height          =   195
            Left            =   60
            TabIndex        =   19
            Top             =   150
            Width           =   855
         End
      End
      Begin VB.Frame frmOno 
         BackColor       =   &H00FFFFFF&
         Height          =   1440
         Left            =   3270
         TabIndex        =   10
         Top             =   105
         Width           =   1665
         Begin VB.TextBox txtOno 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   75
            TabIndex        =   11
            Top             =   420
            Width           =   1170
         End
         Begin MSMask.MaskEdBox txtODt 
            Height          =   285
            Left            =   90
            TabIndex        =   12
            Top             =   1065
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   8
            Format          =   "dd-mmm-yy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bill Date"
            Height          =   195
            Left            =   75
            TabIndex        =   14
            Top             =   780
            Width           =   585
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bill No."
            Height          =   195
            Left            =   60
            TabIndex        =   13
            Top             =   150
            Width           =   495
         End
      End
      Begin VB.TextBox txtSTot 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4830
         TabIndex        =   9
         Top             =   6915
         Width           =   1170
      End
      Begin VB.ComboBox cboBranch 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   645
         Width           =   1440
      End
      Begin MSFlexGridLib.MSFlexGrid flxgrd 
         Height          =   3300
         Left            =   195
         TabIndex        =   15
         Top             =   3555
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   5821
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   -2147483648
         BackColorBkg    =   16777215
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         Appearance      =   0
         FormatString    =   "SlNo. |                  Product Name              | Quantity   |  Rate     |   Amount   | Code  | Tax"
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Round Value"
         Height          =   195
         Left            =   7080
         TabIndex        =   43
         Top             =   4860
         Width           =   930
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax"
         Height          =   195
         Left            =   7095
         TabIndex        =   42
         Top             =   4410
         Width           =   270
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deduction"
         Height          =   195
         Left            =   7110
         TabIndex        =   41
         Top             =   3585
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Freight Chgs"
         Height          =   195
         Left            =   7095
         TabIndex        =   40
         Top             =   4065
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   690
         TabIndex        =   34
         Top             =   2055
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   675
         TabIndex        =   33
         Top             =   1650
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Left            =   135
         TabIndex        =   27
         Top             =   330
         Width           =   510
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   12015
      Top             =   5700
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   11910
      Top             =   2895
   End
   Begin MSComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   7485
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "User name"
            TextSave        =   "User name"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Time"
            TextSave        =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu1 
      Caption         =   "Files"
      Begin VB.Menu mnu2 
         Caption         =   "New Bill"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu3 
         Caption         =   "Edit Bill"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnu4 
         Caption         =   "Cancel Bill"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnu5 
         Caption         =   "Delete Bill"
      End
      Begin VB.Menu mnu6 
         Caption         =   "Print Bill"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnue6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu7 
         Caption         =   "Park a Bill"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnu8 
         Caption         =   "Save"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu9 
         Caption         =   "Save As"
         Shortcut        =   {F3}
         Visible         =   0   'False
      End
      Begin VB.Menu mnu10 
         Caption         =   "-"
      End
      Begin VB.Menu mnu11 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnu12 
      Caption         =   "Edit"
      Begin VB.Menu mnu13 
         Caption         =   "Delete Current Line"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnu14 
         Caption         =   "Clear This Line"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu15 
         Caption         =   "Search Party"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnu16 
      Caption         =   "Master"
      Begin VB.Menu mnu17 
         Caption         =   "Add New Customer"
      End
      Begin VB.Menu mnu18 
         Caption         =   "Add New Product"
      End
      Begin VB.Menu mnu19 
         Caption         =   "Change Rate for Product"
      End
   End
   Begin VB.Menu mnu20 
      Caption         =   "Accounts"
      Begin VB.Menu mnu21 
         Caption         =   "Receipt"
      End
      Begin VB.Menu mnu22 
         Caption         =   "Payment"
      End
   End
   Begin VB.Menu mnu23 
      Caption         =   "Inventory"
      Begin VB.Menu mnu24 
         Caption         =   "Goods Receipt"
      End
   End
   Begin VB.Menu mnu25 
      Caption         =   "Windows"
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lFnd As Boolean
Dim nOpt As Single, nTax(10) As Double
Dim nBcP As Double, nTxP As Double
Dim cCustCode As String


Private Sub CallTot()
Dim nSubTot As Double, nTax As Double, nTotal As Double, nRound As Double
Dim nDed As Double, nBc As Double, nPc As Double, nTrns As Double
For I = 1 To FlxGrd.Rows - 1
    nSubTot = nSubTot + Val(FlxGrd.TextMatrix(I, 4))
Next
txtSTot = nSubTot

nDed = Val(txtDedR)
nTrns = Val(txtFreight)
nRound = Val(txtRound)
nTax = ((nSubTot - nDed) + nTrns) * Val(txtTaxP) / 100
nTotal = ((nSubTot - nDed) + nTrns + nTax) + nRound

txtTotal = nTotal
txtTaxR = nTax
txtFreight = nTrns
txtDedR = nDed
txtAmtR = txtTotal

End Sub


Private Sub ClearData()
FlxGrd.Rows = 2
FlxGrd.Clear
FlxGrd.FormatString = "SlNo. |                  Product Name              | Quantity   |  Rate     |   Amount   | Code  "
If nOpt = 1 Then txtBno = ""
txtdNo = ""
txtOno = ""
txtTotal = ""
txtCardNo = ""
txtSTot = ""
txtPartyName = ""
txtAdd1 = ""
txtAdd2 = ""
txtAdd3 = ""
txtadd4 = ""
txtdedP = ""
txtDedR = ""
txtTaxP = ""
txtTaxR = ""
txtTrnsChg = ""
txtBcP = ""
txtBcR = ""
txtPcP = ""
txtPcR = ""
txtAmtR = ""
txtAmtB = ""
txtTaxP = nTxP
End Sub

Private Function DelData()
   Set rsPur = New ADODB.Recordset
Set rsIled = New ADODB.Recordset
Set rsPlst = New ADODB.Recordset
Set rsItem = New ADODB.Recordset
Set rsGdTrn = New ADODB.Recordset
Set rsAcc = New ADODB.Recordset
Set rsBrStk = New ADODB.Recordset
    
    rsPlst.Open "delete from purlst where val(fvrno)='" & Val(txtBno) & "'", Con, adOpenDynamic, adLockPessimistic

    rsPur.Open "Select * from pur0203d where val(fvrno)='" & Val(txtBno) & "'", Con, adOpenStatic
    If Not rsPur.EOF Then
    If Not rsPur.BOF Then rsPur.MoveFirst
    Do While Not rsPur.EOF
    
        MinusStock rsPur!faccode, rsPur!fqty
        rsBrStk.Open "update brstk set fbal=fbal- '" & rsPur!fqty & "' where faccode='" & rsPur!faccode & "' and fbranch='" & rsPur!fbr & "' ", Con, adOpenDynamic, adLockPessimistic

    
    rsPur.MoveNext
    Loop
    End If
    rsPur.Close

'    If lSuGd Then
'        rsGdTrn.Open "select * from gdtrans where fvrno='" & txtBno & "' and fvrtype=1", Con, adOpenStatic
'        If Not rsGdTrn.EOF Then
'        If Not rsGdTrn.BOF Then rsGdTrn.MoveFirst
'        Do While Not rsGdTrn.EOF
'            AddGdStk rsGdTrn!faccode, rsGdTrn!fqty, rsGdTrn!fgodown
'        rsGdTrn.MoveNext
'        Loop
'
'        End If
'        rsGdTrn.Close
'            rsGdTrn.Open "delete from gdtrans where fvrno='" & txtBno & "' and fvrtype=1", Con, adOpenDynamic, adLockPessimistic
'
'    End If
    
    
    rsPur.Open "delete from pur0203d where val(fvrno)='" & Val(txtBno) & "'", Con, adOpenDynamic, adLockPessimistic
    
    rsIled.Open "delete from ile0203d where fvrno='" & txtBno & "' and fvrtype=1 and fbranch='" & CboBranch.Text & "'", Con, adOpenDynamic, adLockPessimistic
    

    

End Function



Private Sub FillAdd(cCode As String)
Set rsAdd = New ADODB.Recordset
txtAdd1 = ""
txtAdd2 = ""
txtAdd3 = ""
txtadd4 = ""
rsAdd.Open "select * from add0203d where faccode='" & cCode & " '", Con, adOpenStatic
If Not rsAdd.EOF Then
If Not rsAdd.BOF Then rsAdd.MoveFirst
If Not IsNull(rsAdd!add1) Then txtAdd1 = rsAdd!add1
If Not IsNull(rsAdd!add2) Then txtAdd2 = rsAdd!add2
If Not IsNull(rsAdd!add3) Then txtAdd3 = rsAdd!add3
If Not IsNull(rsAdd!add4) Then txtadd4 = rsAdd!add4


End If
rsAdd.Close

Set rsAdd = Nothing

End Sub

Private Sub FillGdStk(cCode As String)
Dim nR As Integer
GdwnGrid.Rows = 2
GdwnGrid.Clear
GdwnGrid.FormatString = "Godown               | Qty  |GC"

Set rsGdSTk = New ADODB.Recordset
rsGdSTk.Open "select * from gdstk where faccode='" & cCode & "' and fbal>0", Con, adOpenStatic
If Not rsGdSTk.EOF Then
If Not rsGdSTk.BOF Then rsGdSTk.MoveFirst
nR = 1

Do While Not rsGdSTk.EOF
GdwnGrid.TextMatrix(nR, 0) = rsGdSTk!fgd
GdwnGrid.TextMatrix(nR, 1) = rsGdSTk!fbal
nR = nR + 1
GdwnGrid.AddItem ""
rsGdSTk.MoveNext
Loop

End If
rsGdSTk.Close

End Sub

Private Sub fillGrdflst()
Set rsItem = New ADODB.Recordset
Set rsBrStk = New ADODB.Recordset
rsItem.Open "select * from ite0203d where faccode='" & ItemLst.TextMatrix(ItemLst.Row, 3) & "'", Con, adOpenStatic
If Not rsItem.EOF Then

    FlxGrd.TextMatrix(FlxGrd.Row, 0) = FlxGrd.Row
    FlxGrd.TextMatrix(FlxGrd.Row, 1) = rsItem!facname
    FlxGrd.TextMatrix(FlxGrd.Row, 3) = rsItem!fSp
    FlxGrd.TextMatrix(FlxGrd.Row, 5) = ItemLst.TextMatrix(ItemLst.Row, 3)
    txtStock = rsItem!fclbal
'    flxgrd.TextMatrix(flxgrd.Row, 7) = IIf(IsNull(rsItem!funit), "", rsItem!funit)
 '   flxgrd.TextMatrix(flxgrd.Row, 9) = rsItem!fweight
    
    If lSuBra Then
     rsBrStk.Open "select * from brstk where faccode='" & ItemLst.TextMatrix(ItemLst.Row, 3) & "' and fbranch='" & CboBranch.Text & "'", Con, adOpenStatic
     If Not rsBrStk.EOF Then
     If Not rsBrStk.BOF Then rsBrStk.MoveFirst
      txtBstk.Text = rsBrStk!fbal
     End If
     rsBrStk.Close
    End If
    If lSuTax Then
        If lSuItax Then
        FlxGrd.TextMatrix(FlxGrd.Row, 6) = IIf(IsNull(rsItem!ftax), "", rsItem!ftax)
        End If
    End If
    fremIte.Visible = False
    FlxGrd.Col = 2
    FlxGrd.SetFocus
End If
rsItem.Close
Set rsItem = Nothing
End Sub

Private Sub LoadProduct()
Dim nR As Integer
Set rsItem = New ADODB.Recordset
ItemLst.Rows = 2
ItemLst.Clear
ItemLst.FormatString = "Product  Name                                                           | Stock| MRP"
nR = 1
rsItem.Open "select * from ite0203d where faclevel<0 order by facname", Con, adOpenStatic
If Not rsItem.EOF Then
If Not rsItem.BOF Then rsItem.MoveFirst
Do While Not rsItem.EOF
ItemLst.TextMatrix(nR, 0) = rsItem!facname
ItemLst.TextMatrix(nR, 1) = rsItem!fclbal
ItemLst.TextMatrix(nR, 2) = rsItem!fmrp
ItemLst.TextMatrix(nR, 3) = rsItem!faccode
ItemLst.AddItem ""
nR = nR + 1
rsItem.MoveNext
Loop
End If
rsItem.Close
Set rsItem = Nothing
End Sub

Private Function SaveData() As Boolean
Dim nStg As Single, nVrtype As Single, cCode As String
nVrtype = VrType("PUR")
Set rsPur = New ADODB.Recordset
Set rsIled = New ADODB.Recordset
Set rsNum = New ADODB.Recordset
Set rsPlst = New ADODB.Recordset
Set rsItem = New ADODB.Recordset
Set rsGdTrn = New ADODB.Recordset
Set rsAcc = New ADODB.Recordset
Set rsArea = New ADODB.Recordset
Set rsBrStk = New ADODB.Recordset

'############################################################################################3
'############################################################################################3
'############################################################################################3
'############################################################################################3
'############################################################################################3




'############################################################################################3
'############################################################################################3
'############################################################################################3




    
    nStg = 0
    If nOpt = 1 Then
            rsNum.Open "select * from num0203d ", Con, adOpenDynamic, adLockPessimistic
             txtBno = rsNum!fpurchase + 1
             rsNum!fpurchase = Val(txtBno)
            rsNum.Update
            rsNum.Close
    End If
    
    Set rsPlst = New ADODB.Recordset
    rsPlst.Open "select * from purlst", Con, adOpenDynamic, adLockPessimistic
            
            rsPlst.AddNew
                   rsPlst!fvrno = txtBno
                   rsPlst!fvrdate = dBilldt.FormattedText
                   rsPlst!frefno = txtOno
                   rsPlst!frefdt = txtODt.FormattedText
                   rsPlst!fcucode = cCustCode
                   rsPlst!fpartyname = txtPartyName.Text
                    rsPlst!fadd1 = txtAdd1
                    rsPlst!fadd2 = txtAdd2
                    rsPlst!fadd3 = txtAdd3
                    rsPlst!fadd4 = txtadd4
                    rsPlst!ftotal = Val(txtTotal)
                    rsPlst!ftaxr = Val(txtTaxR)
                    rsPlst!ftaxp = Val(txtTaxP)
                    rsPlst!ftfri = Val(txtFreight)
                    rsPlst!ftdis = Val(txtDedR)
                    rsPlst!fbr = CboBranch.Text
                    rsPlst!frounded = Val(txtRound)
                    rsPlst!fbalance = Val(txtTotal)
            rsPlst.Update
        nStg = 1
    rsPlst.Close
    
    Set rsPur = New ADODB.Recordset
    Set rsIled = New ADODB.Recordset
   ' Set rsGdTrn = New ADODB.Recordset
    rsPur.Open "Select * from pur0203d", Con, adOpenDynamic, adLockPessimistic
    rsIled.Open "select * from ile0203d", Con, adOpenDynamic, adLockPessimistic
    'rsGdTrn.Open "select * from gdtrans", Con, adOpenDynamic, adLockPessimistic
    
    For I = 1 To FlxGrd.Rows - 1
        If FlxGrd.TextMatrix(I, 5) <> "" Then
            '/update bill/
            rsPur.AddNew
                rsPur!fvrno = txtBno
                rsPur!fvrdate = dBilldt.FormattedText
                rsPur!frefno = txtOno.Text
                rsPur!frefdt = txtODt.FormattedText
                rsPur!fcucode = cCustCode
                rsPur!faccode = FlxGrd.TextMatrix(I, 5)
                rsPur!fqty = Val(FlxGrd.TextMatrix(I, 2))
                rsPur!frate = Val(FlxGrd.TextMatrix(I, 3))
                rsPur!famount = Val(FlxGrd.TextMatrix(I, 4))
                rsPur!facname = FlxGrd.TextMatrix(I, 1)
                rsPur!fbr = CboBranch.Text
                
                rsBrStk.Open "update brstk set fbal=fbal+ '" & Val(FlxGrd.TextMatrix(I, 2)) & "' where faccode='" & FlxGrd.TextMatrix(I, 5) & "' and fbranch='" & CboBranch.Text & "' ", Con, adOpenDynamic, adLockPessimistic
                
            rsPur.Update
            
            '/Update Iledger/
            
            rsIled.AddNew
            
                rsIled!fvrtype = VrType("PUR")
                rsIled!fvrno = txtBno
                rsIled!fvrdate = dBilldt.FormattedText
                rsIled!faccode = FlxGrd.TextMatrix(I, 5)
                rsIled!fqty = Val(FlxGrd.TextMatrix(I, 2))
                rsIled!frate = Val(FlxGrd.TextMatrix(I, 3))
                rsIled!fval = Val(FlxGrd.TextMatrix(I, 4))
                rsIled!ftag = 11
                rsIled!fbranch = CboBranch.Text
            rsIled.Update
            '/Update Godown transaction/
                AddStock FlxGrd.TextMatrix(I, 5), Val(FlxGrd.TextMatrix(I, 2))
        End If
    Next
    nStg = 2
    Set rsLed = New ADODB.Recordset
    rsLed.Open "select * from led0203d", Con, adOpenDynamic, adLockPessimistic
                rsLed.AddNew
                rsLed!fvrtype = nVrtype
                rsLed!fvrno = Val(txtBno)
                rsLed!fvrdate = dBilldt.FormattedText
                   rsLed!faccode = cCustCode
                rsLed!fcrdb = "CR"
                rsLed!famount = Val(txtTotal)
                rsLed!faccode2 = cPurchase
                
                rsLed.Update
rsLed.Close

Set rsBrStk = Nothing
End Function



Private Sub StartAccessLmt()

End Sub


Private Sub stuffData()
Dim nR As Integer
Set rsPur = New ADODB.Recordset
Set rsIled = New ADODB.Recordset
nR = 1
ClearData
rsPur.Open "select * from purlist where val(fvrno)='" & Val(txtBno) & "'", Con, adOpenDynamic, adLockPessimistic


If Not rsPur.EOF Then
    If Not rsPur.BOF Then rsPur.MoveFirst
    
       If Not IsNull(rsPur!partyname) Then txtPartyName = rsPur!partyname
    
    If Not IsNull(rsPur!fadd1) Then txtAdd1 = rsPur!fadd1
    If Not IsNull(rsPur!fadd2) Then txtAdd2 = rsPur!fadd2
    If Not IsNull(rsPur!fadd3) Then txtAdd3 = rsPur!fadd3
    If Not IsNull(rsPur!fadd4) Then txtadd4 = rsPur!fadd4
'    If Not IsNull(rsPur!fnote) Then txtNote = rsPur!fnote
    If Not IsNull(rsPur!fvrdate) Then dBilldt.Text = datecon(rsPur!fvrdate)
      If Not IsNull(rsPur!frefno) Then txtOno = rsPur!frefno
      If Not IsNull(rsPur!frefdt) Then txtODt.Text = datecon(rsPur!frefdt)
       If Not IsNull(rsPur!ftaxr) Then txtTaxR = rsPur!ftaxr
       If Not IsNull(rsPur!ftaxp) Then txtTaxP = rsPur!ftaxp
    
       If Not IsNull(rsPur!ftdis) Then txtDedR = rsPur!ftdis
       If Not IsNull(rsPur!ftfri) Then txtFreight = rsPur!ftfri
       If Not IsNull(rsPur!frounded) Then txtRound = rsPur!frounded
        Do While Not rsPur.EOF
           FlxGrd.TextMatrix(nR, 0) = nR
           FlxGrd.TextMatrix(nR, 1) = rsPur!facname
           FlxGrd.TextMatrix(nR, 2) = Format(rsPur!fqty, cDP)
           FlxGrd.TextMatrix(nR, 3) = Format(rsPur!frate, "#####0.00")
           FlxGrd.TextMatrix(nR, 4) = Format(rsPur!famount, "######0.00")
           FlxGrd.TextMatrix(nR, 5) = rsPur!faccode
           FlxGrd.AddItem ""
           nR = nR + 1
        rsPur.MoveNext
        Loop
End If
rsPur.Close
CallTot
Set rsPur = Nothing
Set rsIled = Nothing
End Sub

Private Sub cboBranch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'   Set rsBBnum = New ADODB.Recordset
'   rsBBnum.Open "select * from bbnum where fbranch='" & cboBranch.Text & "'", Con, adOpenStatic
'   If Not rsBBnum Then
'   If Not rsBBnum.BOF Then rsBBnum.MoveFirst
'      MsgBox rsBBnum!fbillno
'      txtBno.Text = rsBBnum!fbillno
      txtOno.SetFocus
'   End If
'   rsBBnum.Close
'   Set rsBBnum = Nothing
End If
End Sub




Private Sub dBilldt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtOno.SetFocus
End If
End Sub


Private Sub FlxGrd_KeyPress(KeyAscii As Integer)
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
                                    FlxGrd.Col = 0
                                End If
                            End If
                        End If
                    End If
                ElseIf KeyAscii = 32 Then
                
                Else                             'for normal
                    
                    FlxGrd.TextMatrix(FlxGrd.Row, 1) = FlxGrd.TextMatrix(FlxGrd.Row, 1) + Chr(KeyAscii)
                    If nLst = 2 Then
                        txtItemLst.Text = ""
                        txtItemLst.Text = FlxGrd.TextMatrix(FlxGrd.Row, 1)
                                            fremIte.Top = 600
                    fremIte.Left = 6480

                        fremIte.Visible = True
                        txtItemLst.SetFocus
                        SendKeys "{END}"
                        Find ItemLst, UCase(txtItemLst.Text), 0, frmBill.fSp
                        
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
                            FlxGrd.Col = 3
                        End If
                    Else
                        If FlxGrd.Text <> "" Then
                            Stg = FlxGrd.Text
                            Strg = Mid(Stg, 1, (Len(Stg) - 1))
                            FlxGrd.Text = Strg
                        End If
                    End If
                Else
                    If (Val(FlxGrd.TextMatrix(FlxGrd.Row, 3)) <= 0 And FlxGrd.TextMatrix(FlxGrd.Row, 3) <> ".") And Chr(KeyAscii) <> "." Then FlxGrd.TextMatrix(FlxGrd.Row, 3) = ""
                    FlxGrd.TextMatrix(FlxGrd.Row, 3) = FlxGrd.TextMatrix(FlxGrd.Row, 3) + Chr(KeyAscii)
                End If

           
           
           End If
ElseIf KeyAscii = 13 Then
        If FlxGrd.Col = 1 Or FlxGrd.Col = 0 Then
           If FlxGrd.TextMatrix(FlxGrd.Row, 1) = "" And FlxGrd.Row >= 2 Then
              txtDedR.SetFocus
              
           End If
        End If
        
     If FlxGrd.Col = 1 Then
            If nLst = 0 Then
            Set rsItem = New ADODB.Recordset
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
            Set rsItem = Nothing
            End If
              txtStock = GetStock(FlxGrd.TextMatrix(FlxGrd.Row, 5))
              FlxGrd.Col = 2
     ElseIf FlxGrd.Col = 2 Then
            FlxGrd.TextMatrix(FlxGrd.Row, 4) = Val(FlxGrd.TextMatrix(FlxGrd.Row, 2)) * Val(FlxGrd.TextMatrix(FlxGrd.Row, 3))
            If lSuGd Then
            FillGdStk FlxGrd.TextMatrix(FlxGrd.Row, 5)
             fremGdwn.Visible = True
             GdwnGrid.SetFocus
            End If
            FlxGrd.Col = 3
     ElseIf FlxGrd.Col = 3 Then
            FlxGrd.TextMatrix(FlxGrd.Row, 4) = Val(FlxGrd.TextMatrix(FlxGrd.Row, 2)) * Val(FlxGrd.TextMatrix(FlxGrd.Row, 3))
            FlxGrd.AddItem ""
            FlxGrd.Row = FlxGrd.Row + 1
            CallTot
            FlxGrd.Col = 0
     End If
End If


End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
fSp.Visible = False
End Sub

Private Sub Form_Load()
nOpt = 1

LoadProduct
StartAccessLmt
BranchLoad CboBranch
StartSet
FillAcc Grpgrid, -1, "                                                          Account Name                                | Code"

CboBranch.ListIndex = 0
dBilldt.Text = datecon(Date)
txtODt.Text = datecon(Date)
Set rsNum = New ADODB.Recordset
rsNum.Open "select * from num0203d ", Con, adOpenStatic
txtBno = rsNum!fpurchase + 1
rsNum.Close
Set rsNum = Nothing

End Sub




Private Sub GrpGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPartyName.Text = Grpgrid.Text
   FillAdd Grpgrid.TextMatrix(Grpgrid.Row, 1)

txtPartyName_KeyPress 13


End If
End Sub


Private Sub ItemLst_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
fillGrdflst
End If
End Sub



Private Sub mnu11_Click()
Unload Me
End Sub

Private Sub mnu13_Click()
If MsgBox("Remove The Line", vbYesNo + vbDefaultButton2) = vbYes Then
   FlxGrd.RemoveItem (FlxGrd.Row)
End If
End Sub

Private Sub mnu2_Click()
Set rsNum = New ADODB.Recordset
ClearData
nOpt = 1
rsNum.Open "select * from num0203d ", Con, adOpenStatic
txtBno = rsNum!fpurchase + 1
rsNum.Close

dBilldt.SetFocus
End Sub

Public Sub mnu3_Click()
ClearData
nOpt = 2
txtBno.SetFocus
End Sub

Private Sub mnu4_Click()
ClearData
nOpt = 3
txtBno.SetFocus
End Sub

Private Sub mnu5_Click()
ClearData
nOpt = 4
txtBno.SetFocus

End Sub


Private Sub mnu6_Click()
ClearData
nOpt = 5
txtBno.SetFocus

End Sub


Private Sub Timer1_Timer()
sb1.Panels(2).Text = Time
End Sub


Public Sub txtBno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtBno <> "" Then
    If nOpt > 1 Then
         If nOpt = 2 Then
           stuffData
            dBilldt.SetFocus
         ElseIf nOpt = 3 Then
            If MsgBox("cancel this bill", vbYesNo) = vbYes Then
                
            End If
         ElseIf nOpt = 4 Then
            If MsgBox("Delete this bill", vbYesNo) = vbYes Then
                
            End If
         ElseIf nOpt = 5 Then
            If MsgBox("Print this bill", vbYesNo) = vbYes Then
                
            End If
         End If
         
    End If
End If
End Sub













Private Sub txtDedR_GotFocus()
    SendKeys "{HOME}"
    SendKeys "+{END}"

End Sub

Private Sub txtDedR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CallTot
txtFreight.SetFocus
End If
End Sub


Private Sub txtFreight_GotFocus()
    SendKeys "{HOME}"
    SendKeys "+{END}"

End Sub

Private Sub txtFreight_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtTaxP.SetFocus
CallTot
End If
End Sub


Private Sub txtItemLst_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
ItemLst.SetFocus
ElseIf KeyCode = vbKeyUp Then
ItemLst.SetFocus
End If
End Sub

Private Sub txtItemLst_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And lFnd Then

fillGrdflst


ElseIf KeyAscii = 27 Then
    
    fremIte.Visible = False


End If
End Sub


Private Sub txtItemLst_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 13 Then
 lFnd = Find(ItemLst, UCase(txtItemLst.Text), 0, frmBill.fSp)
End If

End Sub



Private Sub txtODt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPartyName.SetFocus
End If
End Sub


Private Sub txtOno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtOno <> "" Then
txtODt.SetFocus
End If
End Sub


Private Sub txtPartyName_GotFocus()
    SendKeys "{HOME}"
    SendKeys "+{END}"
Grpgrid.Top = 3500
Grpgrid.Visible = True
End Sub

Private Sub txtPartyName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
   Grpgrid.SetFocus
ElseIf KeyCode = vbKeyUp Then
   Grpgrid.SetFocus
End If
End Sub


Private Sub txtPartyName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtPartyName <> "" And lFnd Then
   'Find GrpGrid, Format(txtPartyName.Text, ">"), 0  for change option we have check this
    txtPartyName.Text = Grpgrid.TextMatrix(Grpgrid.Row, 0)
    cCustCode = Grpgrid.TextMatrix(Grpgrid.Row, 1)
    cGroup = Grpgrid.TextMatrix(Grpgrid.Row, 2)
    FillAdd cCustCode
    Grpgrid.Visible = False
    FlxGrd.SetFocus
'    If lSuDed Then
'    txtdedP.SetFocus
'    ElseIf lSuTax Then
'    txtTaxP.SetFocus
'    ElseIf lSuBc Then
'    txtBcP.SetFocus
'    ElseIf lSuPc Then
'    txtPcP.SetFocus
'    End If

ElseIf KeyAscii = 13 And txtPartyName <> "j " And Not lFnd Then
    Grpgrid.Visible = False
    txtAdd1.SetFocus
ElseIf KeyAscii = 13 And txtPartyName = "" Then
    If lSuDed Then
    txtdedP.SetFocus
    ElseIf lSuTax Then
    txtTaxP.SetFocus
    ElseIf lSuBc Then
    txtBcP.SetFocus
    ElseIf lSuPc Then
    txtPcP.SetFocus
    End If

ElseIf KeyAscii = 27 Then
    Frame4.Visible = False
    FlxGrd.SetFocus
End If

If KeyAscii = 13 And Not lFnd Then
Grpgrid.Visible = False
End If
End Sub


Private Sub txtPartyName_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyEscape And KeyCode <> vbKeyRight And KeyCode <> vbKeyLeft _
            And txtPartyName.Text <> "" And KeyCode <> vbKeyReturn Then
   lFnd = Find(Grpgrid, Format(txtPartyName.Text, ">"), 0, frmBill.fSp)
   FillAdd Grpgrid.TextMatrix(Grpgrid.Row, 1)
End If

End Sub





Private Sub txtRound_GotFocus()
    SendKeys "{HOME}"
    SendKeys "+{END}"

End Sub

Private Sub txtRound_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CallTot
    If nOpt = 1 Then
        If MsgBox("Save", vbYesNo) = vbYes Then
            SaveData
        End If
    ElseIf nOpt = 2 Then
        If MsgBox("Update", vbYesNo) = vbYes Then
            DelData
            SaveData
        End If
    ElseIf nOpt = 3 Then
    
    End If
ClearData
nOpt = 1
rsNum.Open "select * from num0203d ", Con, adOpenStatic
txtBno = rsNum!fpurchase + 1
rsNum.Close

CboBranch.SetFocus

End If
End Sub


Private Sub txtTaxP_GotFocus()
    SendKeys "{HOME}"
    SendKeys "+{END}"

End Sub

Private Sub txtTaxP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CallTot
txtRound.SetFocus
ElseIf KeyAscii = 27 Then
    Frame4.Visible = False
    FlxGrd.SetFocus

End If

End Sub





