VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptPur 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Purchase Report"
   ClientHeight    =   7995
   ClientLeft      =   2160
   ClientTop       =   2055
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   11220
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   345
      Left            =   8925
      TabIndex        =   7
      Top             =   375
      Width           =   1020
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Height          =   345
      Left            =   7785
      TabIndex        =   6
      Top             =   390
      Width           =   1020
   End
   Begin VB.ComboBox CboPartyname 
      Height          =   315
      Left            =   210
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   345
      Width           =   3450
   End
   Begin MSFlexGridLib.MSFlexGrid GrdRpt 
      Height          =   6000
      Left            =   135
      TabIndex        =   0
      Top             =   1005
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   10583
      _Version        =   393216
      Cols            =   10
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   $"FrmRptPur.frx":0000
   End
   Begin MSMask.MaskEdBox mskFdt 
      Height          =   285
      Left            =   4500
      TabIndex        =   2
      Top             =   435
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      Format          =   "dd-mmm-yy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskTdt 
      Height          =   285
      Left            =   6255
      TabIndex        =   3
      Top             =   435
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "To"
      Height          =   195
      Left            =   5955
      TabIndex        =   5
      Top             =   435
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   195
      Left            =   4020
      TabIndex        =   4
      Top             =   435
      Width           =   345
   End
End
Attribute VB_Name = "frmRptPur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ShowPur(id As Long)
If id > 0 Then
    frmPurchase.Show
    Call frmPurchase.mnu3_Click
    frmPurchase.txtBno = id
    frmPurchase.txtBno_KeyPress (13)
End If
End Sub

Private Sub stuffData()
Dim nR As Integer, nTTax As Double, nTdis As Double, nTfri As Double, nTTotal As Double
GrdRpt.Rows = 2
GrdRpt.Clear
GrdRpt.FormatString = "Voucher    |      Date    |  Bill No    |       Bill Date  |   Supplier Name                      |     Amount     |    Dis       |     Freight   |      Tax      |         Total   "
Set rsPlst = New ADODB.Recordset
nR = 1
If CboPartyname.Text = "All Suppliers" Then
 '   rsPlst.Open "select * from  purlst where (cdate(fvrdate) between '" & CDate(mskFdt.FormattedText) & "' and '" & CDate(mskTdt.FormattedText) & "' order by fvrno", Con, adOpenStatic
     rsPlst.Open "select * from  purlst where (CDATE(fvrdate) between '" & DateValue(mskFdt.FormattedText) & "' and '" & DateValue(mskTdt.FormattedText) & "') order by fvrno", Con, adOpenStatic

Else
    rsPlst.Open "select * from  purlst where (cdate(fvrdate) between #" & CDate(mskFdt.FormattedText) & "# and #" & CDate(mskTdt.FormattedText) & "#) and faccode='" & CboPartyname.ItemData(CboPartyname.ListIndex) & "' order by fvrno", Con, adOpenStatic

End If

If Not rsPlst.EOF Then
If Not rsPlst.BOF Then rsPlst.MoveFirst
With GrdRpt
Do While Not rsPlst.EOF

.TextMatrix(nR, 0) = rsPlst!fvrno
.TextMatrix(nR, 1) = rsPlst!fvrdate
.TextMatrix(nR, 2) = rsPlst!frefno
.TextMatrix(nR, 3) = rsPlst!frefdt
If Not IsNull(rsPlst!fpartyname) Then .TextMatrix(nR, 4) = rsPlst!fpartyname
.TextMatrix(nR, 5) = Format((((rsPlst!ftotal + rsPlst!ftdis) - rsPlst!ftfri) - rsPlst!ftaxr), "#####0.00")
.TextMatrix(nR, 6) = Format(rsPlst!ftdis, "####0.00")
.TextMatrix(nR, 7) = Format(rsPlst!ftfri, "####0.00")
.TextMatrix(nR, 8) = Format(rsPlst!ftaxr, "####0.00")
.TextMatrix(nR, 9) = Format(rsPlst!ftotal, "######0.00")
nTTax = nTTax + rsPlst!ftaxr
nTdis = nTdis + rsPlst!ftdis
nTfri = nTfri + rsPlst!ftfri
nTTotal = nTTotal + rsPlst!ftotal
nR = nR + 1
.AddItem ""
rsPlst.MoveNext
Loop
.AddItem ""
nR = nR + 1
.TextMatrix(nR, 6) = String(20, "-")
.TextMatrix(nR, 7) = String(20, "-")
.TextMatrix(nR, 8) = String(20, "-")
.TextMatrix(nR, 9) = String(20, "-")

.AddItem ""
nR = nR + 1

.TextMatrix(nR, 6) = Format(nTdis, "####0.00")
.TextMatrix(nR, 7) = Format(nTfri, "####0.00")
.TextMatrix(nR, 8) = Format(nTTax, "#####0.00")
.TextMatrix(nR, 9) = Format(nTTotal, "######0.00")

.AddItem ""
nR = nR + 1
.TextMatrix(nR, 6) = String(20, "-")
.TextMatrix(nR, 7) = String(20, "-")
.TextMatrix(nR, 8) = String(20, "-")
.TextMatrix(nR, 9) = String(20, "-")

End With
End If
rsPlst.Close

End Sub

Private Sub Command1_Click()
stuffData

End Sub


Private Sub Command2_Click()
Dim nR As Integer
nR = GrdRpt.Rows - 1
Open "c:\files\testfile.TXT" For Output As #1  ' Open file for output
Print #1, Chr(15)
Print #1, String(122, "-")
Print #1, RPad("Bill No", 7) + " " + RPad("Bill Dt", 10) + " " + RPad("Name", 30) + " " + LPad("Amount", 13) + " " + LPad("Dis", 7) + " " + LPad("Freight", 7) + " " + LPad("Tax", 7) + " " + LPad("Total", 13)
Print #1, String(122, "-")

For I = 1 To nR

Print #1, LPad(GrdRpt.TextMatrix(I, 2), 7) + " " + RPad(GrdRpt.TextMatrix(I, 3), 10) + " " + RPad(GrdRpt.TextMatrix(I, 4), 30) + " " + LPad(GrdRpt.TextMatrix(I, 5), 13) + " " + LPad(GrdRpt.TextMatrix(I, 6), 7) + " " + LPad(GrdRpt.TextMatrix(I, 7), 7) + " "; LPad(GrdRpt.TextMatrix(I, 8), 7) + " " + LPad(GrdRpt.TextMatrix(I, 9), 13)

Next

Close #1



End Sub

Private Sub Form_Load()

FillRptCmbo CboPartyname
CboPartyname.ListIndex = 0
mskFdt.Text = datecon(Date)
mskTdt.Text = datecon(Date)
'StuffData
End Sub


Private Sub GrdRpt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Val(GrdRpt.TextMatrix(GrdRpt.Row, 0)) > 0 Then
    Call ShowPur(Val(GrdRpt.TextMatrix(GrdRpt.Row, 0)))
End If
End Sub


