VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRptSal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sales Report"
   ClientHeight    =   7710
   ClientLeft      =   300
   ClientTop       =   585
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   11340
   Begin VB.TextBox txtSearch 
      Height          =   390
      Left            =   7440
      TabIndex        =   8
      Top             =   315
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   345
      Left            =   6195
      TabIndex        =   7
      Top             =   360
      Width           =   1020
   End
   Begin MSComCtl2.DTPicker mskFdt 
      Height          =   345
      Left            =   2490
      TabIndex        =   5
      Top             =   375
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   609
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   40183
   End
   Begin VB.ComboBox CboBranch 
      Height          =   315
      Left            =   135
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   390
      Width           =   2265
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Height          =   345
      Left            =   5115
      TabIndex        =   0
      Top             =   360
      Width           =   1020
   End
   Begin MSFlexGridLib.MSFlexGrid GrdRpt 
      Height          =   6000
      Left            =   60
      TabIndex        =   1
      Top             =   840
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   10583
      _Version        =   393216
      Cols            =   10
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmRptSal.frx":0000
   End
   Begin MSComCtl2.DTPicker msktdt 
      Height          =   345
      Left            =   3810
      TabIndex        =   6
      Top             =   360
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   609
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   40183
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   195
      Left            =   2490
      TabIndex        =   3
      Top             =   165
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   195
      Left            =   3810
      TabIndex        =   2
      Top             =   150
      Width           =   195
   End
End
Attribute VB_Name = "frmRptSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub StuffSal()
Dim nR As Integer, nTTax As Double, nTdis As Double, nTfri As Double, nTTotal As Double, nTbc As Double
GrdRpt.Rows = 2
GrdRpt.Clear
GrdRpt.FormatString = "Bill No    |       Bill Date  |Pay Type    |            Name                      |       Amount     |    Dis       |            Tax   |     Freight   | Bank Chg |                Total   "
Set rsBLst = New ADODB.Recordset
nR = 1
'#" & DateValue(Dfd.Value) & "#
If lSuBra Then
   If txtSearch.Text <> "" Then
    rsBLst.Open "select * from  salrpt  where (cdate(fbilldt) between #" & DateValue(mskFdt.Value) & "# and #" & DateValue(msktdt.Value) & "# ) and fbranch='" & CboBranch.Text & "' AND   FACNAME like +'%'+'" & txtSearch.Text & "'+'%'", Con, adOpenStatic
   Else
    rsBLst.Open "select * from  salrpt  where (cdate(fbilldt) between #" & DateValue(mskFdt.Value) & "# and #" & DateValue(msktdt.Value) & "# ) and fbranch='" & CboBranch.Text & "'", Con, adOpenStatic
   End If
    
Else
    rsBLst.Open "select * from  salrpt  where (cdate(fbilldt) between '" & CDate(mskFdt.Value) & "' and '" & CDate(msktdt.Value) & "' )  order by val(fbillno)", Con, adOpenStatic

End If

If Not rsBLst.EOF Then
If Not rsBLst.BOF Then rsBLst.MoveFirst
With GrdRpt
Do While Not rsBLst.EOF

.TextMatrix(nR, 0) = rsBLst!Fbillno
.TextMatrix(nR, 1) = rsBLst!fbilldt
.TextMatrix(nR, 2) = rsBLst!paytype
If Not IsNull(rsBLst!facname) Then .TextMatrix(nR, 3) = rsBLst!facname
.TextMatrix(nR, 4) = Format(((rsBLst!ftotal + rsBLst!fded) - rsBLst!ftax) + rsBLst!ftrns, "#####0.00")
.TextMatrix(nR, 5) = Format(rsBLst!fded, "####0.00")
.TextMatrix(nR, 6) = Format(rsBLst!ftax, "####0.00")
.TextMatrix(nR, 7) = Format(rsBLst!ftrns, "####0.00")
.TextMatrix(nR, 8) = Format(rsBLst!fbc, "####0.00")
'.TextMatrix(nR, 8) = Format(rsBLst!ftotal, "######0.00")
.TextMatrix(nR, 9) = Format(rsBLst!ftotal, "######0.00")

'If rsBLst!fnote <> "" Then
'nR = nR + 1
'.AddItem ""
'.TextMatrix(nR, 3) = rsBLst!fnote
'End If
'
'If rsBLst!fcardno <> "" Then
'nR = nR + 1
'.AddItem ""
'.TextMatrix(nR, 3) = rsBLst!fcardno
'End If
'
'If rsBLst!fslno <> "" Then
'nR = nR + 1
'.AddItem ""
'.TextMatrix(nR, 3) = rsBLst!fslno
'End If
nTbc = nTbc + rsBLst!fbc
nTTax = nTTax + rsBLst!ftax
nTdis = nTdis + rsBLst!fded
nTfri = nTfri + rsBLst!ftrns
nTTotal = nTTotal + rsBLst!ftotal
.AddItem ""
nR = nR + 1
rsBLst.MoveNext
Loop

.AddItem ""
nR = nR + 1
.TextMatrix(nR, 5) = String(20, "-")

.TextMatrix(nR, 6) = String(20, "-")
.TextMatrix(nR, 7) = String(20, "-")
.TextMatrix(nR, 8) = String(20, "-")
.TextMatrix(nR, 9) = String(20, "-")

.AddItem ""
nR = nR + 1
.TextMatrix(nR, 5) = Format(nTdis, "####0.00")

.TextMatrix(nR, 6) = Format(nTTax, "#####0.00")

.TextMatrix(nR, 7) = Format(nTfri, "#####0.00")
.TextMatrix(nR, 8) = Format(nTbc, "#####0.00")
.TextMatrix(nR, 9) = Format(nTTotal, "######0.00")

.AddItem ""
nR = nR + 1
.TextMatrix(nR, 5) = String(20, "-")

.TextMatrix(nR, 6) = String(20, "-")
.TextMatrix(nR, 7) = String(20, "-")
.TextMatrix(nR, 8) = String(20, "-")
.TextMatrix(nR, 9) = String(20, "-")

End With
End If
rsBLst.Close

End Sub


Private Sub Command1_Click()
StuffSal
End Sub


Private Sub Command2_Click()
Dim nR As Integer
nR = GrdRpt.Rows - 1
Open "c:\files\testfile.TXT" For Output As #1  ' Open file for output
Print #1, Chr(15)
Print #1, String(122, "-")
Print #1, RPad("Bill No", 7) + " " + RPad("Bill Dt", 10) + " " + RPad("Pay Type", 8) + " " + RPad("Name", 30) + " " + LPad("Amount", 13) + " " + LPad("Dis", 7) + " " + LPad("Tax", 7) + " " + LPad("Freight", 7) + " " + LPad("Bank Chg", 8) + " " + LPad("Total", 13)
Print #1, String(122, "-")

For I = 1 To nR

Print #1, LPad(GrdRpt.TextMatrix(I, 0), 7) + " " + RPad(GrdRpt.TextMatrix(I, 1), 10) + " " + RPad(GrdRpt.TextMatrix(I, 2), 8) + " " + RPad(GrdRpt.TextMatrix(I, 3), 30) + " " + LPad(GrdRpt.TextMatrix(I, 4), 13) + " " + LPad(GrdRpt.TextMatrix(I, 5), 7) + " " + LPad(GrdRpt.TextMatrix(I, 6), 7) + " "; LPad(GrdRpt.TextMatrix(I, 7), 7) + " "; LPad(GrdRpt.TextMatrix(I, 8), 8) + " " + LPad(GrdRpt.TextMatrix(I, 9), 13)

Next

Close #1




'
'Do While True
'RetVal = Shell("c:\files\dosprint.bat", 0)
'If MsgBox("Print Again", vbOKCancel) = vbCancel Then
'Exit Do
'End If
'Loop



End Sub

Private Sub Form_Load()
mskFdt.Value = (Date)
msktdt.Value = (Date)
CboBranch.Visible = False
If lSuBra Then
CboBranch.Visible = True
BranchLoad CboBranch
CboBranch.ListIndex = 0
End If
End Sub


Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

   ' txtSearch.Text = "%" + txtSearch.Text + "%"
Command1.SetFocus
End If
End Sub



