VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStkStmnt 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Stock Statement"
   ClientHeight    =   7740
   ClientLeft      =   2505
   ClientTop       =   2130
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   9930
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7005
      Left            =   75
      TabIndex        =   3
      Top             =   255
      Visible         =   0   'False
      Width           =   9810
      Begin VB.TextBox txtRec 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7605
         TabIndex        =   6
         Top             =   6030
         Width           =   1050
      End
      Begin VB.TextBox txtIss 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8640
         TabIndex        =   5
         Top             =   6030
         Width           =   1020
      End
      Begin VB.TextBox txtSih 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7605
         TabIndex        =   4
         Top             =   6510
         Width           =   1050
      End
      Begin MSFlexGridLib.MSFlexGrid grdIstk 
         Height          =   5295
         Left            =   30
         TabIndex        =   7
         Top             =   585
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   9340
         _Version        =   393216
         Cols            =   9
         FocusRect       =   0
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   $"frmStkStmnt.frx":0000
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name :"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   75
         TabIndex        =   9
         Top             =   180
         Width           =   4395
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock In Hand"
         Height          =   165
         Left            =   6315
         TabIndex        =   8
         Top             =   6540
         Width           =   1155
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   345
      Left            =   1290
      TabIndex        =   2
      Top             =   420
      Width           =   1020
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Height          =   345
      Left            =   195
      TabIndex        =   0
      Top             =   420
      Width           =   1020
   End
   Begin MSFlexGridLib.MSFlexGrid GrdRpt 
      Height          =   6000
      Left            =   195
      TabIndex        =   1
      Top             =   1020
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   10583
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   $"frmStkStmnt.frx":00A2
   End
End
Attribute VB_Name = "frmStkStmnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Fillstk()
Dim nRe As Double, nIs As Double, nCl As Double, nR As Integer
Dim rsItem As New ADODB.Recordset, rsIled As New ADODB.Recordset
rsItem.Open "select * from ite0203d where faclevel<0 order by facname", Con, adOpenForwardOnly
If Not rsItem.EOF Then
If Not rsItem.BOF Then rsItem.MoveFirst
nR = 1
Do While Not rsItem.EOF
nRe = 0
nIs = 0
nCl = 0
rsIled.Open "select * from ile0203d where faccode='" & rsItem!faccode & "'", Con, adOpenForwardOnly
If Not rsIled.EOF Then
If Not rsIled.BOF Then rsIled.MoveFirst
Do While Not rsIled.EOF


If rsIled!fvrtype = 2 Or rsIled!fvrtype = 5 Then 'if receipt
    nRe = nRe + rsIled!fqty
ElseIf rsIled!fvrtype = 6 Or rsIled!fvrtype = 1 Then 'if issue
    nIs = nIs + rsIled!fqty
End If


rsIled.MoveNext
Loop
End If
nCl = nRe - nIs
GrdRpt.TextMatrix(nR, 0) = rsItem!facname
GrdRpt.TextMatrix(nR, 1) = nRe
GrdRpt.TextMatrix(nR, 2) = nIs
GrdRpt.TextMatrix(nR, 3) = nCl
GrdRpt.TextMatrix(nR, 4) = rsItem!faccode
nR = nR + 1
GrdRpt.AddItem ""
rsIled.Close


rsItem.MoveNext
Loop
End If
rsItem.Close
End Sub

Private Sub Command1_Click()
Fillstk
End Sub

Private Sub Command2_Click()
Dim nR As Integer
nR = GrdRpt.Rows - 1
Open "c:\files\testfile.TXT" For Output As #1  ' Open file for output
Print #1, String(80, "-")
Print #1, RPad("Product Name", 45) + " " + RPad("Receipt", 10) + " " + RPad("Issue", 10) + " " + RPad("Closing", 10)
Print #1, String(80, "-")

For I = 1 To nR

Print #1, RPad(GrdRpt.TextMatrix(I, 0), 45) + " " + RPad(GrdRpt.TextMatrix(I, 1), 10) + " " + RPad(GrdRpt.TextMatrix(I, 2), 10) + " " + RPad(GrdRpt.TextMatrix(I, 3), 10)

Next

Close #1


End Sub


Private Sub grdIstk_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Frame1.Visible = False
GrdRpt.SetFocus
End If
End Sub


Private Sub GrdRpt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
FillGrid
Frame1.Visible = True
grdIstk.SetFocus
End If
End Sub


Private Sub FillGrid()
Dim nR As Integer, nIss As Double, nRec As Double
txtIss = ""
txtRec = ""
txtSih = ""
Label1.Caption = " Name : " + GrdRpt.TextMatrix(GrdRpt.Row, 0)


grdIstk.Rows = 2
grdIstk.Clear
If lSuBra Then
grdIstk.FormatString = "Slno |  Voucher |        Date  |   Type                | Branch    |               Tag   |         Remark                         |    Receipt   |       Issue"
Else
grdIstk.FormatString = "Slno |  Voucher |        Date  |   Type                    |               Tag   |                      Remark                         |    Receipt   |       Issue"
End If
Set rsIled = New ADODB.Recordset
rsIled.Open "select * from LstIled where faccode='" & GrdRpt.TextMatrix(GrdRpt.Row, 4) & "'", Con, adOpenStatic
If Not rsIled.EOF Then
nR = 1
If Not rsIled.BOF Then rsIled.MoveFirst
Do While Not rsIled.EOF
grdIstk.TextMatrix(nR, 0) = nR
grdIstk.TextMatrix(nR, 1) = rsIled!fvrno
grdIstk.TextMatrix(nR, 2) = rsIled!fvrdate
grdIstk.TextMatrix(nR, 3) = rsIled!Vrname

If lSuBra Then
    grdIstk.TextMatrix(nR, 4) = IIf(IsNull(rsIled!fbranch), "", rsIled!fbranch)
    grdIstk.TextMatrix(nR, 5) = LstTag(rsIled!ftag)
    grdIstk.TextMatrix(nR, 6) = IIf(IsNull(rsIled!fnote), "", rsIled!fnote)
    If rsIled!fvrtype = 1 Or rsIled!fvrtype = 6 Then  ' Issue and sales
     ' If rsIled!ftag <> 5 Then   ' to be actived for main branch
        grdIstk.TextMatrix(nR, 8) = rsIled!fqty
        nIss = nIss + rsIled!fqty
      ' Else
        'grdistk.TextMatrix(nR, 7) = rsIled!fqty           ' if goods transfer issue will not be considerd as minus stock
       ' grdistk.TextMatrix(nR, 6) = rsIled!fbrancht
       'End If
    ElseIf rsIled!fvrtype = 2 Or rsIled!fvrtype = 5 Then    'receipt and purchase
        grdIstk.TextMatrix(nR, 7) = rsIled!fqty
        nRec = nRec + rsIled!fqty
    End If
Else
    grdIstk.TextMatrix(nR, 4) = LstTag(rsIled!ftag)
    grdIstk.TextMatrix(nR, 5) = rsIled!fnote
    If rsIled!fvrtype = 1 Or rsIled!fvrtype = 6 Then
        grdIstk.TextMatrix(nR, 7) = rsIled!fqty
        nIss = nIss + rsIled!fqty
    ElseIf rsIled!fvrtype = 2 Or rsIled!fvrtype = 5 Then
        nRec = nRec + rsIled!fqty
        grdIstk.TextMatrix(nR, 6) = rsIled!fqty
    End If

End If

grdIstk.AddItem ""
nR = nR + 1



rsIled.MoveNext
Loop
End If
rsIled.Close

txtRec = nRec
txtIss = nIss
txtSih = nRec - nIss
End Sub


