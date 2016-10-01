VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIsCrd 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Item Stock Card"
   ClientHeight    =   8010
   ClientLeft      =   705
   ClientTop       =   450
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   10005
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7005
      Left            =   75
      TabIndex        =   1
      Top             =   915
      Visible         =   0   'False
      Width           =   9810
      Begin VB.TextBox txtSih 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7605
         TabIndex        =   7
         Top             =   6510
         Width           =   1050
      End
      Begin VB.TextBox txtIss 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8640
         TabIndex        =   6
         Top             =   6030
         Width           =   1020
      End
      Begin VB.TextBox txtRec 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7605
         TabIndex        =   5
         Top             =   6030
         Width           =   1050
      End
      Begin MSFlexGridLib.MSFlexGrid GrdRpt 
         Height          =   5295
         Left            =   30
         TabIndex        =   2
         Top             =   585
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   9340
         _Version        =   393216
         Cols            =   9
         Appearance      =   0
         FormatString    =   $"frmIsCrd.frx":0000
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name :"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   75
         TabIndex        =   4
         Top             =   180
         Width           =   4395
      End
   End
   Begin VB.TextBox txtItemLst 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   210
      TabIndex        =   0
      Top             =   975
      Width           =   5310
   End
   Begin MSFlexGridLib.MSFlexGrid ProductList 
      Height          =   6045
      Left            =   210
      TabIndex        =   3
      Top             =   1425
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   10663
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   "Product  Name                                  | Code   |  Stock       | MRP "
   End
End
Attribute VB_Name = "frmIsCrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub FillGrid()
Dim nR As Integer, nIss As Double, nRec As Double
txtIss = ""
txtRec = ""
txtSih = ""
Label1.Caption = " Name : " + ProductList.TextMatrix(ProductList.Row, 0)
GrdRpt.Rows = 2
GrdRpt.Clear
If lSuBra Then
GrdRpt.FormatString = "Slno |  Voucher |        Date  |   Type                | Branch    |               Tag   |         Remark                         |    Receipt   |       Issue"
Else
GrdRpt.FormatString = "Slno |  Voucher |        Date  |   Type                    |               Tag   |                      Remark                         |    Receipt   |       Issue"
End If
Set rsIled = New ADODB.Recordset
rsIled.Open "select * from LstIled where faccode='" & ProductList.TextMatrix(ProductList.Row, 1) & "'", Con, adOpenStatic
If Not rsIled.EOF Then
nR = 1
If Not rsIled.BOF Then rsIled.MoveFirst
Do While Not rsIled.EOF
GrdRpt.TextMatrix(nR, 0) = nR
GrdRpt.TextMatrix(nR, 1) = rsIled!fvrno
GrdRpt.TextMatrix(nR, 2) = rsIled!fvrdate
GrdRpt.TextMatrix(nR, 3) = rsIled!Vrname

If lSuBra Then
    GrdRpt.TextMatrix(nR, 4) = IIf(IsNull(rsIled!fbranch), "", rsIled!fbranch)
    GrdRpt.TextMatrix(nR, 5) = LstTag(rsIled!ftag)
    GrdRpt.TextMatrix(nR, 6) = IIf(IsNull(rsIled!fnote), "", rsIled!fnote)
    If rsIled!fvrtype = 1 Or rsIled!fvrtype = 6 Then  ' Issue and sales
     ' If rsIled!ftag <> 5 Then   ' to be actived for main branch
        GrdRpt.TextMatrix(nR, 8) = rsIled!fqty
        nIss = nIss + rsIled!fqty
      ' Else
        'GrdRpt.TextMatrix(nR, 7) = rsIled!fqty           ' if goods transfer issue will not be considerd as minus stock
       ' GrdRpt.TextMatrix(nR, 6) = rsIled!fbrancht
       'End If
    ElseIf rsIled!fvrtype = 2 Or rsIled!fvrtype = 5 Then    'receipt and purchase
        GrdRpt.TextMatrix(nR, 7) = rsIled!fqty
        nRec = nRec + rsIled!fqty
    End If
Else
    GrdRpt.TextMatrix(nR, 4) = LstTag(rsIled!ftag)
    GrdRpt.TextMatrix(nR, 5) = rsIled!fnote
    If rsIled!fvrtype = 1 Or rsIled!fvrtype = 6 Then
        GrdRpt.TextMatrix(nR, 7) = rsIled!fqty
        nIss = nIss + rsIled!fqty
    ElseIf rsIled!fvrtype = 2 Or rsIled!fvrtype = 5 Then
        nRec = nRec + rsIled!fqty
        GrdRpt.TextMatrix(nR, 6) = rsIled!fqty
    End If

End If

GrdRpt.AddItem ""
nR = nR + 1



rsIled.MoveNext
Loop
End If
rsIled.Close

txtRec = nRec
txtIss = nIss
txtSih = nRec - nIss

End Sub

Private Sub Form_Load()
FillItem ProductList, -1, "Product  Name                                  | Code   |  Stock       | MRP "
End Sub


Private Sub GrdRpt_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Frame1.Visible = False
End Sub


Private Sub productlist_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
FillGrid
Frame1.Visible = True
GrdRpt.SetFocus
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

FillGrid
Frame1.Visible = True
GrdRpt.SetFocus



ElseIf KeyAscii = 27 Then
    
'    fremIte.Visible = False


End If
End Sub


Private Sub txtItemLst_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 13 Then
 lFnd = Find(ProductList, UCase(txtItemLst.Text), 0)
End If

End Sub
