VERSION 5.00
Begin VB.Form frmBranch 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Branch Creation"
   ClientHeight    =   3780
   ClientLeft      =   2820
   ClientTop       =   3645
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   6315
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3480
      Left            =   105
      TabIndex        =   3
      Top             =   180
      Width           =   6045
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   420
         Left            =   4590
         TabIndex        =   5
         Top             =   2805
         Width           =   1260
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save"
         Height          =   420
         Left            =   3270
         TabIndex        =   4
         Top             =   2820
         Width           =   1260
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1305
         MaxLength       =   10
         TabIndex        =   0
         Top             =   750
         Width           =   1680
      End
      Begin VB.TextBox txtLocat 
         Appearance      =   0  'Flat
         Height          =   675
         Left            =   1290
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1290
         Width           =   4590
      End
      Begin VB.ComboBox cboBranch 
         Height          =   315
         ItemData        =   "frmBranch.frx":0000
         Left            =   1320
         List            =   "frmBranch.frx":000A
         TabIndex        =   2
         Text            =   "No"
         Top             =   2250
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Godown"
         Height          =   195
         Left            =   270
         TabIndex        =   8
         Top             =   2355
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         Height          =   195
         Left            =   345
         TabIndex        =   7
         Top             =   1410
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Left            =   330
         TabIndex        =   6
         Top             =   795
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SaveData()
Set rsBranch = New ADODB.Recordset
rsBranch.Open "select * from branch", Con, adOpenDynamic, adLockPessimistic
rsBranch.AddNew
rsBranch!fbranch = txtCode
rsBranch!fbranchname = txtlocate
rsBranch!fgodown = cboBranch.ListIndex
rsBranch.Update
rsBranch.Close
Set rsBranch = Nothing





Set rsItem = New ADODB.Recordset
Set rsBrStk = New ADODB.Recordset
rsItem.Open "select * from ite0203d where faclevel<0 order by faccode", Con, adOpenDynamic
If Not rsItem.EOF Then
If Not rsItem.BOF Then rsItem.MoveFirst
rsBrStk.Open "select * from brstk ", Con, adOpenDynamic, adLockPessimistic
    Do While Not rsItem.EOF
      rsBrStk.AddNew
      rsBrStk!faccode = rsItem!faccode
      rsBrStk!fbranch = txtCode
      rsBrStk!fbal = 0
      rsBrStk.Update
    rsItem.MoveNext
    Loop
End If
rsItem.Close
rsBrStk.Close
txtCode = ""
txtLocat = ""
Set rsItem = Nothing
Set rsBrStk = Nothing
End Sub

Private Sub cboBranch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command3.SetFocus
End If
End Sub


Private Sub Command3_Click()
SaveData
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtCode <> "" Then
txtLocat.SetFocus

End If
End Sub


Private Sub txtLocat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command3.SetFocus
End If
End Sub


