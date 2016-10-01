VERSION 5.00
Begin VB.Form frmGodown 
   BackColor       =   &H00FFFFFF&
   Caption         =   "New Godown Creation"
   ClientHeight    =   4740
   ClientLeft      =   1365
   ClientTop       =   2040
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   8730
   Begin VB.ListBox lstGdwn 
      Appearance      =   0  'Flat
      Height          =   4125
      Left            =   6150
      TabIndex        =   6
      Top             =   210
      Width           =   2460
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4350
      Left            =   60
      TabIndex        =   7
      Top             =   105
      Width           =   6030
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   420
         Left            =   3330
         TabIndex        =   4
         Top             =   3585
         Width           =   1260
      End
      Begin VB.ComboBox cboBrnch 
         Height          =   315
         ItemData        =   "frmGodown.frx":0000
         Left            =   1320
         List            =   "frmGodown.frx":0016
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   2895
         Width           =   1185
      End
      Begin VB.TextBox txtGdName 
         Appearance      =   0  'Flat
         Height          =   1305
         Left            =   1305
         MaxLength       =   5
         TabIndex        =   1
         Top             =   1305
         Width           =   4590
      End
      Begin VB.TextBox txtGd 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1305
         MaxLength       =   10
         TabIndex        =   0
         Top             =   750
         Width           =   1125
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Height          =   420
         Left            =   2010
         TabIndex        =   3
         Top             =   3585
         Width           =   1260
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   420
         Left            =   4665
         TabIndex        =   5
         Top             =   3585
         Width           =   1260
      End
      Begin VB.Label lblBrnch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Left            =   270
         TabIndex        =   10
         Top             =   3015
         Width           =   510
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Godown No"
         Height          =   195
         Left            =   330
         TabIndex        =   9
         Top             =   795
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Left            =   345
         TabIndex        =   8
         Top             =   1410
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmGodown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOpt As Integer
Private Sub LstGodowm()
Set rsGodown = New ADODB.Recordset
rsGodown.Open "select * from godown order by fgodown", Con, adOpenStatic
If Not rsGodown.EOF Then
If Not rsGodown.BOF Then rsGodown.MoveFirst

lstGdwn.Clear
Do While Not rsGodown.EOF
lstGdwn.AddItem rsGodown!fgodown
rsGodown.MoveNext
Loop
End If
rsGodown.Close

End Sub

Private Sub cmdApply_Click()
Set rsGodown = New ADODB.Recordset
rsGodown.Open "select * from godown", Con, adOpenDynamic, adLockPessimistic
rsGodown.AddNew
rsGodown!fgodown = txtGd
rsGodown!fgdes = txtGdName


If lSuBra Then
    rsGodown!fbranch = cboBrnch.ListIndex
End If
rsGodown.Update
rsGodown.Close


Set rsItem = New ADODB.Recordset
Set rsGdSTk = New ADODB.Recordset
rsItem.Open "select * from ite0203d where faclevel<0 order by faccode", Con, adOpenDynamic
If Not rsItem.EOF Then
If Not rsItem.BOF Then rsItem.MoveFirst
rsGdSTk.Open "select * from gdstk ", Con, adOpenDynamic, adLockPessimistic
     '
    Do While Not rsItem.EOF
      
      rsGdSTk.AddNew
      
      rsGdSTk!faccode = rsItem!faccode
      rsGdSTk!fgd = txtGd
      rsGdSTk.Update
    
    rsItem.MoveNext
    Loop


End If
rsItem.Close
rsGdSTk.Close
txtGd = ""
txtGdName = ""
Set rsItem = Nothing
Set rsGdSTk = Nothing
End Sub

Private Sub Form_Load()
nOpt = 1
LstGodowm
If lSuBra Then
    lblBrnch.Visible = True
    cboBrnch.Visible = True
    BranchLoad cboBrnch
Else
    lblBrnch.Visible = False
    cboBrnch.Visible = False
End If
End Sub





Private Sub lstGdwn_Click()
nOpt = 2
'rsGodown.Open "select * from "
End Sub

Private Sub txtGd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtGd <> "" Then
txtGdName.SetFocus
End If
End Sub


Private Sub txtGdName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And lSuBra Then
   cboBrnch.SetFocus
ElseIf KeyAscii = 13 And lSuBra = False Then
    cmdApply.SetFocus
End If
End Sub


