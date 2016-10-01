VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmUserCreation 
   BackColor       =   &H00FFFFFF&
   Caption         =   "User Creation"
   ClientHeight    =   4170
   ClientLeft      =   3345
   ClientTop       =   3375
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   7650
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   6510
      TabIndex        =   8
      Top             =   3570
      Width           =   945
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ok"
      Height          =   330
      Left            =   5385
      TabIndex        =   7
      Top             =   3570
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Height          =   330
      Left            =   4230
      TabIndex        =   3
      Top             =   3570
      Width           =   945
   End
   Begin VB.TextBox TxtPass 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   4725
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1665
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   375
      Left            =   3690
      TabIndex        =   0
      Top             =   255
      Width           =   705
   End
   Begin VB.TextBox txtLogin 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   4740
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid UserLst 
      Height          =   3705
      Left            =   165
      TabIndex        =   4
      Top             =   285
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   6535
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   "User  Name                                  | code"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Left            =   3765
      TabIndex        =   6
      Top             =   1725
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Login"
      Height          =   195
      Left            =   3780
      TabIndex        =   5
      Top             =   1140
      Width           =   765
   End
End
Attribute VB_Name = "frmUserCreation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOpt As Single

Private Sub Command1_Click()
nOpt = 1
txtLogin.SetFocus
End Sub

Private Sub Command2_Click()
Set rsLogin = New ADODB.Recordset

If nOpt = 1 Then
rsLogin.Open "select * from userlogin", Con, adOpenDynamic, adLockPessimistic
rsLogin.AddNew
rsLogin!loginname = txtLogin
rsLogin!fPassword = TxtPass
rsLogin.Update
rsLogin.Close
updateLog "mnu46", nUser, nOpt, Time, ""

ElseIf nOpt = 2 Then
rsLogin.Open "update userlogin set fpassword='" & TxtPass & "' where val(fslno)='" & Val(UserLst.TextMatrix(UserLst.Row, 1)) & "'", Con, adOpenDynamic, adLockPessimistic
updateLog "mnu46", nUser, nOpt, Time, ""

End If
txtLogin = ""
TxtPass = ""

Set rsLogin = Nothing
FillUser UserLst, "User  Name                                  | code"

txtLogin.SetFocus
End Sub

Private Sub Command3_Click()
Set rsLogin = New ADODB.Recordset

If nOpt = 1 Then
rsLogin.Open "select * from userlogin", Con, adOpenDynamic, adLockPessimistic
rsLogin.AddNew
rsLogin!loginname = txtLogin
rsLogin!fPassword = TxtPass
rsLogin.Update
rsLogin.Close
ElseIf nOpt = 2 Then
rsLogin.Open "update userlogin set fpassword='" & TxtPass & "' where val(fslno)='" & Val(UserLst.TextMatrix(UserLst.Row, 1)) & "'", Con, adOpenDynamic, adLockPessimistic

End If
Set rsLogin = Nothing
Unload Me
End Sub


Private Sub Command4_Click()
txtLogin = ""
TxtPass = ""
txtLogin.SetFocus
End Sub

Private Sub Form_Load()
FillUser UserLst, "User  Name                                  | code"
End Sub


Private Sub txtLogin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtLogin <> "" Then
  rsLogin.Open "select * from userlogin where loginname='" & txtLogin & "'", Con, adOpenStatic
  If Not rsLogin.EOF Then

     MsgBox "Allready Exists"
4   Else
           TxtPass.SetFocus
End If
rsLogin.Close
End If
End Sub


Private Sub TxtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And TxtPass <> "" And txtLogin <> "" Then
Command2.SetFocus
End If
End Sub


Private Sub UserLst_Click()
nOpt = 2
txtLogin = UserLst.TextMatrix(UserLst.Row, 0)
TxtPass = UserLst.TextMatrix(UserLst.Row, 2)
txtLogin.SetFocus
End Sub


