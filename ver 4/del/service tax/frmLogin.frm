VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   4410
   ClientTop       =   4245
   ClientWidth     =   4770
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   3480
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2835
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1950
      Width           =   1770
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Left            =   2835
      TabIndex        =   1
      Top             =   1365
      Width           =   1770
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1605
      TabIndex        =   2
      Top             =   1965
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1605
      TabIndex        =   0
      Top             =   1380
      Width           =   1185
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function chkAccess() As Boolean
Set rsLogin = New ADODB.Recordset

chkAccess = False
rsLogin.Open "select * from userlogin", Con, adOpenStatic
If Not rsLogin.EOF Then

If Not rsLogin.BOF Then rsLogin.MoveFirst
    Do While Not rsLogin.EOF
        
       If rsLogin!loginname = txtLogin And rsLogin!fpassword = TxtPass Then
          chkAccess = True
          cUser = rsLogin!loginname
          nUser = rsLogin!fslno
          
          Exit Do
       End If
    
    rsLogin.MoveNext
    Loop
End If
rsLogin.Close


End Function

Private Sub Text2_Change()

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   If MsgBox("Exit", vbYesNo) = vbYes Then
      Unload Me
      End
   End If
End If
End Sub

Private Sub Form_Load()
Main

End Sub


Private Sub Form_Unload(Cancel As Integer)
'Con.Close
End Sub


Private Sub txtLogin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtLogin <> "" Then
TxtPass.SetFocus
End If
End Sub


Private Sub TxtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And TxtPass <> "" And txtLogin <> "" Then
   If chkAccess Then
       frmMain.Show
       Unload Me
frmMain.MainSb.Panels(2).Text = "User Name: " + cUser
'IntMenu nUser

   Else
      MsgBox "Access Denied"
   End If
End If
End Sub

