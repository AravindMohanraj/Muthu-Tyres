VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmAccCtrl 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Access Control Panel"
   ClientHeight    =   7455
   ClientLeft      =   2370
   ClientTop       =   2625
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Lock All"
      Height          =   345
      Left            =   5745
      TabIndex        =   8
      Top             =   510
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ClearAll"
      Height          =   345
      Left            =   4650
      TabIndex        =   7
      Top             =   510
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Full Access"
      Height          =   345
      Left            =   3585
      TabIndex        =   6
      Top             =   510
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   375
      Left            =   7215
      TabIndex        =   4
      Top             =   6480
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9585
      TabIndex        =   3
      Top             =   6480
      Width           =   1080
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ok"
      Height          =   375
      Left            =   8415
      TabIndex        =   2
      Top             =   6480
      Width           =   1080
   End
   Begin MSFlexGridLib.MSFlexGrid MnuLst 
      Height          =   5445
      Left            =   3165
      TabIndex        =   0
      Top             =   945
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   9604
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      RowHeightMin    =   400
      BackColorBkg    =   16777215
      AllowBigSelection=   0   'False
      GridLinesFixed  =   1
      Appearance      =   0
      FormatString    =   "Sl No |              Menu Name        | Add    | Edit  |   Delete   | Print     |    View   |  Cancel |    Visible"
   End
   Begin MSFlexGridLib.MSFlexGrid UserLst 
      Height          =   5430
      Left            =   150
      TabIndex        =   5
      Top             =   945
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   9578
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   7425
      TabIndex        =   9
      Top             =   510
      Width           =   3150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   195
      Left            =   7410
      TabIndex        =   1
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "FrmAccCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim nOpt As Integer
Private Sub ClearAll()
For I = 1 To MnuLst.Rows - 1
For l = 2 To 8
MnuLst.TextMatrix(I, l) = ""
Next
Next
'Ï

End Sub

Private Sub DelData()
rsAccCtrl.Open "delete from accesscntrl where val(fuser)='" & Val(UserLst.TextMatrix(UserLst.Row, 1)) & "'", Con, adOpenDynamic, adLockPessimistic

End Sub

Private Sub FullAccess()
For I = 1 To MnuLst.Rows - 1
MnuLst.Row = I
For l = 2 To 8
MnuLst.Col = l
MnuLst.CellFontName = "Webdings"
MnuLst.CellFontSize = 14
MnuLst.TextMatrix(I, l) = "Ð"
Next
Next
'Ï
End Sub


Private Sub LoadUserCtrl(code As String)
Dim nR As Integer
rsAccCtrl.Open "select * from N_LstMnuUserCtrl where val(fuser)='" & Val(code) & "'", Con, adOpenStatic
MnuLst.Rows = 2
MnuLst.Clear
MnuLst.FormatString = "Sl No |              Menu Name        | Add    | Edit  |   Delete   | Print     |    View   |  Cancel |    Visible"
If Not rsAccCtrl.EOF Then
nR = 1
If Not rsAccCtrl.BOF Then rsAccCtrl.MoveFirst
Do While Not rsAccCtrl.EOF
MnuLst.Row = nR
MnuLst.TextMatrix(nR, 0) = rsAccCtrl!menucode
MnuLst.TextMatrix(nR, 1) = rsAccCtrl!fmenuname
MnuLst.Col = 2
MnuLst.CellFontName = "Webdings"
MnuLst.CellFontSize = 14

MnuLst.TextMatrix(nR, 2) = IIf(rsAccCtrl!fnew, "Ð", "Ï")
MnuLst.Col = 3
MnuLst.CellFontName = "Webdings"
MnuLst.CellFontSize = 14

MnuLst.TextMatrix(nR, 3) = IIf(rsAccCtrl!fedit, "Ð", "Ï")
MnuLst.Col = 4
MnuLst.CellFontName = "Webdings"
MnuLst.CellFontSize = 14
MnuLst.TextMatrix(nR, 4) = IIf(rsAccCtrl!fdelete, "Ð", "Ï")
MnuLst.Col = 5
MnuLst.CellFontName = "Webdings"
MnuLst.CellFontSize = 14
MnuLst.TextMatrix(nR, 5) = IIf(rsAccCtrl!fprint, "Ð", "Ï")
MnuLst.Col = 6
MnuLst.CellFontName = "Webdings"
MnuLst.CellFontSize = 14

MnuLst.TextMatrix(nR, 6) = IIf(rsAccCtrl!fview, "Ð", "Ï")
MnuLst.Col = 7
MnuLst.CellFontName = "Webdings"
MnuLst.CellFontSize = 14
MnuLst.TextMatrix(nR, 7) = IIf(rsAccCtrl!fcancel, "Ð", "Ï")
MnuLst.Col = 8
MnuLst.CellFontName = "Webdings"
MnuLst.CellFontSize = 14
MnuLst.TextMatrix(nR, 8) = IIf(rsAccCtrl!fvisible, "Ð", "Ï")
nR = nR + 1
MnuLst.AddItem ""
rsAccCtrl.MoveNext
Loop
End If

rsAccCtrl.Close
End Sub

Private Sub LockAll()
For I = 1 To MnuLst.Rows - 1
MnuLst.Row = I
For l = 2 To 8
MnuLst.Col = l
MnuLst.CellFontName = "Webdings"
MnuLst.CellFontSize = 14
MnuLst.TextMatrix(I, l) = "Ï"
Next
Next
'
End Sub

Private Sub SaveUserCtrl()
Set rsAccCtrl = New ADODB.Recordset
rsAccCtrl.Open "select * from accesscntrl ", Con, adOpenDynamic, adLockPessimistic
For I = 1 To MnuLst.Rows - 1
If MnuLst.TextMatrix(I, 0) <> "" And MnuLst.TextMatrix(I, 8) <> "" Then
     
        rsAccCtrl.AddNew
        rsAccCtrl!menucode = Val(MnuLst.TextMatrix(I, 0))
        rsAccCtrl!fuser = Val(UserLst.TextMatrix(UserLst.Row, 1))
        rsAccCtrl!fvisible = IIf(MnuLst.TextMatrix(I, 8) = "Ð", True, False)
        rsAccCtrl!fnew = IIf(MnuLst.TextMatrix(I, 2) = "Ð", True, False)
        rsAccCtrl!fedit = IIf(MnuLst.TextMatrix(I, 3) = "Ð", True, False)
        rsAccCtrl!fdelete = IIf(MnuLst.TextMatrix(I, 4) = "Ð", True, False)
        rsAccCtrl!fcancel = IIf(MnuLst.TextMatrix(I, 7) = "Ð", True, False)
        rsAccCtrl!fprint = IIf(MnuLst.TextMatrix(I, 5) = "Ð", True, False)
        rsAccCtrl!fview = IIf(MnuLst.TextMatrix(I, 6) = "Ð", True, False)
        rsAccCtrl.Update
End If
Next
rsAccCtrl.Close
End Sub

Private Sub Command1_Click()
If nOpt = 1 Then
SaveUserCtrl
ElseIf nOpt = 2 Then
DelData
SaveUserCtrl
End If


ClearAll

UserLst.SetFocus
End Sub

Private Sub Command4_Click()
FullAccess
End Sub


Private Sub Command5_Click()
ClearAll
End Sub

Private Sub Command6_Click()
LockAll
End Sub

Private Sub Form_Load()
FillUser UserLst, "User  Name                                  | code"
LoadMnuLst MnuLst
End Sub



Private Sub MnuLst_DblClick()
If MnuLst.Text = "Ð" Then
MnuLst.Text = "Ï"
ElseIf MnuLst.Text = "Ï" Then
MnuLst.Text = "Ð"
End If
If MnuLst.Col = 8 Then
    If MnuLst.Text = "Ï" Then
       For l = 2 To 7
          MnuLst.TextMatrix(MnuLst.Row, l) = "Ï"
       Next
    End If
End If
End Sub


Private Sub UserLst_Click()
Label2.Caption = UserLst.TextMatrix(UserLst.Row, 0)
Set rsAccCtrl = New ADODB.Recordset
rsAccCtrl.Open "select * from accesscntrl where val(fuser)='" & Val(UserLst.TextMatrix(UserLst.Row, 1)) & "'", Con, adOpenStatic
If Not rsAccCtrl.EOF Then
rsAccCtrl.Close
nOpt = 2
LoadUserCtrl UserLst.TextMatrix(UserLst.Row, 1)
Else
rsAccCtrl.Close
MnuLst.Rows = 2
MnuLst.Clear
MnuLst.FormatString = "Sl No |              Menu Name        | Add    | Edit  |   Delete   | Print     |    View   |  Cancel |    Visible"
nOpt = 1
LoadMnuLst MnuLst
End If

End Sub


