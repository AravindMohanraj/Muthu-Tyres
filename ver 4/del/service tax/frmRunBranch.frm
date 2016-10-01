VERSION 5.00
Begin VB.Form frmRunBranch 
   Caption         =   "Generate Branch"
   ClientHeight    =   2340
   ClientLeft      =   5430
   ClientTop       =   3660
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   4020
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Branch"
      Height          =   495
      Left            =   1410
      TabIndex        =   0
      Top             =   930
      Width           =   1215
   End
End
Attribute VB_Name = "frmRunBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set rsItem = New ADODB.Recordset
Set rsBrStk = New ADODB.Recordset


Set rsBranch = New ADODB.Recordset
rsBranch.Open "select * from branch", Con, adOpenDynamic, adLockPessimistic
If Not rsBranch.EOF Then
If Not rsBranch.BOF Then rsBranch.MoveFirst
Do While Not rsBranch.EOF

    rsItem.Open "select * from ite0203d where faclevel<0 order by faccode", Con, adOpenStatic
    
    If Not rsItem.EOF Then
       Do While Not rsItem.EOF
          rsBrStk.Open "select * from brstk where faccode='" & rsItem!faccode & "' and fbranch='" & rsBranch!fbranch & "'", Con, adOpenDynamic, adLockPessimistic
          If rsBrStk.EOF Then
          rsBrStk.AddNew
          rsBrStk!faccode = rsItem!faccode
          rsBrStk!fbranch = rsBranch!fbranch
          rsBrStk!fbal = 0
          rsBrStk.Update
          End If
          rsBrStk.Close
       
       rsItem.MoveNext
       Loop
    End If
    rsItem.Close


rsBranch.MoveNext
Loop

End If
rsBranch.Close



Set rsItem = Nothing
Set rsBrStk = Nothing

End Sub


