VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBom 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "B.O.M"
   ClientHeight    =   6750
   ClientLeft      =   4425
   ClientTop       =   1905
   ClientWidth     =   6990
   FillColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MFxPopUp 
      Height          =   4065
      Left            =   2565
      TabIndex        =   7
      Top             =   1725
      Visible         =   0   'False
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   7170
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   12426621
      ForeColorFixed  =   8454143
      BackColorSel    =   16506822
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   "                                           Description                 |Code"
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4170
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6060
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6060
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1515
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6060
      Width           =   1215
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2835
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6060
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid FlxGrd 
      Height          =   4125
      Left            =   315
      TabIndex        =   3
      Top             =   1125
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   7276
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   12426621
      ForeColorFixed  =   65535
      BackColorSel    =   16506822
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      GridColor       =   0
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      Appearance      =   0
      FormatString    =   "                                               Description            |             Code  |         Quantity"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   330
      MaxLength       =   35
      TabIndex        =   0
      Top             =   570
      Width           =   4980
   End
   Begin VB.TextBox txtPrnName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      Top             =   570
      Width           =   1260
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00A3C05C&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   345
      TabIndex        =   2
      Top             =   165
      Width           =   1200
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edit"
      Begin VB.Menu mnuSrch 
         Caption         =   "Search"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuRR 
         Caption         =   "Remove Record"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "frmBom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsItem As New ADODB.Recordset
Dim rsBom As New ADODB.Recordset
Dim nOpt As Integer, lFrmM As Boolean

Private Sub DelData()
rsBom.Open "delete * from bom where mpartno='" & txtName & "'", Con, adOpenDynamic, adLockPessimistic
End Sub

Private Sub SaveData()
rsBom.Open "select * from bom", Con, adOpenDynamic, adLockPessimistic
    For I = 1 To FlxGrd.Rows - 1
        rsBom.AddNew
        rsBom!spartno = FlxGrd.TextMatrix(I, 1)
        rsBom!qty = Val(FlxGrd.TextMatrix(I, 2))
        rsBom!mpartno = txtPrnName.Text
        rsBom!Desc = txtName.Text
        rsBom.Update
    Next I
    MsgBox "saved"
rsBom.Close
End Sub

Private Function stuffData() As Boolean
    rsBom.Open "select * from bom where mpartno='" & txtPrnName & "'", Con, adOpenDynamic
    If Not rsBom.EOF Then
        If Not rsBom.EOF Then rsBom.MoveFirst
        MFxPopUp.Visible = False
        rsItem.Open "select facname from ite0203d where faccode='" & txtPrnName & "'", Con, adOpenDynamic
        If Not rsItem.EOF Then
        txtPrnName = rsItem!facname
        End If
        rsItem.Close
        nR = 1
        While Not rsBom.EOF
            FlxGrd.TextMatrix(nR, 0) = rsBom!spartno
                rsItem.Open "select facname from ite0203d where faccode='" & rsBom!spartno & "'", Con, adOpenDynamic
                    If Not rsItem.EOF Then
                        FlxGrd.TextMatrix(nR, 1) = rsItem!facname
                    End If
                rsItem.Close
                If Not IsNull(rsBom!qty) Then FlxGrd.TextMatrix(nR, 2) = rsBom!qty
            rsBom.MoveNext
            If Not rsBom.EOF Then FlxGrd.AddItem ""
            nR = nR + 1
        Wend
        stuffData = True
    Else
       stuffData = False
    End If
    
    rsBom.Close
End Function

Private Sub cmdCancel_Click()
clearFunc
nOpt = 1
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
If nOpt = 1 Then
    If VALIDATE Then
        SaveData
        If MsgBox("Do you want to print", vbYesNo) = vbYes Then cmdPrint_Click
        cmdCancel_Click
        txtName.SetFocus
    Else
        MsgBox "few entries are empty"
    End If
ElseIf nOpt = 2 Then
    DelData
    SaveData
    If MsgBox("Do you want to print", vbYesNo) = vbYes Then cmdPrint_Click
    cmdCancel_Click
End If
End Sub

Private Sub swapFunc(KeyCode As Integer)
If FlxGrd.Col = 0 Then
       MFxPopUp.Visible = True
       If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or _
                                               KeyCode = vbKeyPageDown Then
        nPKey = KeyCode
        If KeyCode = vbKeyDown Then
           MFxPopUp.SetFocus
           If (MFxPopUp.Row < MFxPopUp.Rows) Then MFxPopUp.Row = MFxPopUp.Row + 1
           MFxPopUp.Col = 0
        ElseIf KeyCode = vbKeyUp Then
           MFxPopUp.SetFocus
           If (MFxPopUp.Row > 1) Then MFxPopUp.Row = MFxPopUp.Row - 1
           MFxPopUp.Col = 0
        End If
           MFxPopUp.SetFocus
    End If
End If
End Sub

Private Sub cmdPrint_Click()
Dim cPno As String
    cPno = txtName.Text
    If cPno <> "" Then
        dbPath = App.Path + "\" + "winiv.mdb"
        cr.DataFiles(0) = dbPath
        cr.SelectionFormula = "{BomPrint.MPartno}=  '" & txtName.Text & "' "
        cr.ReportFileName = App.Path + "\bomprn.rpt"
        cr.WindowState = crptMaximized
        cr.Action = 1
    Else
         MsgBox "Important fields are empty", vbCritical
    End If
End Sub

Private Sub FlxGrd_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
        If KeyAscii = vbKeyDown Then swapFunc (KeyAscii)
          If FlxGrd.Col = 0 Then
                     MFxPopUp.Visible = True
                   
                      If KeyAscii = 8 Then      ' for backspace
                         If (IsEmpty(FlxGrd.Text)) Then
                               No = FlxGrd.Row
                               If (No > 1) Then
                                 FlxGrd.RemoveItem (FlxGrd.Row)
                                 FlxGrd.Row = No - 1
                                 FlxGrd.Col = 2
                                 MFxPopUp.Visible = False
                               End If
                          Else
                             Stg = FlxGrd.Text
                             Strg = Mid(Stg, 1, (Len(Stg) - 1))
                            FlxGrd.Text = Strg
                          End If
                      Else                             'for normal
                        FlxGrd.TextMatrix(FlxGrd.Row, 0) = FlxGrd.TextMatrix(FlxGrd.Row, 0) + Chr(KeyAscii)
                           Find MFxPopUp, UCase(FlxGrd.Text), 0
                      End If
        ElseIf FlxGrd.Col = 2 Then
              If (KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13 Then
                      KeyAscii = KeyAscii
              Else
                      KeyAscii = 0
              Exit Sub
              End If
              If KeyAscii = 8 Then                   'for back space
                        If (IsEmpty(FlxGrd.Text)) Then
                         Else
                            Stg = FlxGrd.Text
                             Strg = Mid(Stg, 1, (Len(Stg) - 1))
                            FlxGrd.Text = Strg
                          End If
              Else                                       'for normal
                        FlxGrd.TextMatrix(FlxGrd.Row, 2) = FlxGrd.TextMatrix(FlxGrd.Row, 2) + Chr(KeyAscii)
              End If
        End If
        
ElseIf KeyAscii = 13 Then     'on pressing enter
        
        If FlxGrd.Col = 0 Then   'auto display of discription from ite03023d
             If (IsEmpty(FlxGrd.Text)) Then
                 MsgBox " cant be empty"
                 FlxGrd.Row = FlxGrd.Row
                 FlxGrd.Col = 0
                 GoTo er
             Else
                  FlxGrd.TextMatrix(FlxGrd.Row, 0) = MFxPopUp.TextMatrix(MFxPopUp.Row, 0)
                  FlxGrd.TextMatrix(FlxGrd.Row, 1) = MFxPopUp.TextMatrix(MFxPopUp.Row, 1)
                  FlxGrd.Col = 2
                  MFxPopUp.Visible = False
             End If
        ElseIf FlxGrd.Col = 2 Then    'on pressing  enter add inf new record
               If (IsEmpty(FlxGrd.Text)) Then
                 MsgBox " cant be empty"
                 FlxGrd.Row = FlxGrd.Row
                 FlxGrd.Col = FlxGrd.Col
               Else
                    If FlxGrd.Rows - 1 = FlxGrd.Row Then
                        FlxGrd.AddItem ""
                        FlxGrd.Row = FlxGrd.Row + 1
                        FlxGrd.TopRow = FlxGrd.Row
                        FlxGrd.Col = 0
                    End If
               End If
        End If
End If
er:
End Sub



Private Sub Form_Load()
'--------------- form position
Me.Top = 1155
Me.Left = 4365
Me.Height = 7560
Me.Width = 7110
'------------------------------

FillPop
nOpt = 1 ' New mode
lFrmM = True
End Sub

Private Sub MFxPopUp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And lFrmM = False Then FlxGrd_KeyPress (13) Else txtName_KeyPress (13)
End Sub


Private Sub mnuRR_Click()
If MsgBox("Remove a line", vbYesNo + vbDefaultButton2) = vbYes Then FlxGrd.RemoveItem (FlxGrd.Row)
End Sub

Private Sub mnuSrch_Click()
MFxPopUp.Visible = True
End Sub

Private Sub txtName_GotFocus()
lFrmM = True
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 27 Then
      
      If (IsEmpty(txtName.Text)) Then
           MsgBox " cant be empty"
           txtName.SetFocus
      Else
      
             If MFxPopUp.Visible = True Then
              txtName = MFxPopUp.Text
              txtPrnName = MFxPopUp.TextMatrix(MFxPopUp.Row, 1)
             MFxPopUp.Visible = False
             End If
             
            txtName = UCase(txtName)
            txtPrnName = MFxPopUp.TextMatrix(MFxPopUp.Row, 1)
            FlxGrd.Clear
            FlxGrd.FormatString = "                                               Description            |             Code  |         Quantity"
            FlxGrd.Rows = 2
            
          If stuffData Then
             nOpt = 2
             txtPrnName.SetFocus
          Else
             nOpt = 1
               rsItem.Open "select * from ite0203d where faccode='" & txtPrnName & "'", Con, adOpenDynamic
                If Not rsItem.EOF Then
                    txtPrnName = rsItem!faccode
                    FlxGrd.SetFocus
                Else
                    MsgBox "Part No not found", vbCritical
                End If
               rsItem.Close
          End If
      End If
End If

End Sub

Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode <> vbKeyEscape And KeyCode <> vbKeyRight And KeyCode <> vbKeyLeft _
            And txtName.Text <> "" And KeyCode <> vbKeyReturn Then

                           Find MFxPopUp, UCase(txtName.Text), 0
End If

End Sub

Private Sub txtName_LostFocus()
lFrmM = False
End Sub

Private Sub txtPrnName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 27 Then
       If (IsEmpty(txtPrnName.Text)) Then
          MsgBox " cant be empty"
          txtPrnName.SetFocus
       Else
'set the focusing to first cell
          FlxGrd.SetFocus
          FlxGrd.Row = 1
          FlxGrd.Col = 0
       End If
End If
End Sub


Public Function VALIDATE() As Boolean 'validation before saving
For I = 1 To FlxGrd.Rows - 1
  If FlxGrd.TextMatrix(I, 0) <> "" Then
        For j = 0 To FlxGrd.Cols - 1
            If (FlxGrd.TextMatrix(I, j) = "") Then
                VALIDATE = False
                GoTo skiP
            End If
        Next j
  Else
     VALIDATE = False
  End If
Next I


'If flxgrd.Rows > 1 Then
'    For i = 1 To flxgrd.Rows
'        If flxgrd.TextMatrix(i, 0) <> "" And flxgrd.TextMatrix(i, 2) <> "" Then
'
'        End If
'    Next i
'End If
If (IsEmpty(txtName) And IsEmpty(txtPrnName)) Then
VALIDATE = False
GoTo skiP
End If
VALIDATE = True
skiP:
End Function

Public Function IsEmpty(Str As String) As Boolean 'to check if a text is empty
If (Len(Str) < 1) Then
IsEmpty = True
Else
IsEmpty = False
End If
End Function

Public Sub FillPop()            'procedure to fill the content in the popup grid mfxpopup
rsItem.Open "select * from ITE0203D where faclevel<0 ORDER BY facname", Con, adOpenDynamic
If Not rsItem.BOF Then rsItem.MoveFirst
I = 1
While rsItem.EOF = False
MFxPopUp.TextMatrix(I, 0) = rsItem!facname
MFxPopUp.TextMatrix(I, 1) = rsItem!faccode
MFxPopUp.AddItem ""
rsItem.MoveNext
I = I + 1
Wend
rsItem.Close
End Sub

Public Sub clearFunc()
txtName = ""
txtPrnName = ""
MFxPopUp.Visible = False
FlxGrd.Clear
FlxGrd.FormatString = "                                               Description            |             Code  |         Quantity"
FlxGrd.Rows = 2
txtName.SetFocus
End Sub
