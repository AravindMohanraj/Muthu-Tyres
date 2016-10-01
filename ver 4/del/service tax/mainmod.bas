Attribute VB_Name = "MainMod"
Public Con As New ADODB.Connection
Public dFdt As Date, dTdt As Date, dCdt As Date, nLst As Integer
Public lSuGd As Boolean, lSuBra As Boolean, lSuTax As Boolean, lSuBc As Boolean
Public lSuPc As Boolean, lSuDed As Boolean, lSuOlan As Boolean, lSuBar As Boolean, lSuItax As Boolean
Public nBM As Single, lSuDc As Boolean, lSuOn As Boolean, lSuSTax As Boolean, cNtDes As String

Public cDbPath As String, cDbFname As String, rsPlst As ADODB.Recordset
Public rsAcc As ADODB.Recordset, rsItem As ADODB.Recordset, rsBStp As ADODB.Recordset
Public rsAdd As ADODB.Recordset, rsArea As ADODB.Recordset, rsBstrt As ADODB.Recordset
Public rsNum As ADODB.Recordset, rsLogin As ADODB.Recordset, rsLog As ADODB.Recordset
Public rsStartUP As ADODB.Recordset, nUser As Long, cUser As String, rsMenu As ADODB.Recordset
Public rsAccCtrl As ADODB.Recordset, rsBranch As ADODB.Recordset, rsGodown As ADODB.Recordset
Public rsTxM As ADODB.Recordset, rsBC As ADODB.Recordset
Public rsLed As ADODB.Recordset, rsIled As ADODB.Recordset, rsBill As ADODB.Recordset
Public rsPur As ADODB.Recordset, rsRec As ADODB.Recordset, rsPay As ADODB.Recordset
Public rsGdTrn As ADODB.Recordset, rsGdSTk As ADODB.Recordset, rsBLst As ADODB.Recordset
Public rsTemp As ADODB.Recordset, rsBrStk As ADODB.Recordset, rsBBnum As ADODB.Recordset

Public cCash As String, cBank As String, cSales As String, cPurchase As String
Public cVat As String, cBankchg As String, cCard As String, cTrans As String, cDiscount As String
Public nDmal As Single, cDP As String, cHead1 As String, cHead2 As String, cHead3 As String, cHead4 As String
Public cHead5 As String, cHead6 As String, cBot1 As String, cBot2 As String, cBot3 As String, cBot4 As String
Public nNoP As Single, nNoLR As Single, nNolE As Single
Public Function GetCboDataIndex(cBo As ComboBox, SearchData As String) As Integer
'---------------------------------------------------------------------------------------------------
'-------09/12/2008  Rakesh
' CBo->Combobox name to search, SearchKey-> search value for data listed in Cbo
'---------------------------------------------------------------------------------------------------
    GetCboDataIndex = "0"
    For I = 0 To cBo.ListCount - 1
        cBo.ListIndex = I
        If SearchData = cBo.Text Then
            GetCboDataIndex = I
            Exit For
        End If
    Next I
End Function

Public Function Round050(nNum As Double) As Double
Dim nPos, nNv As Integer, cChrVal, cLv, cRv, cFdec As String
nPos = InStr(1, Trim(Str(nNum)), ".")

If nPos <> 0 Then
cChrVal = Mid(Str(nNum), nPos + 2, 2)
If Len(cChrVal) > 1 Then
cLv = Left(cChrVal, 1)
cRv = Right(cChrVal, 1)
nNv = Val(cRv)
If Val(cLv) < 5 Then
    If nNv < 3 Then
        'cFdec = cLv + "0"
        cFdec = "0"
    ElseIf nNv >= 3 And nNv <= 5 Then
        cFdec = cLv + "5"
    ElseIf nNv > 5 Then
        cFdec = Str((Val(cLv) + 1)) + "0"
    End If
Round050 = Format(Val(Mid(Trim(Str(nNum)), 1, nPos - 1) + "." + cFdec), "######0.00")
    
Else
cChrVal = Mid(Trim(Str(nNum)), nPos - 1, 1)
cLv = Mid(Trim(Str(nNum)), 1, nPos - 2)
cFdec = Trim(cLv + Str(Val(cChrVal) + 1))
Round050 = Format(Val(cFdec), "######0.00")

End If
Else
   Round050 = Format(nNum, "######0.00")
End If
Else
    Round050 = Format(nNum, "######0.00")
End If
End Function


Public Function nearRnd(nNum As Double) As Double
Dim nPos, nNv As Integer, cChrVal, cLv, cRv, cFdec As String
nPos = InStr(1, Trim(Str(nNum)), ".")


If nPos <> 0 Then


cChrVal = Left(Mid(Str(nNum), nPos + 2, 2) + String(2, "0"), 2)
cRv = Mid(Trim(Str(nNum)), 1, nPos - 1)

If Val(cChrVal) < 50 Then
nearRnd = Format(cRv, "#####0.00")
ElseIf Val(cChrVal) > 49 Then
nearRnd = Format(Val(cRv) + 1, "#####0.00")
End If
Else
nearRnd = nNum
End If

End Function

Public Function AccountSalVr(cCrcode As String, cDrCode As String, nAmt As Double, Optional nTax As Double) As Boolean


End Function

Public Function RPad(cNum As String, nLen As Long) As String
Dim CPad As String, nlStr As Integer
CPad = Right(cNum, nLen)
nlStr = nLen - Len(CPad)
RPad = Right(cNum, nLen) + String(nlStr, " ")
End Function
Public Function LPad(cNum As String, nLen As Integer) As String
Dim CPad As String, nlStr As Integer
CPad = Left(cNum, nLen)
nlStr = nLen - Len(CPad)
LPad = String(nlStr, " ") + Left(cNum, nLen)
End Function
Public Function CPad(s As String, n As Integer) As String
Dim nLen As Integer
nLen = (n - Len(s)) / 2
c = " "
CPad = String(nLen, c) + s + String(nLen, c)
End Function

Public Function datecon(dDt As Date) As String
Dim cDay As String, cMonth As String, cYear As String, cMon As String, cDt As String
cDt = Trim(Str(dDt))
If cDt = "" Or cDt = "__/__/__" Then
datecon = "__/__/__"
Exit Function
End If
cDay = Left(cDt, 2)
cMonth = Mid(cDt, 4, 3)
cYear = Right(cDt, 2)
If cMonth = "Jan" Then
    cMon = "01"
ElseIf cMonth = "Feb" Then
    cMon = "02"
ElseIf cMonth = "Mar" Then
    cMon = "03"
ElseIf cMonth = "Apr" Then
    cMon = "04"
ElseIf cMonth = "May" Then
    cMon = "05"
ElseIf cMonth = "Jun" Then
    cMon = "06"
ElseIf cMonth = "Jul" Then
    cMon = "07"
ElseIf cMonth = "Aug" Then
    cMon = "08"
ElseIf cMonth = "Sep" Then
    cMon = "09"
ElseIf cMonth = "Oct" Then
    cMon = "10"
ElseIf cMonth = "Nov" Then
    cMon = "11"
ElseIf cMonth = "Dec" Then
    cMon = "12"
End If
datecon = cDay + "/" + cMon + "/" + cYear
End Function
Public Sub AddGdStk(cCode As String, nStock As Double, Optional cGd As String)
    Set rsGdSTk = New ADODB.Recordset
    rsGdSTk.Open "update gdstk set fbal=fbal+'" & nStock & "' where faccode='" & cCode & "' and fgd='" & cGd & "'", Con, adOpenDynamic, adLockPessimistic
  
End Sub

Public Function NumToWord(nNum As Double) As String
Dim a(100) As String
Dim nPos, nInc, nVar As Integer, cChrVal, cLv, cWord As String

cChrVal = Trim(Str(nNum))
nPos = InStr(1, cChrVal, ".")
If nPos <> 0 Then
cLv = Mid(cChrVal, 1, nPos - 1)
cRv = Mid(cChrVal, nPos + 1, 2) + "0"
Else
cLv = cChrVal
cRv = "00"
End If
If nPos - 1 > 9 Then
    NumToWord = " "
    Exit Function
Else
cLv = String(9 - Len(cLv), "0") + cLv
End If









a(1) = "One"
a(2) = "Two"
a(3) = "Three"
a(4) = "Four"
a(5) = "Five"
a(6) = "Six"
a(7) = "Seven"
a(8) = "Eight"
a(9) = "Nine"
a(10) = "Ten"
a(11) = "Eleven"
a(12) = "Twelve"
a(13) = "Thirteen"
a(14) = "Fourteen"
a(15) = "Fifteen"
a(16) = "Sixteen"
a(17) = "Seventeen"
a(18) = "Eighteen"
a(19) = "Nineteen"
a(20) = "Twenty"
a(30) = "Thirty"
a(40) = "Fourty"
a(50) = "Fifty"
a(60) = "Sixty"
a(70) = "Seventy"
a(80) = "Eighty"
a(90) = "Ninty"

'653422365.23
For I = 1 To 3

    If Val(Left(cLv, 2)) = 0 Then
       cLv = Mid(cLv, 3, Len(cLv))
       
    
    ElseIf Val(Left(cLv, 1)) = 0 Then
        If I = 1 Then ' Crore
             cWord = a(Val(Left(cLv, 2))) + " Crore "
             cLv = Mid(cLv, 3, Len(cLv))
        
        ElseIf I = 2 Then ' Lakhs
             cWord = a(Val(Left(cLv, 2))) + " Lakhs "
             cLv = Mid(cLv, 3, Len(cLv))
        
        ElseIf I = 3 Then ' Thousand
             cWord = a(Val(Left(cLv, 2))) + " Thousand "
             cLv = Mid(cLv, 3, Len(cLv))
        End If
    Else
        If I = 1 Then ' crore
          If Val(Left(cLv, 1)) = 1 Then
             cWord = a(Val(Left(cLv, 2))) + " Crore "
             cLv = Mid(cLv, 3, Len(cLv))

          ElseIf Val(Left(cLv, 2)) = 0 Then
             cWord = a(Val(Left(cLv, 2))) + " Crore "
             cLv = Mid(cLv, 3, Len(cLv))
          
          Else
             cWord = a(Val(Left(cLv, 1) + "0")) + a(Val(Mid(cLv, 2, 1))) + " Crore "
             cLv = Mid(cLv, 3, Len(cLv))
    
          End If
             
        ElseIf I = 2 Then 'lakhs
                  If Val(Left(cLv, 1)) = 1 Then
             cWord = cWord + a(Val(Left(cLv, 2))) + " Lakhs "
       cLv = Mid(cLv, 3, Len(cLv))
          ElseIf Val(Left(cLv, 2)) = 0 Then
             cWord = cWord + a(Val(Left(cLv, 2))) + " Lakhs "
       cLv = Mid(cLv, 3, Len(cLv))
          Else
             cWord = cWord + a(Val(Left(cLv, 1) + "0")) + a(Val(Mid(cLv, 2, 1))) + " Lakhs "
       cLv = Mid(cLv, 3, Len(cLv))
          End If
 
        ElseIf I = 3 Then 'Thousand
          If Val(Left(cLv, 1)) = 1 Then
             cWord = cWord + a(Val(Left(cLv, 2))) + " Thousand "
             cLv = Mid(cLv, 3, Len(cLv))
          ElseIf Val(Left(cLv, 2)) = 0 Then
             cWord = cWord + a(Val(Left(cLv, 2))) + " Thousand "
            cLv = Mid(cLv, 3, Len(cLv))
          Else
             cWord = cWord + a(Val(Left(cLv, 1) + "0")) + a(Val(Mid(cLv, 2, 1))) + " Thousand "
             cLv = Mid(cLv, 3, Len(cLv))
          End If
 
        End If
    End If
Next



If Val(Right(cLv, 2)) = 0 Then  'Hundred
    If Val(Left(cLv, 1)) <> 0 Then
        cWord = cWord + IIf(Val(cRv) = 0, "and ", "") + a(Val(Left(cLv, 1))) + " Hundered "
        cLv = Mid(cLv, 2, Len(cLv))
    End If

Else
    If Val(Left(cLv, 1)) <> 0 Then
        cWord = cWord + a(Val(Left(cLv, 1))) + " Hundered" + IIf(Val(cRv) = 0, " and ", "")
        cLv = Mid(cLv, 2, Len(cLv))
    Else
        cLv = Right(cLv, 2)
    End If

          If Val(Left(cLv, 1)) = 1 Then
             cWord = cWord + a(Val(Left(cLv, 2)))
             cLv = Mid(cLv, 3, Len(cLv))
          ElseIf Val(Left(cLv, 2)) = 0 Then
             cWord = cWord + a(Val(Left(cLv, 2)))
             cLv = Mid(cLv, 3, Len(cLv))
          Else
             cWord = cWord + a(Val(Left(cLv, 1) + "0")) + " " + a(Val(Mid(cLv, 2, 1)))
             cLv = Mid(cLv, 3, Len(cLv))
          End If

End If


If Val(cRv) <> 0 Then 'Paise
          If Val(Left(cRv, 1)) = 1 Then
             cWord = cWord + "and " + a(Val(Left(cRv, 2))) + " Paise"
             cRv = Mid(cRv, 3, Len(cRv))
          ElseIf Val(Left(cRv, 2)) = 0 Then
             cWord = cWord + "and " + a(Val(Left(cRv, 2))) + " Paise"
             cRv = Mid(cRv, 3, Len(cRv))
          Else
             cWord = cWord + " and " + a(Val(Left(cRv, 1) + "0")) + " " + a(Val(Mid(cRv, 2, 1))) + " Paise"
             cRv = Mid(cRv, 3, Len(cRv))
          End If
End If

NumToWord = cWord
End Function




Public Sub BranchLoad(oObject As ComboBox)
Set rsBranch = New ADODB.Recordset
rsBranch.Open "select * from branch", Con, adOpenStatic
If Not rsBranch.EOF Then
If Not rsBranch.BOF Then rsBranch.MoveFirst
oObject.Clear
Do While Not rsBranch.EOF
oObject.AddItem rsBranch!FBRANCH

rsBranch.MoveNext
Loop
End If
rsBranch.Close
Set rsBranch = Nothing
End Sub


Public Sub dbStart()
Con.Open "Provider=Microsoft.Jet.Oledb.4.0;data source ='" & cDbPath + "\" + cDbFname & "'"

End Sub

Public Sub FillAcc(oGrid As MSFlexGrid, nLevel As Double, cString As String)
Dim nR As Integer
Set rsAcc = New ADODB.Recordset
If nLevel > 0 Then
rsAcc.Open "select * from acc0203d where val(faclevel) > 0 order by facname ", Con, adOpenStatic
ElseIf nLevel < 0 Then
rsAcc.Open "select * from acc0203d where val(faclevel) < 0 order by facname ", Con, adOpenStatic

End If
If Not rsAcc.EOF Then
oGrid.Rows = 2
oGrid.Clear
oGrid.FormatString = cString
nR = 1
    If Not rsAcc.BOF Then rsAcc.MoveFirst
    Do While Not rsAcc.EOF
    oGrid.TextMatrix(nR, 0) = rsAcc!facname
    oGrid.TextMatrix(nR, 1) = rsAcc!faccode
    oGrid.TextMatrix(nR, 2) = rsAcc!facparent
    
    oGrid.AddItem ""
    nR = nR + 1
    rsAcc.MoveNext
    Loop
End If
rsAcc.Close
Set rsAcc = Nothing

End Sub

Public Sub FillRptCmbo(oCombo As ComboBox)
Dim nR As Long
Set rsAcc = New ADODB.Recordset
rsAcc.Open "select * from acc0203d where faclevel<0 and left(facparent,5)='00001'", Con, adOpenStatic
If Not rsAcc.EOF Then
nR = 0
If Not rsAcc.BOF Then rsAcc.MoveFirst
oCombo.Clear
oCombo.AddItem "All Suppliers"
Do While Not rsAcc.EOF
oCombo.AddItem rsAcc!facname
'oCombo.ItemData(nR) = rsAcc!faccode
rsAcc.MoveNext
Loop
End If
rsAcc.Close
End Sub

Public Sub FillArea(oGrid As MSFlexGrid, cString As String)
Dim nR As Integer
Set rsArea = New ADODB.Recordset
rsArea.Open "select * from arr0203d order by faname ", Con, adOpenStatic
If Not rsArea.EOF Then
oGrid.Rows = 2
oGrid.Clear
oGrid.FormatString = cString
nR = 1
    If Not rsArea.BOF Then rsArea.MoveFirst
    Do While Not rsArea.EOF
    oGrid.TextMatrix(nR, 0) = rsArea!faname
    oGrid.TextMatrix(nR, 1) = rsArea!ftransport
    oGrid.AddItem ""
    nR = nR + 1
    rsArea.MoveNext
    Loop
End If
rsArea.Close
Set rsArea = Nothing

End Sub


Public Sub fillBranch(oCombo As ComboBox)
Set rsBranch = New ADODB.Recordset
rsBranch.Open "select * from branch", Con, adOpenStatic
If Not rsBranch.EOF Then
If Not rsBranch.BOF Then rsBranch.MoveFirst
    oCombo.Clear
    Do While Not rsBranch.EOF
            oCombo.AddItem rsBranch!FBRANCH
    rsBranch.MoveNext
    Loop
End If
rsBranch.Close


Set rsBranch = Nothing
End Sub

Public Sub FillItem(oGrid As MSFlexGrid, nLevel, cString)
Dim nR As Integer
Set rsItem = New ADODB.Recordset
If nLevel > 0 Then
    rsItem.Open "select * from ite0203d where val(faclevel) > 0 order by facname ", Con, adOpenStatic
ElseIf nLevel < 0 Then
    rsItem.Open "select * from ite0203d where val(faclevel) < 0 order by facname ", Con, adOpenStatic
End If
If Not rsItem.EOF Then
oGrid.Rows = 2
oGrid.Clear
oGrid.FormatString = cString
nR = 1
    If Not rsItem.BOF Then rsItem.MoveFirst
    Do While Not rsItem.EOF
    oGrid.TextMatrix(nR, 0) = rsItem!facname
    oGrid.TextMatrix(nR, 1) = rsItem!faccode
    oGrid.TextMatrix(nR, 2) = rsItem!fclbal
    oGrid.TextMatrix(nR, 3) = rsItem!fSp
    
    
    oGrid.AddItem ""
    nR = nR + 1
    rsItem.MoveNext
    Loop
End If
rsItem.Close
Set rsItem = Nothing


End Sub


Public Sub Limitforform(cMnu As String, oForm As Form, nUser As Long)
Set rsMenu = New ADODB.Recordset
rsMenu.Open "select * from N_LstMnuUserCtrl where fmenuvariable='" & cMnu & "' and val(fuser)='" & nUser & "'", Con, adOpenStatic
If Not rsMenu.EOF Then
If Not rsMenu.BOF Then rsMenu.MoveFirst
oForm.mnu2.Visible = rsMenu!fnew
oForm.mnu3.Visible = rsMenu!fedit
oForm.mnu4.Visible = rsMenu!fcancel
oForm.mnu5.Visible = rsMenu!fdelete
oForm.mnu6.Visible = rsMenu!fprint

If rsMenu!fnew = False Then oForm.mnu8.Visible = False
    
End If
rsMenu.Close
Set rsMenu = Nothing


End Sub

Public Function LstTag(nTag As Integer) As String
Select Case nTag
Case Is = 0
    LstTag = "Opening"
Case Is = 1
    LstTag = "Return"
Case Is = 2
    LstTag = "Production"
Case Is = 3
    LstTag = "Damaged"
Case Is = 4
    LstTag = "Scrab"
Case Is = 5
    LstTag = "Transfer"
Case Is = 6
    LstTag = "Purchase without Bill"
Case Is = 10
    LstTag = "Sales"
Case Is = 11
    LstTag = "Purchase"
End Select

End Function

Public Sub MinusGdStk(cCode As String, nStock As Double, cGd As String)
Set rsGdSTk = New ADODB.Recordset
    rsGdSTk.Open "update gdstk set fbal=fbal-'" & nStock & "' where faccode='" & cCode & "' and fgd='" & cGd & "'", Con, adOpenDynamic, adLockPessimistic

End Sub

Public Sub StartSet()
'*****************************************************
' this procedure is to assign avialable or not
'*****************************************************
'*****************************************************
Set rsBstrt = New ADODB.Recordset
rsBstrt.Open "select * from configvoucher", Con, adOpenStatic
If Not rsBstrt.EOF Then
If Not rsBstrt.BOF Then rsBstrt.MoveFirst
    nLst = rsBstrt!fproductlst
    lSuGd = rsBstrt!fgodown
    lSuBra = rsBstrt!FBRANCH
    lSuTax = rsBstrt!ftax
    lSuBc = rsBstrt!fbc
    lSuPc = rsBstrt!fpc
    lSuDed = rsBstrt!fd
    lSuOlan = rsBstrt!fl
    lSuBar = rsBstrt!fbarcode
    nBM = rsBstrt!fbillnum
    lSuDc = rsBstrt!fdc
    lSuOn = rsBstrt!fon
    lSuItax = rsBstrt!ftaxin
    lSuSTax = rsBstrt!fsingletax

End If
rsBstrt.Close
Set rsBstrt = Nothing
End Sub
Public Sub FillUser(oGrid As MSFlexGrid, cFstring As String)
Dim nR As Integer
Set rsLogin = New ADODB.Recordset
rsLogin.Open "select * from userlogin order by loginname", Con, adOpenStatic
If Not rsLogin.EOF Then
oGrid.Rows = 2
oGrid.Clear
oGrid.FormatString = cFstring
nR = 1
If Not rsLogin.BOF Then rsLogin.MoveFirst
Do While Not rsLogin.EOF

oGrid.TextMatrix(nR, 0) = rsLogin!loginname
oGrid.TextMatrix(nR, 1) = rsLogin!fslno
oGrid.TextMatrix(nR, 2) = rsLogin!fPassword
oGrid.AddItem ""
nR = nR + 1
rsLogin.MoveNext
Loop
End If
rsLogin.Close
Set rsLogin = Nothing
End Sub

Public Function GetStock(cCode As String) As Double
Set rsItem = New ADODB.Recordset
rsItem.Open "select * from ite0203d where faccode='" & cCode & "'", Con, adOpenStatic
If Not rsItem.EOF Then
    If Not rsItem.BOF Then rsItem.MoveFirst
    GetStock = rsItem!fclbal
    End If
rsItem.Close
Set rsItem = Nothing
End Function

Public Sub AddStock(cCode As String, nStock As Double, Optional cGd As String)
Set rsItem = New ADODB.Recordset
rsItem.Open "update ite0203d set fclbal=fclbal+ '" & nStock & "' where faccode='" & cCode & "'", Con, adOpenStatic

End Sub


Public Sub IntMenu(nUser As Long)
Set rsAccCtrl = New ADODB.Recordset
rsAccCtrl.Open "select * from N_LstMnuUserCtrl where val(fuser)='" & nUser & "'", Con, adOpenStatic
If Not rsAccCtrl.EOF Then
If Not rsAccCtrl.BOF Then rsAccCtrl.MoveFirst
With frmMain
Do While Not rsAccCtrl.EOF

Select Case rsAccCtrl!fmenuvariable
   Case Is = "mnu0"
         .mnu0.Visible = rsAccCtrl!fvisible

   Case Is = "mnu1"
         .mnu1.Visible = rsAccCtrl!fvisible
   Case Is = "mnu2"
         .mnu2.Visible = rsAccCtrl!fvisible
   Case Is = "mnu3"
         .mnu3.Visible = rsAccCtrl!fvisible
   Case Is = "mnu4"
         .mnu4.Visible = rsAccCtrl!fvisible
   Case Is = "mnu5"
         .mnu5.Visible = rsAccCtrl!fvisible
   Case Is = "mnu6"
         .mnu7.Visible = rsAccCtrl!fvisible
   Case Is = "mnu8"
         .mnu8.Visible = rsAccCtrl!fvisible
   Case Is = "mnu9"
         .mnu9.Visible = rsAccCtrl!fvisible
   Case Is = "mnu10"
         .mnu10.Visible = rsAccCtrl!fvisible
   Case Is = "mnu11"
         .mnu11.Visible = rsAccCtrl!fvisible
   Case Is = "mnu12"
         .mnu12.Visible = rsAccCtrl!fvisible
   Case Is = "mnu13"
         .mnu13.Visible = rsAccCtrl!fvisible
   Case Is = "mnu14"
         .mnu14.Visible = rsAccCtrl!fvisible
   Case Is = "mnu15"
         .mnu15.Visible = rsAccCtrl!fvisible
   Case Is = "mnu16"
         .mnu16.Visible = rsAccCtrl!fvisible
   Case Is = "mnu17"  ' goods receipt issue
         .mnu17.Visible = rsAccCtrl!fvisible
   Case Is = "mnu18"
         .mnu18.Visible = rsAccCtrl!fvisible
   Case Is = "mnu19"
         .mnu19.Visible = rsAccCtrl!fvisible
   Case Is = "mnu20"
         .mnu20.Visible = rsAccCtrl!fvisible
   Case Is = "mnu21"
         .mnu21.Visible = rsAccCtrl!fvisible
   Case Is = "mnu22"
         .mnu22.Visible = rsAccCtrl!fvisible
   Case Is = "mnu23"
         .mnu23.Visible = rsAccCtrl!fvisible
   Case Is = "mnu24"
         .mnu24.Visible = rsAccCtrl!fvisible
   Case Is = "mnu25"
         .mnu25.Visible = rsAccCtrl!fvisible
   Case Is = "mnu26"
         .mnu26.Visible = rsAccCtrl!fvisible
   Case Is = "mnu27"
         .mnu27.Visible = rsAccCtrl!fvisible
   Case Is = "mnu28"
         .mnu28.Visible = rsAccCtrl!fvisible
   Case Is = "mnu29"
         .mnu29.Visible = rsAccCtrl!fvisible
   Case Is = "mnu30"
         .mnu30.Visible = rsAccCtrl!fvisible
   Case Is = "mnu31"
         .mnu31.Visible = rsAccCtrl!fvisible
   Case Is = "mnu32"
         .mnu32.Visible = rsAccCtrl!fvisible
   Case Is = "mnu33"
         .mnu33.Visible = rsAccCtrl!fvisible
   Case Is = "mnu34"
         .mnu34.Visible = rsAccCtrl!fvisible
   Case Is = "mnu35"
         .mnu35.Visible = rsAccCtrl!fvisible
   Case Is = "mnu36"
         .mnu36.Visible = rsAccCtrl!fvisible
   Case Is = "mnu37"
         .mnu37.Visible = rsAccCtrl!fvisible
   Case Is = "mnu38"
         .mnu38.Visible = rsAccCtrl!fvisible
   Case Is = "mnu39"
         .mnu39.Visible = rsAccCtrl!fvisible
   Case Is = "mnu40"
         .mnu40.Visible = rsAccCtrl!fvisible
   Case Is = "mnu41"
         .mnu41.Visible = rsAccCtrl!fvisible
   Case Is = "mnu42"
         .mnu42.Visible = rsAccCtrl!fvisible
   Case Is = "mnu43"
         .mnu43.Visible = rsAccCtrl!fvisible
   Case Is = "mnu44"
         .mnu44.Visible = rsAccCtrl!fvisible
   Case Is = "mnu45"
         .mnu45.Visible = rsAccCtrl!fvisible
   Case Is = "mnu46"
         .mnu46.Visible = rsAccCtrl!fvisible
   Case Is = "mnu47"
         .mnu47.Visible = rsAccCtrl!fvisible
   Case Is = "mnu48"
         .mnu48.Visible = rsAccCtrl!fvisible
   Case Is = "mnu49"
         .mnu49.Visible = rsAccCtrl!fvisible
   Case Is = "mnu50"
         .mnu50.Visible = rsAccCtrl!fvisible
'   Case Is = "mnu51"
'         .mnu51.Visible = rsAccCtrl!fvisible
'   Case Is = "mnu52"
'         .mnu52.Visible = rsAccCtrl!fvisible
'   Case Is = "mnu53"
'         .mnu53.Visible = rsAccCtrl!fvisible
'   Case Is = "mnu54"
'         .mnu54.Visible = rsAccCtrl!fvisible
'   Case Is = "mnu55"
'         .mnu55.Visible = rsAccCtrl!fvisible
'   Case Is = "mnu56"
'         .mnu56.Visible = rsAccCtrl!fvisible
'   Case Is = "mnu57"
'         .mnu57.Visible = rsAccCtrl!fvisible


End Select

rsAccCtrl.MoveNext
Loop
End With

End If
rsAccCtrl.Close
End Sub


Public Sub LoadMnuLst(oGrid As MSFlexGrid)
Dim nR As Integer
Set rsMenu = New ADODB.Recordset
rsMenu.Open "select * from menu order by fslno", Con, adOpenStatic
If Not rsMenu.BOF Then rsMenu.MoveFirst
nR = 1
Do While Not rsMenu.EOF
If rsMenu!fvisible = True Then
oGrid.TextMatrix(nR, 0) = rsMenu!fslno
oGrid.TextMatrix(nR, 1) = rsMenu!fmenuname

oGrid.AddItem ""
nR = nR + 1
End If
rsMenu.MoveNext
Loop
rsMenu.Close
Set rsMenu = Nothing
End Sub

Public Sub MinusStock(cCode As String, nStock As Double, Optional cGd As String)
Set rsItem = New ADODB.Recordset
rsItem.Open "update ite0203d set fclbal=fclbal- '" & nStock & "' where faccode='" & cCode & "'", Con, adOpenDynamic, adLockPessimistic

End Sub

Public Function ShowRate(cCode As String) As Double
Set rssit = New ADODB.Recordset
rssit.Open "select * from ite0203d where faccode='" & cCode & "'", Con, adOpenStatic
If Not rssit.EOF Then
    If Not rssit.BOF Then rssit.MoveFirst
    ShowRate = rssit!fSp
    End If
rssit.Close
Set rssit = Nothing
End Function



Public Sub Main()
cDbPath = App.Path
cDbFname = "winiv.mdb"
Call dbStart
StartUp
End Sub


Public Function Find(FindList As MSFlexGrid, FindStr As String, Fcol As Integer, Optional oFrm As Frame) As Boolean
FindList.Enabled = True
FindList.Visible = True
Find = False
For I = 1 To FindList.Rows - 1
    If Left(FindList.TextMatrix(I, Fcol), Len(FindStr)) = FindStr Then
'        FindList.Row = nOldRow
 '       FindList.CellBackColor = vbWhite
          Find = True
        If Not FindList.RowIsVisible(I) Then
            a = FindList.Row
            FindList.TopRow = I
            If FindList.TopRow < I Then
               FindList.Row = I
               nOldRow = FindList.TopRow
            Else
                        nOldRow = FindList.TopRow
            FindList.Row = FindList.TopRow

            End If
        Else
            FindList.Row = I
  '          nOldRow = FindList.Row
        End If
        
   '     FindList.CellBackColor = vbBlue
        Exit Function
    End If
Next
'If Not Find Then oFrm.Visible = True

End Function

Public Sub StartUp()
'******************************************
'******************************************
' This sub use for current transanction, previous transaction
' assigning current


Set rsStartUP = New ADODB.Recordset
rsStartUP.Open "select * from lastrecord", Con, adOpenDynamic, adLockPessimistic
dCdt = rsStartUP!fcdate
nDmal = rsStartUP!fd

If nDmal > 0 Then
cDP = "#####0" + "." + String(nDmal, "0")
Else
cDP = "##0"
End If
rsStartUP.Close
frmMain.MainSb.Panels(1).Text = "Date:" & dCdt

Set rsTemp = New ADODB.Recordset
rsTemp.Open "select * from template", Con, adOpenStatic
    If Not rsTemp.EOF Then
       If Not rsTemp.BOF Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
           Select Case rsTemp!ftemp
             Case Is = "CASH"
                  cCash = rsTemp!fcode
             Case Is = "BANK"
                  cBank = rsTemp!fcode
             Case Is = "SALES"
                  cSales = rsTemp!fcode
             Case Is = "PURCHASE"
                  cPurchase = rsTemp!fcode
             Case Is = "DISCOUNT"
                  cDiscount = rsTemp!fcode
             Case Is = "VAT"
                  cVat = rsTemp!fcode
             Case Is = "TRANS"
                  cTrans = rsTemp!fcode
             Case Is = "CARD"
                  cCard = rsTemp!fcode
             Case Is = "BANKCHG"
                  cBankchg = rsTemp!fcode
           End Select
        rsTemp.MoveNext
        Loop
    End If
rsTemp.Close

Set rsBStp = New ADODB.Recordset

rsBStp.Open "select * from billsetup", Con, adOpenStatic
If Not rsBStp.EOF Then
cNtDes = rsBStp!fnote
If Not IsNull(rsBStp!fhead1) Then cHead1 = rsBStp!fhead1
If Not IsNull(rsBStp!fhead2) Then cHead2 = rsBStp!fhead2
If Not IsNull(rsBStp!fhead3) Then cHead3 = rsBStp!fhead3
If Not IsNull(rsBStp!fhead4) Then cHead4 = rsBStp!fhead4
If Not IsNull(rsBStp!fhead5) Then cHead5 = rsBStp!fhead5
If Not IsNull(rsBStp!fhead6) Then cHead6 = rsBStp!fhead6
If Not IsNull(rsBStp!fbottom1) Then cBot1 = rsBStp!fbottom1
If Not IsNull(rsBStp!fbottom2) Then cBot2 = rsBStp!fbottom2
If Not IsNull(rsBStp!fbottom3) Then cBot3 = rsBStp!fbottom3
If Not IsNull(rsBStp!fbottom4) Then cBot4 = rsBStp!fbottom4
nNoP = rsBStp!fnofp  ' noproduct
nNoLR = rsBStp!fnolr  ' no of line reverse
nNolE = rsBStp!fnole   'no of line eject


End If
rsBStp.Close


Set rsTemp = Nothing


Set rsStartUP = Nothing
End Sub

Public Sub updateLog(cMnu As String, nUser As Long, nMode As Single, dTime As Date, cDesc As String)
Set rsLog = New ADODB.Recordset
rsLog.Open "select * from logfile", Con, adOpenDynamic, adLockPessimistic
rsLog.AddNew
rsLog!fuser = cUser
rsLog!fdate = dCdt
rsLog!fmenu = cMnu
rsLog!fmode = nMode
rsLog!fdesc = cDesc
rsLog!ftime = dTime
rsLog.Update
rsLog.Close
Set rsLog = Nothing
End Sub


Public Function VrType(cVoucher As String) As Integer

Select Case cVoucher
 Case Is = "SAL"
    VrType = 1
 Case Is = "PUR"
    VrType = 2
 Case Is = "REC"
    VrType = 3
 Case Is = "PAY"
    VrType = 4
 Case Is = "GREC"
    VrType = 5
 Case Is = "GISS"
    VrType = 6
End Select
End Function


