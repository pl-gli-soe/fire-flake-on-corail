Attribute VB_Name = "MakeListModule"
'The MIT License (MIT)
'
'Copyright (c) 2021 FORREST
' Mateusz Milewski mateusz.milewski@opel.com aka FORREST
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

Public Sub clearInputList(ictrl As IRibbonControl)
    
    innerClearInputList False
End Sub

Public Sub innerClearInputList(Optional b As Boolean)

    Dim ish As Worksheet
    
    
    If b = False Then
    
        Dim ans As Variant
        ans = MsgBox("Are you sure you want to clear your input list?", vbYesNo + vbExclamation)
        
        
        If ans = vbYes Then
            
            
            Set ish = ThisWorkbook.Sheets("input")
            
            '    Rows("7:52").Select
            '    Selection.Delete Shift:=xlUp
            ish.Rows("2:100000").Delete Shift:=xlUp
            
            
            MsgBox "input list is clear!", vbInformation
        End If
    Else
    
            
        Set ish = ThisWorkbook.Sheets("input")
        ish.Rows("2:100000").Delete Shift:=xlUp
        
    End If
End Sub

Public Sub makeList(ictrl As IRibbonControl)
    innerMakeList
End Sub

Public Sub innerMakeList()
Attribute innerMakeList.VB_ProcData.VB_Invoke_Func = "Q\n14"


    ThisWorkbook.Sheets(FFOC.G_SH_NM_PRE_LIST).Activate
    
    
    ' starting main logic with creating input list based on scopes from 2720
    
    ' clear plt list on form
    Make1.ListBox1.Clear
    Make1.ListBox2.Clear
    Make1.pb2.Width = 1
    
    Dim ir As Range
    ' plt-list == yellow list
    Set ir = ThisWorkbook.Sheets("plt-list").Range("A2")
    Do
        
        If (UCase(ir.Offset(0, 1).Value) Like "*CORAIL*") And (Not UCase(ir.Offset(0, 1).Value) Like "*CORAIL*MANUAL*") Then
            Make1.ListBox1.addItem ir.Value, 0
            Make1.ListBox1.List(0, 1) = Split(ir.Offset(0, 1), " ")(1)
        End If
        Set ir = ir.Offset(1, 0)
    Loop Until Trim(ir.Value) = ""
    
    Make1.show vbModeless
End Sub


Public Sub quickTestToRefreshOnEI()
    
    ' page=1&max=10&charset=UTF-8&ecwAutoId=false&ecwDivId=productSearch&shortage=false
    '   &provisionerAreaIdArray=43609&provisionerAreaLabel=PIETRASZ MAGDALENA - New flows&sgrLineDetail=false
    '   &errorState=false&ruptureStateEnum=NOT_TREATED&displayLocal=false
    
    ' page=1&max=10&charset=UTF-8&ecwAutoId=false&ecwDivId=productSearch&shortage=false
    '   &provisionerAreaIdArray=35107:35108:43913&provisionerAreaLabel=BREWCZYK PAULINA - All areas
    '   &sgrLineDetail=false&errorState=false&ruptureStateEnum=NOT_TREATED&displayLocal=false
End Sub






Public Sub quikTestOn(arg As String, ByRef frm As Make1)


    frm.ListBox2.Clear


    frm.pb2.Width = 1
    frm.StatusLabel.Caption = "Status: loading areas"
    
    Dim req As WinHttpRequest
    

    Dim doc As HTMLDocument, idoc As IHTMLDocument
    Dim element As IHTMLElement
    Dim col As HTMLElementCollection
    
    Dim tb As IHTMLTable
    
    
    
    Dim httpreqtxt As String, params As String
    Dim login As String, pass As String
    

    login = frm.TextBoxLogin.Value
    pass = frm.TextBoxPassword.Value
    
    
    ' http://ei.control.erp.corail.inetpsa.com/provisionerAreaComboAction.do?sum=true&total=true&own=true
    ' value=BLANKENBURG HEIKO - All areas&page=1&max=20&charset=UTF-8&ecwAutoId=false&ecwDivId=provisionerArea_table
    ' Request URL: http://ei.control.erp.corail.inetpsa.com/provisionerAreaComboAction.do?sum=true&total=true&own=true
    httpreqtxt = "http://" & CStr(arg) & ".control.erp.corail.inetpsa.com/provisionerAreaComboAction.do?sum=true&total=true&own=true"
    
    
    
    Dim currentPage As Long
    currentPage = 1
    
    
    Do
        params = "page=" & CStr(currentPage) & "&max=20&charset=UTF-8&ecwAutoId=false&ecwDivId=provisionerArea_table"
    
        Set doc = New HTMLDocument
        Set idoc = New HTMLDocument

        Set req = New WinHttpRequest
        
        req.Open "POST", httpreqtxt, True
        req.setRequestHeader "Accept", "*/*"
        req.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        req.SetAutoLogonPolicy AutoLogonPolicy_Always
        req.SetCredentials login, pass, 0
        req.send params
        
        ' req.WaitForResponse
        
        
        Do
            DoEvents
            incrementFrmStatusBar frm
            Sleep 100
        Loop Until req.WaitForResponse()
    
    
            
        Debug.Print Split(req.responseText, Chr(10))(0)
        frm.StatusLabel.Caption = "Status: loading areas: " & CStr(Split(req.responseText, Chr(10))(0))
        frm.Repaint
        DoEvents
        'Debug.Print Split(req.responseText, Chr(10))(1)
        'Debug.Print Split(req.responseText, Chr(10))(30)
        
        'Debug.Print req.responseText
        
        If checkIfThisIsAllPages(CStr(Split(CStr(req.responseText), Chr(10))(0))) Then
        
            doc.body.innerHTML = req.responseText
            idoc.body.innerHTML = doc.body.innerHTML
            
            Set tb = doc.getElementsByTagName("table")(0)
            
            
            ' ThisWorkbook.Sheets("test").Range("a1")
            ' fillMatrix tb, False, ThisWorkbook.Sheets("test").Range("a1"), ThisWorkbook.Sheets("test")
            fillListBoxFromMake1 frm, tb
        End If
    
        currentPage = currentPage + 1
    
    Loop While checkIfLastPage(currentPage, CStr(Split(CStr(req.responseText), Chr(10))(0)))
    
    
    
    
    
End Sub




Private Sub fillListBoxFromMake1(ByRef frm As Make1, ByRef tb As IHTMLTable)

    Application.Calculation = xlCalculationManual
    Application.enableEvents = False
    
    
    
    frm.StatusLabel.Caption = "Status: loading reference list"
    incrementFrmStatusBar frm
    
    ' -------------------------------------------------------------
    Dim r As HTMLTableRow
    Dim c As HTMLTableCell
    Dim iter As Long, ktoraKolumna As Integer
    iter = 0
    
    ' begin from second starting from zero!
    ktoraKolumna = 1
    
    ' frm.ListBox2.Clear
    
    Dim tmpStrForAttr As String


    For Each r In tb.Rows
        
        ktoraKolumna = 1
        
        On Error Resume Next
        tmpStrForAttr = r.getAttribute("ecwKeyVal0")
        
        If Trim(tmpStrForAttr) <> "" Then
            frm.ListBox2.addItem tmpStrForAttr, iter
            For Each c In r.Cells
                
                frm.ListBox2.List(iter, ktoraKolumna) = c.innerText
                
                
                ktoraKolumna = ktoraKolumna + 1
                
            Next c
        
            
            iter = iter + 1
            incrementFrmStatusBar frm
        End If
    Next r
    ' -------------------------------------------------------------
    
    Application.Calculation = xlCalculationAutomatic
    Application.enableEvents = True
End Sub

Private Sub fillMatrix(tb, avoidFirstRow, rng, Sh)


    Application.Calculation = xlCalculationManual
    Application.enableEvents = False


    

    Dim r As HTMLTableRow
    Dim c As HTMLTableCell
    Dim tmpStrForAttr As String


    For Each r In tb.Rows
    
    
        ' not really useful
        'rng.Parent.Cells(rng.Row, 30).Value = r.innerHTML
        'rng.Parent.Cells(rng.Row, 35).Value = r.innerText
        
        
        If Not avoidFirstRow Then
        
            On Error Resume Next
            tmpStrForAttr = r.getAttribute("ecwKeyVal0")
            rng.Value = tmpStrForAttr
            Set rng = rng.Offset(0, 1)
            
    
            For Each c In r.Cells
                
                rng.Value = c.innerText
                
                
                Set rng = rng.Offset(0, 1)
                
            Next c
            
            Set rng = rng.Offset(1, 0)
            Set rng = Sh.Cells(rng.Row, 1)
        
        Else
            avoidFirstRow = False
        End If
        

        
    Next r
    
    Application.Calculation = xlCalculationAutomatic
    Application.enableEvents = True

End Sub


Private Function checkIfThisIsAllPages(topLine As String) As Boolean
    
    checkIfThisIsAllPages = False
    
    
    ' <!-- indicePage="54" pageSize="39" numberOfElement="5239" numberOfPage="53"-->
    
    Dim indicePage As Long
    Dim allPages As Long
    
    indicePage = getPageNumber(topLine, "indicePage=")
    allPages = getPageNumber(topLine, "numberOfPage=")
    
    
    If indicePage <= allPages Then
        checkIfThisIsAllPages = True
    End If
    
    
    

End Function



Private Function checkIfLastPage(currentPage As Long, topLine As String) As Boolean
    
    checkIfLastPage = False
    
    
    ' <!-- indicePage="54" pageSize="39" numberOfElement="5239" numberOfPage="53"-->
    Debug.Print topLine
    
    Dim indicePage As Long
    Dim allPages As Long
    
    indicePage = getPageNumber(topLine, "indicePage=")
    allPages = getPageNumber(topLine, "numberOfPage=")
    
    
    If currentPage <= allPages Then
        checkIfLastPage = True
    End If
    
    
    

End Function



Private Function getPageNumber(tp As String, ptrn As String) As Long

    getIndicepage = 0
    arr = Split(tp, ptrn)
    tmpstr = arr(1)
    
    arr = Split(tmpstr, """")
    
    ' Debug.Print arr(0) & " ,  " & arr(1)
    
    getPageNumber = CLng(arr(1))
    

End Function



Private Sub incrementFrmStatusBar(ByRef frm As Make1)
    
    frm.pb2.Width = frm.pb2.Width + 1
    
    If frm.pb2.Width > frm.pb1.Width Then
        frm.pb2.Width = 10
    End If
End Sub





Public Sub quikGetPNsList(plt As String, areaId As String, nm As String, areaStr As String, frm As Make1, e1 As E_PRE_LIST)




    Dim myRng As Range, mySh As Worksheet
    Set myRng = ThisWorkbook.Sheets("pre-list").Range("a1")
    Set mySh = ThisWorkbook.Sheets("pre-list")
    
    If e1 = E_PRE_LIST_NEW Then
    
        
        With mySh.UsedRange
            .ClearComments
            .ClearContents
            .ClearHyperlinks
            .Value = ""
        End With
        mySh.UsedRange.Clear
    ElseIf e1 = E_PRE_LIST_ADD Then
    
    
        ' heurisitc 4th columns
    
        Dim tmprng As Range
        Set tmprng = mySh.Range("C1048576").End(xlUp).Offset(1, 0)
        Set myRng = myRng.Offset(tmprng.Row - 1, 0)
        
        Debug.Print myRng.Address
    End If
    
    
    
    ' after clearing so put labels again
    With mySh
        .Range("C1").Value = "REF"
        .Range("D1").Value = "DESC"
        .Range("E1").Value = "SELLER"
        .Range("F1").Value = "SHIPPER"
        .Range("G1").Value = "C"
        .Range("H1").Value = "NM"
        .Range("I1").Value = "SGR_LINE"
        .Range("J1").Value = "PROC"
        .Range("K1").Value = "CMJ"
        .Range("L1").Value = "SDU"
        .Range("M1").Value = "SHORT1"
        .Range("N1").Value = "SHORT2"
        .Range("O1").Value = "CMNT"
        .Range("P1").Value = "PLT"
    End With



    Dim arg As String
    arg = plt
    

    
    frm.pb2.Width = 1
    frm.StatusLabel.Caption = "Status: loading list"
    
    Dim req As WinHttpRequest
    

    Dim doc As HTMLDocument, idoc As IHTMLDocument
    Dim element As IHTMLElement
    Dim col As HTMLElementCollection
    
    Dim tb As IHTMLTable
    
    
    
    Dim httpreqtxt As String, params As String
    Dim login As String, pass As String
    
    login = frm.TextBoxLogin.Value
    pass = frm.TextBoxPassword.Value
    
    
    ' page=1&max=10&charset=UTF-8&ecwAutoId=false&ecwDivId=productSearch&shortage=false
    '   &provisionerAreaIdArray=46732&provisionerAreaLabel=BLANKENBURG HEIKO - 67SYN
    '   &sgrLineDetail=false&errorState=false&displayLocal=false
    
    ' Request URL: http://ei.control.erp.corail.inetpsa.com/getProductSearchPager.do
    ' request Method: Post
    httpreqtxt = "http://" & CStr(arg) & ".control.erp.corail.inetpsa.com/getProductSearchPager.do"
    
    
    
    Dim currentPage As Long
    currentPage = 1
    
    
    
    Do
    
    
        params = "page=" & CStr(currentPage) & "&max=20&charset=UTF-8&ecwAutoId=false&ecwDivId=productSearch&shortage=false" & _
            "&provisionerAreaIdArray=" & areaId & "&provisionerAreaLabel=" & nm & " - " & areaStr & _
            "&sgrLineDetail=false&errorState=false&displayLocal=false"
    

    
        Set doc = New HTMLDocument
        Set idoc = New HTMLDocument
    

        Set req = New WinHttpRequest
    
        req.Open "POST", httpreqtxt, True
        req.setRequestHeader "Accept", "*/*"
        req.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        req.SetAutoLogonPolicy AutoLogonPolicy_Always
        req.SetCredentials login, pass, 0
        req.send params
        
        ' req.WaitForResponse
    
    
        Do
            DoEvents
            incrementFrmStatusBar frm
            Sleep 50
        Loop Until req.WaitForResponse()
    
    
            
        Debug.Print Split(req.responseText, Chr(10))(0)
        frm.StatusLabel.Caption = "Status: loading list: " & CStr(Split(req.responseText, Chr(10))(0))
        frm.Repaint
        DoEvents
        'Debug.Print Split(req.responseText, Chr(10))(1)
        'Debug.Print Split(req.responseText, Chr(10))(30)
        
        'Debug.Print req.responseText
    
        If checkIfThisIsAllPages(CStr(Split(CStr(req.responseText), Chr(10))(0))) Then
        
            doc.body.innerHTML = req.responseText
            idoc.body.innerHTML = doc.body.innerHTML
            
            Set tb = doc.getElementsByTagName("table")(0)
            
            
            ' ThisWorkbook.Sheets("test").Range("a1")
            fillListMatrix plt, tb, True, myRng, mySh, frm
            'fillListBoxFromMake1 frm, tb
        End If
        
        currentPage = currentPage + 1
    
    Loop While checkIfLastPage(currentPage, CStr(Split(CStr(req.responseText), Chr(10))(0)))
    
    
    
    ' frm.StatusLabel.Caption = "Status: loading list ready!"
    
End Sub





Private Sub fillListMatrix(plt As String, tb As IHTMLTable, avoidFirstRow, rng As Range, Sh As Worksheet, frm As Make1)


    Application.Calculation = xlCalculationManual
    Application.enableEvents = False
    
    
    
    ' in this sub e1 doesn;t make any sense
    ' I need to add - I cant just replace here - i will overwrite prev page!
    ' ------------------------------------------------------------------------
    'If e1 = E_PRE_LIST_NEW Then
    '    sh.UsedRange.ClearComments
    '    sh.UsedRange.ClearContents
    '    sh.UsedRange.ClearFormats
    '    sh.UsedRange.Value = ""
    '    sh.UsedRange.Clear
    'ElseIf e1 = E_PRE_LIST_ADD Then
    '
    '    ' heurisitc 4th columns
    '
    '    Dim tmprng As Range
    '    Set tmprng = sh.Range("D1048576").End(xlUp).Offset(1, 0)
    '    Set rng = rng.Offset(tmprng.Row - 1, 0)
    'End If
    ' ------------------------------------------------------------------------

    Dim tmprng As Range
    Set tmprng = Sh.Range("C1048576").End(xlUp).Offset(1, 0)
    ' return to A column
    Set rng = tmprng.Offset(0, -2)
    
    Debug.Print "rng.Address: " & rng.Address
    

    Dim r As HTMLTableRow
    Dim c As HTMLTableCell
    Dim tmpStrForAttr As String


    For Each r In tb.Rows
    
    
        '<tr ecwkeyval0="" ecwkeyval1="98262472XT" ecwkeyval2="Chine" ecwkeyval3="272922" ecwkeyval4="BREWCZYK PAULINA" ecwkeyval5="" ecwkeyval6="" ecwkeyval7="0.0" ecwkeyval8="0.0" ecwkeyval9="" ecwkeyval10="" ecwkeyval11="" ecwkeyval12="" ecwkeyval13="BITRON" ecwkeyval14="BITRON" ecwkeyval15="24/02/2021" ecwkeyval16="oszhm3j" ecwkeyval17="" class=" ecwTableAlternateLine1" onmouseover="javascript:ecwTableV2LigneOnMouseOver(this, event);" onmouseout="javascript:ecwTableV2LigneOnMouseOut(this, event);" onclick="javascript:ecwTableV2LigneOnClick(this, event);">
        '<td><table cellspacing="0" cellpadding="0" border="0" class="ecwTableListeAction"><tbody><tr><td><span class="ecwConsult" onclick="javascript:window.ecwTableV2CallClickAction(this);" ecwonclick="javascript:fn();" title="Consult"></span></td></tr></tbody></table></td>
        '<td ecwkeyname="warningProduct"><div class="text emptyStyle">&nbsp;</div></td>
        '<td ecwkeyname="productCodeKey"><div class="text">98262472XT</div></td>
        '<td><div class="text">COMMUTATEUR CHAR</div></td>
        '<td><div class="text">A001WV  01</div></td>
        '<td><div class="text">A001WV  01</div></td>
        '<td ecwkeyname="shipperCountryName">Chine</td>
        '<td ecwkeyname="approInCharge">BREWCZYK PAULINA</td>
        '<td ecwkeyname="locationKey"><div class="text">&nbsp;</div></td>
        '<td><div class="text">Not treated</div></td>
        '<td ecwkeyname="cmjJLHebdo">0.0</td>
        '<td ecwkeyname="stockDepartAbsolute">0.0</td>
        '<td ecwkeyname="formattedAbsoluteDisruptionDate"><div class="text">&nbsp;</div></td>
        '<td ecwkeyname="formattedRelativeDisruptionDate"><div class="text">&nbsp;</div></td>
        '<td ecwkeyname="shortComment" id="<table> <tr valign=&quot;top&quot;>  <td>BITRON  </td> </tr> <tr valign=&quot;top&quot;>  <td>[24/02/2021 - oszhm3j]  </td> </tr></table>">BITRON</td>
        '</tr>
        
        
        If Not avoidFirstRow Then
        
            'On Error Resume Next
            'tmpStrForAttr = r.getAttribute("ecwKeyVal0")
            'rng.Value = tmpStrForAttr
            'Set rng = rng.Offset(0, 1)
            
    
            For Each c In r.Cells
                
                rng.Value = c.innerText
                Set rng = rng.Offset(0, 1)
                
            Next c
            
            rng.Value = plt
            
            Set rng = rng.Offset(1, 0)
            Set rng = Sh.Cells(rng.Row, 1)
        
        Else
            avoidFirstRow = False
        End If
        

        incrementFrmStatusBar frm
        
    Next r
    
    Application.Calculation = xlCalculationAutomatic
    Application.enableEvents = True

End Sub














Public Sub quikTestOnEI()


    
    Dim req As WinHttpRequest
    

    Dim doc As HTMLDocument, idoc As IHTMLDocument
    Dim element As IHTMLElement
    Dim col As HTMLElementCollection
    
    Dim tb As IHTMLTable
    
    
    
    Dim httpreqtxt As String, params As String
    Dim login As String, pass As String
    
    login = ""
    pass = ""
    
    
    ' http://ei.control.erp.corail.inetpsa.com/provisionerAreaComboAction.do?sum=true&total=true&own=true
    ' value=BLANKENBURG HEIKO - All areas&page=1&max=20&charset=UTF-8&ecwAutoId=false&ecwDivId=provisionerArea_table
    ' Request URL: http://ei.control.erp.corail.inetpsa.com/provisionerAreaComboAction.do?sum=true&total=true&own=true
    httpreqtxt = "http://ei.control.erp.corail.inetpsa.com/provisionerAreaComboAction.do?sum=true&total=true&own=true"
    
    
    ' Debug.Print "req: " & httpreqtxt & " params: " & params
    
    '
    'With request2510
    '    .Open "POST", url, False
    '    .setRequestHeader "Accept", "*/*"
    '    ' .setRequestHeader "Origin", "http://sx.control.erp.corail.inetpsa.com"
    '    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    '    .SetAutoLogonPolicy AutoLogonPolicy_Always
    '    .SetCredentials login, pass, 0
    '    .send args
    '    .WaitForResponse
    '
    '    httpPost = CStr(.responseText)
    '
    '    ' Debug.Print httpPost
    'End With
    
    Dim currentPage As Long
    currentPage = 1
    
    Do
        
        params = "page=" & CStr(currentPage) & "&max=100&charset=UTF-8&ecwAutoId=false&ecwDivId=provisionerArea_table"

        Set doc = New HTMLDocument
        Set idoc = New HTMLDocument
        
        Set req = New WinHttpRequest
        req.Open "POST", httpreqtxt, False
        req.setRequestHeader "Accept", "*/*"
        req.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        req.SetAutoLogonPolicy AutoLogonPolicy_Always
        req.SetCredentials login, pass, 0
        req.send params
            
        Debug.Print Split(req.responseText, Chr(10))(0)
        'Debug.Print Split(req.responseText, Chr(10))(1)
        'Debug.Print Split(req.responseText, Chr(10))(30)
        
        'Debug.Print req.responseText
        
        If checkIfThisIsAllPages(CStr(Split(CStr(req.responseText), Chr(10))(0))) Then
        
            doc.body.innerHTML = req.responseText
            idoc.body.innerHTML = doc.body.innerHTML
            
            Set tb = doc.getElementsByTagName("table")(0)
            
            
            ' ThisWorkbook.Sheets("test").Range("a1")
            fillMatrix tb, True, ThisWorkbook.Sheets("test").Range("a1"), ThisWorkbook.Sheets("test")
        End If
        
        currentPage = currentPage + 1
    
    Loop While checkIfLastPage(currentPage, CStr(Split(CStr(req.responseText), Chr(10))(0)))
        
    
End Sub
