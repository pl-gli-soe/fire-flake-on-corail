VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HTTPRequestHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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




Private request2720 As WinHttpRequest
Private request2510 As WinHttpRequest

Private authUrl As String
Private url As String
Private urlFor2510 As String
Private pltUrl As String
'----------------added------------------
Private partNr As String
Private plant As String
'----------------------------------
Public dom As DOMHandler

Private dom2510 As DOM2510Handler
Public collectionOfDoms2510 As Collection

'-----added----
Private collectonOfXtraCorailItems As CorailIteration
'-----added----

Private login As String
Private pass As String

Private currentPartNumber As String

'corail_2720_screen:

Private num As Integer

'----------------added------------------
Private preUrl As String
' with this ko stuff request http
Private koUrl As String
Private authUrlExtendedDays As String
Private baseUrl As String
Private kodzik As String
Private screenUrl As String
Private moreRqmsUrl As String
Private urlExtendedDays As String
Private oldPreUrl As String

Private req As HTTPRequestHandler

'----------------------------------



Public Function httpPost(url As Variant, args As Variant, login As String, pass As String, e As E_TYPE_OF_CORAIL) As String


    ' Debug.Print "httpPost: args: " & args
    Dim AuthCredentials As String
    AuthCredentials = "Basic " & Base64Encode(login & ":" & pass)

    With request2510
        
        If authUrl <> "" Then
        
            .Open "GET", authUrl, False
            Debug.Print authUrl
            
            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            .SetAutoLogonPolicy AutoLogonPolicy_Always
            .setRequestHeader "Authorization", AuthCredentials
            .SetCredentials login, pass, 0
            
            .send
            DoEvents
            .WaitForResponse
            
        End If
        ' -----------------------------------------------------------------------------
    
    
        .Open "POST", url, False
        .setRequestHeader "Accept", "*/*"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .SetAutoLogonPolicy AutoLogonPolicy_Always
        .setRequestHeader "Authorization", AuthCredentials
        
        ' Debug.Print args ' OK
        
        .send args
        .WaitForResponse 3000
        
        httpPost = iterationOfgetData(CStr(.responseText), 3, CStr(url), CStr(login), CStr(pass), "POST", args)
        
        ' Debug.Print Len(httpPost) ' usually around 30k
        
        
        ' Debug.Print httpPost
        
        'Debug.Print Split(httpPost, Chr(10))(6) ' first label
        'Debug.Print Split(httpPost, Chr(10))(32) ' content part number
    End With
End Function


Private Function iterationOfgetData(odp As String, times As Integer, url As String, login As String, pass As String, Optional POST_FLAG As String, Optional args As Variant) As String


    Dim tmpArr As Variant, tmpStr1 As String

    
    tmpArr = Split(odp, Chr(10))
    G_TEST_ARR_FOR_RESPONSE = tmpArr
    
    If POST_FLAG = "POST" Then tmpStr1 = tmpArr(0)
    
    
    If times > 0 Then
    
    
        If POST_FLAG = "POST" Then
            
            If tmpStr1 Like "*indicePage=*pageSize=*numberOfElement=*" Then
                iterationOfgetData = CStr(odp)
            Else
                iterationOfgetData = "" ' tryToGetDataAgain(odp, times - 1, url, login, pass, POST_FLAG, args)
            End If
        
        Else
    
        
            If UBound(tmpArr) > 500 Then
                iterationOfgetData = CStr(odp)
            Else
                iterationOfgetData = tryToGetDataAgain(odp, times - 1, url, login, pass, POST_FLAG, args)
            End If
        End If
    
    Else
        iterationOfgetData = CStr(odp)
    End If

End Function



Private Function tryToGetDataAgain(odp As String, times As Integer, url As String, login As String, pass As String, Optional POST_FLAG As String, Optional args As Variant) As String

    tryToGetDataAgain = ""
    
    
    Dim AuthCredentials As String
    AuthCredentials = "Basic " & Base64Encode(login & ":" & pass)
    
    
    
    ' blue and 2720
    If POST_FLAG = "" Then
        With request2720
            .Open "GET", url, False
            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            .SetAutoLogonPolicy AutoLogonPolicy_Always
            .setRequestHeader "Authorization", AuthCredentials
            ' .SetCredentials login, pass, 0
            .send
            
            DoEvents
            .WaitForResponse
            
            tryToGetDataAgain = iterationOfgetData(CStr(.responseText), times, url, login, pass, "")
        End With
    ElseIf POST_FLAG = "POST" Then
    
        With request2510
            .Open "POST", url, False
            .setRequestHeader "Accept", "*/*"
            ' .setRequestHeader "Origin", "http://sx.control.erp.corail.inetpsa.com"
            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            .SetAutoLogonPolicy AutoLogonPolicy_Always
            .setRequestHeader "Authorization", AuthCredentials
            ' .SetCredentials login, pass, 0
            .send args
            .WaitForResponse 3000
            
            tryToGetDataAgain = iterationOfgetData(CStr(.responseText), times, url, login, pass, "POST")
            
            ' Debug.Print httpPost
        End With
    End If
    
End Function



Private Function getData(url As String, login As String, pass As String, e As E_TYPE_OF_CORAIL) As String



    ' Debug.Print url
    'Debug.Print login
    'Debug.Print pass
    getData = ""
    

    If url <> "" Then
    
    
        If e = BLUE Then
    
            With request2720
            
            
                'url = "http://vh.erp.corail.inetpsa.com/auth/oauth/customauthorize?client_id=corailClientId&redirect_uri=http://vh.erp.corail.inetpsa.com/&response_type=code&state=FTcNmg"
                'url = "http://vh.erp.corail.inetpsa.com/?corailsecuritycode=aa55tE&corailsecuritystate=FTcNmg"
                'url = "http://vh.erp.corail.inetpsa.com/ko5/getProductSummaryRead.do?beanId=1400018480"
                ' url = "http://vh.erp.corail.inetpsa.com/ko5/auth/oauth/customauthorize/"
                ' url = "http://ta.control.erp.corail.inetpsa.com/getProductSummaryRead.do?beanId=7552160880"
                
                ' http://vh.erp.corail.inetpsa.com/ko5/getProductSummaryRead.do?beanId=1400018480&corailsecuritycode=5iSG3x&corailsecuritystate=JkAzi7
                ' http://vh.erp.corail.inetpsa.com/auth/oauth/customauthorize?
                '   client_id=corailClientId&redirect_uri=http://vh.erp.corail.inetpsa.com/ko5/getProductSummaryRead.do?beanId%3D1400018480&response_type=code&state=JkAzi7
                
                ' url = "http://vh.erp.corail.inetpsa.com/auth/oauth/customauthorize?client_id=corailClientId&redirect_uri=http://vh.erp.corail.inetpsa.com/&response_type=code&state=7JjvnG"
                
                
                Dim AuthCredentials As String
                AuthCredentials = "Basic " & Base64Encode(login & ":" & pass)
                
                If authUrl = "" Then
                Else
                
                    .Open "GET", authUrl, False
                    Debug.Print authUrl
                    
                    ' .setRequestHeader "Content-Type", "text/json"
                    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                    ' .setRequestHeader "User-Agent", "Mozilla/5.0 (iPad; U; CPU OS 3_2_1 like Mac OS X; en-us) AppleWebKit/531.21.10 (KHTML, like Gecko) Mobile/7B405"
                    
                    

                    ' AuthCredentials = "Bearer " & login & ":" & pass
                    'AuthCredentials = "Bearer " & "token"
                    .setRequestHeader "Authorization", AuthCredentials
                    
                    .SetAutoLogonPolicy AutoLogonPolicy_Always
                    
                    .SetCredentials login, pass, 0
                    
                    
                    .send
                    
                    DoEvents
                    .WaitForResponse
                    
                    Debug.Print .getAllResponseHeaders
                    ' Debug.Print .getResponseHeader("Location")
                    ' Debug.Print .responseText
                End If
                
                
                ' url = "http://vh.erp.corail.inetpsa.com/ko5/getProductSummaryRead.do?beanId=1400018480"
                
                .Open "GET", url, False
                Debug.Print url
                
                ' .setRequestHeader "Content-Type", "text/json"
                .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                ' .setRequestHeader "User-Agent", "Mozilla/5.0 (iPad; U; CPU OS 3_2_1 like Mac OS X; en-us) AppleWebKit/531.21.10 (KHTML, like Gecko) Mobile/7B405"
                
                
                ' Dim AuthCredentials As String
                ' AuthCredentials = "Basic " & Base64Encode(login & ":" & pass)
                ' AuthCredentials = "Bearer " & login & ":" & pass
                'AuthCredentials = "Bearer " & "token"
                .setRequestHeader "Authorization", AuthCredentials
                
                .SetAutoLogonPolicy AutoLogonPolicy_Always
                
                .SetCredentials login, pass, 0
                
                
                .send
                
                DoEvents
                .WaitForResponse
                
                
                
                ' Debug.Print .responseText
                
                
                getData = iterationOfgetData(CStr(.responseText), 5, url, login, pass, "")
                
                ' Debug.Print CStr(.responseText)

                
            End With
        ElseIf e = MAESTRO Then
            
            With request2720
            
                '
                startUrl = "http://maestro.inetpsa.com/authentifier.do"
                .Open "GET", startUrl, False
                .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                .SetAutoLogonPolicy AutoLogonPolicy_Always
                .SetCredentials login, pass, 0
                .send
                
                
                Debug.Print .statusText
            
                'getData = CStr(.responseText)
                'Debug.Print getData
                
                ' pltUrl = "http://maestro.inetpsa.com/selectCentre.do?methode=select&selectedCentre=88+"
                ' pltUrl = "http://maestro.inetpsa.com/selectCentre.do?methode=select&selectedCentre=88+"
                .Open "GET", pltUrl, False
                .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                .SetAutoLogonPolicy AutoLogonPolicy_Always
                .SetCredentials login, pass, 0
                .send
                
                
                Debug.Print .statusText
            
                'getData = CStr(.responseText)
                'Debug.Print getData
                
                
                
                ' url = "http://maestro.inetpsa.com/produit.do?methode=init&selectedcodeProduit=9807026780"
                .Open "GET", url, False
                .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                .SetAutoLogonPolicy AutoLogonPolicy_Always
                .SetCredentials login, pass, 0
                .send
                
                
                Debug.Print .statusText
                
                getData = CStr(.responseText)
                
            End With
        End If
    
    Else
        getData = ""
    End If
End Function


Public Sub init(plt As String, pn As String)

    url = ""
    urlFor2510 = ""
    plant = plt
    partNr = pn
    
    
    ' CORAIL
    ' If Trim(UCase(plt)) = "Y" Then url = "http://ty.control.erp.corail.inetpsa.com/getProductSummaryRead.do?beanId=" & CStr(pn)
    
    
    ' MAESTRO
    ' url = "http://maestro.inetpsa.com/produit.do?methode=init&selectedcodeProduit=9807026780"
    
    Dim r As Range, cde As String
    Set r = ThisWorkbook.Sheets(FFOC.G_SH_NM_PLT_LIST).Range("A2")
    Do

        If UCase(Trim(plt)) = UCase(Trim(r.Value)) Then
            
            
            If Trim(r.Offset(0, 3).Value) = "MAESTRO" Then
                url = "http://maestro.inetpsa.com/produit.do?methode=init&selectedcodeProduit=" & CStr(pn)
                pltUrl = Trim(r.Offset(0, 2).Value)
            Else
                url = Trim(r.Offset(0, 4).Value) & CStr(FFOC.G_URL_EXT) & CStr(pn)
                
                ' very first page! - you need to download rest later!
                urlFor2510 = Trim(r.Offset(0, 4).Value) & CStr(FFOC.G_URL_EXT_FOR_2510) ' & CStr(makeSomeParams(pn, 1, 40)
                ' changed my mind - i will just copy pn to have ref in braceWIthDom sub ooo...
                currentPartNumber = CStr(pn)
                
                
                
                ' xtra auth protocol
                If Trim(r.Offset(0, 2).Value) Like "http://??.commons.erp.corail.inetpsa.com/" Then
                    authUrl = ""
                ElseIf Trim(r.Offset(0, 2).Value) Like "http://??.erp.corail.inetpsa.com/" Then
                
                    ' auth url
                    ' url = "http://vh.erp.corail.inetpsa.com/auth/oauth/customauthorize?client_id=corailClientId&redirect_uri=http://vh.erp.corail.inetpsa.com/&response_type=code&state=7JjvnG"
                    ' mZvCUl
                    ' 7JjvnG
                    ' 3Vzonj
                    ' V8WxMs
                    ' sJ11Gk
                    cde = "7JjvnG"
                    authUrl = Trim(r.Offset(0, 2).Value) & _
                        "auth/oauth/customauthorize?client_id=corailClientId" & "&redirect_uri=" & Trim(r.Offset(0, 2).Value) & "&response_type=code&state=" & CStr(cde)
                        
                Else
                    authUrl = ""
                End If
                
                pltUrl = ""
            End If
            Exit Do
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""

    ' Debug.Print url
End Sub


Private Function makeSomeParams(pn As String, ktoraStrona As Integer, ileLiniiNaStronie As Integer) As String


    Dim d1 As Date, d2 As Date, d_params As String, stringForForecastParam As String
    
    stringForForecastParam = "&forecast=true"
    
    d_params = makeDatesQuickly(d1, d2)

    makeSomeParams = "" & "page=" & ktoraStrona & _
                "&max=" & ileLiniiNaStronie & _
                "&charset=UTF-8" & _
                stringForForecastParam & _
                "&displayLocal=false" & _
                d_params & _
                "&productCode=" & CStr(pn)
                
    'Debug.Print makeSomeParams
End Function


Private Function makeDatesQuickly(ByRef md1 As Date, ByRef md2 As Date) As String

    makeDatesQuickly = ""


    Dim yyyy As String, mm As String, dd As String
    
    
    Dim d1 As Date, d2 As Date
    ' bardzo heurystycznie!
    d1 = Date - 10
    d2 = Date + 60
    
        
    yyyy = "" & Year(d1)
        
    mm = "" & Month(d1)
    If Len(mm) = 1 Then
        mm = "0" & mm
    End If
        
    dd = "" & Day(d1)
    If Len(dd) = 1 Then
        dd = "0" & dd
    End If
        
        
    makeDatesQuickly = "&startDate=" & yyyy & "-" & mm & "-" & dd & "T22:00:00Z"
        
    
    
    yyyy = "" & Year(d2)
    
    mm = "" & Month(d2)
    If Len(mm) = 1 Then
        mm = "0" & mm
    End If
    
    dd = "" & Day(d2)
    If Len(dd) = 1 Then
        dd = "0" & dd
    End If
        
        
    makeDatesQuickly = makeDatesQuickly & "&endDate=" & yyyy & "-" & mm & "-" & dd & "T22:00:00Z"
        
End Function



Private Function stillSomethingToParse(i_html As String, ktoraStrona As Integer) As Boolean


    Dim currentPage As Integer, pages As Integer, answer As String
    
    Dim pagesBool As Boolean, currentPageBool As Boolean
    pagesBool = False
    currentPageBool = False
    
    answer = i_html
    
    
    stillSomethingToParse = False
    
    
    firstLine = Split(answer, Chr(10))(0)
    
    arr = Split(firstLine, " ")
    
    
    Dim newVal As Long
    
    For x = LBound(arr) To UBound(arr)
        ' Debug.Print arr(x)
        
        If CStr(arr(x)) Like "numberOfElement=*" Then
        
            tmp = Split(arr(x), "=")(1)
            tmp = Replace(tmp, Chr(34), "")
            newVal = CLng(tmp)
            numberOfElements = newVal
            ' Exit For
        End If
        
        If CStr(arr(x)) Like "pageSize=*" Then
        
            tmp = Split(arr(x), "=")(1)
            tmp = Replace(tmp, Chr(34), "")
            newVal = Int(tmp)
            currentPageSize = newVal
            ' Exit For
        End If
        
        If CStr(arr(x)) Like "indicePage=*" Then
        
            tmp = Split(arr(x), "=")(1)
            tmp = Replace(tmp, Chr(34), "")
            newVal = Int(tmp)
            currentPage = newVal
            currentPageBool = True
            ' Exit For
        End If
        
        If CStr(arr(x)) Like "numberOfPage=*" Then
            tmp = Split(arr(x), "=")(1)
            tmp = Replace(tmp, Chr(34), "")
            tmp = Replace(tmp, "-->", "")
            newVal = Int(tmp)
            pages = newVal
            pagesBool = True
        End If
    Next x
    
    If pagesBool And currentPageBool Then
        If currentPage < pages Then
            stillSomethingToParse = True
        Else
            stillSomethingToParse = False
        End If
    Else
        stillSomethingToParse = False
    End If
End Function

'added----------------------------------------
'returnig extended horizon col
Public Function getCollectonOfXtraCorailItems() As CorailIteration
    Set getCollectonOfXtraCorailItems = collectonOfXtraCorailItems
End Function
'Private Function getRequestedDomExtendedHorizon()
'    Set getRequestedDomExtendedHorizon = collectonOfXtraCorailItems
'End Function
'added---------------------------------------







'---------------------------------------------

Public Sub braceWithDom(e As E_TYPE_OF_CORAIL)

    
    With dom
    
    
        'Set .doc.body.innerText = getData(url, login, pass)
        'Set .idoc.body.innerText = getData(url, login, pass)
        'dodac dwa elemnty html doc dla zacignietego wygladu podstawoweggo i dla week cov
        Set .doc = New HTMLDocument
        Set .idoc = New HTMLDocument
        Set .extdoc = New HTMLDocument
       
        'Set .iextdoc = New HTMLDocument
        
        .doc.body.innerHTML = getData(url, login, pass, e)
        .extdoc.body.innerHTML = getExtendedCovData("OAUTH", plant, partNr)
        .idoc.body.innerHTML = .doc.body.innerHTML
        

        If e = BLUE Then
        
            ' MsgBox .doc.body.innerHTML
                
            Set .stockInputElement = .doc.getElementById("stockDepart")
            Set .recvInputElement = .doc.getElementById("receivedOnSite")
            Set .backlogInputElement = .doc.getElementById("efXpLate")
            Set .extendedDays = .extdoc.getElementById("d")
            
            ' hazards
            Set .hazardsInputElement = .doc.getElementById("aleatoriesSinceFJU")
            
            
            ' new iBlockagesInProgress - return nothing if html element do not recognized
            On Error Resume Next
            Set .iBlockagesInProgress = pGetBlockagesInProgress()
            
            
            On Error Resume Next
            Set .ipartnameElement = .doc.getElementsByClassName("descriptionProduct")(0).getElementsByTagName("a")(0)
            
            
            On Error Resume Next
            Set .isupplierElement = .doc.getElementById("tableauFluxDePiece")
            ' too static
            ' Set .isupplierElement = .isupplierElement.getElementsByTagName("table")(3).getElementsByTagName("td")(0)
            ' loop and take proper data:
            'Debug.Print .isupplierElement.getElementsByTagName("table")(0).innerHTML
            'Debug.Print .isupplierElement.getElementsByTagName("table")(1).innerHTML
            'Debug.Print .isupplierElement.getElementsByTagName("table")(2).getElementsByTagName("td")(0).outerHTML
            'Debug.Print .isupplierElement.getElementsByTagName("table")(3).getElementsByTagName("td")(0).outerHTML
            
            Dim line As Integer, i As Integer, tfdp As IHTMLTable, tbIter As Integer
            tbIter = 0
            Do
                Set tfdp = Nothing
                On Error Resume Next
                Set tfdp = .isupplierElement.getElementsByTagName("table")(tbIter)
                
                
                If Not tfdp Is Nothing Then
                
                
                    line = 1
                    For Each r In tfdp.Rows
                        i = 1
                        For Each c In r.Cells
                        
                            Debug.Print c.outerHTML
                        
                            If UCase(c.outerHTML) Like "*FONT-WEIGHT: BOLD; TEXT-ALIGN: LEFT*:*" Then
                                Set .isupplierElement = c
                                Exit Do
                            End If
                        Next c
                    Next r
                    tbIter = tbIter + 1
                End If
            Loop Until tfdp Is Nothing
            
            
            
            
                
            'Debug.Print .ipartnameElement.innerText
            'Debug.Print .isupplierElement.innerText
            
            On Error Resume Next
            Set .iapproElement = .doc.getElementById("tableauIdentificationProduct").getElementsByTagName("table")(0).getElementsByTagName("td")(1)
            Set .iapproElement = .iapproElement.getElementsByClassName("INPUT_text_Readonly_Actif INPUT_text_Readonly number")(0)
            Set .iapproElement = .iapproElement.getElementsByTagName("input")(0)
            
            'Debug.Print .iapproElement.outerHTML
            'Debug.Print .iapproElement.innerHTML
            'Debug.Print .iapproElement.Value ' OK
            
            
            'Parse Data From extended Horizon
            Set theParser = New ParserHandler
            
            theParser.importPackageOfData dom.extdoc
    
          
            Set collectonOfXtraCorailItems = theParser.getConvertedDataSuitableForExcel()
            
            
            With theDomHandler
                .doc.body.innerHTML = ""
                .idoc.body.innerHTML = ""
            
                Set .doc = Nothing
                Set .idoc = Nothing
            End With
            
            Set theDomHandler = Nothing
            
            
        ElseIf e = MAESTRO Then
            
                
        
            
            
            Set .stockInputElement = Nothing
            Set .recvInputElement = Nothing
            Set .backlogInputElement = Nothing
            
            Set .istockElement = tryGetDataFromMaestro2("stock")
            Set .ibacklogElement = tryGetDataFromMaestro("backlog")
            
        ElseIf e = MANUAL Then
        
            Set .stockInputElement = Nothing
            Set .recvInputElement = Nothing
            Set .backlogInputElement = Nothing
        End If
        
        
        
    End With
    
    
    
    ' Dim specialTable As IHTMLTable
    
    
    If e = BLUE And GlobalModule.IS2510REQ = True Then
        
        Dim ktoraStrona As Integer
        ktoraStrona = 1
            
            
        Set collectionOfDoms2510 = New Collection
        Do
            
            Set dom2510 = New DOM2510Handler
            With dom2510
                Set .doc = New HTMLDocument
                Set .idoc = New HTMLDocument
                
                
                DoEvents
                
                
                ' without o auth will not work! - new implementation req.
                .rawString = httpPost(urlFor2510, _
                    makeSomeParams(currentPartNumber, ktoraStrona, 40), _
                    login, pass, e)
                    
                    
                DoEvents
                
                
                
                If .rawString <> "" Then
                
                    .doc.body.innerHTML = .rawString
                    .idoc.body.innerHTML = .doc.body.innerHTML
                    
                    If thereIsEmptyListHear(.doc.getElementsByTagName("table")(0)) Then
                    Else
                    
                    
                        'Set specialTable = Nothing
                        'On Error Resume Next
                        'Set specialTable = .doc.getElementsByTagName("table")(0)
                        
                        ' Debug.Print .table.innerHTML
                        ' Debug.Print "page: " & ktoraStrona & ", .table.Rows.Length = " & .table.Rows.Length
                        
                        
                        collectionOfDoms2510.Add .doc
                    End If
                End If
                    
                ktoraStrona = ktoraStrona + 1
                    
                
            End With
                
        Loop While stillSomethingToParse(dom2510.rawString, ktoraStrona)
    End If
End Sub
Private Function thereIsEmptyListHear(t As IHTMLTable) As Boolean



    thereIsEmptyListHear = False


    ' Debug.Print t.outerHTML
    Dim line As Integer, i As Integer
    line = 1
    For Each r In t.Rows
    
        i = 1
        For Each c In r.Cells
            If line = 2 Then
                
                ' Debug.Print c.innerHTML
                If i = 2 Then
                    If c.innerHTML Like "*&nbsp;*" Then
                        thereIsEmptyListHear = True
                    End If
                End If
                
                If i = 3 Then
                    If c.innerHTML Like "*&nbsp;*" Then
                        thereIsEmptyListHear = thereIsEmptyListHear And True
                    End If
                End If
                
                Exit For
                
                i = i + 1
            End If
        Next c
        
        line = line + 1
    Next r
        
   

    
    
End Function



Private Function pGetBlockagesInProgress()


    Set pGetBlockagesInProgress = Nothing

    On Error Resume Next
    Set dom.iElementCollection = dom.doc.getElementsByName("currentBlockingCount")
    
    If Not com.iElementCollection Is Nothing Then
        For Each el In dom.iElementCollection
        
            Set dom.iElement = el
            If Not dom.iElement Is Nothing Then
                If dom.iElement.outerHTML Like "*currentBlockingCount*" Then
                    Set pGetBlockagesInProgress = dom.iElement
                    Exit For
                End If
            End If
            
            
        Next el
               
    
    End If
End Function


Private Sub Class_Initialize()


    'login = "ofzq0l2"
    'pass = pass
    
    'login = CStr(ThisWorkbook.Sheets("register").Range("H4").Value)
    'pass = CStr(ThisWorkbook.Sheets("register").Range("H5").Value)
    
    login = G_LOGIN
    pass = G_PASS
    


    Set dom = New DOMHandler
    Set collectionOfDoms2510 = New Collection

    Set request2720 = New WinHttpRequest
    
    'With request2720
    '    .Option(4) = 13056
    '    .Option(12) = True
    'End With
    
    Set request2510 = New WinHttpRequest
    
    'With request2510
    '    .Option(4) = 13056
    '    .Option(12) = True
    'End With
End Sub


Private Sub Class_Terminate()
    Set dom = Nothing
    Set request2720 = Nothing
    Set request2510 = Nothing
    
    Set collectionOfDoms2510 = Nothing
End Sub


Private Function tryGetDataFromMaestro2(what As String) As IHTMLElement
    Set tryGetDataFromMaestro2 = Nothing
    
    With dom
        
        If what = "stock" Then
            ' Set .iElementCollection = .doc.getElementsByClassName("ligneTableauSerre")
            Set .iElement = .doc.getElementsByClassName("ligneTableauSerre")(0)
            
            On Error Resume Next
            Set .iElementCollection = .iElement.getElementsByTagName("td")(0).getElementsByTagName("td")
            
            If Not .iElementCollection Is Nothing Then
            
                podlicz = 0
                For Each el In .iElementCollection
                
                    Set .iElement = el
                    If Not .iElement Is Nothing Then
                        'Debug.Print iElement.innerHTML
                        'Debug.Print iElement.outerHTML
                        
                        If podlicz = 1 Then
                            
                            If .iElement.innerHTML Like "*FONT*" Then
                            
                                Set .iElement = .iElement.getElementsByTagName("font")(0)
                                ' Debug.Print .iElement.innerHTML ' oK
                                Set tryGetDataFromMaestro2 = .iElement
                                
                                
                            Else
                            
                                Set tryGetDataFromMaestro2 = .iElement
                            End If
                        End If
                    End If
                    
                    podlicz = podlicz + 1
                Next el
            End If
        End If
    End With
End Function

Private Function tryGetDataFromMaestro(what As String) As IHTMLElement

    Set tryGetDataFromMaestro = Nothing
    
    With dom
                
        If what = "stock" Then
            Set .iElementCollection = .doc.getElementsByTagName("font")
            
        End If
        If what = "backlog" Then
            Set .iElementCollection = .doc.getElementsByClassName("Donnee")
        End If
        
        podlicz = 0
        For Each el In .iElementCollection
            
            Set iElement = el
            If Not iElement Is Nothing Then
                Debug.Print iElement.innerHTML
                Debug.Print iElement.outerHTML
                
                If podlicz = 1 Then
                
                    If what = "stock" Then
                        Set tryGetDataFromMaestro = iElement
                        Exit Function
                    End If
                End If
                
                
                If what = "backlog" Then
                
                    
                    ' Debug.Print iElement.innerHTML
                    
                    
                    If podlicz = 22 Then
                        Set tryGetDataFromMaestro = iElement
                        Exit Function
                    End If
                End If
            
            End If
            
            podlicz = podlicz + 1
            
            
        Next el
    End With
End Function

'-------added---------
'getting data from 2720 daily Requirements in extended horizon
Public Function getExtendedCovData(sys As Variant, plt As Variant, partNumber As Variant, Optional e)

    num = 2720

    screenUrl = "/getProductSummaryRead.do?beanId="
    preUrl = ".control.erp.corail.inetpsa.com"
    
    baseUrl = ".erp.corail.inetpsa.com"
    koUrl = baseUrl & "/ko5"
    kodzik = "vjp0Vr"
    oldPreUrl = ".commons.erp.corail.inetpsa.com/"
    urlExtendedDays = ""
    
    ' more rqms
    ' http://sx.control.erp.corail.inetpsa.com/getFbpcForProductSummaryList.do?productCode=YP000501KU
    moreRqmsUrl = "/getFbpcForProductSummaryList.do?productCode="
    
    Set req = Nothing
   ' Set valid = New Validator
    'Set theCorailData = Nothing
    ' check if o-auth req
    If UCase(sys) = "OAUTH" Then
    
        baseUrl = ".erp.corail.inetpsa.com"
    
        authUrlExtendedDays = "http://" & LCase(plt) & _
            ".erp.corail.inetpsa.com/auth/oauth/customauthorize?client_id=corailClientId&redirect_uri=http://" & LCase(plt) & _
            ".erp.corail.inetpsa.com/&response_type=code&state=" & Trim(kodzik)
            
        baseUrl = "http://" & CStr(plt) & Trim(baseUrl)
        url = "http://" & CStr(plt) & koUrl & screenUrl & CStr(partNumber)
    Else
        authUrlExtendedDays = ""
        baseUrl = "http://" & CStr(plt) & Trim(preUrl)
        url = "http://" & CStr(plt) & preUrl & screenUrl & CStr(partNumber)
    End If
    

    Dim answer As String

    'Set ICorailScreen_getData = Nothing


    If UCase(sys) = "OAUTH" Then
        url = "http://" & CStr(plt) & koUrl & moreRqmsUrl & CStr(partNumber)
        
    Else
        url = "http://" & CStr(plt) & preUrl & moreRqmsUrl & CStr(partNumber)
    End If
    
    On Error Resume Next
    'make authorization and retrieve data
    answer = httpGet2(url, authUrlExtendedDays)
    
    If Err.number <> 0 Then
       ''debug.Print "The Error : " & Err.Description & vbCrLf _
              & " was generated in " & Err.Source
              
       Err.Clear
    End If
    
    
    If valid.checkAnswer(num, answer) Then

        'Set theCorailData = New CorailDataFrom2720
        setString answer
        
        
        getExtendedCovData = answer
        
        
        
        
        
        
        
'            Set suitableXtraDataFrom2720 = theCorailData.getSuitableData()
'            theCorailData.clearSuitableData
'            Set theCorailData = Nothing
'            Set theCorail = Nothing
        'Debug.Print (collectonOfXtraCorailItems.xtraRqms.Count)
        '-----added----
    Else
        Set theCorailData = Nothing
    End If
    
    
    'Set ICorailScreen_getData = theCorailData
    'Set theCorailData = Nothing
    Set req = Nothing
    answer = ""
    ' -----------------------------------------------------------------------------
'End If
    
End Function
'getting response from extended days horizon
Public Function httpGet2(url As String, authUrl As String) As String



    Dim AuthCredentials As String, tmpResp As String
    AuthCredentials = "Basic " & Base64Encode(login & ":" & pass)


    With request2720
    
    
        If authUrl <> "" Then

             authUrl = "http://vh.erp.corail.inetpsa.com/auth/oauth/customauthorize?client_id=corailClientId&redirect_uri=http://vh.erp.corail.inetpsa.com/&response_type=code&state=7JjvnG"
            .Open "GET", authUrl, False
            ''debug.Print authUrl
            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            .SetAutoLogonPolicy AutoLogonPolicy_Always
            .setRequestHeader "Authorization", AuthCredentials
            .SetCredentials login, pass, 0


            .send
            ' DoEvents
            .WaitForResponse

            ' 'debug.Print .responseText

        End If
        
        
        .Open "GET", url, False
        ''debug.Print url
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .setRequestHeader "Authorization", AuthCredentials
        .SetAutoLogonPolicy AutoLogonPolicy_Always
        .SetCredentials login, pass, 0
        .send
        ' DoEvents
        .WaitForResponse
        
        ' 'debug.Print .responseText
        
        
        'Dim f As FSOHandler
        'Set f = New FSOHandler
        'f.makeHtmlFile "aaa", .responseText
        httpGet2 = ""
        httpGet2 = CStr(.responseText)
    End With

End Function
Private Sub setString(arg As String)
    rawString = arg
    rawString = "<style> table tr td { padding: 4px; }</style>" & Chr(10) & rawString
    'With theDomHandler
    '    Set .doc = New HTMLDocument
    '    Set .idoc = New HTMLDocument
    '    .doc.body.innerHTML = arg
    '    .idoc.body.innerHTML = .doc.body.innerHTML
    'End With
    
End Sub

Private Function ICorailScreen_openData(plt As Variant, partNumber As Variant) As Boolean
    ICorailScreen_openData = False
End Function

Private Sub ICorailScreen_screenSetup(number As Variant, httpReq As HTTPRequestHandler)

    Set req = httpReq

End Sub

