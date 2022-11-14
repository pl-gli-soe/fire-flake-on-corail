VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParserHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'
'Copyright (c) 2019 FORREST
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


Private dateCatcher As String

Private besoinsPCCatcher As String
Private besoinsBCCatcher As String

Private commandCatcher1 As String
Private commandCatcher2 As String
Private commandCatcher3 As String
Private commandCatcher4 As String
Private cc1 As String
Private cc2 As String
Private cc3 As String
Private cc4 As String
Private expCatcher1 As String
Private expCatcher2 As String
Private expCatcher3 As String
Private expCatcher4 As String
Private ec1 As String
Private ec2 As String
Private ec3 As String
Private ec4 As String
Private collectonOfXtraCorailItems As CorailIteration

Private rawCorailData As HTMLDocument
Private sd As Variant


Public Property Get pDateCatcher() As String
    pDateCatcher = dateCatcher
End Property

Public Property Get pBesoinsPCCattcher() As String
    pBesoinsPCCattcher = besoinsPCCatcher
End Property

Public Property Get pBesoinsBCCatcher() As String
    pBesoinsBCCatcher = besoinsBCCatcher
End Property

Public Property Get pCmdCatcher1() As String
    pCmdCatcher1 = commandCatcher1
End Property

Public Property Get pExpCatcher1() As String
    pExpCatcher1 = expCatcher1
End Property

Public Property Get pCmdCatcher2() As String
    pCmdCatcher2 = commandCatcher2
End Property

Public Property Get pExpCatcher2() As String
    pExpCatcher2 = expCatcher2
End Property

Public Property Get pCmdCatcher3() As String
    pCmdCatcher3 = commandCatcher3
End Property

Public Property Get pExpCatcher3() As String
    pExpCatcher3 = expCatcher3
End Property

Public Property Get pCmdCatcher4() As String
    pCmdCatcher4 = commandCatcher4
End Property

Public Property Get pExpCatcher4() As String
    pExpCatcher4 = expCatcher4
End Property

Public Function getConvertedDataSuitableForExcel() As CorailIteration
    Set getConvertedDataSuitableForExcel = collectonOfXtraCorailItems
End Function

' new from 2020-07-xx
Public Function htmlDataIntoCovertedData(collectonOfXtraCorailItemsBrangToMerge As CorailIteration, tb As HTMLTable, dom As DOMHandler, domsFrom2510 As Collection, e As E_TYPE_OF_CORAIL, Optional BO As String) As ConvertedData


    Dim tmp_cd As ConvertedData
    ' only data from 2710
    Set tmp_cd = htmlTableToRawMatrix(collectonOfXtraCorailItemsBrangToMerge, tb, dom, e)

    Set htmlDataIntoCovertedData = Nothing
    
    If BO = "BO" Then
        ' nop req
    Else
        Set tmp_cd = addSomeDataFrom2510(domsFrom2510, tmp_cd, e)
    End If
    
    
    Set htmlDataIntoCovertedData = tmp_cd
    
        
        
End Function

Private Function addSomeDataFrom2510(domsFrom2510 As Collection, cd As ConvertedData, e As E_TYPE_OF_CORAIL) As ConvertedData



    ' Debug.Print " ile jest stron na 2510: " & domsFrom2510.Count
    Dim d As IHTMLDocument
    'For Each d In domsFrom2510
    '    Debug.Print d.getElementsByTagName("table")(0).Rows.Length
    'Next d

    If e = BLUE And domsFrom2510.Count > 0 Then
    
        Dim itb As IHTMLTable, r As Variant, c As Variant
        Dim i As Integer, line As Integer
        
        Dim checkLabels As String
        
        Dim insideDiv As String
        
        Dim ci2510 As CorailItem2510
        
        
        If domsFrom2510.Count > 0 Then
        
            For Each d In domsFrom2510
                ' Debug.Print TypeName(itb)
                
                Set itb = d.getElementsByTagName("table")(0)
                
                ' MsgBox itb.outerHTML
                
                
                
                
                
                line = 1
                For Each r In itb.Rows
                    
                    ' Debug.Print TypeName(r)
                
                    ' caly row
                    ' Debug.Print r.innerHTML
                    If line = 1 Then
                        checkLabels = r.innerHTML
                    End If

                    
                    If checkLabels Like "<TH>&nbsp;</TH>*" Then
                        i = 1
                    Else
                        i = 2
                    End If
                    
                    For Each c In r.Cells
                    
                    
                        ' pojedyncze celle
                    
                    
                        ' inside table from screen 2510:
                        '&nbsp; ' CONSULT
                        'SGR/Line
                        'Product
                        'Qty
                        'order Number
                        'SID
                        'DHCA
                        'DHEO
                        'DHEF
                        'DHRP
                        'DHXP
                        'DHAS
                        'DHRQ
                        'Production Day(JP)
                        'routing
                        'Q
                        'Seller
                        'Shipper
                        's
                        'Type
                        'UM
                    
                        ' s = c.innerHTML
                        ' Debug.Print "data: " & s & " , " & line & " " & i
                        
                        If line > 1 Then
                            ' it means that this is not label but only data here
                            
                            If i = FFOC.E_2510_SGR_LINE Then
                                ' just constructor
                                Set ci2510 = New CorailItem2510
                            ElseIf i = FFOC.E_2510_DHAS Then
                                insideDiv = removeDivs(c.innerHTML)
                                ci2510.dhxx.DHAS = parseCorailStringDateToDateAndTime(insideDiv)
                            ElseIf i = FFOC.E_2510_DHEO Then
                                insideDiv = removeDivs(c.innerHTML)
                                ci2510.dhxx.DHEO = parseCorailStringDateToDateAndTime(insideDiv)
                            ElseIf i = FFOC.E_2510_DHEF Then
                                insideDiv = removeDivs(c.innerHTML)
                                ci2510.dhxx.dhef = parseCorailStringDateToDateAndTime(insideDiv)
                            ElseIf i = FFOC.E_2510_DHRQ Then
                                insideDiv = removeDivs(c.innerHTML)
                                ci2510.dhxx.dhrq = parseCorailStringDateToDateAndTime(insideDiv)
                                ci2510.dhxx.dhrq2 = parseCorailStringDateToDate(insideDiv)
                                
                            ElseIf i = FFOC.E_2510_orderNumber Then
                                insideDiv = removeDivs(c.innerHTML)
                                ci2510.orderNum = extraBoxing(CStr(insideDiv))
                                
                            ElseIf i = FFOC.E_2510_PRODUCT Then
                            
                                If c.innerHTML Like "*&nbsp;*" Then
                                    Exit For
                                Else
                                    insideDiv = removeDivs(c.innerHTML)
                                    
                                    ci2510.pn = extraBoxing(CStr(insideDiv))
                                End If
                                
                            ElseIf i = FFOC.E_2510_QTY Then
                                insideDiv = removeDivs(c.innerHTML)
                                
                                ci2510.qty = CDbl(ci2510.parseQtyFrom2510(CStr(insideDiv)))
                                
                            ElseIf i = FFOC.E_2510_TYPE Then
                                insideDiv = removeDivs(c.innerHTML)
                                
                                ci2510.itemType = extraBoxing(CStr(insideDiv))
                            End If
                            
                            
                        End If
                        
                        i = i + 1
                    Next c
                    
                    
                    ' tutaj podsumowanie rowa
                    ' ------------------------------------------------------
                    If line > 1 Then
                        cd.iteration.getItemByDate(ci2510.dhxx.dhrq2).addItemFrom2510 ci2510
                        cd.iteration.addCorailItem2510ToExtraCollection ci2510
                    End If
                    
                    ' ------------------------------------------------------
                    
                    line = line + 1
                Next r
            Next d
        End If
    End If
    
    
    
    Set addSomeDataFrom2510 = cd
End Function


Private Function extraBoxing(str As String)
    On Error Resume Next
    extraBoxing = CStr(str)
End Function


 
Private Function removeDivs(strWithDiv As String) As String


    Dim tmpArr As Variant, tmpstr As String
    
    If UCase(strWithDiv) Like "<DIV*</DIV>" Then
        
        tmpArr = Split(strWithDiv, ">")
        tmpstr = tmpArr(1)
        tmpstr = Replace(UCase(tmpstr), "</DIV", "")
        removeDivs = Trim(tmpstr)
    Else
        removeDivs = Trim(strWithDiv)
    End If
End Function

Private Function parseCorailStringDateToDateAndTime(corailStringDate As String) As Date

    
    ' format MM/DD/YY HH:MM
    
    Dim d As Integer
    Dim m As Integer
    Dim y As Integer
    
    Dim hh_mm As String
    
    If checkIfLongDate(corailStringDate) Then
    
        If IsNumeric(Left(corailStringDate, 2)) Then
            d = Int(Left(corailStringDate, 2))
        End If
        
        If IsNumeric(Mid(corailStringDate, 4, 2)) Then
            m = Int(Mid(corailStringDate, 4, 2))
        End If
        
        If IsNumeric(Mid(corailStringDate, 7, 2)) Then
            y = Int(Mid(corailStringDate, 7, 2))
            
            y = 2000 + y
        End If
        
                
        If Right(corailStringDate, 5) Like "??:??" Then
            hh_mm = Right(corailStringDate, 5)
        End If
    
    End If
    
    s = Format(y & "-" & m & "-" & d & " " & CStr(hh_mm), "yyyy-mm-dd hh:mm")
    
    On Error Resume Next
    parseCorailStringDateToDateAndTime = CDate(s)

End Function


Private Function parseCorailStringDateToDate(corailStringDate As String) As Date

    
    ' format MM/DD/YY HH:MM
    
    Dim d As Integer
    Dim m As Integer
    Dim y As Integer
    
    
    
    If checkIfLongDate(corailStringDate) Then
    
        If IsNumeric(Left(corailStringDate, 2)) Then
            d = Int(Left(corailStringDate, 2))
        End If
        
        If IsNumeric(Mid(corailStringDate, 4, 2)) Then
            m = Int(Mid(corailStringDate, 4, 2))
        End If
        
        If IsNumeric(Mid(corailStringDate, 7, 2)) Then
            y = Int(Mid(corailStringDate, 7, 2))
            
            y = 2000 + y
        End If

    
    End If
    
    s = Format(y & "-" & m & "-" & d, "yyyy-mm-dd")
    
    On Error Resume Next
    parseCorailStringDateToDate = CDate(s)

End Function


Private Function checkIfLongDate(strD) As Boolean

    checkIfLongDate = False
    
    If Len(strD) = 14 Then
    
        If Mid(strD, 3, 1) = "/" Then
            If Mid(strD, 6, 1) = "/" Then
                If Mid(strD, 12, 1) = ":" Then
                    checkIfLongDate = True
                End If
            End If
        End If
    End If
End Function

Public Function htmlTableToRawMatrix(collectonOfXtraCorailItemsBrangToMerge As CorailIteration, tb As HTMLTable, dom As DOMHandler, e As E_TYPE_OF_CORAIL) As ConvertedData

    Dim htmlString As String
    Dim tmp As ConvertedData
    
    
    Dim i As CorailItem
    Dim ii As CorailIteration
    
    Dim innerTb As HTMLTable

    Dim r As HTMLTableRow
    Dim c As HTMLTableCell
    
    Dim s As String
    
    ' new 21 oct 2019
    ' Blockages in progress - decrement from stock
    Dim tmpBIP As String
    
    Dim tmpStock As String
    Dim tmpRecv As String
    Dim tmpBacklog As String
    Dim tmpHazards As String
    
    Dim tmpResp2 As String
    
    
    Dim tmpSupplier As String
    Dim tmpPartName As String
    
    Set ii = New CorailIteration
        
    If e = ORANGE Then
    
    
        ' table with rqm, order, ship, balance
        ' ----------------------------------------------------------------
        
        Set tmp = New ConvertedData
        
        If Not tb Is Nothing Then
            htmlString = tb.innerHTML
            If checkIfHTMLTable(htmlString) Then
        
        
                
                Set innerTb = tb.getElementsByTagName("table")(0)
    
                avoidHeading = True
                For Each r In innerTb.Rows
                
                    Set i = Nothing
                    Set i = New CorailItem
                
                    For Each c In r.Cells
                        
                        
                        If avoidHeading Then
                            Exit For
                        Else
                            ' Debug.Print c.innerHTML
                            
                            
                            If UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & UCase(Replace(Me.pDateCatcher, Chr(34), "")) & "*" Then
                                
                                
                                s = CStr(c.innerHTML)
                                arr = Split(s, ">")
                                
                                s = Left(arr(1), 10)
                                
                                i.parseStringToDate s
                                
                                
                            ElseIf UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & UCase(Replace(Me.pBesoinsPCCattcher, Chr(34), "")) & "*" Then
                                
                                s = CStr(c.innerHTML)
                                arr = Split(s, ">")
                                s = arr(1)
                                s = Replace(UCase(s), "</DIV", "")
                                s = Replace(s, ".0", "")
                                s = Replace(s, ",0", "")
                                
                                i.besoinsPC = CLng(s)
                                
                            ElseIf UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & UCase(Replace(Me.pBesoinsBCCatcher, Chr(34), "")) & "*" Then
                            
                                s = CStr(c.innerHTML)
                                arr = Split(s, ">")
                                s = arr(1)
                                s = Replace(UCase(s), "</DIV", "")
                                s = Replace(s, ".0", "")
                                s = Replace(s, ",0", "")
                                
                                i.besoinsBC = CLng(s)
                                
                            ElseIf (UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher1) & "*") Or _
                                (UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher2) & "*") Or _
                                (UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher3) & "*") Or _
                                UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher4) & "*" Then
                            
                                s = CStr(c.innerHTML)
                                arr = Split(s, ">")
                                s = arr(1)
                                s = Replace(UCase(s), "</DIV", "")
                                s = Replace(s, ".0", "")
                                s = Replace(s, ",0", "")
                                
                                i.order = CLng(s)
                                
                            ElseIf (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher1) & "*") Or _
                                (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher2) & "*") Or _
                                (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher3) & "*") Or _
                                (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher4) & "*") Then
                            
                                s = CStr(c.innerHTML)
                                arr = Split(s, ">")
                                s = arr(1)
                                s = Replace(UCase(s), "</DIV", "")
                                s = Replace(s, ".0", "")
                                s = Replace(s, ",0", "")
                                
                                i.ship = CLng(s)
                                
                            End If
                            
                        End If
                        
                    Next c
                    
                    If Not avoidHeading Then
                    
                        ii.addItem i
                    End If
                    
                    avoidHeading = False
                Next r
        
            End If
        End If
        
        tmpStock = dom.stockInputElement.Value
        tmpStock = Replace(tmpStock, ".0", "")
        tmpStock = Replace(tmpStock, ",0", "")
        
        If Not IsNumeric(tmpStock) Then tmpStock = "0"
        
        tmpRecv = dom.recvInputElement.Value
        tmpRecv = Replace(tmpRecv, ".0", "")
        tmpRecv = Replace(tmpRecv, ",0", "")
        
        If Not IsNumeric(tmpRecv) Then tmpRecv = "0"
        
        tmpBacklog = dom.backlogInputElement.Value
        tmpBacklog = Replace(tmpBacklog, ".0", "")
        tmpBacklog = Replace(tmpBacklog, ",0", "")
        
        If Not IsNumeric(tmpBacklog) Then tmpBacklog = "0"
        
        Set tmp.iteration = ii
        tmp.iteration.setStock CLng(tmpStock)
        tmp.iteration.setRecv CLng(tmpRecv)
        tmp.iteration.setBacklog CLng(tmpBacklog)
        
        tmp.iteration.setComment 1, CStr(ieh.txtIsAvail)
        
        Set htmlTableToRawMatrix = tmp
        
    
    ElseIf e = BLUE Then
    
    
        ' table with rqm, order, ship, balance
        ' ----------------------------------------------------------------
        
        Set tmp = New ConvertedData
        
        If Not tb Is Nothing Then
            htmlString = tb.innerHTML
            If checkIfHTMLTable(htmlString) Then
            
            
                Set innerTb = Nothing
                On Error Resume Next
                Set innerTb = tb.getElementsByTagName("table")(0)
                
                If Not innerTb Is Nothing Then
                
                    avoidHeading = True
                    For Each r In innerTb.Rows
                    
                        Set i = Nothing
                        Set i = New CorailItem
                    
                        For Each c In r.Cells
                            
                            
                            If avoidHeading Then
                                Exit For
                            Else
                                ' Debug.Print c.innerHTML
                                
                                
                                If UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & Replace(UCase(Me.pDateCatcher), Chr(34), "") & "*" Then
                                    
                                    
                                    s = CStr(c.innerHTML)
                                    arr = Split(s, ">")
                                    
                                    s = Left(arr(1), 10)
                                    
                                    i.parseStringToDate s
                                    
                                    
                                ElseIf UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & UCase(Replace(Me.pBesoinsPCCattcher, Chr(34), "")) & "*" Then
                                    
                                    s = CStr(c.innerHTML)
                                    arr = Split(s, ">")
                                    s = arr(1)
                                    s = Replace(UCase(s), "</DIV", "")
                                    ' s = replaceAllDecimals(s)
                                    
                                    s = adjustDecimalSeparator(s)
                                    i.besoinsPC = CDbl(s)
                                    
                                ElseIf UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & UCase(Replace(Me.pBesoinsBCCatcher, Chr(34), "")) & "*" Then
                                
                                    s = CStr(c.innerHTML)
                                    arr = Split(s, ">")
                                    s = arr(1)
                                    s = Replace(UCase(s), "</DIV", "")
'                                    s = Replace(s, ".0", "")
'                                    s = Replace(s, ",0", "")
                                    
                                    ' s = replaceAllDecimals(s)
                                    
                                    ' Application.DecimalSeparator
                                    
                                    s = adjustDecimalSeparator(s)
                                    i.besoinsBC = CDbl(s)
                                    
                                    
                                    'Debug.Print i.besoinsBC
                                    
                                ElseIf (UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher1) & "*") Or _
                                    (UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher2) & "*") Or _
                                    (UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher3) & "*") Or _
                                    UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher4) & "*" Then
                                
                                    s = CStr(c.innerHTML)
                                    arr = Split(s, ">")
                                    s = arr(1)
                                    s = Replace(UCase(s), "</DIV", "")
'                                    s = Replace(s, ".0", "")
'                                    s = Replace(s, ",0", "")

                                    ' s = replaceAllDecimals(s)
                                    
                                    s = adjustDecimalSeparator(s)
                                    i.order = CDbl(s)
                                    'Debug.Print i.order
                                    
                                ElseIf (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher1) & "*") Or _
                                    (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher2) & "*") Or _
                                    (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher3) & "*") Or _
                                    (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher4) & "*") Then
                                
                                    s = CStr(c.innerHTML)
                                    arr = Split(s, ">")
                                    s = arr(1)
                                    s = Replace(UCase(s), "</DIV", "")
'                                    s = Replace(s, ".0", "")
'                                    s = Replace(s, ",0", "")

                                    ' s = replaceAllDecimals(s)
                                    
                                    s = adjustDecimalSeparator(s)
                                    i.ship = CDbl(s)
                                    
                                    'Debug.Print i.ship
                                    
                                End If
                                
                            End If
                            
                        Next c
                        
                        If Not avoidHeading Then
                        
                            ii.addItem i
                        End If
                        
                        avoidHeading = False
                    Next r
                    
                    

            Dim nrOfCorailItemsInCorailIteration As Integer
            nrOfCorailItemsInCorailIteration = ii.pItems.Count
            
            Dim lastDateInCollectionFromFF As Date
            lastDateInCollectionFromFF = ii.pItems()(nrOfCorailItemsInCorailIteration).getDate()
            
            Dim filteredCollectonOfXtraCorailItemsBrangToMerge As Collection
            Set filteredCollectonOfXtraCorailItemsBrangToMerge = New Collection
            
            Dim xtraRqmCorailItem As CorailItem
            
            'filtering CorailIteration with data from extended Horizon, by length of sgrLine
            For Each xtraRqmCorailItem In collectonOfXtraCorailItemsBrangToMerge.pItems


                If Len(Trim(xtraRqmCorailItem.sgrLine)) = 2 Then

                    filteredCollectonOfXtraCorailItemsBrangToMerge.Add xtraRqmCorailItem
                    
                End If


             Next xtraRqmCorailItem
             
             'clearing collection of CorailItems of Extended Horizon
             Set collectonOfXtraCorailItemsBrangToMerge.pItems = Nothing
             
             'Setting filtered data collection as new collection of CorailItems of Extended Horizon
             Set collectonOfXtraCorailItemsBrangToMerge.pItems = filteredCollectonOfXtraCorailItemsBrangToMerge
          
             
             Set filteredCollectonOfXtraCorailItemsBrangToMerge = Nothing
           
            'repackaging extended CoraiItem(s) to CorailIteration of FireFlake
             Dim lColItem As CorailItem
             
             For Each lColItem In collectonOfXtraCorailItemsBrangToMerge.pItems
                
                If lColItem.getDate() >= lastDateInCollectionFromFF Then
                    
                    If lColItem.getDate() = lastDateInCollectionFromFF Then
                    
                        ii.pItems()(nrOfCorailItemsInCorailIteration).besoinsBC = lColItem.Total

                        i.besoinsPC = 0
                        
                    Else
                    
                      
                        lColItem.setDate (CStr(lColItem.getDate()))
                        
                        lColItem.besoinsBC = lColItem.Total
                        
                        lColItem.besoinsPC = 0
                        
                        ii.addItem lColItem
                    
                    End If
                End If

             Next
'
                    
                    
                    
                End If
            End If
        End If
        
        If Not dom.stockInputElement Is Nothing Then
            tmpStock = dom.stockInputElement.Value
            tmpStock = adjustDecimalSeparator(tmpStock)
        End If
        
        If Not IsNumeric(tmpStock) Then tmpStock = "0"
        
        If Not dom.recvInputElement Is Nothing Then
            tmpRecv = dom.recvInputElement.Value
            tmpRecv = adjustDecimalSeparator(tmpRecv)
        End If
        
        If Not IsNumeric(tmpRecv) Then tmpRecv = "0"
        
        If Not dom.backlogInputElement Is Nothing Then
            tmpBacklog = dom.backlogInputElement.Value
            tmpBacklog = adjustDecimalSeparator(tmpBacklog)
        End If
        
        If Not dom.hazardsInputElement Is Nothing Then
            tmpHazards = dom.hazardsInputElement.Value
            tmpHazards = adjustDecimalSeparator(tmpHazards)
        End If
        
        If Not dom.iBlockagesInProgress Is Nothing Then
        
            tmpBIP = dom.iBlockagesInProgress.Value
            tmpBIP = adjustDecimalSeparator(tmpBIP)
            
        End If
        
        
        If Not dom.isupplierElement Is Nothing Then
        
            tmpSupplier = Trim(dom.isupplierElement.innerText)
            tmpSupplier = Trim(focusOnSupplierName(tmpSupplier))
        End If
        
        
        If Not dom.ipartnameElement Is Nothing Then
        
            tmpPartName = dom.ipartnameElement.innerText
            tmpPartName = Trim(Split(tmpPartName, "-")(1))
        End If
        
        
        If Not dom.iapproElement Is Nothing Then
            tmpResp2 = dom.iapproElement.Value
        End If
        
        
        If Not IsNumeric(tmpBacklog) Then tmpBacklog = "0"
        If Not IsNumeric(tmpBIP) Then tmpBIP = "0"
        If Not IsNumeric(tmpHazards) Then tmpHazards = "0"
        
        Set tmp.iteration = ii
        'tmp.iteration.setStock CDbl(Replace(tmpStock, ".", Application.DecimalSeparator))
        'tmp.iteration.setStock CDbl(Replace(tmpStock, ",", Application.DecimalSeparator))
        tmp.iteration.setStock CDbl(tmpStock)
        
        'tmp.iteration.setRecv CDbl(Replace(tmpRecv, ".", Application.DecimalSeparator))
        'tmp.iteration.setRecv CDbl(Replace(tmpRecv, ",", Application.DecimalSeparator))
        tmp.iteration.setRecv CDbl(tmpRecv)
        
        'tmp.iteration.setBacklog CDbl(Replace(tmpBacklog, ".", Application.DecimalSeparator))
        'tmp.iteration.setBacklog CDbl(Replace(tmpBacklog, ",", Application.DecimalSeparator))
        tmp.iteration.setBacklog CDbl(tmpBacklog)
        
        tmp.iteration.setPartName CStr(tmpPartName)
        tmp.iteration.setSupplier CStr(tmpSupplier)
        tmp.iteration.setResp CStr(tmpResp2)
        ' aleas == hazards
        tmp.iteration.setAleas CStr(tmpHazards)
        
        ' new element for Blockages in progress
        'tmpBIP = CDbl(Replace(tmpBIP, ".", Application.DecimalSeparator))
        'tmpBIP = CDbl(Replace(tmpBIP, ",", Application.DecimalSeparator))
        tmp.iteration.setBlockagesInProgress CDbl(tmpBIP)
        
        
        
        ' tmp.iteration.setComment 1, CStr(ieh.txtIsAvail)
        
        
        Set htmlTableToRawMatrix = tmp
            
    ElseIf e = MAESTRO Then
    
    
        ' table with rqm, order, ship, balance
        ' ----------------------------------------------------------------
        
        Set tmp = New ConvertedData
        
        If Not tb Is Nothing Then
        
            Set innerTb = tb
            htmlString = tb.innerHTML
            If checkIfHTMLTable(htmlString) Then
        
        
                For Each r In innerTb.Rows
                
                    Set i = Nothing
                    Set i = New CorailItem
                    
                    cCounter = 0
                
                    For Each c In r.Cells
                        
                        ' Debug.Print c.innerHTML
                        
                        ' 0: date
                        ' 1: rqm
                        ' 2: order
                        ' 3: it
                        
                        ' always
                        s = CStr(c.innerHTML)
                        
                        If cCounter = 0 Then
                            i.parseStringToDate Trim(s)
                        ElseIf cCounter = 1 Then
                        
                            arr = Split(s, "<")
                            s = Trim(arr(LBound(arr)))
                            's = Replace(s, ".", Application.DecimalSeparator)
                            's = Replace(s, ",", Application.DecimalSeparator)
                            s = adjustDecimalSeparator(s)
                            
                            i.besoinsBC = 0
                            i.besoinsPC = 0
                            If IsNumeric(s) Then i.besoinsBC = CDbl(s)
                            
                        ElseIf cCounter = 2 Then
                        
                            arr = Split(s, "<")
                            s = Trim(arr(LBound(arr)))
                            's = Replace(s, ".", Application.DecimalSeparator)
                            's = Replace(s, ",", Application.DecimalSeparator)
                            s = adjustDecimalSeparator(s)
                            i.order = 0
                            If IsNumeric(s) Then i.order = CLng(s)
                            
                        ElseIf cCounter = 3 Then
                            If s Like "<FONT*" Then
                                arr = Split(s, ">")
                                arr = Split(arr(LBound(arr) + 1), "</")
                            Else
                                arr = Split(s, "<")
                                arr = Split(arr(LBound(arr)), "</")
                            End If
                            s = Trim(arr(LBound(arr)))
                            's = Replace(s, ".", Application.DecimalSeparator)
                            's = Replace(s, ",", Application.DecimalSeparator)
                            s = adjustDecimalSeparator(s)
                            
                            i.ship = 0
                            If IsNumeric(s) Then i.ship = CLng(s)
                        End If
                        
                        cCounter = cCounter + 1
                    Next c
                    
                    ii.addItem i
                Next r
            End If
        End If
        
        If Not dom.istockElement Is Nothing Then
            tmpStock = dom.istockElement.innerHTML

            tmpStock = Trim(Replace(tmpStock, G_NBSP, ""))
            
            If tmpStock Like "*IMG*" Then
                tmpArr = Split(tmpStock, " ")
                tmpStock = Trim(tmpArr(0))
                'tmpStock = Replace(tmpStock, ".", Application.DecimalSeparator)
                'tmpStock = Replace(tmpStock, ",", Application.DecimalSeparator)
                tmpStock = adjustDecimalSeparator(tmpStock)
            End If
        End If
        
        If Not IsNumeric(tmpStock) Then tmpStock = "0"
        
        If Not dom.ibacklogElement Is Nothing Then
            tmpBacklog = dom.ibacklogElement.innerHTML
            's = Replace(s, ".", Application.DecimalSeparator)
            's = Replace(s, ",", Application.DecimalSeparator)
            tmpBacklog = Trim(Replace(tmpBacklog, G_NBSP, ""))
            
            
            If tmpBacklog Like "*FONT*" Then
                
                tmpArr = Split(tmpBacklog, ">")
                tmpBacklog = tmpArr(1)
                tmpBacklog = Replace(tmpBacklog, "</FONT", "")
                'tmpBacklog = Replace(tmpBacklog, ",", Application.DecimalSeparator)
                'tmpBacklog = Replace(tmpBacklog, ",", Application.DecimalSeparator)
                tmpBacklog = adjustDecimalSeparator(tmpBacklog)
            End If
            
            
        End If
        
        If Not IsNumeric(tmpBacklog) Then tmpBacklog = "0"
        
        
        Set tmp.iteration = ii
        tmp.iteration.setStock CLng(tmpStock)
        tmp.iteration.setRecv CLng(0)
        tmp.iteration.setBacklog CLng(tmpBacklog)
        
        ' tmp.iteration.setComment 1, CStr(ieh.txtIsAvail)
        
        
        
        Set htmlTableToRawMatrix = tmp
        
    Else
        MsgBox "brak zdefiniowanego systemu!"
        End
    End If
        
    
End Function
'NOT USED FROM WEEKLY COVERAGE
'Public Function findCovItem(d As Date, c As Collection) As CoverageItem
'
'    Set findCovItem = Nothing
'
'    Dim ci As CoverageItem
'    For Each ci In c
'        If CDate(ci.getDate()) = CDate(d) Then
'            Set findCovItem = ci
'            Exit Function
'        End If
'    Next ci
'End Function


Private Function focusOnSupplierName(sn As String) As String


    Dim tmpArrayOfStr As Variant
    tmpArrayOfStr = Split(sn, Chr(10))
    
    Dim x As Variant
    For x = LBound(tmpArrayOfStr) To UBound(tmpArrayOfStr)
        
        If tmpArrayOfStr(x) Like "R. S.:*" Then
            focusOnSupplierName = Trim(tmpArrayOfStr(x))
            Exit Function
        End If
    Next x
    
    
    focusOnSupplierName = ""
End Function




Private Function adjustDecimalSeparator(s) As String



    Dim regHalf As Range
    Set regHalf = ThisWorkbook.Sheets(FFOC.G_SH_NM_REG).Range("Q17")
    
    Dim myExcelDecimalPoint As String
    myExcelDecimalPoint = CStr(Mid(regHalf.Value, 2, 1))
    
    If myExcelDecimalPoint = Application.DecimalSeparator Then
        ' nothing to do!
    Else
        ' there is sth strage
    End If
    
    ' in corail always dots


    adjustDecimalSeparator = ""
    
    If Trim(s) <> "" Then

        If myExcelDecimalPoint = "," Then
            ' Debug.Print "Application.DecimalSeparator ,"
            If s Like "*,*" Then
                ' no operation req
            ElseIf s Like "*.*" Then
                s = Replace(s, ".", myExcelDecimalPoint)
            Else
                ' no operation req - int num
            End If
            
        ElseIf myExcelDecimalPoint = "." Then
            ' Debug.Print "Application.DecimalSeparator ."
            If s Like "*.*" Then
                ' no operation req
            ElseIf s Like "*,*" Then
                s = Replace(s, ",", myExcelDecimalPoint)
            Else
                ' no operation req - int num
            End If
        Else
            MsgBox "wrong decimal separator - please change it in control panel!"
            End
        End If
        
        
        adjustDecimalSeparator = CStr(s)
    Else
        adjustDecimalSeparator = ""
    End If


End Function


Private Function checkIfHTMLTable(htmlString As String) As Boolean


    Dim tmp As String
    tmp = CStr(htmlString)
    
    If UCase(tmp) Like UCase("*tr*td*td*tr*") Then
    
        checkIfHTMLTable = True
    Else
    
        checkIfHTMLTable = False
    End If
End Function

Private Sub Class_Initialize()
    dateCatcher = "<div class=" & Chr(34) & "date" & Chr(34) & ">"

    besoinsPCCatcher = "<div class=" & Chr(34) & "number fbpcDetailByDay" & Chr(34) & ""
    besoinsBCCatcher = "<div class=" & Chr(34) & "number" & Chr(34) & ">"

    commandCatcher1 = "<div class=" & Chr(34) & "ecwButtonTexteOverOrdered number" & Chr(34) & ""
    commandCatcher2 = "<div class=" & Chr(34) & "ecwButtonTexteOverOrdered number ecwButtonTexteOver" & Chr(34) & ""
    commandCatcher3 = "<div class=" & Chr(34) & "number ecwButtonTexteOverSauOrdered" & Chr(34) & ""
    commandCatcher4 = "<div class=" & Chr(34) & "number ecwButtonTexteOverSauOrdered ecwButtonTexteOver" & Chr(34) & ""
    
    cc1 = commandCatcher1
    cc2 = commandCatcher2
    cc3 = commandCatcher3
    cc4 = commandCatcher4
    
    
    expCatcher1 = "<div class=" & Chr(34) & "ecwButtonTexteOverShipped number" & Chr(34) & ""
    expCatcher2 = "<div class=" & Chr(34) & "ecwButtonTexteOverShipped number ecwButtonTexteOver" & Chr(34) & ""
    expCatcher3 = "<div class=" & Chr(34) & "number ecwButtonTexteOverSauShipped" & Chr(34) & ""
    expCatcher4 = "<div class=" & Chr(34) & "number ecwButtonTexteOverSauShipped ecwButtonTexteOver" & Chr(34) & ""
    
    ec1 = expCatcher1
    ec2 = expCatcher2
    ec3 = expCatcher3
    ec4 = expCatcher4
    
    Set rawCorailData = Nothing
 '   Set valid = New Validator
    Set sd = Nothing
End Sub



'***********************************************************************************************************************
Public Sub importPackageOfData(pckg As HTMLDocument)
    
    Set rawCorailData = pckg
    
    ' and try parse it!
    ' CorailDataFrom2720 or CorailDataFrom2510
    
    innerParse2720
'    If TypeName(pckg) = "CorailDataFrom2720" Then
'        innerParse2720
'    ElseIf TypeName(pckg) = "CorailDataFrom2510" Then
'        'innerParse2510
'    ElseIf TypeName(pckg) = "CorailDataFrom6510" Then
'        'innerParse6510
'    End If
End Sub









'  ____  ______     _     __
' /'___\/\  ___\  /' \  /'__`\
'/\ \__/\ \ \__/ /\_, \/\ \/\ \
'\ \  _``\ \___``\/_/\ \ \ \ \ \
' \ \ \L\ \/\ \L\ \ \ \ \ \ \_\ \
'  \ \____/\ \____/  \ \_\ \____/
'   \/___/  \/___/    \/_/\/___/
'
'
'



'
' #################################################################################
' #################################################################################
'Private Sub innerParse6510()
'
'    If valid.makeSomeValidationOnImportedPackage(rawCorailData) Then
'        ' it is ok go on we working with this data
'        ' -------------------------------------------
'
'        Dim d As ICorailData ' CorailDataFrom2510
'        Set d = rawCorailData
'
'
'        Set sd = New SuitableData6510
'
'        Dim coll As Collection
'        ''debug.Print "test d.getCollectionOfData().count: " & d.getCollectionOfData().Count
'        Set coll = d.getCollectionOfData()
'
'
'        ' c is an collection of pages
'
'        ' !!!
'        '  -> change into dom
'        Dim stronaOrderow As Variant
'        Dim maintables As IHTMLElementCollection
'        Dim divElement As IHTMLElement
'        Dim divTempElement As IHTMLElement
'        Dim parentDivTable As IHTMLElement
'
'
'        Dim tds As IHTMLElementCollection, t As IHTMLTable, tc As IHTMLTableCell ' divTempElement - juz zadeklarowany i go wykorzystam tutaj
'        Dim r As HTMLTableRow
'        Dim c As HTMLTableCell
'
'        Dim ti As TTItem
'
'
'        ' temps vars - really ugly - i know
'        ' --------------------------------------------------------------------------------------------------
'        ' --------------------------------------------------------------------------------------------------
'
'        Dim tmpStrArr(WCOC.E_6510_CONSULT_ICON To WCOC.E_6510_DOCK) As String
'
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Dim tmpColIndxArr(WCOC.E_6510_CONSULT_ICON To WCOC.E_6510_DOCK)
'        For xx = WCOC.E_6510_CONSULT_ICON To WCOC.E_6510_DOCK
'            tmpColIndxArr(xx) = -1
'        Next xx
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'        Dim tmpDH_Handler As DH_Handler
'
'        ' --------------------------------------------------------------------------------------------------
'        ' --------------------------------------------------------------------------------------------------
'
'
'
'        For Each stronaOrderow In coll
'
'
'            francuski = False
'            angielski = False
'
'
'            ' every page will be treat as seperate web-page
'            d.getDom.idoc.body.innerHTML = stronaOrderow
'            d.getDom.doc.body.innerHTML = stronaOrderow
'
'
'
'            Set maintables = d.getDom.idoc.getElementsByTagName("div")
'
'
'            ' just get first one!
'            ' ===========================================
'            For Each divElement In maintables
'
'                If divElement.className = "ecwTableBody" Then
'
'                    Set parentDivTable = divElement
'                    Exit For
'                End If
'            Next divElement
'            ' ===========================================
'
'
'
'            Set t = parentDivTable.getElementsByTagName("table")(0)
'
'            For Each r In t.Rows
'
'
'                Set ti = Nothing
'
'                ' init by empty
'                ' ----------------------------------------------
'                'Dim tmpStrArr(WCOC.E_2510_CONSULT_ICON To WCOC.E_2510_UM) As String
'                For i = WCOC.E_6510_CONSULT_ICON To WCOC.E_6510_DOCK
'                    tmpStrArr(i) = ""
'                Next i
'
'                Set tmpDH_Handler = New DH_Handler
'                ' ----------------------------------------------
'
'                i = 1
'
'                For Each c In r.Cells
'
'
'                    If c.tagName = "TH" Then
'
'
'                        ' catch the order of columns!!!
'                        ' ---------------------------------
'
'
'
'                        'If c.innerHTML = "DHCA" Then tmpColIndxArr(WCOC.E_2510_DHCA) = i
'                        If c.innerHTML = "DHEF" Then tmpColIndxArr(WCOC.E_6510_DHEF) = i
'                        'If c.innerHTML = "DHEO" Then tmpColIndxArr(WCOC.E_2510_DHEO) = i
'                        If c.innerHTML = "DHRQ" Then tmpColIndxArr(WCOC.E_6510_DHRQ) = i
'                        'If c.innerHTML = "DHAS" Then tmpColIndxArr(WCOC.E_2510_DHAS) = i
'
'
'                        ' ---------------------------------
'                    End If
'
'
'                    If c.tagName = "TD" Then
'                        Set divTempElement = c.getElementsByTagName("div")(0)
'
'
'
'                        ' DHXX here
'                        ' --------------------------------------------------------------------------
'
'                        'If i = tmpColIndxArr(WCOC.E_2510_DHEO) Then tmpStrArr(WCOC.E_2510_DHEO) = CStr(divTempElement.innerHTML)
'                        If i = tmpColIndxArr(WCOC.E_6510_DHEF) Then
'                            tmpStrArr(WCOC.E_6510_DHEF) = CStr(divTempElement.innerHTML)
'                            Set tmpDH_Handler = Nothing
'                            Set tmpDH_Handler = New DH_Handler
'
'                            If checkIfDateAndTime(tmpStrArr(WCOC.E_6510_DHEF)) Then
'                                tmpDH_Handler.dhef = parseStringToDateTime(tmpStrArr(WCOC.E_6510_DHEF))
'                            End If
'
'                            Set ti = Nothing
'                            Set ti = New TTItem
'
'
'                        End If
'                        'If i = tmpColIndxArr(WCOC.E_2510_DHAS) Then tmpStrArr(WCOC.E_2510_DHAS) = CStr(divTempElement.innerHTML)
'                        If i = tmpColIndxArr(WCOC.E_6510_DHRQ) Then
'                            tmpStrArr(WCOC.E_6510_DHRQ) = CStr(divTempElement.innerHTML)
'                            'Set tmpDH_Handler = Nothing
'                            'Set tmpDH_Handler = New DH_Handler
'
'                            If checkIfDateAndTime(tmpStrArr(WCOC.E_6510_DHRQ)) Then
'                                tmpDH_Handler.dhrq = parseStringToDateTime(tmpStrArr(WCOC.E_6510_DHRQ))
'                            End If
'
'
'                            With ti
'                                .dhConstructor tmpDH_Handler
'                                .calc
'                            End With
'                        End If
'
'
'
'                        ' --------------------------------------------------------------------------
'
'
'
'                        'Dim qtyAsDouble As Double
'                        'qtyAsDouble = 0
'                        'On Error Resume Next
'                        'qtyAsDouble = CDbl(forceStringToBeDouble(tmpStrArr(WCOC.E_2510_QTY)))
'
'                        'With oItem
'                        '    .coforConstructor tmpStrArr(WCOC.E_2510_SELLER), tmpStrArr(WCOC.E_2510_SHIPPER)
'                        '    .dhConstructor tmpDH_Handler
'                        '    .commonConstructor tmpStrArr, CDbl(qtyAsDouble)
'                        'End With
'
'
'                        '
'                        ''
'                        ' ==========================================================================
'                        ' ==========================================================================
'                    End If
'
'                    i = i + 1
'                Next c
'
'
'
'                If Not ti Is Nothing Then
'                    If CLng(ti.getTT()) <> 0 Then
'                        sd.addTT ti
'                    End If
'                End If
'
'
'            Next r
'
'
'
'            ' robimy tylko pierwsza strone!
'            Exit For
'
'
'        Next stronaOrderow
'
'
'
'        Set d = Nothing
'
'        ' -------------------------------------------
'    Else
'        MsgBox "Wrong type of package!"
'    End If
'End Sub
'''
'''
'' #################################################################################
'
'
'
'
'
'
'
'
'
'
'
'
'' 222222222222222    555555555555555555   1111111        000000000
''2:::::::::::::::22  5::::::::::::::::5  1::::::1      00:::::::::00
''2::::::222222:::::2 5::::::::::::::::5 1:::::::1    00:::::::::::::00
''2222222     2:::::2 5:::::555555555555 111:::::1   0:::::::000:::::::0
''            2:::::2 5:::::5               1::::1   0::::::0   0::::::0
''            2:::::2 5:::::5               1::::1   0:::::0     0:::::0
''         2222::::2  5:::::5555555555      1::::1   0:::::0     0:::::0
''    22222::::::22   5:::::::::::::::5     1::::l   0:::::0 000 0:::::0
''  22::::::::222     555555555555:::::5    1::::l   0:::::0 000 0:::::0
'' 2:::::22222                    5:::::5   1::::l   0:::::0     0:::::0
''2:::::2                         5:::::5   1::::l   0:::::0     0:::::0
''2:::::2             5555555     5:::::5   1::::l   0::::::0   0::::::0
''2:::::2       2222225::::::55555::::::5111::::::1110:::::::000:::::::0
''2::::::2222222:::::2 55:::::::::::::55 1::::::::::1 00:::::::::::::00
''2::::::::::::::::::2   55:::::::::55   1::::::::::1   00:::::::::00
''22222222222222222222     555555555     111111111111     000000000
''
''
'' #################################################################################
'' #################################################################################
'Private Sub innerParse2510()
'
'    If valid.makeSomeValidationOnImportedPackage(rawCorailData) Then
'        ' it is ok go on we working with this data
'        ' -------------------------------------------
'
'        Dim d As ICorailData ' CorailDataFrom2510
'        Set d = rawCorailData
'
'
'        Set sd = New SuitableData2510
'
'        Dim coll As Collection
'        ' 'debug.Print "test d.getCollectionOfData().count: " & d.getCollectionOfData().Count
'
'        If Not d Is Nothing Then
'            If Not d.getCollectionOfData() Is Nothing Then
'                If d.getCollectionOfData().Count > 0 Then
'
'
'                    ''debug.Print "Private Sub innerParse2510: d.getCollectionOfData().count: " & d.getCollectionOfData().Count
'
'                    Set coll = d.getCollectionOfData()
'
'
'
'
'
'                    ' c is an collection of pages
'
'                    ' !!!
'                    '  -> change into dom
'                    Dim stronaOrderow As Variant
'                    Dim maintables As IHTMLElementCollection
'                    Dim divElement As IHTMLElement
'                    Dim divTempElement As IHTMLElement
'                    Dim parentDivTable As IHTMLElement
'
'
'                    Dim tds As IHTMLElementCollection, t As IHTMLTable, tc As IHTMLTableCell ' divTempElement - juz zadeklarowany i go wykorzystam tutaj
'                    Dim r As HTMLTableRow
'                    Dim c As HTMLTableCell
'
'                    Dim oItem As OrderItem
'
'
'                    ' temps vars - really ugly - i know
'                    ' --------------------------------------------------------------------------------------------------
'                    ' --------------------------------------------------------------------------------------------------
'
'                    Dim tmpStrArr(WCOC.E_2510_CONSULT_ICON To WCOC.E_2510_UM) As String
'
'                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                    Dim tmpColIndxArr(WCOC.E_2510_CONSULT_ICON To WCOC.E_2510_UM)
'                    For xx = WCOC.E_2510_CONSULT_ICON To WCOC.E_2510_UM
'                        tmpColIndxArr(xx) = -1
'                    Next xx
'                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'                    Dim tmpDH_Handler As DH_Handler
'
'                    ' --------------------------------------------------------------------------------------------------
'                    ' --------------------------------------------------------------------------------------------------
'
'
'
'                    For Each stronaOrderow In coll
'
'
'                        francuski = False
'                        angielski = False
'
'
'                        ' every page will be treat as seperate web-page
'                        d.getDom.idoc.body.innerHTML = stronaOrderow
'                        d.getDom.doc.body.innerHTML = stronaOrderow
'
'
'
'                        Set maintables = d.getDom.idoc.getElementsByTagName("div")
'
'
'                        ' just get first one!
'                        ' ===========================================
'                        For Each divElement In maintables
'
'                            If divElement.className = "ecwTableBody" Then
'
'                                Set parentDivTable = divElement
'                                Exit For
'                            End If
'                        Next divElement
'                        ' ===========================================
'
'
'
'                        Set t = parentDivTable.getElementsByTagName("table")(0)
'
'                        For Each r In t.Rows
'
'                            Set oItem = Nothing
'                            Set oItem = New OrderItem
'
'
'                            ' init by empty
'                            ' ----------------------------------------------
'                            'Dim tmpStrArr(WCOC.E_2510_CONSULT_ICON To WCOC.E_2510_UM) As String
'                            For i = WCOC.E_2510_CONSULT_ICON To WCOC.E_2510_UM
'                                tmpStrArr(i) = ""
'                            Next i
'
'                            Set tmpDH_Handler = New DH_Handler
'                            ' ----------------------------------------------
'
'                            i = 1
'
'                            For Each c In r.Cells
'                                ' 'debug.Print c.outerHTML & ", class: " & c.className & " inner: " & c.innerHTML & " tag: " & c.tagName
'                                ' <TH class=ecwTableSortable>SGR/Line</TH>, class: ecwTableSortable inner: SGR/Line tag: TH
'                                ' ...
'                                ' <TH>Qty</TH>, class:  inner: Qty tag: TH
'                                ' <TH class=ecwTableSortable>DHEO</TH>, class: ecwTableSortable inner: DHEO tag: TH
'                                ' ...
'                                ' <TH>Seller </TH>, class:  inner: Seller  tag: TH
'                                ' <TH>Shipper </TH>, class:  inner: Shipper  tag: TH
'                                If c.tagName = "TH" Then
'                                    ' catch the order of columns
'                                    ' ---------------------------------
'
'                                    If Left(c.innerHTML, 3) = "SGR" Then tmpColIndxArr(WCOC.E_2510_SGR_LINE) = i
'                                    If c.innerHTML = "Product" Or c.innerHTML = "Produit" Then
'                                        tmpColIndxArr(WCOC.E_2510_PRODUCT) = i
'
'                                        If c.innerHTML = "Product" Then
'                                            francuski = False
'                                            angielski = True
'                                        ElseIf c.innerHTML = "Produit" Then
'                                            francuski = True
'                                            angielski = False
'                                        End If
'
'                                    End If
'                                    If Left(c.innerHTML, 2) = "Qt" Then tmpColIndxArr(WCOC.E_2510_QTY) = i
'                                    If c.innerHTML = "Order number" Or c.innerHTML = "Order No" Or c.innerHTML Like "*Ordre" Then tmpColIndxArr(WCOC.E_2510_ORDER_NUMBER) = i
'                                    If c.innerHTML = "SID" Then tmpColIndxArr(WCOC.E_2510_DN) = i
'
'                                    If c.innerHTML = "DHCA" Then tmpColIndxArr(WCOC.E_2510_DHCA) = i
'                                    If c.innerHTML = "DHEF" Then tmpColIndxArr(WCOC.E_2510_DHEF) = i
'                                    If c.innerHTML = "DHEO" Then tmpColIndxArr(WCOC.E_2510_DHEO) = i
'                                    If c.innerHTML = "DHRQ" Then tmpColIndxArr(WCOC.E_2510_DHRQ) = i
'                                    If c.innerHTML = "DHAS" Then tmpColIndxArr(WCOC.E_2510_DHAS) = i
'
'                                    If UCase(c.innerHTML) = UCase("Routing") Or UCase(c.innerHTML) = UCase("Routage") Then
'                                        tmpColIndxArr(WCOC.E_2510_ROUTING) = i
'
'                                        If UCase(c.innerHTML) = UCase("Routing") Then
'                                            francuski = False
'                                            angielski = True
'                                        ElseIf UCase(c.innerHTML) = UCase("Routage") Then
'                                            francuski = True
'                                            angielski = False
'                                        End If
'
'                                    End If
'                                    If Trim(c.innerHTML) = "Seller" Or c.innerHTML = "Vendeur" Then tmpColIndxArr(WCOC.E_2510_SELLER) = i
'                                    If Trim(c.innerHTML) = "Shipper" Or (Left(c.innerHTML, 3) = "Exp" And Right(c.innerHTML, 4) = "teur") Then tmpColIndxArr(WCOC.E_2510_SHIPPER) = i
'                                    If Trim(c.innerHTML) = "S" Then tmpColIndxArr(WCOC.E_2510_S2) = i
'                                    If Trim(c.innerHTML) = "Type" Then tmpColIndxArr(WCOC.E_2510_TYPE) = i
'                                    If Trim(c.innerHTML) = "UM" Then tmpColIndxArr(WCOC.E_2510_UM) = i
'
'
'
'                                    ' ---------------------------------
'                                End If
'
'
'                                If c.tagName = "TD" Then
'                                    Set divTempElement = c.getElementsByTagName("div")(0)
'
'                                    ' 'debug.Print "c.getElementsByTagName(div)(0) > o: " & divTempElement.outerHTML & " i: " & divTempElement.innerHTML
'
'                                    ' &nbsp (consult button)
'                                    If i = tmpColIndxArr(WCOC.E_2510_SGR_LINE) Then tmpStrArr(WCOC.E_2510_SGR_LINE) = CStr(divTempElement.innerHTML)
'                                    If i = tmpColIndxArr(WCOC.E_2510_PRODUCT) Then tmpStrArr(WCOC.E_2510_PRODUCT) = CStr(divTempElement.innerHTML)
'                                    If i = tmpColIndxArr(WCOC.E_2510_QTY) Then
'                                        tmpStrArr(WCOC.E_2510_QTY) = CStr(divTempElement.innerHTML)
'                                    End If
'                                    If i = tmpColIndxArr(WCOC.E_2510_ORDER_NUMBER) Then tmpStrArr(WCOC.E_2510_ORDER_NUMBER) = CStr(divTempElement.innerHTML)
'                                    If i = tmpColIndxArr(WCOC.E_2510_DN) Then tmpStrArr(E_2510_DN) = CStr(divTempElement.innerHTML)
'
'
'                                    ' DHXX here
'                                    ' --------------------------------------------------------------------------
'
'                                    If i = tmpColIndxArr(WCOC.E_2510_DHEO) Then tmpStrArr(WCOC.E_2510_DHEO) = CStr(divTempElement.innerHTML)
'                                    If i = tmpColIndxArr(WCOC.E_2510_DHEF) Then tmpStrArr(WCOC.E_2510_DHEF) = CStr(divTempElement.innerHTML)
'                                    If i = tmpColIndxArr(WCOC.E_2510_DHAS) Then tmpStrArr(WCOC.E_2510_DHAS) = CStr(divTempElement.innerHTML)
'                                    If i = tmpColIndxArr(WCOC.E_2510_DHRQ) Then tmpStrArr(WCOC.E_2510_DHRQ) = CStr(divTempElement.innerHTML)
'                                    ' --------------------------------------------------------------------------
'
'
'                                    If i = tmpColIndxArr(WCOC.E_2510_ROUTING) Then tmpStrArr(WCOC.E_2510_ROUTING) = CStr(divTempElement.innerHTML)
'                                    If i = tmpColIndxArr(WCOC.E_2510_D) Then tmpStrArr(WCOC.E_2510_D) = CStr(divTempElement.innerHTML)
'                                    If i = tmpColIndxArr(WCOC.E_2510_SELLER) Then tmpStrArr(WCOC.E_2510_SELLER) = CStr(divTempElement.innerHTML)
'                                    If i = tmpColIndxArr(WCOC.E_2510_SHIPPER) Then tmpStrArr(WCOC.E_2510_SHIPPER) = CStr(divTempElement.innerHTML)
'                                    '(status)
'                                    If i = tmpColIndxArr(WCOC.E_2510_S2) Then tmpStrArr(WCOC.E_2510_S2) = CStr(divTempElement.innerHTML)
'                                    If i = tmpColIndxArr(WCOC.E_2510_TYPE) Then tmpStrArr(WCOC.E_2510_TYPE) = CStr(divTempElement.innerHTML)
'                                    ' (UM the last one)
'                                    If i = tmpColIndxArr(WCOC.E_2510_UM) Then
'                                        tmpStrArr(WCOC.E_2510_UM) = CStr(divTempElement.innerHTML)
'
'
'
'                                        ' and this is the last one - here we can make some inital setup
'                                        ' on data...
'                                        'setup all data from i = 1 to 18
'                                        ' ==========================================================================
'                                        ' ==========================================================================
'                                        ''
'                                        '
'                                        Set tmpDH_Handler = Nothing
'                                        Set tmpDH_Handler = New DH_Handler
'
'                                        If checkIfDateAndTime(tmpStrArr(WCOC.E_2510_DHEO)) Then
'                                            tmpDH_Handler.DHEO = parseStringToDateTime(tmpStrArr(WCOC.E_2510_DHEO), CBool(francuski))
'                                        End If
'
'                                        If checkIfDateAndTime(tmpStrArr(WCOC.E_2510_DHEF)) Then
'                                            tmpDH_Handler.dhef = parseStringToDateTime(tmpStrArr(WCOC.E_2510_DHEF), CBool(francuski))
'                                        End If
'
'
'                                        If checkIfDateAndTime(tmpStrArr(WCOC.E_2510_DHAS)) Then
'                                            tmpDH_Handler.DHAS = parseStringToDateTime(tmpStrArr(WCOC.E_2510_DHAS), CBool(francuski))
'                                        End If
'
'                                        If checkIfDateAndTime(tmpStrArr(WCOC.E_2510_DHRQ)) Then
'                                            tmpDH_Handler.dhrq = parseStringToDateTime(tmpStrArr(WCOC.E_2510_DHRQ), CBool(francuski))
'                                        End If
'
'
'                                        Dim qtyAsDouble As Double
'                                        qtyAsDouble = 0
'                                        'On Error Resume Next
'                                        qtyAsDouble = CDbl(forceStringToBeDouble(tmpStrArr(WCOC.E_2510_QTY)))
'
'                                        With oItem
'                                            .coforConstructor tmpStrArr(WCOC.E_2510_SELLER), tmpStrArr(WCOC.E_2510_SHIPPER)
'                                            .dhConstructor tmpDH_Handler
'                                            .commonConstructor tmpStrArr, CDbl(qtyAsDouble)
'                                        End With
'
'
'                                        '
'                                        ''
'                                        ' ==========================================================================
'                                        ' ==========================================================================
'                                    End If
'                                End If
'
'                                i = i + 1
'                            Next c
'
'
'
'                            If oItem.checkIsObjectReady() Then
'                                sd.addOrder oItem
'                            End If
'
'
'                        Next r
'
'
'                    Next stronaOrderow
'
'
'
'
'
'
'                    'For Each oItem In sd.getOrders()
'                    '     'debug.Print "oi pn: " & oItem.getElement(E_2510_PRODUCT) & " qty: " & oItem.getElement(E_2510_QTY)
'                    'Next oItem
'                End If
'            End If
'
'
'            Set d = Nothing
'        End If
'
'
'        ' -------------------------------------------
'    Else
'        MsgBox "Wrong type of package!"
'    End If
'End Sub
'
'
' #################################################################################



                                                                               
'
' 222222222222222    77777777777777777777 222222222222222         000000000
'2:::::::::::::::22  7::::::::::::::::::72:::::::::::::::22     00:::::::::00
'2::::::222222:::::2 7::::::::::::::::::72::::::222222:::::2  00:::::::::::::00
'2222222     2:::::2 777777777777:::::::72222222     2:::::2 0:::::::000:::::::0
'            2:::::2            7::::::7             2:::::2 0::::::0   0::::::0
'            2:::::2           7::::::7              2:::::2 0:::::0     0:::::0
'         2222::::2           7::::::7            2222::::2  0:::::0     0:::::0
'    22222::::::22           7::::::7        22222::::::22   0:::::0 000 0:::::0
'  22::::::::222            7::::::7       22::::::::222     0:::::0 000 0:::::0
' 2:::::22222              7::::::7       2:::::22222        0:::::0     0:::::0
'2:::::2                  7::::::7       2:::::2             0:::::0     0:::::0
'2:::::2                 7::::::7        2:::::2             0::::::0   0::::::0
'2:::::2       222222   7::::::7         2:::::2       2222220:::::::000:::::::0
'2::::::2222222:::::2  7::::::7          2::::::2222222:::::2 00:::::::::::::00
'2::::::::::::::::::2 7::::::7           2::::::::::::::::::2   00:::::::::00
'2222222222222222222277777777            22222222222222222222     000000000
'
'
'
' #################################################################################
' #################################################################################
Private Sub innerParse2720()
    
    'If valid.makeSomeValidationOnImportedPackage(rawCorailData) Then
    
    
    
        Set sd = Nothing
    
        Dim tb As IHTMLTable
        Dim tds As IHTMLElementCollection
        Dim t As IHTMLTable
        Dim tc As IHTMLTableCell
        Dim r As HTMLTableRow
        Dim c As HTMLTableCell
        
        Dim divElement As IHTMLElement, tempTableElement As IHTMLElement
        
        Set divElement = rawCorailData.getElementById("dailyRequirement")

        
        'Dim d As ICorailData ' CorailDataFrom2720
        'Set d = rawCorailData
        
        
        If Not divElement Is Nothing Then
        
            Set collectonOfXtraCorailItems = New CorailIteration
        
            ' xtra rqm pop-up window
            ' -------------------------------------------------------------------------------------
            'Set sd = New SuitableXtraData2720 ' inside new dictionary now
            ' sd.xtraRqms
            Dim txri As CorailItem
            

            Set tb = divElement.getElementsByTagName("table")(0)
            
            
            ' <TH class=ecwTableSortable>Date</TH>
            ' <TH class=ecwTableSortable>SGR/Line</TH>
            ' <TH>CLV</TH>
            ' <TH>Fab Plan</TH>
            ' <th class="">Assembly</th>
            ' <th class="">Machining</th>
            ' <TH>Total</TH>
            howManyI = 0
            For Each r In tb.Rows
            
                i = 0
                Set txri = Nothing
        
                For Each c In r.Cells
                
                
                    ' 'debug.Print c.outerHTML & " " & i
                    
                    ' using here:
                    ' forceStringToBeDouble()
                    ' parseStringToDate()
                    
                    
                    If c.tagName = "TH" Then
                        howManyI = howManyI + 1
                    End If
                    
                    
                    
                    If c.tagName = "TD" Then
                    
                        If i = 0 Then
                            ' div date
                            
                            Set txri = New CorailItem
                            
                            strD = Trim(c.getElementsByTagName("div")(0).innerText)
                            Debug.Print (strD)
                            txri.setDate (CStr(parseStringToDate(strD)))
                        End If
                        
                        If i = 1 Then
                            ' div sgr line
                            txri.sgrLine = Trim(c.innerText)
                        End If
                        
                        If i = 2 Then
                            ' div clv
                            txri.clv = forceStringToBeDouble(Trim(c.innerText))
                        End If
                        
                        If i = 3 Then
                            ' div tab plan
                            txri.fabPlan = forceStringToBeDouble(Trim(c.innerText))
                        End If
                        
                        If i = (howManyI - 1) Then
                            ' div total
                            txri.Total = forceStringToBeDouble(Trim(c.innerText))
                        End If
                        
                        
                    End If
                    
                    i = i + 1
                
                Next c
                
                If Not txri Is Nothing Then collectonOfXtraCorailItems.addItem txri
            Next r
            
            
            
            ' -------------------------------------------------------------------------------------
        
'        ElseIf divElement Is Nothing Then
'
'            ' common data from 2720
'
'            ' contructor init also collection for coverage items
'            Set sd = New SuitableData2720
'
'            Dim inputElement As IHTMLInputElement
'
'
'            ' suitable data Part Number
'            ' ---------------------------------------------------------------------
'
'            Set inputElement = d.getDom.idoc.getElementById("detailBeanId")
'
'            ' On Error Resume Next
'            sd.pn = inputElement.Value
'            ' 'debug.Print sd.pn OK!
'            ' ---------------------------------------------------------------------
'
'
'
'
'            ' suitable data for PLT
'            ' ---------------------------------------------------------------------
'            Set inputElement = d.getDom.idoc.getElementById("stockDepart")
'            sd.stock = CDbl(forceStringToBeDouble(CStr(inputElement.Value)))
'            ' ---------------------------------------------------------------------
'
'
'            ' suitable data recv
'            ' ---------------------------------------------------------------------
'            Set inputElement = d.getDom.idoc.getElementById("receivedOnSite")
'            sd.recv = CDbl(forceStringToBeDouble(CStr(inputElement.Value)))
'            ' ---------------------------------------------------------------------
'
'
'            ' suitable data Hazards from FJU
'            ' ---------------------------------------------------------------------
'            Set inputElement = d.getDom.idoc.getElementById("aleatoriesSinceFJU")
'            sd.hazardsFromFJU = CDbl(forceStringToBeDouble(CStr(inputElement.Value)))
'            ' ---------------------------------------------------------------------
'
'
'
'
'            ' suitable data Blockages in progress
'            ' ---------------------------------------------------------------------
'            ' ---------------------------------------------------------------------
'            ' name = currentBlockingCount
'
'            Dim cbcColl As IHTMLElementCollection
'            Set cbcColl = Nothing
'
'            ' On Error Resume Next
'            Set cbcColl = d.getDom.idoc.getElementsByName("currentBlockingCount")
'
'            'If Err.number <> 0 Then
'            '    'debug.Print "The Error : " & Err.Description & vbCrLf & " was generated in " & Err.Source
'            '    Err.Clear
'            'End If
'
'            If Not cbcColl Is Nothing Then
'                For Each el In cbcColl
'
'                    Set inputElement = el
'                    If Not inputElement Is Nothing Then
'                        If inputElement.outerHTML Like "*currentBlockingCount*" Then
'
'                            sd.blockagesInProgress = CDbl(forceStringToBeDouble(CStr(inputElement.Value)))
'                            Exit For
'                        End If
'                    End If
'
'
'                Next el
'
'
'            End If
'
'            ' ---------------------------------------------------------------------
'            ' ---------------------------------------------------------------------
'
'
'
'
'
'            ' suitable data for cofors
'            ' ---------------------------------------------------------------------
'            Set inputElement = d.getDom.idoc.getElementById("psp.vendor.cofor")
'            sd.vendorCofor = CStr(inputElement.Value)
'
'            Set inputElement = d.getDom.idoc.getElementById("psp.shipper.cofor")
'            sd.shipperCofor = CStr(inputElement.Value)
'
'            Set inputElement = d.getDom.idoc.getElementById("psp.manufacturer.cofor")
'            sd.manufacturerCofor = CStr(inputElement.Value)
'
'
'            ' z supplier name jest nieco bardziej tricky
'            ' dodatkowo dorzucimy jednoczenie ua i uc
'            ' ------------------------------------------------------------------------------
'            Dim ec As IHTMLElementCollection
'            Dim tdElement As IHTMLElement
'
'            ' pod UC and UA !!!
'            ' ==================================================================================
'            Set tb = d.getDom.idoc.getElementById("tableauFluxDePiece")
'            ' w tabeli powyzej znajduja sie dodatkowo cztery sub-tabele - wybieramy ostatnia
'            ' wg tego co widac w html to wlasnie w ostatniej znajduje sie supplier name
'            ' to jest druga kolumna, gdzie znajduja sie packi
'            Set tb = tb.getElementsByTagName("table")(1) ' jako celle 4tej sub-tabeli html
'            Set ec = tb.getElementsByTagName("td")
'            ' ==================================================================================
'
'
'            iterator = 1
'            For Each tdElement In ec
'
'                ' 'debug.Print tdElement.Style & " outer: " & tdElement.outerHTML
'
'                If UCase(tdElement.innerHTML) Like "*UC*UC*" Or iterator = 1 Then
'                    sd.uc = Application.WorksheetFunction.Trim(tdElement.innerHTML)
'                End If
'
'                If UCase(tdElement.innerHTML) Like "*UA*UA*" Or iterator = 2 Then
'                    sd.ua = Application.WorksheetFunction.Trim(tdElement.innerHTML)
'                End If
'
'                If iterator = 3 Then
'                    sd.dock = Application.WorksheetFunction.Trim(tdElement.innerHTML)
'                End If
'
'
'                iterator = iterator + 1
'            Next
'
'
'            ' pod supplier name !!!
'            ' ==================================================================================
'            Set tb = d.getDom.idoc.getElementById("tableauFluxDePiece")
'            ' w tabeli powyzej znajduja sie dodatkowo cztery sub-tabele - wybieramy ostatnia
'            ' wg tego co widac w html to wlasnie w ostatniej znajduje sie supplier name
'            ' Set tb = tb.getElementsByTagName("table")(3) ' jako celle 4tej sub-tabeli html
'            Set ec = tb.getElementsByTagName("td")
'            ' ==================================================================================
'            ' szukamy teraz statycznego formotowania, ktore jest specyficzne dla nazwy dostawcy
'            ' oczywiscie heurystycznie, wiec nie zdziwie sie jak kiedys przestanie dzialac poprawnie :D
'            ' szukamy boldu i dopasowania textu:
'            ' font-weight: bold; text-align: left
'            ' Dim tdElement As IHTMLElement ' look up
'            iterator = 1
'            For Each tdElement In ec
'
'                ' 'debug.Print tdElement.Style & " outer: " & tdElement.outerHTML
'
'                If UBound(Split(tdElement.outerHTML, Chr(10))) < 2 Then
'
'                    If tdElement.outerHTML Like "*FONT-WEIGHT: bold*TEXT-ALIGN: left*colSpan=2*" Then
'                        ' On Error Resume Next
'                        sd.supplierName = Trim(Left(tdElement.innerHTML, 100))
'                        sd.supplierName = focusOnSupplierName(Trim(sd.supplierName))
'                        Exit For
'                    End If
'                End If
'
'                iterator = iterator + 1
'            Next
'
'
'            ' get full table from top of the 2720
'            ' !!!
'            ' table#tableauIdentificationProduct
'            ''
'            '
'            '
'
'
'            Set tb = d.getDom.idoc.getElementById("tableauIdentificationProduct")
'
'
'            ' .getDom.doc.getElementsByTagName
'            Set ec = tb.getElementsByTagName("input")
'
'            iterator = 1
'            For Each inputElement In ec
'                ' 'debug.Print inputElement.Value & " " & inputElement.outerHTML
'
'
'                If iterator = 6 Then
'                    ' (ID appro)
'                    If Trim(inputElement.Value) <> "" Then
'                        sd.approId = Split(inputElement.Value, " - ")(0)
'                        sd.approName = Split(inputElement.Value, " - ")(1)
'
'                        ' 'debug.Print sd.approId & " 666 " & sd.approName ' OK!
'                    End If
'
'                ElseIf iterator = 7 Then
'                    ' (Area)
'                    sd.area = inputElement.Value
'
'                ElseIf iterator = 8 Then
'                    ' (Transfert in product date)
'                    sd.tipd = inputElement.Value
'                ElseIf iterator = 11 Then
'                    ' (CMJ JL Hebdo)
'                    sd.hebdo = inputElement.Value
'                ElseIf iterator = 12 Then
'                    '(CMJ JL Périodique)
'                    sd.periodique = inputElement.Value
'                ElseIf iterator = 13 Then
'                    sd.upstreamFlow = inputElement.Value
'                ElseIf iterator = 14 Then
'                    sd.internals = inputElement.Value
'                End If
'
'                iterator = iterator + 1
'            Next
'
'            ' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'            ' in ec indxs on 23 of may 2019:
'            ' 1: <INPUT id=detailBeanId maxLength=20 size=15 value=9810867580>
'            ' 2: <INPUT id=hidDetailBeanId type=hidden name=detailBeanId>
'            ' 3: <INPUT id=sgrCode type=hidden name=sgrCode>
'            ' 4: <INPUT id=lineCode type=hidden name=lineCode>
'            ' 5: <INPUT id=lineId type=hidden name=lineId>
'            ' 6: P526139 - LAMBERT GEORGES <INPUT class=number readOnly maxLength=20 size=15 value="P526139 - LAMBERT GEORGES"> (ID appro)
'            ' 7: E41 <INPUT class=number readOnly maxLength=20 size=15 value=E41> (Area)
'            ' 8: 14/06/2016 <INPUT class=number readOnly maxLength=20 size=15 value=14/06/2016> (Transfert in product date)
'            ' 9: <INPUT class=number readOnly maxLength=20 size=15>
'            ' 10: <INPUT class=number readOnly maxLength=20 size=15>
'            ' 11: 2544.0 <INPUT class=number readOnly maxLength=20 size=15 value=2544.0> (CMJ JL Hebdo)
'            ' 12: 2632.0 <INPUT class=number readOnly maxLength=20 size=15 value=2632.0> (CMJ JL Périodique)
'            ' 13: 9 <INPUT class=number readOnly maxLength=20 size=15 value=9> (Upstream Flow)
'            ' 14 743 <INPUT class=number readOnly maxLength=20 size=15 value=743> (Internals)
'            ' 15: 2019-05-16T07:08:14Z <INPUT id=internalAlertBeginDate type=hidden value=2019-05-16T07:08:14Z name=internalAlertBeginDate>
'            ' 16: 2019-05-23T07:08:14Z <INPUT id=internalAlertEndDate type=hidden value=2019-05-23T07:08:14Z name=internalAlertEndDate>
'            ' 17: on <INPUT disabled CHECKED type=checkbox value="">
'            ' 18: on <INPUT disabled CHECKED type=checkbox value="">
'            ' 19: on <INPUT disabled CHECKED type=checkbox value="">
'
'            ''
'            '
'            '
'
'
'
'
'            ' -------------------------------------------
'
'
'            ' TEXTAREA MAJOR COMMENT (LAST UPDATE):
'            Dim textAreaElement As IHTMLElement
'
'            Set textAreaElement = d.getDom.idoc.getElementById("getProductSummaryRead_detailBean_productSummaryNoteList_lastProductSummaryNotes_noteComments")
'            sd.majorComment = textAreaElement.innerText
'            ' 'debug.Print sd.majorComment ' OK!
'
'
'
'
'            ' tables of rqms, transit end disruption
'            ' ----------------------------------------------------------------------------------------------
'            ' ----------------------------------------------------------------------------------------------
'            ' ----------------------------------------------------------------------------------------------
'            ''
'            ''
'            '
'
'            Dim mainTableDiv As IHTMLElement
'            Set mainTableDiv = d.getDom.idoc.getElementById("demandAndSupplyPerDay")
'
'            ' now inside this mainTableDiv we have another div.ecwTableHead on top and on bottom div.ecwTableBody
'            ' those 2 have direct connection with tables which includes data of our interests :)
'
'            Dim maintables As IHTMLElementCollection
'            Dim divTableHead As IHTMLElement, divTableBody As IHTMLElement, divTempElement As IHTMLElement
'            'Set maintables = d.getDom.idoc.getElementsByClassName("ecwTableHead")
'            Set maintables = mainTableDiv.getElementsByTagName("div")
'
'            For Each divTempElement In maintables
'
'                If divTempElement.className = "ecwTableHead" Then
'                    Set divTableHead = divTempElement
'                End If
'
'                If divTempElement.className = "ecwTableBody" Then
'                    Set divTableBody = divTempElement
'                End If
'
'            Next
'
'
'
'            ' 'debug.Print divTableBody.innerHTML ' OK
'            ' ecwTableHead should have exactly same data as ecwTableBody - and I do not know why
'            ' but at the I'll go with body
'
'
'
'            Set t = divTableBody.getElementsByTagName("table")(0)
'
'
'
'            Dim tmpDate As String, tmpRqm1 As String, tmpRqm2 As String, tmpOrderStr As String, tmpShipStr As String
'            Dim tmpDateDate As Date, tmpRqm1Dbl As Double, tmpRqm2Dbl As Double, tmpOrderStrDbl As Double, tmpShipStrDbl As Double
'
'
'            Dim covItem As CoverageItem
'
'            For Each r In t.Rows
'
'                Set covItem = Nothing
'                Set covItem = New CoverageItem
'
'                For Each c In r.Cells
'                    ''debug.Print c.outerHTML & ", class: " & c.className & " inner: " & c.innerHTML & " tag: " & c.tagName
'                    ' <TD ecwKeyName="date"><DIV class=date>22/05/2019</DIV></TD>, class:  inner: <DIV class=date>22/05/2019</DIV> tag: TD
'                    If c.tagName = "TD" Then
'                        Set divTempElement = c.getElementsByTagName("div")(0)
'
'                        ' div.innerHTML: 22/05/2019 class: [date] (DATE)
'                        ' div.innerHTML: 113.0 class: [number fbpcDetailByDay] (PC DEMAND)
'                        ' div.innerHTML: 0.0 class: [number] (Incl. BC) - not suff -> ?
'                        ' div.innerHTML: 0.0 class: [ecwButtonTexteOverOrdered number] ' <TD ecwKeyName="orderedQuantity"
'                        ' div.innerHTML: 0.0 class: [ecwButtonTexteOverShipped number] ' <TD ecwKeyName="shippedQuantity"
'                        ' div.innerHTML: &nbsp; class: [dateRupture]
'                        ' div.innerHTML: &nbsp; class: [dateRupture]
'                        ' div.innerHTML: &nbsp; class: [dateRupture] ' <TD ecwKeyName="formattedRelativeDisruptionDate"
'
'                        ' 'debug.Print "div.innerHTML: " & divTempElement.innerHTML & " class: [" & divTempElement.className & "]"
'
'                        If divTempElement.className = "date" Then
'                            ' date
'                            tmpDate = divTempElement.innerText
'                            tmpDateDate = parseStringToDate(tmpDate)
'
'
'
'                        ElseIf divTempElement.className = "number fbpcDetailByDay" Then
'                            ' pc demand
'                            tmpRqm1 = divTempElement.innerText
'                            tmpRqm1Dbl = CDbl(forceStringToBeDouble(tmpRqm1))
'
'                        ElseIf divTempElement.className = "number" And c.outerHTML Like "*requiredBCQuantity*" Then
'                            ' incl bc
'                            tmpRqm2 = divTempElement.innerText
'                            tmpRqm2Dbl = CDbl(forceStringToBeDouble(tmpRqm2))
'
'
'                        ElseIf divTempElement.className Like "ecwButtonTexteOverOrdered*" Then
'                            ' ordered
'                            tmpOrderStr = divTempElement.innerText
'                            tmpOrderStrDbl = CDbl(forceStringToBeDouble(tmpOrderStr))
'
'                        ElseIf divTempElement.className Like "ecwButtonTexteOverShipped*" Then
'                            ' shipped
'                            tmpShipStr = divTempElement.innerText
'                            tmpShipStrDbl = CDbl(forceStringToBeDouble(tmpShipStr))
'
'                            covItem.makeReady
'                            Exit For
'
'                        End If
'
'
'                    Else
'                        Exit For
'                    End If
'                Next c
'
'
'                If covItem.isReady() Then
'                    covItem.constructor CDate(tmpDateDate), CDbl(tmpRqm1Dbl), CDbl(tmpRqm2Dbl), CDbl(tmpOrderStrDbl), CDbl(tmpShipStrDbl)
'                    sd.addCovLine covItem
'                End If
'            Next r
'
            
            
            
            ' OK!
            ' For Each covItem In sd.getCollectionOfItems()
            '     'debug.Print "ci : " & covItem.getStrDate & " r " & covItem.getRqms & " t " & " o " & covItem.getOrder & covItem.getShip
            ' Next covItem
            
            '
            ''
            ''
            ' ----------------------------------------------------------------------------------------------
            ' ----------------------------------------------------------------------------------------------
            ' ----------------------------------------------------------------------------------------------
        'End If
        
        
        
        
        '
        
        
        
        Set tb = Nothing
        Set tds = Nothing
        Set t = Nothing
        Set tc = Nothing
        Set r = Nothing
        Set c = Nothing
        
        Set divElement = Nothing
        Set tempTableElement = Nothing
        
        ' big fish?
        Set d = Nothing
        
    Else
        MsgBox "Wrong type of package!"
    End If
End Sub
'
'
' #################################################################################


Private Function checkIfDate(strD) As Boolean

    checkIfDate = False
    
    If Len(strD) = 10 Then
    
        If Mid(strD, 3, 1) = "/" Then
            If Mid(strD, 6, 1) = "/" Then
                checkIfDate = True
            End If
        End If
    End If
End Function





Private Function forceStringToBeDouble(arg As String) As Double

    ' Application.DecimalSeparator
    Dim convertedValue As String
    convertedValue = CStr(arg)
    ' double sure shot! :D
    
    ' just in case, if empty then change to string "0"
    If Trim(convertedValue) = "" Then
        convertedValue = "0"
    End If
    
    
    convertedValue = Replace(convertedValue, ",", Application.DecimalSeparator)
    convertedValue = Replace(convertedValue, ".", Application.DecimalSeparator)
    
    ' On Error Resume Next
    If IsNumeric(convertedValue) Then
        forceStringToBeDouble = CDbl(convertedValue)
    End If
    
    'If Err.number <> 0 Then
    '    'debug.Print "The Error : " & Err.Description & vbCrLf & " was generated in " & Err.Source
    '    Err.Clear
    'End If
    
    If CStr(convertedValue) <> "0" And forceStringToBeDouble = 0 Then
    
        ' forceStringToBeDouble = convertedValue
        
        If Left(convertedValue, 1) = "-" Then
        
            convertedValue = Mid(convertedValue, 2, Len(convertedValue) - 3)
            forceStringToBeDouble = Int(convertedValue) * (-1)
            
        
        Else
        
            convertedValue = Mid(convertedValue, 1, Len(convertedValue) - 2)
            forceStringToBeDouble = CDbl(convertedValue)
        
        End If
    End If
    
    
End Function

Private Function parseStringToDate(strD) As Date
    
    ' format MM/DD/YYYY
    
    Dim d As Integer
    Dim m As Integer
    Dim y As Integer
    
    If checkIfDate(strD) Then
    
        If IsNumeric(Left(strD, 2)) Then
            d = Int(Left(strD, 2))
        End If
        
        If IsNumeric(Mid(strD, 4, 2)) Then
            m = Int(Mid(strD, 4, 2))
        End If
        
        If IsNumeric(Mid(strD, 7, 4)) Then
            y = Int(Right(strD, 4))
        End If
    
    End If
    
    Dim s As String
    s = Format(y & "-" & m & "-" & d, "yyyy-mm-dd")
    
    parseStringToDate = CDate(CStr(s))
    
End Function

'NOT USED, FROM WEEKLY COVERAGE
'Private Function parseStringToDateTime(strDT, Optional isFrance As Boolean) As Date
'
'
'    ' angielski
'    ' format 19/05/29 10:20
'    ' format yy/mm/dd hh:mm
'    '
'    '
'    ' francuski
'    ' format dd/mm/yy hh:mm
'
'
'    Dim d1 As Integer
'    Dim m1 As Integer
'    Dim y1 As Integer
'
'    Dim d2 As Integer
'    Dim m2 As Integer
'    Dim y2 As Integer
'
'    Dim hh As Integer
'    Dim mm As Integer
'
'    Dim dzis As Date
'    dzis = Date
'
'
'
'    If Len(strDT) = 14 Then
'
'
'        If IsNumeric(Left(strDT, 2)) Then
'            d1 = Int(Left(strDT, 2))
'        End If
'
'        If IsNumeric(Mid(strDT, 4, 2)) Then
'            m1 = Int(Mid(strDT, 4, 2))
'        End If
'
'        If IsNumeric(Mid(strDT, 7, 2)) Then
'            y1 = Int(Mid(strDT, 7, 2))
'        End If
'
'        If IsNumeric(Left(strDT, 2)) Then
'            y2 = Int(Left(strDT, 2))
'        End If
'
'        If IsNumeric(Mid(strDT, 4, 2)) Then
'            m2 = Int(Mid(strDT, 4, 2))
'        End If
'
'        If IsNumeric(Mid(strDT, 7, 2)) Then
'            d2 = Int(Mid(strDT, 7, 2))
'        End If
'
'        If IsNumeric(Mid(strDT, 10, 2)) Then
'            hh = Int(Mid(strDT, 10, 2))
'        End If
'
'        If IsNumeric(Mid(strDT, 13, 2)) Then
'            mm = Int(Mid(strDT, 13, 2))
'        End If
'
'    ' prototypowa logika sprawdzajaca, czy wygenerowana data jest w miare blisko daty dzisiejszej
'    '
'
'        Dim s1 As String, S2 As String
'        Dim SD1 As Date, SD2 As Date
'
'        s1 = Format(y1 & "-" & m1 & "-" & d1 & " " & hh & ":" & mm, "yy-mm-dd hh:mm")
'        S2 = Format(y2 & "-" & m2 & "-" & d2 & " " & hh & ":" & mm, "yy-mm-dd hh:mm")
'
'        SD1 = CDate(s1)
'        SD2 = CDate(S2)
'
'        diff1 = Math.Abs(dzis - SD1)
'        diff2 = Math.Abs(dzis - SD2)
'
'        If diff1 < diff2 Then
'            parseStringToDateTime = CDate(CStr(SD1))
'        Else
'            parseStringToDateTime = CDate(CStr(SD2))
'        End If
'
'    ElseIf Len(strDT) = 16 Then
'
'
'        If IsNumeric(Left(strDT, 2)) Then
'            d1 = Int(Left(strDT, 2))
'        End If
'
'        If IsNumeric(Mid(strDT, 4, 2)) Then
'            m1 = Int(Mid(strDT, 4, 2))
'        End If
'
'        If IsNumeric(Mid(strDT, 7, 4)) Then
'            y1 = Int(Mid(strDT, 7, 4))
'        End If
'
'
'        Dim d16size As Date
'        d16size = Date
'        d16size = CDate(Format(y1 & "-" & m1 & "-" & d1, "yyyy-mm-dd"))
'        parseStringToDateTime = CDate(d16size)
'
'
'    End If
'
'
'
'End Function
''NOT USED, FROM WEEKLY COVERAGE
'Private Function checkIfDateAndTime(strDT) As Boolean
'
'    checkIfDateAndTime = False
'
'    If Len(strDT) = 14 Or Len(strDT) = 16 Then
'
'        If Mid(strDT, 3, 1) = "/" Then
'            If Mid(strDT, 6, 1) = "/" Then
'                If Mid(strDT, 9, 1) = " " Or Mid(strDT, 11, 1) = " " Then
'                    If Mid(strDT, 12, 1) = ":" Or Mid(strDT, 14, 1) = ":" Then
'                        checkIfDateAndTime = True
'                    End If
'                End If
'            End If
'        End If
'    End If
'End Function

' OBSOLETE
'Private Function replaceAllDecimals(s) As String
'
'    s = Replace(s, ".0", "")
'    s = Replace(s, ",0", "")
'
'    s = Replace(s, ".1", "")
'    s = Replace(s, ",1", "")
'
'    s = Replace(s, ".2", "")
'    s = Replace(s, ",2", "")
'
'    s = Replace(s, ".3", "")
'    s = Replace(s, ",3", "")
'
'    s = Replace(s, ".4", "")
'    s = Replace(s, ",4", "")
'
'    s = Replace(s, ".5", "")
'    s = Replace(s, ",5", "")
'
'    s = Replace(s, ".6", "")
'    s = Replace(s, ",6", "")
'
'    s = Replace(s, ".7", "")
'    s = Replace(s, ",7", "")
'
'    s = Replace(s, ".8", "")
'    s = Replace(s, ",8", "")
'
'    s = Replace(s, ".9", "")
'    s = Replace(s, ",9", "")
'
'    replaceAllDecimals = CStr(s)
'
'End Function
