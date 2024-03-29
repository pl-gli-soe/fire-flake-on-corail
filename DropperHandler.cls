VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DropperHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'
'Copyright (c) 2017 FORREST
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


Private dane As ConvertedData
Private plt As String
Private pn As String
Private cmnt1 As String
Private cmnt2 As String

Public Sub init(argPlt, argPn, argCmnt1, argCmnt2, argDane As ConvertedData)

    Set dane = argDane
    plt = argPlt
    pn = argPn
    cmnt1 = argCmnt1
    cmnt2 = argCmnt2
End Sub


Public Sub run(Sh As Worksheet, repRow As Long, Optional BO As String)

    Dim r As Range
    
    
    
    'pn
    Sh.Cells(repRow, 2 + FFOC.E_COMMON_PN - 1).Value = pn
    'plt
    Sh.Cells(repRow, 2 + FFOC.E_COMMON_PLT_CODE - 1).Value = plt
    ' plt name
    Sh.Cells(repRow, 2 + FFOC.E_COMMON_PLT_NAME - 1).Value = getPlantName(plt)
    
    If Not dane Is Nothing Then
        ' comments
        If Trim(CStr(dane.iteration.getComment1())) <> "" Then
            Sh.Cells(repRow, 2 + FFOC.E_COMMON_CMNT1 - 1).Value = cmnt1 & ", " & CStr(dane.iteration.getComment1())
        Else
            Sh.Cells(repRow, 2 + FFOC.E_COMMON_CMNT1 - 1).Value = cmnt1
        End If
        Sh.Cells(repRow, 2 + FFOC.E_COMMON_CMNT2 - 1).Value = cmnt2
    Else
        Sh.Cells(repRow, 2 + FFOC.E_COMMON_CMNT1 - 1).Value = cmnt1
        Sh.Cells(repRow, 2 + FFOC.E_COMMON_CMNT2 - 1).Value = cmnt2
    End If
    
    
    
    If Not dane Is Nothing Then
        'stock + recv - hazards !!! but only when checkbox avail.
        'If FFOC.G_HAZARDS Then
        '
        '
        '    ' stock depart and hazards
        '    sh.Cells(repRow, 2 + FFOC.E_COMMON_STOCK - 1).AddComment "STOCK: " & _
        '    CStr(CDbl(dane.iteration.getStock)) & chr(10) & _
        '    "HAZARDS: " & CStr(CDbl(dane.iteration.getAleas))
        '
        '    sh.Cells(repRow, 2 + FFOC.E_COMMON_STOCK - 1).Value = CDbl(dane.iteration.getStock) + CDbl(dane.iteration.getAleas)
        '
        'Else
        '    ' stock depart only
        '    sh.Cells(repRow, 2 + FFOC.E_COMMON_STOCK - 1).Value = CDbl(dane.iteration.getStock)
        'End If
        
        ' ^^^^^^^^^^^^^
        ' stock depart only - starting from version 0.6 hazards will have seperate column!
        Sh.Cells(repRow, 2 + FFOC.E_COMMON_Hazards - 1).Value = CDbl(dane.iteration.getAleas)
        Sh.Cells(repRow, 2 + FFOC.E_COMMON_STOCK - 1).Value = CDbl(dane.iteration.getStock)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Sh.Cells(repRow, 2 + FFOC.E_COMMON_RECV - 1).Value = CDbl(dane.iteration.getRecv)
        Sh.Cells(repRow, 2 + FFOC.E_COMMON_BACKLOG - 1).Value = CDbl(dane.iteration.getBacklog)
        
        ' NEW 21 OCT 2019
        Sh.Cells(repRow, 2 + FFOC.E_COMMON_Blockages_in_progress - 1).Value = dane.iteration.getBlockagesInProgress
        
        Sh.Cells(repRow, 2 + FFOC.E_COMMON_PART_NAME - 1).Value = dane.iteration.getPartName
        Sh.Cells(repRow, 2 + FFOC.E_COMMON_SUPPLIER - 1).Value = dane.iteration.getSupplier
        
        Sh.Cells(repRow, 2 + FFOC.E_COMMON_RESP - 1).Value = dane.iteration.getresp
        
    Else
        Sh.Cells(repRow, 2 + FFOC.E_COMMON_STOCK - 1).Value = 0#
        Sh.Cells(repRow, 2 + FFOC.E_COMMON_BACKLOG - 1).Value = 0#
        
    End If
    
    
    If BO = "" Then
    
    
        Dim item As CorailItem, gdzieJestRecvOffsetowo As Integer
        gdzieJestRecvOffsetowo = -4
        
        If Not dane Is Nothing Then
        
        
            Dim tbInTxtBox As TableInTextBox
            Dim txtForCmmnt As String
            Dim ci2510 As CorailItem2510
            
            Set r = Sh.Cells(repRow, FFOC.E_COMMON_FIRST_RQM + 1)
            i = 0
            Do
                dstr = CStr(r.Parent.Cells(3, r.Column).Value)
                
                If dstr = "" Then
                    Exit Do
                End If
                
                If CDate(dstr) < Date Then
                    gdzieJestRecvOffsetowo = gdzieJestRecvOffsetowo - 4
                End If
                
                If Trim(dstr) <> "" Then
                
                    Application.enableEvents = False
            
                    For Each item In dane.iteration.pItems
                    
                        
                    
                        If matchDates(dstr, item.getStringDate) Then
                        
                            
                    
                            r.Value = CDbl(item.besoinsBC + item.besoinsPC)
                            r.Offset(0, 1).Value = CDbl(item.order)
                            r.Offset(0, 2).Value = CDbl(item.ship)
                            
                            
                            If item.getCollection2510.Count > 0 Then
                            
                                txtForCmmnt = ""
                            
                                Set tbInTxtBox = New TableInTextBox
                                With tbInTxtBox
                                    .defineLables "ORDERED", "PICKUP", "DELIVERY", "QTY", "TYPE", "NM"
                                    .defineSizes 20, 20, 20, 12, 12, 14
                                    .setCols 6
                                End With
                                
                                
                                tbInTxtBox.tryToAddLineOfTheData "       DHEO         ", "       DHEF         ", _
                                    "       DHRQ         ", "    QTY     ", "    TYPE    ", "    NM        "
                                ' tbInTxtBox.tryToAddLineOfTheData "ORDERED", "PICKUP", "DELIVERY", "QTY", "TYPE", "NM"
                                
                                For Each ci2510 In item.getCollection2510
                                    tbInTxtBox.tryToAddLineOfTheData forceChangeFormat(ci2510.dhxx.DHEO), _
                                        forceChangeFormat(ci2510.dhxx.dhef), _
                                        forceChangeFormat(ci2510.dhxx.dhrq), _
                                        forceChangeFormat(ci2510.qty, True), _
                                        forceChangeFormat(ci2510.itemType), _
                                        forceChangeFormat(ci2510.orderNum)
                                Next ci2510
                                
                                txtForCmmnt = tbInTxtBox.drawThisTable()
                                
                                r.Offset(0, 1).AddComment txtForCmmnt
                                On Error Resume Next
                                r.Offset(0, 1).Comment.Shape.TextFrame.AutoSize = True
                                r.Offset(0, 1).Comment.Shape.TextFrame.Characters.Font.Name = "Courier New"
                                r.Offset(0, 1).Comment.Shape.TextFrame.Characters.Font.Size = 8
                            
                            End If
                            
                            
                            Exit For
                        End If
        
                    
                    Next item
                
                
                    ' so now this line is really important from balance perspective - need to adjust it properly
                    ' first of all need to check if it is a very first element of the item so ...
                    ' another thing is to add recv dynamically to the today date
                    ' Debug.Print dstr
                    
                    If r.Column = (FFOC.E_COMMON_FIRST_RQM + 1) And CDate(dstr) < Date Then
                        ' -RC[-6] new on 21 oct 2019 - for Blockages_in_progress
                        If FFOC.G_HAZARDS Then
                            r.Offset(0, 3).FormulaR1C1 = "=RC[-1]-RC[-3]+RC[-5]+RC[-6]-RC[-7]"
                        Else
                            r.Offset(0, 3).FormulaR1C1 = "=RC[-1]-RC[-3]+RC[-5]-RC[-7]"
                        End If
                        
                    ElseIf r.Column = (FFOC.E_COMMON_FIRST_RQM + 1) And CDate(dstr) >= Date Then
                        r.Offset(0, 3).FormulaR1C1 = "=RC[-1]-RC[-3]+RC[-5]+RC[" & CStr(gdzieJestRecvOffsetowo) & "]"
                        
                    ElseIf CDate(dstr) = Date Then
                        ' addtionally adding dynamically recv param
                        r.Offset(0, 3).FormulaR1C1 = "=RC[-1]-RC[-3]+RC[-4]+RC[" & CStr(gdzieJestRecvOffsetowo) & "]"
                        
                    Else
                        r.Offset(0, 3).FormulaR1C1 = "=RC[-1]-RC[-3]+RC[-4]"
                        
                    End If
                    
                    
                    Application.enableEvents = True
                
                    i = i + 1
                
                    Set r = r.Offset(0, G_ITEM_OFFSET)
                Else
                    Exit Do
                End If
            Loop Until CStr(dstr) = ""
        Else
            
            Set r = Sh.Cells(repRow, FFOC.E_COMMON_FIRST_RQM + 1)
            Do
            
                dstr = CStr(r.Parent.Cells(3, r.Column).Value)
                
                If dstr = "" Then
                    Exit Do
                End If
                
                
                If CDate(dstr) < Date Then
                    gdzieJestRecvOffsetowo = gdzieJestRecvOffsetowo - 4
                End If
                
                If Trim(dstr) <> "" Then
                    Application.enableEvents = False
                
                    If r.Column = (FFOC.E_COMMON_FIRST_RQM + 1) And CDate(dstr) < Date Then
                        r.Offset(0, 3).FormulaR1C1 = "=RC[-1]-RC[-3]+RC[-5]"
                        
                    ElseIf r.Column = (FFOC.E_COMMON_FIRST_RQM + 1) And CDate(dstr) >= Date Then
                        r.Offset(0, 3).FormulaR1C1 = "=RC[-1]-RC[-3]+RC[-5]+RC[" & CStr(gdzieJestRecvOffsetowo) & "]"
                        
                    ElseIf CDate(dstr) = Date Then
                        ' addtionally adding dynamically recv param
                        r.Offset(0, 3).FormulaR1C1 = "=RC[-1]-RC[-3]+RC[-4]+RC[" & CStr(gdzieJestRecvOffsetowo) & "]"
                        
                    Else
                        r.Offset(0, 3).FormulaR1C1 = "=RC[-1]-RC[-3]+RC[-4]"
                        
                    End If
                    
                    
                    Application.enableEvents = True
                    Set r = r.Offset(0, G_ITEM_OFFSET)
                    
                    
                Else
                    Exit Do
                End If
            Loop Until CStr(dstr) = ""
        End If
    End If
End Sub



Private Function forceChangeFormat(xd As Variant, Optional isQty As Boolean) As String



    If isQty Then
    
        forceChangeFormat = CStr(xd)
        
    ElseIf IsDate(xd) Then
        
        forceChangeFormat = Format(xd, "yyyy-mm-dd hh:mm")
    Else
        forceChangeFormat = CStr(xd)
    End If
End Function


Private Function matchDates(d1 As Variant, d2 As Variant) As Boolean
    matchDates = False
    
    innerd1 = CDate(d1)
    innerd2 = CDate(d2)
    
    
    If Year(innerd1) = Year(innerd2) And Month(innerd1) = Month(innerd2) And Day(innerd1) = Day(innerd2) Then
        matchDates = True
    End If
End Function

Public Function getFirstDate() As Date

    If Not dane Is Nothing Then
    
        Dim item As CorailItem
        Set item = Nothing
        If dane.iteration.pItems.Count > 0 Then Set item = dane.iteration.pItems(1)
        Dim tmpDate As Date
        
        If Not item Is Nothing Then
        
            If CLng(item.getDate()) > 0 Then
            
                tmpDate = CDate(item.getDate())
                
            
                For Each item In dane.iteration.pItems
                    
                    If CDate(item.getDate) < CDate(tmpDate) Then
                        tmpDate = item.getDate
                    End If
                    
                Next item
                
                getFirstDate = tmpDate
            Else
                getFirstDate = Date
            End If
            
        Else
            getFirstDate = Date
        End If
        
    Else
        getFirstDate = Date
    End If
End Function


Public Function getLastDate() As Date

    If Not dane Is Nothing Then
    
        Dim item As CorailItem
        Set item = Nothing
        If dane.iteration.pItems.Count > 0 Then Set item = dane.iteration.pItems(1)
        Dim tmpDate As Date
        
        ' Debug.Print item.getDate() & " " & CLng(item.getDate())
        If Not item Is Nothing Then
        
            If CLng(item.getDate()) > 0 Then
            
                tmpDate = CDate(item.getDate())
                
            
                For Each item In dane.iteration.pItems
                    
                    If CDate(item.getDate) > CDate(tmpDate) Then
                        tmpDate = item.getDate
                    End If
                    
                Next item
                
                getLastDate = tmpDate
            Else
                getLastDate = Date
            End If
        Else
            getLastDate = Date
        End If
        
    Else
        getLastDate = Date
    End If
End Function
