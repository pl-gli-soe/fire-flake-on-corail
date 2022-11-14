VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CorailItem2510"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'The MIT License (MIT)
'
'Copyright (c) 2020 FORREST
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

Public pn As String
Public dhxx As CorailItemDH
Public qty As Double
Public status As String
Public orderNum As String
Public seller As String
Public shipper As String
Public routing As String
Public sid As String
Public itemType As String


Private Sub Class_Initialize()
    Set dhxx = New CorailItemDH
End Sub

Private Sub Class_Terminate()
    Set dhxx = Nothing
End Sub



Public Function parseQtyFrom2510(strQty As String) As Double
    
    parseQtyFrom2510 = 0
    
    Dim tmpstr As String
    Dim separatorString As String
    
    If IsNumeric(strQty) Then
    
        tmpstr = Replace(strQty, ".", Application.DecimalSeparator)
        parseQtyFrom2510 = CDbl(tmpstr)
        
    ElseIf strQty Like "*.*" Then
        
        'tmpStr = Replace(strQty, ".", Application.DecimalSeparator)
        'parseQtyFrom2510 = CDbl(tmpStr)
        
        
        separatorString = Mid(ThisWorkbook.Sheets("register").Range("Q17").Value, 2, 1)
        
        tmpstr = Replace(strQty, ".", separatorString)
        
        
        ' new in 076
        ' look like tmpstr is still not a valid number
        If IsNumeric(tmpstr) Then
            
            
            ' just assign - OK validation
            parseQtyFrom2510 = CDbl(tmpstr)
            
        Else
        
           ' still something wrong #1
            
            ' try static replace - very specific situation
            tmpstr = Replace(tmpstr, "</div", "")
            
            If IsNumeric(tmpstr) Then
                
                ' just assign - OK validation
                parseQtyFrom2510 = CDbl(tmpstr)
            Else
            
                ' still something wrong #2 - the second layer of trying catching an issue
                parseQtyFrom2510 = CDbl(loopTroughTextAndTryToReceiveNumber(tmpstr))
                
            End If
            
        End If
        
        
        
        
        
    ElseIf strQty Like "*,*" Then
        
        'tmpStr = Replace(strQty, ".", Application.DecimalSeparator)
        'parseQtyFrom2510 = CDbl(tmpStr)
        
        
        separatorString = Mid(ThisWorkbook.Sheets("register").Range("Q17").Value, 2, 1)
        
        tmpstr = Replace(strQty, ",", separatorString)
        parseQtyFrom2510 = CDbl(tmpstr)
        
        
    End If
End Function



Private Function loopTroughTextAndTryToReceiveNumber(S1 As String) As Double
    
    loopTroughTextAndTryToReceiveNumber = 0
    
    
    Dim possibleAnswer As Double, tempPossibleStr As String
    possibleAnswer = 0
    tempPossibleStr = ""
    
    Dim x As Variant
    For x = 1 To Len(S1)
    
    
        If IsNumeric(Mid(S1, x, 1)) Or (Mid(S1, x, 1) = ".") Or (Mid(S1, x, 1) = ",") Then
            tempPossibleStr = tempPossibleStr & Mid(S1, CLng(x), 1)
        End If
        
    
    Next x
    
    tempPossibleStr = Replace(tempPossibleStr, ".", Application.DecimalSeparator)
    tempPossibleStr = Replace(tempPossibleStr, ",", Application.DecimalSeparator)
    
    If IsNumeric(tempPossibleStr) Then
        possibleAnswer = CDbl(tempPossibleStr)
    End If
    
    
    
End Function
