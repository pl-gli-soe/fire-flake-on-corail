VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CorailItem"
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



Public sgrLine As String
Public clv As Double
Public fabPlan As Double
Public Total As Double

' date of item
Private d As Date
' required qty
Public besoinsPC As Double
' required qty (BC)
Public besoinsBC As Double
' order - ordered Qty
Public order As Double
' ship - shipped Qty
Public ship As Double

Private lineFrom2510 As CorailItem2510
Private collectionOfCorailItemsFrom2510 As Collection

Public Sub addItemFrom2510(obj As CorailItem2510)
    
    If Not collectionOfCorailItemsFrom2510 Is Nothing Then
        collectionOfCorailItemsFrom2510.Add obj
    Else
        Set collectionOfCorailItemsFrom2510 = New Collection
        collectionOfCorailItemsFrom2510.Add obj
    End If
End Sub

Public Function getCollection2510() As Collection
    Set getCollection2510 = collectionOfCorailItemsFrom2510
End Function



Public Function getDate() As Date
    
    getDate = CDate(d)
End Function


Public Sub setDate(arg As String)

    d = CDate(arg)
End Sub


Public Function getStringDate() As String

    getStringDate = CStr(d)
End Function

Public Function getRqms() As Double
    getRqms = pcDemand + incBc
End Function

Public Sub parseStringToDate(strD)
    
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
    
    s = Format(y & "-" & m & "-" & d, "yyyy-mm-dd")
    
    Me.setDate CStr(s)
End Sub

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

Private Sub Class_Initialize()
    Set collectionOfCorailItemsFrom2510 = New Collection
End Sub

Private Sub Class_Terminate()
    Set collectionOfCorailItemsFrom2510 = Nothing
End Sub
