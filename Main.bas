Attribute VB_Name = "Main"
'The MIT License (MIT)
'
'Copyright (c) 2018 FORREST
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

Public Sub runMain(ictrl As IRibbonControl)
    
    login.CheckBoxHazards.Value = False
    login.show
End Sub


Public Sub innerFromLogin()


    ThisWorkbook.Sheets(FFOC.G_SH_NM_REG).Range("RUN").Value = 1
    Application.enableEvents = False
    
    ' --------------------------------------------------------------
    Main
    ' --------------------------------------------------------------
    
    ThisWorkbook.Sheets(FFOC.G_SH_NM_REG).Range("RUN").Value = 0
    Application.enableEvents = True
    
    MsgBox "Gotowe!"
End Sub

Public Sub innerFromLoginBO()


    
    Application.enableEvents = False
    ThisWorkbook.Sheets(FFOC.G_SH_NM_REG).Range("RUN").Value = 1
    
    ' --------------------------------------------------------------
    MainBO
    ' --------------------------------------------------------------
    
    ThisWorkbook.Sheets(FFOC.G_SH_NM_REG).Range("RUN").Value = 0
    Application.enableEvents = True
    
    MsgBox "Gotowe!"
End Sub

Public Sub Main()
    
    ' ------------------------------------------------------------------------
    
    Dim c As CorailHelper, i As InputListHelper
    Set c = New CorailHelper
    
    
    ' this object is an component for main object CorailHelper
    Set i = New InputListHelper
    c.run i
    
    c.putDataOnReportSheet
    c.makeLayout
    
    ' ! as global one : ' public sub from global module
    FFOC.RunOnSelectionChangeModule.recalcLayoutAndColors ActiveSheet, Range("B4")
    FFOC.FirstRunoutModule.firstRunoutFormulaFilling ActiveSheet, Range("B4")
    
    ' ------------------------------------------------------------------------
End Sub

Public Sub MainBO()
    
    ' ------------------------------------------------------------------------
    
    Dim c As CorailHelper, i As InputListHelper
    Set c = New CorailHelper
    
    
    ' this object is an component for main object CorailHelper
    Set i = New InputListHelper
    c.runBO i
    
    c.putDataOnReportSheetBO
    c.makeLayoutBO
    
    ' ! as global one : ' public sub from global module
    FFOC.RunOnSelectionChangeModule.recalcLayoutAndColors ActiveSheet, Range("B4")
    FFOC.FirstRunoutModule.firstRunoutFormulaFilling ActiveSheet, Range("B4")
    
    ' ------------------------------------------------------------------------
End Sub


Public Sub aQuickRun()

    G_LOGIN = ""
    G_PASS = ""
    G_HAZARDS = False
    
    innerFromLogin
End Sub
