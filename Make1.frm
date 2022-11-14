VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Make1 
   Caption         =   "Make Input List"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7995
   OleObjectBlob   =   "Make1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Make1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Public login As String
Public pass As String



Public plt As String

Private Sub BtnAdd_Click()

    ' page=1&max=10&charset=UTF-8&ecwAutoId=false&ecwDivId=productSearch&shortage=false
    '   &provisionerAreaIdArray=46732&provisionerAreaLabel=BLANKENBURG HEIKO - 67SYN
    '   &sgrLineDetail=false&errorState=false&displayLocal=false
    
    
    Dim params(0 To 2) As String
    
    
    Dim x As Variant
    For x = 1 To Me.ListBox2.ListCount
        
        If Me.ListBox2.Selected(x) Then
            
            ' =======================================================================================================
            ' OK OK OK
            Debug.Print Me.ListBox2.List(x, 0) & " " & Me.ListBox2.List(x, 1) & " " & Me.ListBox2.List(x, 2)
            params(0) = Me.ListBox2.List(x, 0)
            params(1) = Me.ListBox2.List(x, 1)
            params(2) = Me.ListBox2.List(x, 2)
            
            Exit For
            ' =======================================================================================================
        End If
    Next x
    
    
    quikGetPNsList plt, params(0), params(1), params(2), Me, FFOC.E_PRE_LIST.E_PRE_LIST_ADD
    
    ' ---------------------------------------------
    Me.pb2.Width = Me.pb1.Width
    ' ---------------------------------------------
End Sub

Private Sub BtnMove_Click()
    
    ' move viisble data from pre-list
    
    Dim sh1 As Worksheet, sh2 As Worksheet
    
    Set sh1 = ThisWorkbook.Sheets(FFOC.G_SH_NM_PRE_LIST)
    Set sh2 = ThisWorkbook.Sheets(FFOC.G_SH_NM_IN)
    
    
    Dim d As Dictionary
    Set d = Nothing
    Set d = New Dictionary
    
    
    Dim r As Range, tmpstr As Variant
    Set r = sh1.Range("C2")
    
    Do
        If Not r.EntireRow.Hidden Then
        
            tmpstr = CStr(r.Value) & "__" & CStr(r.Offset(0, FFOC.E_2720_IN_CMNT).Value)
        
            If d.Exists(tmpstr) Then
                
            Else
                d.Add tmpstr, Array(r.Value, r.Offset(0, FFOC.E_2720_IN_CMNT).Value)
            End If
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r.Value) = ""
    
    
    innerClearInputList True
    
    
    Dim ir As Range
    sh2.Activate
    Set ir = sh2.Range("A2")
    
    For Each tmpstr In d.Keys
        ' PLT
        ir.Value = "" & d(tmpstr)(1)
        ir.Offset(0, 1).Value = d(tmpstr)(0)
        ir.Offset(0, 2).Value = "BLUE"
        ir.Offset(0, 3).Value = "" ' r.Offset(0, FFOC.E_2720_IN_CMNT - 1).Value
        
        Set ir = ir.Offset(1, 0)
    Next
    
    
    
    
    sh2.Activate
    
    
End Sub

Private Sub BtnReplace_Click()
    ' page=1&max=10&charset=UTF-8&ecwAutoId=false&ecwDivId=productSearch&shortage=false
    '   &provisionerAreaIdArray=46732&provisionerAreaLabel=BLANKENBURG HEIKO - 67SYN
    '   &sgrLineDetail=false&errorState=false&displayLocal=false
    
    
    Dim params(0 To 2) As String
    
    
    Dim x As Variant
    For x = 1 To Me.ListBox2.ListCount
        
        If Me.ListBox2.Selected(x) Then
            
            ' =======================================================================================================
            ' OK OK OK
            Debug.Print Me.ListBox2.List(x, 0) & " " & Me.ListBox2.List(x, 1) & " " & Me.ListBox2.List(x, 2)
            params(0) = Me.ListBox2.List(x, 0)
            params(1) = Me.ListBox2.List(x, 1)
            params(2) = Me.ListBox2.List(x, 2)
            
            Exit For
            ' =======================================================================================================
        End If
    Next x
    
    
    quikGetPNsList plt, params(0), params(1), params(2), Me, FFOC.E_PRE_LIST.E_PRE_LIST_NEW
    
    ' ---------------------------------------------
    Me.pb2.Width = Me.pb1.Width
    ' ---------------------------------------------
End Sub

Private Sub CommandButton1_Click()
    hide
End Sub

Private Sub CommandButton2_Click()
    End
End Sub

Private Sub ListBox1_Click()



    

    Me.ListBox2.Clear
    
    Debug.Print Me.ListBox1.Value
    
    Me.pb2.Width = 1
    
    plt = CStr(Me.ListBox1.Value)
    quikTestOn CStr(Me.ListBox1.Value), Me
    
    Me.pb2.Width = Me.pb1.Width
End Sub

