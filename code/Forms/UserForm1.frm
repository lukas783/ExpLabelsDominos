VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PrintButton_Click()
    Dim path As String
    path = ThisWorkbook.path
    If (InStr(path, "/") > 0) Then
        path = path & "/"
    Else
        path = path & "\"
    End If
    Dim labelPath As String
    labelPath = path & "label.txt"
    Open labelPath For Output As #1
    
    ' START PRINTING STUFF?
    Print #1, "! 0 100 350 1"
    Print #1, "DRAW_BOX 0 115 850 1 2"
    Print #1, "DRAW_BOX 0 240 850 1 2"
    
    Print #1, "30 0 4 TEST1"
    Print #1, "30 40 4 TEST2"
    
    Print #1, "END"
    Close #1
    Shell ("notepad.exe /PT " & Chr(34) & labelPath & Chr(34) & " " & Chr(34) & "Label" & Chr(34))
    'END PRINTING STUFF?
End Sub
