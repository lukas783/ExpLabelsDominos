VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mainForm 
   Caption         =   "Dominos Experiation Labels"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15120
   OleObjectBlob   =   "mainForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "mainForm"
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
    
    Print #1, "TEXT 30 0 4 TEST1"
    Print #1, "TEXT 30 40 4 TEST2"
    
    Print #1, "END"
    Close #1
    Shell ("notepad.exe /PT " & Chr(34) & labelPath & Chr(34) & " " & Chr(34) & "Label" & Chr(34))
    'END PRINTING STUFF?
End Sub

Private Sub TextBox1_Change()
CommandButton1.Caption = TextBox1.Text
End Sub
