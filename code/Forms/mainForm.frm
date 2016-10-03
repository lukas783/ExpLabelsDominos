VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mainForm 
   Caption         =   "Dominos Expiration Labels"
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

Private Sub OptionButton2_Click()

End Sub

Private Sub OptionButton1_Click()

End Sub

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
    Print #1, "DRAW_BOX 425 0 1 500 2"
    
    Print #1, "TEXT 3 30 20 ItemLine1"
    Print #1, "TEXT 3 450 30 stuffs"
    Print #1, "TEXT 3 30 65 ItemLine2"
    Print #1, "TEXT 2 30 100 OP / Rec/ In-Use"
    Print #1, "TEXT 2 30 140 DATE!"
    Print #1, "TEXT 3 30 200 EXPIRES"
    Print #1, "TEXT 2 30 240 date.."
     
    Print #1, "END"
    
    Close #1
    Shell ("notepad.exe /PT " & Chr(34) & labelPath & Chr(34) & " " & Chr(34) & "Label" & Chr(34))
    'END PRINTING STUFF?
End Sub

Private Sub TextBox1_Change()
CommandButton1.Caption = TextBox1.Text
End Sub


