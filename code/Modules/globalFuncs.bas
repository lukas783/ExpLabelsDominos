Attribute VB_Name = "globalFuncs"

Public ItemCol As New Collection


Public Function getExpDate(txtBox As String, optHrs As Boolean, optDays As Boolean) As String
    Dim currDate As Date
    Dim txtInt As Integer
    currDate = Now
    
    If IsNumeric(txtBox) = "True" Then
        txtInt = CInt(txtBox)
    Else
        txtInt = 0
    End If
    
    If optHrs = "True" Then
        currDate = DateAdd("h", txtInt, currDate)
        getExpDate = Format(currDate, "hh:nn ampm")
        ' Do the hrs forward by # in txtBox
    ElseIf optDays = "True" Then
        ' Do the days forward by # in txtBox
        currDate = DateAdd("d", txtInt, currDate)
        getExpDate = Format(currDate, "mm/dd/yyyy")
    Else
        getExpDate = ""
    End If
    
End Function

Public Function getPath() As String
    Dim path As String
    path = ThisWorkbook.path
    If (InStr(path, "/") > 0) Then
        getPath = path & "/"
    Else
        getPath = path & "\"
    End If
End Function

Public Function getDate() As String
    getDate = Format(Date, "mm/dd/yyyy")
End Function

Public Function getUsage() As String
    If mainForm.MultiPage1.Value = 0 Then
    
        If mainForm.receivedButton.Value = "True" Then
            getUsage = "Received"
        ElseIf mainForm.openButton.Value = "True" Then
            getUsage = "Opened"
        ElseIf mainForm.preppedButton.Value = "True" Then
            getUsage = "Prepped"
        ElseIf mainForm.useButton.Value = "True" Then
            getUsage = "In-Use"
        Else
            getUsage = "N/A"
        End If
    End If
    If mainForm.MultiPage1.Value = 3 Then
        If mainForm.In_Use_Alex.Value = "True" Then
            getUsage = "In-Use"
        ElseIf mainForm.Prepped_Alex.Value = "True" Then
            getUsage = "Prepped"
        Else
            getUsage = "N/A"
        End If
    End If
    
End Function
Sub printLabel(item1 As String, item2 As String, expdate As String)

    Dim labelPath As String
    labelPath = getPath & "label.txt"
    Open labelPath For Output As #1
    
    
    
    If mainForm.Labels2.Value = "True" Then
        ' START PRINTING STUFF FOR 2 Labels Option
        Print #1, "! 0 100 350 1"
        Print #1, "DRAW_BOX 425 0 1 500 2"
        Print #1, "TEXT 3 40 20 " & item1
        Print #1, "TEXT 3 40 75 " & item2
        Print #1, "TEXT 2 40 125 " & getUsage
        Print #1, "TEXT 2 40 150 Prepped on: " & getDate
        Print #1, "TEXT 2 40 185 By: " & mainForm.nameText.Text
        Print #1, "TEXT 3 40 220 EXPIRES"
        Print #1, "TEXT 4 40 260 " & expdate
    
        Print #1, "TEXT 3 450 20 " & item1
        Print #1, "TEXT 3 450 75 " & item2
        Print #1, "TEXT 2 450 125 " & getUsage
        Print #1, "TEXT 2 450 150 Prepped on: " & getDate
        Print #1, "TEXT 2 450 185 By: " & mainForm.nameText.Text
        Print #1, "TEXT 3 450 220 EXPIRES"
        Print #1, "TEXT 4 450 260 " & expdate
     
    Else
        ' START PRINTING STUFF FOR 4 LABELS OPTION
        Print #1, "! 0 100 350 1"
        
        
    End If
    
    Print #1, "END"
    
    Close #1
    Shell ("notepad.exe /PT " & Chr(34) & labelPath & Chr(34) & " " & Chr(34) & "Label" & Chr(34))
    'END PRINTING STUFF?
End Sub

