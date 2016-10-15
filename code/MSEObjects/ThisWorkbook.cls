VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long   ' Function to pull variables from system

 
 
Sub File_Download_From_Website()
    'Declare Local Variables to be Used in this Sub Module.
    Dim InpUrl As String
    Dim OutFilePath As String
    Dim DownloadStatus As Long
    Dim path As String
    path = ThisWorkbook.path
    If (InStr(path, "/") > 0) Then
        path = path & "/"
    Else
        path = path & "\"
    End If
    'Read Input Path for the File and Output File Destination Path
    InpUrl = "https://docs.google.com/document/d/1Cc5LsYeSXc7BAHZoy0zJIYLWXfZtP5Z0WEvvnJsUnU0/export?format=docx" 'ThisWorkbook.Sheets(1).Cells(1, 1)"
    OutFilePath = path & "ExpLabelsLucas.xlsm" 'ThisWorkbook.Sheets(1).Cells(2, 1)
    
    'Invoke API to download file from the website.
    DownloadStatus = URLDownloadToFile(0, InpUrl, OutFilePath, 0, 0)

    'Check return status from the API and Intimate the User accordingly.
    If DownloadStatus = 0 Then
        Application.Speech.Speak "No. Don't touch that... wait.. noooo...ooo...ooo...ooo...ERROR. Deleting win-thirty-two dot DLL" 'File Downloaded. Check in this path: " & OutFilePath, True
        MsgBox "File Downloaded. Check in this path: " & OutFilePath
    Else
        Application.Speech.Speak "Download File Process Failed"
        MsgBox "Download File Process Failed"
    End If
End Sub


Private Sub Workbook_Open()
    
    Dim response As VbMsgBoxResult
    ' response = MsgBox("This is a test program, don't run this if you don't know what it is!\n", vbDefaultButton1, "TEST MSG BOX")
    
    mainForm.Width = GetSystemMetrics(0) * 0.75             ' System screen width * (3/4)
    mainForm.Height = GetSystemMetrics(1) * 0.75            ' System screen height * (3/4)
    mainForm.MultiPage1.Width = mainForm.Width              ' tabbed pane, set width and height to appropriate dimensions
    mainForm.MultiPage1.Height = mainForm.Height - mainForm.MultiPage1.Top
    mainForm.MultiPage2.Width = mainForm.Width
    mainForm.MultiPage2.Height = mainForm.MultiPage1.Height - mainForm.MultiPage2.Top
    
    mainForm.dateText.Caption = getDate
    
    mainForm.Show                                           ' Inits the main form for the application

End Sub
