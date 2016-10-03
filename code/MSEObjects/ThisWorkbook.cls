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

Private Sub Workbook_Open()
    Dim response As VbMsgBoxResult
    response = MsgBox("This is a test program, don't run this if you don't know what it is!\n", vbDefaultButton1, "TEST MSG BOX")
    
    mainForm.width = GetSystemMetrics(0) * 0.75             ' System screen width * (3/4)
    mainForm.Height = GetSystemMetrics(1) * 0.75            ' System screen height * (3/4)
    mainForm.MultiPage1.width = mainForm.width              ' tabbed pane, set width and height to appropriate dimensions
    mainForm.MultiPage1.Height = mainForm.Height - mainForm.MultiPage1.Top
    
    mainForm.Show                                           ' Inits the main form for the application
    
End Sub
