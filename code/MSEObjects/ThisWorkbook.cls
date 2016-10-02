VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Dim response As VbMsgBoxResult
    response = MsgBox("This is a test program, don't run this if you don't know what it is!\n", vbDefaultButton1, "TEST MSG BOX")
    
    UserForm1.Show                                  ' Inits the main form for the application
    'Sheet1.Application.WindowState = xlMinimized    ' Minimizes excel workbook
End Sub
