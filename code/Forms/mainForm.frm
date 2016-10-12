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

Private Sub CommandButton1_Click()
    printLabel "Marinara", "Sauce"
End Sub

Private Sub MultiPage1_Change()
    dateText.Caption = getDate
End Sub

Private Sub PrintButton_Click()
        printLabel "Green", "Peppers"
        
End Sub

Private Sub TextBox1_Change()
CommandButton1.Caption = TextBox1.Text
End Sub


