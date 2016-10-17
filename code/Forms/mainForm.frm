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
Private Sub aclear_Button_Click()
ListBox1.Clear
j = ItemCol.Count - 1
Set ItemCol = Nothing

End Sub

Private Sub aPrint_Button_Click()
    Dim lab2col As New Collection
    Dim lab4col As New Collection
    Dim name1 As String
    Dim name_2 As String
    Dim expdate As String
    Dim colCount As Integer
    colCount = ItemCol.Count
    
    For i = 1 To colCount
        If ItemCol(i).labelType = "2" Then
            lab2col.Add (ItemCol(i))
        Else
            lab4col.Add (ItemCol(i))
        End If
    Next i
    
    For i = 1 To lab2col.Count
        MsgBox ("2: " & lab2col(i).name)
    Next i
    For i = 1 To lab4col.Count
        MsgBox ("4: " & lab4col(i).name)
    Next i
    ListBox1.Clear
    Set ItemCol = Nothing
    
    

End Sub

Private Sub dateLabel_Click()

End Sub

' Option to set date
Private Sub dateText_Click()

End Sub
'start buttons for choosing label styles

Private Sub Label1_Click()

End Sub

Private Sub label2_Alex_Click()

End Sub

Private Sub label4_Alex_Click()

End Sub

Private Sub Labels2_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

' end lebel style choices
' start choices for opened, received, prepped, and in-use
Private Sub openButton_Click()

End Sub

Private Sub In_Use_Alex_Click()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub Prepped_Alex_Click()

End Sub

Private Sub preppedButton_Click()

End Sub

Private Sub receivedButton_Click()

End Sub

Private Sub useButton_Click()

End Sub
' end choices for opened, received, prepped, and in-use
' start textbox for prepper's name
Private Sub nameText_Change()

End Sub
' end for textbox for prepper's name
' start set date
Private Sub MultiPage1_Change()
    dateText.Caption = getDate
End Sub
'end set date
'start topping clicks

Private Sub Banana_Pep_Button_Click()
Dim prep As New pItem

prep.name = "Banana"
prep.name2 = "Peppers"
prep.expdate = getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
If mainForm.In_Use_Alex.Value = "True" Then
    prep.useType = "In-Use"
Else
    prep.useType = "Prepped"
End If
If mainForm.label2_Alex.Value = "True" Then
    prep.labelType = "2"
Else
    prep.labelType = "4"
End If

ItemCol.Add prep
ListBox1.AddItem ItemCol(ItemCol.Count).name
ListBox1.List((ItemCol.Count - 1), 1) = ItemCol(ItemCol.Count).name2


'.List(ItemCol.Count - 1, 0) = ItemCol(ItemCol.Count).name
'.List(ItemCol.Count - 1, 1) = ItemCol(ItemCol.Count).name2
End Sub

Private Sub Boned_Wings_Button_Click()
printLabel "Bone-In", "Wings", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub Boneless_Wings_Click()
printLabel "Boneless", "Wings", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub Brownies_Button_Click()
printLabel "Brownies", "", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub
Private Sub Pasta_Button_Click()
printLabel "Pasta", "", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub GOlives_Button_Click()
printLabel "Green", "Olives", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub Jalapeno_Button_Click()
printLabel "Jalapeno", "Peppers", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub Lava_Cakes_Button_Click()
printLabel "Lava", "Cakes", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub RedPeppers_Button_Click()
printLabel "Red", "Peppers", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub Sand_Bread_Button_Click()
printLabel "Sandwich", "Bread", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub Specialty_Chick_Button_Click()
printLabel "Specialty", "Chicken", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub Tomatoes_Button_Click()
printLabel "Tomatoes", "", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub
Private Sub pep_Button_Click()
printLabel "Pepperoni", "", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
MsgBox (TextBox1.Text)

End Sub

Private Sub Philly_Button_Click()
printLabel "Philly", "", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)

End Sub

Private Sub Pineapple_Button_Click()
printLabel "Pineapple", "", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub
Private Sub Feta_Button_Click()
printLabel "Feta", "", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub GPep_Button_Click()
printLabel "Green", "Peppers", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub Ham_Button_Click()
printLabel "Ham", "", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub
Private Sub Mushroom_Button_Click()
Dim inputDate As String

inputDate = InputBox("Enter Expiration Date")
If (inputDate = "") Then
Else
    printLabel "Mushroom", "", inputDate
End If

End Sub

Private Sub Onion_Button_Click()
printLabel "Onlon", "", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub
Private Sub Marinara_Click()
    printLabel "Marinara", "Sauce", "FFSDate"
End Sub
Private Sub American_Button_Click()
printLabel "American", "Cheese", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub Asiago_Button_Click()
printLabel "Asiago", "", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub Bacon_Button_Click()
printLabel "Bacon", "", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub aBeef_Button_Click()
printLabel "Beef", "", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub BOlives_Button_Click()
printLabel "Black", "Olives", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub Cheddar_Button_Click()
printLabel "Cheddar", "", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub CheeseBlend_Button_Click()
printLabel "50/50", "CheeseBlend", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub Chicken_Button_Click()
printLabel "Grilled", "Chicken", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub Sausage_Button_Click()
printLabel "Sausage", "", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub Shredded_Prov_Button_Click()
printLabel "Shredded", "Provolone", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub Sliced_Prov_Button_Click()
printLabel "Sliced", "Prov", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub

Private Sub Spinach_Button_Click()
printLabel "Spinach", "", getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
End Sub
'end toppings

Private Sub PrintButton_Click()
    Dim expdate As String
    expdate = getExpDate(TextBox1.Text, OptionButton1.Value, OptionButton2.Value)
    MsgBox ("... " & getExpDate(TextBox1, OptionButton1, OptionButton2))
    
    printLabel "Green", "Peppers", expdate
        
End Sub


Private Sub sauceLabel_Click()

End Sub

Private Sub TextBox1_Change()
    CommandButton1.Caption = TextBox1.Text
End Sub

