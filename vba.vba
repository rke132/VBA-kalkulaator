Option Explicit
Dim firstNumber As Long
Dim secondNumber As Long
Dim operator As String


Private Sub AdditionButton_Click()
firstNumber = DisplayBox.Text
operator = "+"
DisplayBox.Text = ""

End Sub

Private Sub ClearButton_Click()
DisplayBox.Text = "0"

End Sub

Private Sub CommandButton0_Click()
DisplayBox.Text = ""
If (DisplayBox.Text = "0") Then
DisplayBox.Text = "0"
Else
DisplayBox.Text = DisplayBox.Text + "0"

End If
End Sub

Private Sub CommandButton1_Click()
DisplayBox.Text = ""
If (DisplayBox.Text = "0") Then
DisplayBox.Text = "1"
Else
DisplayBox.Text = DisplayBox.Text + "1"

End If

End Sub

Private Sub CommandButton2_Click()
If (DisplayBox.Text = "0") Then
DisplayBox.Text = "2"
Else
DisplayBox.Text = DisplayBox.Text + "2"

End If


End Sub

Private Sub CommandButton3_Click()
If (DisplayBox.Text = "0") Then
DisplayBox.Text = "3"
Else
DisplayBox.Text = DisplayBox.Text + "3"

End If
End Sub


Private Sub CommandButton4_Click()
If (DisplayBox.Text = "0") Then
DisplayBox.Text = "4"
Else
DisplayBox.Text = DisplayBox.Text + "4"

End If
End Sub

Private Sub CommandButton5_Click()
If (DisplayBox.Text = "0") Then
DisplayBox.Text = "5"
Else
DisplayBox.Text = DisplayBox.Text + "5"

End If
End Sub


Private Sub CommandButton6_Click()
If (DisplayBox.Text = "0") Then
DisplayBox.Text = "6"
Else
DisplayBox.Text = DisplayBox.Text + "6"

End If
End Sub

Private Sub CommandButton7_Click()
If (DisplayBox.Text = "0") Then
DisplayBox.Text = "7"
Else
DisplayBox.Text = DisplayBox.Text + "7"

End If
End Sub


Private Sub CommandButton8_Click()
If (DisplayBox.Text = "0") Then
DisplayBox.Text = "8"
Else
DisplayBox.Text = DisplayBox.Text + "8"

End If
End Sub


Private Sub CommandButton9_Click()
If (DisplayBox.Text = "0") Then
DisplayBox.Text = "9"
Else
DisplayBox.Text = DisplayBox.Text + "9"

End If
End Sub

Private Sub DivisionButton_Click()
firstNumber = DisplayBox.Text
operator = "/"
DisplayBox.Text = ""

End Sub

Private Sub MultiplicationButton_Click()

firstNumber = DisplayBox.Text
operator = "*"
DisplayBox.Text = ""

End Sub


Private Sub PlusMinusButton_Click()

Dim x
x = DisplayBox.Text
DisplayBox.Text = (-1 * x)

End Sub

Private Sub ResultButton_Click()
secondNumber = DisplayBox.Text

Select Case operator
Case "+"
    DisplayBox.Text = firstNumber + secondNumber
Case "-"
    DisplayBox.Text = firstNumber - secondNumber
Case "*"
     DisplayBox.Text = firstNumber * secondNumber
Case "/"
If secondNumber = 0 Then
    DisplayBox.Text = "Cannot divide by zero"
Else
     DisplayBox.Text = firstNumber / secondNumber

End If

End Select


End Sub

Private Sub SubtractionButton_Click()
firstNumber = DisplayBox.Text
operator = "-"
DisplayBox.Text = ""

End Sub


Private Sub UserForm_Click()

End Sub
