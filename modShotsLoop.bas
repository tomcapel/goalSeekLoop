Attribute VB_Name = "modShotsLoop"
Sub GoalSeekLoop()

Application.ScreenUpdating = False

Dim intRow As Integer
intRow = 3 '3rd row down is where my data starts

Do Until IsEmpty(Cells(intRow, 2)) 'Perform until the defined row and 2nd column is empty

Cells(intRow, 6).GoalSeek Goal:=Cells(intRow, 3), ChangingCell:=Cells(intRow, 5) 'select defined row, column 8 and perform goal seek by changing the value of defined row, column 7
intRow = intRow + 1 'add one for next line

Loop

MsgBox "SoT Totals Updated"

End Sub

