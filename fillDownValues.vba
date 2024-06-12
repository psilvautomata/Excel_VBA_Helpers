Sub filldown_values()

Dim i As Integer
Dim max As Integer
Dim column As String

Application.ScreenUpdating = False 'Freeze the screen

max = Range("YourTable").Rows.Count + 1 'Count YourTable rows
column = Range("YourCell").Value 'A Cell to insert the Column letter to fill down the values - Better updating a cell value then the code everytime

For i = 2 To max 'Rows Iteration (starts at 2, jumps the header)
    If Range(column & i).Value = "" Then 'If cell value equal to "" (empty)
    Range(column & i).Select 'Select the empty cell
    Selection.FillDown 'Fill down the empty row (Repeats the above cell value)

    End If 'End the loop

Next

Application.ScreenUpdating = True 'Unfreeze the screen


Range("YourCell").Select 'Select a cell on the sheet


End Sub
