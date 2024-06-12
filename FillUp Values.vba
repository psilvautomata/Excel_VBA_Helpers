Sub fillup_values()

Dim i As Integer
Dim max As Integer
Dim column As String

Application.ScreenUpdating = False 'Freeze the screen

max = Range("YourTable").Rows.Count 'Count the rows from a table
column = Range("YourCell").Value 'A Cell to insert the Column letter to fill up the values - Better updating a cell value then the code everytime

For i = max To 2 Step -1 'Iteration from bottom to the top (Step -1 means reverse iteration)
    If Range(column & i).Value = "" Then 'If value is "" (Empty)
    Range(column & i).Select 'Select the empty cell
    Selection.FillUp 'Fill up based on the cell below
    End If 'End the Loop

Next

Application.ScreenUpdating = True 'Unfreeze the screen


Range("YourCell").Select 'Select a specific cell of the sheet


End Sub
