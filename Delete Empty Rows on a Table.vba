Sub emptyrow()

Dim i As Integer
Dim max As Integer

Application.ScreenUpdating = False 'Freezes the screen updating

max = Range("YourTableName").Rows.Count + 1 'Count the rows from YourTableName

Range("A1").Select 'Selects the beginning of the active sheet on dataflow

For i = 1 To max 'Iteration o variable i
    If Cells(i, 1) = "" And Cells(i, 2) = "" And Cells(i, 3) = "" And ... Cells(i, n) = "" Then  'Checks if the Cells value are "" (empty)
    Rows(i).Select 'Selects the entire empty row
    Selection.Delete 'Delete the entire empty row

    End If 'Stops the loop

Next

Application.ScreenUpdating = True 'Activate the screen

End Sub
