
Private Sub Worksheet_Change(ByVal Target As Range)

Dim row As Long
Dim column As Long

row = Target.Row
column = Target.Column 'Get the row and column of the changed cell

If column = xX And row = yY Then 'Check if the change occurred in a specific cell (Replace xX and yY with your row and column numbers)
    Application.EnableEvents = False 'Useful when you are making changes to multiple cells or sheets and don't want events related to those changes to trigger other macros or automated actions.

    If Range("Z").Value = "Value1" Then
    	  Rows("z:z").EntireRow.Hidden = False 'Unhides rows, because if they are hidden and not unhidden, new rows will be hidden incorrectly
    	  Rows("z:z").EntireRow.Hidden = True 'Hide rows
    ElseIf Range("W").Value = "Value2" Then
        Rows("z:z").EntireRow.Hidden = False 'Unhide rows, it is so important to unhide, cause if they are hidden your rows will change and it will be hidden two times
        Rows("w:w").EntireRow.Hidden = True 'Hide anoter section
    ElseIf Range("K").Value = "Value3" Then
        Rows("z:z").EntireRow.Hidden = False 'Unhide rows in the specified range
        Rows("k:k").EntireRow.Hidden = True 'Hide another section of rows
    End If

    Application.EnableEvents = True 'Re-enable Excel events

End If
    
Range("YourCell").Select 'Select a specific cell after the changes are made (Replace "YourCell")

End Sub
