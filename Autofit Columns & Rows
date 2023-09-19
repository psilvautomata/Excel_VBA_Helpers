Sub autofit()
'
    Rows.Range("(1:1):(500:500)").EntireRow.AutoFit 'Insert your Rows range (X:X):(Y:Y)
    Columns("A:E").EntireColumn.AutoFit 'Columns 'Insert your Columns (X:Y)

End Sub


or

Sub autofit()

Sheets("YourSheet").Activate 'Insert your sheet name that you want to set active
ActiveSheet.UsedRange.EntireRow.AutoFit 'UsedRange means an Area in your sheet that cointains some data
ActiveSheet.UsedRange.EntireColumn.AutoFit
    
End Sub
