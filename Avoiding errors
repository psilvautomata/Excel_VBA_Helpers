Sub avoidingErrors ()

On Error Resume Next 'When using a loop (for) -> resume to next
. . . Your code . . .
Next
End Sub
'
'
On Error GoTo 0 'Forces Excel to threat erros in the standard way
'
'
On Error GoTo ExampleVariable
. . . Your code . . . 
Exit Sub
ExampleVariable
End Sub
'
'
If range("xX").value = " " then 'Change xX for your cell index
Exit Sub 'If cell value is empty exit sub
End If
End Sub
'
'
