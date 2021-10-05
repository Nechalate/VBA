Attribute VB_Name = "Module4"
Sub ListLoop()
Dim count As Integer
For count = 0 To 10
Selection.Value = count
Selection.Offset(0, 1).Select
Next count
End Sub
