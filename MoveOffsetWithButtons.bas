Attribute VB_Name = "Module3"
Sub MoveUp()
Selection.Offset(-1, 0).Select
End Sub
Sub MoveDown()
Selection.Offset(1, 0).Select
End Sub
Sub MoveRight()
Selection.Offset(0, 1).Select
End Sub
Sub MoveLeft()
Selection.Offset(0, -1).Select
End Sub
