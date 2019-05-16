Sub move()
Dim shp As Object

xmove = InputBox("x adjustment")

If ActiveWindow.Selection.Type = ppSelectionNone Then
   MsgBox "Please select objects", vbExclamation, "Make Selection"
Else
   For Each shp In ActiveWindow.Selection.ShapeRange
       'MsgBox shp.Name
       shp.Left = shp.Left + xmove
   Next shp
End If

Set shp = Nothing
End Sub
