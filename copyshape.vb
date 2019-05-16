Sub copyshape()

Dim sld As Slide
Dim shp As Shape
Set sld = Application.ActiveWindow.View.Slide

s = sld.SlideIndex
Windows(1).Selection.Copy
aantal = InputBox("Aantal slides")


For i = s + 1 To s + aantal
    ActivePresentation.Slides(i).Shapes.Paste
Next i


End Sub
