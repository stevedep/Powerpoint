Sub context()
   
    
    Dim sld As Slide
    Dim shp As Shape
    
    Set sld = Application.ActiveWindow.View.Slide
    
    sld.Duplicate
    
      
    y = ActiveWindow.Selection.ShapeRange.Top
    x = ActiveWindow.Selection.ShapeRange.Left
    w = ActiveWindow.Selection.ShapeRange.Width
    h = ActiveWindow.Selection.ShapeRange.Height
    
    
    Windows(1).Selection.Copy
    ActivePresentation.Slides(sld.SlideIndex).Shapes.Paste
    
    ActivePresentation.Slides(sld.SlideIndex).Shapes(ActivePresentation.Slides(sld.SlideIndex).Shapes.Count).Top = y
    ActivePresentation.Slides(sld.SlideIndex).Shapes(ActivePresentation.Slides(sld.SlideIndex).Shapes.Count).Left = x
    n = ActivePresentation.Slides(sld.SlideIndex).Shapes.Count
    
sld.Shapes.AddShape Type:=msoShapeRectangle, _
    Left:=0, Top:=0, Width:=960, Height:=540
    
    ActivePresentation.Slides(sld.SlideIndex).Shapes(ActivePresentation.Slides(sld.SlideIndex).Shapes.Count).Fill.ForeColor.RGB = RGB(172, 185, 202)
    ActivePresentation.Slides(sld.SlideIndex).Shapes(ActivePresentation.Slides(sld.SlideIndex).Shapes.Count).Fill.Transparency = 0.2

    sld.Shapes.Range(Array(ActivePresentation.Slides(sld.SlideIndex).Shapes.Count, n)).Select
    
    ActiveWindow.Selection.ShapeRange.MergeShapes msoMergeCombine
   

Set shp = Nothing

Dim ffb As FreeformBuilder
Dim myshape As Shape
Dim currentslide As Slide

Set currentslide = sld 'ActivePresentation.Slides(ActiveWindow.View.Slide.SlideIndex)

Links = MsgBox("Left?", vbYesNo + vbQuestion, "Side")

If Links = vbYes Then
       Set ffb = currentslide.Shapes.BuildFreeform(msoEditingCorner, x, y)
        With ffb
            .AddNodes msoSegmentLine, msoEditingAuto, x - 40, 135
            .AddNodes msoSegmentLine, msoEditingAuto, x - 40, 30
            .AddNodes msoSegmentLine, msoEditingAuto, x - 300, 30
            .AddNodes msoSegmentLine, msoEditingAuto, x - 300, 450
            .AddNodes msoSegmentLine, msoEditingAuto, x - 40, 450
            .AddNodes msoSegmentLine, msoEditingAuto, x - 40, 350
            .AddNodes msoSegmentLine, msoEditingAuto, x, y + h
        End With
Else
         Set ffb = currentslide.Shapes.BuildFreeform(msoEditingCorner, x + w, y)
        With ffb
            .AddNodes msoSegmentLine, msoEditingAuto, x + w + 40, 135
            .AddNodes msoSegmentLine, msoEditingAuto, x + w + 40, 30
            .AddNodes msoSegmentLine, msoEditingAuto, x + w + 40 + 300, 30
            .AddNodes msoSegmentLine, msoEditingAuto, x + w + 40 + 300, 450
            .AddNodes msoSegmentLine, msoEditingAuto, x + w + 40, 450
            .AddNodes msoSegmentLine, msoEditingAuto, x + w + 40, 350
            .AddNodes msoSegmentLine, msoEditingAuto, x + w, y + h
        End With
End If

Set myshape = ffb.ConvertToShape
myshape.Fill.ForeColor.RGB = RGB(256, 256, 256)


Set sld = Nothing
Set myshape = Nothing

End Sub


