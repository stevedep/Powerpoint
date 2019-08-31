    
    Dim sld As Slide
    Dim shp As Shape
    
    Set sld = Application.ActiveWindow.View.Slide
    
    
    
    y = ActiveWindow.Selection.ShapeRange.Top
    x = ActiveWindow.Selection.ShapeRange.Left
    
    
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
   

Set sld = Nothing
Set shp = Nothing
