Sub CreateRectanglesWithCircles() 'houden

    RemoveShapeByName 'cleanup before start

    ' Declare variables
    Dim oSlide As slide
    Dim oShape As shape
    Dim oGroup As shape
    Dim oCircle As shape
    Dim oText As shape
    Dim i As Integer
    Dim NumRectangles As Integer
    Dim RectHeight As Single
    Dim RectTop As Single
    Dim CircleDiameter As Single
    Dim CircleLeft As Single
    Dim CircleTop As Single
    Dim TextLeft As Single
    Dim TextTop As Single
    Dim RectSpacing As Single ' Variable for spacing between rectangles
    
    ' Set number of rectangles to create and spacing between rectangles
    RectSpacing = 10 ' Set the vertical spacing between rectangles
    
    ' Get active slide
    Set oSlide = Application.ActiveWindow.View.slide
        
    slidesArray = FindMenuShapes()
    NumRectangles = UBound(slidesArray, 1)
    ' Calculate height of each rectangle
    RectHeight = (Application.ActivePresentation.PageSetup.SlideHeight - RectSpacing * (NumRectangles - 1)) / NumRectangles ' Modify calculation to include spacing
    
    ' Calculate diameter of circle to fit within each rectangle
    CircleDiameter = RectHeight * 0.8
  
    
    Dim j As Integer
    For i = 1 To UBound(slidesArray, 1)
        j = i * 2 - 1
        
        
               ' Calculate top position of current rectangle
        RectTop = (i - 1) * (RectHeight + RectSpacing) ' Modify calculation to include spacing
        
        ' Add new rectangle shape to slide
        Set oShape = oSlide.shapes.AddShape(msoShapeRectangle, 0, RectTop, Application.ActivePresentation.PageSetup.SlideWidth / 3, RectHeight)
                
        
        ' Add circle shape to rectangle
        CircleLeft = 20 '(Application.ActivePresentation.PageSetup.SlideWidth - CircleDiameter) / 2
        CircleTop = RectTop + (RectHeight - CircleDiameter) / 2
        Set oCircle = oSlide.shapes.AddShape(msoShapeOval, CircleLeft, CircleTop, CircleDiameter, CircleDiameter)
        
       ' rectangleNames(j + 1) = oCircle.Name
        
        
        ' Add text shape to circle
        TextLeft = CircleLeft + CircleDiameter / 2.5
        TextTop = CircleTop + CircleDiameter / 3
        Set oText = oSlide.shapes.AddTextbox(msoTextOrientationHorizontal, TextLeft, TextTop, CircleDiameter, CircleDiameter)
        oText.TextFrame2.textRange.Text = i ' Set text to instance number
        oText.TextFrame2.textRange.Font.Bold = msoTrue ' Make text bold
        oText.TextFrame2.textRange.Font.Size = 20
        oText.TextFrame2.VerticalAnchor = msoAnchorMiddle ' Align text vertically in center of circle
        
        'rectangleNames(j) = oShape.Name
      '  rectangleNames(j + 1) = oText.Name
        
        Set oText2 = oSlide.shapes.AddTextbox(msoTextOrientationHorizontal, TextLeft + 40, TextTop, Application.ActivePresentation.PageSetup.SlideWidth / 4, CircleDiameter)
        oText2.TextFrame2.textRange.Text = slidesArray(i, 2) ' Set text to instance number
        oText2.TextFrame2.textRange.Font.Bold = msoTrue ' Make text bold
        oText2.TextFrame2.VerticalAnchor = msoAnchorMiddle ' Align text vertically in center of circle
        
        
        ' Group circle and text shapes and move them to the front of the rectangle
        Set oGroup = oSlide.shapes.Range(Array(oCircle.Name, oText.Name)).Group
        oGroup.Name = "mn_gr_circle_" & oText.TextFrame.textRange.Text
        oGroup.ZOrder msoBringToFront
        
        ' Format rectangle shape as desired (e.g. fill color, outline, etc.)
        oShape.Fill.ForeColor.RGB = RGB(255, 0, 0) ' Red fill color
        oShape.Line.ForeColor.RGB = RGB(0, 0, 255) ' Blue outline color
        'add link to menu item
        
        With oShape.ActionSettings(ppMouseClick).Hyperlink
            .SubAddress = slidesArray(i, 3) & ". " & slidesArray(i, 4)
        End With
        
        Set oGroup2 = oSlide.shapes.Range(Array(oShape.Name, oText2.Name)).Group
        oGroup2.Name = "mn_gr_rec_" & oText2.TextFrame.textRange.Text
      
        
    Next i
  
   GroupShapesByName
   Dim r As shape
   Set r = FindShapeByName(oSlide.shapes, "MenuTbl")
   
   PasteShapeToMenuSlides r
   Highlightmenu
   GroupShapesByName
End Sub

Sub GroupShapesByName()
    Dim slide As slide
    Dim shape As shape
    Dim groupShape As shape
    Dim shapeArray() As String
    Dim i As Integer
    
    For Each slide In ActivePresentation.Slides
    i = 0
        For Each shape In slide.shapes
            If Left(shape.Name, 7) Like "mn_gr_*" Then
                ' Add shape to the array
                i = i + 1
                ReDim Preserve shapeArray(1 To i)
                shapeArray(i) = shape.Name
            End If
        Next shape
            ' Create a new group shape using the shape array
    If i > 1 Then
        slide.shapes.Range(shapeArray).Group.Name = "MenuTbl"
     
    End If
    Next slide
End Sub


Function FindShapeByName(shapes As shapes, shapeName As String) As shape
    Dim shape As shape
        For Each shape In shapes
            If shape.Name = shapeName Then
                Set FindShapeByName = shape
                Exit Function
            End If
        Next shape
    Set FindShapeByName = Nothing ' Shape not found
End Function

Sub PasteShapeToMenuSlides(ByVal shape As shape)
    Dim slide As slide
    Dim textBox As shape
    Dim r As shape
    
    For Each slide In ActivePresentation.Slides
        For Each textBox In slide.shapes
            If textBox.Type = msoTextBox Then
                If textBox.Name = "Menu" Then
                    shape.copy
                    slide.shapes.Paste
                    Exit For
                End If
            End If
        Next textBox
    Next slide
End Sub


Sub Highlightmenu()
    Dim slide As slide
    Dim textBox As shape
    Dim r As shape
      Dim shapeToFind As shape
      Dim shapeToFind2 As shape
      Dim rec As shape
    Dim shapeToFind3 As shape
    
    For Each slide In ActivePresentation.Slides
        For Each textBox In slide.shapes
            If textBox.Type = msoTextBox Then
                If textBox.Name = "Menu" Then
                    Debug.Print textBox.TextFrame.textRange.Text
                    
                    Set shapeToFind = FindShapeByName(slide.shapes, "MenuTbl")
                    
                    ' Ungroup the group shape
                    shapeToFind.Ungroup
                      
                    Set shapeToFind2 = FindShapeInGroup(slide.shapes, textBox.TextFrame.textRange.Text)
                    Set shapeToFind3 = FindRectangleInGroup(shapeToFind2)
                    
                    If Not shapeToFind3 Is Nothing Then
                        shapeToFind3.Fill.ForeColor.RGB = RGB(255, 255, 0) ' Change the color as needed
                    End If
                    
                    Exit For
                End If
            End If
        Next textBox
    Next slide
End Sub

Sub RemoveShapeByName()
    Dim slide As slide
    Dim shape As shape
    
    For Each slide In ActivePresentation.Slides
        For Each shape In slide.shapes
            If shape.Name = "MenuTbl" Then
                shape.Delete
                Exit For
            End If
        Next shape
    Next slide
End Sub


Function FindMenuShapes() As Variant
    Dim slideArray() As Variant
    Dim i As Long
    Dim j As Long
    Dim x As Integer
    Dim numSlides As Long
    x = 1
    
    numSlides = ActivePresentation.Slides.Count
    For i = 1 To numSlides
        For j = 1 To ActivePresentation.Slides(i).shapes.Count
            If InStr(1, ActivePresentation.Slides(i).shapes(j).Name, "menu", vbTextCompare) > 0 Then
                x = x + 1
                Exit For
            End If
        Next j
    Next i
    
    ReDim slideArray(1 To x - 1, 1 To 4)
    x = 1
    For i = 1 To numSlides
        For j = 1 To ActivePresentation.Slides(i).shapes.Count
            If InStr(1, ActivePresentation.Slides(i).shapes(j).Name, "menu", vbTextCompare) > 0 Then
                slideArray(x, 1) = i
                slideArray(x, 2) = ActivePresentation.Slides(i).shapes(j).TextFrame.textRange.Text
                slideArray(x, 3) = ActivePresentation.Slides(i).SlideNumber
                slideArray(x, 4) = ActivePresentation.Slides(i).Name
                
                x = x + 1
                Exit For
            End If
        Next j
    Next i
    
    'ReDim Preserve slideArray(x)
    FindMenuShapes = slideArray
End Function

Function FindRectangleInGroup(ByVal groupShape As shape) As shape
    Dim shape As shape
    
    For Each shape In groupShape.GroupItems
        If shape.Type = msoShapeRectangle Then
            Set FindRectangleInGroup = shape
            Exit Function
        End If
    Next shape
    
    Set FindRectangleInGroup = Nothing ' Rectangle not found in the group
End Function

Function FindShapeInGroup(shapes As shapes, targetText As String) As shape 'deze
    Dim shape As shape
    For Each shape In shapes
    Debug.Print shape.Type
        If shape.Type = msoGroup Then
            Dim groupShape As shape
            For Each groupShape In shape.GroupItems
                If groupShape.HasTextFrame Then
                    Dim textRange As textRange
                    Set textRange = groupShape.TextFrame.textRange
                    If InStr(1, textRange.Text, targetText, vbTextCompare) > 0 Then
                        Dim rectangleShape As shape
              
              
                        Set FindShapeInGroup = shape
                        Exit Function
                    End If
                End If
            Next
        End If
    Next
End Function

