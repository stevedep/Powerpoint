Sub menu()


    Dim sld As Slide
    Dim sld2 As Slide
    Dim shp As Shape
        
    Set sld = Application.ActiveWindow.View.Slide
     With sld
            .Shapes.AddTable _
            NumRows:=1, _
            NumColumns:=ActivePresentation.Slides.Count - 2, _
            Left:=0, _
            Top:=0, _
            Width:=700, _
            Height:=5
    End With
    i = 1
    
    For Each shp In sld.Shapes
     If shp.HasTable Then
        Id = i
     End If
     i = i + 1
    Next shp
    
    j = 1
    ActivePresentation.Slides(1).Shapes(Id).Table.FirstRow = False
    
        For Each sld2 In ActivePresentation.Slides
            For Each shp In sld2.Shapes
                If shp.HasTextFrame Then
                    If shp.Visible = msoFalse And shp.Name  Like "*" & Title & "*" Then
                             With ActivePresentation.Slides(1).Shapes(Id).Table.Cell(1, j).Shape.TextFrame.TextRange
                                .Font.Name = "Arial"
                                .Font.Size = 6
                                .Font.Bold = msoFalse
                                
                                .Text = shp.TextFrame.TextRange.Text
                                End With
                            
                            ActivePresentation.Slides(1).Shapes(Id).Table.Cell(1, j).Shape.Fill.ForeColor.RGB = RGB(250, 250, 250)
                            
                             
                             With ActivePresentation.Slides(1).Shapes(Id).Table.Cell(1, j).Shape.TextFrame.TextRange.ActionSettings(ppMouseClick).Hyperlink
                                .SubAddress = sld2.SlideNumber & ". " & sld2.Name
                             End With
                             j = j + 1
                        End If
                End If
        Next shp
    Next sld2


End Sub

Sub copy()


    Dim sld As Slide
    Dim sld2 As Slide
    Dim shp As Shape
    
    Set sld = Application.ActiveWindow.View.Slide
       i = 1
    
    For Each shp In sld.Shapes
        If shp.HasTable Then
           Id = i
        End If
        i = i + 1
    Next shp
    
    ActivePresentation.Slides(1).Shapes(Id).copy
    
    j = 1
        For Each sld2 In ActivePresentation.Slides
           If sld2.SlideIndex > 1 Then
            sld2.Shapes.Paste
           End If
           sld2.ColorScheme.Colors(ppActionHyperlink).RGB = RGB(90, 90, 90)
        
            'titel van slide bepalen
            For Each shp In sld2.Shapes
                    If shp.HasTextFrame Then
                            If shp.Visible = msoFalse And shp.TextFrame.TextRange.Text Like "*" & Title & "*" Then
                                titel = shp.TextFrame.TextRange.Text
                            End If
                    End If
            Next shp
        
            'tabel zoeken
            i = 1
            For Each shp2 In sld2.Shapes
             If shp2.HasTable Then
                Id = i
             End If
             i = i + 1
            Next shp2
        
            'cellen nalopen
            For j = 1 To sld2.Shapes(Id).Table.Columns.Count
                If sld2.Shapes(Id).Table.Cell(1, j).Shape.TextFrame.TextRange.Text = titel Then
                    sld2.Shapes(Id).Table.Cell(1, j).Shape.Fill.ForeColor.RGB = RGB(217, 217, 217)
                End If
            Next j
            
        
        Next sld2


End Sub
