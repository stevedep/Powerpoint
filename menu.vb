Sub menu()
    Dim sld As Slide
    Dim sld2 As Slide
    Dim shp As Shape
        
    Set sld = Application.ActiveWindow.View.Slide
    wdth = Application.ActivePresentation.PageSetup.SlideWidth
    
    'count slides with menu text box
    Dim i As Integer
    i = 0
           For Each sld2 In ActivePresentation.Slides
            For Each shp In sld2.Shapes
                If shp.HasTextFrame Then
                    'If shp.Visible = msoFalse And shp.Name Like "* Menu *" Then
                    If shp.Name Like "*Menu*" Then
                             i = i + 1
                        End If
                End If
        Next shp
    Next sld2
    
    'add table
     With sld
            .Shapes.AddTable _
            NumRows:=1, _
            NumColumns:=i, _
            Left:=0, _
            Top:=0, _
            Width:=wdth, _
            Height:=5
    End With
    sld.Shapes(sld.Shapes.Count).Name = "MenuTbl"
    i = 1
    
    'remove first row from table
    For Each shp In sld.Shapes
     If shp.HasTable Then
        Id = i
     End If
     i = i + 1
    Next shp
    
    j = 1
    ActivePresentation.Slides(1).Shapes(Id).Table.FirstRow = False
    
    'loop slides, then loop shapes, search for textbox shape with name title and copy text content, and populate that value in the tablecell
    ' loop slids
        For Each sld2 In ActivePresentation.Slides
            'loop shapes
            For Each shp In sld2.Shapes
                'if textframe then continue
                If shp.HasTextFrame Then
                    'If shp.Visible = msoFalse And shp.Name Like "* Menu *" Then
                    ' if named Manu then copy value to table in the right column (index)
                    If shp.Name Like "*Menu*" Then
                             With ActivePresentation.Slides(1).Shapes(Id).Table.Cell(1, j).Shape.TextFrame.TextRange
                                .Font.Name = "Arial"
                                .Font.Size = 6
                                .Font.Bold = msoFalse
                                .Text = shp.TextFrame.TextRange.Text
                                End With
                            'set colour
                            ActivePresentation.Slides(1).Shapes(Id).Table.Cell(1, j).Shape.Fill.ForeColor.RGB = RGB(250, 250, 250)
                             'add link to menu item
                             With ActivePresentation.Slides(1).Shapes(Id).Table.Cell(1, j).Shape.TextFrame.TextRange.ActionSettings(ppMouseClick).Hyperlink
                                .SubAddress = sld2.SlideNumber & ". " & sld2.Name
                             End With
                             ' moving the next cell:
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
    
    'make sure to have the first navigation slide selected! From there we will search for the table and copy it. Only expects one table in the slide!
    For Each shp In sld.Shapes
        If shp.HasTable Then
           Id = i
        End If
        i = i + 1
    Next shp
    '  copy the table
    ActivePresentation.Slides(1).Shapes(Id).copy
    
    j = 1
    'loop over the slides, then the shapes, when we find a shaped Named Menu we will paste the table,
        For Each sld2 In ActivePresentation.Slides
        
            'titel van slide bepalen
            For Each shp In sld2.Shapes
                    If shp.HasTextFrame Then
                            'If shp.Visible = msoFalse And shp.TextFrame.TextRange.Text Like "*" & Title & "*" Then
                              If shp.Name Like "*Menu*" Then
                                       If sld2.SlideIndex > 1 Then
                                                sld2.Shapes.Paste
                                        End If
                                        sld2.ColorScheme.Colors(ppActionHyperlink).RGB = RGB(90, 90, 90)
                                        titel = shp.TextFrame.TextRange.Text
                                             ' get the menu table id
                                            i = 1
                                            For Each shp2 In sld2.Shapes
                                               If shp2.Name = "MenuTbl" Then
                                                Id = i
                                             End If
                                             i = i + 1
                                            Next shp2
                                        
                                            ' loop over the cells and make the background darker
                                            For j = 1 To sld2.Shapes(Id).Table.Columns.Count
                                                If sld2.Shapes(Id).Table.Cell(1, j).Shape.TextFrame.TextRange.Text = titel Then
                                                    sld2.Shapes(Id).Table.Cell(1, j).Shape.Fill.ForeColor.RGB = RGB(217, 217, 217)
                                                End If
                                            Next j
                            End If
                    End If
            Next shp
        Next sld2
End Sub


Sub new_font_size()
    Dim sld As Slide
    Dim sld2 As Slide
    Dim shp As Shape
    Set sld = Application.ActiveWindow.View.Slide
    i = 1
    FontSize = InputBox("Increase FontSize by how much")
    
    j = 1
    'loop over the slides, then the shapes, when we find a shaped Named Menu we will paste the table,
        For Each sld2 In ActivePresentation.Slides
            For Each shp In sld2.Shapes
                     ' get the menu table id
                    i = 1
                    For Each shp2 In sld2.Shapes
                       If shp2.Name = "MenuTbl" Then
                        Id = i
                        For j = 1 To sld2.Shapes(Id).Table.Columns.Count
                         sld2.Shapes(Id).Table.Cell(1, j).Shape.TextFrame.TextRange.Font.Size = FontSize
                    Next j
                     End If
                     i = i + 1
                    Next shp2
                    ' loop over the cells and increase the font
            Next shp
        Next sld2
End Sub
