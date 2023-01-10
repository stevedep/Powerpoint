Sub create_callouts()
 'load excel
     Dim app As New Excel.Application
     app.Visible = False 'Visible is False by default, so this isn't necessary
     Dim book As Excel.Workbook
     Set book = app.Workbooks.Add(FileOpenDialogBox)
    
    Dim planningslidenr As Integer
    planningslidenr = InputBox("Which Slide contains the planning")
    
    'retreive the first table
    Set lstActivities = book.Worksheets("Sheet1").ListObjects(1)
    

    'for each task
        For i = 1 To lstActivities.ListRows.Count
                'if the task ID is greater than 0
               If Len(LTrim(RTrim(lstActivities.ListColumns("ID").DataBodyRange(i).Value))) > 0 Then
                   ' start creating a new fn_new_callout, navigate to the function to learn more.
                   If fn_search_bar_and_create_callout(lstActivities.ListColumns("ID").DataBodyRange(i).Value, planningslidenr) = 1 Then
                    'once the new slide with the shapes has been created we populate the last shape (text box with text)
                    Set currentslide = ActivePresentation.Slides(ActiveWindow.View.Slide.SlideIndex)
                    Set tr = currentslide.Shapes(currentslide.Shapes.Count).TextFrame.TextRange
                    tr.Font.Size = 12
                    'We populate the textbox over here, for some reason sorting/filtering needed to be disables for this to work properly
                    tr.Text = lstActivities.ListColumns("Title").DataBodyRange(i).Value & _
                    vbNewLine & vbNewLine & "Objective" & vbNewLine & _
                    lstActivities.ListColumns("Objective").DataBodyRange(i).Value & _
                    vbNewLine & vbNewLine & "Approach" & vbNewLine & _
                    lstActivities.ListColumns("Approach").DataBodyRange(i).Value & _
                    vbNewLine & vbNewLine & "Responsible" & vbNewLine & _
                    lstActivities.ListColumns("Responsible").DataBodyRange(i).Value
    
                    'set some formatting.
                    tr.Paragraphs(1).Font.Bold = msoTrue
                    tr.Paragraphs(1).Font.Size = 18
                    tr.Paragraphs(1).Font.Color = RGB(0, 137, 196)
                    
                   'We want to make certain fields bold.
                  Dim arr As Variant
                    arr = Array("Objective", "Approach", "Responsible")
                  For Each word In arr
                    Set foundText = tr.Find(FindWhat:=word)
                    Do While Not (foundText Is Nothing)
                        With foundText
                            .Font.Bold = True
                            Set foundText = _
                                tr.Find(FindWhat:=word, _
                                After:=.Start + .Length - 1)
                        End With
                    Loop
                Next
               End If
            End If
  
        Next i
                    
'close off
    book.Close SaveChanges:=False
    app.Quit
    Set app = Nothing
    Set shps = Nothing
    Set currentslide = Nothing
    Set tr = Nothing
    

End Sub

Function fn_search_bar_and_create_callout(id As String, planningslidenr As Integer) As Integer

fn_search_bar_and_create_callout = 0
    Set shps = ActivePresentation.Slides(CInt(planningslidenr)).Shapes
    Dim a As Integer

' We iterate over the shapes in search for the one that has a name that is equal to the id of the task in excel
    For a = 1 To shps.Count
            If shps(a).Name = id Then
                'once found we create a new call out
                fn_new_callout CInt(planningslidenr), a
                fn_search_bar_and_create_callout = 1
            End If
    Next a

End Function


Sub fn_new_callout(slidenr As Integer, shapenr As Integer)
    Dim sld As Slide
    Dim shp As Shape
    
    Set sld = ActivePresentation.Slides(slidenr)
    Set shp = sld.Shapes(shapenr)
    
    sld.Duplicate 'duplicate the planning slide to add shapes to that one
    
    ' Copy the x, y, width etc from the bar
    y = shp.Top: x = shp.Left: w = shp.Width: h = shp.Height
       
    Set sld = ActivePresentation.Slides(slidenr + 1)
   
    'Add the semi transparant overlay
     sld.Shapes.AddShape Type:=msoShapeRectangle, _
    Left:=0, Top:=0, Width:=960, Height:=540
     sld.Shapes(sld.Shapes.Count).Fill.ForeColor.RGB = RGB(172, 185, 202)
     sld.Shapes(sld.Shapes.Count).Fill.Transparency = 0.35
     ActiveWindow.View.GotoSlide slidenr + 1
    
    ' bring the bar back to the front
    sld.Shapes(shapenr).ZOrder (msoBringToFront)
            
    ' cleanup for this part
    Set shp = Nothing
    Dim ffb As FreeformBuilder
    Dim myshape As Shape
    Dim currentslide As Slide
        
    Set currentslide = sld 'bit redundant..
    
    ' depending on the x position of the bar we draw the call out either on the left or right side
    If x > Application.ActivePresentation.PageSetup.SlideWidth * 0.4 Then
           Set ffb = currentslide.Shapes.BuildFreeform(msoEditingCorner, x, y)
            With ffb
                .AddNodes msoSegmentLine, msoEditingAuto, x - 40, 135
                .AddNodes msoSegmentLine, msoEditingAuto, x - 40, 30
                .AddNodes msoSegmentLine, msoEditingAuto, x - 300, 30
                .AddNodes msoSegmentLine, msoEditingAuto, x - 300, 520
                .AddNodes msoSegmentLine, msoEditingAuto, x - 40, 520
                .AddNodes msoSegmentLine, msoEditingAuto, x - 40, 350
                .AddNodes msoSegmentLine, msoEditingAuto, x, y + h
            End With
            
            currentslide.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                Left:=x - 300, Top:=30, Width:=250, Height:=500).TextFrame _
                .TextRange.Text = ""
    Else
             Set ffb = currentslide.Shapes.BuildFreeform(msoEditingCorner, x + w, y)
            With ffb
                .AddNodes msoSegmentLine, msoEditingAuto, x + w + 40, 135
                .AddNodes msoSegmentLine, msoEditingAuto, x + w + 40, 30
                .AddNodes msoSegmentLine, msoEditingAuto, x + w + 40 + 300, 30
                .AddNodes msoSegmentLine, msoEditingAuto, x + w + 40 + 300, 520
                .AddNodes msoSegmentLine, msoEditingAuto, x + w + 40, 520
                .AddNodes msoSegmentLine, msoEditingAuto, x + w + 40, 350
                .AddNodes msoSegmentLine, msoEditingAuto, x + w, y + h
            End With
            'adding a textbox
           currentslide.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                Left:=x + w + 40, Top:=30, Width:=300, Height:=500).TextFrame _
                .TextRange.Text = ""
    End If
    'making the call out white
    Set myshape = ffb.ConvertToShape
    myshape.Fill.ForeColor.RGB = RGB(256, 256, 256)
    currentslide.Shapes(currentslide.Shapes.Count).ZOrder (msoSendBackward)
    
    Set tb = Nothing
    Set sld = Nothing
    Set myshape = Nothing
    Set currentslide = Nothing

End Sub

Function FileOpenDialogBox() As String
 
'Display a Dialog Box that allows to select a single file.
'The path for the file picked will be stored in fullpath variable
  With Application.FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        'Filter to just the following types of files to narrow down selection options
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb", 1
        'Show the dialog box
        .Show
        
        'Store in fullpath variable
        FileOpenDialogBox = .SelectedItems.Item(1)
        
    End With
 
End Function







