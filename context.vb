Sub loop_excel()
 'load excel
     Dim app As New Excel.Application
     app.Visible = False 'Visible is False by default, so this isn't necessary
     Dim book As Excel.Workbook
     Set book = app.Workbooks.Add(FileOpenDialogBox)
    

    Dim planningslidenr As Integer
    planningslidenr = InputBox("Which Slide contains the planning")
    Set lstActivities = book.Worksheets("Sheet1").ListObjects(1)
        For i = 1 To lstActivities.ListRows.Count
            'If lstActivities.ListColumns("ID").DataBodyRange(i).Value = "1" Then
                'MsgBox lstActivities.ListColumns("Name").DataBodyRange(i).Value
               
               If Len(LTrim(RTrim(lstActivities.ListColumns("External ID").DataBodyRange(i).Value))) > 0 Then
                   
                   If context_new(lstActivities.ListColumns("External ID").DataBodyRange(i).Value, planningslidenr) = 1 Then
                    
                    Set currentslide = ActivePresentation.Slides(ActiveWindow.View.Slide.SlideIndex)
                    Set tr = currentslide.Shapes(currentslide.Shapes.Count).TextFrame.TextRange
                    tr.Font.Size = 12
                    tr.Text = lstActivities.ListColumns("Title").DataBodyRange(i).Value & _
                    vbNewLine & vbNewLine & "Objective" & vbNewLine & _
                    lstActivities.ListColumns("Description").DataBodyRange(i).Value & _
                    vbNewLine & vbNewLine & "Approach" & vbNewLine & _
                    lstActivities.ListColumns("Generic03").DataBodyRange(i).Value & _
                    vbNewLine & vbNewLine & "Input" & vbNewLine & _
                    lstActivities.ListColumns("Generic02").DataBodyRange(i).Value & _
                    vbNewLine & vbNewLine & "Output / Deliverable" & vbNewLine & _
                    lstActivities.ListColumns("Acceptance Criteria").DataBodyRange(i).Value & _
                    vbNewLine & vbNewLine & "Responsible" & vbNewLine & _
                    lstActivities.ListColumns("Assigned To").DataBodyRange(i).Value & _
                    vbNewLine & vbNewLine & "Timing" & vbNewLine & _
                    lstActivities.ListColumns("Iteration Path").DataBodyRange(i).Value '& _
                   ' vbNewLine & vbNewLine & "Status" & vbNewLine & _
                   ' lstActivities.ListColumns("Status").DataBodyRange(i).Value
    
    Debug.Print "id:"
    Debug.Print lstActivities.ListColumns("External ID").DataBodyRange(i).Value & vbNewLine
    
                    tr.Paragraphs(1).Font.Bold = msoTrue
                    tr.Paragraphs(1).Font.Size = 18
                    tr.Paragraphs(1).Font.Color = RGB(0, 137, 196)
                    
                  Dim arr As Variant
                    arr = Array("Objective", "Approach", "Input", "Output / Deliverable", "Responsible", "Timing")
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
                    
                    'tr.Paragraphs(3).Font.Bold = msoTrue
                    'tr.Paragraphs(6).Font.Bold = msoTrue
                    'tr.Paragraphs(9).Font.Bold = msoTrue
                    'tr.Paragraphs(11).Font.Bold = msoTrue
                    'tr.Paragraphs(15).Font.Bold = msoTrue
                    'tr.Paragraphs(18).Font.Bold = msoTrue
                'tr.Paragraphs(21).Font.Bold = msoTrue
                
                ' tr.Lines(5).Font.Bold = msoTrue
               End If
            End If
            'End If
        Next i
                    
'close off
    book.Close SaveChanges:=False
    app.Quit
    Set app = Nothing
    Set shps = Nothing
    Set currentslide = Nothing
    Set tr = Nothing
    
    'MsgBox "done"

End Sub

Function context_new(id As String, planningslidenr As Integer) As Integer
'select slide
context_new = 0
    
    Set shps = ActivePresentation.Slides(CInt(planningslidenr)).Shapes

    Dim a As Integer

    For a = 1 To shps.Count
            If shps(a).Name = id Then
                context CInt(planningslidenr), a
                context_new = 1
            End If
    Next a
    
   

End Function


Sub context(slidenr As Integer, shapenr As Integer)

   
    
    Dim sld As Slide
    Dim shp As Shape
    
    Set sld = ActivePresentation.Slides(slidenr) 'Application.ActiveWindow.View.Slide
    Set shp = sld.Shapes(shapenr)
    
    sld.Duplicate
    
      
    y = shp.Top ' ActiveWindow.Selection.ShapeRange.Top
    x = shp.Left 'ActiveWindow.Selection.ShapeRange.Left
    w = shp.Width 'ActiveWindow.Selection.ShapeRange.Width
    h = shp.Height 'ActiveWindow.Selection.ShapeRange.Height
    
   
    shp.Copy 'Windows(1).Selection.Copy
  '  MsgBox sld.Shapes.Count
    Set sld = ActivePresentation.Slides(slidenr + 1)
    sld.Shapes.Paste
     sld.Shapes(sld.Shapes.Count).Top = y
     sld.Shapes(sld.Shapes.Count).Left = x
     
     n = sld.Shapes.Count
     sld.Shapes.AddShape Type:=msoShapeRectangle, _
    Left:=0, Top:=0, Width:=960, Height:=540
     sld.Shapes(sld.Shapes.Count).Fill.ForeColor.RGB = RGB(172, 185, 202)
     sld.Shapes(sld.Shapes.Count).Fill.Transparency = 0.35
     ActiveWindow.View.GotoSlide slidenr + 1
    
    sld.Shapes.Range(Array(sld.Shapes.Count, n)).Select
    ActiveWindow.Selection.ShapeRange.MergeShapes msoMergeCombine
        
    
    
    
    Set shp = Nothing
    Dim ffb As FreeformBuilder
    Dim myshape As Shape
    Dim currentslide As Slide
    
    Set currentslide = sld 'ActivePresentation.Slides(ActiveWindow.View.Slide.SlideIndex)

    'Links = MsgBox("Left?", vbYesNo + vbQuestion, "Side")
    
    If x > 600 Then
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
            
           currentslide.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                Left:=x + w + 40, Top:=30, Width:=300, Height:=500).TextFrame _
                .TextRange.Text = ""

        
    
    
    End If
    
    Set myshape = ffb.ConvertToShape
    myshape.Fill.ForeColor.RGB = RGB(256, 256, 256)
    
    'MsgBox currentslide.Shapes(currentslide.Shapes.Count).Name
    currentslide.Shapes(currentslide.Shapes.Count).ZOrder (msoSendBackward)
   ' MsgBox currentslide.Shapes(currentslide.Shapes.Count).TextFrame.TextRange.Lines.Count
    
    'Set tr = currentslide.Shapes(currentslide.Shapes.Count).TextFrame.TextRange
    
    'tr.Lines(2, 2).Font.Italic = msoTrue
    
    
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




