
Sub loop_excel()
 'load excel
    Dim app As New Excel.Application
    app.Visible = False 'Visible is False by default, so this isn't necessary
    Dim book As Excel.Workbook
    Set book = app.Workbooks.Add("C:\Users\310267217\OneDrive - Philips\vba\contexttable.xlsx")

    Set lstActivities = book.Worksheets("Sheet1").Range("Table2").ListObject
        For i = 1 To lstActivities.ListRows.Count
            If lstActivities.ListColumns("ID").DataBodyRange(i).Value = "1" Then
                'MsgBox lstActivities.ListColumns("Name").DataBodyRange(i).Value
               context_new lstActivities.ListColumns("ID").DataBodyRange(i).Value
                
                Set currentslide = ActivePresentation.Slides(ActiveWindow.View.Slide.SlideIndex)
                Set tr = currentslide.Shapes(currentslide.Shapes.Count).TextFrame.TextRange
                tr.Font.Size = 12
                tr.Text = lstActivities.ListColumns("Name").DataBodyRange(i).Value & _
                vbNewLine & vbNewLine & "Objective" & vbNewLine & _
                lstActivities.ListColumns("Objective").DataBodyRange(i).Value & _
                vbNewLine & vbNewLine & "Approach" & vbNewLine & _
                lstActivities.ListColumns("Approach").DataBodyRange(i).Value & _
                vbNewLine & vbNewLine & "Input" & vbNewLine & _
                lstActivities.ListColumns("Input").DataBodyRange(i).Value & _
                vbNewLine & vbNewLine & "Output / Deliverable" & vbNewLine & _
                lstActivities.ListColumns("Output").DataBodyRange(i).Value & _
                vbNewLine & vbNewLine & "Responsible" & vbNewLine & _
                lstActivities.ListColumns("Responsible").DataBodyRange(i).Value & _
                vbNewLine & vbNewLine & "Timing" & vbNewLine & _
                lstActivities.ListColumns("Timing").DataBodyRange(i).Value & _
                vbNewLine & vbNewLine & "Status" & vbNewLine & _
                lstActivities.ListColumns("Status").DataBodyRange(i).Value

                tr.Lines(2, 1).Font.Bold = msoTrue
                 tr.Lines(5, 1).Font.Bold = msoTrue
                
              
               
            End If
        Next i
                    
'close off
    book.Close SaveChanges:=False
    app.Quit
    Set app = Nothing
    Set shps = Nothing
    Set currentslide = Nothing
    Set tr = Nothing
    
    MsgBox "done"

End Sub

Sub context_new(id As String)
'select slide
    planningslidenr = InputBox("Which Slide contains the planning")
    Set shps = ActivePresentation.Slides(CInt(planningslidenr)).Shapes

    Dim a As Integer

    For a = 1 To shps.Count
            If shps(a).Name = "1" Then
                context CInt(planningslidenr), a
            End If
    Next a
    
   

End Sub


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
     sld.Shapes(sld.Shapes.Count).Fill.Transparency = 0.2
     ActiveWindow.View.GotoSlide slidenr + 1
    
    sld.Shapes.Range(Array(sld.Shapes.Count, n)).Select
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
            
           currentslide.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                Left:=x + 80, Top:=30, Width:=300, Height:=500).TextFrame _
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




