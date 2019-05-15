Sub time()


    Dim sld As Slide
    Dim sld2 As Slide
    Dim shp As Shape

    Set sld = Application.ActiveWindow.View.Slide

   
    i = 1
    For Each shp In sld.Shapes
        If shp.Name = "txtTime" Then
           Id = i
        End If
        i = i + 1
    Next shp
    nu = Hour(Now) & ":" & Minute(Now)
    
    'MsgBox sld.Shapes(Id).TextFrame.TextRange.Text
    Status = DateDiff("n", sld.Shapes(Id).TextFrame.TextRange.Text, nu)
    
   i = 1
    For Each shp In sld.Shapes
        If shp.Name = "txtStatus" Then
           Id = i
        End If
        i = i + 1
    Next shp
    
    sld.Shapes(Id).TextFrame.TextRange.Text = Status
    
End Sub
