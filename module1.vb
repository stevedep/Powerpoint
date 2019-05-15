
Dim cPPTObject As New cEventClass
Dim TrapFlag As Boolean

Sub TrapEvents()
If TrapFlag = True Then
   MsgBox "Relax, my friend, the EventHandler is already active.", vbInformation + vbOKOnly, "PowerPoint Event Handler Example"
   Exit Sub
End If
   Set cPPTObject.PPTEvent = Application
   TrapFlag = True
End Sub

Sub ReleaseTrap()
If TrapFlag = True Then
   Set cPPTObject.PPTEvent = Nothing
   Set cPPTObject = Nothing
   TrapFlag = False
End If
End Sub



Sub time(sldnr As Integer)


    Dim Sld As Slide
    Dim sld2 As Slide
    Dim shp As Shape

    Set Sld = ActivePresentation.Slides(sldnr)
    
    i = 1
    For Each shp In Sld.Shapes
        If shp.Name = "txtTime" Then
           Id = i
        End If
        i = i + 1
    Next shp
    nu = Hour(Now) & ":" & Minute(Now)
    
    'MsgBox sld.Shapes(Id).TextFrame.TextRange.Text
   If Id Then
    Status = DateDiff("n", Sld.Shapes(Id).TextFrame.TextRange.Text, nu)
    
   i = 1
    For Each shp In Sld.Shapes
        If shp.Name = "txtStatus" Then
           Id = i
        End If
        i = i + 1
    Next shp
    
    Sld.Shapes(Id).TextFrame.TextRange.Text = Status
    
    End If
    
End Sub
