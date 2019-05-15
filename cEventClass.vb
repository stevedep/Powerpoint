Option Explicit
Const COMPANY_NAME = "Whatever_U_May_Call_It INC"

Public WithEvents PPTEvent As Application

Private Sub Class_Initialize()
MsgBox "This demo traps some Presentation events:" & vbCrLf & _
       "NewPresentation, PresentationOpen, PresentationNewSlide, PresentationSave and PresentationPrint" & vbCrLf & vbCrLf & _
       "Some Slide Show Events like: " & vbCrLf & _
       "SlideShowBegin & SlideShowEnd" & vbCrLf & vbCrLf & _
       "And Slide View mode Window events like :" & vbCrLf & _
       "WindowSelectionChange", vbInformation + vbOKOnly, _
       "The EventHandler class has been initialized."
End Sub

Private Sub Class_Terminate()
MsgBox "EventHandler is now inactive.", vbInformation + vbOKOnly, "PowerPoint Event Handler Example"
End Sub


Private Sub PPTEvent_SlideShowNextSlide(ByVal Wn As SlideShowWindow)
    'MsgBox Wn.View.CurrentShowPosition + 1
    time (Wn.View.CurrentShowPosition)
End Sub




