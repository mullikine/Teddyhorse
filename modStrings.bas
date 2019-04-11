Attribute VB_Name = "modStrings"

Public sStrings(10) As String


Sub Init()

    sStrings(0) = "Important" & vbCrLf & vbCrLf & _
                    "The following procedure might require you to restart your computer, which will close this troubleshooter. If possible, view this troubleshooter on another computer while you perform the steps on the computer you are troubleshooting." & vbCrLf & _
                    "To continue troubleshooting if no other computer is available" & vbCrLf & vbCrLf & _
                    "Right-click the page displayed on your screen, and then click Print." & vbCrLf & _
                    "Follow the steps in the printed copy of the procedure." & vbCrLf & _
                    "After your computer restarts, reopen this troubleshooter and answer each question as you answered it initially." & vbCrLf & vbCrLf & _
                    "When you reach this page again, answer the question at the bottom, and then click Next." & vbCrLf & _
                    "Some display properties, if set incorrectly for your display adapter (video card), can prevent Windows from working correctly. Adjusting these settings might fix the problem." & vbCrLf & vbCrLf & _
                    "If you cannot see anything on your computer screen, restart your computer in safe mode and adjust your display settings"

End Sub
