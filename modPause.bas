Attribute VB_Name = "modPause"
Public Sub Pause(ByVal Interval As Single)
Dim t As Single

    t = Timer + Interval
    While t > Timer
        DoEvents
    Wend

End Sub
