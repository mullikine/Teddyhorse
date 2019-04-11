Attribute VB_Name = "modCursor"
Sub MoveCursor(ByVal X As Integer, ByVal Y As Integer)
Dim pos As POINTAPI

    GetCursorPos pos
    SetCursorPos pos.X + X, pos.Y + Y

End Sub
