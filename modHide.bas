Attribute VB_Name = "modHide"
Public Function hwndStartButton() As Long

   hwndStartButton = FindWindowEx(FindWindow("Shell_TrayWnd", ""), 0, "Button", vbNullString)
    
End Function

Public Sub HideHWND(ByVal hwnd As Long)

    ShowWindow hwnd, 0

End Sub

Public Sub ShowHWND(ByVal hwnd As Long)

    ShowWindow hwnd, 5

End Sub
