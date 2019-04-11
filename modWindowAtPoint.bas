Attribute VB_Name = "modWindowAtPoint"
'Dim hWndOver As Long
'Dim hWndParent As Long
'Dim sParentClassName As String * 100
'Dim wID As Long

'Dim hInstance As Long
'Dim sParentWindowText As String * 100
'Dim sModuleFileName As String * 100


Public Function WindowAtCursor() As Long

Dim pt32 As POINTAPI

    Call GetCursorPos(pt32)
    WindowAtCursor = WindowFromPointXY(pt32.x, pt32.y)
    
End Function

Public Function WindowClassName(ByVal hwnd As Long) As String
Dim sClassName As String * 100
Dim r As Long

       r = GetClassName(hWndOver, sClassName, 100)         ' Window Class
       tClassName = Left(sClassName, r)

End Function

Public Function WindowText(ByVal hwnd As Long) As String
Dim sWindowText As String * 100
Dim r As Long

    r = GetWindowText(hwnd, sWindowText, 100)
    WindowText = Left(sWindowText, r)

End Function

Public Function WindowStyle(ByVal hwnd As Long) As Long

    WindowStyle = GetWindowLong(hwnd, GWL_STYLE)

End Function

'Public Sub SaveWindowDetails()
'
'
'
'
'       ' Get handle of parent window:
'       hWndParent = GetParent(hWndOver)
'       tPHandle = hWndParent
'
'       ' If there is a parent get more info:
'       If hWndParent <> 0 Then
'          ' Get ID of window:
'          wID = GetWindowWord(hWndOver, GWW_ID)
'          tPID = wID
'          tPHandle = hWndParent
'
'          ' Get the text of the Parent window:
'          r = GetWindowText(hWndParent, sParentWindowText, 100)
'          tPText = Left(sParentWindowText, r)
'
'          ' Get the class name of the parent window:
'          r = GetClassName(hWndParent, sParentClassName, 100)
'          tPClassName = Left(sParentClassName, r)
'       Else
'          ' Update fields when no parent:
'          tPID = -1
'          tPHandle = -1
'          tPText = ""
'          tPClassName = ""
'       End If
'
'       ' Get window instance:
'       hInstance = GetWindowWord(hWndOver, GWW_HINSTANCE)
'
'       ' Get module file name:
'       r = GetModuleFileName(hInstance, sModuleFileName, 100)
'       tModuleFileName = Left(sModuleFileName, r)
'    End If
'
'End Sub
