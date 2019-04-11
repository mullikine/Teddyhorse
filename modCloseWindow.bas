Attribute VB_Name = "modCloseWindow"
Option Explicit

Public Sub QuitFromTitle(ByVal sTitle As String)
Dim iHwnd As Long
Dim ihTask As Long
Dim iReturn As Long

     iHwnd = FindWindow(0&, sTitle)
     iReturn = PostMessage(iHwnd, WM_QUIT, 0&, 0&)

End Sub
