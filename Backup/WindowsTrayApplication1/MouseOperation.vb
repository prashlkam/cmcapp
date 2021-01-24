Imports WindowsTrayApplication1.MouseMoveActions
Imports WindowsTrayApplication1.ChangeCursor

Public Class MouseOperation

    Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

    Private Const VK_LBUTTON = &H1


    Public MoveCursorBy As System.Drawing.Point
    Public timer1 As Timer = New Timer
    Public mclk As Boolean
    Public firstclick As Boolean
    Public mousemoving As Boolean
    Public dragaction As Boolean
    Public mouseaction As UShort

    Public Sub timer1_Timer()

        If mouseaction > 9 Then
            mouseaction = 1
        End If
        If dragaction = True And mouseaction > 5 Then
            mouseaction = 1
        End If

        mclk = DetectMouseLeftClick()

        If mouseaction >= 2 And mouseaction <= 5 And mclk = True Then
            firstclick = False
            mousemoving = True
        End If

        Select Case mouseaction
            Case 1, 6, 7, 8, 9
                firstclick = True
                MoveCursorBy.X = 0
                MoveCursorBy.Y = 0
                OtherMouseActions()
                firstclick = False
            Case 2
                MoveCursorBy.X = 0
                MoveCursorBy.Y = -1
                If dragaction = True Then
                    DragAndMoveMouse(MoveCursorBy.X, MoveCursorBy.Y)
                Else
                    MoveMouse(MoveCursorBy.X, MoveCursorBy.Y)
                End If

            Case 3
                MoveCursorBy.X = 1
                MoveCursorBy.Y = 0
                If dragaction = True Then
                    DragAndMoveMouse(MoveCursorBy.X, MoveCursorBy.Y)
                Else
                    MoveMouse(MoveCursorBy.X, MoveCursorBy.Y)
                End If

            Case 4
                MoveCursorBy.X = 0
                MoveCursorBy.Y = 1
                If dragaction = True Then
                    DragAndMoveMouse(MoveCursorBy.X, MoveCursorBy.Y)
                Else
                    MoveMouse(MoveCursorBy.X, MoveCursorBy.Y)
                End If

            Case 5
                MoveCursorBy.X = -1
                MoveCursorBy.Y = 0
                If dragaction = True Then
                    DragAndMoveMouse(MoveCursorBy.X, MoveCursorBy.Y)
                Else
                    MoveMouse(MoveCursorBy.X, MoveCursorBy.Y)
                End If

        End Select
        If mousemoving = False Then

            mouseaction = mouseaction + 1

        End If

    End Sub

    Private Shared Function DetectMouseLeftClick() As Boolean

        If GetAsyncKeyState(VK_LBUTTON) Then
            Return True
        End If

        Return False

    End Function

    Public Sub mouseop()

        timer1.Interval = 1000
        timer1.Start()
        timer1_Timer()


    End Sub

    Public Sub OtherMouseActions()
        Select Case mouseaction
            Case 6
                LeftClick()
            Case 7
                LeftClick()
                LeftClick()
            Case 8
                dragaction = True
            Case 9
                RightClick()
        End Select
    End Sub

End Class
