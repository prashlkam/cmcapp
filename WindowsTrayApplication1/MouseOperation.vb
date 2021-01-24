Imports WindowsTrayApplication1.MouseMoveActions
Imports WindowsTrayApplication1.ChangeCursor

Public Class MouseOperation

    Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

    Private Const VK_LBUTTON = &H1

    Public cc As ChangeCursor = New ChangeCursor()

    Public MoveCursorBy As System.Drawing.Point
    Public timer1 As Timer = New Timer
    Public mclk As Boolean
    Public firstclick As Boolean
    Public mousemoving As Boolean
    Public dragaction As Boolean
    Public mouseaction As UShort
    Private isactive As Boolean = False

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
            Case 1, 6, 7, 8, 9, 10
                firstclick = True
                MoveCursorBy.X = 0
                MoveCursorBy.Y = 0
                OtherMouseActions()
                firstclick = False
            Case 2
                '' move up
                MoveCursorBy.X = 0
                MoveCursorBy.Y = -1
                If dragaction = True Then
                    cc.SetCursor(mouseaction - 1, dragaction)
                    DragAndMoveMouse(MoveCursorBy.X, MoveCursorBy.Y)
                Else
                    cc.SetCursor(mouseaction - 1, dragaction)
                    MoveMouse(MoveCursorBy.X, MoveCursorBy.Y)
                End If

            Case 3
                '' move right
                MoveCursorBy.X = 1
                MoveCursorBy.Y = 0
                If dragaction = True Then
                    cc.SetCursor(mouseaction - 1, dragaction)
                    DragAndMoveMouse(MoveCursorBy.X, MoveCursorBy.Y)
                Else
                    cc.SetCursor(mouseaction - 1, dragaction)
                    MoveMouse(MoveCursorBy.X, MoveCursorBy.Y)
                End If

            Case 4
                '' move down 
                MoveCursorBy.X = 0
                MoveCursorBy.Y = 1
                If dragaction = True Then
                    cc.SetCursor(mouseaction - 1, dragaction)
                    DragAndMoveMouse(MoveCursorBy.X, MoveCursorBy.Y)
                Else
                    cc.SetCursor(mouseaction - 1, dragaction)
                    MoveMouse(MoveCursorBy.X, MoveCursorBy.Y)
                End If

            Case 5
                '' move left
                MoveCursorBy.X = -1
                MoveCursorBy.Y = 0
                If dragaction = True Then
                    cc.SetCursor(mouseaction - 1, dragaction)
                    DragAndMoveMouse(MoveCursorBy.X, MoveCursorBy.Y)
                Else
                    cc.SetCursor(mouseaction - 1, dragaction)
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
        isactive = True
        timer1.Start()
        While isactive = True
            timer1_Timer()
        End While

    End Sub

    Public Sub OtherMouseActions()
        dragaction = False
        Select Case mouseaction
            Case 6
                '' single click
                cc.SetCursor(1, dragaction)
                LeftClick()
            Case 7
                '' double click
                cc.SetCursor(1, dragaction)
                LeftClick()
                LeftClick()
            Case 8
                '' drag action
                dragaction = True
                cc.SetCursor(0, dragaction)
            Case 9
                '' right click
                RightClick()
            Case 10
                '' exit loop
                isactive = False
                cc.SetCursor(10, dragaction)
        End Select
    End Sub

End Class
