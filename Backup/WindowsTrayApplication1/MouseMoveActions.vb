Public Class MouseMoveActions
    ' global declarations 
    Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)
    Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
    Private Declare Function SetCursorPos Lib "User32.dll" (ByVal X As Integer, ByVal Y As Integer) As Integer
    Private Declare Function GetCursorPos Lib "User32.dll" (ByRef lpPoint As System.Drawing.Point) As Integer
    Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Integer) As Integer
    Private Const MOUSEEVENTF_LEFTDOWN = &H2
    Private Const MOUSEEVENTF_LEFTUP = &H4
    Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
    Private Const MOUSEEVENTF_MIDDLEUP = &H40
    Private Const MOUSEEVENTF_RIGHTDOWN = &H8
    Private Const MOUSEEVENTF_RIGHTUP = &H10
    Private Const MOUSEEVENTF_MOVE = &H1
    Private Const VK_LBUTTON = &H1
    Private Const VK_RBUTTON = &H2

    Public Shared Function X() As Integer
        Dim mousePoint As System.Drawing.Point
        GetCursorPos(mousePoint)
        Return mousePoint.X
    End Function

    Public Shared Function Y() As Integer
        Dim mousePoint As System.Drawing.Point
        GetCursorPos(mousePoint)
        Return mousePoint.Y
    End Function

    Public Shared Sub LeftClick()
        LeftDown()
        LeftUp()
    End Sub

    Public Shared Sub LeftDown()
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    End Sub

    Public Shared Sub LeftUp()
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    End Sub

    Public Shared Sub MiddleClick()
        MiddleDown()
        MiddleUp()
    End Sub

    Public Shared Sub MiddleDown()
        mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0)
    End Sub

    Public Shared Sub MiddleUp()
        mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0)
    End Sub

    Public Shared Sub RightClick()
        RightDown()
        RightUp()
    End Sub

    Public Shared Sub RightDown()
        mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
    End Sub

    Public Shared Sub RightUp()
        mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
    End Sub

    Public Shared Sub MoveMouse(ByVal xMove As Integer, ByVal yMove As Integer)
        'Dim curpoint As System.Drawing.Point
        'curpoint.X = X()
        'curpoint.Y = Y()
        'Dim xMovex As Integer
        'Dim yMovey As Integer
        ''Dim destreachd As Boolean = False

        ''While (destreachd = False)

        ' '' suppose curpoint is (169,132)    and destn point is (100,100)   
        ' '' actual movement should be (-69,-32)

        'If curpoint.X < xMove Then
        '    xMovex = xMove - curpoint.X  ' should come to 1
        'ElseIf curpoint.X = xMove Then
        '    xMovex = xMove - curpoint.X ' should come to 0
        'ElseIf curpoint.X > xMove Then
        '    xMovex = xMove - curpoint.X ' should come to -1
        'End If

        'If curpoint.Y < yMove Then
        '    yMovey = yMove - curpoint.Y  ' should come to 1
        'ElseIf curpoint.Y = yMove Then
        '    yMovey = yMove - curpoint.Y ' should come to 0
        'ElseIf curpoint.Y > yMove Then
        '    yMovey = yMove - curpoint.Y ' should come to -1
        'End If

        'curpoint.X = curpoint.X + xMovex
        'curpoint.Y = curpoint.Y + yMovey

        '    mouse_event(MOUSEEVENTF_MOVE, xMovex, yMovey, 0, 0)

        '    If ((curpoint.X = xMove) And (curpoint.Y = yMove)) Then
        '        destreachd = True
        '    Else
        '        destreachd = False
        '    End If

        'End While

        mouse_event(MOUSEEVENTF_MOVE, xMove, yMove, 0, 0)

    End Sub

    Public Shared Sub DragAndMoveMouse(ByVal xMove As Integer, ByVal yMove As Integer)
        'Dim curpoint As System.Drawing.Point
        'curpoint.X = X()
        'curpoint.Y = Y()
        'Dim xMovex As Integer ' = curpoint.X
        'Dim yMovey As Integer ' = curpoint.Y
        ''Dim destreachd As Boolean = False

        ''While (destreachd = False)

        'If curpoint.X < xMove Then
        '    xMovex = xMove - curpoint.X  ' should come to 1
        'ElseIf curpoint.X = xMove Then
        '    xMovex = xMove - curpoint.X ' should come to 0
        'ElseIf curpoint.X > xMove Then
        '    xMovex = xMove - curpoint.X ' should come to -1
        'End If

        'If curpoint.Y < yMove Then
        '    yMovey = yMove - curpoint.Y  ' should come to 1
        'ElseIf curpoint.Y = yMove Then
        '    yMovey = yMove - curpoint.Y ' should come to 0
        'ElseIf curpoint.Y > yMove Then
        '    yMovey = yMove - curpoint.Y ' should come to -1
        'End If

        ''curpoint.X = curpoint.X + xMovex
        ''curpoint.Y = curpoint.Y + yMovey

        ''mouse_event(MOUSEEVENTF_LEFTDOWN, xMovex, yMovey, 0, 0)

        ''If ((curpoint.X = xMove) And (curpoint.Y = yMove)) Then
        ''    destreachd = True
        ''Else
        ''    destreachd = False
        ''End If

        ''End While

        ''mouse left click on start point
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        ''move to destinATION POINT
        mouse_event(MOUSEEVENTF_MOVE, xMove, yMove, 0, 0)
        ''release mouse left click
        'mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    End Sub

    Public Shared Sub SetMousePos(ByVal xPos As Integer, ByVal yPos As Integer)
        SetCursorPos(xPos, yPos)
    End Sub

    Public Shared Sub MouseHide()
        Dim ReturnValue As Integer = ShowCursor(0)
    End Sub

    Public Shared Sub MouseShow()
        Dim ReturnValue As Integer = ShowCursor(1)
    End Sub

End Class
