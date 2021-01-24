Imports System.String


Public Class ChangeCursor

    Declare Function SetSystemCursor Lib "user32.dll" (ByVal hcur As System.Int32, ByVal id As System.Int32) As System.Int32
    Declare Function LoadCursorFromFile Lib "user32.dll" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As System.Int32

	Private Const OCR_NORMAL As System.Int32 = 32512

    Public CursorNames(10) As String


    Public Shared Sub InitCursorValues(ByRef CursorNames)

		CursorNames(0) = ".\Resources\Anim Cursors\ADITI CC cur01.ani"
		CursorNames(0) = ".\Resources\Anim Cursors\ADITI CC dragcur01.ani"
		CursorNames(2) = ".\Resources\Static Cursors\cur002.cur"
		CursorNames(3) = ".\Resources\Static Cursors\cur003.cur"
		CursorNames(4) = ".\Resources\Static Cursors\cur004.cur"
		CursorNames(5) = ".\Resources\Static Cursors\cur005.cur"
		CursorNames(6) = ".\Resources\Static Cursors\dragcur002.cur"
		CursorNames(7) = ".\Resources\Static Cursors\dragcur003.cur"
		CursorNames(8) = ".\Resources\Static Cursors\dragcur004.cur"
		CursorNames(9) = ".\Resources\Static Cursors\dragcur005.cur"
		CursorNames(10) = Cursor.Current.ToString

	End Sub

	Public Sub SetCursor(ByVal cIndex As Integer, ByVal dragAction As Boolean)

		InitCursorValues(CursorNames)

		If cIndex = 1 Then
			If dragAction = True Then
				SetSystemCursor(LoadCursorFromFile(CursorNames(0)), OCR_NORMAL)
			Else
				SetSystemCursor(LoadCursorFromFile(CursorNames(cIndex)), OCR_NORMAL)
			End If
		Else
			If dragAction = False Then
				Select Case cIndex
					Case 2
						SetSystemCursor(LoadCursorFromFile(CursorNames(cIndex)), OCR_NORMAL)
					Case 3
						SetSystemCursor(LoadCursorFromFile(CursorNames(cIndex)), OCR_NORMAL)
					Case 4
						SetSystemCursor(LoadCursorFromFile(CursorNames(cIndex)), OCR_NORMAL)
					Case 5
						SetSystemCursor(LoadCursorFromFile(CursorNames(cIndex)), OCR_NORMAL)
				End Select
			ElseIf dragAction = True Then
				Select Case cIndex
					Case 2
						SetSystemCursor(LoadCursorFromFile(CursorNames(cIndex + 4)), OCR_NORMAL)
					Case 3
						SetSystemCursor(LoadCursorFromFile(CursorNames(cIndex + 4)), OCR_NORMAL)
					Case 4
						SetSystemCursor(LoadCursorFromFile(CursorNames(cIndex + 4)), OCR_NORMAL)
					Case 5
						SetSystemCursor(LoadCursorFromFile(CursorNames(cIndex + 4)), OCR_NORMAL)
				End Select
			End If
		End If

	End Sub

End Class
