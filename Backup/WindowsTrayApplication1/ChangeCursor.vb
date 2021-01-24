Imports System.String


Public Class ChangeCursor

    Declare Function SetSystemCursor Lib "user32.dll" (ByVal hcur As System.Int32, ByVal id As System.Int32) As System.Int32
    Declare Function LoadCursorFromFile Lib "user32.dll" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As System.Int32

	Private Const OCR_NORMAL As System.Int32 = 32512

    Public CursorNames(10) As String


    Public Shared Sub InitCursorValues(ByRef CursorNames)

        CursorNames(0) = "C:\Users\m9000808\Documents\Visual Studio 2008\Projects\WindowsTrayApplication1\WindowsTrayApplication1\Resources\Anim Cursors\ADITI CC cur01.ani"
		CursorNames(0) = "C:\Users\m9000808\Documents\Visual Studio 2008\Projects\WindowsTrayApplication1\WindowsTrayApplication1\Resources\Anim Cursors\ADITI CC dragcur01.ani"
        CursorNames(2) = "C:\Users\m9000808\Documents\Visual Studio 2008\Projects\WindowsTrayApplication1\WindowsTrayApplication1\Resources\Static Cursors\cur002.cur"
        CursorNames(3) = "C:\Users\m9000808\Documents\Visual Studio 2008\Projects\WindowsTrayApplication1\WindowsTrayApplication1\Resources\Static Cursors\cur003.cur"
        CursorNames(4) = "C:\Users\m9000808\Documents\Visual Studio 2008\Projects\WindowsTrayApplication1\WindowsTrayApplication1\Resources\Static Cursors\cur004.cur"
        CursorNames(5) = "C:\Users\m9000808\Documents\Visual Studio 2008\Projects\WindowsTrayApplication1\WindowsTrayApplication1\Resources\Static Cursors\cur005.cur"
        CursorNames(6) = "C:\Users\m9000808\Documents\Visual Studio 2008\Projects\WindowsTrayApplication1\WindowsTrayApplication1\Resources\Static Cursors\dragcur002.cur"
        CursorNames(7) = "C:\Users\m9000808\Documents\Visual Studio 2008\Projects\WindowsTrayApplication1\WindowsTrayApplication1\Resources\Static Cursors\dragcur003.cur"
        CursorNames(8) = "C:\Users\m9000808\Documents\Visual Studio 2008\Projects\WindowsTrayApplication1\WindowsTrayApplication1\Resources\Static Cursors\dragcur004.cur"
        CursorNames(9) = "C:\Users\m9000808\Documents\Visual Studio 2008\Projects\WindowsTrayApplication1\WindowsTrayApplication1\Resources\Static Cursors\dragcur005.cur"

    End Sub

    Public Sub SetCursor(ByVal cIndex As Integer)

        InitCursorValues(CursorNames)

		if cIndex = 1 then 
			if dragaction = true then
				SetSystemCursor(LoadCursorFromFile(CursorNames(0)), OCR_NORMAL)
			else 
				SetSystemCursor(LoadCursorFromFile(CursorNames(cIndex)), OCR_NORMAL)
			end if
		else
			if dragaction = false then
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
			else if dragaction = true then
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
			end if 
		end if
		
    End Sub

	
	
End Class
