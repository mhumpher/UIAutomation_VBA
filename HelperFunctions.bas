'INTRODUCTION#############################################################################################
'Set of helper functions using the WinAPI
'Author: Michael Humpherys
'Last Updated: 6/15/2017
'Version: 0.1a
'Requirements:
'
'TO DO: - Mouse Move
'       - Mouse Click
'       - Send Keys
'Resources:
'       http://www.jkp-ads.com/articles/apideclarations.asp
'       http://www.vbforums.com/showthread.php?734167-Mouse-Move-and-Click-with-Windows-API-Function-SendInput
'       https://support.microsoft.com/nl-nl/help/2030490/office-2010-help-files-win32api-ptrsafe-with-64-bit-support
'=======================================================================================================
'TYPE DECLARATIONS#######################################################################################
Type POINTAPI
        x As Long
        y As Long
End Type
'========================================================================================================

'DECLARE FUNCTIONS#######################################################################################
'========================================================================================================
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Declare PtrSafe Function LineTo _
    Lib "gdi32" _
    (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long

'https://msdn.microsoft.com/en-us/library/windows/desktop/dd145069(v=vs.85).aspx
Declare PtrSafe Function MoveToEx _
    Lib "gdi32" _
    (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long

Declare PtrSafe Function GetDC _
    Lib "user32" _
    (ByVal hwnd As LongPtr) As LongPtr

Declare PtrSafe Function CreatePen _
    Lib "gdi32" _
    (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As LongPtr

Declare PtrSafe Function SelectObject _
    Lib "gdi32" _
    (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
'==========================================================================================================

'FUNCTIONS###########################################################################################################
'===========================================================================================================
'Function: drawLine
'Purpose: Draw a red line on desktop
'Parameters:    x0, y0 - coordinates for start of line
'               x1, y1 - coordinates for end of line
'Returns: N/A
'Notes:
'Const PS_SOLID = 0
'Const PS_DASH = 1
'Const PS_DOT = 2
'Const PS_DASHDOT = 3
'Const PS_DASHDOTDOT = 4
'Const PS_NULL = 5
'Const PS_INSIDEFRAME = 6
'http://www.jasinskionline.com/windowsapi/ref/c/createpen.html
'===========================================================================================================
Sub drawLine(x0 As Long, y0 As Long, x1 As Long, y1 As Long)
    Dim hdc As Variant
    Dim pen As Variant
    Dim res As Variant
    Dim lp As POINTAPI

    hdc = GetDC(0)
    'pen = CreatePen(PS_SOLID, 4, RGB(255, 0, 0))
    pen = CreatePen(0, 4, RGB(255, 0, 0))
    res = SelectObject(hdc, pen)
    res = MoveToEx(hdc, x0, y0, lp)
    res = LineTo(hdc, x1, y1)

End Sub

'===========================================================================================================
'Function: drawRect
'Purpose: Draw a red rectangle
'Parameters:    x0, y0 - coordinates for top left of rectangle
'               x1, y1 - coordinates for bottom right of rectangle
'Returns: N/A
'Notes:
'===========================================================================================================
Sub drawRect(x0 As Long, y0 As Long, x1 As Long, y1 As Long)
    drawLine x0, y0, x1, y0 'top
    drawLine x0, y0, x0, y1 'left
    drawLine x1, y0, x1, y1 'right
    drawLine x0, y1, x1, y1 'bottom
End Sub
