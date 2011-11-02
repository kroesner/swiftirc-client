Attribute VB_Name = "modWin32Methods"
Option Explicit

'GDI functions
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long) As Long

Public Declare Function GetCurrentObject Lib "gdi32" (ByVal hdc As Long, ByVal uObjectType As Long) _
    As Long

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As _
    Long) As Long

Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal _
    X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
    
Public Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As _
    Long) As Long
    
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As _
    Long

Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal _
    crColor As Long) As Long
    
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As _
    Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As _
    Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Public Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As _
    Long) As Long
    
Public Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush _
    As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    
Public Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As _
    Long) As Long
    
Public Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc _
    As Long, ByVal dwRop As Long) As Long
    
Public Declare Function TransparentBlt Lib "Msimg32" (ByVal hDestDC As Long, ByVal x As Long, ByVal _
    y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) _
    As Long
    
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, _
    lpMetrics As TEXTMETRIC) As Long

'Text functions
Public Declare Function GetTextExtentPoint32A Lib "gdi32" (ByVal hdc As Long, ByVal lpsz As String, _
    ByVal cbString As Long, lpSize As SIZE) As Long
    
Public Declare Function GetTextExtentPoint32W Lib "gdi32" (ByVal hdc As Long, ByVal lpsz As Long, _
    ByVal cbString As Long, lpSize As SIZE) As Long
    
Public Declare Function GetTextExtentExPointA Lib "gdi32" (ByVal hdc As Long, ByVal lpszStr As _
    String, ByVal cchString As Long, ByVal nMaxExtent As Long, lpnFit As Long, alpDx As Long, lpSize As _
    SIZE) As Long
    
Public Declare Function GetTextExtentExPointW Lib "gdi32" (ByVal hdc As Long, ByVal lpszStr As Long, _
    ByVal cchString As Long, ByVal nMaxExtent As Long, lpnFit As Long, alpDx As Long, lpSize As SIZE) _
    As Long
    
Public Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal _
    nCount As Long, ByVal lpRect As Long, ByVal wFormat As Long) As Long

Public Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal _
    nCount As Long, ByVal lpRect As Long, ByVal wFormat As Long) As Long
    
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, _
    ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
    
'Window functions
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As _
    Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As _
    Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal _
    hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As _
    Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

'Scrollbar functions
Public Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal _
    bShow As Long) As Long
    
Public Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, _
    lpScrollInfo As SCROLLINFO) As Long
    
Public Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, _
    lpcScrollInfo As SCROLLINFO, ByVal reDraw As Long) As Long
    
Public Declare Function EnableScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wSBflags As Long, _
    ByVal wArrows As Long) As Long

'Image functions
Public Declare Function LoadImageA Lib "user32" (ByVal hInst As Long, ByVal lpsz _
    As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    
Public Declare Function LoadImageAPtr Lib "user32" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpsz _
    As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    
Public Declare Function LoadImageW Lib "user32" (ByVal hInst As Long, ByVal lpsz _
    As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc _
    As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    
Public Declare Function AlphaBlend Lib "MSIMG32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As _
    Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc _
    As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal _
    nHeightSrc As Long, ByVal lBlendFunction As Long) As Long

'Mouse capturing
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal _
    nCount As Long, lpObject As Any) As Long

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As _
    OSVERSIONINFO) As Long

Public Declare Function GetBitmapDimensionEx Lib "gdi32" (ByVal hBitmap As Long, lpDimension As _
    SIZE) As Long

Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal _
    uElapse As Long, ByVal lpTimerFunc As Long) As Long
    
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As _
    TIME_ZONE_INFORMATION) As Long

'Hooking
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As _
    Any, ByVal cbCopy As Long)
    
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As _
    Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, _
    ByVal wParam As Long, lParam As Any) As Long
    
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As _
    Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal _
    lpClassName As String, ByVal nMaxCount As Long) As Long

Public Declare Function ExtTextOutA Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As _
    Long, ByVal wOptions As Long, ByVal lpRect As Long, ByVal lpString As String, ByVal nCount As Long, _
    lpDx As Long) As Long
    
Public Declare Function ExtTextOutW Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As _
    Long, ByVal wOptions As Long, ByVal lpRect As Long, ByVal lpString As Long, ByVal nCount As Long, _
    lpDx As Long) As Long

Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As _
    LOGFONT) As Long
    
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, _
    ByVal nDenominator As Long) As Long

Public Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal _
    hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal _
    wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, _
    ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal _
    dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As _
    Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal _
    nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
    
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal _
    nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
    
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Public Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As _
    LARGE_INTEGER) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
    ByVal hIcon As Long) As Long
    
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove _
    As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Public Declare Function SetFilePointerEx Lib "kernel32" (ByVal hFile As Long, ByVal _
    liDistanceToMove1 As Long, ByVal liDistanceToMove2 As Long, ByVal lpNewFilePointer As Long, ByVal _
    dwMoveMethod As Long) As Long
    
Public Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As _
    tChooseFont) As Long

Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColorStruct) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal _
    lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
    ByVal lpSubKey As String, ByVal Reserved As Long, ByVal samDesired As Long, phkResult As Long) As _
    Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As _
    Long, ByVal lpValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, ByVal lpData As _
    String, ByRef lpcbData As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Declare Function RegOpenCurrentUser Lib "advapi32.dll" (ByVal samDesired As Long, ByRef _
    phkResult As Long) As Long

Public Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal _
    lpCaption As Long, ByVal uType As Long) As Long

Public Function swiftGetTextExtentPoint32(ByVal hdc As Long, ByVal lpsz As String, lpSize As SIZE)
    If g_canUseUnicode Then
        swiftGetTextExtentPoint32 = GetTextExtentPoint32W(hdc, StrPtr(lpsz), Len(lpsz), lpSize)
    Else
        swiftGetTextExtentPoint32 = GetTextExtentPoint32A(hdc, lpsz, Len(lpsz), lpSize)
    End If
End Function

Public Function swiftGetTextExtentExPoint(ByVal hdc As Long, ByVal lpszStr As String, _
    ByVal nMaxExtent As Long, lpnFit As Long, lpSize As SIZE) As Long

    If g_canUseUnicode Then
        swiftGetTextExtentExPoint = GetTextExtentExPointW(hdc, StrPtr(lpszStr), Len(lpszStr), nMaxExtent, lpnFit, ByVal 0, lpSize)
    Else
        swiftGetTextExtentExPoint = GetTextExtentExPointA(hdc, lpszStr, Len(lpszStr), nMaxExtent, lpnFit, ByVal 0, lpSize)
    End If
End Function
    
Public Function swiftTextOut(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As _
    Long, ByVal lpRect As Long, ByVal lpString As String) As Long

    If g_canUseUnicode Then
        swiftTextOut = ExtTextOutW(hdc, x, y, wOptions, lpRect, StrPtr(lpString), Len(lpString), ByVal 0)
    Else
        swiftTextOut = ExtTextOutA(hdc, x, y, wOptions, lpRect, lpString, Len(lpString), ByVal 0)
    End If
End Function

Public Function swiftDrawText(ByVal hdc As Long, ByVal lpStr As _
    String, ByVal lpRect As Long, ByVal wFormat As Long) As Long
    
    If g_canUseUnicode Then
        swiftDrawText = DrawTextW(hdc, StrPtr(lpStr), -1, lpRect, wFormat)
    Else
        swiftDrawText = DrawTextA(hdc, lpStr, -1, lpRect, wFormat)
    End If
End Function

Public Function swiftLoadImage(ByVal hInst As Long, ByVal lpsz _
    As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    
    If g_canUseUnicode Then
        swiftLoadImage = LoadImageW(hInst, StrPtr(lpsz), un1, n1, n2, un2)
    Else
        swiftLoadImage = LoadImageA(hInst, lpsz, un1, n1, n2, un2)
    End If
End Function

Public Function LoWord(DWord As Long) As Integer
    If DWord And &H8000& Then ' &H8000& = &H00008000
        LoWord = DWord Or &HFFFF0000
    Else
        LoWord = DWord And &HFFFF&
    End If
End Function

Public Function HiWord(DWord As Long) As Integer
    HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Public Function MakeLong(wLow As Long, wHigh As Long) As Long
    MakeLong = LoWord(wLow) Or (&H10000 * LoWord(wHigh))
End Function

Public Function MakeWord(msb As Byte, lsb As Byte) As Long
    MakeWord = CLng("&h" & Trim$(Hex(msb)) & Trim$(Hex(lsb)))
End Function

Public Function DivDiv(lOne As Long, lTwo As Long, lThree As Long) As Long
    Dim lTemp As Single
    
    lTemp = lTwo / lThree
    lTemp = lOne / lTemp
    
    DivDiv = CLng(lTemp)
End Function
