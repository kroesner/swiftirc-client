Attribute VB_Name = "modHacks"
Option Explicit

Public g_hook As Long
Public m_oldWndProc As Long

Public Function listBoxHook(ByVal lHookID As Long, ByVal wParam As Long, ByVal lParam As Long) As _
    Long
    Dim CWP As CWPSTRUCT
    Dim className As String
    Dim Length As Long
    
    CopyMemory CWP, ByVal lParam, Len(CWP)

    If CWP.message = WM_CREATE Then
        className = Space$(200)
        Length = GetClassName(CWP.hwnd, className, Len(className))
        className = left$(className, Length)
        
        If className = "ThunderListBox" Or className = "ThunderRT6ListBox" Then
            m_oldWndProc = SetWindowLong(CWP.hwnd, GWL_WNDPROC, AddressOf listBoxStyleHook)
        End If
    End If
    
    listBoxHook = CallNextHookEx(g_hook, lHookID, wParam, ByVal lParam)
End Function

Public Function listBoxStyleHook(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal _
    lParam As Long) As Long
    Dim cStruct As CREATESTRUCT
    
    If Msg = WM_CREATE Then
        Dim style As Long
        
        style = GetWindowLong(hwnd, GWL_STYLE)
        
        If style And WS_BORDER Then
            style = style - WS_BORDER
        End If
        
        style = style Or LBS_HASSTRINGS Or LBS_OWNERDRAWFIXED Or LBS_NOINTEGRALHEIGHT Or _
            LBS_EXTENDEDSEL
        
        CopyMemory cStruct, ByVal lParam, Len(cStruct)
        cStruct.style = style
        CopyMemory ByVal lParam, cStruct, Len(cStruct)
        
        SetWindowLong hwnd, GWL_STYLE, style
        SetWindowLong hwnd, GWL_WNDPROC, m_oldWndProc
    End If
    
    listBoxStyleHook = CallWindowProc(m_oldWndProc, hwnd, Msg, wParam, lParam)
End Function

