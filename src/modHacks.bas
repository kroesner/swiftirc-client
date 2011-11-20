Attribute VB_Name = "modHacks"
Option Explicit

Public g_hook As Long
Public m_oldWndProc As Long

Public Function listBoxHook(ByVal lHookID As Long, ByVal wParam As Long, ByVal lParam As Long) As _
    Long
          Dim CWP As CWPSTRUCT
          Dim className As String
          Dim Length As Long
          
10        CopyMemory CWP, ByVal lParam, Len(CWP)

20        If CWP.message = WM_CREATE Then
30            className = Space$(200)
40            Length = GetClassName(CWP.hwnd, className, Len(className))
50            className = left$(className, Length)
              
60            If className = "ThunderListBox" Or className = "ThunderRT6ListBox" Then
70                m_oldWndProc = SetWindowLong(CWP.hwnd, GWL_WNDPROC, AddressOf listBoxStyleHook)
80            End If
90        End If
          
100       listBoxHook = CallNextHookEx(g_hook, lHookID, wParam, ByVal lParam)
End Function

Public Function listBoxStyleHook(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal _
    lParam As Long) As Long
          Dim cStruct As CREATESTRUCT
          
10        If Msg = WM_CREATE Then
              Dim style As Long
              
20            style = GetWindowLong(hwnd, GWL_STYLE)
              
30            If style And WS_BORDER Then
40                style = style - WS_BORDER
50            End If
              
60            style = style Or LBS_HASSTRINGS Or LBS_OWNERDRAWFIXED Or LBS_NOINTEGRALHEIGHT Or _
                  LBS_EXTENDEDSEL
              
70            CopyMemory cStruct, ByVal lParam, Len(cStruct)
80            cStruct.style = style
90            CopyMemory ByVal lParam, cStruct, Len(cStruct)
              
100           SetWindowLong hwnd, GWL_STYLE, style
110           SetWindowLong hwnd, GWL_WNDPROC, m_oldWndProc
120       End If
          
130       listBoxStyleHook = CallWindowProc(m_oldWndProc, hwnd, Msg, wParam, lParam)
End Function

