Attribute VB_Name = "modMain"
Option Explicit

Public Const ALL_BITS As Long = &H7FFFFFFF

Public eventColours As CEventColours
Public colourManager As CColourManager
Public serverProfiles As CServerProfileManager
Public colourThemes As CColourThemeManager
Public textManager As CTextManager
Public settings As CSettings
Public highlights As CHighlightManager

Public ignoreManager As CIgnoreManager

Public prefixStyles As cArrayList
Public styleNormal As CUserStyle
Public styleMe As CUserStyle
Public styleMeOp As CUserStyle
Public styleMeHalfop As CUserStyle
Public styleMeVoice As CUserStyle

Public imageManager As CImageManager

Public g_handCursor As IPictureDisp

Public g_iconSBStatus As CImage
Public g_iconSBChannel As CImage
Public g_iconSBQuery As CImage
Public g_iconSBGeneric As CImage
Public g_iconSBList As CImage

Public osVersion As OSVERSIONINFO

Public g_timestamps As Boolean
Public g_timestampFormat As String

Public g_canUseUnicode As Boolean
Public g_initialized As Boolean

Public g_debugModeEx As Boolean

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Enum eSwitchbarPosition
    sbpTop
    sbpBottom
End Enum

Public Enum eSwitchbarTabState
    stsNormal
    stsMouseOver
    stsMouseDown
    stsSelected
End Enum

Public Enum eTabActivityState
    tasNormal
    tasEvent
    tasMessage
    tasAlert
    tasHighlight
End Enum

Public Enum eWndType
    ewtStatus
    ewtChannel
    ewtQuery
    ewtGeneric
End Enum

Public Enum eDataEntryMode
    demAdd
    demEdit
End Enum


Public Enum eTabStripItemState
    tisNormal
    tisMouseOver
    tisSelected
End Enum

Public Enum eButtonMode
    bmNormal
    bmTab
End Enum

Public Enum eIalType
    ialAll
    ialHost
    ialIdent
    ialIdentHost
End Enum

Public Enum eChannelModeType
    cmtList
    cmtParam
    cmtSetOnly
    cmtNormal
    cmtUnknown
End Enum

Public Const SWIFTCOLOUR_WINDOW As Integer = 0
Public Const SWIFTCOLOUR_CONTROLBACK As Integer = 1
Public Const SWIFTCOLOUR_CONTROLFORE As Integer = 2
Public Const SWIFTCOLOUR_CONTROLFOREOVER As Integer = 3
Public Const SWIFTCOLOUR_CONTROLBORDER As Integer = 4
Public Const SWIFTCOLOUR_FRAMEBACK As Integer = 5
Public Const SWIFTCOLOUR_FRAMEBORDER As Integer = 6

Public Const SWIFTCOLOUR_TEXTVIEWBACK = 6
Public Const SWIFTCOLOUR_TEXTVIEWFORE = 7

Public Const SWIFTPEN_BORDER As Integer = 0
Public Const SWIFTPEN_FRAMEBACK As Integer = 1
Public Const SWIFTPEN_THICKBORDER As Integer = 2
Public Const SWIFTPEN_FRAMEBORDER As Integer = 3

Public g_clientCount As Long

Public g_AssetPath As String
Public g_userPath As String

Public Const LOG_DIR = "logs\"


Public g_errorShown As Boolean
Public g_hideErrors As Boolean

Public Const g_cryptKey As String = "m_buttonCancel"

Public Function getPaletteEntry(entry As Long) As Long
10        getPaletteEntry = colourThemes.currentTheme.paletteEntry(entry)
End Function

Public Function getSettingsPaletteEntry(entry As Long) As Long
10        getSettingsPaletteEntry = colourThemes.currentSettingsTheme.paletteEntry(entry)
End Function

Public Function getSystemTime() As Long
10        getSystemTime = CLng(DateDiff("s", #1/1/1970#, Now))
End Function

Public Sub handleError(context As String, errorCode As Long, errorText As String, ByVal line As Long, detailText As String)
          Dim errorReport As String
          
10        errorReport = String(10, "-") & vbCrLf & "Error occured on " & formatTime(getSystemTime()) _
              & vbCrLf & "SwiftIRC version: " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf _
              & vbCrLf & "Error: " & errorText & " (code: " & errorCode & ") in " & context & " on line " & line

20        If LenB(detailText) Then
30            errorReport = errorReport & vbCrLf & "Details: " & vbCrLf & vbCrLf & detailText
40        End If
          
50        errorReport = errorReport & vbCrLf & String(10, "-")
60        showError errorReport
End Sub

Public Function boolToText(bool As Boolean) As String
10        If bool Then
20            boolToText = "yes"
30        Else
40            boolToText = "no"
50        End If
End Function

Public Function textToBool(text As String) As Boolean
10        If LCase$(text) = "yes" Then
20            textToBool = True
30        End If
End Function
    
    
Public Sub showError(text As String)
10        If g_hideErrors Then
20            Exit Sub
30        End If
          
40        If g_errorShown Then
50            Exit Sub
60        End If

70        g_errorShown = True
          
          Dim errorWindow As New frmErrorReport
          
80        errorWindow.txtErrorLog.text = text
90        errorWindow.Show vbModal
100       Unload errorWindow
          
110       g_errorShown = False
End Sub

Public Sub Main()
10        osVersion.dwOSVersionInfoSize = Len(osVersion)
20        GetVersionEx osVersion
          
30        If osVersion.dwMajorVersion >= 5 Then
40            g_canUseUnicode = True
50        End If

    #If debugmode Then
              Dim count As Long
              Dim testContainer As frmTestContainer
              
60            Set testContainer = New frmTestContainer
70            testContainer.Show
              
              'Set testContainer = New frmTestContainer
              'testContainer.Show
    #End If
End Sub

Public Function makeStringArray(ParamArray params()) As String()
          Dim newArray() As String
          
10        ReDim newArray(UBound(params))
          
          Dim count As Integer
          
20        For count = 0 To UBound(params)
30            newArray(count) = CStr(params(count))
40        Next count
          
50        makeStringArray = newArray
End Function

Public Function getRealWindow(window As IWindow) As VBControlExtender
10        Set getRealWindow = window.realWindow
End Function

Public Sub openOptions(client As swiftIrc.SwiftIrcClient, Optional session As CSession = Nothing)
          Dim options As New frmOptions
          
10        options.client = client
20        options.Show vbModal, client
30        Unload options
End Sub

Public Function makeRect(ByVal left As Long, ByVal right As Long, ByVal top As Long, ByVal bottom _
    As Long) As RECT
10        makeRect.left = left
20        makeRect.right = right
30        makeRect.top = top
40        makeRect.bottom = bottom
End Function

Public Function timeString(time As Long)
          Dim days As Long
          Dim hours As Long
          Dim minutes As Long
          Dim seconds As Long
          Dim remaining As Long
          
10        days = time \ 86400
20        remaining = time Mod 86400
30        hours = remaining \ 3600
40        remaining = remaining Mod 3600
50        minutes = remaining \ 60
60        remaining = remaining Mod 60
70        seconds = remaining
          
80        timeString = ""

90        If seconds > 0 Or time = 0 Then
100           timeString = seconds & "sec" & IIf(seconds = 1, "", "s")
110       End If
          
120       If minutes > 0 Then
130           timeString = minutes & "min" & IIf(minutes = 1, "", "s") & IIf(timeString = "", "", " ") & timeString
140       End If
          
150       If hours > 0 Then
160           timeString = hours & "hour" & IIf(hours = 1, "", "s") & IIf(timeString = "", "", " ") & timeString
170       End If
          
180       If days > 0 Then
190           timeString = days & "day" & IIf(days = 1, "", "s") & IIf(timeString = "", "", " ") & timeString
200       End If
End Function

Public Function formatTime(ByVal time As Long) As String
          Dim timeZoneInfo As TIME_ZONE_INFORMATION
          Dim aDate As Date
          Dim timeZone As String
          Dim count As Integer
          Dim ret As Long
          
          Dim tzName() As Integer
            
10        ret = GetTimeZoneInformation(timeZoneInfo)

20        aDate = CDate((25569 + (time / 86400)) - ((timeZoneInfo.Bias * 60) / 86400))
          
30        If ret = 0 Or ret = 1 Then
40            tzName = timeZoneInfo.StandardName
50        Else
60            tzName = timeZoneInfo.DaylightName
70        End If
          
80        For count = 0 To UBound(timeZoneInfo.StandardName)
90            If timeZoneInfo.StandardName(count) = 0 Then
100               Exit For
110           End If
              
120           timeZone = timeZone & ChrW$(timeZoneInfo.StandardName(count))
130       Next count
          
140       formatTime = format(aDate, "ddd, d MMM yyyy HH:mm:ss ") & timeZone
End Function




Public Function getFileNameDate() As String
10        getFileNameDate = format(Now, "yyyy-mm-dd")
End Function

Public Sub debugLog(text As String)
    #If debugmode Then
10            Open App.path & "\debug.log" For Append As #1
20                Print #1, text
30            Close #1
    #End If
End Sub

Public Sub debugLogEx(text As String)
10        If g_debugModeEx Then
20            Open g_userPath & "\debugEx.log" For Append As #1
30                Print #1, text
40            Close #1
50        End If
End Sub

Public Function launchDefaultBrowser(url As String)
10        If osVersion.dwMajorVersion >= 5 Then
20            If osVersion.dwMajorVersion = 5 Then
30                If osVersion.dwMinorVersion >= 1 Then
40                    launchDefaultBrowserModern url
50                Else
60                    launchDefaultBrowserLegacy url
70                End If
80            Else
90                launchDefaultBrowserModern url
100           End If
110       Else
120           launchDefaultBrowserLegacy url
130       End If
End Function

Private Function launchDefaultBrowserModern(url As String)
          Dim ret As Long
          Dim key As Long

10        ret = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice", 0, KEY_READ, key)
              
20        If ret <> 0 Then
30            launchDefaultBrowserLegacy url
40            Exit Function
50        End If

          Dim data As String
          Dim dataSize As Long
          
60        data = Space$(1024)
70        dataSize = 1024
          
80        ret = RegQueryValueEx(key, StrPtr(StrConv("Progid", vbFromUnicode)), 0, 0, data, dataSize)
90        RegCloseKey key
          
100       If ret <> 0 Then
110           launchDefaultBrowserLegacy url
120           Exit Function
130       End If
          
          Dim subKey As String
          
140       data = Mid$(data, 1, dataSize - 1)
          
150       ret = RegOpenKeyEx(HKEY_CLASSES_ROOT, data & "\shell\open\command", 0, KEY_READ, key)
          
160       If ret <> 0 Then
170           launchDefaultBrowserLegacy url
180           Exit Function
190       End If
          
200       data = Space$(1024)
210       dataSize = 1024
          
220       ret = RegQueryValueEx(key, 0, 0, 0, data, dataSize)
230       RegCloseKey key
          
240       If ret <> 0 Then
250           launchDefaultBrowserLegacy url
260           Exit Function
270       End If
          
          Dim path As String
          Dim args As String
          
280       data = Mid$(data, 1, dataSize - 1)
          
290       parseBrowserCommand data, path, args, url
          
300       ShellExecute 0, "open", path, args, 0, SW_SHOWNORMAL
End Function

Private Function launchDefaultBrowserLegacy(url As String)
          Dim ret As Long
          Dim key As Long

10        ret = RegOpenKeyEx(HKEY_CLASSES_ROOT, "http\shell\open\command", 0, KEY_READ, key)
          
20        If ret <> 0 Then
30            Exit Function
40        End If
          
          Dim data As String
          Dim dataSize As Long
          
50        data = Space$(1024)
60        dataSize = 1024
          
70        ret = RegQueryValueEx(key, 0, 0, 0, data, dataSize)
80        RegCloseKey key
          
90        If ret <> 0 Then
100           Exit Function
110       End If
          
          Dim path As String
          Dim args As String
          
120       data = Mid$(data, 1, dataSize - 1)

130       parseBrowserCommand data, path, args, url

140       ShellExecute 0, "open", path, args, 0, SW_SHOWNORMAL
End Function

Private Function parseBrowserCommand(ByVal command As String, ByRef path As String, _
    ByRef args As String, url As String)
          
          Dim pos As Long

10        If left$(command, 1) = Chr$(34) Then
20            pos = InStr(2, command, Chr$(34))
30            path = Mid$(command, 2, pos - 2)
40            args = Mid$(command, pos + 1)
              
50            If InStr(args, "%1") Then
60                args = Replace$(args, "%1", url)
70            Else
80                args = args & " " & url
90            End If
100       Else
110           pos = InStr(command, " ")
              
120           If pos > 0 Then
130               path = Mid(command, 1, pos - 1)
140               args = Mid$(command, pos + 1)
                  
150               If InStr(args, "%1") Then
160                   args = Replace$(args, "%1", url)
170               Else
180                   args = args & " " & url
190               End If
200           Else
210               path = command
220               args = url
230           End If
240       End If
End Function
