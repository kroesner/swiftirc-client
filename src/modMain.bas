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
    getPaletteEntry = colourThemes.currentTheme.paletteEntry(entry)
End Function

Public Function getSettingsPaletteEntry(entry As Long) As Long
    getSettingsPaletteEntry = colourThemes.currentSettingsTheme.paletteEntry(entry)
End Function

Public Function getSystemTime() As Long
    getSystemTime = CLng(DateDiff("s", #1/1/1970#, Now))
End Function

Public Sub handleError(context As String, errorCode As Long, errorText As String, ByVal line As Long, detailText As String)
    Dim errorReport As String
    
    errorReport = String(10, "-") & vbCrLf & "Error occured on " & formatTime(getSystemTime()) & vbCrLf & "SwiftIRC version: " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & "Error: " & errorText & " (code: " & errorCode & ") in " & context & " on line " & line

    If LenB(detailText) Then
        errorReport = errorReport & vbCrLf & "Details: " & vbCrLf & vbCrLf & detailText
    End If
    
    errorReport = errorReport & vbCrLf & String(10, "-")
    showError errorReport
End Sub

Public Function boolToText(bool As Boolean) As String
    If bool Then
        boolToText = "yes"
    Else
        boolToText = "no"
    End If
End Function

Public Function textToBool(text As String) As Boolean
    If LCase$(text) = "yes" Then
        textToBool = True
    End If
End Function
    
    
Public Sub showError(text As String)
    If g_hideErrors Then
        Exit Sub
    End If
    
    If g_errorShown Then
        Exit Sub
    End If

    g_errorShown = True
    
    Dim errorWindow As New frmErrorReport
    
    errorWindow.txtErrorLog.text = text
    errorWindow.Show vbModal
    Unload errorWindow
    
    g_errorShown = False
End Sub

Public Sub Main()
    osVersion.dwOSVersionInfoSize = Len(osVersion)
    GetVersionEx osVersion
    
    If osVersion.dwMajorVersion >= 5 Then
        g_canUseUnicode = True
    End If

    #If debugmode Then
        Dim count As Long
        Dim testContainer As frmTestContainer
        
        Set testContainer = New frmTestContainer
        
        testContainer.Show
    #End If
End Sub

Public Function makeStringArray(ParamArray params()) As String()
    Dim newArray() As String
    
    ReDim newArray(UBound(params))
    
    Dim count As Integer
    
    For count = 0 To UBound(params)
        newArray(count) = CStr(params(count))
    Next count
    
    makeStringArray = newArray
End Function

Public Function getRealWindow(window As IWindow) As VBControlExtender
    Set getRealWindow = window.realWindow
End Function

Public Sub openOptions(client As swiftIrc.SwiftIrcClient, Optional session As CSession = Nothing)
    Dim options As New frmOptions
    
    options.client = client
    options.Show vbModal, client
    Unload options
End Sub

Public Function makeRect(ByVal left As Long, ByVal right As Long, ByVal top As Long, ByVal bottom As Long) As RECT
    makeRect.left = left
    makeRect.right = right
    makeRect.top = top
    makeRect.bottom = bottom
End Function

Public Function timeString(time As Long)
    Dim days As Long
    Dim hours As Long
    Dim minutes As Long
    Dim seconds As Long
    Dim remaining As Long
    
    days = time \ 86400
    remaining = time Mod 86400
    hours = remaining \ 3600
    remaining = remaining Mod 3600
    minutes = remaining \ 60
    remaining = remaining Mod 60
    seconds = remaining
    
    timeString = ""

    If seconds > 0 Or time = 0 Then
        timeString = seconds & "sec" & IIf(seconds = 1, "", "s")
    End If
    
    If minutes > 0 Then
        timeString = minutes & "min" & IIf(minutes = 1, "", "s") & IIf(timeString = "", "", " ") & timeString
    End If
    
    If hours > 0 Then
        timeString = hours & "hour" & IIf(hours = 1, "", "s") & IIf(timeString = "", "", " ") & timeString
    End If
    
    If days > 0 Then
        timeString = days & "day" & IIf(days = 1, "", "s") & IIf(timeString = "", "", " ") & timeString
    End If
End Function

Public Function formatTime(ByVal time As Long) As String
    Dim timeZoneInfo As TIME_ZONE_INFORMATION
    Dim aDate As Date
    Dim timeZone As String
    Dim count As Integer
    Dim ret As Long
    
    Dim tzName() As Integer
      
    ret = GetTimeZoneInformation(timeZoneInfo)

    aDate = CDate((25569 + (time / 86400)) - ((timeZoneInfo.Bias * 60) / 86400))
    
    If ret = 0 Or ret = 1 Then
        tzName = timeZoneInfo.StandardName
    Else
        tzName = timeZoneInfo.DaylightName
    End If
    
    For count = 0 To UBound(timeZoneInfo.StandardName)
        If timeZoneInfo.StandardName(count) = 0 Then
            Exit For
        End If
        
        timeZone = timeZone & ChrW$(timeZoneInfo.StandardName(count))
    Next count
    
    formatTime = format(aDate, "ddd, d MMM yyyy HH:mm:ss ") & timeZone
End Function

Public Function getFileNameDate() As String
    getFileNameDate = format(Now, "yyyy-mm-dd")
End Function

Public Function launchDefaultBrowser(url As String)
    If osVersion.dwMajorVersion >= 5 Then
        If osVersion.dwMajorVersion = 5 Then
            If osVersion.dwMinorVersion >= 1 Then
                launchDefaultBrowserModern url
            Else
                launchDefaultBrowserLegacy url
            End If
        Else
            launchDefaultBrowserModern url
        End If
    Else
        launchDefaultBrowserLegacy url
    End If
End Function

Private Function launchDefaultBrowserModern(url As String)
    Dim ret As Long
    Dim key As Long

    ret = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice", 0, KEY_READ, key)
        
    If ret <> 0 Then
        launchDefaultBrowserLegacy url
        Exit Function
    End If

    Dim data As String
    Dim dataSize As Long
    
    data = Space$(1024)
    dataSize = 1024
    
    ret = RegQueryValueEx(key, StrPtr(StrConv("Progid", vbFromUnicode)), 0, 0, data, dataSize)
    RegCloseKey key
    
    If ret <> 0 Then
        launchDefaultBrowserLegacy url
        Exit Function
    End If
    
    Dim subKey As String
    
    data = Mid$(data, 1, dataSize - 1)
    
    ret = RegOpenKeyEx(HKEY_CLASSES_ROOT, data & "\shell\open\command", 0, KEY_READ, key)
    
    If ret <> 0 Then
        launchDefaultBrowserLegacy url
        Exit Function
    End If
    
    data = Space$(1024)
    dataSize = 1024
    
    ret = RegQueryValueEx(key, 0, 0, 0, data, dataSize)
    RegCloseKey key
    
    If ret <> 0 Then
        launchDefaultBrowserLegacy url
        Exit Function
    End If
    
    Dim path As String
    Dim args As String
    
    data = Mid$(data, 1, dataSize - 1)
    
    parseBrowserCommand data, path, args, url
    
    ShellExecute 0, "open", path, args, 0, SW_SHOWNORMAL
End Function

Private Function launchDefaultBrowserLegacy(url As String)
    Dim ret As Long
    Dim key As Long

    ret = RegOpenKeyEx(HKEY_CLASSES_ROOT, "http\shell\open\command", 0, KEY_READ, key)
    
    If ret <> 0 Then
        Exit Function
    End If
    
    Dim data As String
    Dim dataSize As Long
    
    data = Space$(1024)
    dataSize = 1024
    
    ret = RegQueryValueEx(key, 0, 0, 0, data, dataSize)
    RegCloseKey key
    
    If ret <> 0 Then
        Exit Function
    End If
    
    Dim path As String
    Dim args As String
    
    data = Mid$(data, 1, dataSize - 1)

    parseBrowserCommand data, path, args, url

    ShellExecute 0, "open", path, args, 0, SW_SHOWNORMAL
End Function

Private Function parseBrowserCommand(ByVal command As String, ByRef path As String, ByRef args As String, url As String)
    
    Dim pos As Long

    If left$(command, 1) = Chr$(34) Then
        pos = InStr(2, command, Chr$(34))
        path = Mid$(command, 2, pos - 2)
        args = Mid$(command, pos + 1)
        
        If InStr(args, "%1") Then
            args = Replace$(args, "%1", url)
        Else
            args = args & " " & url
        End If
    Else
        pos = InStr(command, " ")
        
        If pos > 0 Then
            path = Mid(command, 1, pos - 1)
            args = Mid$(command, pos + 1)
            
            If InStr(args, "%1") Then
                args = Replace$(args, "%1", url)
            Else
                args = args & " " & url
            End If
        Else
            path = command
            args = url
        End If
    End If
End Function
