VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.UserControl ctlOptionsIrc 
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5835
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   389
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   240
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "ctlOptionsIrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow
Implements IColourUser

Private m_realWindow As VBControlExtender
Private m_labelManager As New CLabelManager

Private m_highlights As New cArrayList

Private m_checkEnableHighlighting As VB.CheckBox
Private m_checkHighlightNickname As VB.CheckBox
Private m_listHighlights As VB.ListBox
Private WithEvents m_buttonHighlightAdd As ctlButton
Attribute m_buttonHighlightAdd.VB_VarHelpID = -1
Private WithEvents m_buttonHighlightDel As ctlButton
Attribute m_buttonHighlightDel.VB_VarHelpID = -1
Private WithEvents m_buttonHighlightClear As ctlButton
Attribute m_buttonHighlightClear.VB_VarHelpID = -1
Private m_checkEnableHighlightSound As VB.CheckBox
Attribute m_checkEnableHighlightSound.VB_VarHelpID = -1
Private m_fieldHighlightSoundPath As swiftIrc.ctlField
Private WithEvents m_buttonHighlightCustom As ctlButton
Attribute m_buttonHighlightCustom.VB_VarHelpID = -1

Private WithEvents m_checkEnableLogging As VB.CheckBox
Attribute m_checkEnableLogging.VB_VarHelpID = -1
Private m_checkLogStatus As VB.CheckBox
Private m_checkLogChannel As VB.CheckBox
Private m_checkLogQuery As VB.CheckBox
Private m_checkLogGeneric As VB.CheckBox
Private m_checkLogDirectories As VB.CheckBox
Private m_checkLogDirectoriesProfile As VB.CheckBox

Private m_checkLogIncludeCodes As VB.CheckBox
Private m_checkLogIncludeTimeStamp As VB.CheckBox

Private WithEvents m_buttonViewLogs As swiftIrc.ctlButton
Attribute m_buttonViewLogs.VB_VarHelpID = -1

Private m_checkEnableFiltering As VB.CheckBox
Private m_checkAutoRejoinOnKick As VB.CheckBox
Private m_checkTimestamps As VB.CheckBox
Private m_fieldTimestampFormat As swiftIrc.ctlField

Private Sub IColourUser_coloursUpdated()
10        UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
20        updateColours Controls
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Sub initControls()
10        m_labelManager.addLabel "Highlights", ltHeading, 10, 10
20        m_labelManager.addLabel "Logging", ltHeading, 260, 10
          
30        Set m_checkEnableHighlighting = addCheckBox(Controls, "Enable highlighting", 10, 30, 200, 20)
40        Set m_checkHighlightNickname = addCheckBox(Controls, "Highlight current nickname", 10, 50, 200, 20)
          
50        Set m_checkEnableFiltering = addCheckBox(Controls, "Enable message filtering", 10, 215, 200, 20)
60        Set m_checkAutoRejoinOnKick = addCheckBox(Controls, "Auto rejoin on kick", 10, 235, 200, 20)
          
70        Set m_checkTimestamps = addCheckBox(Controls, "Enable timestamps", 10, 260, 200, 20)
80        Set m_fieldTimestampFormat = addField(Controls, "Timestamp format:", 10, 285, 210, 20)
90        m_fieldTimestampFormat.setFieldWidth 130, 80
          
100       Set m_buttonHighlightAdd = addButton(Controls, "Add", 165, 80, 70, 20)
110       Set m_buttonHighlightDel = addButton(Controls, "Remove", 165, 104, 70, 20)
          
120       Set m_buttonHighlightClear = addButton(Controls, "Clear", 165, 128, 70, 19)
          
130       Set m_listHighlights = createControl(Controls, "VB.ListBox", "highlights")
          
140       Set m_checkEnableHighlightSound = addCheckBox(Controls, "Custom highlight sound", 10, 155, 200, 20)
150       Set m_fieldHighlightSoundPath = addField(Controls, "", 10, 177, 150, 20)
160       m_fieldHighlightSoundPath.setFieldWidth 1, 149
          
170       Set m_buttonHighlightCustom = addButton(Controls, "Browse", 165, 177, 70, 20)
          
          
180       Set m_checkEnableLogging = addCheckBox(Controls, "Enable logging", 260, 30, 200, 20)
          
190       Set m_checkLogStatus = addCheckBox(Controls, "Log status window", 260, 60, 140, 20)
200       Set m_checkLogChannel = addCheckBox(Controls, "Log channels", 400, 60, 140, 20)
210       Set m_checkLogQuery = addCheckBox(Controls, "Log queries/PMs", 260, 80, 140, 20)
220       Set m_checkLogGeneric = addCheckBox(Controls, "Log other windows", 400, 80, 150, 20)
          
230       Set m_checkLogDirectories = addCheckBox(Controls, "Store logs in seperate folders", 260, 110, 200, 20)
240       Set m_checkLogDirectoriesProfile = addCheckBox(Controls, "Use profile names for folders", 260, 130, 200, 20)
          
250       Set m_checkLogIncludeCodes = addCheckBox(Controls, "Include formatting codes", 260, 160, 200, 20)
260       Set m_checkLogIncludeTimeStamp = addCheckBox(Controls, "Include timestamps", 260, 180, 200, 20)
          
270       Set m_buttonViewLogs = addButton(Controls, "View logs", 490, 30, 70, 20)
          
280       m_listHighlights.Move 10, 80, 150, 73
          
          'm_fieldHighlightSoundPath.enabled = False
End Sub

Private Sub m_buttonHighlightCustom_clicked()
          Dim highlightFilename As String
          
10        CommonDialog.Filter = "Sound file (*.wav;*.mp3)|*.wav;*.mp3"
20        CommonDialog.DefaultExt = "txt"
30        CommonDialog.DialogTitle = "Select File"
40        CommonDialog.ShowOpen
          
50        If Not CommonDialog.CancelError Then
60            m_fieldHighlightSoundPath.value = CommonDialog.filename
70        End If
End Sub

Private Sub m_buttonHighlightAdd_clicked()
          Dim result As Variant
          
10        result = requestInput("Highlighting", "Highlight text:", vbNullString, Me)
          
20        If result = False Then
30            Exit Sub
40        End If

          Dim highlight As New CHighlight
          
50        highlight.text = CStr(result)

60        m_highlights.Add highlight
70        m_listHighlights.addItem highlight.text
End Sub

Private Sub m_buttonHighlightClear_clicked()
10        m_listHighlights.clear
20        m_highlights.clear
End Sub

Private Sub m_buttonHighlightDel_clicked()
10        If m_listHighlights.ListIndex = -1 Then
20            Exit Sub
30        End If
          
          Dim index As Long
          
40        index = m_listHighlights.ListIndex
          
50        m_highlights.Remove m_listHighlights.ListIndex + 1
60        m_listHighlights.removeItem m_listHighlights.ListIndex
          
70        If index = 0 Then
80            If m_listHighlights.ListCount > 0 Then
90                m_listHighlights.ListIndex = 0
100           End If
110       Else
120           m_listHighlights.ListIndex = index - 1
130       End If
End Sub

Private Sub m_buttonViewLogs_clicked()
10        ShellExecute UserControl.hwnd, "explore", g_userPath & LOG_DIR, 0, 0, SW_SHOW
End Sub

Private Sub m_checkEnableLogging_Click()
10        If m_checkEnableLogging.value = 0 Then
20            m_checkLogStatus.enabled = False
30            m_checkLogChannel.enabled = False
40            m_checkLogQuery.enabled = False
50            m_checkLogGeneric.enabled = False
60            m_checkLogDirectories.enabled = False
70            m_checkLogDirectoriesProfile.enabled = False
80            m_checkLogIncludeCodes.enabled = False
90            m_checkLogIncludeTimeStamp.enabled = False
100       Else
110           m_checkLogStatus.enabled = True
120           m_checkLogChannel.enabled = True
130           m_checkLogQuery.enabled = True
140           m_checkLogGeneric.enabled = True
150           m_checkLogDirectories.enabled = True
160           m_checkLogDirectoriesProfile.enabled = True
170           m_checkLogIncludeCodes.enabled = True
180           m_checkLogIncludeTimeStamp.enabled = True
190       End If
End Sub

Private Sub UserControl_Initialize()
10        initControls
          
          Dim count As Long
          
20        m_checkEnableHighlighting.value = -settings.enableHighlighting
30        m_checkHighlightNickname.value = -settings.highlightNickname

          Dim highlight As CHighlight
          
40        For count = 1 To highlights.highlightCount
50            Set highlight = New CHighlight
60            highlights.highlightItem(count).copy highlight
70            m_highlights.Add highlight
80            m_listHighlights.addItem highlight.text
90        Next count
          
100       m_checkEnableFiltering.value = -settings.enableFiltering
110       m_checkAutoRejoinOnKick.value = -settings.autoRejoinOnKick
          
120       m_checkTimestamps.value = -settings.setting("timestamps", estBoolean)
130       m_fieldTimestampFormat.value = settings.setting("timestampFormat", estString)
          
140       m_checkEnableLogging.value = -settings.setting("enableLogging", estBoolean)
150       m_checkLogStatus.value = -settings.setting("logStatus", estBoolean)
160       m_checkLogChannel.value = -settings.setting("logChannel", estBoolean)
170       m_checkLogQuery.value = -settings.setting("logQuery", estBoolean)
180       m_checkLogGeneric.value = -settings.setting("logGeneric", estBoolean)
190       m_checkLogDirectories.value = -settings.setting("logDirectories", estBoolean)
200       m_checkLogDirectoriesProfile.value = -settings.setting("logDirectoriesProfile", estBoolean)
210       m_checkLogIncludeCodes.value = -settings.setting("logIncludeCodes", estBoolean)
220       m_checkLogIncludeTimeStamp.value = -settings.setting("logIncludeTimestamP", estBoolean)
230       m_checkEnableHighlightSound.value = -settings.setting("highlightCustomSound", estBoolean)
240       m_fieldHighlightSoundPath.value = settings.setting("highlightSoundPath", estString)
          
250       If m_checkEnableLogging.value = 0 Then
260           m_checkLogStatus.enabled = False
270           m_checkLogChannel.enabled = False
280           m_checkLogQuery.enabled = False
290           m_checkLogGeneric.enabled = False
300           m_checkLogDirectories.enabled = False
310           m_checkLogDirectoriesProfile.enabled = False
320           m_checkLogIncludeCodes.enabled = False
330           m_checkLogIncludeTimeStamp.enabled = False
340       Else
350           m_checkLogStatus.enabled = True
360           m_checkLogChannel.enabled = True
370           m_checkLogQuery.enabled = True
380           m_checkLogGeneric.enabled = True
390           m_checkLogDirectories.enabled = True
400           m_checkLogDirectoriesProfile.enabled = True
410           m_checkLogIncludeCodes.enabled = True
420           m_checkLogIncludeTimeStamp.enabled = True
430       End If
End Sub

Friend Sub saveSettings()
10        highlights.clearHighlights
          
          Dim count As Long
          
20        settings.enableHighlighting = -m_checkEnableHighlighting.value
30        settings.highlightNickname = -m_checkHighlightNickname.value
          
40        For count = 1 To m_highlights.count
50            highlights.addHighlightIndirect m_highlights.item(count)
60        Next count
          
70        settings.enableFiltering = -m_checkEnableFiltering.value
80        settings.autoRejoinOnKick = -m_checkAutoRejoinOnKick.value
          
90        settings.setting("timestamps", estBoolean) = -m_checkTimestamps.value
100       settings.setting("timestampFormat", estString) = m_fieldTimestampFormat.value
110       g_timestamps = -m_checkTimestamps.value
120       g_timestampFormat = m_fieldTimestampFormat.value
          
130       settings.setting("enableLogging", estBoolean) = -m_checkEnableLogging.value
140       settings.setting("logStatus", estBoolean) = -m_checkLogStatus.value
150       settings.setting("logChannel", estBoolean) = -m_checkLogChannel.value
160       settings.setting("logQuery", estBoolean) = -m_checkLogQuery.value
170       settings.setting("logGeneric", estBoolean) = -m_checkLogGeneric.value
180       settings.setting("logDirectories", estBoolean) = -m_checkLogDirectories.value
190       settings.setting("logDirectoriesProfile", estBoolean) = -m_checkLogDirectoriesProfile.value
200       settings.setting("logIncludeCodes", estBoolean) = -m_checkLogIncludeCodes.value
210       settings.setting("logIncludeTimestamP", estBoolean) = -m_checkLogIncludeTimeStamp.value
220       settings.setting("highlightCustomSound", estBoolean) = -m_checkEnableHighlightSound.value
230       settings.setting("highlightSoundPath", estString) = m_fieldHighlightSoundPath.value
          
240       highlights.save
End Sub

Private Sub UserControl_Paint()
10        FillRect UserControl.hdc, makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), _
              colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
              
20        FrameRect UserControl.hdc, makeRect(5, 245, 5, UserControl.ScaleHeight - 120), _
              colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
              
30        FrameRect UserControl.hdc, makeRect(255, UserControl.ScaleWidth - 5, 5, UserControl.ScaleHeight - 120), _
              colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
              
40        m_labelManager.renderLabels UserControl.hdc
End Sub

Private Sub UserControl_Terminate()
10        debugLog "ctlOptionsIrc terminating"
End Sub
