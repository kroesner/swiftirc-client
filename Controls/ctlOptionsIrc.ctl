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
    UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    updateColours Controls
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Sub initControls()
    m_labelManager.addLabel "Highlights", ltHeading, 10, 10
    m_labelManager.addLabel "Logging", ltHeading, 260, 10
    
    Set m_checkEnableHighlighting = addCheckBox(Controls, "Enable highlighting", 10, 30, 200, 20)
    Set m_checkHighlightNickname = addCheckBox(Controls, "Highlight current nickname", 10, 50, 200, 20)
    
    Set m_checkEnableFiltering = addCheckBox(Controls, "Enable message filtering", 10, 215, 200, 20)
    Set m_checkAutoRejoinOnKick = addCheckBox(Controls, "Auto rejoin on kick", 10, 235, 200, 20)
    
    Set m_checkTimestamps = addCheckBox(Controls, "Enable timestamps", 10, 260, 200, 20)
    Set m_fieldTimestampFormat = addField(Controls, "Timestamp format:", 10, 285, 210, 20)
    m_fieldTimestampFormat.setFieldWidth 130, 80
    
    Set m_buttonHighlightAdd = addButton(Controls, "Add", 165, 80, 70, 20)
    Set m_buttonHighlightDel = addButton(Controls, "Remove", 165, 104, 70, 20)
    
    Set m_buttonHighlightClear = addButton(Controls, "Clear", 165, 128, 70, 19)
    
    Set m_listHighlights = createControl(Controls, "VB.ListBox", "highlights")
    
    Set m_checkEnableHighlightSound = addCheckBox(Controls, "Custom highlight sound", 10, 155, 200, 20)
    Set m_fieldHighlightSoundPath = addField(Controls, "", 10, 177, 150, 20)
    m_fieldHighlightSoundPath.setFieldWidth 1, 149
    
    Set m_buttonHighlightCustom = addButton(Controls, "Browse", 165, 177, 70, 20)
    
    
    Set m_checkEnableLogging = addCheckBox(Controls, "Enable logging", 260, 30, 200, 20)
    
    Set m_checkLogStatus = addCheckBox(Controls, "Log status window", 260, 60, 140, 20)
    Set m_checkLogChannel = addCheckBox(Controls, "Log channels", 400, 60, 140, 20)
    Set m_checkLogQuery = addCheckBox(Controls, "Log queries/PMs", 260, 80, 140, 20)
    Set m_checkLogGeneric = addCheckBox(Controls, "Log other windows", 400, 80, 150, 20)
    
    Set m_checkLogDirectories = addCheckBox(Controls, "Store logs in seperate folders", 260, 110, 200, 20)
    Set m_checkLogDirectoriesProfile = addCheckBox(Controls, "Use profile names for folders", 260, 130, 200, 20)
    
    Set m_checkLogIncludeCodes = addCheckBox(Controls, "Include formatting codes", 260, 160, 200, 20)
    Set m_checkLogIncludeTimeStamp = addCheckBox(Controls, "Include timestamps", 260, 180, 200, 20)
    
    Set m_buttonViewLogs = addButton(Controls, "View logs", 490, 30, 70, 20)
    
    m_listHighlights.Move 10, 80, 150, 73
    
    'm_fieldHighlightSoundPath.enabled = False
End Sub

Private Sub m_buttonHighlightCustom_clicked()
    Dim highlightFilename As String
    
    CommonDialog.Filter = "Sound file (*.wav;*.mp3)|*.wav;*.mp3"
    CommonDialog.DefaultExt = "txt"
    CommonDialog.DialogTitle = "Select File"
    CommonDialog.ShowOpen
    
    If Not CommonDialog.CancelError Then
        m_fieldHighlightSoundPath.value = CommonDialog.fileName
    End If
End Sub

Private Sub m_buttonHighlightAdd_clicked()
    Dim result As Variant
    
    result = requestInput("Highlighting", "Highlight text:", vbNullString, Me)
    
    If result = False Then
        Exit Sub
    End If

    Dim highlight As New CHighlight
    
    highlight.text = CStr(result)

    m_highlights.Add highlight
    m_listHighlights.addItem highlight.text
End Sub

Private Sub m_buttonHighlightClear_clicked()
    m_listHighlights.clear
    m_highlights.clear
End Sub

Private Sub m_buttonHighlightDel_clicked()
    If m_listHighlights.ListIndex = -1 Then
        Exit Sub
    End If
    
    Dim index As Long
    
    index = m_listHighlights.ListIndex
    
    m_highlights.Remove m_listHighlights.ListIndex + 1
    m_listHighlights.removeItem m_listHighlights.ListIndex
    
    If index = 0 Then
        If m_listHighlights.ListCount > 0 Then
            m_listHighlights.ListIndex = 0
        End If
    Else
        m_listHighlights.ListIndex = index - 1
    End If
End Sub

Private Sub m_buttonViewLogs_clicked()
    ShellExecute UserControl.hwnd, "explore", g_userPath & LOG_DIR, 0, 0, SW_SHOW
End Sub

Private Sub m_checkEnableLogging_Click()
    If m_checkEnableLogging.value = 0 Then
        m_checkLogStatus.enabled = False
        m_checkLogChannel.enabled = False
        m_checkLogQuery.enabled = False
        m_checkLogGeneric.enabled = False
        m_checkLogDirectories.enabled = False
        m_checkLogDirectoriesProfile.enabled = False
        m_checkLogIncludeCodes.enabled = False
        m_checkLogIncludeTimeStamp.enabled = False
    Else
        m_checkLogStatus.enabled = True
        m_checkLogChannel.enabled = True
        m_checkLogQuery.enabled = True
        m_checkLogGeneric.enabled = True
        m_checkLogDirectories.enabled = True
        m_checkLogDirectoriesProfile.enabled = True
        m_checkLogIncludeCodes.enabled = True
        m_checkLogIncludeTimeStamp.enabled = True
    End If
End Sub

Private Sub UserControl_Initialize()
    initControls
    
    Dim count As Long
    
    m_checkEnableHighlighting.value = -settings.enableHighlighting
    m_checkHighlightNickname.value = -settings.highlightNickname

    Dim highlight As CHighlight
    
    For count = 1 To highlights.highlightCount
        Set highlight = New CHighlight
        highlights.highlightItem(count).copy highlight
        m_highlights.Add highlight
        m_listHighlights.addItem highlight.text
    Next count
    
    m_checkEnableFiltering.value = -settings.enableFiltering
    m_checkAutoRejoinOnKick.value = -settings.autoRejoinOnKick
    
    m_checkTimestamps.value = -settings.setting("timestamps", estBoolean)
    m_fieldTimestampFormat.value = settings.setting("timestampFormat", estString)
    
    m_checkEnableLogging.value = -settings.setting("enableLogging", estBoolean)
    m_checkLogStatus.value = -settings.setting("logStatus", estBoolean)
    m_checkLogChannel.value = -settings.setting("logChannel", estBoolean)
    m_checkLogQuery.value = -settings.setting("logQuery", estBoolean)
    m_checkLogGeneric.value = -settings.setting("logGeneric", estBoolean)
    m_checkLogDirectories.value = -settings.setting("logDirectories", estBoolean)
    m_checkLogDirectoriesProfile.value = -settings.setting("logDirectoriesProfile", estBoolean)
    m_checkLogIncludeCodes.value = -settings.setting("logIncludeCodes", estBoolean)
    m_checkLogIncludeTimeStamp.value = -settings.setting("logIncludeTimestamP", estBoolean)
    m_checkEnableHighlightSound.value = -settings.setting("highlightCustomSound", estBoolean)
    m_fieldHighlightSoundPath.value = settings.setting("highlightSoundPath", estString)
    
    If m_checkEnableLogging.value = 0 Then
        m_checkLogStatus.enabled = False
        m_checkLogChannel.enabled = False
        m_checkLogQuery.enabled = False
        m_checkLogGeneric.enabled = False
        m_checkLogDirectories.enabled = False
        m_checkLogDirectoriesProfile.enabled = False
        m_checkLogIncludeCodes.enabled = False
        m_checkLogIncludeTimeStamp.enabled = False
    Else
        m_checkLogStatus.enabled = True
        m_checkLogChannel.enabled = True
        m_checkLogQuery.enabled = True
        m_checkLogGeneric.enabled = True
        m_checkLogDirectories.enabled = True
        m_checkLogDirectoriesProfile.enabled = True
        m_checkLogIncludeCodes.enabled = True
        m_checkLogIncludeTimeStamp.enabled = True
    End If
End Sub

Friend Sub saveSettings()
    highlights.clearHighlights
    
    Dim count As Long
    
    settings.enableHighlighting = -m_checkEnableHighlighting.value
    settings.highlightNickname = -m_checkHighlightNickname.value
    
    For count = 1 To m_highlights.count
        highlights.addHighlightIndirect m_highlights.item(count)
    Next count
    
    settings.enableFiltering = -m_checkEnableFiltering.value
    settings.autoRejoinOnKick = -m_checkAutoRejoinOnKick.value
    
    settings.setting("timestamps", estBoolean) = -m_checkTimestamps.value
    settings.setting("timestampFormat", estString) = m_fieldTimestampFormat.value
    g_timestamps = -m_checkTimestamps.value
    g_timestampFormat = m_fieldTimestampFormat.value
    
    settings.setting("enableLogging", estBoolean) = -m_checkEnableLogging.value
    settings.setting("logStatus", estBoolean) = -m_checkLogStatus.value
    settings.setting("logChannel", estBoolean) = -m_checkLogChannel.value
    settings.setting("logQuery", estBoolean) = -m_checkLogQuery.value
    settings.setting("logGeneric", estBoolean) = -m_checkLogGeneric.value
    settings.setting("logDirectories", estBoolean) = -m_checkLogDirectories.value
    settings.setting("logDirectoriesProfile", estBoolean) = -m_checkLogDirectoriesProfile.value
    settings.setting("logIncludeCodes", estBoolean) = -m_checkLogIncludeCodes.value
    settings.setting("logIncludeTimestamP", estBoolean) = -m_checkLogIncludeTimeStamp.value
    settings.setting("highlightCustomSound", estBoolean) = -m_checkEnableHighlightSound.value
    settings.setting("highlightSoundPath", estString) = m_fieldHighlightSoundPath.value
    
    highlights.save
End Sub

Private Sub UserControl_Paint()
    FillRect UserControl.hdc, makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), _
        colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
        
    FrameRect UserControl.hdc, makeRect(5, 245, 5, UserControl.ScaleHeight - 120), _
        colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
        
    FrameRect UserControl.hdc, makeRect(255, UserControl.ScaleWidth - 5, 5, UserControl.ScaleHeight - 120), _
        colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
        
    m_labelManager.renderLabels UserControl.hdc
End Sub

Private Sub UserControl_Terminate()
    debugLog "ctlOptionsIrc terminating"
End Sub
