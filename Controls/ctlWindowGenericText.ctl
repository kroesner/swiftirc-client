VERSION 5.00
Begin VB.UserControl ctlWindowGenericText 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "ctlWindowGenericText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow
Implements ITextWindow
Implements ITabWindow
Implements IFontUser
Implements IColourUser
Implements IOptionsChangeListener

Private m_textInputHeight As Long
Private m_children As New cArrayList
Private m_nameCounter As Long
Private m_fontmanager As CFontManager
Private m_tab As CTab

Private m_name As String

Private m_session As CSession
Private m_realWindow As VBControlExtender

Private WithEvents m_textView As ctlTextView
Attribute m_textView.VB_VarHelpID = -1
Private m_textViewControl As VBControlExtender
Private WithEvents m_textInput As ctlTextInput
Attribute m_textInput.VB_VarHelpID = -1
Private m_textInputControl As VBControlExtender

Private Sub m_textInput_pageDown()
    m_textView.pageDown
End Sub

Private Sub m_textInput_pageUp()
    m_textView.pageUp
End Sub

Private Property Get ITextWindow_textview() As ctlTextView
    Set ITextWindow_textview = m_textView
End Property

Private Sub ITextWindow_clear()
    m_textView.clear
End Sub

Private Sub m_textInput_mouseWheel(delta As Long)
    m_textView.processMouseWheel delta
End Sub

Private Sub m_textView_clickedUrl(url As String)
    If left$(url, 1) = "#" Then
        If InStr(url, ",0") = 0 Then
            m_session.joinChannel url
        End If
    Else
        m_session.client.visitUrl url, True
    End If
End Sub

Private Sub m_textView_noLongerNeedFocus()
    If getRealWindow(m_textInput).visible Then
        getRealWindow(m_textInput).setFocus
    End If
End Sub

Private Sub ITextWindow_focusInput()
    If getRealWindow(m_textInput).visible Then
        getRealWindow(m_textInput).setFocus
    End If
End Sub

Private Sub IColourUser_coloursUpdated()
    updateColours Controls
End Sub

Public Sub init(session As CSession, name As String)
    Set m_session = session
    m_name = name
        
    refreshLogging
    
    registerForOptionsChanges Me
End Sub

Public Sub deInit()
    Set m_session = Nothing
    Set m_realWindow = Nothing
    
    unregisterForOptionsChanges Me
End Sub

Public Property Get switchbartab() As CTab
    Set switchbartab = m_tab
End Property

Public Property Let switchbartab(newValue As CTab)
    Set m_tab = newValue
End Property

Public Property Get session() As CSession
    Set session = m_session
End Property

Private Property Let IFontUser_fontManager(RHS As CFontManager)
    Set m_fontmanager = RHS
End Property

Private Sub IFontUser_fontsUpdated()
    Dim count As Long
    Dim fontUser As IFontUser
    
    m_textInputHeight = m_fontmanager.fontHeight + 5
    
    For count = 1 To m_children.count
        If TypeOf m_children.item(count).object Is IFontUser Then
            Set fontUser = m_children.item(count).object
            fontUser.fontManager = m_fontmanager
            fontUser.fontsUpdated
        End If
    Next count
    
    UserControl_Resize
End Sub

Private Property Get ITabWindow_getTab() As CTab
    Set ITabWindow_getTab = m_tab
End Property

Private Sub ITextWindow_addEvent(eventName As String, params() As String)
    m_textView.addEvent eventName, params
    m_session.client.switchbar.tabActivity m_tab, tasEvent
End Sub

Private Sub ITextWindow_addEventEx(eventName As String, userStyle As CUserStyle, username As String, _
    flags As Long, params() As String)
    
    m_textView.addEventEx eventName, userStyle, username, flags, params
    m_session.client.switchbar.tabActivity m_tab, tasEvent
End Sub

Private Sub ITextWindow_addText(text As String)
    m_textView.addRawText "$0", makeStringArray(text)
    m_session.client.switchbar.tabActivity m_tab, tasEvent
End Sub

Private Sub ITextWindow_addTextEx(eventColour As CEventColour, foreColour As Byte, format As String, _
    userStyle As CUserStyle, username As String, flags As Long, params() As String)
    
    m_textView.addRawTextEx eventColour, foreColour, format, userStyle, username, flags, params
    m_session.client.switchbar.tabActivity m_tab, tasEvent
End Sub

Private Property Let ITextWindow_eventManager(RHS As CEventManager)
    m_textView.eventManager = RHS
End Property

Private Property Let ITextWindow_inputText(RHS As String)
    m_textInput.text = RHS
End Property

Private Property Get ITextWindow_session() As CSession
    Set ITextWindow_session = m_session
End Property

Private Sub ITextWindow_Update()
    m_textView.updateVisibility
End Sub
Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Function createNewWindow(progId As String, name As String) As IWindow
    Dim newControl As VBControlExtender
    
    Set newControl = Controls.Add(progId, name & m_nameCounter)
    
    Dim window As IWindow
    
    Set window = newControl.object
    window.realWindow = newControl
    
    m_children.Add newControl
    m_nameCounter = m_nameCounter + 1
    
    If TypeOf window Is IFontUser Then
        If Not m_fontmanager Is Nothing Then
            Dim fontUser As IFontUser
            
            Set fontUser = window
            fontUser.fontManager = m_fontmanager
            fontUser.fontsUpdated
        End If
    End If
    
    Set createNewWindow = window
End Function

Private Sub m_textInput_textSubmitted(text As String, ctrl As Boolean)
    If Not ctrl Then
        m_session.textInput Me, text
    End If
End Sub

Private Sub UserControl_Initialize()
    m_textInputHeight = 30

    Dim window As IWindow

    Set m_textView = createNewWindow("swiftirc.ctlTextView", "textview")
    Set window = m_textView
    
    m_textView.foreColour = 1
    m_textView.backColour = 0
    
    window.realWindow.visible = True
    
    Set m_textInput = createNewWindow("swiftirc.ctlTextInput", "textInput")
    Set window = m_textInput
    
    window.realWindow.visible = True
End Sub

Private Sub UserControl_Resize()
    Dim window As IWindow
    
    If Not m_textView Is Nothing Then
        Set window = m_textView
        window.realWindow.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - _
            m_textInputHeight
    End If
    
    If Not m_textInput Is Nothing Then
        Set window = m_textInput
        window.realWindow.Move 0, UserControl.ScaleHeight - m_textInputHeight, _
            UserControl.ScaleWidth, m_textInputHeight
    End If
End Sub

Private Sub IOptionsChangeListener_optionsChanged()
    refreshLogging
End Sub

Private Function refreshLogging()
    m_textView.logName = combinePath(m_session.baseLogPath, sanitizeFilename(m_name))
    m_textView.enableLogging = shouldLog
End Function

Private Function shouldLog() As Boolean
    shouldLog = settings.setting("enableLogging", estBoolean) And settings.setting("logGeneric", estBoolean)
End Function


