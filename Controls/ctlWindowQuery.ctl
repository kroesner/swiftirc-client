VERSION 5.00
Begin VB.UserControl ctlWindowQuery 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "ctlWindowQuery"
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

Private m_textInputHeight As Long
Private m_children As New cArrayList
Private m_nameCounter As Long
Private m_fontManager As CFontManager
Private m_tab As CTab

Private m_session As CSession
Private m_query As CQuery

Private m_realWindow As VBControlExtender

Private WithEvents m_textView As ctlTextView
Attribute m_textView.VB_VarHelpID = -1
Private m_textViewControl As VBControlExtender
Private WithEvents m_textInput As ctlTextInput
Attribute m_textInput.VB_VarHelpID = -1
Private m_textInputControl As VBControlExtender

Private Sub m_textInput_mouseWheel(delta As Long)
10        m_textView.processMouseWheel delta
End Sub

Private Sub m_textInput_pageDown()
10        m_textView.pageDown
End Sub

Private Sub m_textInput_pageUp()
10        m_textView.pageUp
End Sub


Private Property Get ITextWindow_textview() As ctlTextView
10        Set ITextWindow_textview = m_textView
End Property

Private Sub ITextWindow_clear()
10        m_textView.clear
End Sub

Private Sub m_textView_clickedUrl(url As String)
10        If left$(url, 1) = "#" Then
20            If InStr(url, ",0") = 0 Then
30                m_session.joinChannel url
40            End If
50        Else
60            m_session.client.visitUrl url, True
70        End If
End Sub

Private Sub m_textView_noLongerNeedFocus()
10        If getRealWindow(m_textInput).visible Then
20            getRealWindow(m_textInput).setFocus
30        End If
End Sub

Private Sub ITextWindow_focusInput()
10        If getRealWindow(m_textInput).visible Then
20            getRealWindow(m_textInput).setFocus
30        End If
End Sub

Private Sub IColourUser_coloursUpdated()
10        updateColours Controls
End Sub

Public Sub init(session As CSession, query As CQuery)
10        Set m_session = session
20        Set m_query = query
End Sub

Public Sub deInit()
10        Set m_session = Nothing
20        Set m_query = Nothing
30        Set m_realWindow = Nothing
End Sub

Public Property Get switchbartab() As CTab
10        Set switchbartab = m_tab
End Property

Public Property Let switchbartab(newValue As CTab)
10        Set m_tab = newValue
End Property

Public Property Get session() As CSession
10        Set session = m_session
End Property

Public Property Get query() As CQuery
10        Set query = m_query
End Property

Private Property Let IFontUser_fontManager(RHS As CFontManager)
10        Set m_fontManager = RHS
End Property

Private Sub IFontUser_fontsUpdated()
          Dim count As Long
          Dim fontUser As IFontUser
          
10        m_textInputHeight = m_fontManager.fontHeight + 5
          
20        For count = 1 To m_children.count
30            If TypeOf m_children.item(count).object Is IFontUser Then
40                Set fontUser = m_children.item(count).object
50                fontUser.fontManager = m_fontManager
60                fontUser.fontsUpdated
70            End If
80        Next count
          
90        UserControl_Resize
End Sub

Private Property Get ITabWindow_getTab() As CTab
10        Set ITabWindow_getTab = m_tab
End Property

Private Sub ITextWindow_addEvent(eventName As String, params() As String)
10        m_textView.addEvent eventName, params
20        m_session.client.switchbar.tabActivity m_tab, tasEvent
End Sub

Private Sub ITextWindow_addEventEx(eventName As String, userStyle As CUserStyle, username As String, _
    flags As Long, params() As String)
          
10        m_textView.addEventEx eventName, userStyle, username, flags, params
20        m_session.client.switchbar.tabActivity m_tab, tasEvent
End Sub

Private Sub ITextWindow_addText(text As String)
10        m_textView.addRawText "$0", makeStringArray(text)
20        m_session.client.switchbar.tabActivity m_tab, tasEvent
End Sub

Private Sub ITextWindow_addTextEx(eventColour As CEventColour, foreColour As Byte, format As String, _
    userStyle As CUserStyle, username As String, flags As Long, params() As String)
          
10        m_textView.addRawTextEx eventColour, foreColour, format, userStyle, username, flags, params
20        m_session.client.switchbar.tabActivity m_tab, tasEvent
End Sub

Private Property Let ITextWindow_eventManager(RHS As CEventManager)
10        m_textView.eventManager = RHS
End Property

Private Property Let ITextWindow_inputText(RHS As String)
10        m_textInput.text = RHS
End Property

Private Property Get ITextWindow_session() As CSession
10        Set ITextWindow_session = m_session
End Property

Private Sub ITextWindow_Update()
10        m_textView.updateVisibility
End Sub
Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Function createNewWindow(progId As String, name As String) As IWindow
          Dim newControl As VBControlExtender
          
10        Set newControl = Controls.Add(progId, name & m_nameCounter)
          
          Dim window As IWindow
          
20        Set window = newControl.object
30        window.realWindow = newControl
          
40        m_children.Add newControl
50        m_nameCounter = m_nameCounter + 1
          
60        If TypeOf window Is IFontUser Then
70            If Not m_fontManager Is Nothing Then
                  Dim fontUser As IFontUser
                  
80                Set fontUser = window
90                fontUser.fontManager = m_fontManager
100               fontUser.fontsUpdated
110           End If
120       End If
          
130       Set createNewWindow = window
End Function

Private Sub m_textInput_textSubmitted(text As String, ctrl As Boolean)
10        If left(LTrim$(text), 1) = "/" And Not ctrl Then
20            m_session.textInput Me, text
30        Else
40            m_query.textEntered text
50        End If
End Sub

Private Sub UserControl_Initialize()
10        m_textInputHeight = 30

          Dim window As IWindow

20        Set m_textView = createNewWindow("swiftirc.ctlTextView", "textview")
30        Set window = m_textView
          
40        m_textView.foreColour = 1
50        m_textView.backColour = 0
          
60        window.realWindow.visible = True
          
70        Set m_textInput = createNewWindow("swiftirc.ctlTextInput", "textInput")
80        Set window = m_textInput
          
90        window.realWindow.visible = True
          
100       updateColours Controls
End Sub

Private Sub UserControl_Resize()
          Dim window As IWindow
          
10        If Not m_textView Is Nothing Then
20            Set window = m_textView
30            window.realWindow.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - _
                  m_textInputHeight
40        End If
          
50        If Not m_textInput Is Nothing Then
60            Set window = m_textInput
70            window.realWindow.Move 0, UserControl.ScaleHeight - m_textInputHeight, _
                  UserControl.ScaleWidth, m_textInputHeight
80        End If
End Sub

Private Sub m_textInput_tabbed(text As String, start As Long, Length As Long)
10        If left$(text, 1) = "#" Then
20            m_session.channelTabbing m_textInput, text, start, Length
30            Exit Sub
40        End If
End Sub

Private Sub UserControl_Terminate()
10        debugLog "ctlWindowQuery terminated"
End Sub
