VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_name As String
Private m_session As CSession
Private m_window As ctlWindowQuery
Private m_windowText As ITextWindow

Public Sub init(name As String, window As ctlWindowQuery, session As CSession)
10        m_name = name
20        Set m_window = window
30        Set m_windowText = window
40        Set m_session = session
          
50        If settings.setting("enableLogging", estBoolean) And settings.setting("logQuery", estBoolean) Then
60            m_windowText.textview.logName = m_session.baseLogPath & sanitizeFilename(m_name)
70            m_windowText.textview.enableLogging = True
80        End If
End Sub

Public Sub deInit()
10        m_window.deInit
          
20        Set m_window = Nothing
30        Set m_windowText = Nothing
40        Set m_session = Nothing
End Sub

Public Property Get name() As String
10        name = m_name
End Property

Public Property Get window() As ctlWindowQuery
10        Set window = m_window
End Property
    
Public Sub userMessage(origin As COrigin, text As String, highlight As Boolean)
          Dim count As Long

10        If highlight Then
20            m_windowText.addEventEx "PRIVMSG_HIGHLIGHT", Nothing, origin.nickname, TVE_NONE, makeStringArray(text)
30            m_session.client.switchbar.tabActivity m_window.switchbartab, tasHighlight
40        Else
50            m_windowText.addEventEx "PRIVMSG", Nothing, origin.nickname, TVE_NONE, makeStringArray(text)
60            m_session.client.switchbar.tabActivity m_window.switchbartab, tasAlert
70        End If
End Sub

Public Sub userEmote(origin As COrigin, text As String, highlight As Boolean)
10        If highlight Then
20            m_windowText.addEvent "EMOTE_HIGHLIGHT", makeStringArray(origin.nickname, text)
30            m_session.client.switchbar.tabActivity m_window.switchbartab, tasHighlight
40        Else
50            m_windowText.addEvent "EMOTE", makeStringArray(origin.nickname, text)
60            m_session.client.switchbar.tabActivity m_window.switchbartab, tasAlert
70        End If
End Sub

Public Sub nickChange(origin As COrigin, newNick As String)
10        m_windowText.addEvent "QUERY_NICK_CHANGE", makeStringArray(origin.nickname, newNick)
20        m_session.client.switchbar.tabActivity m_window.switchbartab, tasEvent
30        m_name = newNick
          
40        m_window.switchbartab.caption = newNick
50        m_session.client.redrawSwitchbarTab m_window.switchbartab
End Sub

Public Sub textEntered(text As String)
10        m_session.sendLine "PRIVMSG " & m_name & " :" & text
20        m_windowText.addEventEx "PRIVMSG", Nothing, m_session.currentNickname, TVE_NONE, _
              makeStringArray(text)
End Sub

Public Sub sendEmote(text As String)
10        m_session.sendLine "PRIVMSG " & m_name & " :" & Chr$(1) & "ACTION " & text & Chr$(1)
20        m_windowText.addEvent "EMOTE", makeStringArray(m_session.currentNickname, text)
End Sub

Private Sub Class_Terminate()
10        debugLog "CQuery terminated"
End Sub
