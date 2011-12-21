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
    m_name = name
    Set m_window = window
    Set m_windowText = window
    Set m_session = session
    
    If settings.setting("enableLogging", estBoolean) And settings.setting("logQuery", estBoolean) Then
        m_windowText.textview.logName = m_session.baseLogPath & sanitizeFilename(m_name)
        m_windowText.textview.enableLogging = True
    End If
End Sub

Public Sub deInit()
    m_window.deInit
    
    Set m_window = Nothing
    Set m_windowText = Nothing
    Set m_session = Nothing
End Sub

Public Property Get name() As String
    name = m_name
End Property

Public Property Get window() As ctlWindowQuery
    Set window = m_window
End Property
    
Public Sub userMessage(origin As COrigin, text As String, highlight As Boolean)
    Dim count As Long

    If highlight Then
        m_windowText.addEventEx "PRIVMSG_HIGHLIGHT", Nothing, origin.nickname, TVE_NONE, makeStringArray(text)
        m_session.client.switchbar.tabActivity m_window.switchbartab, tasHighlight
    Else
        m_windowText.addEventEx "PRIVMSG", Nothing, origin.nickname, TVE_NONE, makeStringArray(text)
        m_session.client.switchbar.tabActivity m_window.switchbartab, tasAlert
    End If
End Sub

Public Sub userEmote(origin As COrigin, text As String, highlight As Boolean)
    If highlight Then
        m_windowText.addEvent "EMOTE_HIGHLIGHT", makeStringArray(origin.nickname, text)
        m_session.client.switchbar.tabActivity m_window.switchbartab, tasHighlight
    Else
        m_windowText.addEvent "EMOTE", makeStringArray(origin.nickname, text)
        m_session.client.switchbar.tabActivity m_window.switchbartab, tasAlert
    End If
End Sub

Public Sub nickChange(origin As COrigin, newNick As String)
    m_windowText.addEvent "QUERY_NICK_CHANGE", makeStringArray(origin.nickname, newNick)
    m_session.client.switchbar.tabActivity m_window.switchbartab, tasEvent
    m_name = newNick
    
    m_window.switchbartab.caption = newNick
    m_session.client.redrawSwitchbarTab m_window.switchbartab
End Sub

Public Sub textEntered(text As String)
    m_session.sendLine "PRIVMSG " & m_name & " :" & text
    m_windowText.addEventEx "PRIVMSG", Nothing, m_session.currentNickname, TVE_NONE, _
        makeStringArray(text)
End Sub

Public Sub sendEmote(text As String)
    m_session.sendLine "PRIVMSG " & m_name & " :" & Chr$(1) & "ACTION " & text & Chr$(1)
    m_windowText.addEvent "EMOTE", makeStringArray(m_session.currentNickname, text)
End Sub

Private Sub Class_Terminate()
    debugLog "CQuery terminated"
End Sub
