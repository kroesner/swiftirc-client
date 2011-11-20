VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CServerProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_name As String

'Server details
Private m_hostname As String
Private m_port As Long

Private m_serverPassword As String

'Login details
Private m_primaryNickname As String
Private m_backupNickname As String
Private m_realName As String
Private m_nicknamePassword As String

Private m_enablePerform As Boolean
Private m_perform As String

Private m_enableAutoJoin As Boolean
Private m_enableAutoIdentify As Boolean

Private m_enableReconnect As Boolean
Private m_enableConnectRetry As Boolean

Private m_autoJoinChannels As New cArrayList

Public Property Get name() As String
10        name = m_name
End Property

Public Property Let name(newValue As String)
10        m_name = newValue
End Property

Public Property Get hostname() As String
10        hostname = m_hostname
End Property

Public Property Let hostname(newValue As String)
10        m_hostname = newValue
End Property

Public Property Get port() As Long
10        port = m_port
End Property

Public Property Let port(newValue As Long)
10        m_port = newValue
End Property

Public Property Get serverPassword() As String
10        serverPassword = m_serverPassword
End Property

Public Property Let serverPassword(newValue As String)
10        m_serverPassword = newValue
End Property

Public Property Get primaryNickname() As String
10        primaryNickname = m_primaryNickname
End Property

Public Property Let primaryNickname(newValue As String)
10        m_primaryNickname = newValue
End Property

Public Property Get backupNickname() As String
10        backupNickname = m_backupNickname
End Property

Public Property Let backupNickname(newValue As String)
10        m_backupNickname = newValue
End Property

Public Property Get realName() As String
10        realName = m_realName
End Property

Public Property Let realName(newValue As String)
10        m_realName = newValue
End Property

Public Property Get nicknamePassword() As String
10        nicknamePassword = m_nicknamePassword
End Property

Public Property Let nicknamePassword(newValue As String)
10        m_nicknamePassword = newValue
End Property

Public Property Get enablePerform() As Boolean
10        enablePerform = m_enablePerform
End Property

Public Property Let enablePerform(newValue As Boolean)
10        m_enablePerform = newValue
End Property

Public Property Get perform() As String
10        perform = m_perform
End Property

Public Property Let perform(newValue As String)
10        m_perform = newValue
End Property

Public Property Get enableAutoJoin() As Boolean
10        enableAutoJoin = m_enableAutoJoin
End Property

Public Property Let enableAutoJoin(newValue As Boolean)
10        m_enableAutoJoin = newValue
End Property

Public Property Get enableAutoIdentify() As Boolean
10        enableAutoIdentify = m_enableAutoIdentify
End Property

Public Property Let enableAutoIdentify(newValue As Boolean)
10        m_enableAutoIdentify = newValue
End Property

Public Property Get enableReconnect() As Boolean
10        enableReconnect = m_enableReconnect
End Property

Public Property Let enableReconnect(newValue As Boolean)
10        m_enableReconnect = newValue
End Property

Public Property Get enableConnectRetry() As Boolean
10        enableConnectRetry = m_enableConnectRetry
End Property

Public Property Let enableConnectRetry(newValue As Boolean)
10        m_enableConnectRetry = newValue
End Property

Public Sub addAutoJoinChannel(channel As String, key As String)
          Dim newChannel As New CAutoJoinChannel
          
10        newChannel.channel = channel
20        newChannel.key = key
          
30        m_autoJoinChannels.Add newChannel
End Sub

Public Sub removeAutoJoinChannel(index As Long)
10        m_autoJoinChannels.Remove index
End Sub

Public Function autoJoinChannel(index As Long) As CAutoJoinChannel
10        Set autoJoinChannel = m_autoJoinChannels.item(index)
End Function

Public Function autoJoinChannelCount() As Long
10        autoJoinChannelCount = m_autoJoinChannels.count
End Function

Public Sub clearAutoJoinChannels()
10        m_autoJoinChannels.clear
End Sub
