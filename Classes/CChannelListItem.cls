VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CChannelListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_name As String
Private m_userCount As Long
Private m_topic As String

Public Property Get name() As String
    name = m_name
End Property

Public Property Let name(newValue As String)
    m_name = newValue
End Property

Public Property Get userCount() As Long
    userCount = m_userCount
End Property

Public Property Let userCount(newValue As Long)
    m_userCount = newValue
End Property

Public Property Get topic() As String
    topic = m_topic
End Property

Public Property Let topic(newValue As String)
    m_topic = newValue
End Property

