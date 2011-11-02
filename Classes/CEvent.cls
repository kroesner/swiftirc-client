VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_eventName As String
Private m_format As String
Private m_eventColour As CEventColour
Private m_flags As Long

Public Sub init(name As String, format As String, eventColour As CEventColour, flags As Long)
    m_eventName = name
    m_format = format
    Set m_eventColour = eventColour
    m_flags = flags
End Sub

Public Sub update(format As String, eventColour As CEventColour, flags As Long)
    m_format = format
    Set m_eventColour = eventColour
    m_flags = flags
End Sub

Public Property Get eventName() As String
    eventName = m_eventName
End Property

Public Property Get format() As String
    format = m_format
End Property

Public Property Get eventColour() As CEventColour
    Set eventColour = m_eventColour
End Property

Public Property Get flags() As Long
    flags = m_flags
End Property

Public Function hasFlag(flag As Long) As Boolean
    hasFlag = (m_flags And flag)
End Function

Public Property Get visible() As Boolean
    visible = m_flags And TVE_VISIBLE
End Property

Public Property Let visible(newValue As Boolean)
    If newValue Then
        m_flags = m_flags Or TVE_VISIBLE
    Else
        If m_flags And TVE_VISIBLE Then
            m_flags = m_flags - TVE_VISIBLE
        End If
    End If
End Property
