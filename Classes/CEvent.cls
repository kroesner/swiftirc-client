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
10        m_eventName = name
20        m_format = format
30        Set m_eventColour = eventColour
40        m_flags = flags
End Sub

Public Sub update(format As String, eventColour As CEventColour, flags As Long)
10        m_format = format
20        Set m_eventColour = eventColour
30        m_flags = flags
End Sub

Public Property Get eventName() As String
10        eventName = m_eventName
End Property

Public Property Get format() As String
10        format = m_format
End Property

Public Property Get eventColour() As CEventColour
10        Set eventColour = m_eventColour
End Property

Public Property Get flags() As Long
10        flags = m_flags
End Property

Public Function hasFlag(flag As Long) As Boolean
10        hasFlag = (m_flags And flag)
End Function

Public Property Get visible() As Boolean
10        visible = m_flags And TVE_VISIBLE
End Property

Public Property Let visible(newValue As Boolean)
10        If newValue Then
20            m_flags = m_flags Or TVE_VISIBLE
30        Else
40            If m_flags And TVE_VISIBLE Then
50                m_flags = m_flags - TVE_VISIBLE
60            End If
70        End If
End Property
