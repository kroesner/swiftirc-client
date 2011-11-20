VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTextTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_name As String
Private m_events As New cArrayList

Public Property Get name() As String
10        name = m_name
End Property

Public Property Get eventCount() As String
10        eventCount = m_events.count
End Property

Public Property Get getEvent(index As Long) As CEvent
10        Set getEvent = m_events.item(index)
End Property

Public Sub addEvent(name As String, format As String, eventColour As CEventColour, flags As Long)
          Dim newEvent As New CEvent
          
10        newEvent.init name, format, eventColour, flags
20        m_events.Add newEvent
End Sub
