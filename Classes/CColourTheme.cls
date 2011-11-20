VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColourTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_name As String
Private m_palette(0 To 15) As Long
Private m_eventColours(1 To 17) As Byte

Private m_switchbarColourEvent As Byte
Private m_switchbarColourMessage As Byte
Private m_switchbarColourHighlight As Byte

Private m_backgroundColour As Byte

Friend Sub copy(target As CColourTheme)
10        target.name = m_name
20        target.setEventColours m_eventColours
30        target.setPalette m_palette
40        target.switchbarColourEvent = m_switchbarColourEvent
50        target.switchbarColourMessage = m_switchbarColourMessage
60        target.switchbarColourHighlight = m_switchbarColourHighlight
70        target.backgroundColour = m_backgroundColour
End Sub

Public Property Get backgroundColour() As Byte
10        backgroundColour = m_backgroundColour
End Property

Public Property Let backgroundColour(newValue As Byte)
10        m_backgroundColour = newValue
End Property

Public Property Get switchbarColourEvent() As Byte
10        switchbarColourEvent = m_switchbarColourEvent
End Property

Public Property Let switchbarColourEvent(newValue As Byte)
10        m_switchbarColourEvent = newValue
End Property

Public Property Get switchbarColourMessage() As Byte
10        switchbarColourMessage = m_switchbarColourMessage
End Property

Public Property Let switchbarColourMessage(newValue As Byte)
10        m_switchbarColourMessage = newValue
End Property

Public Property Get switchbarColourHighlight() As Byte
10        switchbarColourHighlight = m_switchbarColourHighlight
End Property

Public Property Let switchbarColourHighlight(newValue As Byte)
10        m_switchbarColourHighlight = newValue
End Property

Public Property Get name() As String
10        name = m_name
End Property

Public Property Let name(newValue As String)
10        m_name = newValue
End Property

Public Function getPalette() As Long()
10        getPalette = m_palette()
End Function

Public Sub setPalette(newPalette() As Long)
          Dim count As Long
          
10        For count = 0 To UBound(m_palette)
20            m_palette(count) = newPalette(count)
30        Next count
End Sub

Public Function getEventColours() As Byte()
10        getEventColours = m_eventColours()
End Function

Public Sub setEventColours(newEventColours() As Byte)
          Dim count As Long
          
10        For count = 1 To UBound(m_eventColours)
20            m_eventColours(count) = newEventColours(count)
30        Next count
End Sub

Public Property Get paletteEntry(index As Long) As Long
10        paletteEntry = m_palette(index)
End Property

Public Property Let paletteEntry(index As Long, newValue As Long)
10        m_palette(index) = newValue
End Property

Public Property Get paletteCount() As Long
10        paletteCount = UBound(m_palette) + 1
End Property

Public Property Get eventColour(index As Long) As Byte
10        eventColour = m_eventColours(index)
End Property

Public Property Let eventColour(index As Long, newValue As Byte)
10        m_eventColours(index) = newValue
End Property

Public Property Get eventColourCount() As Long
10        eventColourCount = UBound(m_eventColours) + 1
End Property
