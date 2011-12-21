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
    target.name = m_name
    target.setEventColours m_eventColours
    target.setPalette m_palette
    target.switchbarColourEvent = m_switchbarColourEvent
    target.switchbarColourMessage = m_switchbarColourMessage
    target.switchbarColourHighlight = m_switchbarColourHighlight
    target.backgroundColour = m_backgroundColour
End Sub

Public Property Get backgroundColour() As Byte
    backgroundColour = m_backgroundColour
End Property

Public Property Let backgroundColour(newValue As Byte)
    m_backgroundColour = newValue
End Property

Public Property Get switchbarColourEvent() As Byte
    switchbarColourEvent = m_switchbarColourEvent
End Property

Public Property Let switchbarColourEvent(newValue As Byte)
    m_switchbarColourEvent = newValue
End Property

Public Property Get switchbarColourMessage() As Byte
    switchbarColourMessage = m_switchbarColourMessage
End Property

Public Property Let switchbarColourMessage(newValue As Byte)
    m_switchbarColourMessage = newValue
End Property

Public Property Get switchbarColourHighlight() As Byte
    switchbarColourHighlight = m_switchbarColourHighlight
End Property

Public Property Let switchbarColourHighlight(newValue As Byte)
    m_switchbarColourHighlight = newValue
End Property

Public Property Get name() As String
    name = m_name
End Property

Public Property Let name(newValue As String)
    m_name = newValue
End Property

Public Function getPalette() As Long()
    getPalette = m_palette()
End Function

Public Sub setPalette(newPalette() As Long)
    Dim count As Long
    
    For count = 0 To UBound(m_palette)
        m_palette(count) = newPalette(count)
    Next count
End Sub

Public Function getEventColours() As Byte()
    getEventColours = m_eventColours()
End Function

Public Sub setEventColours(newEventColours() As Byte)
    Dim count As Long
    
    For count = 1 To UBound(m_eventColours)
        m_eventColours(count) = newEventColours(count)
    Next count
End Sub

Public Property Get paletteEntry(index As Long) As Long
    paletteEntry = m_palette(index)
End Property

Public Property Let paletteEntry(index As Long, newValue As Long)
    m_palette(index) = newValue
End Property

Public Property Get paletteCount() As Long
    paletteCount = UBound(m_palette) + 1
End Property

Public Property Get eventColour(index As Long) As Byte
    eventColour = m_eventColours(index)
End Property

Public Property Let eventColour(index As Long, newValue As Byte)
    m_eventColours(index) = newValue
End Property

Public Property Get eventColourCount() As Long
    eventColourCount = UBound(m_eventColours) + 1
End Property
