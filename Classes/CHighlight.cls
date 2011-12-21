VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHighlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_text As String
Private m_type As eHighlightType
Private m_colour As Byte

Public Sub copy(target As CHighlight)
    target.text = m_text
End Sub

Public Property Get text() As String
    text = m_text
End Property

Public Property Let text(newValue As String)
    m_text = newValue
End Property

Public Property Get highlightType() As eHighlightType
    highlightType = m_type
End Property

Public Property Let highlightType(newValue As eHighlightType)
    m_type = newValue
End Property

Public Property Get colour() As Byte
    colour = m_colour
End Property

Public Property Let colour(newValue As Byte)
    m_colour = newValue
End Property
