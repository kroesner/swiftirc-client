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
10        target.text = m_text
End Sub

Public Property Get text() As String
10        text = m_text
End Property

Public Property Let text(newValue As String)
10        m_text = newValue
End Property

Public Property Get highlightType() As eHighlightType
10        highlightType = m_type
End Property

Public Property Let highlightType(newValue As eHighlightType)
10        m_type = newValue
End Property

Public Property Get colour() As Byte
10        colour = m_colour
End Property

Public Property Let colour(newValue As Byte)
10        m_colour = newValue
End Property
