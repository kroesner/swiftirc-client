VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmoticon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_text As String
Private m_icon As CImage

Public Sub init(text As String, icon As CImage)
    m_text = text
    Set m_icon = icon
End Sub

Public Property Get text() As String
    text = m_text
End Property

Public Property Get icon() As CImage
    Set icon = m_icon
End Property

