VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNicklistItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_text As String
Private m_prefix As String

Public Property Get text() As String
    text = m_text
End Property

Public Property Let text(newValue As String)
    m_text = newValue
End Property

Public Property Get prefix() As String
    prefix = m_prefix
End Property

Public Property Let prefix(newValue As String)
    m_prefix = newValue
End Property

