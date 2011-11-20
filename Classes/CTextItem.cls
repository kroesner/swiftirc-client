VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTextItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_name As String
Private m_text As String

Public Property Get name() As String
10        name = m_name
End Property

Public Property Let name(newValue As String)
10        m_name = newValue
End Property

Public Property Get text() As String
10        text = m_text
End Property

Public Property Let text(newValue As String)
10        m_text = newValue
End Property
