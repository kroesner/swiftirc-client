VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrefixStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_symbol As String
Private m_name As String
Private m_style As CUserStyle

Public Property Get symbol() As String
    symbol = m_symbol
End Property

Public Property Let symbol(newValue As String)
    m_symbol = newValue
End Property

Public Property Get name() As String
    name = m_name
End Property

Public Property Let name(newValue As String)
    m_name = newValue
End Property

Public Property Get style() As CUserStyle
    Set style = m_style
End Property

Public Property Let style(newValue As CUserStyle)
    Set m_style = newValue
End Property
