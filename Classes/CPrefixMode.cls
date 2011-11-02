VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrefixMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_mode As String
Private m_symbol As String

Public Sub init(mode As String, symbol As String)
    m_mode = mode
    m_symbol = symbol
End Sub

Public Property Get mode() As String
    mode = m_mode
End Property

Public Property Get symbol() As String
    symbol = m_symbol
End Property
