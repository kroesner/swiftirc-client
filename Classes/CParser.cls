VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_name As String
Private m_minParams As Byte
Private m_maxParams As Byte
Private m_func As String

Public Property Get name() As String
10        name = m_name
End Property

Public Property Let name(newValue As String)
10        m_name = newValue
End Property

Public Property Get minParams() As Byte
10        minParams = m_minParams
End Property

Public Property Let minParams(newValue As Byte)
10        m_minParams = newValue
End Property

Public Property Get maxParams() As Byte
10        maxParams = m_maxParams
End Property

Public Property Let maxParams(newValue As Byte)
10        m_maxParams = newValue
End Property

Public Property Get func() As String
10        func = m_func
End Property

Public Property Let func(newValue As String)
10        m_func = newValue
End Property
