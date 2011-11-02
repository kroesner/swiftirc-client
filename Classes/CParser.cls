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
    name = m_name
End Property

Public Property Let name(newValue As String)
    m_name = newValue
End Property

Public Property Get minParams() As Byte
    minParams = m_minParams
End Property

Public Property Let minParams(newValue As Byte)
    m_minParams = newValue
End Property

Public Property Get maxParams() As Byte
    maxParams = m_maxParams
End Property

Public Property Let maxParams(newValue As Byte)
    m_maxParams = newValue
End Property

Public Property Get func() As String
    func = m_func
End Property

Public Property Let func(newValue As String)
    m_func = newValue
End Property
