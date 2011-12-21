VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIalEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_nick As String
Private m_ident As String
Private m_host As String

Public Property Get nick() As String
    nick = m_nick
End Property

Public Property Let nick(newValue As String)
    m_nick = newValue
End Property

Public Property Get ident() As String
    ident = m_ident
End Property

Public Property Let ident(newValue As String)
    m_ident = newValue
End Property

Public Property Get host() As String
    host = m_host
End Property

Public Property Let host(newValue As String)
    m_host = newValue
End Property
