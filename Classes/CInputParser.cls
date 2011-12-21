VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInputParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_command As String
Private m_minParams As Byte
Private m_maxParams As Byte
Private m_func As String

Public Sub init(command As String, minParams As Byte, maxParams As Byte, func As String)
    m_command = command
    m_minParams = minParams
    m_maxParams = maxParams
    m_func = func
End Sub

Public Property Get command() As String
    command = m_command
End Property

Public Property Get minParams() As Byte
    minParams = m_minParams
End Property

Public Property Get maxParams() As Byte
    maxParams = m_maxParams
End Property

Public Property Get func() As String
    func = m_func
End Property

