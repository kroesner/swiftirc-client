VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_control As VBControlExtender
Private m_id As String

Public Property Get control() As VBControlExtender
10        Set control = m_control
End Property

Public Property Let control(newValue As VBControlExtender)
10        Set m_control = newValue
End Property

Public Property Get id() As String
10        id = m_id
End Property

Public Property Let id(newValue As String)
10        m_id = newValue
End Property
