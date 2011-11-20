VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParamMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_mode As String
Private m_param As String

Public Sub init(mode As String, param As String)
10        m_mode = mode
20        m_param = param
End Sub

Public Property Get mode() As String
10        mode = m_mode
End Property

Public Property Get param() As String
10        param = m_param
End Property
