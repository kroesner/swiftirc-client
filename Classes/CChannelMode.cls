VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CChannelMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_mode As String
Private m_type As eChannelModeType

Public Sub init(mode As String, modeType As eChannelModeType)
10        m_mode = mode
20        m_type = modeType
End Sub

Public Property Get mode() As String
10        mode = m_mode
End Property

Public Property Get modeType() As eChannelModeType
10        modeType = m_type
End Property


