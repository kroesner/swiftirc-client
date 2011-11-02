VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNicknameAppearance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_userStyle As CUserStyle
Public mask As String
Public modes As String

Public matchMask As Boolean
Public matchModes As Boolean

Public Function copy() As CNicknameAppearance
    Set copy = New CNicknameAppearance
    copy.matchMask = matchMask
    copy.matchModes = matchModes
    copy.mask = mask
    copy.modes = modes
    copy.userStyle = m_userStyle.copy
End Function

Public Property Get userStyle() As CUserStyle
    Set userStyle = m_userStyle
End Property

Public Property Let userStyle(newValue As CUserStyle)
    Set m_userStyle = newValue
End Property
