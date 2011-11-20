VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEventColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_name As String
Private m_confName As String
Private m_colour As Byte

Public Property Get name() As String
10        name = m_name
End Property

Public Property Let name(newValue As String)
10        m_name = newValue
End Property

Public Property Get confName() As String
10        confName = m_confName
End Property

Public Property Let confName(newValue As String)
10        m_confName = newValue
End Property

Public Property Get colour() As Byte
10        colour = m_colour
End Property

Public Property Let colour(newValue As Byte)
10        m_colour = newValue
End Property
