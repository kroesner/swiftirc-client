VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COrigin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_nickname As String
Private m_username As String
Private m_hostname As String

Public Sub parseOrigin(origin As String)
10        If Not origin Like "*!*@*" Then
20            m_nickname = origin
30            Exit Sub
40        End If
          
          Dim temp As String

50        Do While LenB(origin) <> 0
60            If left$(origin, 1) = "!" Then
70                m_nickname = temp
80                temp = vbNullString
90            ElseIf left$(origin, 1) = "@" Then
100               m_username = temp
110               temp = vbNullString
120               m_hostname = Mid$(origin, 2)
130           Else
140               temp = temp & left$(origin, 1)
150           End If
              
160           origin = Mid$(origin, 2)
170       Loop
End Sub

Public Property Get nickname() As String
10        nickname = m_nickname
End Property

Public Property Get username() As String
10        username = m_username
End Property

Public Property Get hostname() As String
10        hostname = m_hostname
End Property
