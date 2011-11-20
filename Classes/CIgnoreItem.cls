VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIgnoreItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum ignoreFlags
    IGNORE_NONE = 0
    IGNORE_PRIVATE = &H1
    IGNORE_CHANNEL = &H2
    IGNORE_NOTICE = &H4
    IGNORE_CTCP = &H8
    IGNORE_INVITE = &H10
    IGNORE_CODES = &H20
    IGNORE_EXCLUDE = &H40
    IGNORE_PRIVATE_EXTENDED = IGNORE_PRIVATE Or IGNORE_NOTICE Or IGNORE_CTCP Or IGNORE_INVITE
    IGNORE_ALL = IGNORE_PRIVATE Or IGNORE_CHANNEL Or IGNORE_NOTICE Or IGNORE_CTCP Or IGNORE_INVITE Or IGNORE_CODES
End Enum

Private Type ignoreFlagChar
    flag As ignoreFlags
    char As String
End Type

Private m_mask As String
Private m_flags As Long
Private m_flagChars(6) As ignoreFlagChar

Public Property Get mask() As String
10        mask = m_mask
End Property

Public Property Let mask(newValue As String)
10        m_mask = newValue
End Property

Public Property Get flags() As Long
10        flags = m_flags
End Property

Public Property Let flags(newValue As Long)
10        m_flags = newValue
End Property

Public Property Get flagChars() As String
          Dim count As Long
          
10        For count = 0 To UBound(m_flagChars)
20            If m_flags And m_flagChars(count).flag Then
30                flagChars = flagChars & m_flagChars(count).char
40            End If
50        Next count
End Property

Public Function parseFlagChars(flagChars As String) As Boolean
          Dim count As Long
          Dim flagCount As Long
          Dim char As String
          
10        m_flags = 0
20        flagChars = LCase(flagChars)
          
30        For count = 1 To Len(flagChars)
40            char = Mid$(flagChars, count, 1)
              
50            For flagCount = 0 To UBound(m_flagChars)
60                If char = m_flagChars(flagCount).char Then
70                    m_flags = m_flags Or m_flagChars(flagCount).flag
80                    Exit For
90                End If
100           Next flagCount
110       Next count
          
120       If m_flags <> 0 Then
130           parseFlagChars = True
140       End If
End Function

Private Sub Class_Initialize()
10        m_flagChars(0).flag = IGNORE_PRIVATE
20        m_flagChars(0).char = "p"
30        m_flagChars(1).flag = IGNORE_CHANNEL
40        m_flagChars(1).char = "c"
50        m_flagChars(2).flag = IGNORE_NOTICE
60        m_flagChars(2).char = "n"
70        m_flagChars(3).flag = IGNORE_CTCP
80        m_flagChars(3).char = "t"
90        m_flagChars(4).flag = IGNORE_INVITE
100       m_flagChars(4).char = "i"
110       m_flagChars(5).flag = IGNORE_CODES
120       m_flagChars(5).char = "f"
130       m_flagChars(6).flag = IGNORE_EXCLUDE
140       m_flagChars(6).char = "w"
End Sub
