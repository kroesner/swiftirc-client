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
    IGNORE_ANY = &H0
    IGNORE_NONE = &H0
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
    mask = m_mask
End Property

Public Property Let mask(newValue As String)
    m_mask = newValue
End Property

Public Property Get flags() As Long
    flags = m_flags
End Property

Public Property Let flags(newValue As Long)
    m_flags = newValue
End Property

Public Property Get flagChars() As String
    Dim count As Long
    
    For count = 0 To UBound(m_flagChars)
        If m_flags And m_flagChars(count).flag Then
            flagChars = flagChars & m_flagChars(count).char
        End If
    Next count
End Property

Public Function parseFlagChars(flagChars As String) As Boolean
    Dim count As Long
    Dim flagCount As Long
    Dim char As String
    
    m_flags = 0
    flagChars = LCase(flagChars)
    
    For count = 1 To Len(flagChars)
        char = Mid$(flagChars, count, 1)
        
        For flagCount = 0 To UBound(m_flagChars)
            If char = m_flagChars(flagCount).char Then
                m_flags = m_flags Or m_flagChars(flagCount).flag
                Exit For
            End If
        Next flagCount
    Next count
    
    If m_flags <> 0 Then
        parseFlagChars = True
    End If
End Function

Private Sub Class_Initialize()
    m_flagChars(0).flag = IGNORE_PRIVATE
    m_flagChars(0).char = "p"
    m_flagChars(1).flag = IGNORE_CHANNEL
    m_flagChars(1).char = "c"
    m_flagChars(2).flag = IGNORE_NOTICE
    m_flagChars(2).char = "n"
    m_flagChars(3).flag = IGNORE_CTCP
    m_flagChars(3).char = "t"
    m_flagChars(4).flag = IGNORE_INVITE
    m_flagChars(4).char = "i"
    m_flagChars(5).flag = IGNORE_CODES
    m_flagChars(5).char = "f"
    m_flagChars(6).flag = IGNORE_EXCLUDE
    m_flagChars(6).char = "w"
End Sub
