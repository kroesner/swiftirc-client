VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFontManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_normal As Long
Private m_bold As Long
Private m_underline As Long
Private m_italic As Long
Private m_boldUnderline As Long
Private m_boldItalic As Long
Private m_boldUnderlineItalic As Long
Private m_underlineItalic As Long
Private m_fontHeight As Long
Private m_logFont As LOGFONT

Public Property Get fontHeight() As Long
    fontHeight = m_fontHeight
End Property

Friend Property Get fontStruct() As LOGFONT
    fontStruct = m_logFont
End Property

Public Function changeFont(hdc As Long, fontName As String, fontSize As Integer, fontBold As _
    Boolean, fontItalic As Boolean) As Integer
    
    releaseFonts
    
    Dim pixelHeight As Long
    Dim fontInfo As LOGFONT
    
    pixelHeight = -MulDiv(fontSize, GetDeviceCaps(hdc, LOGPIXELSY), 72)
        
    fontInfo.lfHeight = pixelHeight
    
    Dim count As Integer
    
    For count = 1 To Len(fontName)
        fontInfo.lfFaceName(count) = AscW(Mid$(fontName, count, 1))
        
        If count = 32 Then
            Exit For
        End If
    Next count
    
    If fontItalic Then
        fontInfo.lfItalic = 1
    Else
        fontInfo.lfItalic = 0
    End If
    
    If fontBold Then
        fontInfo.lfWeight = FW_NORMAL
    Else
        fontInfo.lfWeight = FW_BOLD
    End If

    m_bold = CreateFontIndirect(fontInfo)
    
    If fontItalic Then
        fontInfo.lfItalic = 0
    Else
        fontInfo.lfItalic = 1
    End If
    
    m_boldItalic = CreateFontIndirect(fontInfo)
    fontInfo.lfUnderline = 1
    m_boldUnderlineItalic = CreateFontIndirect(fontInfo)
    
    If fontBold Then
        fontInfo.lfWeight = FW_BOLD
    Else
        fontInfo.lfWeight = FW_NORMAL
    End If
    
    m_underlineItalic = CreateFontIndirect(fontInfo)
    
    fontInfo.lfUnderline = 0
    m_italic = CreateFontIndirect(fontInfo)
    
    If fontItalic Then
        fontInfo.lfItalic = 1
    Else
        fontInfo.lfItalic = 0
    End If
    
    fontInfo.lfUnderline = 1
    m_underline = CreateFontIndirect(fontInfo)
    fontInfo.lfUnderline = 0
    m_normal = CreateFontIndirect(fontInfo)
    
    Dim oldFont As Long
    Dim tm As TEXTMETRIC
    
    oldFont = SelectObject(hdc, m_normal)
    GetTextMetrics hdc, tm
    SelectObject hdc, oldFont
    
    m_fontHeight = tm.tmHeight
    changeFont = tm.tmHeight
    
    m_logFont = fontInfo
End Function

Public Function getFont(bold As Boolean, italic As Boolean, underline As Boolean) As Long
    If bold Then
        If italic Then
            If underline Then
                getFont = m_boldUnderlineItalic
            Else
                getFont = m_boldItalic
            End If
        ElseIf underline Then
            getFont = m_boldUnderline
        Else
            getFont = m_bold
        End If
    ElseIf italic Then

        If underline Then
            getFont = m_underlineItalic
        Else
            getFont = m_italic
        End If
    ElseIf underline Then
        getFont = m_underline
    Else
        getFont = m_normal
    End If
End Function

Public Function getDefaultFont() As Long
    getDefaultFont = m_normal
End Function

Private Sub releaseFonts()
    If m_normal <> 0 Then
        DeleteObject m_normal
        m_normal = 0
    End If
    
    If m_bold <> 0 Then
        DeleteObject m_bold
        m_bold = 0
    End If
    
    If m_underline <> 0 Then
        DeleteObject m_underline
        m_underline = 0
    End If
    
    If m_italic <> 0 Then
        DeleteObject m_italic
        m_italic = 0
    End If
    
    If m_boldUnderline <> 0 Then
        DeleteObject m_boldUnderline
        m_boldUnderline = 0
    End If
    
    If m_boldItalic <> 0 Then
        DeleteObject m_boldItalic
        m_boldItalic = 0
    End If
    
    If m_boldUnderlineItalic <> 0 Then
        DeleteObject m_boldUnderlineItalic
        m_boldUnderlineItalic = 0
    End If
    
    If m_underlineItalic <> 0 Then
        DeleteObject m_underlineItalic
        m_underlineItalic = 0
    End If
End Sub

Private Sub Class_Terminate()
    releaseFonts
End Sub
