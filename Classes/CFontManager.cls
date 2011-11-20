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
10        fontHeight = m_fontHeight
End Property

Friend Property Get fontStruct() As LOGFONT
10        fontStruct = m_logFont
End Property

Public Function changeFont(hdc As Long, fontName As String, fontSize As Integer, fontBold As _
    Boolean, fontItalic As Boolean) As Integer
          
10        releaseFonts
          
          Dim pixelHeight As Long
          Dim fontInfo As LOGFONT
          
20        pixelHeight = -MulDiv(fontSize, GetDeviceCaps(hdc, LOGPIXELSY), 72)
              
30        fontInfo.lfHeight = pixelHeight
          
          Dim count As Integer
          
40        For count = 1 To Len(fontName)
50            fontInfo.lfFaceName(count) = AscW(Mid$(fontName, count, 1))
              
60            If count = 32 Then
70                Exit For
80            End If
90        Next count
          
100       If fontItalic Then
110           fontInfo.lfItalic = 1
120       Else
130           fontInfo.lfItalic = 0
140       End If
          
150       If fontBold Then
160           fontInfo.lfWeight = FW_NORMAL
170       Else
180           fontInfo.lfWeight = FW_BOLD
190       End If

200       m_bold = CreateFontIndirect(fontInfo)
          
210       If fontItalic Then
220           fontInfo.lfItalic = 0
230       Else
240           fontInfo.lfItalic = 1
250       End If
          
260       m_boldItalic = CreateFontIndirect(fontInfo)
270       fontInfo.lfUnderline = 1
280       m_boldUnderlineItalic = CreateFontIndirect(fontInfo)
          
290       If fontBold Then
300           fontInfo.lfWeight = FW_BOLD
310       Else
320           fontInfo.lfWeight = FW_NORMAL
330       End If
          
340       m_underlineItalic = CreateFontIndirect(fontInfo)
          
350       fontInfo.lfUnderline = 0
360       m_italic = CreateFontIndirect(fontInfo)
          
370       If fontItalic Then
380           fontInfo.lfItalic = 1
390       Else
400           fontInfo.lfItalic = 0
410       End If
          
420       fontInfo.lfUnderline = 1
430       m_underline = CreateFontIndirect(fontInfo)
440       fontInfo.lfUnderline = 0
450       m_normal = CreateFontIndirect(fontInfo)
          
          Dim oldFont As Long
          Dim tm As TEXTMETRIC
          
460       oldFont = SelectObject(hdc, m_normal)
470       GetTextMetrics hdc, tm
480       SelectObject hdc, oldFont
          
490       m_fontHeight = tm.tmHeight
500       changeFont = tm.tmHeight
          
510       m_logFont = fontInfo
End Function

Public Function getFont(bold As Boolean, italic As Boolean, underline As Boolean) As Long
10        If bold Then
20            If italic Then
30                If underline Then
40                    getFont = m_boldUnderlineItalic
50                Else
60                    getFont = m_boldItalic
70                End If
80            ElseIf underline Then
90                getFont = m_boldUnderline
100           Else
110               getFont = m_bold
120           End If
130       ElseIf italic Then

140           If underline Then
150               getFont = m_underlineItalic
160           Else
170               getFont = m_italic
180           End If
190       ElseIf underline Then
200           getFont = m_underline
210       Else
220           getFont = m_normal
230       End If
End Function

Public Function getDefaultFont() As Long
10        getDefaultFont = m_normal
End Function

Private Sub releaseFonts()
10        If m_normal <> 0 Then
20            DeleteObject m_normal
30            m_normal = 0
40        End If
          
50        If m_bold <> 0 Then
60            DeleteObject m_bold
70            m_bold = 0
80        End If
          
90        If m_underline <> 0 Then
100           DeleteObject m_underline
110           m_underline = 0
120       End If
          
130       If m_italic <> 0 Then
140           DeleteObject m_italic
150           m_italic = 0
160       End If
          
170       If m_boldUnderline <> 0 Then
180           DeleteObject m_boldUnderline
190           m_boldUnderline = 0
200       End If
          
210       If m_boldItalic <> 0 Then
220           DeleteObject m_boldItalic
230           m_boldItalic = 0
240       End If
          
250       If m_boldUnderlineItalic <> 0 Then
260           DeleteObject m_boldUnderlineItalic
270           m_boldUnderlineItalic = 0
280       End If
          
290       If m_underlineItalic <> 0 Then
300           DeleteObject m_underlineItalic
310           m_underlineItalic = 0
320       End If
End Sub

Private Sub Class_Terminate()
10        releaseFonts
End Sub
