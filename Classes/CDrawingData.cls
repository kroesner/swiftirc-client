VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDrawingData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ignoreSeperators As Boolean

Private m_top As Integer
Private m_bottom As Integer
Private m_left As Integer
Private m_right As Integer

Private m_posX As Integer
Private m_posY As Integer

Private m_width As Integer

Private m_realPosY As Integer

Private m_dc As Long
Private m_bold As Boolean
Private m_italic As Boolean
Private m_underline As Boolean
Private m_reverse As Boolean

Private m_hasForeColour As Boolean
Private m_hasBackColour As Boolean

Private m_foreColour As Byte
Private m_backColour As Byte

Private m_fontManager As CFontManager
Private m_fontHeight As Long
Private m_palette() As Long

Private m_fillbrush As Long
Private m_altFillBrush As Long

Private m_defaultForeColour As Byte
Private m_defaultBackColour As Byte

Private m_lineForeColour As Byte

Public Property Get ignoreSeperators() As Boolean
10        ignoreSeperators = m_ignoreSeperators
End Property

Public Property Let ignoreSeperators(newValue As Boolean)
10        m_ignoreSeperators = newValue
End Property

Public Property Get fillBrush() As Long
10        fillBrush = m_fillbrush
End Property

Public Property Get altFillBrush() As Long
10        altFillBrush = m_altFillBrush
End Property

Public Property Get top() As Integer
10        top = m_top
End Property

Public Property Let top(newValue As Integer)
10        m_top = newValue
End Property

Public Property Get bottom() As Integer
10        bottom = m_bottom
End Property

Public Property Let bottom(newValue As Integer)
10        m_bottom = newValue
End Property

Public Property Get left() As Integer
10        left = m_left
End Property

Public Property Let left(newValue As Integer)
10        m_left = newValue
End Property

Public Property Get right() As Integer
10        right = m_right
End Property

Public Property Let right(newValue As Integer)
10        m_right = newValue
End Property

Public Property Get x() As Integer
10        x = m_posX
End Property

Public Property Let x(newValue As Integer)
10        m_posX = newValue
End Property

Public Property Get y() As Integer
10        y = m_posY
End Property

Public Property Let y(newValue As Integer)
10        m_posY = newValue
End Property

Public Property Get realY() As Integer
10        realY = m_realPosY
End Property

Public Property Let realY(newValue As Integer)
10        m_realPosY = newValue
End Property

Public Property Get width() As Integer
10        width = m_width
End Property

Public Property Let width(newValue As Integer)
10        m_width = newValue
End Property

Public Property Get Dc() As Long
10        Dc = m_dc
End Property

Public Property Let Dc(newValue As Long)
10        m_dc = newValue
End Property

Public Property Get defaultForeColour() As Byte
10        defaultForeColour = m_defaultForeColour
End Property

Public Property Let defaultForeColour(newValue As Byte)
10        m_defaultForeColour = newValue
20        m_lineForeColour = newValue
End Property

Public Property Get defaultBackColour() As Byte
10        defaultBackColour = m_defaultBackColour
End Property

Public Property Let defaultBackColour(newValue As Byte)
10        m_defaultBackColour = newValue
20        updateFillBrush
End Property

Public Property Get lineForeColour() As Byte
10        lineForeColour = m_lineForeColour
End Property

Public Property Let lineForeColour(newValue As Byte)
10        m_lineForeColour = newValue
End Property

Public Property Get bold() As Boolean
10        bold = m_bold
End Property

Public Property Let bold(newValue As Boolean)
10        If newValue <> m_bold Then
20            m_bold = newValue
30            SelectObject m_dc, m_fontManager.getFont(m_bold, m_italic, m_underline)
40        End If
End Property

Public Property Get italic() As Boolean
10        italic = m_italic
End Property

Public Property Let italic(newValue As Boolean)
10        If newValue <> m_italic Then
20            m_italic = newValue
30            SelectObject m_dc, m_fontManager.getFont(m_bold, m_italic, m_underline)
40        End If
End Property

Public Property Get underline() As Boolean
10        underline = m_underline
End Property

Public Property Let underline(newValue As Boolean)
10        If newValue <> m_underline Then
20            m_underline = newValue
30            SelectObject m_dc, m_fontManager.getFont(m_bold, m_italic, m_underline)
40        End If
End Property

Public Property Get reverse() As Boolean
10        reverse = m_reverse
End Property

Public Property Let reverse(newValue As Boolean)
10        If newValue <> m_reverse Then
20            m_reverse = newValue
              
30            If m_reverse = True Then
40                SetTextColor m_dc, m_palette(m_defaultBackColour)
50                SetBkColor m_dc, m_palette(m_defaultForeColour)
60            Else
70                If m_hasForeColour Then
80                    SetTextColor m_dc, m_palette(foreColour)
90                Else
100                   SetTextColor m_dc, m_palette(m_lineForeColour)
110               End If
                  
120               If m_hasBackColour Then
130                   SetBkColor m_dc, m_palette(m_backColour)
140               Else
150                   SetBkColor m_dc, m_palette(m_defaultBackColour)
160               End If
170           End If
180       End If
End Property

Public Property Get hasForeColour() As Boolean
10        hasForeColour = m_hasForeColour
End Property

Public Property Let hasForeColour(newValue As Boolean)
10        If newValue <> m_hasForeColour Then
20            m_hasForeColour = newValue
              
30            If Not m_reverse Then
40                If m_hasForeColour Then
50                    SetTextColor m_dc, m_palette(m_foreColour)
60                Else
70                    SetTextColor m_dc, m_palette(m_lineForeColour)
80                End If
90            End If
100       End If
End Property

Public Property Get hasBackColour() As Boolean
10        hasBackColour = m_hasBackColour
End Property

Public Property Let hasBackColour(newValue As Boolean)
10        If newValue <> m_hasBackColour Then
20            m_hasBackColour = newValue
              
30            If Not m_reverse Then
40                If m_hasBackColour Then
50                    SetBkColor m_dc, m_palette(m_backColour)
60                Else
70                    SetBkColor m_dc, m_palette(m_defaultBackColour)
80                End If
90            End If
100       End If
End Property

Public Property Get foreColour() As Byte
10        foreColour = m_foreColour
End Property

Public Property Let foreColour(newValue As Byte)
10        If newValue <> m_foreColour Then
20            m_foreColour = newValue
              
30            If Not m_reverse Then
40                If m_hasForeColour Then
50                    SetTextColor m_dc, m_palette(m_foreColour)
60                End If
70            End If
80        End If
End Property

Public Property Get backColour() As Byte
10        backColour = m_backColour
End Property

Public Property Let backColour(newValue As Byte)
10        If newValue <> m_backColour Then
20            m_backColour = newValue
              
30            If Not m_reverse Then
40                If m_hasBackColour Then
50                    SetBkColor m_dc, m_palette(m_backColour)
60                End If
70            End If
80        End If
End Property

Public Sub setPalette(newPalette() As Long)
10        m_palette = newPalette
20        updateFillBrush
End Sub

Public Property Get fontHeight() As Long
10        fontHeight = m_fontHeight
End Property

Public Property Let fontHeight(newValue As Long)
10        m_fontHeight = newValue
End Property

Public Property Get fontManager() As CFontManager
10        Set fontManager = m_fontManager
End Property

Public Property Let fontManager(newValue As CFontManager)
10        Set m_fontManager = newValue
End Property

Private Sub updateFillBrush()
10        If m_fillbrush <> 0 Then
20            DeleteObject m_fillbrush
30        End If
          
40        If m_altFillBrush <> 0 Then
50            DeleteObject m_altFillBrush
60        End If
          
70        m_fillbrush = CreateSolidBrush(m_palette(m_defaultBackColour))
80        m_altFillBrush = CreateSolidBrush(m_palette(m_defaultForeColour))
End Sub

Public Sub fillLineIndent(left As Integer)
          Dim fill As RECT
          
10        fill.top = m_posY
20        fill.bottom = m_posY + m_fontHeight
30        fill.left = 0
40        fill.right = left
50        FillRect m_dc, fill, m_fillbrush
End Sub

Public Sub fillRemainingLine()
10        If m_posX < m_right Then
              Dim fill As RECT
              
20            fill.top = m_posY
30            fill.bottom = m_posY + m_fontHeight
40            fill.left = m_posX
50            fill.right = m_right
              
60            FillRect m_dc, fill, m_fillbrush
70        End If
End Sub

Public Sub fillSpace()
10        If m_posY > m_top - m_fontHeight Then
              Dim fill As RECT
              
20            fill.top = m_top
30            fill.bottom = m_posY + m_fontHeight
40            fill.left = m_left
50            fill.right = m_right
              
60            FillRect m_dc, fill, m_fillbrush
70        End If
End Sub

Public Sub reset()
10        m_bold = False
20        m_underline = False
30        m_italic = False
40        m_reverse = False
          
50        m_hasForeColour = False
60        m_hasBackColour = False
          
70        If Not m_fontManager Is Nothing Then
80            SelectObject m_dc, m_fontManager.getFont(False, False, False)
90        End If
          
100       SetTextColor m_dc, m_palette(m_lineForeColour)
110       SetBkColor m_dc, m_palette(m_defaultBackColour)
End Sub
