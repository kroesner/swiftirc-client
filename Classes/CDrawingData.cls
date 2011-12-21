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
    ignoreSeperators = m_ignoreSeperators
End Property

Public Property Let ignoreSeperators(newValue As Boolean)
    m_ignoreSeperators = newValue
End Property

Public Property Get fillBrush() As Long
    fillBrush = m_fillbrush
End Property

Public Property Get altFillBrush() As Long
    altFillBrush = m_altFillBrush
End Property

Public Property Get top() As Integer
    top = m_top
End Property

Public Property Let top(newValue As Integer)
    m_top = newValue
End Property

Public Property Get bottom() As Integer
    bottom = m_bottom
End Property

Public Property Let bottom(newValue As Integer)
    m_bottom = newValue
End Property

Public Property Get left() As Integer
    left = m_left
End Property

Public Property Let left(newValue As Integer)
    m_left = newValue
End Property

Public Property Get right() As Integer
    right = m_right
End Property

Public Property Let right(newValue As Integer)
    m_right = newValue
End Property

Public Property Get x() As Integer
    x = m_posX
End Property

Public Property Let x(newValue As Integer)
    m_posX = newValue
End Property

Public Property Get y() As Integer
    y = m_posY
End Property

Public Property Let y(newValue As Integer)
    m_posY = newValue
End Property

Public Property Get realY() As Integer
    realY = m_realPosY
End Property

Public Property Let realY(newValue As Integer)
    m_realPosY = newValue
End Property

Public Property Get width() As Integer
    width = m_width
End Property

Public Property Let width(newValue As Integer)
    m_width = newValue
End Property

Public Property Get Dc() As Long
    Dc = m_dc
End Property

Public Property Let Dc(newValue As Long)
    m_dc = newValue
End Property

Public Property Get defaultForeColour() As Byte
    defaultForeColour = m_defaultForeColour
End Property

Public Property Let defaultForeColour(newValue As Byte)
    m_defaultForeColour = newValue
    m_lineForeColour = newValue
End Property

Public Property Get defaultBackColour() As Byte
    defaultBackColour = m_defaultBackColour
End Property

Public Property Let defaultBackColour(newValue As Byte)
    m_defaultBackColour = newValue
    updateFillBrush
End Property

Public Property Get lineForeColour() As Byte
    lineForeColour = m_lineForeColour
End Property

Public Property Let lineForeColour(newValue As Byte)
    m_lineForeColour = newValue
End Property

Public Property Get bold() As Boolean
    bold = m_bold
End Property

Public Property Let bold(newValue As Boolean)
    If newValue <> m_bold Then
        m_bold = newValue
        SelectObject m_dc, m_fontManager.getFont(m_bold, m_italic, m_underline)
    End If
End Property

Public Property Get italic() As Boolean
    italic = m_italic
End Property

Public Property Let italic(newValue As Boolean)
    If newValue <> m_italic Then
        m_italic = newValue
        SelectObject m_dc, m_fontManager.getFont(m_bold, m_italic, m_underline)
    End If
End Property

Public Property Get underline() As Boolean
    underline = m_underline
End Property

Public Property Let underline(newValue As Boolean)
    If newValue <> m_underline Then
        m_underline = newValue
        SelectObject m_dc, m_fontManager.getFont(m_bold, m_italic, m_underline)
    End If
End Property

Public Property Get reverse() As Boolean
    reverse = m_reverse
End Property

Public Property Let reverse(newValue As Boolean)
    If newValue <> m_reverse Then
        m_reverse = newValue
        
        If m_reverse = True Then
            SetTextColor m_dc, m_palette(m_defaultBackColour)
            SetBkColor m_dc, m_palette(m_defaultForeColour)
        Else
            If m_hasForeColour Then
                SetTextColor m_dc, m_palette(foreColour)
            Else
                SetTextColor m_dc, m_palette(m_lineForeColour)
            End If
            
            If m_hasBackColour Then
                SetBkColor m_dc, m_palette(m_backColour)
            Else
                SetBkColor m_dc, m_palette(m_defaultBackColour)
            End If
        End If
    End If
End Property

Public Property Get hasForeColour() As Boolean
    hasForeColour = m_hasForeColour
End Property

Public Property Let hasForeColour(newValue As Boolean)
    If newValue <> m_hasForeColour Then
        m_hasForeColour = newValue
        
        If Not m_reverse Then
            If m_hasForeColour Then
                SetTextColor m_dc, m_palette(m_foreColour)
            Else
                SetTextColor m_dc, m_palette(m_lineForeColour)
            End If
        End If
    End If
End Property

Public Property Get hasBackColour() As Boolean
    hasBackColour = m_hasBackColour
End Property

Public Property Let hasBackColour(newValue As Boolean)
    If newValue <> m_hasBackColour Then
        m_hasBackColour = newValue
        
        If Not m_reverse Then
            If m_hasBackColour Then
                SetBkColor m_dc, m_palette(m_backColour)
            Else
                SetBkColor m_dc, m_palette(m_defaultBackColour)
            End If
        End If
    End If
End Property

Public Property Get foreColour() As Byte
    foreColour = m_foreColour
End Property

Public Property Let foreColour(newValue As Byte)
    If newValue <> m_foreColour Then
        m_foreColour = newValue
        
        If Not m_reverse Then
            If m_hasForeColour Then
                SetTextColor m_dc, m_palette(m_foreColour)
            End If
        End If
    End If
End Property

Public Property Get backColour() As Byte
    backColour = m_backColour
End Property

Public Property Let backColour(newValue As Byte)
    If newValue <> m_backColour Then
        m_backColour = newValue
        
        If Not m_reverse Then
            If m_hasBackColour Then
                SetBkColor m_dc, m_palette(m_backColour)
            End If
        End If
    End If
End Property

Public Sub setPalette(newPalette() As Long)
    m_palette = newPalette
    updateFillBrush
End Sub

Public Property Get fontHeight() As Long
    fontHeight = m_fontHeight
End Property

Public Property Let fontHeight(newValue As Long)
    m_fontHeight = newValue
End Property

Public Property Get fontManager() As CFontManager
    Set fontManager = m_fontManager
End Property

Public Property Let fontManager(newValue As CFontManager)
    Set m_fontManager = newValue
End Property

Private Sub updateFillBrush()
    If m_fillbrush <> 0 Then
        DeleteObject m_fillbrush
    End If
    
    If m_altFillBrush <> 0 Then
        DeleteObject m_altFillBrush
    End If
    
    m_fillbrush = CreateSolidBrush(m_palette(m_defaultBackColour))
    m_altFillBrush = CreateSolidBrush(m_palette(m_defaultForeColour))
End Sub

Public Sub fillLineIndent(left As Integer)
    Dim fill As RECT
    
    fill.top = m_posY
    fill.bottom = m_posY + m_fontHeight
    fill.left = 0
    fill.right = left
    FillRect m_dc, fill, m_fillbrush
End Sub

Public Sub fillRemainingLine()
    If m_posX < m_right Then
        Dim fill As RECT
        
        fill.top = m_posY
        fill.bottom = m_posY + m_fontHeight
        fill.left = m_posX
        fill.right = m_right
        
        FillRect m_dc, fill, m_fillbrush
    End If
End Sub

Public Sub fillSpace()
    If m_posY > m_top - m_fontHeight Then
        Dim fill As RECT
        
        fill.top = m_top
        fill.bottom = m_posY + m_fontHeight
        fill.left = m_left
        fill.right = m_right
        
        FillRect m_dc, fill, m_fillbrush
    End If
End Sub

Public Sub reset()
    m_bold = False
    m_underline = False
    m_italic = False
    m_reverse = False
    
    m_hasForeColour = False
    m_hasBackColour = False
    
    If Not m_fontManager Is Nothing Then
        SelectObject m_dc, m_fontManager.getFont(False, False, False)
    End If
    
    SetTextColor m_dc, m_palette(m_lineForeColour)
    SetBkColor m_dc, m_palette(m_defaultBackColour)
End Sub
