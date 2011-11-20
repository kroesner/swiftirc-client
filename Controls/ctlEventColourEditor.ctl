VERSION 5.00
Begin VB.UserControl ctlEventColourEditor 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "ctlEventColourEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow
Implements IColourUser
Implements IFontUser

Private Const COLUMN_SIZE = 125
Private Const COLUMN_PADDING = 10

Private m_realWindow As VBControlExtender

Private m_backgroundSelected As Boolean
Private m_selectedLabel As Long

Private m_backgroundColour As Byte

Private m_eventColours() As Byte
Private m_palette() As Long

Private m_rects() As RECT

Private m_paletteSet As Boolean

Public Event itemClicked(colourIndex As Byte)

Public Property Get backgroundSelected() As Boolean
10        backgroundSelected = m_backgroundSelected
End Property

Public Property Let backgroundColour(newValue As Long)
10        m_backgroundColour = newValue
20        reDraw
End Property

Public Property Get selectedItem() As Long
10        selectedItem = m_selectedLabel
End Property

Public Sub setPalette(newPalette() As Long)
10        m_paletteSet = True
20        m_palette() = newPalette
30        reDraw
End Sub

Public Sub setEventColours(newEventColours() As Byte)
10        m_eventColours() = newEventColours
20        reDraw
End Sub

Private Sub IColourUser_coloursUpdated()

End Sub

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Sub IFontUser_fontsUpdated()

End Sub

Private Property Let IFontUser_fontManager(RHS As CFontManager)

End Property

Public Sub reDraw()
10        If Not m_paletteSet Then
20            Exit Sub
30        End If

          Dim backBuffer As Long
          Dim backBitmap As Long
          Dim oldBitmap As Long
          
          Dim coordX As Long
          Dim coordY As Long
          
40        backBuffer = CreateCompatibleDC(UserControl.hdc)
50        backBitmap = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, _
              UserControl.ScaleHeight)
          
60        oldBitmap = SelectObject(backBuffer, backBitmap)
          
70        SetBkMode backBuffer, TRANSPARENT
          
          Dim oldBrush As Long
          Dim oldPen As Long
          Dim fillBrush As Long
          
80        fillBrush = CreateSolidBrush(m_palette(m_backgroundColour))
          
90        oldBrush = SelectObject(backBuffer, fillBrush)
100       oldPen = SelectObject(backBuffer, colourManager.getPen(SWIFTPEN_BORDER))
          
110       Rectangle backBuffer, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
          
120       SelectObject backBuffer, oldBrush
130       SelectObject backBuffer, oldPen
          
140       DeleteObject fillBrush
          
          Dim count As Long
          Dim textMetrics As TEXTMETRIC
          
150       GetTextMetrics UserControl.hdc, textMetrics
          
          Dim textRect As RECT
          
160       coordX = 5
170       coordY = 5
           
          Dim textSize As SIZE
          Dim tempRect As RECT
          
180       ReDim m_rects(UBound(m_eventColours))
           
190       For count = 1 To UBound(m_eventColours)
200           If m_eventColours(count) <= UBound(m_palette) Then
210               SetTextColor backBuffer, m_palette(m_eventColours(count))
220           End If
          
230           swiftGetTextExtentPoint32 backBuffer, eventColours.getName(count), textSize
          
240           m_rects(count) = makeRect(coordX, coordX + textSize.cx, coordY, coordY + textMetrics.tmHeight)
              
250           swiftTextOut backBuffer, coordX, coordY, ETO_CLIPPED, VarPtr(m_rects(count)), eventColours.getName(count)
              
260           If m_selectedLabel = count Then
270               SetTextColor backBuffer, m_palette(g_textViewFore)
                  
280               tempRect = m_rects(count)
290               tempRect.left = tempRect.left - 1
300               tempRect.right = tempRect.right + 1
310               DrawFocusRect backBuffer, tempRect
320           End If
              
330           coordY = coordY + textMetrics.tmHeight
              
340           If coordY + textMetrics.tmHeight > UserControl.ScaleHeight - 5 Then
350               coordY = 5
360               coordX = coordX + COLUMN_SIZE + COLUMN_PADDING
370           End If
380       Next count
          
390       If m_backgroundSelected Then
400           DrawFocusRect backBuffer, makeRect(3, UserControl.ScaleWidth - 3, 3, UserControl.ScaleHeight - 3)
410       End If
          
420       BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, backBuffer, 0, 0, _
              vbSrcCopy
          
430       SelectObject backBuffer, oldBitmap
          
440       DeleteObject backBitmap
450       DeleteDC backBuffer
End Sub

Public Sub colourSelected(index As Long)
10        If m_backgroundSelected Then
20            m_backgroundColour = index
30        Else
40            m_eventColours(m_selectedLabel) = index
50        End If
          
60        reDraw
End Sub

Public Sub paletteEntryUpdated(index As Long, newColour As Long)
10        m_palette(index) = newColour
20        reDraw
End Sub

Private Sub UserControl_Initialize()
10        ReDim m_palette(0)
20        ReDim m_eventColours(0)
          
30        m_selectedLabel = 1
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
10        If Button <> vbKeyLButton Then
20            Exit Sub
30        End If
          
          Dim fontHeight As Long
          Dim textMetrics As TEXTMETRIC
          
40        GetTextMetrics UserControl.hdc, textMetrics
50        fontHeight = textMetrics.tmHeight
          
          Dim column As Long
          Dim row As Long
          Dim labelIndex As Long
          
60        row = Fix((y - 5) / fontHeight) + 1
70        column = Fix((x - COLUMN_PADDING) / COLUMN_SIZE)
          
80        labelIndex = (Fix((UserControl.ScaleHeight - 10) / fontHeight) * column) + row
          
90        If labelIndex < 1 Or labelIndex > UBound(m_eventColours) Then
100           selectBackground
110           Exit Sub
120       End If
          
130       If PtInRect(m_rects(labelIndex), x, y) = 0 Then
140           selectBackground
150           Exit Sub
160       End If
          
170       m_backgroundSelected = False
180       m_selectedLabel = labelIndex
190       reDraw
          
200       RaiseEvent itemClicked(m_eventColours(labelIndex))
End Sub

Private Sub selectBackground()
10        m_backgroundSelected = True
20        m_selectedLabel = -1
30        reDraw
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
10        If Button = vbKeyLButton Then
20            UserControl_MouseDown Button, Shift, x, y
30        End If
End Sub

Private Sub UserControl_Paint()
10        reDraw
End Sub

Private Sub UserControl_Terminate()
10        debugLog "ctlEventColourEditor terminating"
End Sub
