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
    backgroundSelected = m_backgroundSelected
End Property

Public Property Let backgroundColour(newValue As Long)
    m_backgroundColour = newValue
    reDraw
End Property

Public Property Get selectedItem() As Long
    selectedItem = m_selectedLabel
End Property

Public Sub setPalette(newPalette() As Long)
    m_paletteSet = True
    m_palette() = newPalette
    reDraw
End Sub

Public Sub setEventColours(newEventColours() As Byte)
    m_eventColours() = newEventColours
    reDraw
End Sub

Private Sub IColourUser_coloursUpdated()

End Sub

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Sub IFontUser_fontsUpdated()

End Sub

Private Property Let IFontUser_fontManager(RHS As CFontManager)

End Property

Public Sub reDraw()
    If Not m_paletteSet Then
        Exit Sub
    End If

    Dim backBuffer As Long
    Dim backBitmap As Long
    Dim oldBitmap As Long
    
    Dim coordX As Long
    Dim coordY As Long
    
    backBuffer = CreateCompatibleDC(UserControl.hdc)
    backBitmap = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, _
        UserControl.ScaleHeight)
    
    oldBitmap = SelectObject(backBuffer, backBitmap)
    
    SetBkMode backBuffer, TRANSPARENT
    
    Dim oldBrush As Long
    Dim oldPen As Long
    Dim fillBrush As Long
    
    fillBrush = CreateSolidBrush(m_palette(m_backgroundColour))
    
    oldBrush = SelectObject(backBuffer, fillBrush)
    oldPen = SelectObject(backBuffer, colourManager.getPen(SWIFTPEN_BORDER))
    
    Rectangle backBuffer, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    
    SelectObject backBuffer, oldBrush
    SelectObject backBuffer, oldPen
    
    DeleteObject fillBrush
    
    Dim count As Long
    Dim textMetrics As TEXTMETRIC
    
    GetTextMetrics UserControl.hdc, textMetrics
    
    Dim textRect As RECT
    
    coordX = 5
    coordY = 5
     
    Dim textSize As SIZE
    Dim tempRect As RECT
    
    ReDim m_rects(UBound(m_eventColours))
     
    For count = 1 To UBound(m_eventColours)
        If m_eventColours(count) <= UBound(m_palette) Then
            SetTextColor backBuffer, m_palette(m_eventColours(count))
        End If
    
        swiftGetTextExtentPoint32 backBuffer, eventColours.getName(count), textSize
    
        m_rects(count) = makeRect(coordX, coordX + textSize.cx, coordY, coordY + textMetrics.tmHeight)
        
        swiftTextOut backBuffer, coordX, coordY, ETO_CLIPPED, VarPtr(m_rects(count)), eventColours.getName(count)
        
        If m_selectedLabel = count Then
            SetTextColor backBuffer, m_palette(g_textViewFore)
            
            tempRect = m_rects(count)
            tempRect.left = tempRect.left - 1
            tempRect.right = tempRect.right + 1
            DrawFocusRect backBuffer, tempRect
        End If
        
        coordY = coordY + textMetrics.tmHeight
        
        If coordY + textMetrics.tmHeight > UserControl.ScaleHeight - 5 Then
            coordY = 5
            coordX = coordX + COLUMN_SIZE + COLUMN_PADDING
        End If
    Next count
    
    If m_backgroundSelected Then
        DrawFocusRect backBuffer, makeRect(3, UserControl.ScaleWidth - 3, 3, UserControl.ScaleHeight - 3)
    End If
    
    BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, backBuffer, 0, 0, _
        vbSrcCopy
    
    SelectObject backBuffer, oldBitmap
    
    DeleteObject backBitmap
    DeleteDC backBuffer
End Sub

Public Sub colourSelected(index As Long)
    If m_backgroundSelected Then
        m_backgroundColour = index
    Else
        m_eventColours(m_selectedLabel) = index
    End If
    
    reDraw
End Sub

Public Sub paletteEntryUpdated(index As Long, newColour As Long)
    m_palette(index) = newColour
    reDraw
End Sub

Private Sub UserControl_Initialize()
    ReDim m_palette(0)
    ReDim m_eventColours(0)
    
    m_selectedLabel = 1
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbKeyLButton Then
        Exit Sub
    End If
    
    Dim fontHeight As Long
    Dim textMetrics As TEXTMETRIC
    
    GetTextMetrics UserControl.hdc, textMetrics
    fontHeight = textMetrics.tmHeight
    
    Dim column As Long
    Dim row As Long
    Dim labelIndex As Long
    
    row = Fix((y - 5) / fontHeight) + 1
    column = Fix((x - COLUMN_PADDING) / COLUMN_SIZE)
    
    labelIndex = (Fix((UserControl.ScaleHeight - 10) / fontHeight) * column) + row
    
    If labelIndex < 1 Or labelIndex > UBound(m_eventColours) Then
        selectBackground
        Exit Sub
    End If
    
    If PtInRect(m_rects(labelIndex), x, y) = 0 Then
        selectBackground
        Exit Sub
    End If
    
    m_backgroundSelected = False
    m_selectedLabel = labelIndex
    reDraw
    
    RaiseEvent itemClicked(m_eventColours(labelIndex))
End Sub

Private Sub selectBackground()
    m_backgroundSelected = True
    m_selectedLabel = -1
    reDraw
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbKeyLButton Then
        UserControl_MouseDown Button, Shift, x, y
    End If
End Sub

Private Sub UserControl_Paint()
    reDraw
End Sub

Private Sub UserControl_Terminate()
    debugLog "ctlEventColourEditor terminating"
End Sub
