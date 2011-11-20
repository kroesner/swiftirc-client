VERSION 5.00
Begin VB.UserControl ctlColourPalette 
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3450
   ScaleHeight     =   70
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   230
End
Attribute VB_Name = "ctlColourPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow

Private m_realWindow As VBControlExtender

Private m_selectedIndex As Byte
Private m_palette() As Long

Private Const MARGIN_X As Integer = 5
Private Const MARGIN_Y As Integer = 5
Private Const SPACING_X As Integer = 3
Private Const SPACING_Y As Integer = 3

Private Const ITEM_SIZE As Integer = 20

Private m_allowPaletteChange As Boolean

Public Event colourChanged(index As Long, colour As Long)
Public Event colourSelected(index As Long)

Public Property Let allowPaletteChange(newValue As Long)
10        m_allowPaletteChange = newValue
End Property

Public Sub setPalette(newPalette() As Long)
10        m_palette() = newPalette()
20        UserControl_Paint
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Sub UserControl_Initialize()
10        ReDim m_palette(0)
End Sub

Public Sub setFocus(index As Byte)
10        m_selectedIndex = index
20        UserControl_Paint
End Sub

Private Function getIndexByCoords(x As Single, y As Single) As Integer
10        getIndexByCoords = -1

20        If x < MARGIN_X Or x >= UserControl.ScaleWidth - MARGIN_X Then
30            Exit Function
40        End If
          
          Dim buttonsPerRow As Integer
          
50        buttonsPerRow = Fix((UserControl.ScaleWidth - (MARGIN_X * 2)) / (ITEM_SIZE + SPACING_X))
          
          Dim buttonIndex As Integer
          Dim rowIndex As Integer
          
60        rowIndex = Fix((y - MARGIN_Y) / (ITEM_SIZE + SPACING_Y))
70        buttonIndex = Fix((x - MARGIN_X) / (ITEM_SIZE + SPACING_X)) + (rowIndex * buttonsPerRow)
          
80        If buttonIndex > UBound(m_palette) + 1 Then
90            Exit Function
100       End If
          
          Dim edgeX As Integer
          Dim edgeY As Integer
          
110       edgeX = MARGIN_X + ((ITEM_SIZE + SPACING_X) * ((buttonIndex + 1) - (rowIndex * buttonsPerRow))) _
              - SPACING_X
120       edgeY = MARGIN_Y + ((ITEM_SIZE + SPACING_Y) * (rowIndex + 1)) - SPACING_Y
          
130       If x > edgeX Or y > edgeY Then
140           Exit Function
150       End If
          
160       getIndexByCoords = buttonIndex
End Function

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
          Dim itemIndex As Long
          
10        itemIndex = getIndexByCoords(x, y)
          
20        If itemIndex = -1 Then
30            Exit Sub
40        End If
          
50        If Button = vbKeyLButton Then
60            m_selectedIndex = itemIndex
70            UserControl_Paint
80            RaiseEvent colourSelected(itemIndex)
90        ElseIf Button = vbKeyRButton Then
100           If m_allowPaletteChange Then
110               changeColour itemIndex
120           End If
130       End If
End Sub

Private Sub UserControl_Paint()
          Dim backBuffer As Long
          Dim backBitmap As Long
          Dim oldBitmap As Long
          
10        backBuffer = CreateCompatibleDC(UserControl.hdc)
20        backBitmap = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, _
              UserControl.ScaleHeight)
          
30        oldBitmap = SelectObject(backBuffer, backBitmap)
          
40        FillRect backBuffer, makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), _
              colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
          
          Dim itemRect As RECT
          Dim colourBrush As Long
          Dim oldBrush As Long
          Dim count As Long
          
          Dim coordY As Long
          Dim coordX As Long
          
50        coordY = MARGIN_Y
60        coordX = MARGIN_X
          
70        For count = 0 To UBound(m_palette)
80            itemRect.left = coordX
90            itemRect.right = coordX + ITEM_SIZE
100           itemRect.top = coordY
110           itemRect.bottom = coordY + ITEM_SIZE
              
120           colourBrush = CreateSolidBrush(m_palette(count))
130           FillRect backBuffer, itemRect, colourBrush
140           DeleteObject colourBrush
              
150           FrameRect backBuffer, itemRect, colourManager.getBrush(SWIFTCOLOUR_CONTROLFORE)
              
160           If count = m_selectedIndex Then
170               itemRect.left = itemRect.left + 1
180               itemRect.right = itemRect.right - 1
190               itemRect.top = itemRect.top + 1
200               itemRect.bottom = itemRect.bottom - 1
                  
210               DrawFocusRect backBuffer, itemRect
220           End If
              
230           coordX = coordX + ITEM_SIZE + SPACING_X
              
240           If coordX + ITEM_SIZE > UserControl.ScaleWidth - MARGIN_X Then
250               coordX = MARGIN_X
260               coordY = coordY + ITEM_SIZE + SPACING_Y
270           End If
280       Next count
          
290       BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, backBuffer, 0, 0, _
              vbSrcCopy
              
300       SelectObject backBuffer, oldBitmap
              
310       DeleteDC backBuffer
320       DeleteObject backBitmap
End Sub

Private Sub changeColour(index As Long)
          Dim cc As ChooseColorStruct
          
10        cc.lStructSize = Len(cc)
20        cc.hwndOwner = UserControl.hwnd
30        cc.lpCustColors = VarPtr(custChooseColorPalette(0))
          
40        If ChooseColor(cc) <> 0 Then
50            m_palette(index) = cc.rgbResult
60            UserControl_Paint
70            RaiseEvent colourChanged(index, m_palette(index))
80        End If
End Sub

Private Sub UserControl_Terminate()
10        debugLog "ctlColourPalette terminating"
End Sub
