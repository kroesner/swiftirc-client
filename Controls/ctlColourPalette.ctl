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
    m_allowPaletteChange = newValue
End Property

Public Sub setPalette(newPalette() As Long)
    m_palette() = newPalette()
    UserControl_Paint
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Sub UserControl_Initialize()
    ReDim m_palette(0)
End Sub

Public Sub setFocus(index As Byte)
    m_selectedIndex = index
    UserControl_Paint
End Sub

Private Function getIndexByCoords(x As Single, y As Single) As Integer
    getIndexByCoords = -1

    If x < MARGIN_X Or x >= UserControl.ScaleWidth - MARGIN_X Then
        Exit Function
    End If
    
    Dim buttonsPerRow As Integer
    
    buttonsPerRow = Fix((UserControl.ScaleWidth - (MARGIN_X * 2)) / (ITEM_SIZE + SPACING_X))
    
    Dim buttonIndex As Integer
    Dim rowIndex As Integer
    
    rowIndex = Fix((y - MARGIN_Y) / (ITEM_SIZE + SPACING_Y))
    buttonIndex = Fix((x - MARGIN_X) / (ITEM_SIZE + SPACING_X)) + (rowIndex * buttonsPerRow)
    
    If buttonIndex > UBound(m_palette) + 1 Then
        Exit Function
    End If
    
    Dim edgeX As Integer
    Dim edgeY As Integer
    
    edgeX = MARGIN_X + ((ITEM_SIZE + SPACING_X) * ((buttonIndex + 1) - (rowIndex * buttonsPerRow))) - SPACING_X
    edgeY = MARGIN_Y + ((ITEM_SIZE + SPACING_Y) * (rowIndex + 1)) - SPACING_Y
    
    If x > edgeX Or y > edgeY Then
        Exit Function
    End If
    
    getIndexByCoords = buttonIndex
End Function

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim itemIndex As Long
    
    itemIndex = getIndexByCoords(x, y)
    
    If itemIndex = -1 Then
        Exit Sub
    End If
    
    If Button = vbKeyLButton Then
        m_selectedIndex = itemIndex
        UserControl_Paint
        RaiseEvent colourSelected(itemIndex)
    ElseIf Button = vbKeyRButton Then
        If m_allowPaletteChange Then
            changeColour itemIndex
        End If
    End If
End Sub

Private Sub UserControl_Paint()
    Dim backBuffer As Long
    Dim backBitmap As Long
    Dim oldBitmap As Long
    
    backBuffer = CreateCompatibleDC(UserControl.hdc)
    backBitmap = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight)
    
    oldBitmap = SelectObject(backBuffer, backBitmap)
    
    FillRect backBuffer, makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
    
    Dim itemRect As RECT
    Dim colourBrush As Long
    Dim oldBrush As Long
    Dim count As Long
    
    Dim coordY As Long
    Dim coordX As Long
    
    coordY = MARGIN_Y
    coordX = MARGIN_X
    
    For count = 0 To UBound(m_palette)
        itemRect.left = coordX
        itemRect.right = coordX + ITEM_SIZE
        itemRect.top = coordY
        itemRect.bottom = coordY + ITEM_SIZE
        
        colourBrush = CreateSolidBrush(m_palette(count))
        FillRect backBuffer, itemRect, colourBrush
        DeleteObject colourBrush
        
        FrameRect backBuffer, itemRect, colourManager.getBrush(SWIFTCOLOUR_CONTROLFORE)
        
        If count = m_selectedIndex Then
            itemRect.left = itemRect.left + 1
            itemRect.right = itemRect.right - 1
            itemRect.top = itemRect.top + 1
            itemRect.bottom = itemRect.bottom - 1
            
            DrawFocusRect backBuffer, itemRect
        End If
        
        coordX = coordX + ITEM_SIZE + SPACING_X
        
        If coordX + ITEM_SIZE > UserControl.ScaleWidth - MARGIN_X Then
            coordX = MARGIN_X
            coordY = coordY + ITEM_SIZE + SPACING_Y
        End If
    Next count
    
    BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, backBuffer, 0, 0, vbSrcCopy
        
    SelectObject backBuffer, oldBitmap
        
    DeleteDC backBuffer
    DeleteObject backBitmap
End Sub

Private Sub changeColour(index As Long)
    Dim cc As ChooseColorStruct
    
    cc.lStructSize = Len(cc)
    cc.hwndOwner = UserControl.hwnd
    cc.lpCustColors = VarPtr(custChooseColorPalette(0))
    
    If ChooseColor(cc) <> 0 Then
        m_palette(index) = cc.rgbResult
        UserControl_Paint
        RaiseEvent colourChanged(index, m_palette(index))
    End If
End Sub

