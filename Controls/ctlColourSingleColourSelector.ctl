VERSION 5.00
Begin VB.UserControl ctlSingleColourSelector 
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   360
   ScaleHeight     =   23
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   24
End
Attribute VB_Name = "ctlSingleColourSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow

Private m_realWindow As VBControlExtender

Private m_colour As Long
Private m_gotPalette As Boolean
Private m_palette() As Long

Public Event colourChanged()

Public Sub setPalette(newPalette() As Long)
    m_palette = newPalette
    m_gotPalette = True
    UserControl_Paint
End Sub

Public Property Get colour() As Long
    colour = m_colour
End Property

Public Property Let colour(newValue As Long)
    m_colour = newValue
    UserControl_Paint
End Property

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Sub UserControl_Click()
    Dim frmColourSelector As New frmColourSelector
    
    frmColourSelector.setPalette m_palette
    frmColourSelector.Show vbModal, Me
    
    If frmColourSelector.selectedColour <> -1 Then
        m_colour = frmColourSelector.selectedColour
        RaiseEvent colourChanged
        UserControl_Paint
    End If
    
    Unload frmColourSelector
End Sub

Private Sub UserControl_Paint()
    If Not m_gotPalette Then
        Exit Sub
    End If

    Dim brush As Long
    
    brush = CreateSolidBrush(m_palette(m_colour))

    Dim controlRect As RECT

    controlRect = makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight)

    FillRect UserControl.hdc, controlRect, brush
    FrameRect UserControl.hdc, controlRect, colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
    
    DeleteObject brush
End Sub

Private Sub UserControl_Terminate()
    debugLog "ctlSingleColourSelector terminating"
End Sub
