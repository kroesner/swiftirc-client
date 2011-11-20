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
10        m_palette = newPalette
20        m_gotPalette = True
30        UserControl_Paint
End Sub

Public Property Get colour() As Long
10        colour = m_colour
End Property

Public Property Let colour(newValue As Long)
10        m_colour = newValue
20        UserControl_Paint
End Property

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Sub UserControl_Click()
          Dim frmColourSelector As New frmColourSelector
          
10        frmColourSelector.setPalette m_palette
20        frmColourSelector.Show vbModal, Me
          
30        If frmColourSelector.selectedColour <> -1 Then
40            m_colour = frmColourSelector.selectedColour
50            RaiseEvent colourChanged
60            UserControl_Paint
70        End If
          
80        Unload frmColourSelector
End Sub

Private Sub UserControl_Paint()
10        If Not m_gotPalette Then
20            Exit Sub
30        End If

          Dim brush As Long
          
40        brush = CreateSolidBrush(m_palette(m_colour))

          Dim controlRect As RECT

50        controlRect = makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight)

60        FillRect UserControl.hdc, controlRect, brush
70        FrameRect UserControl.hdc, controlRect, colourManager.getBrush(SWIFTCOLOUR_CONTROLBORDER)
          
80        DeleteObject brush
End Sub

Private Sub UserControl_Terminate()
10        debugLog "ctlSingleColourSelector terminating"
End Sub
