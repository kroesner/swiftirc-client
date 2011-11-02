VERSION 5.00
Begin VB.Form frmColourSelector 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select colour"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   65
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmColourSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_colourPalette As ctlColourPalette
Attribute m_colourPalette.VB_VarHelpID = -1
Private m_palette() As Long
Private m_selectedColour As Long

Public Property Get selectedColour() As Long
    selectedColour = m_selectedColour
End Property

Public Sub setPalette(newPalette() As Long)
    m_palette = newPalette
End Sub

Private Sub Form_Load()
    m_selectedColour = -1

    Set m_colourPalette = createControl(Controls, "swiftIrc.ctlColourPalette", "palette")
    m_colourPalette.setPalette m_palette
    Me.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    MoveWindow Me.hwnd, Me.left, Me.top, 200, 75, 1
End Sub

Private Sub Form_Resize()
    getRealWindow(m_colourPalette).Move 0, 0, 200, 50
End Sub

Private Sub m_colourPalette_colourSelected(index As Long)
    m_selectedColour = index
    Me.Hide
End Sub
