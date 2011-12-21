VERSION 5.00
Begin VB.Form frmEditHighlight 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit highlight"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3930
   Icon            =   "frmEditHighlight.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   194
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   262
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox comboHighlightType 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   315
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   210
      Width           =   1275
   End
End
Attribute VB_Name = "frmEditHighlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_comboList As VB.ComboBox
Private m_fieldHighlight As swiftIrc.ctlField
Private m_colourPalette As swiftIrc.ctlColourPalette
Private m_labelManager As New CLabelManager
Private m_buttonSave As swiftIrc.ctlButton
Private m_buttonCancel As swiftIrc.ctlButton

Private m_highlight As CHighlight

Private Sub initControls()
    Set m_fieldHighlight = addField(Controls, "Highlight text:", 20, 20, Me.ScaleWidth - 40, 20)
    m_labelManager.addLabel "Match:", ltSubHeading, 20, 50
    m_labelManager.addLabel "Colour:", ltSubHeading, 20, 80
    
    Set m_colourPalette = createControl(Controls, "swiftIrc.ctlColourPalette", "palette")
    
    Set m_buttonSave = addButton(Controls, "&Save", Me.ScaleWidth - 165, Me.ScaleHeight - 40, 70, 20)
    Set m_buttonCancel = addButton(Controls, "&Cancel", Me.ScaleWidth - 90, Me.ScaleHeight - 40, 70, 20)
    
    getRealWindow(m_colourPalette).Move 20, 95, Me.ScaleWidth - 40, 50
    m_colourPalette.setPalette colourThemes.currentTheme.getPalette
End Sub

Private Sub Form_Load()
    Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
    SetBkColor Me.hdc, colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    initControls
    updateColours Controls
End Sub

Private Sub Form_Paint()
    SelectObject Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
    RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
    
    m_labelManager.renderLabels Me.hdc
End Sub

Private Sub Form_Resize()
    comboHighlightType.top = 50
    comboHighlightType.left = Me.ScaleWidth - (Me.ScaleWidth / 2)
    comboHighlightType.width = ((Me.ScaleWidth - 40) / 2)
End Sub
