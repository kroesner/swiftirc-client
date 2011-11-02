VERSION 5.00
Begin VB.Form frmEditNicknameAppearance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit nickname appearance"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmEditNicknameAppearance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditNicknameAppearance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_buttonOk As swiftIrc.ctlButton
Attribute m_buttonOk.VB_VarHelpID = -1
Private WithEvents m_buttonCancel As swiftIrc.ctlButton
Attribute m_buttonCancel.VB_VarHelpID = -1
Private WithEvents m_checkMatchMyself As swiftIrc.ctlButton
Attribute m_checkMatchMyself.VB_VarHelpID = -1
Private WithEvents m_checkMatchMask As VB.CheckBox
Attribute m_checkMatchMask.VB_VarHelpID = -1
Private WithEvents m_checkMatchModes As VB.CheckBox
Attribute m_checkMatchModes.VB_VarHelpID = -1
Private WithEvents m_checkColour As VB.CheckBox
Attribute m_checkColour.VB_VarHelpID = -1
Private WithEvents m_checkIcon As VB.CheckBox
Attribute m_checkIcon.VB_VarHelpID = -1
Private WithEvents m_fieldMask As swiftIrc.ctlField
Attribute m_fieldMask.VB_VarHelpID = -1
Private WithEvents m_fieldModes As swiftIrc.ctlField
Attribute m_fieldModes.VB_VarHelpID = -1
Private m_colourSelector As swiftIrc.ctlSingleColourSelector
Attribute m_colourSelector.VB_VarHelpID = -1
Private m_labelManager As New CLabelManager

Private Sub initControls()
    m_labelManager.addLabel "Criteria", ltHeading, 20, 20
    
    Set m_buttonCancel = addButton(Controls, "&Cancel", Me.ScaleWidth - 95, Me.ScaleHeight - 40, 75, 20)
    Set m_buttonOk = addButton(Controls, "&OK", Me.ScaleWidth - 175, Me.ScaleHeight - 40, 75, 20)
    
    set m_checkmatchmysql = addcheckbox(controls,
    Set m_checkMatchMask = addCheckBox(Controls, "Match nick!user@host mask", 20, 40, 190, 20)
    Set m_checkMatchModes = addCheckBox(Controls, "Match channel modes", 20, 90, 190, 20)
    'Set m_checkColour = addCheckBox(Controls, "Different colour", 20, 140, 100, 20)
    
    Set m_fieldMask = addField(Controls, "Mask:", 20, 65, 190, 20)
    Set m_fieldModes = addField(Controls, "Modes:", 20, 115, 190, 20)
    
    m_fieldMask.setFieldWidth 50, 140
    m_fieldModes.setFieldWidth 50, 140
    
    m_labelManager.addLabel "Appearance", ltHeading, 20, 140
    
    'Set m_colourSelector = createControl(Controls, "swiftIrc.ctlSingleColourSelector", "colourSelector")
    'getRealWindow(m_colourSelector).Move 130, 150, 20, 20
    
   ' m_colourSelector.setPalette colourThemes.currentSettingsTheme.getPalette()
    
    'm_labelManager.addLabel "Nickname colour:", ltNormal, 20, 150
End Sub

Private Sub Form_Initialize()
    initControls
    updateColours Controls
End Sub

Private Sub Form_Load()
    Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
End Sub

Private Sub Form_Paint()
    Dim oldBrush As Long

    SetBkColor Me.hdc, colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    oldBrush = SelectObject(Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK))
    RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
    SelectObject Me.hdc, oldBrush
    
    m_labelManager.renderLabels Me.hdc
End Sub
