VERSION 5.00
Begin VB.Form frmChannelCentral 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Channel central"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6045
   Icon            =   "frmChannelCentral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   403
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmChannelCentral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IFontUser

Private m_fontmanager As CFontManager

Private WithEvents m_buttonOk As ctlButton
Attribute m_buttonOk.VB_VarHelpID = -1
Private WithEvents m_buttonCancel As ctlButton
Attribute m_buttonCancel.VB_VarHelpID = -1
Private WithEvents m_buttonApply As ctlButton
Attribute m_buttonApply.VB_VarHelpID = -1
Private WithEvents m_tabStrip As ctlTabStrip
Attribute m_tabStrip.VB_VarHelpID = -1

Private m_ctlTopic As swiftIrc.ctlCCTopic
Private m_ctlBans As swiftIrc.ctlCCBans
Private m_ctlModes As swiftIrc.ctlCCModes

Private m_tabTopic As CTabStripItem
Private m_tabBans As CTabStripItem
Private m_tabModes As CTabStripItem

Private m_textviewTopicPreview As ctlTextView

Private m_channel As CChannel

Public Property Let channel(newValue As CChannel)
    Set m_channel = newValue
    m_ctlTopic.channel = m_channel
    m_ctlBans.channel = m_channel
    m_ctlModes.channel = m_channel
End Property

Private Sub initControls()
    Set m_buttonOk = addButton(Controls, "Ok", Me.ScaleWidth - 170, Me.ScaleHeight - 30, 50, 20)
    Set m_buttonCancel = addButton(Controls, "Cancel", Me.ScaleWidth - 115, Me.ScaleHeight - 30, 50, 20)
    Set m_buttonApply = addButton(Controls, "Apply", Me.ScaleWidth - 60, Me.ScaleHeight - 30, 50, 20)
    
    Set m_tabStrip = createControl(Controls, "swiftIrc.ctlTabStrip", "tabStrip")
    getRealWindow(m_tabStrip).Move 10, 0, Me.ScaleWidth - 20, 30
    getRealWindow(m_tabStrip).visible = True
    
    Set m_tabTopic = m_tabStrip.addTab("Topic")
    Set m_tabBans = m_tabStrip.addTab("Bans")
    Set m_tabModes = m_tabStrip.addTab("Modes")
    
    Set m_ctlTopic = createControl(Controls, "swiftIrc.ctlCCTopic", "ctlTopic")
    Set m_ctlBans = createControl(Controls, "swiftIrc.ctlCCBans", "ctlBans")
    Set m_ctlModes = createControl(Controls, "swiftIrc.ctlCCModes", "ctlModes")

    getRealWindow(m_ctlTopic).Move 10, 30, Me.ScaleWidth - 20, Me.ScaleHeight - 65
    getRealWindow(m_ctlBans).Move 10, 30, Me.ScaleWidth - 20, Me.ScaleHeight - 65
    getRealWindow(m_ctlModes).Move 10, 30, Me.ScaleWidth - 20, Me.ScaleHeight - 65
    
    m_tabStrip.selectTab m_tabTopic, False
End Sub

Private Sub colourUpdate()
    Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
    updateColours Controls
End Sub

Private Sub Form_Initialize()
    initControls
    colourUpdate
End Sub

Private Property Let IFontUser_fontManager(RHS As CFontManager)
    Set m_fontmanager = RHS
    
    Dim fontUser As IFontUser
    
    Set fontUser = m_ctlTopic
    fontUser.fontManager = m_fontmanager
    fontUser.fontsUpdated
End Property

Private Sub IFontUser_fontsUpdated()

End Sub

Private Sub m_buttonApply_clicked()
    applyChanges
End Sub

Private Sub m_buttonCancel_clicked()
    Me.Hide
End Sub

Private Sub m_buttonOk_clicked()
    applyChanges
    Me.Hide
End Sub

Private Sub m_tabStrip_tabSelected(selectedTab As CTabStripItem)
    If selectedTab Is m_tabTopic Then
        getRealWindow(m_ctlTopic).visible = True
        getRealWindow(m_ctlBans).visible = False
        getRealWindow(m_ctlModes).visible = False
    ElseIf selectedTab Is m_tabBans Then
        getRealWindow(m_ctlBans).visible = True
        getRealWindow(m_ctlTopic).visible = False
        getRealWindow(m_ctlModes).visible = False
    ElseIf selectedTab Is m_tabModes Then
        getRealWindow(m_ctlModes).visible = True
        getRealWindow(m_ctlTopic).visible = False
        getRealWindow(m_ctlBans).visible = False
    End If
End Sub

Private Sub applyChanges()
    If StrComp(m_ctlTopic.topic, m_channel.topic, vbBinaryCompare) <> 0 Then
        m_channel.session.sendLine "TOPIC " & m_channel.name & " :" & m_ctlTopic.topic
    End If
    
    m_ctlModes.applyModes
End Sub
