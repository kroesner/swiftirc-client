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

Private m_fontManager As CFontManager

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
10        Set m_channel = newValue
20        m_ctlTopic.channel = m_channel
30        m_ctlBans.channel = m_channel
40        m_ctlModes.channel = m_channel
End Property

Private Sub initControls()
10        Set m_buttonOk = addButton(Controls, "Ok", Me.ScaleWidth - 170, Me.ScaleHeight - 30, 50, 20)
20        Set m_buttonCancel = addButton(Controls, "Cancel", Me.ScaleWidth - 115, Me.ScaleHeight - 30, 50, _
              20)
30        Set m_buttonApply = addButton(Controls, "Apply", Me.ScaleWidth - 60, Me.ScaleHeight - 30, 50, _
              20)
          
40        Set m_tabStrip = createControl(Controls, "swiftIrc.ctlTabStrip", "tabStrip")
50        getRealWindow(m_tabStrip).Move 10, 0, Me.ScaleWidth - 20, 30
60        getRealWindow(m_tabStrip).visible = True
          
70        Set m_tabTopic = m_tabStrip.addTab("Topic")
80        Set m_tabBans = m_tabStrip.addTab("Bans")
90        Set m_tabModes = m_tabStrip.addTab("Modes")
          
100       Set m_ctlTopic = createControl(Controls, "swiftIrc.ctlCCTopic", "ctlTopic")
110       Set m_ctlBans = createControl(Controls, "swiftIrc.ctlCCBans", "ctlBans")
120       Set m_ctlModes = createControl(Controls, "swiftIrc.ctlCCModes", "ctlModes")

130       getRealWindow(m_ctlTopic).Move 10, 30, Me.ScaleWidth - 20, Me.ScaleHeight - 65
140       getRealWindow(m_ctlBans).Move 10, 30, Me.ScaleWidth - 20, Me.ScaleHeight - 65
150       getRealWindow(m_ctlModes).Move 10, 30, Me.ScaleWidth - 20, Me.ScaleHeight - 65
          
160       m_tabStrip.selectTab m_tabTopic, False
End Sub

Private Sub colourUpdate()
10        Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
20        updateColours Controls
End Sub

Private Sub Form_Initialize()
10        initControls
20        colourUpdate
End Sub

Private Property Let IFontUser_fontManager(RHS As CFontManager)
10        Set m_fontManager = RHS
          
          Dim fontUser As IFontUser
          
20        Set fontUser = m_ctlTopic
30        fontUser.fontManager = m_fontManager
40        fontUser.fontsUpdated
End Property

Private Sub IFontUser_fontsUpdated()

End Sub

Private Sub m_buttonApply_clicked()
10        applyChanges
End Sub

Private Sub m_buttonCancel_clicked()
10        Me.Hide
End Sub

Private Sub m_buttonOk_clicked()
10        applyChanges
20        Me.Hide
End Sub

Private Sub m_tabStrip_tabSelected(selectedTab As CTabStripItem)
10        If selectedTab Is m_tabTopic Then
20            getRealWindow(m_ctlTopic).visible = True
30            getRealWindow(m_ctlBans).visible = False
40            getRealWindow(m_ctlModes).visible = False
50        ElseIf selectedTab Is m_tabBans Then
60            getRealWindow(m_ctlBans).visible = True
70            getRealWindow(m_ctlTopic).visible = False
80            getRealWindow(m_ctlModes).visible = False
90        ElseIf selectedTab Is m_tabModes Then
100           getRealWindow(m_ctlModes).visible = True
110           getRealWindow(m_ctlTopic).visible = False
120           getRealWindow(m_ctlBans).visible = False
130       End If
End Sub

Private Sub applyChanges()
10        If StrComp(m_ctlTopic.topic, m_channel.topic, vbBinaryCompare) <> 0 Then
20            m_channel.session.sendLine "TOPIC " & m_channel.name & " :" & m_ctlTopic.topic
30        End If
          
40        m_ctlModes.applyModes
End Sub
