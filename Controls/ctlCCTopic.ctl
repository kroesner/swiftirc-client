VERSION 5.00
Begin VB.UserControl ctlCCTopic 
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   250
   Begin VB.TextBox textTopicBuilder 
      Appearance      =   0  'Flat
      Height          =   765
      Left            =   585
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "ctlCCTopic.ctx":0000
      Top             =   1440
      Width           =   2235
   End
   Begin VB.ComboBox comboTopicHistory 
      Height          =   315
      Left            =   585
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   825
      Width           =   2220
   End
End
Attribute VB_Name = "ctlCCTopic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow
Implements IColourUser
Implements IFontUser

Private m_realWindow As VBControlExtender
Private m_labelManager As New CLabelManager
Private m_fontmanager As CFontManager

Private m_textviewTopicPreview As swiftIrc.ctlTextView

Private WithEvents m_buttonBold As swiftIrc.ctlButton
Attribute m_buttonBold.VB_VarHelpID = -1
Private WithEvents m_buttonItalic As swiftIrc.ctlButton
Attribute m_buttonItalic.VB_VarHelpID = -1
Private WithEvents m_buttonUnderline As swiftIrc.ctlButton
Attribute m_buttonUnderline.VB_VarHelpID = -1

Private m_channel As CChannel

Public Property Get topic() As String
    topic = textTopicBuilder.text
End Property

Public Property Let channel(newValue As CChannel)
    Set m_channel = newValue
    textTopicBuilder.text = m_channel.topic
    
    Dim count As Long
    
    For count = 1 To m_channel.getTopicHistoryCount
        comboTopicHistory.addItem m_channel.getTopicHistory(count)
    Next count
End Property

Private Sub comboTopicHistory_Click()
    If comboTopicHistory.ListIndex <> -1 Then
        textTopicBuilder.text = comboTopicHistory.list(comboTopicHistory.ListIndex)
    End If
End Sub

Private Sub IColourUser_coloursUpdated()
    UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    updateColours Controls
    
    Dim colourUser As IColourUser
    
    Set colourUser = m_textviewTopicPreview
    colourUser.coloursUpdated
End Sub

Private Property Let IFontUser_fontManager(RHS As CFontManager)
    Set m_fontmanager = RHS
    
    Dim fontUser As IFontUser
    
    Set fontUser = m_textviewTopicPreview
    fontUser.fontManager = RHS
    fontUser.fontsUpdated
End Property

Private Sub IFontUser_fontsUpdated()
    Dim fontUser As IFontUser
    
    Set fontUser = m_textviewTopicPreview
    fontUser.fontsUpdated
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Sub initControls()
    m_labelManager.addLabel "Topic preview", ltSubHeading, 10, 10
    m_labelManager.addLabel "Topic history", ltSubHeading, 10, 115
    m_labelManager.addLabel "Topic builder", ltSubHeading, 10, 165
    
    Set m_textviewTopicPreview = createControl(Controls, "swiftIrc.ctlTextView", "topicPreview")
    
    Set m_buttonBold = addButton(Controls, "B", 10, 265, 25, 25)
    Set m_buttonItalic = addButton(Controls, "I", 40, 265, 25, 25)
    Set m_buttonUnderline = addButton(Controls, "U", 70, 265, 25, 25)
End Sub

Private Sub m_buttonBold_clicked()
    InsertSymbol Chr$(2)
    textTopicBuilder.setFocus
End Sub

Private Sub m_buttonItalic_clicked()
    InsertSymbol Chr$(4)
    textTopicBuilder.setFocus
End Sub

Private Sub m_buttonUnderline_clicked()
    InsertSymbol Chr$(31)
    textTopicBuilder.setFocus
End Sub

Private Sub textTopicBuilder_Change()
    m_textviewTopicPreview.clear
    
    If LenB(textTopicBuilder.text) <> 0 Then
        m_textviewTopicPreview.addRawTextEx eventColours.topicChange, 0, "$0", Nothing, _
            vbNullString, 0, makeStringArray(textTopicBuilder.text)
    End If
End Sub

Private Sub textTopicBuilder_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            Beep
        Case 2
            InsertSymbol Chr$(2)
            KeyAscii = 0
        Case 11
            InsertSymbol Chr$(3)
            KeyAscii = 0
        Case 21
            InsertSymbol Chr$(31)
            KeyAscii = 0
        Case 18
            InsertSymbol Chr$(22)
            KeyAscii = 0
        Case 15
            InsertSymbol Chr$(15)
            KeyAscii = 0
    End Select
End Sub

Private Sub UserControl_Initialize()
    initControls
End Sub

Private Sub UserControl_Paint()
    m_labelManager.renderLabels UserControl.hdc
End Sub

Private Sub UserControl_Resize()
    getRealWindow(m_textviewTopicPreview).Move 10, 30, UserControl.ScaleWidth - 20, 75
    comboTopicHistory.left = 10
    comboTopicHistory.top = 135
    comboTopicHistory.width = UserControl.ScaleWidth - 20
    textTopicBuilder.Move 10, 185, UserControl.ScaleWidth - 20, 75
End Sub

Private Sub InsertSymbol(symbol As String)
    If textTopicBuilder.SelLength > 0 Then
        textTopicBuilder.SelText = symbol & textTopicBuilder.SelText & symbol
        textTopicBuilder.selStart = textTopicBuilder.selStart + textTopicBuilder.SelLength
        textTopicBuilder.SelLength = 0
    Else
        textTopicBuilder.SelText = symbol
    End If
End Sub

