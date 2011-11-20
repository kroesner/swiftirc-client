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
Private m_fontManager As CFontManager

Private m_textviewTopicPreview As swiftIrc.ctlTextView

Private WithEvents m_buttonBold As swiftIrc.ctlButton
Attribute m_buttonBold.VB_VarHelpID = -1
Private WithEvents m_buttonItalic As swiftIrc.ctlButton
Attribute m_buttonItalic.VB_VarHelpID = -1
Private WithEvents m_buttonUnderline As swiftIrc.ctlButton
Attribute m_buttonUnderline.VB_VarHelpID = -1

Private m_channel As CChannel

Public Property Get topic() As String
10        topic = textTopicBuilder.text
End Property

Public Property Let channel(newValue As CChannel)
10        Set m_channel = newValue
20        textTopicBuilder.text = m_channel.topic
          
          Dim count As Long
          
30        For count = 1 To m_channel.getTopicHistoryCount
40            comboTopicHistory.addItem m_channel.getTopicHistory(count)
50        Next count
End Property

Private Sub comboTopicHistory_Click()
10        If comboTopicHistory.ListIndex <> -1 Then
20            textTopicBuilder.text = comboTopicHistory.list(comboTopicHistory.ListIndex)
30        End If
End Sub

Private Sub IColourUser_coloursUpdated()
10        UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
20        updateColours Controls
          
          Dim colourUser As IColourUser
          
30        Set colourUser = m_textviewTopicPreview
40        colourUser.coloursUpdated
End Sub

Private Property Let IFontUser_fontManager(RHS As CFontManager)
10        Set m_fontManager = RHS
          
          Dim fontUser As IFontUser
          
20        Set fontUser = m_textviewTopicPreview
30        fontUser.fontManager = RHS
40        fontUser.fontsUpdated
End Property

Private Sub IFontUser_fontsUpdated()
          Dim fontUser As IFontUser
          
10        Set fontUser = m_textviewTopicPreview
20        fontUser.fontsUpdated
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Sub initControls()
10        m_labelManager.addLabel "Topic preview", ltSubHeading, 10, 10
20        m_labelManager.addLabel "Topic history", ltSubHeading, 10, 115
30        m_labelManager.addLabel "Topic builder", ltSubHeading, 10, 165
          
40        Set m_textviewTopicPreview = createControl(Controls, "swiftIrc.ctlTextView", "topicPreview")
          
50        Set m_buttonBold = addButton(Controls, "B", 10, 265, 25, 25)
60        Set m_buttonItalic = addButton(Controls, "I", 40, 265, 25, 25)
70        Set m_buttonUnderline = addButton(Controls, "U", 70, 265, 25, 25)
End Sub

Private Sub m_buttonBold_clicked()
10        InsertSymbol Chr$(2)
20        textTopicBuilder.setFocus
End Sub

Private Sub m_buttonItalic_clicked()
10        InsertSymbol Chr$(4)
20        textTopicBuilder.setFocus
End Sub

Private Sub m_buttonUnderline_clicked()
10        InsertSymbol Chr$(31)
20        textTopicBuilder.setFocus
End Sub

Private Sub textTopicBuilder_Change()
10        m_textviewTopicPreview.clear
          
20        If LenB(textTopicBuilder.text) <> 0 Then
30            m_textviewTopicPreview.addRawTextEx eventColours.topicChange, 0, "$0", Nothing, _
                  vbNullString, 0, makeStringArray(textTopicBuilder.text)
40        End If
End Sub

Private Sub textTopicBuilder_KeyPress(KeyAscii As Integer)
10        Select Case KeyAscii
              Case vbKeyReturn
20                KeyAscii = 0
30                Beep
40            Case 2
50                InsertSymbol Chr$(2)
60                KeyAscii = 0
70            Case 11
80                InsertSymbol Chr$(3)
90                KeyAscii = 0
100           Case 21
110               InsertSymbol Chr$(31)
120               KeyAscii = 0
130           Case 18
140               InsertSymbol Chr$(22)
150               KeyAscii = 0
160           Case 15
170               InsertSymbol Chr$(15)
180               KeyAscii = 0
190       End Select
End Sub

Private Sub UserControl_Initialize()
10        initControls
End Sub

Private Sub UserControl_Paint()
10        m_labelManager.renderLabels UserControl.hdc
End Sub

Private Sub UserControl_Resize()
10        getRealWindow(m_textviewTopicPreview).Move 10, 30, UserControl.ScaleWidth - 20, 75
20        comboTopicHistory.left = 10
30        comboTopicHistory.top = 135
40        comboTopicHistory.width = UserControl.ScaleWidth - 20
50        textTopicBuilder.Move 10, 185, UserControl.ScaleWidth - 20, 75
End Sub

Private Sub InsertSymbol(symbol As String)
10        If textTopicBuilder.SelLength > 0 Then
20            textTopicBuilder.SelText = symbol & textTopicBuilder.SelText & symbol
30            textTopicBuilder.selStart = textTopicBuilder.selStart + textTopicBuilder.SelLength
40            textTopicBuilder.SelLength = 0
50        Else
60            textTopicBuilder.SelText = symbol
70        End If
End Sub

