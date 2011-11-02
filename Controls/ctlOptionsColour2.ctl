VERSION 5.00
Begin VB.UserControl ctlOptionsColour2 
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7005
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   467
End
Attribute VB_Name = "ctlOptionsColour2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IWindow

Private m_realWindow As VBControlExtender
Private m_nicknameStyles As ctlNicknameStyleList
Private m_client As SwiftIrcClient

Public Property Let client(newValue As SwiftIrcClient)
    Set m_client = newValue
    m_nicknameStyles.client = m_client
End Property

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Sub UserControl_Initialize()
    UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    initControls
End Sub

Private Sub initControls()
    Set m_nicknameStyles = createControl(Controls, "swiftIrc.ctlNicknameStyleList", "nickstyle")
    getRealWindow(m_nicknameStyles).Move 5, 5, 300, 175
End Sub

Friend Sub saveSettings()
    m_nicknameStyles.saveSettings
End Sub
