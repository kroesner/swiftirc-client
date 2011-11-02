VERSION 5.00
Begin VB.Form frmTestContainer 
   Caption         =   "Test container"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   484
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   584
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmTestContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_irc As SwiftIrcClient
Attribute m_irc.VB_VarHelpID = -1
Private m_ircControl As VBControlExtender

Private WithEvents m_newClient As VB.CommandButton
Attribute m_newClient.VB_VarHelpID = -1

Private Sub Form_Load()
    Set m_ircControl = Controls.Add("swiftirc.swiftircclient", "irc")
    Set m_irc = m_ircControl.object
    
    m_irc.debugEx = True
    m_ircControl.visible = True
    m_irc.userPath = "SwiftIRC User Data\"
    m_irc.assetPath = "assets\"
    m_irc.init
    
    Set m_newClient = Controls.Add("VB.CommandButton", "newclient")
    m_newClient.visible = True
    m_newClient.Move 0, 0, 75, 20
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_irc.deInit
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        m_ircControl.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    End If
End Sub

Private Sub m_irc_visitUrl(url As String)
    MsgBox url
End Sub

Private Sub m_newClient_Click()
    MsgBox m_irc.getVersion
End Sub
