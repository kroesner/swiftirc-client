VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   4065
   ClientTop       =   5670
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   571
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   705
      TabIndex        =   0
      Top             =   3510
      Width           =   525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_fontManager As New CFontManager
Private m_textView As ctlTextView
Private m_textViewControl As VBControlExtender
Private WithEvents m_button As swiftIrc.ctlButton
Attribute m_button.VB_VarHelpID = -1
Private m_buttonControl As VBControlExtender

Private a As Integer

Private Sub Form_Load()
    Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
    m_fontManager.changeFont Me.hdc, "Segoe UI", 10

    Set m_buttonControl = Controls.Add("swiftirc.ctlWindowStart", "start")
    m_buttonControl.Move 0, 0, 500, 500
    m_buttonControl.visible = True
End Sub

Private Sub m_button_clicked()
    a = a + 1
    Debug.Print a
End Sub
