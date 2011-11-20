VERSION 5.00
Begin VB.Form frmErrorReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SwiftIRC - An error has occured"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5400
   Icon            =   "frmErrorReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   226
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtErrorLog 
      Height          =   1185
      Left            =   165
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   165
      Width           =   2820
   End
End
Attribute VB_Name = "frmErrorReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_buttonOk As swiftIrc.ctlButton
Attribute m_buttonOk.VB_VarHelpID = -1
Private WithEvents m_buttonCopy As swiftIrc.ctlButton
Attribute m_buttonCopy.VB_VarHelpID = -1
Private m_checkHideErrors As VB.CheckBox

Private Sub initControls()
10        Set m_buttonOk = addButton(Controls, "Ok", Me.ScaleWidth - 95, Me.ScaleHeight - 40, 75, 20)
20        Set m_buttonCopy = addButton(Controls, "Copy error to clipboard", Me.ScaleWidth - 250, Me.ScaleHeight - 40, 150, 20)
30        Set m_checkHideErrors = addCheckBox(Controls, "Hide further errors", Me.ScaleWidth - 145, Me.ScaleHeight - 70, 130, 25)
40        SendMessage txtErrorLog.hwnd, WM_SETFONT, g_fontUI, ByVal 0
End Sub

Private Sub Form_Load()
10        initControls
          
20        Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
30        updateColours Controls
          
40        If g_hideErrors Then
50            m_checkHideErrors.value = 1
60        End If
End Sub

Private Sub Form_Paint()
10        SelectObject Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK)
20        RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
          
          Dim icon As Long
          
30        icon = LoadImageAPtr(ByVal 0&, IDI_HAND, IMAGE_ICON, 0, 0, LR_SHARED)
40        DrawIcon Me.hdc, 20, 20, icon
50        DestroyIcon icon
          
60        SetBkMode Me.hdc, TRANSPARENT
70        SetTextColor Me.hdc, colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
80        SelectObject Me.hdc, g_fontUI
          
          Dim labelText As String
          Dim textRect As RECT
          
90        textRect.left = 65
100       textRect.right = Me.ScaleWidth - 20
110       textRect.top = 20
120       textRect.bottom = 100
          
130       labelText = "An error has occured in the SwiftIRC client.  " _
              & "The client may be able to continue running, but unexpected behavior may result." & vbCrLf & vbCrLf _
              & "Please report the following error text to help improve SwiftIRC."
              
140       swiftDrawText Me.hdc, labelText, VarPtr(textRect), DT_WORDBREAK
End Sub

Private Sub Form_Resize()
10        txtErrorLog.left = 20
20        txtErrorLog.width = Me.ScaleWidth - 40
30        txtErrorLog.top = 110
40        txtErrorLog.height = 50
End Sub

Private Sub m_buttonCopy_clicked()
10        Clipboard.clear
20        Clipboard.SetText txtErrorLog.text
End Sub

Private Sub m_buttonOk_clicked()
10        g_hideErrors = -m_checkHideErrors.value
20        Me.Hide
End Sub
