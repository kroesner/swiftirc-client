VERSION 5.00
Begin VB.Form frmIgnoreEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit ignore"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6690
   Icon            =   "frmIgnoreEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   242
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   446
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmIgnoreEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_fieldAddress As swiftIrc.ctlField
Private m_ignorePrivate As VB.CheckBox
Private m_ignoreChannel As VB.CheckBox
Private m_ignoreNotice As VB.CheckBox
Private m_ignoreCtcp As VB.CheckBox
Private m_ignoreInvite As VB.CheckBox
Private m_ignoreCodes As VB.CheckBox
Private WithEvents m_ignoreExclude As VB.CheckBox
Attribute m_ignoreExclude.VB_VarHelpID = -1
Private WithEvents m_ignoreAll As VB.CheckBox
Attribute m_ignoreAll.VB_VarHelpID = -1

Private m_session As CSession

Private WithEvents m_buttonOk As swiftIrc.ctlButton
Attribute m_buttonOk.VB_VarHelpID = -1
Private WithEvents m_buttonCancel As swiftIrc.ctlButton
Attribute m_buttonCancel.VB_VarHelpID = -1

Private m_cancelled As Boolean

Private m_labelManager As New CLabelManager

Public Property Let session(newValue As CSession)
10        Set m_session = newValue
End Property

Public Property Get address() As String
10        If Not swiftMatch("*!*@*", m_fieldAddress.value) Then
20            If Not m_session Is Nothing Then
30                address = m_session.getIal(m_fieldAddress.value, ialHost)
40            Else
50                address = m_fieldAddress.value & "!*@*"
60            End If
70        Else
80            address = m_fieldAddress.value
90        End If
End Property

Public Property Get cancelled() As Boolean
10        cancelled = m_cancelled
End Property

Public Property Get flags() As Long
10        If m_ignoreExclude.value Then
20            flags = IGNORE_EXCLUDE
30            Exit Property
40        End If

50        If m_ignoreAll.value Then
60            flags = IGNORE_ALL
70            Exit Property
80        End If

90        If m_ignorePrivate.value Then
100           flags = flags Or IGNORE_PRIVATE
110       End If
          
120       If m_ignoreChannel.value Then
130           flags = flags Or IGNORE_CHANNEL
140       End If
          
150       If m_ignoreNotice.value Then
160           flags = flags Or IGNORE_NOTICE
170       End If
          
180       If m_ignoreCtcp.value Then
190           flags = flags Or IGNORE_CTCP
200       End If
          
210       If m_ignoreInvite.value Then
220           flags = flags Or IGNORE_INVITE
230       End If
          
240       If m_ignoreCodes.value Then
250           flags = flags Or IGNORE_CODES
260       End If
End Property

Private Sub Form_Initialize()
10        Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
20        initControls
30        updateColours Controls
End Sub

Public Sub loadIgnore(ignore As CIgnoreItem)
10        m_fieldAddress.value = ignore.mask
          
20        If ignore.flags And IGNORE_EXCLUDE Then
30            m_ignoreExclude.value = 1
40            Exit Sub
50        End If
          
60        If (ignore.flags And IGNORE_ALL) = IGNORE_ALL Then
70            m_ignoreAll.value = 1
80            Exit Sub
90        Else
100           m_ignoreAll.value = 0
110       End If
          
120       If ignore.flags And IGNORE_PRIVATE Then
130           m_ignorePrivate.value = 1
140       End If
          
150       If ignore.flags And IGNORE_CHANNEL Then
160           m_ignoreChannel.value = 1
170       End If
          
180       If ignore.flags And IGNORE_NOTICE Then
190           m_ignoreNotice.value = 1
200       End If
          
210       If ignore.flags And IGNORE_CTCP Then
220           m_ignoreCtcp.value = 1
230       End If
          
240       If ignore.flags And IGNORE_INVITE Then
250           m_ignoreInvite.value = 1
260       End If
          
270       If ignore.flags And IGNORE_CODES Then
280           m_ignoreCodes.value = 1
290       End If
End Sub

Private Sub initControls()
10        Set m_fieldAddress = addField(Controls, "Nickname or mask:", 20, 20, Me.ScaleWidth - 40, 20)
          Dim fieldWidth
          
20        fieldWidth = Me.ScaleWidth - 40
          
30        m_fieldAddress.setFieldWidth (fieldWidth / 5) * 2, (fieldWidth / 5) * 3
          
40        m_labelManager.addLabel "Ignore options", ltSubHeading, 20, 50
          
50        Set m_ignoreAll = addCheckBox(Controls, "&All", 20, 70, 100, 20)
          
60        Set m_ignorePrivate = addCheckBox(Controls, "&Private", 20, 100, 100, 20)
70        Set m_ignoreChannel = addCheckBox(Controls, "C&hannel", 20, 120, 100, 20)
80        Set m_ignoreNotice = addCheckBox(Controls, "&Notice", 20, 140, 100, 20)
90        Set m_ignoreCtcp = addCheckBox(Controls, "C&TCP", 20, 160, 100, 20)
100       Set m_ignoreInvite = addCheckBox(Controls, "&Invite", 125, 100, 100, 20)
110       Set m_ignoreCodes = addCheckBox(Controls, "C&odes", 125, 120, 100, 20)
120       Set m_ignoreExclude = addCheckBox(Controls, "&Whitelist", 125, 140, 100, 20)
          
130       Set m_buttonOk = addButton(Controls, "&OK", Me.ScaleWidth - 185, Me.ScaleHeight - 40, 75, 20)
140       Set m_buttonCancel = addButton(Controls, "&Cancel", Me.ScaleWidth - 105, Me.ScaleHeight - 40, 75, 20)
          
150       m_ignoreAll.value = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
10        If KeyCode = vbKeyEscape Then
20            m_buttonCancel_clicked
30        End If
End Sub

Private Sub Form_Paint()
          Dim oldBrush As Long
          
10        oldBrush = SelectObject(Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK))
20        RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
30        SelectObject Me.hdc, oldBrush
          
40        SetBkColor Me.hdc, colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
50        m_labelManager.renderLabels Me.hdc
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
10        m_cancelled = True
End Sub

Private Sub m_buttonCancel_clicked()
10        m_cancelled = True
20        Me.Hide
End Sub

Private Sub m_buttonOk_clicked()
10        If LenB(m_fieldAddress.value) = 0 Then
20            MsgBox "Please enter a nickname or address to ignore.", vbCritical, "Missing information"
30            Exit Sub
40        End If
          
50        Me.Hide
End Sub

Private Sub m_ignoreAll_Click()
10        If m_ignoreAll.value = 1 Then
20            m_ignorePrivate.enabled = False
30            m_ignoreChannel.enabled = False
40            m_ignoreNotice.enabled = False
50            m_ignoreCtcp.enabled = False
60            m_ignoreInvite.enabled = False
70            m_ignoreCodes.enabled = False
80            m_ignoreExclude.enabled = False
              
90            m_ignorePrivate.value = 1
100           m_ignoreChannel.value = 1
110           m_ignoreNotice.value = 1
120           m_ignoreCtcp.value = 1
130           m_ignoreInvite.value = 1
140           m_ignoreCodes.value = 1
150           m_ignoreExclude.value = 0
160       Else
170           m_ignorePrivate.enabled = True
180           m_ignoreChannel.enabled = True
190           m_ignoreNotice.enabled = True
200           m_ignoreCtcp.enabled = True
210           m_ignoreInvite.enabled = True
220           m_ignoreCodes.enabled = True
230           m_ignoreExclude.enabled = True
              
240           m_ignorePrivate.value = 0
250           m_ignoreChannel.value = 0
260           m_ignoreNotice.value = 0
270           m_ignoreCtcp.value = 0
280           m_ignoreInvite.value = 0
290           m_ignoreCodes.value = 0
300       End If
End Sub

Private Sub m_ignoreExclude_Click()
10        If m_ignoreExclude.value = 1 Then
20            m_ignoreAll.enabled = False
30            m_ignorePrivate.enabled = False
40            m_ignoreChannel.enabled = False
50            m_ignoreNotice.enabled = False
60            m_ignoreCtcp.enabled = False
70            m_ignoreInvite.enabled = False
80            m_ignoreCodes.enabled = False
              
90            m_ignoreAll.value = 0
100           m_ignorePrivate.value = 0
110           m_ignoreChannel.value = 0
120           m_ignoreNotice.value = 0
130           m_ignoreCtcp.value = 0
140           m_ignoreInvite.value = 0
150           m_ignoreCodes.value = 0
160       Else
170           m_ignoreAll.enabled = True
180           m_ignorePrivate.enabled = True
190           m_ignoreChannel.enabled = True
200           m_ignoreNotice.enabled = True
210           m_ignoreCtcp.enabled = True
220           m_ignoreInvite.enabled = True
230           m_ignoreCodes.enabled = True
240       End If
End Sub
