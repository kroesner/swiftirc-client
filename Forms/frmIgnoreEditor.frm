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
    Set m_session = newValue
End Property

Public Property Get address() As String
    If Not swiftMatch("*!*@*", m_fieldAddress.value) Then
        If Not m_session Is Nothing Then
            address = m_session.getIal(m_fieldAddress.value, ialHost)
        Else
            address = m_fieldAddress.value & "!*@*"
        End If
    Else
        address = m_fieldAddress.value
    End If
End Property

Public Property Get cancelled() As Boolean
    cancelled = m_cancelled
End Property

Public Property Get flags() As Long
    If m_ignoreExclude.value Then
        flags = IGNORE_EXCLUDE
        Exit Property
    End If

    If m_ignoreAll.value Then
        flags = IGNORE_ALL
        Exit Property
    End If

    If m_ignorePrivate.value Then
        flags = flags Or IGNORE_PRIVATE
    End If
    
    If m_ignoreChannel.value Then
        flags = flags Or IGNORE_CHANNEL
    End If
    
    If m_ignoreNotice.value Then
        flags = flags Or IGNORE_NOTICE
    End If
    
    If m_ignoreCtcp.value Then
        flags = flags Or IGNORE_CTCP
    End If
    
    If m_ignoreInvite.value Then
        flags = flags Or IGNORE_INVITE
    End If
    
    If m_ignoreCodes.value Then
        flags = flags Or IGNORE_CODES
    End If
End Property

Private Sub Form_Initialize()
    Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
    initControls
    updateColours Controls
End Sub

Public Sub loadIgnore(ignore As CIgnoreItem)
    m_fieldAddress.value = ignore.mask
    
    If ignore.flags And IGNORE_EXCLUDE Then
        m_ignoreExclude.value = 1
        Exit Sub
    End If
    
    If (ignore.flags And IGNORE_ALL) = IGNORE_ALL Then
        m_ignoreAll.value = 1
        Exit Sub
    Else
        m_ignoreAll.value = 0
    End If
    
    If ignore.flags And IGNORE_PRIVATE Then
        m_ignorePrivate.value = 1
    End If
    
    If ignore.flags And IGNORE_CHANNEL Then
        m_ignoreChannel.value = 1
    End If
    
    If ignore.flags And IGNORE_NOTICE Then
        m_ignoreNotice.value = 1
    End If
    
    If ignore.flags And IGNORE_CTCP Then
        m_ignoreCtcp.value = 1
    End If
    
    If ignore.flags And IGNORE_INVITE Then
        m_ignoreInvite.value = 1
    End If
    
    If ignore.flags And IGNORE_CODES Then
        m_ignoreCodes.value = 1
    End If
End Sub

Private Sub initControls()
    Set m_fieldAddress = addField(Controls, "Nickname or mask:", 20, 20, Me.ScaleWidth - 40, 20)
    Dim fieldWidth
    
    fieldWidth = Me.ScaleWidth - 40
    
    m_fieldAddress.setFieldWidth (fieldWidth / 5) * 2, (fieldWidth / 5) * 3
    
    m_labelManager.addLabel "Ignore options", ltSubHeading, 20, 50
    
    Set m_ignoreAll = addCheckBox(Controls, "&All", 20, 70, 100, 20)
    
    Set m_ignorePrivate = addCheckBox(Controls, "&Private", 20, 100, 100, 20)
    Set m_ignoreChannel = addCheckBox(Controls, "C&hannel", 20, 120, 100, 20)
    Set m_ignoreNotice = addCheckBox(Controls, "&Notice", 20, 140, 100, 20)
    Set m_ignoreCtcp = addCheckBox(Controls, "C&TCP", 20, 160, 100, 20)
    Set m_ignoreInvite = addCheckBox(Controls, "&Invite", 125, 100, 100, 20)
    Set m_ignoreCodes = addCheckBox(Controls, "C&odes", 125, 120, 100, 20)
    Set m_ignoreExclude = addCheckBox(Controls, "&Whitelist", 125, 140, 100, 20)
    
    Set m_buttonOk = addButton(Controls, "&OK", Me.ScaleWidth - 185, Me.ScaleHeight - 40, 75, 20)
    Set m_buttonCancel = addButton(Controls, "&Cancel", Me.ScaleWidth - 105, Me.ScaleHeight - 40, 75, 20)
    
    m_ignoreAll.value = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        m_buttonCancel_clicked
    End If
End Sub

Private Sub Form_Paint()
    Dim oldBrush As Long
    
    oldBrush = SelectObject(Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK))
    RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
    SelectObject Me.hdc, oldBrush
    
    SetBkColor Me.hdc, colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    m_labelManager.renderLabels Me.hdc
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    m_cancelled = True
End Sub

Private Sub m_buttonCancel_clicked()
    m_cancelled = True
    Me.Hide
End Sub

Private Sub m_buttonOk_clicked()
    If LenB(m_fieldAddress.value) = 0 Then
        MsgBox "Please enter a nickname or address to ignore.", vbCritical, "Missing information"
        Exit Sub
    End If
    
    Me.Hide
End Sub

Private Sub m_ignoreAll_Click()
    If m_ignoreAll.value = 1 Then
        m_ignorePrivate.enabled = False
        m_ignoreChannel.enabled = False
        m_ignoreNotice.enabled = False
        m_ignoreCtcp.enabled = False
        m_ignoreInvite.enabled = False
        m_ignoreCodes.enabled = False
        m_ignoreExclude.enabled = False
        
        m_ignorePrivate.value = 1
        m_ignoreChannel.value = 1
        m_ignoreNotice.value = 1
        m_ignoreCtcp.value = 1
        m_ignoreInvite.value = 1
        m_ignoreCodes.value = 1
        m_ignoreExclude.value = 0
    Else
        m_ignorePrivate.enabled = True
        m_ignoreChannel.enabled = True
        m_ignoreNotice.enabled = True
        m_ignoreCtcp.enabled = True
        m_ignoreInvite.enabled = True
        m_ignoreCodes.enabled = True
        m_ignoreExclude.enabled = True
        
        m_ignorePrivate.value = 0
        m_ignoreChannel.value = 0
        m_ignoreNotice.value = 0
        m_ignoreCtcp.value = 0
        m_ignoreInvite.value = 0
        m_ignoreCodes.value = 0
    End If
End Sub

Private Sub m_ignoreExclude_Click()
    If m_ignoreExclude.value = 1 Then
        m_ignoreAll.enabled = False
        m_ignorePrivate.enabled = False
        m_ignoreChannel.enabled = False
        m_ignoreNotice.enabled = False
        m_ignoreCtcp.enabled = False
        m_ignoreInvite.enabled = False
        m_ignoreCodes.enabled = False
        
        m_ignoreAll.value = 0
        m_ignorePrivate.value = 0
        m_ignoreChannel.value = 0
        m_ignoreNotice.value = 0
        m_ignoreCtcp.value = 0
        m_ignoreInvite.value = 0
        m_ignoreCodes.value = 0
    Else
        m_ignoreAll.enabled = True
        m_ignorePrivate.enabled = True
        m_ignoreChannel.enabled = True
        m_ignoreNotice.enabled = True
        m_ignoreCtcp.enabled = True
        m_ignoreInvite.enabled = True
        m_ignoreCodes.enabled = True
    End If
End Sub
