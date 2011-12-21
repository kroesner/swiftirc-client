VERSION 5.00
Begin VB.UserControl ctlCCModes 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "ctlCCModes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow

Private m_realWindow As VBControlExtender

Private m_checkOnlyOpsSetTopic As VB.CheckBox
Private m_checkNoExternalMessages As VB.CheckBox
Private m_checkInviteOnly As VB.CheckBox
Private m_checkModerated As VB.CheckBox
Private m_checkSecret As VB.CheckBox
Private m_checkNoColours As VB.CheckBox
Private m_checkNoNickChanges As VB.CheckBox
Private m_checkNoEmotes As VB.CheckBox

Private m_checkKey As VB.CheckBox
Private m_textKey As VB.TextBox
Private m_checkLimit As VB.CheckBox
Private m_textLimit As VB.TextBox

Private m_channel As CChannel

Public Property Let channel(newValue As CChannel)
    Set m_channel = newValue
    
    If m_channel.session.getChannelModeType("c") <> cmtNormal Then
        m_checkNoColours.enabled = False
    End If
    
    If m_channel.session.getChannelModeType("N") <> cmtNormal Then
        m_checkNoNickChanges.enabled = False
    End If
    
    If m_channel.session.getChannelModeType("E") <> cmtNormal Then
        m_checkNoEmotes.enabled = False
    End If
    
    m_checkOnlyOpsSetTopic.value = -m_channel.hasMode("t")
    m_checkNoExternalMessages.value = -m_channel.hasMode("n")
    m_checkInviteOnly.value = -m_channel.hasMode("i")
    m_checkModerated.value = -m_channel.hasMode("m")
    m_checkSecret.value = -m_channel.hasMode("s")
    m_checkNoColours.value = -m_channel.hasMode("c")
    m_checkNoNickChanges.value = -m_channel.hasMode("N")
    m_checkNoEmotes.value = -m_channel.hasMode("E")
    
    If LenB(m_channel.key) <> 0 Then
        m_checkKey.value = 1
        m_textKey.text = m_channel.getModeParam("k")
    End If
    
    If m_channel.limit <> 0 Then
        m_checkLimit.value = 1
        m_textLimit.text = CStr(m_channel.limit)
    End If
End Property

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Sub initControls()
    Set m_checkOnlyOpsSetTopic = addCheckBox(Controls, "Only Ops can set topic (+t)", 10, 10, 200, _
        15)
    Set m_checkNoExternalMessages = addCheckBox(Controls, "No messages from outside channel (+n)", _
        10, 30, 250, 15)
    Set m_checkInviteOnly = addCheckBox(Controls, "Invite only (+i)", 10, 50, 200, 15)
    Set m_checkModerated = addCheckBox(Controls, "Only voices+ can send messages (+m)", 10, 70, 250, _
        15)
    Set m_checkSecret = addCheckBox(Controls, _
        "Channel will not appear in channel list or WHOIS (+s)", 10, 90, 350, 15)
    Set m_checkNoColours = addCheckBox(Controls, "Block colour codes (+c)", 10, 110, 300, 15)
    Set m_checkNoNickChanges = addCheckBox(Controls, "Disallow nickname changes (+N)", 10, 130, 300, _
        15)
    Set m_checkNoEmotes = addCheckBox(Controls, "Disallow emotes/actions (+E)", 10, 150, 300, 15)
    
    Set m_checkKey = addCheckBox(Controls, "Set channel key/password:", 10, 190, 180, 15)
    Set m_textKey = createControl(Controls, "VB.TextBox", "textKey")
    
    m_textKey.Move 190, 190, 75, 15
    
    Set m_checkLimit = addCheckBox(Controls, "Set channel user limit:", 10, 210, 150, 15)
    Set m_textLimit = createControl(Controls, "VB.TextBox", "textLimit")
    
    m_textLimit.Move 160, 210, 35, 15
End Sub

Private Sub UserControl_Initialize()
    initControls
    UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    updateColours Controls
End Sub

Public Sub applyModes()
    Dim modes As String
    Dim params As String
    
    If m_checkOnlyOpsSetTopic.value <> -m_channel.hasMode("t") Then
        If m_checkOnlyOpsSetTopic.value = 1 Then
            modes = modes & "+t"
        Else
            modes = modes & "-t"
        End If
    End If
    
    If m_checkNoExternalMessages.value <> -m_channel.hasMode("n") Then
        If m_checkNoExternalMessages.value = 1 Then
            modes = modes & "+n"
        Else
            modes = modes & "-n"
        End If
    End If
    
    If m_checkInviteOnly.value <> -m_channel.hasMode("i") Then
        If m_checkInviteOnly.value = 1 Then
            modes = modes & "+i"
        Else
            modes = modes & "-i"
        End If
    End If
    
    If m_checkModerated.value <> -m_channel.hasMode("m") Then
        If m_checkModerated.value = 1 Then
            modes = modes & "+m"
        Else
            modes = modes & "-m"
        End If
    End If
    
    If m_checkSecret.value <> -m_channel.hasMode("s") Then
        If m_checkSecret.value = 1 Then
            modes = modes & "+s"
        Else
            modes = modes & "-s"
        End If
    End If
    
    If m_checkNoColours.value <> -m_channel.hasMode("c") Then
        If m_checkNoColours.value = 1 Then
            modes = modes & "+c"
        Else
            modes = modes & "-c"
        End If
    End If
    
    If m_checkNoNickChanges.value <> -m_channel.hasMode("N") Then
        If m_checkNoNickChanges.value = 1 Then
            modes = modes & "+N"
        Else
            modes = modes & "-N"
        End If
    End If
    
    If m_checkNoEmotes.value <> -m_channel.hasMode("E") Then
        If m_checkNoEmotes.value = 1 Then
            modes = modes & "+E"
        Else
            modes = modes & "-E"
        End If
    End If
    
    If m_checkKey.value = 1 Then
        If LenB(m_channel.key) = 0 Then
            modes = modes & "+k"
            params = params & m_textKey.text & " "
        Else
            If m_channel.key <> m_textKey.text Then
                m_channel.session.sendLine "MODE " & m_channel.name & " -k " & m_channel.key
                modes = modes & "+k"
                params = params & m_textKey.text & " "
            End If
        End If
    Else
        If LenB(m_channel.key) <> 0 Then
            modes = modes & "-k"
            params = params & m_channel.key & " "
        End If
    End If
    
    If m_checkLimit.value = 1 Then
        modes = modes & "+l"
        params = params & m_textLimit.text & " "
    Else
        If m_channel.limit <> 0 Then
            modes = modes & "-l"
        End If
    End If
    
    m_channel.session.sendModeChange m_channel.name, modes, params
End Sub
