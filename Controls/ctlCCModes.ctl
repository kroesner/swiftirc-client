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
10        Set m_channel = newValue
          
20        If m_channel.session.getChannelModeType("c") <> cmtNormal Then
30            m_checkNoColours.enabled = False
40        End If
          
50        If m_channel.session.getChannelModeType("N") <> cmtNormal Then
60            m_checkNoNickChanges.enabled = False
70        End If
          
80        If m_channel.session.getChannelModeType("E") <> cmtNormal Then
90            m_checkNoEmotes.enabled = False
100       End If
          
110       m_checkOnlyOpsSetTopic.value = -m_channel.hasMode("t")
120       m_checkNoExternalMessages.value = -m_channel.hasMode("n")
130       m_checkInviteOnly.value = -m_channel.hasMode("i")
140       m_checkModerated.value = -m_channel.hasMode("m")
150       m_checkSecret.value = -m_channel.hasMode("s")
160       m_checkNoColours.value = -m_channel.hasMode("c")
170       m_checkNoNickChanges.value = -m_channel.hasMode("N")
180       m_checkNoEmotes.value = -m_channel.hasMode("E")
          
190       If LenB(m_channel.key) <> 0 Then
200           m_checkKey.value = 1
210           m_textKey.text = m_channel.getModeParam("k")
220       End If
          
230       If m_channel.limit <> 0 Then
240           m_checkLimit.value = 1
250           m_textLimit.text = CStr(m_channel.limit)
260       End If
End Property

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Sub initControls()
10        Set m_checkOnlyOpsSetTopic = addCheckBox(Controls, "Only Ops can set topic (+t)", 10, 10, 200, _
              15)
20        Set m_checkNoExternalMessages = addCheckBox(Controls, "No messages from outside channel (+n)", _
              10, 30, 250, 15)
30        Set m_checkInviteOnly = addCheckBox(Controls, "Invite only (+i)", 10, 50, 200, 15)
40        Set m_checkModerated = addCheckBox(Controls, "Only voices+ can send messages (+m)", 10, 70, 250, _
              15)
50        Set m_checkSecret = addCheckBox(Controls, _
              "Channel will not appear in channel list or WHOIS (+s)", 10, 90, 350, 15)
60        Set m_checkNoColours = addCheckBox(Controls, "Block colour codes (+c)", 10, 110, 300, 15)
70        Set m_checkNoNickChanges = addCheckBox(Controls, "Disallow nickname changes (+N)", 10, 130, 300, _
              15)
80        Set m_checkNoEmotes = addCheckBox(Controls, "Disallow emotes/actions (+E)", 10, 150, 300, 15)
          
90        Set m_checkKey = addCheckBox(Controls, "Set channel key/password:", 10, 190, 180, 15)
100       Set m_textKey = createControl(Controls, "VB.TextBox", "textKey")
          
110       m_textKey.Move 190, 190, 75, 15
          
120       Set m_checkLimit = addCheckBox(Controls, "Set channel user limit:", 10, 210, 150, 15)
130       Set m_textLimit = createControl(Controls, "VB.TextBox", "textLimit")
          
140       m_textLimit.Move 160, 210, 35, 15
End Sub

Private Sub UserControl_Initialize()
10        initControls
20        UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
30        updateColours Controls
End Sub

Public Sub applyModes()
          Dim modes As String
          Dim params As String
          
10        If m_checkOnlyOpsSetTopic.value <> -m_channel.hasMode("t") Then
20            If m_checkOnlyOpsSetTopic.value = 1 Then
30                modes = modes & "+t"
40            Else
50                modes = modes & "-t"
60            End If
70        End If
          
80        If m_checkNoExternalMessages.value <> -m_channel.hasMode("n") Then
90            If m_checkNoExternalMessages.value = 1 Then
100               modes = modes & "+n"
110           Else
120               modes = modes & "-n"
130           End If
140       End If
          
150       If m_checkInviteOnly.value <> -m_channel.hasMode("i") Then
160           If m_checkInviteOnly.value = 1 Then
170               modes = modes & "+i"
180           Else
190               modes = modes & "-i"
200           End If
210       End If
          
220       If m_checkModerated.value <> -m_channel.hasMode("m") Then
230           If m_checkModerated.value = 1 Then
240               modes = modes & "+m"
250           Else
260               modes = modes & "-m"
270           End If
280       End If
          
290       If m_checkSecret.value <> -m_channel.hasMode("s") Then
300           If m_checkSecret.value = 1 Then
310               modes = modes & "+s"
320           Else
330               modes = modes & "-s"
340           End If
350       End If
          
360       If m_checkNoColours.value <> -m_channel.hasMode("c") Then
370           If m_checkNoColours.value = 1 Then
380               modes = modes & "+c"
390           Else
400               modes = modes & "-c"
410           End If
420       End If
          
430       If m_checkNoNickChanges.value <> -m_channel.hasMode("N") Then
440           If m_checkNoNickChanges.value = 1 Then
450               modes = modes & "+N"
460           Else
470               modes = modes & "-N"
480           End If
490       End If
          
500       If m_checkNoEmotes.value <> -m_channel.hasMode("E") Then
510           If m_checkNoEmotes.value = 1 Then
520               modes = modes & "+E"
530           Else
540               modes = modes & "-E"
550           End If
560       End If
          
570       If m_checkKey.value = 1 Then
580           If LenB(m_channel.key) = 0 Then
590               modes = modes & "+k"
600               params = params & m_textKey.text & " "
610           Else
620               If m_channel.key <> m_textKey.text Then
630                   m_channel.session.sendLine "MODE " & m_channel.name & " -k " & m_channel.key
640                   modes = modes & "+k"
650                   params = params & m_textKey.text & " "
660               End If
670           End If
680       Else
690           If LenB(m_channel.key) <> 0 Then
700               modes = modes & "-k"
710               params = params & m_channel.key & " "
720           End If
730       End If
          
740       If m_checkLimit.value = 1 Then
750           modes = modes & "+l"
760           params = params & m_textLimit.text & " "
770       Else
780           If m_channel.limit <> 0 Then
790               modes = modes & "-l"
800           End If
810       End If
          
820       m_channel.session.sendModeChange m_channel.name, modes, params
End Sub
