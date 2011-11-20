Attribute VB_Name = "modUI"
Option Explicit

Private nameCounter As Integer

Private uiFonts As CFontManager
Private uiFonts2 As CFontManager

Public g_fontUI As Long
Public g_fontSubHeading As Long
Public g_fontHeading As Long

Public Enum eHighlightType
    ehtMessages
    ehtNicks
    ehtBoth
End Enum

Public Enum eLabelType
    ltNormal
    ltSubHeading
    ltHeading
End Enum

Public Enum eFieldJustification
    fjleft
    fjRight
End Enum

Public g_textViewBack As Long
Public g_textViewFore As Long
Public g_textInputBack As Long
Public g_textInputFore As Long
Public g_nicklistBack As Long
Public g_nicklistFore As Long
Public g_channelListBack As Long
Public g_channelListFore As Long

Public g_requestInputLocked As Boolean

Public Function getBestDefaultFont() As String
10        If osVersion.dwMajorVersion >= 6 Then
20            getBestDefaultFont = "Calibri"
30        ElseIf osVersion.dwMajorVersion = 5 Then
40            getBestDefaultFont = "Tahoma"
50        Else
60            getBestDefaultFont = "MS Sans Serif"
70        End If
End Function

Public Function requestInput(title As String, request As String, Optional defaultValue As String = _
    vbNullString, Optional parent As Object = Nothing, Optional password As Boolean = False) As Variant
10        If g_requestInputLocked Then
              'Events could trigger another requestInput(), which will cause
              'problems because of the modal display.  Return a cancel
20            requestInput = False
30            Exit Function
40        End If
          
50        g_requestInputLocked = True
          Dim requestForm As New frmGenericInput
          
60        requestForm.init title, request, defaultValue, password
70        requestForm.Show vbModal, parent
80        g_requestInputLocked = False
          
90        If requestForm.cancelled Then
100           requestInput = False
110           Unload requestForm
120           Exit Function
130       End If
          
140       requestInput = requestForm.value
150       Unload requestForm
End Function

Public Sub initUIFonts(hdc As Long)
10        Set uiFonts = New CFontManager
20        Set uiFonts2 = New CFontManager
          
30        uiFonts.changeFont hdc, "Verdana", 8, False, False
40        uiFonts2.changeFont hdc, "Verdana", 10, False, False
          
50        g_fontUI = uiFonts.getFont(False, False, False)
60        g_fontSubHeading = uiFonts.getFont(True, False, False)
70        g_fontHeading = uiFonts2.getFont(True, False, False)
End Sub

Public Function createControl(container As Object, progId As String, name As String) As control
          Dim newControl As control
          
10        Set newControl = container.Add(progId, name & nameCounter)
20        nameCounter = nameCounter + 1
30        newControl.visible = True
          
40        If TypeOf newControl Is VBControlExtender Then
              Dim controlExtender As VBControlExtender
              
50            Set controlExtender = newControl
                  
60            If TypeOf controlExtender.object Is IWindow Then
                  Dim window As IWindow
                  
70                Set window = controlExtender.object
80                window.realWindow = controlExtender
90            End If
100       ElseIf TypeOf newControl Is VB.ListBox Then
              Dim aListBox As VB.ListBox
              
110           Set aListBox = newControl
120           SendMessage aListBox.hwnd, WM_SETFONT, g_fontUI, 1
130       End If
          
140       Set createControl = newControl
End Function

Public Function addField(container As Object, caption As String, x As Long, y As Long, width As _
    Long, height As Long) As ctlField
          
          Dim newField As ctlField
          
10        Set newField = createControl(container, "swiftIrc.ctlField", "field")
20        newField.caption = caption
30        getRealWindow(newField).Move x, y, width, height
          
40        Set addField = newField
End Function

Public Function addButton(container As Object, caption As String, x As Long, y As Long, width As _
    Long, height As Long) As ctlButton

          Dim newButton As ctlButton
          
10        Set newButton = createControl(container, "swiftIrc.ctlButton", "button")
20        newButton.caption = caption
30        getRealWindow(newButton).Move x, y, width, height
          
40        Set addButton = newButton
End Function

Public Function addCheckBox(container As Object, caption As String, x As Long, y As Long, width As _
    Long, height As Long) As VB.CheckBox
          
          Dim newCheckBox As VB.CheckBox
          
10        Set newCheckBox = createControl(container, "VB.CheckBox", "checkbox")
20        newCheckBox.caption = caption
30        newCheckBox.Move x, y, width, height
40        SendMessage newCheckBox.hwnd, WM_SETFONT, g_fontUI, 1
          
50        newCheckBox.Appearance = 0
          
60        Set addCheckBox = newCheckBox
End Function

Public Function addFrame(container As Object, x As Long, y As Long, width As Long, height As Long) _
    As VB.Frame
          Dim newFrame As VB.Frame
          
10        Set newFrame = createControl(container, "VB.Frame", "frame")
20        newFrame.Move x, y, width, height
30        newFrame.BorderStyle = 0
40        newFrame.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
          
50        Set addFrame = newFrame
End Function

Public Function addComboList(container As Object, x As Long, y As Long, width As Long) As _
    VB.ComboBox
          Dim newCombo As VB.ComboBox
          
10        Set newCombo = createControl(container, "VB.ComboBox", "comboBox")
20        newCombo.left = x
30        newCombo.top = y
40        newCombo.width = width
50        newCombo.Appearance = 0
          
60        Set addComboList = newCombo
End Function

Public Sub moveFrameChild(window As IWindow, x As Long, y As Long, width As Long, height As Long)
10        window.realWindow.Move x * Screen.TwipsPerPixelX, y * Screen.TwipsPerPixelY, width * _
              Screen.TwipsPerPixelX, height * Screen.TwipsPerPixelY
End Sub

'Private Function addField(caption As String, x As Long, y As Long, width As Long, height As Long) As ctlField
'    Set addField = createControl("swiftIrc.ctlField", "field", True)
'    addField.caption = caption
'    getRealWindow(addField).Move x, y, width, height
'End Function

Public Function updateFonts(container As Object, fontManager As CFontManager)
          Dim aControl As VB.control
          Dim controlExtender As VB.VBControlExtender
          Dim fontUser As IFontUser
          
10        For Each aControl In container
20            If TypeOf aControl Is VBControlExtender Then
30                Set controlExtender = aControl
                  
40                If TypeOf controlExtender.object Is IFontUser Then
50                    Set fontUser = controlExtender.object
60                    fontUser.fontManager = fontManager
70                    fontUser.fontsUpdated
80                End If
90            Else
100               SendMessage aControl.hwnd, WM_SETFONT, fontManager.getDefaultFont, 1
110           End If
120       Next aControl
End Function

Public Function updateColours(container As Object)
          Dim aControl As VB.control
          Dim controlExtender As VB.VBControlExtender
          Dim colourUser As IColourUser
          Dim aCheckBox As VB.CheckBox
          Dim aListBox As VB.ListBox
          Dim aComboBox As VB.ComboBox
          Dim aTextBox As VB.TextBox
          
10        For Each aControl In container
20            If TypeOf aControl Is VBControlExtender Then
30                Set controlExtender = aControl
                  
40                If TypeOf controlExtender.object Is IColourUser Then
50                    Set colourUser = controlExtender.object
60                    colourUser.coloursUpdated
70                End If
80            ElseIf TypeOf aControl Is VB.CheckBox Then
90                Set aCheckBox = aControl
100               aCheckBox.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
110               aCheckBox.ForeColor = colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
120           ElseIf TypeOf aControl Is VB.ListBox Then
130               Set aListBox = aControl
140               aListBox.Appearance = 0
150               aListBox.BackColor = colourManager.getColour(SWIFTCOLOUR_CONTROLBACK)
160               aListBox.ForeColor = colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
170           ElseIf TypeOf aControl Is VB.ComboBox Then
180               Set aComboBox = aControl
190               aComboBox.Appearance = 0
200               aComboBox.BackColor = colourManager.getColour(SWIFTCOLOUR_CONTROLBACK)
210               aComboBox.ForeColor = colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
220           ElseIf TypeOf aControl Is VB.TextBox Then
230               Set aTextBox = aControl
240               aTextBox.Appearance = 0
250               aTextBox.BackColor = colourManager.getColour(SWIFTCOLOUR_CONTROLBACK)
260               aTextBox.ForeColor = colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
270           End If
280       Next aControl
End Function


