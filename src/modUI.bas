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
    If osVersion.dwMajorVersion >= 6 Then
        getBestDefaultFont = "Calibri"
    ElseIf osVersion.dwMajorVersion = 5 Then
        getBestDefaultFont = "Tahoma"
    Else
        getBestDefaultFont = "MS Sans Serif"
    End If
End Function

Public Function requestInput(title As String, request As String, Optional defaultValue As String = vbNullString, Optional parent As Object = Nothing, Optional password As Boolean = False) As Variant
    If g_requestInputLocked Then
        'Events could trigger another requestInput(), which will cause
        'problems because of the modal display.  Return a cancel
        requestInput = False
        Exit Function
    End If
    
    g_requestInputLocked = True
    Dim requestForm As New frmGenericInput
    
    requestForm.init title, request, defaultValue, password
    requestForm.Show vbModal, parent
    g_requestInputLocked = False
    
    If requestForm.cancelled Then
        requestInput = False
        Unload requestForm
        Exit Function
    End If
    
    requestInput = requestForm.value
    Unload requestForm
End Function

Public Sub initUIFonts(hdc As Long)
    Set uiFonts = New CFontManager
    Set uiFonts2 = New CFontManager
    
    uiFonts.changeFont hdc, "Verdana", 8, False, False
    uiFonts2.changeFont hdc, "Verdana", 10, False, False
    
    g_fontUI = uiFonts.getFont(False, False, False)
    g_fontSubHeading = uiFonts.getFont(True, False, False)
    g_fontHeading = uiFonts2.getFont(True, False, False)
End Sub

Public Function createControl(container As Object, progId As String, name As String) As control
    Dim newControl As control
    
    Set newControl = container.Add(progId, name & nameCounter)
    nameCounter = nameCounter + 1
    newControl.visible = True
    
    If TypeOf newControl Is VBControlExtender Then
        Dim controlExtender As VBControlExtender
        
        Set controlExtender = newControl
            
        If TypeOf controlExtender.object Is IWindow Then
            Dim window As IWindow
            
            Set window = controlExtender.object
            window.realWindow = controlExtender
        End If
    ElseIf TypeOf newControl Is VB.ListBox Then
        Dim aListBox As VB.ListBox
        
        Set aListBox = newControl
        SendMessage aListBox.hwnd, WM_SETFONT, g_fontUI, 1
    End If
    
    Set createControl = newControl
End Function

Public Function addField(container As Object, caption As String, x As Long, y As Long, width As Long, height As Long) As ctlField
    
    Dim newField As ctlField
    
    Set newField = createControl(container, "swiftIrc.ctlField", "field")
    newField.caption = caption
    getRealWindow(newField).Move x, y, width, height
    
    Set addField = newField
End Function

Public Function addButton(container As Object, caption As String, x As Long, y As Long, width As Long, height As Long) As ctlButton

    Dim newButton As ctlButton
    
    Set newButton = createControl(container, "swiftIrc.ctlButton", "button")
    newButton.caption = caption
    getRealWindow(newButton).Move x, y, width, height
    
    Set addButton = newButton
End Function

Public Function addCheckBox(container As Object, caption As String, x As Long, y As Long, width As Long, height As Long) As VB.CheckBox
    
    Dim newCheckBox As VB.CheckBox
    
    Set newCheckBox = createControl(container, "VB.CheckBox", "checkbox")
    newCheckBox.caption = caption
    newCheckBox.Move x, y, width, height
    SendMessage newCheckBox.hwnd, WM_SETFONT, g_fontUI, 1
    
    newCheckBox.Appearance = 0
    
    Set addCheckBox = newCheckBox
End Function

Public Function addFrame(container As Object, x As Long, y As Long, width As Long, height As Long) As VB.Frame
    Dim newFrame As VB.Frame
    
    Set newFrame = createControl(container, "VB.Frame", "frame")
    newFrame.Move x, y, width, height
    newFrame.BorderStyle = 0
    newFrame.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    
    Set addFrame = newFrame
End Function

Public Function addComboList(container As Object, x As Long, y As Long, width As Long) As VB.ComboBox
    Dim newCombo As VB.ComboBox
    
    Set newCombo = createControl(container, "VB.ComboBox", "comboBox")
    newCombo.left = x
    newCombo.top = y
    newCombo.width = width
    newCombo.Appearance = 0
    
    Set addComboList = newCombo
End Function

Public Sub moveFrameChild(window As IWindow, x As Long, y As Long, width As Long, height As Long)
    window.realWindow.Move x * Screen.TwipsPerPixelX, y * Screen.TwipsPerPixelY, width * Screen.TwipsPerPixelX, height * Screen.TwipsPerPixelY
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
    
    For Each aControl In container
        If TypeOf aControl Is VBControlExtender Then
            Set controlExtender = aControl
            
            If TypeOf controlExtender.object Is IFontUser Then
                Set fontUser = controlExtender.object
                fontUser.fontManager = fontManager
                fontUser.fontsUpdated
            End If
        Else
            SendMessage aControl.hwnd, WM_SETFONT, fontManager.getDefaultFont, 1
        End If
    Next aControl
End Function

Public Function updateColours(container As Object)
    Dim aControl As VB.control
    Dim controlExtender As VB.VBControlExtender
    Dim colourUser As IColourUser
    Dim aCheckBox As VB.CheckBox
    Dim aListBox As VB.ListBox
    Dim aComboBox As VB.ComboBox
    Dim aTextBox As VB.TextBox
    
    For Each aControl In container
        If TypeOf aControl Is VBControlExtender Then
            Set controlExtender = aControl
            
            If TypeOf controlExtender.object Is IColourUser Then
                Set colourUser = controlExtender.object
                colourUser.coloursUpdated
            End If
        ElseIf TypeOf aControl Is VB.CheckBox Then
            Set aCheckBox = aControl
            aCheckBox.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
            aCheckBox.ForeColor = colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
        ElseIf TypeOf aControl Is VB.ListBox Then
            Set aListBox = aControl
            aListBox.Appearance = 0
            aListBox.BackColor = colourManager.getColour(SWIFTCOLOUR_CONTROLBACK)
            aListBox.ForeColor = colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
        ElseIf TypeOf aControl Is VB.ComboBox Then
            Set aComboBox = aControl
            aComboBox.Appearance = 0
            aComboBox.BackColor = colourManager.getColour(SWIFTCOLOUR_CONTROLBACK)
            aComboBox.ForeColor = colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
        ElseIf TypeOf aControl Is VB.TextBox Then
            Set aTextBox = aControl
            aTextBox.Appearance = 0
            aTextBox.BackColor = colourManager.getColour(SWIFTCOLOUR_CONTROLBACK)
            aTextBox.ForeColor = colourManager.getColour(SWIFTCOLOUR_CONTROLFORE)
        End If
    Next aControl
End Function


