VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColourThemeManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_themes As New Collection
Private m_currentTheme As CColourTheme
Private m_currentSettingsTheme As CColourTheme
Private m_defaultTheme As CColourTheme

Private Const THEMES_FILENAME  As String = "swiftirc_themes.xml"

Friend Property Get currentSettingsTheme() As CColourTheme
    Set currentSettingsTheme = m_currentSettingsTheme
End Property

Friend Property Let currentTheme(newValue As CColourTheme)
    Set m_currentTheme = newValue
End Property

Friend Property Get currentTheme() As CColourTheme
    Set currentTheme = m_currentTheme
End Property

Friend Property Let currentSettingsTheme(newValue As CColourTheme)
    Set m_currentSettingsTheme = newValue
End Property

Friend Function newTheme(name As String) As CColourTheme
    Set newTheme = addTheme(name)
End Function

Private Function addTheme(name As String) As CColourTheme
    If Not findTheme(name) Is Nothing Then
        Exit Function
    End If

    Dim theme As New CColourTheme
    
    theme.name = name
    m_themes.Add theme, LCase$(name)
    
    Set addTheme = theme
End Function

Friend Sub addThemeIndirect(theme As CColourTheme)
    If Not findTheme(theme.name) Is Nothing Then
        Exit Sub
    End If

    m_themes.Add theme, LCase$(theme.name)
End Sub

Friend Function findTheme(name As String) As CColourTheme
    On Error Resume Next
    Set findTheme = m_themes.item(LCase$(name))
End Function

Friend Sub clear()
    Do Until m_themes.count = 0
        m_themes.Remove 1
    Loop
End Sub

Friend Sub loadThemes()
    Dim xml As New DOMDocument30
    Dim paletteNode As IXMLDOMNode
    Dim themeNode As IXMLDOMNode
    Dim node As IXMLDOMNode
    Dim attrib As IXMLDOMAttribute
    Dim theme As CColourTheme
    Dim colourIndex As Long
    Dim eventCount As Long
    
    Dim file As String
    
    file = combinePath(g_userPath, THEMES_FILENAME)
    
    If Not xml.load(file) Then
        Set m_currentTheme = addDefaultTheme
        Exit Sub
    End If
    
    For Each themeNode In xml.documentElement.childNodes
        If themeNode.nodeName = "theme" Then
            Set theme = New CColourTheme
        
            Set attrib = themeNode.Attributes.getNamedItem("name")
            theme.name = attrib.text
            
            Set paletteNode = themeNode.selectSingleNode("palette")
            
            colourIndex = 0
            
            If Not paletteNode Is Nothing Then
                For Each node In paletteNode.childNodes
                    theme.paletteEntry(colourIndex) = Val(node.text)
                    colourIndex = colourIndex + 1
                    
                    If colourIndex = 16 Then
                        Exit For
                    End If
                Next node
            End If
            
            For eventCount = 1 To theme.eventColourCount - 1
                If xmlElementExists(themeNode, eventColours.getConfName(eventCount)) Then
                    theme.eventColour(eventCount) = Val(xmlGetElementText(themeNode, eventColours.getConfName(eventCount)))
                Else
                    theme.eventColour(eventCount) = m_defaultTheme.eventColour(eventCount)
                End If
            Next eventCount
            
            theme.backgroundColour = Val(xmlGetElementText(themeNode, "background"))
            theme.switchbarColourEvent = Val(xmlGetElementText(themeNode, "switchbarEvent"))
            theme.switchbarColourMessage = Val(xmlGetElementText(themeNode, "switchbarMessage"))
            theme.switchbarColourHighlight = Val(xmlGetElementText(themeNode, "switchbarHighlight"))
            
            addThemeIndirect theme
        End If
    Next themeNode
    
    If m_themes.count = 0 Then
        Set m_currentTheme = addDefaultTheme
    Else
        Dim currentThemeName As String
        
        currentThemeName = xmlGetElementText(xml.documentElement, "currentTheme")
        
        Set m_currentTheme = findTheme(currentThemeName)
        
        If m_currentTheme Is Nothing Then
            Set m_currentTheme = m_themes.item(1)
        End If
    End If
End Sub

Friend Sub saveThemes()
    Dim xml As New DOMDocument30
    Dim themeNode As IXMLDOMNode
    Dim paletteNode As IXMLDOMNode
    Dim eventColourNode As IXMLDOMNode
    Dim node As IXMLDOMNode
    Dim attrib As IXMLDOMAttribute
    Dim count As Long
    Dim paletteCount As Long
    Dim eventCount As Long
    Dim theme As CColourTheme
    
    Dim file As String
    
    file = combinePath(g_userPath, THEMES_FILENAME)
    
    Set xml.documentElement = xml.createNode(NODE_ELEMENT, "themes", "")

    xmlAddElement xml, xml.documentElement, "currentTheme", m_currentTheme.name

    For Each theme In m_themes
        Set themeNode = xml.createNode(NODE_ELEMENT, "theme", "")
        xml.documentElement.appendChild themeNode
        
        Set attrib = xml.createAttribute("name")
        attrib.text = theme.name
        themeNode.Attributes.setNamedItem attrib
        
        Set paletteNode = xml.createNode(NODE_ELEMENT, "palette", "")
        themeNode.appendChild paletteNode
        
        For paletteCount = 0 To theme.paletteCount - 1
            xmlAddElement xml, paletteNode, "colour", CStr(theme.paletteEntry(paletteCount))
        Next paletteCount
        
        For eventCount = 1 To theme.eventColourCount - 1
                xmlAddElement xml, themeNode, eventColours.getConfName(eventCount), CStr(theme.eventColour(eventCount))
        Next eventCount
        
        xmlAddElement xml, themeNode, "background", CStr(theme.backgroundColour)
        
        xmlAddElement xml, themeNode, "switchbarEvent", CStr(theme.switchbarColourEvent)
        xmlAddElement xml, themeNode, "switchbarMessage", CStr(theme.switchbarColourMessage)
        xmlAddElement xml, themeNode, "switchbarHighlight", CStr(theme.switchbarColourHighlight)
    Next theme

    On Error GoTo ThemesSave_Error
    
    saveXml xml, file
    
Exit Sub
ThemesSave_Error:
    handleError "saveThemes", Err.Number, Err.Description, Erl, vbNullString
End Sub

Friend Property Get themeCount() As Long
    themeCount = m_themes.count
End Property

Friend Property Get theme(index As Long) As CColourTheme
    Set theme = m_themes.item(index)
End Property

Friend Sub copyThemes(target As Collection)
    Dim theme As CColourTheme
    Dim newTheme As CColourTheme
    
    For Each theme In m_themes
        Set newTheme = New CColourTheme
        
        theme.copy newTheme
        target.Add newTheme, LCase$(newTheme.name)
    Next theme
End Sub

Private Function addDefaultTheme() As CColourTheme
    Dim theme As New CColourTheme
    
    m_defaultTheme.copy theme
    
    addThemeIndirect theme
    
    Set addDefaultTheme = theme
End Function

Private Sub Class_Initialize()
    Dim theme As New CColourTheme
    
    theme.name = "Default"
    
    theme.paletteEntry(0) = RGB(255, 255, 255)
    theme.paletteEntry(1) = RGB(0, 0, 0)
    theme.paletteEntry(2) = RGB(0, 0, 127)
    theme.paletteEntry(3) = RGB(0, 147, 0)
    theme.paletteEntry(4) = RGB(255, 0, 0)
    theme.paletteEntry(5) = RGB(127, 0, 0)
    theme.paletteEntry(6) = RGB(156, 0, 156)
    theme.paletteEntry(7) = RGB(252, 127, 0)
    theme.paletteEntry(8) = RGB(255, 255, 0)
    theme.paletteEntry(9) = RGB(0, 252, 0)
    theme.paletteEntry(10) = RGB(0, 147, 147)
    theme.paletteEntry(11) = RGB(0, 255, 255)
    theme.paletteEntry(12) = RGB(0, 0, 252)
    theme.paletteEntry(13) = RGB(255, 0, 255)
    theme.paletteEntry(14) = RGB(127, 127, 127)
    theme.paletteEntry(15) = RGB(210, 210, 210)
    
    theme.eventColour(1) = 1
    theme.eventColour(2) = 1
    theme.eventColour(3) = 5
    theme.eventColour(4) = 6
    theme.eventColour(5) = 3
    theme.eventColour(6) = 3
    theme.eventColour(7) = 3
    theme.eventColour(8) = 2
    theme.eventColour(9) = 3
    theme.eventColour(10) = 3
    theme.eventColour(11) = 3
    theme.eventColour(12) = 1
    theme.eventColour(13) = 4
    theme.eventColour(14) = 2
    theme.eventColour(15) = 1
    theme.eventColour(16) = 4
    theme.eventColour(17) = 2
    
    theme.switchbarColourEvent = 2
    theme.switchbarColourMessage = 4
    theme.switchbarColourHighlight = 3
    
    Set m_defaultTheme = theme
End Sub
