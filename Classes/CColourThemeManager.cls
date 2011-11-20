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

Friend Property Get currentSettingsTheme() As CColourTheme
10        Set currentSettingsTheme = m_currentSettingsTheme
End Property

Friend Property Let currentTheme(newValue As CColourTheme)
10        Set m_currentTheme = newValue
End Property

Friend Property Get currentTheme() As CColourTheme
10        Set currentTheme = m_currentTheme
End Property

Friend Property Let currentSettingsTheme(newValue As CColourTheme)
10        Set m_currentSettingsTheme = newValue
End Property

Friend Function newTheme(name As String) As CColourTheme
10        Set newTheme = addTheme(name)
End Function

Private Function addTheme(name As String) As CColourTheme
10        If Not findTheme(name) Is Nothing Then
20            Exit Function
30        End If

          Dim theme As New CColourTheme
          
40        theme.name = name
50        m_themes.Add theme, LCase$(name)
          
60        Set addTheme = theme
End Function

Friend Sub addThemeIndirect(theme As CColourTheme)
10        If Not findTheme(theme.name) Is Nothing Then
20            Exit Sub
30        End If

40        m_themes.Add theme, LCase$(theme.name)
End Sub

Friend Function findTheme(name As String) As CColourTheme
10        On Error Resume Next
20        Set findTheme = m_themes.item(LCase$(name))
End Function

Friend Sub clear()
10        Do Until m_themes.count = 0
20            m_themes.Remove 1
30        Loop
End Sub

Friend Sub loadThemes(file As String)
          Dim xml As New DOMDocument30
          Dim paletteNode As IXMLDOMNode
          Dim themeNode As IXMLDOMNode
          Dim node As IXMLDOMNode
          Dim attrib As IXMLDOMAttribute
          Dim theme As CColourTheme
          Dim colourIndex As Long
          Dim eventCount As Long
          
10        If Not xml.load(file) Then
20            Set m_currentTheme = addDefaultTheme
30            Exit Sub
40        End If
          
50        For Each themeNode In xml.documentElement.childNodes
60            If themeNode.nodeName = "theme" Then
70                Set theme = New CColourTheme
              
80                Set attrib = themeNode.Attributes.getNamedItem("name")
90                theme.name = attrib.text
                  
100               Set paletteNode = themeNode.selectSingleNode("palette")
                  
110               colourIndex = 0
                  
120               If Not paletteNode Is Nothing Then
130                   For Each node In paletteNode.childNodes
140                       theme.paletteEntry(colourIndex) = Val(node.text)
150                       colourIndex = colourIndex + 1
                          
160                       If colourIndex = 16 Then
170                           Exit For
180                       End If
190                   Next node
200               End If
                  
210               For eventCount = 1 To theme.eventColourCount - 1
220                   If xmlElementExists(themeNode, eventColours.getConfName(eventCount)) Then
230                       theme.eventColour(eventCount) = Val(xmlGetElementText(themeNode, _
                              eventColours.getConfName(eventCount)))
240                   Else
250                       theme.eventColour(eventCount) = m_defaultTheme.eventColour(eventCount)
260                   End If
270               Next eventCount
                  
280               theme.backgroundColour = Val(xmlGetElementText(themeNode, "background"))
290               theme.switchbarColourEvent = Val(xmlGetElementText(themeNode, "switchbarEvent"))
300               theme.switchbarColourMessage = Val(xmlGetElementText(themeNode, "switchbarMessage"))
310               theme.switchbarColourHighlight = Val(xmlGetElementText(themeNode, "switchbarHighlight"))
                  
320               addThemeIndirect theme
330           End If
340       Next themeNode
          
350       If m_themes.count = 0 Then
360           Set m_currentTheme = addDefaultTheme
370       Else
              Dim currentThemeName As String
              
380           currentThemeName = xmlGetElementText(xml.documentElement, "currentTheme")
              
390           Set m_currentTheme = findTheme(currentThemeName)
              
400           If m_currentTheme Is Nothing Then
410               Set m_currentTheme = m_themes.item(1)
420           End If
430       End If
End Sub

Friend Sub saveThemes(file As String)
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
          
10        Set xml.documentElement = xml.createNode(NODE_ELEMENT, "themes", "")

20        xmlAddElement xml, xml.documentElement, "currentTheme", m_currentTheme.name

30        For Each theme In m_themes
40            Set themeNode = xml.createNode(NODE_ELEMENT, "theme", "")
50            xml.documentElement.appendChild themeNode
              
60            Set attrib = xml.createAttribute("name")
70            attrib.text = theme.name
80            themeNode.Attributes.setNamedItem attrib
              
90            Set paletteNode = xml.createNode(NODE_ELEMENT, "palette", "")
100           themeNode.appendChild paletteNode
              
110           For paletteCount = 0 To theme.paletteCount - 1
120               xmlAddElement xml, paletteNode, "colour", CStr(theme.paletteEntry(paletteCount))
130           Next paletteCount
              
140           For eventCount = 1 To theme.eventColourCount - 1
150                   xmlAddElement xml, themeNode, eventColours.getConfName(eventCount), _
                          CStr(theme.eventColour(eventCount))
160           Next eventCount
              
170           xmlAddElement xml, themeNode, "background", CStr(theme.backgroundColour)
              
180           xmlAddElement xml, themeNode, "switchbarEvent", CStr(theme.switchbarColourEvent)
190           xmlAddElement xml, themeNode, "switchbarMessage", CStr(theme.switchbarColourMessage)
200           xmlAddElement xml, themeNode, "switchbarHighlight", CStr(theme.switchbarColourHighlight)
210       Next theme

220       On Error GoTo ThemesSave_Error
          
230       saveXml xml, file
          
240   Exit Sub
ThemesSave_Error:
250       handleError "saveThemes", Err.Number, Err.Description, Erl, vbNullString
End Sub

Friend Property Get themeCount() As Long
10        themeCount = m_themes.count
End Property

Friend Property Get theme(index As Long) As CColourTheme
10        Set theme = m_themes.item(index)
End Property

Friend Sub copyThemes(target As Collection)
          Dim theme As CColourTheme
          Dim newTheme As CColourTheme
          
10        For Each theme In m_themes
20            Set newTheme = New CColourTheme
              
30            theme.copy newTheme
40            target.Add newTheme, LCase$(newTheme.name)
50        Next theme
End Sub

Private Function addDefaultTheme() As CColourTheme
          Dim theme As New CColourTheme
          
10        m_defaultTheme.copy theme
          
20        addThemeIndirect theme
          
30        Set addDefaultTheme = theme
End Function

Private Sub Class_Initialize()
          Dim theme As New CColourTheme
          
10        theme.name = "Default"
          
20        theme.paletteEntry(0) = RGB(255, 255, 255)
30        theme.paletteEntry(1) = RGB(0, 0, 0)
40        theme.paletteEntry(2) = RGB(0, 0, 127)
50        theme.paletteEntry(3) = RGB(0, 147, 0)
60        theme.paletteEntry(4) = RGB(255, 0, 0)
70        theme.paletteEntry(5) = RGB(127, 0, 0)
80        theme.paletteEntry(6) = RGB(156, 0, 156)
90        theme.paletteEntry(7) = RGB(252, 127, 0)
100       theme.paletteEntry(8) = RGB(255, 255, 0)
110       theme.paletteEntry(9) = RGB(0, 252, 0)
120       theme.paletteEntry(10) = RGB(0, 147, 147)
130       theme.paletteEntry(11) = RGB(0, 255, 255)
140       theme.paletteEntry(12) = RGB(0, 0, 252)
150       theme.paletteEntry(13) = RGB(255, 0, 255)
160       theme.paletteEntry(14) = RGB(127, 127, 127)
170       theme.paletteEntry(15) = RGB(210, 210, 210)
          
180       theme.eventColour(1) = 1
190       theme.eventColour(2) = 1
200       theme.eventColour(3) = 5
210       theme.eventColour(4) = 6
220       theme.eventColour(5) = 3
230       theme.eventColour(6) = 3
240       theme.eventColour(7) = 3
250       theme.eventColour(8) = 2
260       theme.eventColour(9) = 3
270       theme.eventColour(10) = 3
280       theme.eventColour(11) = 3
290       theme.eventColour(12) = 1
300       theme.eventColour(13) = 4
310       theme.eventColour(14) = 2
320       theme.eventColour(15) = 1
330       theme.eventColour(16) = 4
340       theme.eventColour(17) = 2
          
350       theme.switchbarColourEvent = 2
360       theme.switchbarColourMessage = 4
370       theme.switchbarColourHighlight = 3
          
380       Set m_defaultTheme = theme
End Sub
