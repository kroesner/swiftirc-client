VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHighlightManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_highlights As New cArrayList

Public Sub addHighlight(text As String)
    Dim highlight As New CHighlight
    
    highlight.text = text
    m_highlights.Add highlight
End Sub

Public Sub addHighlightIndirect(highlight As CHighlight)
    m_highlights.Add highlight
End Sub

Public Property Get highlightCount() As Long
    highlightCount = m_highlights.count
End Property

Public Property Get highlightItem(index As Long) As CHighlight
    Set highlightItem = m_highlights.item(index)
End Property

Public Sub clearHighlights()
    m_highlights.clear
End Sub

Public Property Get getStandardHighlight() As CHighlight
    Dim highlight As New CHighlight
    
    highlight.colour = 4
    Set getStandardHighlight = highlight
End Property

Public Sub save()
    Dim xml As New DOMDocument30
    Set xml.documentElement = xml.createNode(NODE_ELEMENT, "highlights", "")
    
    Dim count As Long
    
    For count = 1 To m_highlights.count
        xmlAddElement xml, xml.documentElement, "highlight", m_highlights.item(count).text
    Next count

    On Error GoTo HighlightsSave_Error
    
    saveXml xml, g_userPath & "swiftirc_highlights.xml"
    
Exit Sub
HighlightsSave_Error:
    handleError "saveHighlights", Err.Number, Err.Description, Erl, vbNullString
End Sub

Public Sub load()
    Dim xml As New DOMDocument30
    
    If Not xml.load(g_userPath & "swiftirc_highlights.xml") Then
        Exit Sub
    End If
    
    Dim node As IXMLDOMNode
    
    m_highlights.clear
    
    For Each node In xml.documentElement.childNodes
        If StrComp(node.nodeName, "highlight", vbTextCompare) = 0 Then
            addHighlight node.text
        End If
    Next node
End Sub
