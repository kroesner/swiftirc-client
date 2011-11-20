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
          
10        highlight.text = text
20        m_highlights.Add highlight
End Sub

Public Sub addHighlightIndirect(highlight As CHighlight)
10        m_highlights.Add highlight
End Sub

Public Property Get highlightCount() As Long
10        highlightCount = m_highlights.count
End Property

Public Property Get highlightItem(index As Long) As CHighlight
10        Set highlightItem = m_highlights.item(index)
End Property

Public Sub clearHighlights()
10        m_highlights.clear
End Sub

Public Property Get getStandardHighlight() As CHighlight
          Dim highlight As New CHighlight
          
10        highlight.colour = 4
20        Set getStandardHighlight = highlight
End Property

Public Sub save()
          Dim xml As New DOMDocument30
10        Set xml.documentElement = xml.createNode(NODE_ELEMENT, "highlights", "")
          
          Dim count As Long
          
20        For count = 1 To m_highlights.count
30            xmlAddElement xml, xml.documentElement, "highlight", m_highlights.item(count).text
40        Next count

50        On Error GoTo HighlightsSave_Error
          
60        saveXml xml, g_userPath & "swiftirc_highlights.xml"
          
70    Exit Sub
HighlightsSave_Error:
80        handleError "saveHighlights", Err.Number, Err.Description, Erl, vbNullString
End Sub

Public Sub load()
          Dim xml As New DOMDocument30
          
10        If Not xml.load(g_userPath & "swiftirc_highlights.xml") Then
20            Exit Sub
30        End If
          
          Dim node As IXMLDOMNode
          
40        m_highlights.clear
          
50        For Each node In xml.documentElement.childNodes
60            If StrComp(node.nodeName, "highlight", vbTextCompare) = 0 Then
70                addHighlight node.text
80            End If
90        Next node
End Sub
