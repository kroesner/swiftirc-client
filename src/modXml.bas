Attribute VB_Name = "modXml"
Option Explicit

Public Function xmlAddElement(xml As DOMDocument30, parent As IXMLDOMNode, name As String, text As _
    String) As IXMLDOMNode
    Dim newNode As IXMLDOMNode
    
    Set newNode = xml.createNode(NODE_ELEMENT, prepareXmlString(name), "")
    newNode.text = prepareXmlString(text)
    parent.appendChild newNode
    
    Set xmlAddElement = newNode
End Function

Public Function prepareXmlString(text As String) As String
    Dim count As Long
    Dim char As String
    Dim output As String
    
    Dim disallowedChars As String
    disallowedChars = Chr$(3)
    
    For count = 1 To Len(text)
        char = Mid$(text, count, 1)
        
        If InStr(disallowedChars, char) = 0 Then
            output = output & char
        End If
    Next count
    
    prepareXmlString = output
End Function

Public Function xmlGetAttributeText(node As IXMLDOMNode, name As String, attributeName As String) As String
    Dim tempNode As IXMLDOMNode
    Dim tempNode2 As IXMLDOMNode
    
    Set tempNode = node.selectSingleNode(name)
    
    If Not tempNode Is Nothing Then
        Set tempNode2 = tempNode.Attributes.getNamedItem("crypt")
        
        If Not tempNode2 Is Nothing Then
            xmlGetAttributeText = tempNode2.text
        End If
    End If
End Function

Public Function xmlGetElementText(node As IXMLDOMNode, name As String) As String
    Dim tempNode As IXMLDOMNode
    
    Set tempNode = node.selectSingleNode(name)
    
    If Not tempNode Is Nothing Then
        xmlGetElementText = tempNode.text
    End If
End Function

Public Function xmlElementExists(node As IXMLDOMNode, name As String) As Boolean
    xmlElementExists = Not node.selectSingleNode(name) Is Nothing
End Function

Public Sub saveXml(xml As DOMDocument30, filename As String)
    Dim rdr As New SAXXMLReader30
    Dim wrt As New MXXMLWriter30
    Dim Stream As New ADODB.Stream
    
    Stream.open
    Stream.Charset = "UTF-8"
    
    wrt.encoding = "UTF-8"
    wrt.indent = True
    wrt.byteOrderMark = True
    wrt.output = Stream
    wrt.standalone = True
    
    Set rdr.contentHandler = wrt
    Set rdr.dtdHandler = wrt
    Set rdr.errorHandler = wrt

    rdr.putProperty "http://xml.org/sax/properties/lexical-handler", wrt
    rdr.putProperty "http://xml.org/sax/properties/declaration-handler", wrt
    
    rdr.parse xml
    
    Stream.SaveToFile filename, adSaveCreateOverWrite
    Stream.Close
End Sub
