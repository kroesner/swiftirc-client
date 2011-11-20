Attribute VB_Name = "modXml"
Option Explicit

Public Function xmlAddElement(xml As DOMDocument30, parent As IXMLDOMNode, name As String, text As _
    String) As IXMLDOMNode
          Dim newNode As IXMLDOMNode
          
10        Set newNode = xml.createNode(NODE_ELEMENT, prepareXmlString(name), "")
20        newNode.text = prepareXmlString(text)
30        parent.appendChild newNode
          
40        Set xmlAddElement = newNode
End Function

Public Function prepareXmlString(text As String) As String
          Dim count As Long
          Dim char As String
          Dim output As String
          
          Dim disallowedChars As String
10        disallowedChars = Chr$(3)
          
20        For count = 1 To Len(text)
30            char = Mid$(text, count, 1)
              
40            If InStr(disallowedChars, char) = 0 Then
50                output = output & char
60            End If
70        Next count
          
80        prepareXmlString = output
End Function

Public Function xmlGetAttributeText(node As IXMLDOMNode, name As String, attributeName As String) As String
          Dim tempNode As IXMLDOMNode
          Dim tempNode2 As IXMLDOMNode
          
10        Set tempNode = node.selectSingleNode(name)
          
20        If Not tempNode Is Nothing Then
30            Set tempNode2 = tempNode.Attributes.getNamedItem("crypt")
              
40            If Not tempNode2 Is Nothing Then
50                xmlGetAttributeText = tempNode2.text
60            End If
70        End If
End Function

Public Function xmlGetElementText(node As IXMLDOMNode, name As String) As String
          Dim tempNode As IXMLDOMNode
          
10        Set tempNode = node.selectSingleNode(name)
          
20        If Not tempNode Is Nothing Then
30            xmlGetElementText = tempNode.text
40        End If
End Function

Public Function xmlElementExists(node As IXMLDOMNode, name As String) As Boolean
10        xmlElementExists = Not node.selectSingleNode(name) Is Nothing
End Function

Public Sub saveXml(xml As DOMDocument30, filename As String)
          Dim rdr As New SAXXMLReader30
          Dim wrt As New MXXMLWriter30
          Dim Stream As New ADODB.Stream
          
10        Stream.open
20        Stream.Charset = "UTF-8"
          
30        wrt.encoding = "UTF-8"
40        wrt.indent = True
50        wrt.byteOrderMark = True
60        wrt.output = Stream
70        wrt.standalone = True
          
80        Set rdr.contentHandler = wrt
90        Set rdr.dtdHandler = wrt
100       Set rdr.errorHandler = wrt

110       rdr.putProperty "http://xml.org/sax/properties/lexical-handler", wrt
120       rdr.putProperty "http://xml.org/sax/properties/declaration-handler", wrt
          
130       rdr.parse xml
          
140       Stream.SaveToFile filename, adSaveCreateOverWrite
150       Stream.Close
End Sub
