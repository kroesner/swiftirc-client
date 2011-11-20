VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTextManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_textEntries As New Collection

Private Sub addText(name As String, text As String)
10        If Not getTextItem(name) Is Nothing Then
20            Exit Sub
30        End If

          Dim textItem As New CTextItem
          
40        textItem.name = name
50        textItem.text = text
          
60        m_textEntries.Add textItem, LCase$(name)
End Sub

Private Function getTextItem(name As String) As CTextItem
10        On Error Resume Next
20        Set getTextItem = m_textEntries.item(LCase$(name))
End Function

Public Function getText(name As String) As String
          Dim textItem As CTextItem
          
10        Set textItem = getTextItem(name)
          
20        If Not textItem Is Nothing Then
30            getText = textItem.text
40        End If
End Function

Public Sub loadText(filename As String)
          Dim xml As New msxml2.DOMDocument30
          
10        If Not xml.load(filename) Then
20            Exit Sub
30        End If
          
          Dim textNode As msxml2.IXMLDOMNode
          Dim textAttrib As msxml2.IXMLDOMAttribute
          
40        For Each textNode In xml.documentElement.childNodes
50            If LCase$(textNode.nodeName) = "text" Then
60                Set textAttrib = textNode.Attributes.getNamedItem("name")
                  
70                addText textAttrib.text, textNode.text
80            End If
90        Next textNode
End Sub
