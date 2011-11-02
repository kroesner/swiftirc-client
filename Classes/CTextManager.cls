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
    If Not getTextItem(name) Is Nothing Then
        Exit Sub
    End If

    Dim textItem As New CTextItem
    
    textItem.name = name
    textItem.text = text
    
    m_textEntries.Add textItem, LCase$(name)
End Sub

Private Function getTextItem(name As String) As CTextItem
    On Error Resume Next
    Set getTextItem = m_textEntries.item(LCase$(name))
End Function

Public Function getText(name As String) As String
    Dim textItem As CTextItem
    
    Set textItem = getTextItem(name)
    
    If Not textItem Is Nothing Then
        getText = textItem.text
    End If
End Function

Public Sub loadText(fileName As String)
    Dim xml As New msxml2.DOMDocument30
    
    If Not xml.load(fileName) Then
        Exit Sub
    End If
    
    Dim textNode As msxml2.IXMLDOMNode
    Dim textAttrib As msxml2.IXMLDOMAttribute
    
    For Each textNode In xml.documentElement.childNodes
        If LCase$(textNode.nodeName) = "text" Then
            Set textAttrib = textNode.Attributes.getNamedItem("name")
            
            addText textAttrib.text, textNode.text
        End If
    Next textNode
End Sub
