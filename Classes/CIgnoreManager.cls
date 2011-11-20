VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIgnoreManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_ignoreList As New cArrayList
Private m_ignoreConfNames(6) As ignoreConfName

Private Type ignoreConfName
    flag As ignoreFlags
    confName As String
End Type

Public Property Get ignoreCount() As Long
10        ignoreCount = m_ignoreList.count
End Property

Public Property Get ignore(index As Long)
10        Set ignore = m_ignoreList.item(index)
End Property

Public Sub addIgnore(ignore As CIgnoreItem)
10        m_ignoreList.Add ignore
End Sub

Public Sub clearIgnores()
10        m_ignoreList.clear
End Sub

Public Function getIgnoreByMask(mask As String) As CIgnoreItem
          Dim count As Long
          
10        For count = 1 To m_ignoreList.count
20            If swiftMatch(m_ignoreList.item(count).mask, mask) Then
30                Set getIgnoreByMask = m_ignoreList.item(count)
40                Exit Function
50            End If
60        Next count
End Function

Public Function removeIgnoreByMask(mask As String) As Boolean
          Dim count As Long
          
10        For count = m_ignoreList.count To 1 Step -1
20            If swiftMatch(m_ignoreList.item(count).mask, mask) Then
30                m_ignoreList.Remove count
40                removeIgnoreByMask = True
50            End If
60        Next count
End Function

Public Function isIgnored(mask As String, flag As Long)
          Dim count As Long
          Dim ignore As CIgnoreItem
          
10        For count = 1 To m_ignoreList.count
20            Set ignore = m_ignoreList.item(count)
              
30            If swiftMatch(ignore.mask, mask) Then
40                If ignore.flags And IGNORE_EXCLUDE Then
50                    isIgnored = False
60                    Exit Function
70                End If
                  
80                If (ignore.flags And flag) = flag Then
90                    isIgnored = True
100               End If
110           End If
120       Next count
End Function

Public Sub loadIgnoreList(filename As String)
          Dim xml As New DOMDocument30
          
10        If Not xml.load(filename) Then
20            Exit Sub
30        End If
          
40        ignoreManager.clearIgnores
          
          Dim ignoreNode As IXMLDOMNode
          Dim optionNode As IXMLDOMNode
          Dim ignore As CIgnoreItem
          Dim count As Long
          
50        For Each ignoreNode In xml.documentElement.childNodes
60            Set ignore = New CIgnoreItem
70            ignore.mask = xmlGetElementText(ignoreNode, "mask")
              
80            Set optionNode = ignoreNode.selectSingleNode("options")
              
90            If Not optionNode Is Nothing Then
100               For count = 0 To UBound(m_ignoreConfNames)
110                   If xmlElementExists(optionNode, m_ignoreConfNames(count).confName) Then
120                       ignore.flags = ignore.flags Or m_ignoreConfNames(count).flag
130                   End If
140               Next count
150           End If
              
160           If LenB(ignore.mask) <> 0 Then
170               ignoreManager.addIgnore ignore
180           End If
190       Next ignoreNode
End Sub

Public Sub saveIgnoreList(filename As String)
          Dim xml As New DOMDocument30
          Dim count As Long
          Dim optionCount As Long
          Dim ignoreNode As IXMLDOMNode
          Dim optionsNode As IXMLDOMNode
          Dim ignore As CIgnoreItem
          
10        Set xml.documentElement = xml.createNode(NODE_ELEMENT, "ignoreList", vbNullString)
          
20        For count = 1 To m_ignoreList.count
30            Set ignore = m_ignoreList.item(count)
          
40            Set ignoreNode = xmlAddElement(xml, xml.documentElement, "ignoreItem", vbNullString)
50            xmlAddElement xml, ignoreNode, "mask", ignore.mask
60            Set optionsNode = xmlAddElement(xml, ignoreNode, "options", vbNullString)
              
70            For optionCount = 0 To UBound(m_ignoreConfNames)
80                If ignore.flags And m_ignoreConfNames(optionCount).flag Then
90                    xmlAddElement xml, optionsNode, m_ignoreConfNames(optionCount).confName, vbNullString
100               End If
110           Next optionCount
120       Next count
          
130       saveXml xml, filename
End Sub

Private Sub Class_Initialize()
10        m_ignoreConfNames(0).flag = IGNORE_PRIVATE
20        m_ignoreConfNames(0).confName = "privateMessages"
30        m_ignoreConfNames(1).flag = IGNORE_CHANNEL
40        m_ignoreConfNames(1).confName = "channelMessages"
50        m_ignoreConfNames(2).flag = IGNORE_NOTICE
60        m_ignoreConfNames(2).confName = "notices"
70        m_ignoreConfNames(3).flag = IGNORE_CTCP
80        m_ignoreConfNames(3).confName = "ctcp"
90        m_ignoreConfNames(4).flag = IGNORE_INVITE
100       m_ignoreConfNames(4).confName = "invite"
110       m_ignoreConfNames(5).flag = IGNORE_CODES
120       m_ignoreConfNames(5).confName = "codes"
130       m_ignoreConfNames(6).flag = IGNORE_EXCLUDE
140       m_ignoreConfNames(6).confName = "whitelist"
End Sub
