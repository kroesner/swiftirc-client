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
    ignoreCount = m_ignoreList.count
End Property

Public Property Get ignore(index As Long)
    Set ignore = m_ignoreList.item(index)
End Property

Public Sub addIgnore(ignore As CIgnoreItem)
    m_ignoreList.Add ignore
End Sub

Public Sub clearIgnores()
    m_ignoreList.clear
End Sub

Public Function getIgnoreByMask(mask As String) As CIgnoreItem
    Dim count As Long
    
    For count = 1 To m_ignoreList.count
        If swiftMatch(m_ignoreList.item(count).mask, mask) Then
            Set getIgnoreByMask = m_ignoreList.item(count)
            Exit Function
        End If
    Next count
End Function

Public Function removeIgnoreByMask(mask As String) As Boolean
    Dim count As Long
    
    For count = m_ignoreList.count To 1 Step -1
        If swiftMatch(m_ignoreList.item(count).mask, mask) Then
            m_ignoreList.Remove count
            removeIgnoreByMask = True
        End If
    Next count
End Function

Public Function isIgnored(mask As String, flag As Long)
    Dim count As Long
    Dim ignore As CIgnoreItem
    
    For count = 1 To m_ignoreList.count
        Set ignore = m_ignoreList.item(count)
        
        If swiftMatch(ignore.mask, mask) Then
            If ignore.flags And IGNORE_EXCLUDE Then
                isIgnored = False
                Exit Function
            End If
            
            If (ignore.flags And flag) = flag Then
                isIgnored = True
            End If
        End If
    Next count
End Function

Public Sub loadIgnoreList(filename As String)
    Dim xml As New DOMDocument30
    
    If Not xml.load(filename) Then
        Exit Sub
    End If
    
    ignoreManager.clearIgnores
    
    Dim ignoreNode As IXMLDOMNode
    Dim optionNode As IXMLDOMNode
    Dim ignore As CIgnoreItem
    Dim count As Long
    
    For Each ignoreNode In xml.documentElement.childNodes
        Set ignore = New CIgnoreItem
        ignore.mask = xmlGetElementText(ignoreNode, "mask")
        
        Set optionNode = ignoreNode.selectSingleNode("options")
        
        If Not optionNode Is Nothing Then
            For count = 0 To UBound(m_ignoreConfNames)
                If xmlElementExists(optionNode, m_ignoreConfNames(count).confName) Then
                    ignore.flags = ignore.flags Or m_ignoreConfNames(count).flag
                End If
            Next count
        End If
        
        If LenB(ignore.mask) <> 0 Then
            ignoreManager.addIgnore ignore
        End If
    Next ignoreNode
End Sub

Public Sub saveIgnoreList(filename As String)
    Dim xml As New DOMDocument30
    Dim count As Long
    Dim optionCount As Long
    Dim ignoreNode As IXMLDOMNode
    Dim optionsNode As IXMLDOMNode
    Dim ignore As CIgnoreItem
    
    Set xml.documentElement = xml.createNode(NODE_ELEMENT, "ignoreList", vbNullString)
    
    For count = 1 To m_ignoreList.count
        Set ignore = m_ignoreList.item(count)
    
        Set ignoreNode = xmlAddElement(xml, xml.documentElement, "ignoreItem", vbNullString)
        xmlAddElement xml, ignoreNode, "mask", ignore.mask
        Set optionsNode = xmlAddElement(xml, ignoreNode, "options", vbNullString)
        
        For optionCount = 0 To UBound(m_ignoreConfNames)
            If ignore.flags And m_ignoreConfNames(optionCount).flag Then
                xmlAddElement xml, optionsNode, m_ignoreConfNames(optionCount).confName, vbNullString
            End If
        Next optionCount
    Next count
    
    saveXml xml, filename
End Sub

Private Sub Class_Initialize()
    m_ignoreConfNames(0).flag = IGNORE_PRIVATE
    m_ignoreConfNames(0).confName = "privateMessages"
    m_ignoreConfNames(1).flag = IGNORE_CHANNEL
    m_ignoreConfNames(1).confName = "channelMessages"
    m_ignoreConfNames(2).flag = IGNORE_NOTICE
    m_ignoreConfNames(2).confName = "notices"
    m_ignoreConfNames(3).flag = IGNORE_CTCP
    m_ignoreConfNames(3).confName = "ctcp"
    m_ignoreConfNames(4).flag = IGNORE_INVITE
    m_ignoreConfNames(4).confName = "invite"
    m_ignoreConfNames(5).flag = IGNORE_CODES
    m_ignoreConfNames(5).confName = "codes"
    m_ignoreConfNames(6).flag = IGNORE_EXCLUDE
    m_ignoreConfNames(6).confName = "whitelist"
End Sub
