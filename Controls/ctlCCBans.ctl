VERSION 5.00
Begin VB.UserControl ctlCCBans 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.ComboBox comboListSelector 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "ctlCCBans.ctx":0000
      Left            =   840
      List            =   "ctlCCBans.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2430
      Width           =   2235
   End
   Begin VB.ListBox listBans 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   465
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   615
      Width           =   1575
   End
End
Attribute VB_Name = "ctlCCBans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IColourUser
Implements IWindow

Private m_labelManager As New CLabelManager

Private m_realWindow As VBControlExtender

Private WithEvents m_buttonAdd As swiftIrc.ctlButton
Attribute m_buttonAdd.VB_VarHelpID = -1
Private WithEvents m_buttonEdit As swiftIrc.ctlButton
Attribute m_buttonEdit.VB_VarHelpID = -1
Private WithEvents m_buttonRemove As swiftIrc.ctlButton
Attribute m_buttonRemove.VB_VarHelpID = -1
Private WithEvents m_buttonClear As swiftIrc.ctlButton
Attribute m_buttonClear.VB_VarHelpID = -1

Private m_currentMode As String

Private WithEvents m_channel As CChannel
Attribute m_channel.VB_VarHelpID = -1

Public Property Let channel(newValue As CChannel)
    Set m_channel = newValue
    populateListSelector
End Property

Private Sub comboListSelector_Click()
    m_currentMode = Mid$(m_channel.window.session.getListModes, comboListSelector.ListIndex + 1, 1)
    displayList
End Sub

Private Sub IColourUser_coloursUpdated()
    UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
    updateColours Controls
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Sub populateListSelector()
    Dim count As Long
    Dim listModes As String
    
    comboListSelector.clear
    listModes = m_channel.window.session.getListModes
    
    For count = 1 To Len(listModes)
        Select Case Mid$(listModes, count, 1)
            Case "b"
                comboListSelector.addItem "Bans (+b)"
            Case "e"
                comboListSelector.addItem "Exceptions (+e)"
            Case "I"
                comboListSelector.addItem "Invite exception (+I)"
            Case "a"
                comboListSelector.addItem "Protected users (+a)"
            Case "q"
                comboListSelector.addItem "Channel owners (+q)"
            Case Else
                comboListSelector.addItem "Unknown (mode +" & Mid$(listModes, count, 1) & ")"
        End Select
    Next count
    
    If comboListSelector.ListCount <> 0 Then
        comboListSelector.ListIndex = 0
        m_currentMode = Mid$(listModes, 1, 1)
        displayList
    End If
End Sub

Private Sub displayList()
    listBans.clear

    If m_channel.listIsSynced(m_currentMode) Then
        Dim list As New cArrayList
        
        m_channel.getModeList m_currentMode, list
        
        Dim count As Long
        
        For count = 1 To list.count
            listBans.addItem list.item(count).param
        Next count
    Else
        m_channel.syncListMode m_currentMode
    End If
End Sub

Private Sub initControls()
    m_labelManager.addLabel "Select list type", ltSubHeading, 10, 10
    m_labelManager.addLabel "Modify list", ltSubHeading, 10, 65
    
    Set m_buttonAdd = addButton(Controls, "Add", UserControl.ScaleWidth - 85, 85, 75, 20)
    Set m_buttonEdit = addButton(Controls, "Edit", UserControl.ScaleWidth - 85, 110, 75, 20)
    Set m_buttonRemove = addButton(Controls, "Remove", UserControl.ScaleWidth - 85, 135, 75, 20)
    Set m_buttonClear = addButton(Controls, "Clear", UserControl.ScaleWidth - 85, 160, 75, 20)
End Sub

Private Sub m_buttonAdd_clicked()
    Dim editAddress As frmEditAddress

    Set editAddress = New frmEditAddress
    editAddress.Show vbModal, Me
    
    If editAddress.success Then
        listBans.addItem editAddress.address
        m_channel.session.sendModeChange m_channel.name, "+" & m_currentMode, editAddress.address
    End If
End Sub

Private Sub m_buttonEdit_clicked()
    If listBans.ListIndex = -1 Or listBans.ListCount = 0 Or listBans.SelCount > 1 Then
        Exit Sub
    End If

    Dim editAddress As frmEditAddress
    
    Set editAddress = New frmEditAddress
    editAddress.address = listBans.list(listBans.ListIndex)
    editAddress.Show vbModal, Me

    If editAddress.success Then
        m_channel.session.sendModeChange m_channel.name, "-" & m_currentMode & "+" & m_currentMode, listBans.list(listBans.ListIndex) & " " & editAddress.address
            
        listBans.removeItem listBans.ListIndex
        listBans.addItem editAddress.address
    End If
End Sub

Private Sub m_buttonRemove_clicked()
    If listBans.SelCount = 0 Then
        Exit Sub
    End If
    
    Dim count As Long
    Dim modes As String
    Dim params As String
    
    modes = "-" & String(listBans.SelCount, m_currentMode)
    
    LockWindowUpdate listBans.hwnd
    
    For count = listBans.ListCount - 1 To 0 Step -1
        If listBans.selected(count) = True Then
            params = params & listBans.list(count) & " "
            listBans.removeItem count
        End If
    Next count
    
    LockWindowUpdate 0
    listBans.refresh
    
    m_channel.session.sendModeChange m_channel.name, modes, params
End Sub

Private Sub m_buttonClear_clicked()
    Dim modes As String
    Dim params As String
    Dim count As Long
    
    Dim list As New cArrayList
    
    m_channel.getModeList m_currentMode, list
    
    modes = "-" & String(list.count, m_currentMode)
    
    For count = 1 To list.count
        params = params & list.item(count).param & " "
    Next count
    
    listBans.clear
    
    m_channel.window.session.sendModeChange m_channel.name, modes, params
End Sub

Private Sub m_channel_modeListSynced(mode As String)
    If mode = m_currentMode Then
        displayList
    End If
End Sub

Private Sub UserControl_Initialize()
    initControls
End Sub

Private Sub UserControl_Paint()
    m_labelManager.renderLabels UserControl.hdc
End Sub

Private Sub UserControl_Resize()
    comboListSelector.left = 10
    comboListSelector.top = 30
    comboListSelector.width = UserControl.ScaleWidth - 20
    listBans.Move 10, 85, UserControl.ScaleWidth - 105, UserControl.ScaleHeight - 85
    getRealWindow(m_buttonAdd).Move UserControl.ScaleWidth - 85, 85, 75, 20
    getRealWindow(m_buttonEdit).Move UserControl.ScaleWidth - 85, 110, 75, 20
    getRealWindow(m_buttonRemove).Move UserControl.ScaleWidth - 85, 135, 75, 20
    getRealWindow(m_buttonClear).Move UserControl.ScaleWidth - 85, 160, 75, 20
End Sub
