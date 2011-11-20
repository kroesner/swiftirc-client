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
10        Set m_channel = newValue
20        populateListSelector
End Property

Private Sub comboListSelector_Click()
10        m_currentMode = Mid$(m_channel.window.session.getListModes, comboListSelector.ListIndex + 1, 1)
20        displayList
End Sub

Private Sub IColourUser_coloursUpdated()
10        UserControl.BackColor = colourManager.getColour(SWIFTCOLOUR_FRAMEBACK)
20        updateColours Controls
End Sub

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Sub populateListSelector()
          Dim count As Long
          Dim listModes As String
          
10        comboListSelector.clear
20        listModes = m_channel.window.session.getListModes
          
30        For count = 1 To Len(listModes)
40            Select Case Mid$(listModes, count, 1)
                  Case "b"
50                    comboListSelector.addItem "Bans (+b)"
60                Case "e"
70                    comboListSelector.addItem "Exceptions (+e)"
80                Case "I"
90                    comboListSelector.addItem "Invite exception (+I)"
100               Case "a"
110                   comboListSelector.addItem "Protected users (+a)"
120               Case "q"
130                   comboListSelector.addItem "Channel owners (+q)"
140               Case Else
150                   comboListSelector.addItem "Unknown (mode +" & Mid$(listModes, count, 1) & ")"
160           End Select
170       Next count
          
180       If comboListSelector.ListCount <> 0 Then
190           comboListSelector.ListIndex = 0
200           m_currentMode = Mid$(listModes, 1, 1)
210           displayList
220       End If
End Sub

Private Sub displayList()
10        listBans.clear

20        If m_channel.listIsSynced(m_currentMode) Then
              Dim list As New cArrayList
              
30            m_channel.getModeList m_currentMode, list
              
              Dim count As Long
              
40            For count = 1 To list.count
50                listBans.addItem list.item(count).param
60            Next count
70        Else
80            m_channel.syncListMode m_currentMode
90        End If
End Sub

Private Sub initControls()
10        m_labelManager.addLabel "Select list type", ltSubHeading, 10, 10
20        m_labelManager.addLabel "Modify list", ltSubHeading, 10, 65
          
30        Set m_buttonAdd = addButton(Controls, "Add", UserControl.ScaleWidth - 85, 85, 75, 20)
40        Set m_buttonEdit = addButton(Controls, "Edit", UserControl.ScaleWidth - 85, 110, 75, 20)
50        Set m_buttonRemove = addButton(Controls, "Remove", UserControl.ScaleWidth - 85, 135, 75, 20)
60        Set m_buttonClear = addButton(Controls, "Clear", UserControl.ScaleWidth - 85, 160, 75, 20)
End Sub

Private Sub m_buttonAdd_clicked()
          Dim editAddress As frmEditAddress

10        Set editAddress = New frmEditAddress
20        editAddress.Show vbModal, Me
          
30        If editAddress.success Then
40            listBans.addItem editAddress.address
50            m_channel.session.sendModeChange m_channel.name, "+" & m_currentMode, editAddress.address
60        End If
End Sub

Private Sub m_buttonEdit_clicked()
10        If listBans.ListIndex = -1 Or listBans.ListCount = 0 Or listBans.SelCount > 1 Then
20            Exit Sub
30        End If

          Dim editAddress As frmEditAddress
          
40        Set editAddress = New frmEditAddress
50        editAddress.address = listBans.list(listBans.ListIndex)
60        editAddress.Show vbModal, Me

70        If editAddress.success Then
80            m_channel.session.sendModeChange m_channel.name, "-" & m_currentMode & "+" & m_currentMode, _
                  listBans.list(listBans.ListIndex) & " " & editAddress.address
                  
90            listBans.removeItem listBans.ListIndex
100           listBans.addItem editAddress.address
110       End If
End Sub

Private Sub m_buttonRemove_clicked()
10        If listBans.SelCount = 0 Then
20            Exit Sub
30        End If
          
          Dim count As Long
          Dim modes As String
          Dim params As String
          
40        modes = "-" & String(listBans.SelCount, m_currentMode)
          
50        LockWindowUpdate listBans.hwnd
          
60        For count = listBans.ListCount - 1 To 0 Step -1
70            If listBans.selected(count) = True Then
80                params = params & listBans.list(count) & " "
90                listBans.removeItem count
100           End If
110       Next count
          
120       LockWindowUpdate 0
130       listBans.refresh
          
140       m_channel.session.sendModeChange m_channel.name, modes, params
End Sub

Private Sub m_buttonClear_clicked()
          Dim modes As String
          Dim params As String
          Dim count As Long
          
          Dim list As New cArrayList
          
10        m_channel.getModeList m_currentMode, list
          
20        modes = "-" & String(list.count, m_currentMode)
          
30        For count = 1 To list.count
40            params = params & list.item(count).param & " "
50        Next count
          
60        listBans.clear
          
70        m_channel.window.session.sendModeChange m_channel.name, modes, params
End Sub

Private Sub m_channel_modeListSynced(mode As String)
10        If mode = m_currentMode Then
20            displayList
30        End If
End Sub

Private Sub UserControl_Initialize()
10        initControls
End Sub

Private Sub UserControl_Paint()
10        m_labelManager.renderLabels UserControl.hdc
End Sub

Private Sub UserControl_Resize()
10        comboListSelector.left = 10
20        comboListSelector.top = 30
30        comboListSelector.width = UserControl.ScaleWidth - 20
40        listBans.Move 10, 85, UserControl.ScaleWidth - 105, UserControl.ScaleHeight - 85
50        getRealWindow(m_buttonAdd).Move UserControl.ScaleWidth - 85, 85, 75, 20
60        getRealWindow(m_buttonEdit).Move UserControl.ScaleWidth - 85, 110, 75, 20
70        getRealWindow(m_buttonRemove).Move UserControl.ScaleWidth - 85, 135, 75, 20
80        getRealWindow(m_buttonClear).Move UserControl.ScaleWidth - 85, 160, 75, 20
End Sub
