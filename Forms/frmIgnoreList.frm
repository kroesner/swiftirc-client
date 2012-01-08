VERSION 5.00
Begin VB.Form frmIgnoreList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ignore list"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5505
   Icon            =   "frmIgnoreList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   367
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox listIgnores 
      Height          =   1740
      IntegralHeight  =   0   'False
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmIgnoreList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents m_buttonAdd As swiftIrc.ctlButton
Attribute m_buttonAdd.VB_VarHelpID = -1
Private WithEvents m_buttonEdit As swiftIrc.ctlButton
Attribute m_buttonEdit.VB_VarHelpID = -1
Private WithEvents m_buttonDel As swiftIrc.ctlButton
Attribute m_buttonDel.VB_VarHelpID = -1
Private WithEvents m_buttonClear As swiftIrc.ctlButton
Attribute m_buttonClear.VB_VarHelpID = -1
Private WithEvents m_buttonOk As swiftIrc.ctlButton
Attribute m_buttonOk.VB_VarHelpID = -1
Private WithEvents m_buttonCancel As swiftIrc.ctlButton
Attribute m_buttonCancel.VB_VarHelpID = -1

Private m_tempIgnoreList As New cArrayList
Private m_session As CSession

Public Property Let session(newValue As CSession)
    Set m_session = newValue
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        m_buttonCancel_clicked
    End If
End Sub

Private Sub Form_Load()
    Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
    initControls
    loadIgnores
End Sub

Private Sub initControls()
    listIgnores.Move 20, 20, Me.ScaleWidth - 120, Me.ScaleHeight - 70
    
    Set m_buttonAdd = addButton(Controls, "&Add", Me.ScaleWidth - 95, 20, 75, 20)
    Set m_buttonEdit = addButton(Controls, "&Edit", Me.ScaleWidth - 95, 45, 75, 20)
    Set m_buttonDel = addButton(Controls, "&Remove", Me.ScaleWidth - 95, 70, 75, 20)
    Set m_buttonClear = addButton(Controls, "C&lear", Me.ScaleWidth - 95, 95, 75, 20)
    
    Set m_buttonOk = addButton(Controls, "&OK", Me.ScaleWidth - 175, Me.ScaleHeight - 40, 75, 20)
    Set m_buttonCancel = addButton(Controls, "&Cancel", Me.ScaleWidth - 95, Me.ScaleHeight - 40, 75, 20)
    
    updateColours Controls
End Sub

Private Sub Form_Paint()
    Dim oldBrush As Long
    
    oldBrush = SelectObject(Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK))
    RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
    SelectObject Me.hdc, oldBrush
End Sub

Private Sub updateIgnoreList()
    listIgnores.clear
    
    Dim count As Long
    
    For count = 1 To m_tempIgnoreList.count
        listIgnores.addItem m_tempIgnoreList.item(count).mask
    Next count
End Sub

Private Sub removeIgnore(ignore As CIgnoreItem)
    Dim count As Long
    
    For count = 1 To m_tempIgnoreList.count
        If m_tempIgnoreList.item(count) Is ignore Then
            m_tempIgnoreList.Remove count
            Exit Sub
        End If
    Next count
End Sub

Private Sub listIgnores_DblClick()
    m_buttonEdit_clicked
End Sub

Private Sub m_buttonAdd_clicked()
    Dim ignoreEditor As New frmIgnoreEditor
    
    ignoreEditor.session = m_session
    ignoreEditor.Show vbModal, Me
    
    If Not ignoreEditor.cancelled Then
        Dim ignore As New CIgnoreItem
        
        ignore.mask = ignoreEditor.address
        ignore.flags = ignoreEditor.flags
        
        m_tempIgnoreList.Add ignore
        
        updateIgnoreList
    End If
    
    Unload ignoreEditor
End Sub

Private Sub m_buttonEdit_clicked()
    If listIgnores.SelCount = 1 Then
        Dim count As Long
        Dim ignore As CIgnoreItem
        
        For count = 0 To listIgnores.ListCount - 1
            If listIgnores.selected(count) Then
                Set ignore = m_tempIgnoreList.item(count + 1)
                Exit For
            End If
        Next count
        
        If Not ignore Is Nothing Then
            Dim ignoreEditor  As New frmIgnoreEditor
            
            ignoreEditor.session = m_session
            ignoreEditor.loadIgnore ignore
            ignoreEditor.Show vbModal, Me
            
            If Not ignoreEditor.cancelled Then
                ignore.mask = ignoreEditor.address
                ignore.flags = ignoreEditor.flags
                updateIgnoreList
            End If
            
            Unload ignoreEditor
        End If
    End If
End Sub

Private Sub m_buttonClear_clicked()
    m_tempIgnoreList.clear
    updateIgnoreList
End Sub

Private Sub m_buttonDel_clicked()
    If listIgnores.SelCount = 0 Then
        Exit Sub
    End If
    
    Dim count As Long
    Dim toRemove As New cArrayList
    
    For count = 0 To listIgnores.ListCount - 1
        If listIgnores.selected(count) Then
            toRemove.Add m_tempIgnoreList.item(count + 1)
        End If
    Next count
    
    For count = 1 To toRemove.count
        removeIgnore toRemove.item(count)
    Next count
    
    updateIgnoreList
End Sub

Private Sub loadIgnores()
    Dim count As Long
    Dim ignore As CIgnoreItem
    Dim newIgnore As CIgnoreItem
    
    For count = 1 To ignoreManager.ignoreCount
        Set newIgnore = New CIgnoreItem
        
        Set ignore = ignoreManager.ignore(count)
        
        newIgnore.mask = ignore.mask
        newIgnore.flags = ignore.flags
        
        m_tempIgnoreList.Add newIgnore
    Next count
    
    updateIgnoreList
End Sub

Private Sub saveIgnores()
    ignoreManager.clearIgnores
    
    Dim count As Long
    Dim ignore As CIgnoreItem
    
    For count = 1 To m_tempIgnoreList.count
        Set ignore = m_tempIgnoreList.item(count)
        ignoreManager.addIgnore ignore
    Next count
    
    ignoreManager.saveIgnoreList
End Sub

Private Sub m_buttonOk_clicked()
    saveIgnores
    Me.Hide
End Sub

Private Sub m_buttonCancel_clicked()
    Me.Hide
End Sub
