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
10        Set m_session = newValue
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
10        If KeyCode = vbKeyEscape Then
20            m_buttonCancel_clicked
30        End If
End Sub

Private Sub Form_Load()
10        Me.BackColor = colourManager.getColour(SWIFTCOLOUR_WINDOW)
20        initControls
30        loadIgnores
End Sub

Private Sub initControls()
10        listIgnores.Move 20, 20, Me.ScaleWidth - 120, Me.ScaleHeight - 70
          
20        Set m_buttonAdd = addButton(Controls, "&Add", Me.ScaleWidth - 95, 20, 75, 20)
30        Set m_buttonEdit = addButton(Controls, "&Edit", Me.ScaleWidth - 95, 45, 75, 20)
40        Set m_buttonDel = addButton(Controls, "&Remove", Me.ScaleWidth - 95, 70, 75, 20)
50        Set m_buttonClear = addButton(Controls, "C&lear", Me.ScaleWidth - 95, 95, 75, 20)
          
60        Set m_buttonOk = addButton(Controls, "&OK", Me.ScaleWidth - 175, Me.ScaleHeight - 40, 75, 20)
70        Set m_buttonCancel = addButton(Controls, "&Cancel", Me.ScaleWidth - 95, Me.ScaleHeight - 40, 75, 20)
          
80        updateColours Controls
End Sub

Private Sub Form_Paint()
          Dim oldBrush As Long
          
10        oldBrush = SelectObject(Me.hdc, colourManager.getBrush(SWIFTCOLOUR_FRAMEBACK))
20        RoundRect Me.hdc, 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10, 10, 10
30        SelectObject Me.hdc, oldBrush
End Sub

Private Sub updateIgnoreList()
10        listIgnores.clear
          
          Dim count As Long
          
20        For count = 1 To m_tempIgnoreList.count
30            listIgnores.addItem m_tempIgnoreList.item(count).mask
40        Next count
End Sub

Private Sub removeIgnore(ignore As CIgnoreItem)
          Dim count As Long
          
10        For count = 1 To m_tempIgnoreList.count
20            If m_tempIgnoreList.item(count) Is ignore Then
30                m_tempIgnoreList.Remove count
40                Exit Sub
50            End If
60        Next count
End Sub

Private Sub listIgnores_DblClick()
10        m_buttonEdit_clicked
End Sub

Private Sub m_buttonAdd_clicked()
          Dim ignoreEditor As New frmIgnoreEditor
          
10        ignoreEditor.session = m_session
20        ignoreEditor.Show vbModal, Me
          
30        If Not ignoreEditor.cancelled Then
              Dim ignore As New CIgnoreItem
              
40            ignore.mask = ignoreEditor.address
50            ignore.flags = ignoreEditor.flags
              
60            m_tempIgnoreList.Add ignore
              
70            updateIgnoreList
80        End If
          
90        Unload ignoreEditor
End Sub

Private Sub m_buttonEdit_clicked()
10        If listIgnores.SelCount = 1 Then
              Dim count As Long
              Dim ignore As CIgnoreItem
              
20            For count = 0 To listIgnores.ListCount - 1
30                If listIgnores.selected(count) Then
40                    Set ignore = m_tempIgnoreList.item(count + 1)
50                    Exit For
60                End If
70            Next count
              
80            If Not ignore Is Nothing Then
                  Dim ignoreEditor  As New frmIgnoreEditor
                  
90                ignoreEditor.session = m_session
100               ignoreEditor.loadIgnore ignore
110               ignoreEditor.Show vbModal, Me
                  
120               If Not ignoreEditor.cancelled Then
130                   ignore.mask = ignoreEditor.address
140                   ignore.flags = ignoreEditor.flags
150                   updateIgnoreList
160               End If
                  
170               Unload ignoreEditor
180           End If
190       End If
End Sub

Private Sub m_buttonClear_clicked()
10        m_tempIgnoreList.clear
20        updateIgnoreList
End Sub

Private Sub m_buttonDel_clicked()
10        If listIgnores.SelCount = 0 Then
20            Exit Sub
30        End If
          
          Dim count As Long
          Dim toRemove As New cArrayList
          
40        For count = 0 To listIgnores.ListCount - 1
50            If listIgnores.selected(count) Then
60                toRemove.Add m_tempIgnoreList.item(count + 1)
70            End If
80        Next count
          
90        For count = 1 To toRemove.count
100           removeIgnore toRemove.item(count)
110       Next count
          
120       updateIgnoreList
End Sub

Private Sub loadIgnores()
          Dim count As Long
          Dim ignore As CIgnoreItem
          Dim newIgnore As CIgnoreItem
          
10        For count = 1 To ignoreManager.ignoreCount
20            Set newIgnore = New CIgnoreItem
              
30            Set ignore = ignoreManager.ignore(count)
              
40            newIgnore.mask = ignore.mask
50            newIgnore.flags = ignore.flags
              
60            m_tempIgnoreList.Add newIgnore
70        Next count
          
80        updateIgnoreList
End Sub

Private Sub saveIgnores()
10        ignoreManager.clearIgnores
          
          Dim count As Long
          Dim ignore As CIgnoreItem
          
20        For count = 1 To m_tempIgnoreList.count
30            Set ignore = m_tempIgnoreList.item(count)
40            ignoreManager.addIgnore ignore
50        Next count
          
60        saveIgnoreFile
End Sub

Private Sub m_buttonOk_clicked()
10        saveIgnores
20        Me.Hide
End Sub

Private Sub m_buttonCancel_clicked()
10        Me.Hide
End Sub
