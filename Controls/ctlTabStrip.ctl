VERSION 5.00
Begin VB.UserControl ctlTabStrip 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   398
End
Attribute VB_Name = "ctlTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_realWindow As VBControlExtender
Private m_tabs As New cArrayList

Private m_selectedTab As CTabStripItem
Private m_overTab As CTabStripItem

Private Const TAB_WIDTH = 85
Private Const TAB_SPACING_X = 5
Private Const TAB_HEIGHT = 30

Implements IWindow

Public Event tabSelected(selectedTab As CTabStripItem)

Public Function addTab(caption As String) As CTabStripItem
    Dim aTab As New CTabStripItem
    Dim prefix As String
    
    prefix = parsePrefix(caption)
    
    aTab.init caption, prefix
    m_tabs.Add aTab
    Set addTab = aTab
    
    reDraw
End Function

Private Sub reDraw()
    Dim count As Long
    Dim x As Long
    Dim y As Long
    
    Dim backBuffer As Long
    Dim backBitmap As Long
    Dim oldBitmap As Long
    Dim oldFont As Long
    
    backBuffer = CreateCompatibleDC(UserControl.hdc)
    backBitmap = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, _
        UserControl.ScaleHeight)
    
    oldBitmap = SelectObject(backBuffer, backBitmap)
    oldFont = SelectObject(backBuffer, GetCurrentObject(hdc, OBJ_FONT))
    
    FillRect backBuffer, makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), _
        colourManager.getBrush(SWIFTCOLOUR_WINDOW)
    
    y = 5
    
    For count = 1 To m_tabs.count
        m_tabs.item(count).render backBuffer, x, y, TAB_WIDTH, TAB_HEIGHT
        x = x + TAB_WIDTH + TAB_SPACING_X
    Next count
    
    BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, backBuffer, 0, 0, _
        vbSrcCopy
        
    SelectObject backBuffer, oldBitmap
    SelectObject backBuffer, oldFont
    
    DeleteDC backBuffer
    DeleteObject backBitmap
End Sub

Private Property Get IWindow_realWindow() As Object
    Set IWindow_realWindow = m_realWindow
End Property

Private Property Let IWindow_realWindow(RHS As Object)
    Set m_realWindow = RHS
End Property

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    Dim count As Long
    Dim char As String
    
    char = Chr$(KeyAscii)
    
    For count = 1 To m_tabs.count
        If LCase$(m_tabs.item(count).prefix) = LCase$(char) Then
            If Not m_selectedTab Is Nothing Then
                m_selectedTab.state = tisNormal
            End If
            
            Set m_selectedTab = m_tabs.item(count)
            m_selectedTab.state = tisSelected
            RaiseEvent tabSelected(m_selectedTab)
            reDraw
            
            Exit Sub
        End If
    Next count
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If y < 5 Or y > UserControl.ScaleHeight Or x < 0 Or x > UserControl.ScaleWidth Then
        If Not m_overTab Is Nothing Then
            If Not m_overTab.state = tisSelected Then
                m_overTab.state = tisNormal
                reDraw
            End If
            
            Set m_overTab = Nothing
        End If
        
        ReleaseCapture
        Exit Sub
    End If
    
    Dim tabIndex As Long
    Dim aTab As CTabStripItem
    
    tabIndex = Fix(x / (TAB_WIDTH + TAB_SPACING_X)) + 1
    
    If tabIndex > 0 And tabIndex <= m_tabs.count Then
        Set aTab = m_tabs.item(tabIndex)
        
        If aTab Is m_overTab Then
            Exit Sub
        End If
        
        If Not m_overTab Is Nothing Then
            If Not m_selectedTab Is m_overTab Then
                m_overTab.state = tisNormal
            End If
        End If
        
        Set m_overTab = aTab
        
        If Not aTab.state = tisSelected Then
            aTab.state = tisMouseOver
        End If
        
        reDraw
        SetCapture UserControl.hwnd
    Else
        If Not m_overTab Is Nothing Then
            If Not m_overTab.state = tisSelected Then
                m_overTab.state = tisNormal
                reDraw
            End If
            
            Set m_overTab = Nothing
        End If
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If y < 5 Then
        Exit Sub
    End If
    
    Dim tabIndex As Long
    Dim aTab As CTabStripItem
    
    tabIndex = Fix(x / (TAB_WIDTH + TAB_SPACING_X)) + 1
    
    If tabIndex > 0 And tabIndex <= m_tabs.count Then
        Set aTab = m_tabs.item(tabIndex)
        selectTab aTab, False
    End If
End Sub

Public Sub selectTab(aTab As CTabStripItem, noEvent As Boolean)
    If m_selectedTab Is aTab Then
        Exit Sub
    End If
    
    If Not m_selectedTab Is Nothing Then
        m_selectedTab.state = tisNormal
    End If
    
    Set m_selectedTab = aTab
    m_selectedTab.state = tisSelected
    
    reDraw
    
    If Not noEvent Then
        RaiseEvent tabSelected(m_selectedTab)
    End If
End Sub

Private Function parsePrefix(caption As String) As String
    Dim count As Long
    
    For count = 1 To Len(caption)
        If Mid$(caption, count, 1) = "&" Then
            If count < Len(caption) Then
                If Mid$(caption, count + 1, 1) <> "&" Then
                    UserControl.AccessKeys = UserControl.AccessKeys & Mid$(caption, count + 1, 1)
                    parsePrefix = Mid$(caption, count + 1, 1)
                End If
            End If
        End If
    Next count
End Function

Private Sub UserControl_Paint()
    reDraw
End Sub

