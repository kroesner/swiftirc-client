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
          
10        prefix = parsePrefix(caption)
          
20        aTab.init caption, prefix
30        m_tabs.Add aTab
40        Set addTab = aTab
          
50        reDraw
End Function

Private Sub reDraw()
          Dim count As Long
          Dim x As Long
          Dim y As Long
          
          Dim backBuffer As Long
          Dim backBitmap As Long
          Dim oldBitmap As Long
          Dim oldFont As Long
          
10        backBuffer = CreateCompatibleDC(UserControl.hdc)
20        backBitmap = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, _
              UserControl.ScaleHeight)
          
30        oldBitmap = SelectObject(backBuffer, backBitmap)
40        oldFont = SelectObject(backBuffer, GetCurrentObject(hdc, OBJ_FONT))
          
50        FillRect backBuffer, makeRect(0, UserControl.ScaleWidth, 0, UserControl.ScaleHeight), _
              colourManager.getBrush(SWIFTCOLOUR_WINDOW)
          
60        y = 5
          
70        For count = 1 To m_tabs.count
80            m_tabs.item(count).render backBuffer, x, y, TAB_WIDTH, TAB_HEIGHT
90            x = x + TAB_WIDTH + TAB_SPACING_X
100       Next count
          
110       BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, backBuffer, 0, 0, _
              vbSrcCopy
              
120       SelectObject backBuffer, oldBitmap
130       SelectObject backBuffer, oldFont
          
140       DeleteDC backBuffer
150       DeleteObject backBitmap
End Sub

Private Property Get IWindow_realWindow() As Object
10        Set IWindow_realWindow = m_realWindow
End Property

Private Property Let IWindow_realWindow(RHS As Object)
10        Set m_realWindow = RHS
End Property

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
          Dim count As Long
          Dim char As String
          
10        char = Chr$(KeyAscii)
          
20        For count = 1 To m_tabs.count
30            If LCase$(m_tabs.item(count).prefix) = LCase$(char) Then
40                If Not m_selectedTab Is Nothing Then
50                    m_selectedTab.state = tisNormal
60                End If
                  
70                Set m_selectedTab = m_tabs.item(count)
80                m_selectedTab.state = tisSelected
90                RaiseEvent tabSelected(m_selectedTab)
100               reDraw
                  
110               Exit Sub
120           End If
130       Next count
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
10        If y < 5 Or y > UserControl.ScaleHeight Or x < 0 Or x > UserControl.ScaleWidth Then
20            If Not m_overTab Is Nothing Then
30                If Not m_overTab.state = tisSelected Then
40                    m_overTab.state = tisNormal
50                    reDraw
60                End If
                  
70                Set m_overTab = Nothing
80            End If
              
90            ReleaseCapture
100           Exit Sub
110       End If
          
          Dim tabIndex As Long
          Dim aTab As CTabStripItem
          
120       tabIndex = Fix(x / (TAB_WIDTH + TAB_SPACING_X)) + 1
          
130       If tabIndex > 0 And tabIndex <= m_tabs.count Then
140           Set aTab = m_tabs.item(tabIndex)
              
150           If aTab Is m_overTab Then
160               Exit Sub
170           End If
              
180           If Not m_overTab Is Nothing Then
190               If Not m_selectedTab Is m_overTab Then
200                   m_overTab.state = tisNormal
210               End If
220           End If
              
230           Set m_overTab = aTab
              
240           If Not aTab.state = tisSelected Then
250               aTab.state = tisMouseOver
260           End If
              
270           reDraw
280           SetCapture UserControl.hwnd
290       Else
300           If Not m_overTab Is Nothing Then
310               If Not m_overTab.state = tisSelected Then
320                   m_overTab.state = tisNormal
330                   reDraw
340               End If
                  
350               Set m_overTab = Nothing
360           End If
370       End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
10        If y < 5 Then
20            Exit Sub
30        End If
          
          Dim tabIndex As Long
          Dim aTab As CTabStripItem
          
40        tabIndex = Fix(x / (TAB_WIDTH + TAB_SPACING_X)) + 1
          
50        If tabIndex > 0 And tabIndex <= m_tabs.count Then
60            Set aTab = m_tabs.item(tabIndex)
70            selectTab aTab, False
80        End If
End Sub

Public Sub selectTab(aTab As CTabStripItem, noEvent As Boolean)
10        If m_selectedTab Is aTab Then
20            Exit Sub
30        End If
          
40        If Not m_selectedTab Is Nothing Then
50            m_selectedTab.state = tisNormal
60        End If
          
70        Set m_selectedTab = aTab
80        m_selectedTab.state = tisSelected
          
90        reDraw
          
100       If Not noEvent Then
110           RaiseEvent tabSelected(m_selectedTab)
120       End If
End Sub

Private Function parsePrefix(caption As String) As String
          Dim count As Long
          
10        For count = 1 To Len(caption)
20            If Mid$(caption, count, 1) = "&" Then
30                If count < Len(caption) Then
40                    If Mid$(caption, count + 1, 1) <> "&" Then
50                        UserControl.AccessKeys = UserControl.AccessKeys & Mid$(caption, count + 1, 1)
60                        parsePrefix = Mid$(caption, count + 1, 1)
70                    End If
80                End If
90            End If
100       Next count
End Function

Private Sub UserControl_Paint()
10        reDraw
End Sub

Private Sub UserControl_Terminate()
10        debugLog "ctlTabStrip terminating"
End Sub
