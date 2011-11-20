VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_image As Long
Private m_width As Integer
Private m_height As Integer
Private m_blendFunc As Long
Private m_caption As String

Public Function load(path As String) As Boolean
10        If m_image Then
20            DeleteObject m_image
30        End If
          
40        m_image = swiftLoadImage(ByVal 0&, path, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
          
50        If m_image = 0 Then
60            load = False
70            Exit Function
80        End If
          
          Dim imageInfo As BITMAP
          
90        GetObject m_image, Len(imageInfo), imageInfo
          
100       m_width = imageInfo.bmWidth
110       m_height = imageInfo.bmHeight
          
120       load = True
End Function

Public Sub draw(hdc As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As _
    Long)
          Dim tempDc As Long
          Dim tempDc2 As Long
          Dim tempBitmap As Long
          Dim oldBitmap As Long
          Dim oldBitmap2 As Long
          
10        tempDc = CreateCompatibleDC(hdc)
20        oldBitmap = SelectObject(tempDc, m_image)
          
30        tempDc2 = CreateCompatibleDC(hdc)
40        tempBitmap = CreateCompatibleBitmap(hdc, width, height)
50        oldBitmap2 = SelectObject(tempDc2, tempBitmap)
          
60        StretchBlt tempDc2, 0, 0, width, height, tempDc, 0, 0, m_width, m_height, vbSrcCopy
70        TransparentBlt hdc, x, y, width, height, tempDc2, 0, 0, width, height, RGB(255, 255, 255)
          
80        SelectObject tempDc, oldBitmap
90        SelectObject tempDc2, oldBitmap2
          
100       DeleteDC tempDc
110       DeleteDC tempDc2
120       DeleteObject tempBitmap
End Sub

Public Sub drawInverted(hdc As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As _
    Long)
          Dim tempDc As Long
          Dim tempDc2 As Long
          Dim tempBitmap As Long
          Dim oldBitmap As Long
          Dim oldBitmap2 As Long
          
10        tempDc = CreateCompatibleDC(hdc)
20        oldBitmap = SelectObject(tempDc, m_image)
          
30        tempDc2 = CreateCompatibleDC(hdc)
40        tempBitmap = CreateCompatibleBitmap(hdc, width, height)
50        oldBitmap2 = SelectObject(tempDc2, tempBitmap)
          
60        StretchBlt tempDc2, 0, 0, width, height, tempDc, 0, 0, m_width, m_height, vbNotSrcCopy
70        TransparentBlt hdc, x, y, width, height, tempDc2, 0, 0, width, height, RGB(255, 255, 255)
          
80        SelectObject tempDc, oldBitmap
90        SelectObject tempDc2, oldBitmap2
          
100       DeleteDC tempDc
110       DeleteDC tempDc2
120       DeleteObject tempBitmap
End Sub


Public Sub drawAlpha(hdc As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, backBrush As Long)
          Dim tempDc As Long
          Dim tempBitmap As Long
          Dim oldBitmap As Long
          
10        tempDc = CreateCompatibleDC(hdc)
20        tempBitmap = CreateCompatibleBitmap(hdc, width, height)
          
30        oldBitmap = SelectObject(tempDc, tempBitmap)
          
40        FillRect tempDc, makeRect(0, width, 0, height), backBrush
          
50        Me.draw tempDc, 0, 0, width, height
60        AlphaBlend hdc, x, y, width, height, tempDc, 0, 0, width, height, m_blendFunc
          
70        SelectObject tempDc, oldBitmap
          
80        DeleteDC tempDc
90        DeleteObject tempBitmap
End Sub

Public Property Get image() As Long
10        image = m_image
End Property

Public Property Get width() As Long
10        width = m_width
End Property

Public Property Get height() As Long
10        height = m_height
End Property

Public Property Get caption() As String
10        caption = m_caption
End Property

Public Property Let caption(newValue As String)
10        m_caption = newValue
End Property

Private Sub Class_Initialize()
          Dim BlendFunc As BLENDFUNCTION
          
10        BlendFunc.BlendOp = AC_SRC_OVER
20        BlendFunc.BlendFlags = 0
30        BlendFunc.SourceConstantAlpha = 128
40        BlendFunc.AlphaFormat = 0

50        CopyMemory m_blendFunc, BlendFunc, 4
End Sub

Private Sub Class_Terminate()
10        If m_image Then
20            DeleteObject m_image
30        End If
End Sub
