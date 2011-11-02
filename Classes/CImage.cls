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
    If m_image Then
        DeleteObject m_image
    End If
    
    m_image = swiftLoadImage(ByVal 0&, path, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
    
    If m_image = 0 Then
        load = False
        Exit Function
    End If
    
    Dim imageInfo As BITMAP
    
    GetObject m_image, Len(imageInfo), imageInfo
    
    m_width = imageInfo.bmWidth
    m_height = imageInfo.bmHeight
    
    load = True
End Function

Public Sub draw(hdc As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As _
    Long)
    Dim tempDc As Long
    Dim tempDc2 As Long
    Dim tempBitmap As Long
    Dim oldBitmap As Long
    Dim oldBitmap2 As Long
    
    tempDc = CreateCompatibleDC(hdc)
    oldBitmap = SelectObject(tempDc, m_image)
    
    tempDc2 = CreateCompatibleDC(hdc)
    tempBitmap = CreateCompatibleBitmap(hdc, width, height)
    oldBitmap2 = SelectObject(tempDc2, tempBitmap)
    
    StretchBlt tempDc2, 0, 0, width, height, tempDc, 0, 0, m_width, m_height, vbSrcCopy
    TransparentBlt hdc, x, y, width, height, tempDc2, 0, 0, width, height, RGB(255, 255, 255)
    
    SelectObject tempDc, oldBitmap
    SelectObject tempDc2, oldBitmap2
    
    DeleteDC tempDc
    DeleteDC tempDc2
    DeleteObject tempBitmap
End Sub

Public Sub drawInverted(hdc As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As _
    Long)
    Dim tempDc As Long
    Dim tempDc2 As Long
    Dim tempBitmap As Long
    Dim oldBitmap As Long
    Dim oldBitmap2 As Long
    
    tempDc = CreateCompatibleDC(hdc)
    oldBitmap = SelectObject(tempDc, m_image)
    
    tempDc2 = CreateCompatibleDC(hdc)
    tempBitmap = CreateCompatibleBitmap(hdc, width, height)
    oldBitmap2 = SelectObject(tempDc2, tempBitmap)
    
    StretchBlt tempDc2, 0, 0, width, height, tempDc, 0, 0, m_width, m_height, vbNotSrcCopy
    TransparentBlt hdc, x, y, width, height, tempDc2, 0, 0, width, height, RGB(255, 255, 255)
    
    SelectObject tempDc, oldBitmap
    SelectObject tempDc2, oldBitmap2
    
    DeleteDC tempDc
    DeleteDC tempDc2
    DeleteObject tempBitmap
End Sub


Public Sub drawAlpha(hdc As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, backBrush As Long)
    Dim tempDc As Long
    Dim tempBitmap As Long
    Dim oldBitmap As Long
    
    tempDc = CreateCompatibleDC(hdc)
    tempBitmap = CreateCompatibleBitmap(hdc, width, height)
    
    oldBitmap = SelectObject(tempDc, tempBitmap)
    
    FillRect tempDc, makeRect(0, width, 0, height), backBrush
    
    Me.draw tempDc, 0, 0, width, height
    AlphaBlend hdc, x, y, width, height, tempDc, 0, 0, width, height, m_blendFunc
    
    SelectObject tempDc, oldBitmap
    
    DeleteDC tempDc
    DeleteObject tempBitmap
End Sub

Public Property Get image() As Long
    image = m_image
End Property

Public Property Get width() As Long
    width = m_width
End Property

Public Property Get height() As Long
    height = m_height
End Property

Public Property Get caption() As String
    caption = m_caption
End Property

Public Property Let caption(newValue As String)
    m_caption = newValue
End Property

Private Sub Class_Initialize()
    Dim BlendFunc As BLENDFUNCTION
    
    BlendFunc.BlendOp = AC_SRC_OVER
    BlendFunc.BlendFlags = 0
    BlendFunc.SourceConstantAlpha = 128
    BlendFunc.AlphaFormat = 0

    CopyMemory m_blendFunc, BlendFunc, 4
End Sub

Private Sub Class_Terminate()
    If m_image Then
        DeleteObject m_image
    End If
End Sub
