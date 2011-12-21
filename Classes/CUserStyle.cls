VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUserStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_hasForeColour As Boolean
Private m_foreColour As Byte
Private m_text As String
Private m_image As CImage

Public Function copy() As CUserStyle
    Set copy = New CUserStyle
    copy.init m_text, m_image, m_foreColour
End Function

Public Sub init(text As String, image As CImage, foreColour As Byte)
    m_text = text
    Set m_image = image
    
    If Not m_image Is Nothing Then
        m_image.caption = m_text
    End If
    
    m_hasForeColour = True
    m_foreColour = foreColour
End Sub

Public Property Get text() As String
    text = m_text
End Property

Public Property Get image() As CImage
    Set image = m_image
End Property

Public Property Get foreColour() As Byte
    foreColour = m_foreColour
End Property

Public Property Let foreColour(ByVal newValue As Byte)
    m_foreColour = newValue
End Property

Public Property Get hasForeColour() As Boolean
    hasForeColour = m_hasForeColour
End Property

Public Property Let hasForeColour(newValue As Boolean)
    m_hasForeColour = newValue
End Property

Public Sub buildFormat(blocks As cArrayList)
    If settings.setting("boldNicks", estBoolean) Then
        Dim blockBold As New CBlockBold
        blockBold.bold = True
        blocks.Add blockBold
    End If

    If m_hasForeColour Then
        Dim blockForeColour As New CBlockForeColour
        blockForeColour.hasForeColour = True
        blockForeColour.foreColour = m_foreColour
        blocks.Add blockForeColour
    End If
    
    If Not m_image Is Nothing And settings.setting("nicknameIcons", estBoolean) Then
        Dim blockImage As New CBlockImage
        blockImage.caption = m_text
        blockImage.image = m_image
        blocks.Add blockImage
    Else
        Dim blockText As New CBlockText
        blockText.text = m_text
        blocks.Add blockText
    End If
End Sub

