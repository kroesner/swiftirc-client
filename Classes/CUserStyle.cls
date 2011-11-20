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
10        Set copy = New CUserStyle
20        copy.init m_text, m_image, m_foreColour
End Function

Public Sub init(text As String, image As CImage, foreColour As Byte)
10        m_text = text
20        Set m_image = image
          
30        If Not m_image Is Nothing Then
40            m_image.caption = m_text
50        End If
          
60        m_hasForeColour = True
70        m_foreColour = foreColour
End Sub

Public Property Get text() As String
10        text = m_text
End Property

Public Property Get image() As CImage
10        Set image = m_image
End Property

Public Property Get foreColour() As Byte
10        foreColour = m_foreColour
End Property

Public Property Let foreColour(ByVal newValue As Byte)
10        m_foreColour = newValue
End Property

Public Property Get hasForeColour() As Boolean
10        hasForeColour = m_hasForeColour
End Property

Public Property Let hasForeColour(newValue As Boolean)
10        m_hasForeColour = newValue
End Property

Public Sub buildFormat(blocks As cArrayList)
10        If settings.setting("boldNicks", estBoolean) Then
              Dim blockBold As New CBlockBold
20            blockBold.bold = True
30            blocks.Add blockBold
40        End If

50        If m_hasForeColour Then
              Dim blockForeColour As New CBlockForeColour
60            blockForeColour.hasForeColour = True
70            blockForeColour.foreColour = m_foreColour
80            blocks.Add blockForeColour
90        End If
          
100       If Not m_image Is Nothing And settings.setting("nicknameIcons", estBoolean) Then
              Dim blockImage As New CBlockImage
110           blockImage.caption = m_text
120           blockImage.image = m_image
130           blocks.Add blockImage
140       Else
              Dim blockText As New CBlockText
150           blockText.text = m_text
160           blocks.Add blockText
170       End If
End Sub

