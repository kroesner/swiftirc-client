VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CImageManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_images As New cArrayList
Private m_rootPath As String

Public Property Get rootPath() As String
10        rootPath = m_rootPath
End Property

Public Property Let rootPath(newValue As String)
10        m_rootPath = newValue
End Property

Public Function addImage(imagePath As String) As CImage
          Dim newImage As New CImage
          
10        If newImage.load(m_rootPath & imagePath) = False Then
20            Exit Function
30        End If
          
40        m_images.Add newImage
50        Set addImage = newImage
End Function

Public Function getImage(index As Integer) As CImage
10        If index < 1 Or index > m_images.count Then
20            Exit Function
30        End If
          
40        Set getImage = m_images.item(index)
End Function

Public Sub clearImages()
10        m_images.clear
End Sub

Private Sub Class_Terminate()
10        m_images.clear
End Sub
