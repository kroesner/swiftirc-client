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
    rootPath = m_rootPath
End Property

Public Property Let rootPath(newValue As String)
    m_rootPath = newValue
End Property

Public Function addImage(imagePath As String) As CImage
    Dim newImage As New CImage
    
    If newImage.load(m_rootPath & imagePath) = False Then
        Exit Function
    End If
    
    m_images.Add newImage
    Set addImage = newImage
End Function

Public Function getImage(index As Integer) As CImage
    If index < 1 Or index > m_images.count Then
        Exit Function
    End If
    
    Set getImage = m_images.item(index)
End Function

Public Sub clearImages()
    m_images.clear
End Sub

Private Sub Class_Terminate()
    m_images.clear
End Sub
