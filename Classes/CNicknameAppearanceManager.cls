VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNicknameAppearanceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_naItems As New cArrayList

Public Sub addItemIndirect(na As CNicknameAppearance)
    m_naItems.Add na
End Sub

Public Property Get naCount() As Long
    naCount = m_naItems.count
End Property

Public Property Get naItem(index As Long) As CNicknameAppearance
    naItem = m_naItems.item(index)
End Property

Public Sub copyItems(target As cArrayList)
    Dim count As Long
    
    For count = 1 To m_naItems.count
        target.Add m_naItems.item(count).copy
    Next count
End Sub
