VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITextRenderBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get left() As Long
    
End Property

Public Property Let left(newValue As Long)

End Property

Public Property Get width() As Long

End Property

Public Property Let width(newValue As Long)

End Property

Public Property Get selected() As Boolean
    
End Property

Public Sub render(drawingData As CDrawingData)
    
End Sub

Public Sub selectAll()
    
End Sub

Public Sub selectSome(left As Long, width As Long)
    
End Sub

Public Function getSelectedText(drawingData As CDrawingData, withFormatCodes As Boolean) As String

End Function

Public Sub unSelect()
    
End Sub

