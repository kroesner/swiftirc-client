VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColourManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colours(0 To 6) As Long
Private m_brushes(0 To 5) As Long
Private m_pens(3) As Long

Private Sub initColours()
    m_colours(SWIFTCOLOUR_WINDOW) = RGB(0, 0, 0)
    m_colours(SWIFTCOLOUR_CONTROLBACK) = 3289650
    m_colours(SWIFTCOLOUR_CONTROLFORE) = 11513775
    m_colours(SWIFTCOLOUR_CONTROLFOREOVER) = 16777215
    m_colours(SWIFTCOLOUR_CONTROLBORDER) = 8355711
    m_colours(SWIFTCOLOUR_FRAMEBACK) = 4934475
    m_colours(SWIFTCOLOUR_FRAMEBORDER) = 4934475
End Sub

Private Sub updatePens()
    Dim count As Long
    
    For count = 0 To UBound(m_pens)
        If m_pens(count) <> 0 Then
            DeleteObject m_pens(count)
        End If
    Next count
    
    m_pens(0) = CreatePen(PS_SOLID, 1, Me.getColour(SWIFTCOLOUR_CONTROLBORDER))
    m_pens(1) = CreatePen(PS_SOLID, 1, Me.getColour(SWIFTCOLOUR_FRAMEBACK))
    m_pens(2) = CreatePen(PS_SOLID, 2, Me.getColour(SWIFTCOLOUR_CONTROLBORDER))
    m_pens(3) = CreatePen(PS_SOLID, 1, Me.getColour(SWIFTCOLOUR_FRAMEBORDER))
End Sub

Private Sub updateBrushes()
    Dim count As Integer
    
    For count = 0 To UBound(m_brushes)
        If m_brushes(count) <> 0 Then
            DeleteObject m_brushes(count)
        End If
        
        m_brushes(count) = CreateSolidBrush(m_colours(count))
    Next count
End Sub

Public Function getColour(index As Long) As Long
    If index <= UBound(m_colours) Then
        getColour = m_colours(index)
    End If
End Function

Public Sub setColour(index As Long, newColour As Long)
    m_colours(index) = newColour
    updateBrushes
    updatePens
End Sub


Public Function getBrush(index As Long) As Long
    If index <= UBound(m_brushes) Then
        getBrush = m_brushes(index)
    End If
End Function

Public Function getPen(index As Long) As Long
    If index <= UBound(m_pens) Then
        getPen = m_pens(index)
    End If
End Function

Private Sub Class_Initialize()
    initColours
    updateBrushes
    updatePens
End Sub
