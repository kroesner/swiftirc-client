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
10        m_colours(SWIFTCOLOUR_WINDOW) = RGB(0, 0, 0)
20        m_colours(SWIFTCOLOUR_CONTROLBACK) = 3289650
30        m_colours(SWIFTCOLOUR_CONTROLFORE) = 11513775
40        m_colours(SWIFTCOLOUR_CONTROLFOREOVER) = 16777215
50        m_colours(SWIFTCOLOUR_CONTROLBORDER) = 8355711
60        m_colours(SWIFTCOLOUR_FRAMEBACK) = 4934475
70        m_colours(SWIFTCOLOUR_FRAMEBORDER) = 4934475
End Sub

Private Sub updatePens()
          Dim count As Long
          
10        For count = 0 To UBound(m_pens)
20            If m_pens(count) <> 0 Then
30                DeleteObject m_pens(count)
40            End If
50        Next count
          
60        m_pens(0) = CreatePen(PS_SOLID, 1, Me.getColour(SWIFTCOLOUR_CONTROLBORDER))
70        m_pens(1) = CreatePen(PS_SOLID, 1, Me.getColour(SWIFTCOLOUR_FRAMEBACK))
80        m_pens(2) = CreatePen(PS_SOLID, 2, Me.getColour(SWIFTCOLOUR_CONTROLBORDER))
90        m_pens(3) = CreatePen(PS_SOLID, 1, Me.getColour(SWIFTCOLOUR_FRAMEBORDER))
End Sub

Private Sub updateBrushes()
          Dim count As Integer
          
10        For count = 0 To UBound(m_brushes)
20            If m_brushes(count) <> 0 Then
30                DeleteObject m_brushes(count)
40            End If
              
50            m_brushes(count) = CreateSolidBrush(m_colours(count))
60        Next count
End Sub

Public Function getColour(index As Long) As Long
10        If index <= UBound(m_colours) Then
20            getColour = m_colours(index)
30        End If
End Function

Public Sub setColour(index As Long, newColour As Long)
10        m_colours(index) = newColour
20        updateBrushes
30        updatePens
End Sub


Public Function getBrush(index As Long) As Long
10        If index <= UBound(m_brushes) Then
20            getBrush = m_brushes(index)
30        End If
End Function

Public Function getPen(index As Long) As Long
10        If index <= UBound(m_pens) Then
20            getPen = m_pens(index)
30        End If
End Function

Private Sub Class_Initialize()
10        initColours
20        updateBrushes
30        updatePens
End Sub
