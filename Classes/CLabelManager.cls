VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLabelManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_labels As New Collection

Public Function addLabel(caption As String, labelType As eLabelType, x As Long, y As Long) As CLabel
          Dim newLabel As New CLabel
          
10        newLabel.init caption, labelType, x, y
20        m_labels.Add newLabel
          
30        Set addLabel = newLabel
End Function

Public Sub renderLabels(hdc As Long)
          Dim aLabel As CLabel
          
10        SetBkMode hdc, OPAQUE
          
20        For Each aLabel In m_labels
30            If aLabel.visible Then
40                aLabel.render hdc
50            End If
60        Next aLabel
End Sub
