VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITextWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub addEvent(eventName As String, params() As String)
End Sub

Public Sub addEventEx(eventName As String, userStyle As CUserStyle, username As String, flags As Long, params() As String)
End Sub

Public Sub addText(text As String)
End Sub

Public Sub addTextEx(eventColour As CEventColour, foreColour As Byte, format As String, userStyle As CUserStyle, username As String, flags As Long, params() As String)
    
End Sub

Public Property Let eventManager(newValue As CEventManager)

End Property

Public Sub update()

End Sub

Public Property Let inputText(newValue As String)
    
End Property

Public Sub focusInput()

End Sub

Public Property Get session() As CSession

End Property

Public Sub clear()

End Sub

Public Property Get textview() As swiftIrc.ctlTextView

End Property


