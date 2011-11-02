VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_socket As MSWinsockLib.Winsock
Attribute m_socket.VB_VarHelpID = -1

Public Event readable()
Public Event connected()
Public Event connectionRequest(requestId As Long)
Public Event disconnected()
Public Event error(number As Integer, description As String)
Public Event sendComplete()

Public Property Get socket() As MSWinsockLib.Winsock
    Set socket = m_socket
End Property

Public Property Let socket(newSock As MSWinsockLib.Winsock)
    Set m_socket = newSock
End Property

Public Sub connect(host As String, port As Long)
    If m_socket.state <> sckClosed Then
        m_socket.Close
    End If
    
    m_socket.connect host, port
End Sub

Public Sub acceptConnection(requestId As Long)
    If m_socket.state <> sckClosed Then
        m_socket.Close
    End If
    
    m_socket.Accept requestId
End Sub

Public Sub listen(port As String, Optional host As String)
    m_socket.Bind port, host
    m_socket.listen
End Sub

Public Sub receive(ByRef data As String)
    m_socket.GetData data
End Sub

Public Sub send(data As String)
    If m_socket.state = sckConnected Then
        m_socket.SendData data
    End If
End Sub

Public Sub sendLine(line As String)
    Me.send line & vbCrLf
End Sub

Public Sub closeSocket()
    If m_socket.state <> sckClosed Then
        m_socket.Close
    End If
End Sub

Private Sub m_socket_Close()
    m_socket.Close
    RaiseEvent disconnected
End Sub

Private Sub m_socket_DataArrival(ByVal bytesTotal As Long)
    RaiseEvent readable
End Sub

Private Sub m_socket_Connect()
    RaiseEvent connected
End Sub

Private Sub m_socket_ConnectionRequest(ByVal requestId As Long)
    RaiseEvent connectionRequest(requestId)
End Sub

Private Sub m_socket_error(ByVal number As Integer, description As String, ByVal Scode As Long, _
    ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As _
    Boolean)
    RaiseEvent error(number, description)
End Sub

Public Property Get connected() As Boolean
    connected = (m_socket.state = sckConnected)
End Property

Private Sub m_socket_SendComplete()
    RaiseEvent sendComplete
End Sub
