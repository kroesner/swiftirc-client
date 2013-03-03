VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSwiftSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISubclass

Private Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal _
    level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As _
    Long
Private Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal _
    cmd As Long, ByRef argp As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal versionRequested As _
    Integer, lpWSAData As WSAData) As Long
Private Declare Function WSAAsyncGetHostByName Lib "ws2_32.dll" (ByVal hwnd As _
    Long, ByVal wMsg As Long, ByVal addr As String, hent As HostEntAsync, ByVal _
    bufLen As Long) As Long
Private Declare Function WSAAsyncGetHostByAddr Lib "ws2_32.dll" (ByVal hwnd As _
    Long, ByVal wMsg As Long, addr As Long, ByVal addrLen As Long, ByVal addrType _
    As Long, hent As HostEntAsync, ByVal bufLen As Long) As Long
Private Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal socket As Long, _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long

Private Declare Function APISocket Lib "ws2_32.dll" Alias "socket" (ByVal af As _
    Long, ByVal sockType As Long, ByVal protocol As Long) As Long
Private Declare Function APICloseSocket Lib "ws2_32.dll" Alias "closesocket" _
    (ByVal socket As Long) As Long
Private Declare Function APIConnect Lib "ws2_32.dll" Alias "connect" (ByVal s _
    As Long, name As sockaddr_in, ByVal namelen As Long) As Long
Private Declare Function APIRecv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, _
    ByVal buf As Long, ByVal bufLen As Long, ByVal flags As Long) As Long
Private Declare Function APISend Lib "ws2_32.dll" Alias "send" (ByVal s As Long, _
    ByVal buf As String, ByVal bufLen As Long, ByVal flags As Long) As Long

Private Declare Function WSANtohl Lib "ws2_32.dll" (ByVal socket As Long, ByVal _
    netlong As Long, iphostlong As Long) As Long
Private Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Private Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Long) As _
    Integer

Private Declare Function strLen Lib "kernel32" Alias "lstrlenA" (lpString As _
    Any) As Long

Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long

Private Const SOCKET_ERROR = -1
Private Const INVALID_SOCKET = -1

Private Const WSAEWOULDBLOCK = 10035

Private Const WSADESCRIPTION_LEN = 256 + 1
Private Const WSASYS_STATUS_LEN = 128 + 1

Private Const AF_INET As Long = 2
Private Const SOCK_STREAM As Long = 1
Private Const IPPROTO_TCP As Long = 6

Private Const FD_READ As Long = &H1
Private Const FD_WRITE As Long = &H2
Private Const FD_CONNECT As Long = &H10
Private Const FD_CLOSE As Long = &H20

Private Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type in_addr
    S_addr As Long
End Type

Private Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As in_addr
    sin_zero(0 To 7) As Byte
End Type

Private Type HostEntAsync
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
    hentbuf(1024) As Byte
End Type

Private Type LINGERSTRUCT
    l_onoff As Integer
    l_linger As Integer
End Type

Private Const SOL_SOCKET = &HFFFF&
Private Const SO_LINGER = &H80&
Private Const FIONBIO = &H8004667E

Private Const WSM_RESOLVECOMPLETE_A As Long = WM_USER + 1
Private Const WSM_RESOLVECOMPLETE_PTR As Long = WM_USER + 2
Private Const WSM_SOCKET_ACTIVITY As Long = WM_USER + 3

Public Event resolveFailed(requestId As Long, errorCode As Long)
Public Event connectFailed(errorCode As Long)
Public Event closed(errorCode As Long)

Public Event resolveCompleteA(requestId As Long, addrList As cArrayList)
Public Event resolveCompletePtr(requestId As Long, host As String)
Public Event connected()
Public Event readable()

Private m_initialized As Boolean
Private m_connected As Boolean
Private m_socket As Long

Private m_hostEnt As HostEntAsync

Private m_buffer As String

Private m_hwnd As Long
Private wsData As WSAData

Public Sub init(hwnd As Long)
    Dim ret As Long

    ret = WSAStartup(MakeWord(2, 2), wsData)
    
    If ret Then
        Exit Sub
    End If
    
    m_hwnd = hwnd
    initMessages
    
    m_initialized = True
End Sub

Public Sub deInit()
    deInitMessages
    closeSocket False
    m_initialized = False
End Sub

Public Function resolveAddrByName(host As String) As Long
    resolveAddrByName = WSAAsyncGetHostByName(m_hwnd, WSM_RESOLVECOMPLETE_A, host, m_hostEnt, 1024)
End Function

Public Function resolveNameByAddr(ip As String) As Long
    resolveNameByAddr = WSAAsyncGetHostByAddr(m_hwnd, WSM_RESOLVECOMPLETE_PTR, inet_addr(ip), 4, 2, m_hostEnt, 1024)
End Function

Public Property Get lastError() As Long
    lastError = Err.LastDllError
End Property

Public Sub connect(ByVal addr As Long, ByVal port As Long)
    m_connected = False
    m_socket = APISocket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    
    If m_socket = INVALID_SOCKET Then
        Exit Sub
    End If
    
    setSelect False
    
    Dim sockaddr As sockaddr_in
    
    sockaddr.sin_family = AF_INET
    sockaddr.sin_addr.S_addr = addr
    sockaddr.sin_port = htons(port)
    
    APIConnect m_socket, sockaddr, Len(sockaddr)
End Sub

Public Sub closeSocket(Optional hard As Boolean = False)
    If m_socket = INVALID_SOCKET Then
        Exit Sub
    End If

    stopAsyncSelect

    If Not hard Then
        If LenB(m_buffer) <> 0 Then
            If m_connected Then
                flushBuffer
            End If
        End If
    End If
    
    APICloseSocket m_socket
    m_socket = INVALID_SOCKET

    m_connected = False
End Sub

Private Sub setSelect(wantWrite As Boolean)
    Dim activityFlags As Long
    
    activityFlags = FD_READ Or FD_CONNECT Or FD_CLOSE
    
    If wantWrite Then
        activityFlags = activityFlags Or FD_WRITE
    End If
    
    WSAAsyncSelect m_socket, m_hwnd, WSM_SOCKET_ACTIVITY, activityFlags
End Sub

Private Sub stopAsyncSelect()
    WSAAsyncSelect m_socket, m_hwnd, 0, 0
End Sub

Public Function sockRead(ByRef buf As String, Optional readLen As Long = 4096) As Long
    buf = Space$(readLen)
    buf = StrConv(buf, vbFromUnicode)
    
    sockRead = APIRecv(m_socket, StrPtr(buf), readLen, 0)
    
    If sockRead > 0 Then
        buf = Mid$(buf, 1, sockRead / 2)
        buf = UTF8Decode(StrConv(buf, vbUnicode))
    Else
        buf = vbNullString
    
        If lastError = WSAEWOULDBLOCK Then
            'Fake success
            sockRead = 1
        End If
    End If
End Function

Public Sub send(ByRef buf As String)
    sockWrite buf
End Sub

Public Sub sendLine(ByRef buf As String)
    sockWrite buf & vbCrLf
End Sub

Public Sub sockWrite(ByRef buf As String)
    If LenB(buf) = 0 Then
        Exit Sub
    End If

    If LenB(m_buffer) = 0 Then
        setSelect True
    End If

    m_buffer = m_buffer & UTF8Encode(buf)
End Sub

Private Sub flushBuffer()
    Dim result As Long
    
    Do While LenB(m_buffer) <> 0
        result = APISend(m_socket, m_buffer, Len(m_buffer), 0)
        
        If result > 0 Then
            If result = Len(m_buffer) Then
                m_buffer = vbNullString
            Else
                m_buffer = Mid$(m_buffer, result)
            End If
        Else
            If result = SOCKET_ERROR Then
                If lastError <> WSAEWOULDBLOCK Then
                    m_connected = False
                    closeSocket
                    RaiseEvent closed(lastError)
                End If
            End If
            
            Exit Do
        End If
    Loop
End Sub

Private Sub initMessages()
    If m_hwnd = 0 Then
        Exit Sub
    End If

    AttachMessage Me, m_hwnd, WSM_RESOLVECOMPLETE_A
    AttachMessage Me, m_hwnd, WSM_RESOLVECOMPLETE_PTR
    AttachMessage Me, m_hwnd, WSM_SOCKET_ACTIVITY
End Sub

Private Sub deInitMessages()
    If m_hwnd = 0 Then
        Exit Sub
    End If

    DetachMessage Me, m_hwnd, WSM_RESOLVECOMPLETE_A
    DetachMessage Me, m_hwnd, WSM_RESOLVECOMPLETE_PTR
    DetachMessage Me, m_hwnd, WSM_SOCKET_ACTIVITY
End Sub

Private Sub Class_Initialize()
    m_socket = INVALID_SOCKET
End Sub

Private Sub Class_Terminate()
    deInit
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    Select Case CurrentMessage
        Case WSM_RESOLVECOMPLETE_A
            ISubclass_MsgResponse = emrConsume
        Case WSM_RESOLVECOMPLETE_PTR
            ISubclass_MsgResponse = emrConsume
        Case WSM_SOCKET_ACTIVITY
            ISubclass_MsgResponse = emrConsume
        Case Else
            ISubclass_MsgResponse = emrPostProcess
    End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim error As Long
    
    Select Case iMsg
        Case WSM_RESOLVECOMPLETE_A
            resolveCompleteA wParam, lParam
        Case WSM_RESOLVECOMPLETE_PTR
            resolveCompletePtr wParam, lParam
        Case WSM_SOCKET_ACTIVITY
            If wParam = m_socket Then
                error = HiWord(lParam)
            
                Select Case LoWord(lParam)
                    Case FD_READ
                        If error Then
                            m_connected = False
                            closeSocket
                            RaiseEvent closed(error)
                        Else
                            RaiseEvent readable
                        End If
                    Case FD_WRITE
                        If error Then
                            m_connected = False
                            closeSocket
                            RaiseEvent closed(error)
                        Else
                            flushBuffer
                            
                            If LenB(m_buffer) = 0 Then
                                setSelect False
                            End If
                        End If
                    Case FD_CONNECT
                        If error Then
                            RaiseEvent connectFailed(error)
                        Else
                            If Not m_connected Then
                                m_connected = True
                                RaiseEvent connected
                            End If
                        End If
                    Case FD_CLOSE
                        m_connected = False
                        closeSocket
                        RaiseEvent closed(error)
                End Select
            End If
    End Select
End Function

Private Sub resolveCompleteA(wParam As Long, lParam As Long)
    Dim listPtr As Long
    Dim addrPtr As Long
    Dim addr As Long
    
    Dim list As cArrayList
    Dim addrObj As CIpAddr

    If HiWord(lParam) <> 0 Then
        RaiseEvent resolveFailed(wParam, HiWord(lParam))
        Exit Sub
    End If
    
    listPtr = m_hostEnt.hAddrList
    CopyMemory addrPtr, ByVal listPtr, 4
    
    Set list = New cArrayList
    
    Do While addrPtr <> 0
        CopyMemory addr, ByVal addrPtr, 4
        
        Set addrObj = New CIpAddr
        addrObj.addr = addr
        list.Add addrObj
        
        listPtr = listPtr + 4
        CopyMemory addrPtr, ByVal listPtr, 4
    Loop
    
    RaiseEvent resolveCompleteA(wParam, list)
End Sub

Private Sub resolveCompletePtr(wParam As Long, lParam As Long)
    Dim host As String
    Dim hostLen As Long

    If HiWord(lParam) <> 0 Then
        RaiseEvent resolveFailed(wParam, HiWord(lParam))
        Exit Sub
    End If
    
    hostLen = strLen(ByVal m_hostEnt.hName)
    
    host = Space$(hostLen)
    CopyMemory ByVal host, ByVal m_hostEnt.hName, hostLen
    
    RaiseEvent resolveCompletePtr(wParam, host)
End Sub
