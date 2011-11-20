Attribute VB_Name = "modIOleInPlaceActiveObjectChannellList"
Option Explicit

' ===========================================================================
' Filename:    mIOleInPlaceActiveObject.bas
' Author:      Mike Gainer, Matt Curland and Bill Storage
' Date:        09 January 1999
'
' Requires:    OleGuids.tlb (in IDE only)
'
' Description:
' Allows you to replace the standard IOLEInPlaceActiveObject interface for a
' UserControl with a customisable one.  This allows you to take control
' of focus in VB controls.

' The code could be adapted to replace other UserControl OLE interfaces.
'
' ---------------------------------------------------------------------------
' Visit vbAccelerator, advanced, free source for VB programmers
'     http://vbaccelerator.com
' ===========================================================================
Private Type GUID
    Data1 As Long
    data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As Long, rguid As GUID) As Long


Public Type IPAOHookStructChannelList 'IOleInPlaceActiveObjectHook
    lpVTable As Long 'VTable pointer
    IPAOReal As IOleInPlaceActiveObject 'Un-AddRefed pointer for forwarding calls
    TBEx As ctlChannelList 'Un-AddRefed native class pointer for making Friend calls
    ThisPointer As Long
End Type

Private Const strIID_IOleInPlaceActiveObject As String = "{00000117-0000-0000-C000-000000000046}"
Private IID_IOleInPlaceActiveObject As GUID
Private m_IPAOVTable(9) As Long
Private m_lpIPAOVTable As Long

Public Property Get IPAOVTableChannelList() As Long
         ' Set up the vTable for the interface and return a pointer to it:
10       If m_lpIPAOVTable = 0 Then
20           m_IPAOVTable(0) = AddressOfFunction(AddressOf QueryInterface)
30           m_IPAOVTable(1) = AddressOfFunction(AddressOf AddRef)
40           m_IPAOVTable(2) = AddressOfFunction(AddressOf Release)
50           m_IPAOVTable(3) = AddressOfFunction(AddressOf GetWindow)
60           m_IPAOVTable(4) = AddressOfFunction(AddressOf ContextSensitiveHelp)
70           m_IPAOVTable(5) = AddressOfFunction(AddressOf TranslateAccelerator)
80           m_IPAOVTable(6) = AddressOfFunction(AddressOf OnFrameWindowActivate)
90           m_IPAOVTable(7) = AddressOfFunction(AddressOf OnDocWindowActivate)
100          m_IPAOVTable(8) = AddressOfFunction(AddressOf ResizeBorder)
110          m_IPAOVTable(9) = AddressOfFunction(AddressOf EnableModeless)
120          m_lpIPAOVTable = VarPtr(m_IPAOVTable(0))
130          CLSIDFromString StrPtr(strIID_IOleInPlaceActiveObject), IID_IOleInPlaceActiveObject
140      End If
150      IPAOVTableChannelList = m_lpIPAOVTable
End Property

Private Function AddressOfFunction(lpfn As Long) As Long
         ' Work around, VB thinks lPtr = AddressOf Method is an error
10       AddressOfFunction = lpfn
End Function

Private Function AddRef(This As IPAOHookStructChannelList) As Long
         ' Call the UserControl's standard AddRef method:
10       AddRef = This.IPAOReal.AddRef
End Function

Private Function Release(This As IPAOHookStructChannelList) As Long
         ' Call the UserControl's standard Release method:
10       Release = This.IPAOReal.Release
End Function

Private Function QueryInterface(This As IPAOHookStructChannelList, riid As GUID, pvObj As Long) As Long
         ' Install the interface if required:
10       If IsEqualGUID(riid, IID_IOleInPlaceActiveObject) Then
            ' Install alternative IOleInPlaceActiveObject interface implemented here
20          pvObj = This.ThisPointer
30          AddRef This
40          QueryInterface = 0
50       Else
            ' Use the default support for the interface:
60          QueryInterface = This.IPAOReal.QueryInterface(ByVal VarPtr(riid), pvObj)
70       End If
End Function

Private Function GetWindow(This As IPAOHookStructChannelList, phwnd As Long) As Long
         ' Call user controls' GetWindow method:
10       GetWindow = This.IPAOReal.GetWindow(phwnd)
End Function

Private Function ContextSensitiveHelp(This As IPAOHookStructChannelList, ByVal fEnterMode As Long) As Long
         ' Call the user control's ContextSensitiveHelp method:
10       ContextSensitiveHelp = This.IPAOReal.ContextSensitiveHelp(fEnterMode)
End Function

Private Function TranslateAccelerator(This As IPAOHookStructChannelList, lpMsg As VBOleGuids.Msg) As Long
      Dim hRes As Long
         
         ' Check if we want to override the handling of this key code:
10       hRes = S_FALSE
20       hRes = This.TBEx.TranslateAccelerator(lpMsg)
30       If hRes Then
            ' If not pass it on to the standard UserControl TranslateAccelerator method:
40          hRes = This.IPAOReal.TranslateAccelerator(ByVal VarPtr(lpMsg))
50       End If
60       TranslateAccelerator = hRes

End Function

Private Function OnFrameWindowActivate(This As IPAOHookStructChannelList, ByVal fActivate As Long) As Long
         ' Call the user control's OnFrameWindow activate interface:
10       OnFrameWindowActivate = This.IPAOReal.OnFrameWindowActivate(fActivate)
End Function

Private Function OnDocWindowActivate(This As IPAOHookStructChannelList, ByVal fActivate As Long) As Long
         ' Call the user control's OnDocWindow activate interface:
10       OnDocWindowActivate = This.IPAOReal.OnDocWindowActivate(fActivate)
End Function

Private Function ResizeBorder(This As IPAOHookStructChannelList, prcBorder As RECT, ByVal puiWindow As _
    IOleInPlaceUIWindow, ByVal fFrameWindow As Long) As Long
         ' Call the user control's ResizeBorder interface
10       ResizeBorder = This.IPAOReal.ResizeBorder(VarPtr(prcBorder), puiWindow, fFrameWindow)
End Function

Private Function EnableModeless(This As IPAOHookStructChannelList, ByVal fEnable As Long) As Long
         ' Call the user control's EnableModeless interface
10       EnableModeless = This.IPAOReal.EnableModeless(fEnable)
End Function

Private Function IsEqualGUID(iid1 As GUID, iid2 As GUID) As Boolean
      Dim Tmp1 As Currency
      Dim Tmp2 As Currency

         ' Check for match in GUIDs.
10       If iid1.Data1 = iid2.Data1 Then
20          If iid1.data2 = iid2.data2 Then
30             If iid1.Data3 = iid2.Data3 Then
                  ' compare last 8 bytes of GUID in one chunk:
40                CopyMemory Tmp1, iid1.Data4(0), 8
50                CopyMemory Tmp2, iid2.Data4(0), 8
60                If Tmp1 = Tmp2 Then
70                   IsEqualGUID = True
80                End If
90             End If
100         End If
110      End If
         
         ' This could alternatively be done by matching the result
         ' of StringFromCLSID called on both GUIDs.

End Function




