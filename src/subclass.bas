Attribute VB_Name = "MSubclass"
Option Explicit

' ======================================================================================
' Name:     vbAccelerator SSubTmr object
'           MSubClass.bas
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     25 June 1998
'
' Requires: None
'
' Copyright © 1998-2003 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' The implementation of the Subclassing part of the SSubTmr object.
' Use this module + ISubClass.Cls to replace dependency on the DLL.
'
' Fixes:
' 23 Jan 03
' SPM: Fixed multiple attach/detach bug which resulted in incorrectly setting
' the message count.
' SPM: Refactored code
' SPM: Added automated detach on WM_DESTROY
' 27 Dec 99
' DetachMessage: Fixed typo in DetachMessage which removed more messages than it should
'   (Thanks to Vlad Vissoultchev <wqw@bora.exco.net>)
' DetachMessage: Fixed resource leak (very slight) due to failure to remove property
'   (Thanks to Andrew Smith <asmith2@optonline.net>)
' AttachMessage: Added extra error handlers
'
' ======================================================================================


' declares:
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Const GWL_WNDPROC = (-4)
Private Const WM_DESTROY = &H2

' SubTimer is independent of VBCore, so it hard codes error handling

Public Enum EErrorWindowProc
    eeBaseWindowProc = 13080 ' WindowProc
    eeCantSubclass           ' Can't subclass window
    eeAlreadyAttached        ' Message already handled by another class
    eeInvalidWindow          ' Invalid window
    eeNoExternalWindow       ' Can't modify external window
End Enum

Private m_iCurrentMessage As Long
Private m_iProcOld As Long
Private m_f As Long


Public Property Get CurrentMessage() As Long
10       CurrentMessage = m_iCurrentMessage
End Property

Private Sub ErrRaise(e As Long)
      Dim sText As String, sSource As String
10       If e > 1000 Then
20          sSource = App.EXEName & ".WindowProc"
30          Select Case e
            Case eeCantSubclass
40             sText = "Can't subclass window"
50          Case eeAlreadyAttached
60             sText = "Message already handled by another class"
70          Case eeInvalidWindow
80             sText = "Invalid window"
90          Case eeNoExternalWindow
100            sText = "Can't modify external window"
110         End Select
120         Err.Raise e Or vbObjectError, sSource, sText
130      Else
            ' Raise standard Visual Basic error
140         Err.Raise e, sSource
150      End If
End Sub

Private Property Get MessageCount(ByVal hwnd As Long) As Long
      Dim sName As String
10       sName = "C" & hwnd
20       MessageCount = GetProp(hwnd, sName)
End Property
Private Property Let MessageCount(ByVal hwnd As Long, ByVal count As Long)
      Dim sName As String
10       m_f = 1
20       sName = "C" & hwnd
30       m_f = SetProp(hwnd, sName, count)
40       If (count = 0) Then
50          RemoveProp hwnd, sName
60       End If
70       logMessage "Changed message count for " & Hex(hwnd) & " to " & count
End Property

Private Property Get OldWindowProc(ByVal hwnd As Long) As Long
      Dim sName As String
10       sName = hwnd
20       OldWindowProc = GetProp(hwnd, sName)
End Property
Private Property Let OldWindowProc(ByVal hwnd As Long, ByVal lPtr As Long)
      Dim sName As String
10       m_f = 1
20       sName = hwnd
30       m_f = SetProp(hwnd, sName, lPtr)
40       If (lPtr = 0) Then
50          RemoveProp hwnd, sName
60       End If
70       logMessage "Changed Window Proc for " & Hex(hwnd) & " to " & Hex(lPtr)
End Property

Private Property Get MessageClassCount(ByVal hwnd As Long, ByVal iMsg As Long) As Long
      Dim sName As String
10       sName = hwnd & "#" & iMsg & "C"
20       MessageClassCount = GetProp(hwnd, sName)
End Property

Private Property Let MessageClassCount(ByVal hwnd As Long, ByVal iMsg As Long, ByVal count As Long)
      Dim sName As String
10       sName = hwnd & "#" & iMsg & "C"
20       m_f = SetProp(hwnd, sName, count)
30       If (count = 0) Then
40          RemoveProp hwnd, sName
50       End If
60       logMessage "Changed message count for " & Hex(hwnd) & " Message " & iMsg & " to " & count
End Property

Private Property Get MessageClass(ByVal hwnd As Long, ByVal iMsg As Long, ByVal index As Long) As Long
      Dim sName As String
10       sName = hwnd & "#" & iMsg & "#" & index
20       MessageClass = GetProp(hwnd, sName)
End Property
Private Property Let MessageClass(ByVal hwnd As Long, ByVal iMsg As Long, ByVal index As Long, ByVal classPtr As Long)
      Dim sName As String
10       sName = hwnd & "#" & iMsg & "#" & index
20       m_f = SetProp(hwnd, sName, classPtr)
30       If (classPtr = 0) Then
40          RemoveProp hwnd, sName
50       End If
60       logMessage "Changed message class for " & Hex(hwnd) & " Message " & iMsg & " Index " & index & " to " & Hex(classPtr)
End Property

Sub AttachMessage( _
      iwp As ISubclass, _
      ByVal hwnd As Long, _
      ByVal iMsg As Long _
   )
      Dim procOld As Long
      Dim msgCount As Long
      Dim msgClassCount As Long
      Dim msgClass As Long
          
         ' --------------------------------------------------------------------
         ' 1) Validate window
         ' --------------------------------------------------------------------
10       If IsWindow(hwnd) = False Then
20          ErrRaise eeInvalidWindow
30          Exit Sub
40       End If
50       If IsWindowLocal(hwnd) = False Then
60          ErrRaise eeNoExternalWindow
70          Exit Sub
80       End If

         ' --------------------------------------------------------------------
         ' 2) Check if this class is already attached for this message:
         ' --------------------------------------------------------------------
90       msgClassCount = MessageClassCount(hwnd, iMsg)
100      If (msgClassCount > 0) Then
110         For msgClass = 1 To msgClassCount
120            If (MessageClass(hwnd, iMsg, msgClass) = ObjPtr(iwp)) Then
130               ErrRaise eeAlreadyAttached
140               Exit Sub
150            End If
160         Next msgClass
170      End If

         ' --------------------------------------------------------------------
         ' 3) Associate this class with this message for this window:
         ' --------------------------------------------------------------------
180      MessageClassCount(hwnd, iMsg) = MessageClassCount(hwnd, iMsg) + 1
190      If (m_f = 0) Then
            ' Failed, out of memory:
200         ErrRaise 5
210         Exit Sub
220      End If
         
         ' --------------------------------------------------------------------
         ' 4) Associate the class pointer:
         ' --------------------------------------------------------------------
230      MessageClass(hwnd, iMsg, MessageClassCount(hwnd, iMsg)) = ObjPtr(iwp)
240      If (m_f = 0) Then
            ' Failed, out of memory:
250         MessageClassCount(hwnd, iMsg) = MessageClassCount(hwnd, iMsg) - 1
260         ErrRaise 5
270         Exit Sub
280      End If

         ' --------------------------------------------------------------------
         ' 5) Get the message count
         ' --------------------------------------------------------------------
290      msgCount = MessageCount(hwnd)
300      If msgCount = 0 Then
            
            ' Subclass window by installing window procedure
310         procOld = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
320         If procOld = 0 Then
               ' remove class:
330            MessageClass(hwnd, iMsg, MessageClassCount(hwnd, iMsg)) = 0
               ' remove class count:
340            MessageClassCount(hwnd, iMsg) = MessageClassCount(hwnd, iMsg) - 1
               
350            ErrRaise eeCantSubclass
360            Exit Sub
370         End If
            
            ' Associate old procedure with handle
380         OldWindowProc(hwnd) = procOld
390         If m_f = 0 Then
               ' SPM: Failed to VBSetProp, windows properties database problem.
               ' Has to be out of memory.
               
               ' Put the old window proc back again:
400            SetWindowLong hwnd, GWL_WNDPROC, procOld
               ' remove class:
410            MessageClass(hwnd, iMsg, MessageClassCount(hwnd, iMsg)) = 0
               ' remove class count:
420            MessageClassCount(hwnd, iMsg) = MessageClassCount(hwnd, iMsg) - 1
               
               ' Raise an error:
430            ErrRaise 5
440            Exit Sub
450         End If
460      End If
         
            
         ' Count this message
470      MessageCount(hwnd) = MessageCount(hwnd) + 1
480      If m_f = 0 Then
            ' SPM: Failed to set prop, windows properties database problem.
            ' Has to be out of memory
            
            ' remove class:
490         MessageClass(hwnd, iMsg, MessageClassCount(hwnd, iMsg)) = 0
            ' remove class count contribution:
500         MessageClassCount(hwnd, iMsg) = MessageClassCount(hwnd, iMsg) - 1
            
            ' If we haven't any messages on this window then remove the subclass:
510         If (MessageCount(hwnd) = 0) Then
               ' put old window proc back again:
520            procOld = OldWindowProc(hwnd)
530            If Not (procOld = 0) Then
540               SetWindowLong hwnd, GWL_WNDPROC, procOld
550               OldWindowProc(hwnd) = 0
560            End If
570         End If
            
            ' Raise the error:
580         ErrRaise 5
590         Exit Sub
600      End If
             
End Sub

Sub DetachMessage( _
      iwp As ISubclass, _
      ByVal hwnd As Long, _
      ByVal iMsg As Long _
   )
      Dim msgClassCount As Long
      Dim msgClass As Long
      Dim msgClassIndex As Long
      Dim msgCount As Long
      Dim procOld As Long
          
         ' --------------------------------------------------------------------
         ' 1) Validate window
         ' --------------------------------------------------------------------
10       If IsWindow(hwnd) = False Then
            ' for compatibility with the old version, we don't
            ' raise a message:
            ' ErrRaise eeInvalidWindow
20          Exit Sub
30       End If
40       If IsWindowLocal(hwnd) = False Then
            ' for compatibility with the old version, we don't
            ' raise a message:
            ' ErrRaise eeNoExternalWindow
50          Exit Sub
60       End If
          
         ' --------------------------------------------------------------------
         ' 2) Check if this message is attached for this class:
         ' --------------------------------------------------------------------
70       msgClassCount = MessageClassCount(hwnd, iMsg)
80       If (msgClassCount > 0) Then
90          msgClassIndex = 0
100         For msgClass = 1 To msgClassCount
110            If (MessageClass(hwnd, iMsg, msgClass) = ObjPtr(iwp)) Then
120               msgClassIndex = msgClass
130               Exit For
140            End If
150         Next msgClass
            
160         If (msgClassIndex = 0) Then
               ' fail silently
170            Exit Sub
180         Else
               ' remove this message class:
               
               ' a) Anything above this index has to be shifted up:
190            For msgClass = msgClassIndex To msgClassCount - 1
200               MessageClass(hwnd, iMsg, msgClass) = MessageClass(hwnd, iMsg, msgClass + 1)
210            Next msgClass
               
               ' b) The message class at the end can be removed:
220            MessageClass(hwnd, iMsg, msgClassCount) = 0
               
               ' c) Reduce the message class count:
230            MessageClassCount(hwnd, iMsg) = MessageClassCount(hwnd, iMsg) - 1
               
240         End If
            
250      Else
            ' fail silently
260         Exit Sub
270      End If
         
         ' ---------------------------------------------------------------------
         ' 3) Reduce the message count:
         ' ---------------------------------------------------------------------
280      msgCount = MessageCount(hwnd)
290      If (msgCount = 1) Then
            ' remove the subclass:
300         procOld = OldWindowProc(hwnd)
310         If Not (procOld = 0) Then
               ' Unsubclass by reassigning old window procedure
320            Call SetWindowLong(hwnd, GWL_WNDPROC, procOld)
330         End If
            ' remove the old window proc:
340         OldWindowProc(hwnd) = 0
350      End If
360      MessageCount(hwnd) = MessageCount(hwnd) - 1
         
End Sub

Private Function WindowProc( _
      ByVal hwnd As Long, _
      ByVal iMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long _
   ) As Long
         
      Dim procOld As Long
      Dim msgClassCount As Long
      Dim bCalled As Boolean
      Dim pSubClass As Long
      Dim iwp As ISubclass
      Dim iwpT As ISubclass
      Dim iIndex As Long
      Dim bDestroy As Boolean
          
         ' Get the old procedure from the window
10       procOld = OldWindowProc(hwnd)
20       Debug.Assert procOld <> 0
          
30       If (procOld = 0) Then
            ' we can't work, we're not subclassed properly.
40          Exit Function
50       End If
          
         ' SPM - in this version I am allowing more than one class to
         ' make a subclass to the same hWnd and Msg.  Why am I doing
         ' this?  Well say the class in question is a control, and it
         ' wants to subclass its container.  In this case, we want
         ' all instances of the control on the form to receive the
         ' form notification message.
          
         ' Get the number of instances for this msg/hwnd:
60       bCalled = False
         
70       If (MessageClassCount(hwnd, iMsg) > 0) Then
80          iIndex = MessageClassCount(hwnd, iMsg)
            
90          Do While (iIndex >= 1)
100            pSubClass = MessageClass(hwnd, iMsg, iIndex)
               
110            If (pSubClass = 0) Then
                  ' Not handled by this instance
120            Else
                  ' Turn pointer into a reference:
130               CopyMemory iwpT, pSubClass, 4
140               Set iwp = iwpT
150               CopyMemory iwpT, 0&, 4
                  
                  ' Store the current message, so the client can check it:
160               m_iCurrentMessage = iMsg
                  
170               With iwp
                     ' Preprocess (only checked first time around):
180                  If (iIndex = 1) Then
190                     If (.MsgResponse = emrPreprocess) Then
200                        If Not (bCalled) Then
210                           WindowProc = CallWindowProc(procOld, hwnd, iMsg, _
                                                        wParam, ByVal lParam)
220                           bCalled = True
230                        End If
240                     End If
250                  End If
                     ' Consume (this message is always passed to all control
                     ' instances regardless of whether any single one of them
                     ' requests to consume it):
260                  WindowProc = .WindowProc(hwnd, iMsg, wParam, ByVal lParam)
270               End With
280            End If
               
290            iIndex = iIndex - 1
300         Loop
            
            ' PostProcess (only check this the last time around):
310         If Not (iwp Is Nothing) And Not (procOld = 0) Then
320             If iwp.MsgResponse = emrPostProcess Then
330                If Not (bCalled) Then
340                   WindowProc = CallWindowProc(procOld, hwnd, iMsg, _
                                                wParam, ByVal lParam)
350                   bCalled = True
360                End If
370             End If
380         End If
                  
390      Else
            ' Not handled:
400         If (iMsg = WM_DESTROY) Then
               ' If WM_DESTROY isn't handled already, we should
               ' clear up any subclass
410            pClearUp hwnd
420            WindowProc = CallWindowProc(procOld, hwnd, iMsg, _
                                          wParam, ByVal lParam)
               
430         Else
440            WindowProc = CallWindowProc(procOld, hwnd, iMsg, _
                                          wParam, ByVal lParam)
450         End If
460      End If
          
End Function
Public Function CallOldWindowProc( _
      ByVal hwnd As Long, _
      ByVal iMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long _
   ) As Long
      Dim iProcOld As Long
10       iProcOld = OldWindowProc(hwnd)
20       If Not (iProcOld = 0) Then
30          CallOldWindowProc = CallWindowProc(iProcOld, hwnd, iMsg, wParam, lParam)
40       End If
End Function

Function IsWindowLocal(ByVal hwnd As Long) As Boolean
          Dim idWnd As Long
10        Call GetWindowThreadProcessId(hwnd, idWnd)
20        IsWindowLocal = (idWnd = GetCurrentProcessId())
End Function

Private Sub logMessage(ByVal sMsg As String)
10       Debug.Print sMsg
End Sub


Private Sub pClearUp(ByVal hwnd As Long)
      Dim msgCount As Long
      Dim procOld As Long
         ' this is only called if you haven't explicitly cleared up
         ' your subclass from the caller.  You will get a minor
         ' resource leak as it does not clear up any message
         ' specific properties.
10       msgCount = MessageCount(hwnd)
20       If (msgCount > 0) Then
            ' remove the subclass:
30          procOld = OldWindowProc(hwnd)
40          If Not (procOld = 0) Then
               ' Unsubclass by reassigning old window procedure
50             Call SetWindowLong(hwnd, GWL_WNDPROC, procOld)
60          End If
            ' remove the old window proc:
70          OldWindowProc(hwnd) = 0
80          MessageCount(hwnd) = 0
90       End If
End Sub
