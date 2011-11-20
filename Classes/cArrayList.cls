VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' cIndexCollection
'
' Implements a collection intended to be accessed by
' Index using an array.  This fixes the performance
' problems when inserting or removing items from the
' middle of the array by using the CopyMemory to
' shift all the existing items in a similar chunk.
'

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private m_lItem() As Long
Private m_lCount As Long
Private m_lArraySize As Long
Private m_lAllocationSize As Long

Public Property Get AllocationSize() As Long
10       AllocationSize = m_lAllocationSize
End Property
Public Property Let AllocationSize(ByVal lAllocationSize As Long)
10       m_lAllocationSize = lAllocationSize
End Property

Public Property Get item(ByVal index As Long) As Object
         ' Return the item at Index.  Note since this
         ' array does not automatically reduce size we
         ' need to check that index is in bounds:
10       If index > 0 And index <= m_lCount Then
20          Set item = ObjectFromPtr(m_lItem(index))
30       Else
            ' Subscript out of range
      '      Err.Raise 9
40       End If
End Property
Public Property Let item(ByVal index As Long, ByVal value As Object)
      Dim iU As IShellFolderEx_TLB.IUnknown

10       If index > 0 And index <= m_lCount Then
20          If (m_lItem(index) <> 0) Then
30             Set iU = ObjectFromPtr(m_lItem(index))
40             iU.Release
50             Set iU = Nothing
60          End If
         
            ' Add a reference to the cData object so it
            ' won't be freed until the list item
            ' is removed:
70          Set iU = value
80          iU.AddRef
90          Set iU = Nothing
100         m_lItem(index) = ObjPtr(value)
110      Else
            ' Subscript out of range
120         Err.Raise 9
130      End If
End Property
Public Sub Add(ByVal value As Object, Optional ByVal index As Long = -1)
      Dim j As Long
      Dim lTemp As Long

      Dim iU As IShellFolderEx_TLB.IUnknown

10       If index <= 0 Or index > m_lCount Then
            ' Add to end
20          m_lCount = m_lCount + 1
30          If m_lCount > m_lArraySize Then
40             m_lArraySize = m_lCount
50             ReDim Preserve m_lItem(1 To m_lArraySize) As Long
60          End If
70          Set iU = value
80          iU.AddRef
90          Set iU = Nothing
100         m_lItem(m_lCount) = ObjPtr(value)
110      Else
120         lTemp = m_lItem(m_lCount)
            ' Increase array size as required:
130         m_lCount = m_lCount + 1
140         If m_lCount > m_lArraySize Then
150            m_lArraySize = m_lArraySize + m_lAllocationSize
160            ReDim Preserve m_lItem(1 To m_lArraySize) As Long
170         End If
            ' Shift up from Index to m_lCount-1
180         CopyMemory m_lItem(index + 1), m_lItem(index), (m_lCount - index) * 4
            ' Increase array size:
            ' Set item at Index
190         Set iU = value
200         iU.AddRef
210         Set iU = Nothing
220         m_lItem(index) = ObjPtr(value)
            ' Set last item:
230         m_lItem(m_lCount) = lTemp
            
240      End If
         
End Sub
Public Sub Remove(Optional ByVal index As Long = -1)
      Dim j As Long
         ' Reduce size:
10       If index <= 0 Or index >= m_lCount Then
20          If (m_lCount > 0) Then
               ' Call the object's IUnknown_Release method
               ' using the TLB
               Dim iU As IShellFolderEx_TLB.IUnknown
30             Set iU = ObjectFromPtr(m_lItem(m_lCount))
40             iU.Release
50             Set iU = Nothing
60          End If
70          m_lCount = m_lCount - 1
80          If m_lCount < 0 Then
90             m_lCount = 0
100         End If
110      Else
120         If (index > 0) And (index <= m_lCount) Then
               ' Call the object's IUnknown_Release method
               ' using the TLB
130            Set iU = ObjectFromPtr(m_lItem(index))
140            iU.Release
150            Set iU = Nothing
160         End If
170         m_lCount = m_lCount - 1
180         If m_lCount > 0 Then
               ' shift down from Index to m_lCount-1
190            CopyMemory m_lItem(index), m_lItem(index + 1), (m_lCount - index + 1) * 4
200         Else
               ' the end
210            m_lCount = 0
220         End If
230      End If
         
End Sub
Public Sub HeapMinimize()
         ' Reduce the array storage size to
         ' match the number of items in it:
10       If m_lArraySize > m_lCount Then
20          If m_lCount <= 0 Then
30             Erase m_lItem
40             m_lArraySize = 0
50          Else
60             ReDim Preserve m_lItem(1 To m_lCount) As Long
70             m_lArraySize = m_lCount
80          End If
90       End If
End Sub
Public Property Get count() As Long
         ' Number of items in the array:
10       count = m_lCount
End Property
Public Property Get Exists(ByVal index As Long)
         ' Does the item at Index exist?
10       Exists = (index > 0 And index <= m_lCount)
End Property
Public Sub clear()
         ' Release any object pointers:
         Dim i As Long
         Dim iU As IShellFolderEx_TLB.IUnknown

10       For i = 1 To m_lCount
            ' Call the object's IUnknown_Release method
            ' using the TLB
20          Set iU = ObjectFromPtr(m_lItem(i))
30          iU.Release
40          Set iU = Nothing
50       Next i
         
         ' Clear down the array:
60       m_lCount = 0
End Sub

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
      Dim objT As Object
         ' Bruce McKinney's code for getting an Object from the
         ' object pointer:
10       CopyMemory objT, lPtr, 4
20       Set ObjectFromPtr = objT
30       CopyMemory objT, 0&, 4
End Property

Private Sub Class_Initialize()
10       m_lAllocationSize = 1
End Sub


Private Sub Class_Terminate()
10       clear
End Sub
