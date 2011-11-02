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
   AllocationSize = m_lAllocationSize
End Property
Public Property Let AllocationSize(ByVal lAllocationSize As Long)
   m_lAllocationSize = lAllocationSize
End Property

Public Property Get item(ByVal index As Long) As Object
   ' Return the item at Index.  Note since this
   ' array does not automatically reduce size we
   ' need to check that index is in bounds:
   If index > 0 And index <= m_lCount Then
      Set item = ObjectFromPtr(m_lItem(index))
   Else
      ' Subscript out of range
'      Err.Raise 9
   End If
End Property
Public Property Let item(ByVal index As Long, ByVal value As Object)
Dim iU As IShellFolderEx_TLB.IUnknown

   If index > 0 And index <= m_lCount Then
      If (m_lItem(index) <> 0) Then
         Set iU = ObjectFromPtr(m_lItem(index))
         iU.Release
         Set iU = Nothing
      End If
   
      ' Add a reference to the cData object so it
      ' won't be freed until the list item
      ' is removed:
      Set iU = value
      iU.AddRef
      Set iU = Nothing
      m_lItem(index) = ObjPtr(value)
   Else
      ' Subscript out of range
      Err.Raise 9
   End If
End Property
Public Sub Add(ByVal value As Object, Optional ByVal index As Long = -1)
Dim j As Long
Dim lTemp As Long

Dim iU As IShellFolderEx_TLB.IUnknown

   If index <= 0 Or index > m_lCount Then
      ' Add to end
      m_lCount = m_lCount + 1
      If m_lCount > m_lArraySize Then
         m_lArraySize = m_lCount
         ReDim Preserve m_lItem(1 To m_lArraySize) As Long
      End If
      Set iU = value
      iU.AddRef
      Set iU = Nothing
      m_lItem(m_lCount) = ObjPtr(value)
   Else
      lTemp = m_lItem(m_lCount)
      ' Increase array size as required:
      m_lCount = m_lCount + 1
      If m_lCount > m_lArraySize Then
         m_lArraySize = m_lArraySize + m_lAllocationSize
         ReDim Preserve m_lItem(1 To m_lArraySize) As Long
      End If
      ' Shift up from Index to m_lCount-1
      CopyMemory m_lItem(index + 1), m_lItem(index), (m_lCount - index) * 4
      ' Increase array size:
      ' Set item at Index
      Set iU = value
      iU.AddRef
      Set iU = Nothing
      m_lItem(index) = ObjPtr(value)
      ' Set last item:
      m_lItem(m_lCount) = lTemp
      
   End If
   
End Sub
Public Sub Remove(Optional ByVal index As Long = -1)
Dim j As Long
   ' Reduce size:
   If index <= 0 Or index >= m_lCount Then
      If (m_lCount > 0) Then
         ' Call the object's IUnknown_Release method
         ' using the TLB
         Dim iU As IShellFolderEx_TLB.IUnknown
         Set iU = ObjectFromPtr(m_lItem(m_lCount))
         iU.Release
         Set iU = Nothing
      End If
      m_lCount = m_lCount - 1
      If m_lCount < 0 Then
         m_lCount = 0
      End If
   Else
      If (index > 0) And (index <= m_lCount) Then
         ' Call the object's IUnknown_Release method
         ' using the TLB
         Set iU = ObjectFromPtr(m_lItem(index))
         iU.Release
         Set iU = Nothing
      End If
      m_lCount = m_lCount - 1
      If m_lCount > 0 Then
         ' shift down from Index to m_lCount-1
         CopyMemory m_lItem(index), m_lItem(index + 1), (m_lCount - index + 1) * 4
      Else
         ' the end
         m_lCount = 0
      End If
   End If
   
End Sub
Public Sub HeapMinimize()
   ' Reduce the array storage size to
   ' match the number of items in it:
   If m_lArraySize > m_lCount Then
      If m_lCount <= 0 Then
         Erase m_lItem
         m_lArraySize = 0
      Else
         ReDim Preserve m_lItem(1 To m_lCount) As Long
         m_lArraySize = m_lCount
      End If
   End If
End Sub
Public Property Get count() As Long
   ' Number of items in the array:
   count = m_lCount
End Property
Public Property Get Exists(ByVal index As Long)
   ' Does the item at Index exist?
   Exists = (index > 0 And index <= m_lCount)
End Property
Public Sub clear()
   ' Release any object pointers:
   Dim i As Long
   Dim iU As IShellFolderEx_TLB.IUnknown

   For i = 1 To m_lCount
      ' Call the object's IUnknown_Release method
      ' using the TLB
      Set iU = ObjectFromPtr(m_lItem(i))
      iU.Release
      Set iU = Nothing
   Next i
   
   ' Clear down the array:
   m_lCount = 0
End Sub

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim objT As Object
   ' Bruce McKinney's code for getting an Object from the
   ' object pointer:
   CopyMemory objT, lPtr, 4
   Set ObjectFromPtr = objT
   CopyMemory objT, 0&, 4
End Property

Private Sub Class_Initialize()
   m_lAllocationSize = 1
End Sub


Private Sub Class_Terminate()
   clear
End Sub
