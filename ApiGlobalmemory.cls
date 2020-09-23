VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiGlobalmemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'\\ Global memory management functions
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private mMyData() As Byte
Private mMyDataSize As Long
Private mHmem As Long


Public Enum enGlobalmemoryAllocationConstants
    GMEM_FIXED = &H0
    GMEM_DISCARDABLE = &H100
    GMEM_MOVEABLE = &H2
    GMEM_NOCOMPACT = &H10
    GMEM_NODISCARD = &H20
    GMEM_ZEROINIT = &H40
End Enum

Private mAllocationType As enGlobalmemoryAllocationConstants

Public Property Let AllocationType(ByVal newType As enGlobalmemoryAllocationConstants)

mAllocationType = newType

End Property

Public Property Get AllocationType() As enGlobalmemoryAllocationConstants

    AllocationType = mAllocationType
    
End Property


Private Sub CopyDataToGlobal()

Dim lRet As Long

If mHmem > 0 Then
    lRet = GlobalLock(mHmem)
    If lRet > 0 Then
        Call CopyMemory(ByVal mHmem, mMyData(0), mMyDataSize)
        Call GlobalUnlock(mHmem)
    End If
End If

End Sub

Public Sub CopyFromHandle(ByVal hMemHandle As Long)

Dim lRet As Long
Dim lPtr As Long

lRet = GlobalSize(hMemHandle)
If lRet > 0 Then
    mMyDataSize = lRet
    lPtr = GlobalLock(hMemHandle)
    If lPtr > 0 Then
        ReDim mMyData(0 To mMyDataSize - 1) As Byte
        CopyMemory mMyData(0), ByVal lPtr, mMyDataSize
        Call GlobalUnlock(hMemHandle)
    End If
End If

End Sub

Public Sub CopyToHandle(ByVal hMemHandle As Long)

Dim lSize As Long
Dim lPtr As Long

'\\ Don't copy if its empty
If Not Me.IsEmpty Then
    lSize = GlobalSize(hMemHandle)
    '\\ Don't attempt to copy if zero size...
    If lSize > 0 Then
        If lPtr > 0 Then
            CopyMemory ByVal lPtr, mMyData(0), lSize
            Call GlobalUnlock(hMemHandle)
        End If
    End If
End If

End Sub


'\\ --[Handle]------------------------------------------------------
'\\ Returns a Global Memroy handle that is valid and filled with the
'\\ info held in this object's private byte array
'\\ ----------------------------------------------------------------
Public Property Get Handle() As Long

If mHmem = 0 Then
    If mMyDataSize > 0 Then
        mHmem = GlobalAlloc(AllocationType, mMyDataSize)
    End If
End If
Call CopyDataToGlobal

Handle = mHmem

End Property

Public Property Get IsEmpty() As Boolean

    IsEmpty = (mMyDataSize = 0)

End Property

Public Sub Free()

If mHmem > 0 Then
    Call GlobalFree(mHmem)
    mHmem = 0
    mMyDataSize = 0
    ReDim mMyData(0) As Byte
End If

End Sub

Private Sub Class_Terminate()

If mHmem > 0 Then
    Call GlobalFree(mHmem)
End If

End Sub


