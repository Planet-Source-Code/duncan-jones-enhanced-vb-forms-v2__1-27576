VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "APIFileTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'\\ Createion successful?
Public CreatedOK As Boolean
'\\ Member variables
Public dwLowDateTime As Long
Public dwHighDateTime As Long

'\\ Private memory handling functions
Private Declare Sub CopyMemoryFiletime Lib "kernel32" Alias "RtlMoveMemory" (Destination As FILETIME, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrFiletime Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrFiletime Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

'\\ Conversion to a system time....
Private Declare Function FileTimeToSystemTimeApi Lib "kernel32" Alias "FileTimeToSystemTime" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this Filetime object from the location poiunted to by
'\\ lpFiletime
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function CreateFromPointer(lpFileTime As Long) As Boolean

Dim ftThis As FILETIME

CreatedOK = False

If Not IsBadReadPtrFiletime(lpFileTime, Len(ftThis)) Then
    Call CopyMemoryFiletime(ftThis, lpFileTime, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            dwLowDateTime = .dwLowDateTime
            dwHighDateTime = .dwHighDateTime
            If Err.LastDllError = 0 Then
                CreatedOK = True
            End If
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function

Public Function FileTimeToSystemTime() As APISystemTime

Dim ftThis As FILETIME
Dim stThis As SYSTEMTIME
Dim oSystemTime As APISystemTime

Dim lRet As Long

With ftThis
    .dwHighDateTime = dwHighDateTime
    .dwLowDateTime = dwLowDateTime
End With

lRet = FileTimeToSystemTimeApi(ftThis, stThis)
If Err.LastDllError = 0 Then
    Set oSystemTime = New APISystemTime
    If oSystemTime.CreateFromPointer(VarPtr(stThis)) Then
        FileTimeToSystemTime = oSystemTime
    End If
End If

End Function


