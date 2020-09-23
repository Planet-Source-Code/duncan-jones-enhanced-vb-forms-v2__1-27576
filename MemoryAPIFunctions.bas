Attribute VB_Name = "MemoryAPIFunctions"
'\\ -----[MemoryAPIFunctions]----------------------------------------------------
'\\ For copying structures (and strings) to/from pointers
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------

Option Explicit

'\\ Memory manipulation routines
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'\\ Pointer validation in StringFromPointer
Private Declare Function IsBadStringPtrByLong Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long









'\\ --[StringFromPointer]-------------------------------------------------------------------
'\\ Returns a VB string from an API returned string pointer
'\\ Parameters:
'\\   lpString - The long pointer to the string
'\\   lMaxlength - the size of empty buffer to allow
'\\ HISTORY:
'\\  DEJ 28/02/2001 Check pointer is a valid string pointer...
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function StringFromPointer(lpString As Long, lMaxLength As Long) As String

Dim sRet As String
Dim lRet As Long

If IsBadStringPtrByLong(lpString, lMaxLength) Then
    '\\ An error has occured - do not attempt to use this pointer
    Call ReportError(Err.LastDllError, "StringFromPointer", "Attempt to read bad string pointer: " & lpString)
    StringFromPointer = ""
    Exit Function
End If

'\\ Pre-initialise the return string...
sRet = Space$(lMaxLength)
CopyMemory ByVal sRet, ByVal lpString, ByVal Len(sRet)
If Err.LastDllError = 0 Then
    If InStr(sRet, Chr$(0)) > 0 Then
        sRet = Left$(sRet, InStr(sRet, Chr$(0)) - 1)
    End If
End If

StringFromPointer = sRet

End Function

