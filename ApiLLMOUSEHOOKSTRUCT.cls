VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiLLMOUSEHOOKSTRUCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type LLMOUSEHOOKSTRUCT
    ptX As Long 'MouseHookStruct
    ptY As Long
    MouseData As Long
    Flags As Long
End Type

Public pt As APIPoint
Public MouseData As Long
Public Flags As Long

Public CreatedOK As Boolean

'\\ Private memory handling functions
Private Declare Sub CopyMemoryLLMouseHookStruct Lib "kernel32" Alias "RtlMoveMemory" (Destination As LLMOUSEHOOKSTRUCT, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrLLMouseHookStruct Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrLLMouseHookStruct Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this MouseHookStruct object from the location poiunted to by
'\\ lpMouseHookStruct
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for corMouseHookStructness
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function CreateFromPointer(lpllMouseHookStruct As Long) As Boolean

Dim ftThis As LLMOUSEHOOKSTRUCT

CreatedOK = False

If Not IsBadReadPtrLLMouseHookStruct(lpllMouseHookStruct, Len(ftThis)) Then
    Call CopyMemoryLLMouseHookStruct(ftThis, lpllMouseHookStruct, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            Flags = .Flags
            MouseData = .MouseData
            Set pt = New APIPoint
            pt.x = .ptX
            pt.y = .ptY
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function


