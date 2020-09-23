VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiDEBUGHOOKINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type DEBUGHOOKINFO
    hModuleHook As Long
    Reserved As Long
    lParam As Long
    wParam As Long
    code As Long
End Type

Public hModuleHook As Long
Public Reserved As Long
Public lParam As Long
Public wParam As Long
Public code As Long

Public CreatedOK As Boolean

'\\ Private memory handling functions
Private Declare Sub CopyMemoryDebugHookInfo Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEBUGHOOKINFO, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrDebugHookInfo Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrDebugHookInfo Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this DebugHookInfo object from the location poiunted to by
'\\ lpDebugHookInfo
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function CreateFromPointer(lpDebugHookInfo As Long) As Boolean

Dim ftThis As DEBUGHOOKINFO

CreatedOK = False

If Not IsBadReadPtrDebugHookInfo(lpDebugHookInfo, Len(ftThis)) Then
    Call CopyMemoryDebugHookInfo(ftThis, lpDebugHookInfo, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            code = .code
            hModuleHook = .hModuleHook
            lParam = .lParam
            Reserved = .Reserved
            wParam = .wParam
            If Err.LastDllError = 0 Then
                CreatedOK = True
            End If
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function

