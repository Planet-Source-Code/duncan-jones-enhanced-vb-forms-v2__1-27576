VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiMOUSEHOOKSTRUCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type MOUSEHOOKSTRUCT
    ptX As Long 'MouseHookStruct
    ptY As Long
    hwnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type

Public pt As APIPoint
Public hwnd As Long
Public wHitTestCode As Long
Public dwExtraInfo As Long

Public CreatedOK As Boolean

'\\ Private memory handling functions
Private Declare Sub CopyMemoryMouseHookStruct Lib "kernel32" Alias "RtlMoveMemory" (Destination As MOUSEHOOKSTRUCT, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrMouseHookStruct Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrMouseHookStruct Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this MouseHookStruct object from the location poiunted to by
'\\ lpMouseHookStruct
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for corMouseHookStructness
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function CreateFromPointer(lpMouseHookStruct As Long) As Boolean

Dim ftThis As MOUSEHOOKSTRUCT

CreatedOK = False

If Not IsBadReadPtrMouseHookStruct(lpMouseHookStruct, Len(ftThis)) Then
    Call CopyMemoryMouseHookStruct(ftThis, lpMouseHookStruct, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            dwExtraInfo = .dwExtraInfo
            hwnd = .hwnd
            wHitTestCode = .wHitTestCode
            Set pt = New APIPoint
            pt.x = .ptX
            pt.y = .ptY
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function

