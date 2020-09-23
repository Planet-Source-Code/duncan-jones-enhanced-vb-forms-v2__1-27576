VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiCWPRETSTRUCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type CWPRETSTRUCT
    lResult As Long
    lParam As Long
    wParam As Long
    Message As Long
    hwnd As Long
End Type

Public lResult As Long
Public lParam As Long
Public wParam As Long
Public Message As Long
Public hwnd As Long

Public CreatedOK As Boolean

'\\ Private memory handling functions
Private Declare Sub CopyMemoryCWPRetStruct Lib "kernel32" Alias "RtlMoveMemory" (Destination As CWPRETSTRUCT, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrCWPRetStruct Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrCWPRetStruct Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long


'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this CWPRetStructobject from the location poiunted to by
'\\ lpCWPRetStruct
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function CreateFromPointer(lpCWPRetStruct As Long) As Boolean

Dim ftThis As CWPRETSTRUCT

CreatedOK = False

If Not IsBadReadPtrCWPRetStruct(lpCWPRetStruct, Len(ftThis)) Then
    Call CopyMemoryCWPRetStruct(ftThis, lpCWPRetStruct, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            hwnd = .hwnd
            lParam = .lParam
            lResult = .lResult
            Message = .Message
            wParam = .wParam
            If Err.LastDllError = 0 Then
                CreatedOK = True
            End If
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function

