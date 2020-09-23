VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiCBTACTIVATESTRUCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type CBTACTIVATESTRUCT
     fMouse As Long
     hWndActive As Long
End Type

Public fMouse As Long
Public hWndActive As Long
Public CreatedOK As Boolean

'\\ Private memory handling functions
Private Declare Sub CopyMemoryCBTActivateStruct Lib "kernel32" Alias "RtlMoveMemory" (Destination As CBTACTIVATESTRUCT, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrCBTActivateStruct Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrCBTActivateStruct Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this CBTActivateStruct object from the location poiunted to by
'\\ lpCBTActivateStruct
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function CreateFromPointer(lpCBTActivateStruct As Long) As Boolean

Dim ftThis As CBTACTIVATESTRUCT

CreatedOK = False

If Not IsBadReadPtrCBTActivateStruct(lpCBTActivateStruct, Len(ftThis)) Then
    Call CopyMemoryCBTActivateStruct(ftThis, lpCBTActivateStruct, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            fMouse = .fMouse
            hWndActive = .hWndActive
            If Err.LastDllError = 0 Then
                CreatedOK = True
            End If
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function
