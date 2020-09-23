VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiEVENTMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type EVENTMSG
    Message As Long
    paramL As Long
    paramH As Long
    time As Long
    hwnd As Long
End Type

Public Message As Long
Public paramL As Long
Public paramH As Long
Public time As Long
Public hwnd As Long

Public CreatedOK As Boolean
'\\ Private memory handling functions
Private Declare Sub CopyMemoryEventMsg Lib "kernel32" Alias "RtlMoveMemory" (Destination As EVENTMSG, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrEventMsg Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrEventMsg Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this EventMsg object from the location poiunted to by
'\\ lpEventMsg
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function CreateFromPointer(lpEventMsg As Long) As Boolean

Dim ftThis As EVENTMSG

CreatedOK = False

If Not IsBadReadPtrEventMsg(lpEventMsg, Len(ftThis)) Then
    Call CopyMemoryEventMsg(ftThis, lpEventMsg, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            hwnd = .hwnd
            Message = .Message
            paramH = .paramH
            paramL = .paramL
            time = .time
            If Err.LastDllError = 0 Then
                CreatedOK = True
            End If
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function
