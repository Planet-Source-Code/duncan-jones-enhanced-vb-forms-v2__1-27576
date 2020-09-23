VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiNotifyIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'\\ To put this app in the Task Tray:
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public CreatedOK As Boolean

'\\ Private memory handling functions
Private Declare Sub CopyMemoryNotifyIconData Lib "kernel32" Alias "RtlMoveMemory" (Destination As NOTIFYICONDATA, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrNotifyIconData Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrNotifyIconData Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

Public Enum enShellNotifyMessages
    NIM_ADD = &H0
    NIM_MODIFY = &H1
    NIM_DELETE = &H2
End Enum

Public Enum enShellNotifyFormats
    NIF_MESSAGE = &H1
    NIF_ICON = &H2
    NIF_TIP = &H4
End Enum

'\\ member variables
Private mszTip As String
Private mcbSize As Long
Private mhwnd As Long
Private muID As Long
Private muFlags As Long
Private mhIcon As Long
Private muCallbackMessage As Long
Private mAdded As Boolean

Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private mMessage As Long

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this NotifyIconData object from the location poiunted to by
'\\ lpPointAPI
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for corPointAPIness
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function CreateFromPointer(lpNotifyIconData As Long) As Boolean

Dim niThis As NOTIFYICONDATA

CreatedOK = False

If Not IsBadReadPtrNotifyIconData(lpNotifyIconData, Len(niThis)) Then
    Call CopyMemoryNotifyIconData(niThis, lpNotifyIconData, Len(niThis))
    If Err.LastDllError = 0 Then
        With niThis
            
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function

Public Property Set Icon(ByVal newIcon As ApiIcon)

If mhIcon <> newIcon.hIcon Then
    mhIcon = newIcon.hIcon
End If

End Property

Public Property Get Icon() As ApiIcon

Dim myIcon As ApiIcon

Set myIcon = New ApiIcon
myIcon.hIcon = mhIcon
Set Icon = myIcon

End Property

Public Property Set NotifyWindow(ByVal wndNotify As ApiWindow)

    mhwnd = wndNotify.hwnd
    
End Property

Public Property Get NotifyWindowMessage() As Long

    If mMessage = 0 Then
        mMessage = RegisterWindowMessage("APINOTIFYICONMSG")
        If Err.LastDllError > 0 Then
            ReportError Err.LastDllError, "ApiNotifyIcon:NotifyWindowMessage", APIDispenser.LastSystemError
        End If
    End If
    NotifyWindowMessage = mMessage
    
End Property

Public Sub RefreshNotifyIcon()

Dim niThis As NOTIFYICONDATA
Dim lRet As Long

'\\ Initialise the message to use if not already set...
If muCallbackMessage = 0 Then
    muCallbackMessage = Me.NotifyWindowMessage
End If

If mAdded Then
    With niThis
        .hIcon = mhIcon
        .hwnd = mhwnd
        .szTip = mszTip
        .uCallbackMessage = muCallbackMessage
        .uFlags = muFlags
        .uID = muID
        .cbSize = Len(niThis)
    End With
    
    lRet = Shell_NotifyIcon(NIM_MODIFY, niThis)
    If Err.LastDllError > 0 Then
        ReportError Err.LastDllError, "ApiNotifyIcon:RefreshNotifyIcon", APIDispenser.LastSystemError
    End If
End If

End Sub

'\\ --[SetNotifyIcon]--------------------------------------------------------
'\\ Sets the icon (and other info) from this ApiNotifyIcon
'\\ into the system tray
'\\ -------------------------------------------------------------------------
Public Sub SetNotifyIcon()

Dim niThis As NOTIFYICONDATA
Dim lRet As Long

'\\ Initialise the message to use if not already set...
If muCallbackMessage = 0 Then
    muCallbackMessage = Me.NotifyWindowMessage
End If

With niThis
    .hIcon = mhIcon
    .hwnd = mhwnd
    .szTip = mszTip
    .uCallbackMessage = muCallbackMessage
    .uFlags = muFlags
    .uID = muID
    .cbSize = Len(niThis)
End With

lRet = Shell_NotifyIcon(NIM_ADD, niThis)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiNotifyIcon:SetNotifyIcon", APIDispenser.LastSystemError
Else
    mAdded = True
End If

End Sub

Public Property Get ToolTip() As String

    ToolTip = mszTip
    
End Property

Public Property Let ToolTip(ByVal NewTip As String)

    '\\ Limit tootltip to 63 chars + null terminator...
    If Len(NewTip) > 63 Then
        NewTip = Left$(NewTip, 63)
    End If
    
    '\\ Initialise all to nulls
    mszTip = NewTip & String$(64 - Len(NewTip), 0)

    
End Property


Public Sub UnsetNotifyIcon()

Dim niThis As NOTIFYICONDATA
Dim lRet As Long

'\\ Initialise the message to use if not already set...
If muCallbackMessage = 0 Then
    muCallbackMessage = Me.NotifyWindowMessage
End If

With niThis
    .hIcon = mhIcon
    .hwnd = mhwnd
    .szTip = mszTip
    .uCallbackMessage = muCallbackMessage
    .uFlags = muFlags
    .uID = muID
    .cbSize = Len(niThis)
End With

lRet = Shell_NotifyIcon(NIM_DELETE, niThis)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiNotifyIcon:UnsetNotifyIcon", APIDispenser.LastSystemError
Else
    mAdded = False
End If

End Sub


Private Sub Class_Initialize()

muFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE


End Sub


