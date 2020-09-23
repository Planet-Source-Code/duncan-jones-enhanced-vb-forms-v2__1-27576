VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "APIWndClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type WndClass
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As Long '\\ String conversion after api calls
    lpszClassName As Long '\\ String conversion afer api calls
End Type

Private Type WNDCLASSEX
    cbSize As Long
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As Long
    lpszClassName As Long
    hIconSm As Long
End Type

'\\ Creation successful?
Public CreatedOK As Boolean

'\\ Members
Public style As Long
Public cbClsExtra As Long
Public cbWndExtra2 As Long
Public hInstance As Long
'Public hIcon As Long
Public hCursor As Long
Public hbrBackground As Long
Public lpszMenuName As String
Public lpszClassName As String

Private m_lpfnwndproc As Long
Private mIcon As ApiIcon

'\\ Private APIs
Private Declare Function RegisterClassExApi Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
Private Declare Function UnregisterClassApi Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long


'\\ Private memory handling functions
Private Declare Sub CopyMemoryWndClass Lib "kernel32" Alias "RtlMoveMemory" (Destination As WndClass, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrWndclass Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrWndclass Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long


Private Declare Function GetClassInfoApi Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WndClass) As Long

Public Enum enStandardWindowClasses
    SWC_BUTTON = 1 '"BUTTON"
    SWC_COMBOBOX = 2 '"COMBOBOX"
    SWC_EDIT = 3 '"EDIT"
    SWC_LISTBOX = 4 '"LISTBOX"
    SWC_SCROLLBAR = 5 '"SCROLLBAR"
End Enum

Public Enum enCommonControlClasses
    CCC_BUTTONSLISTBOX = 1
    CCC_HOTKEY_CLASS = 2
    CCC_PROGRESS_CLASS = 3
    CCC_STATUSCLASSNAME = 4
    CCC_RICHEDIT = 5
    CCC_TOOLBARCLASSNAME = 6
    CCC_TOOLTIPS_CLASS = 7
    CCC_TRACKBARCLASS = 8
    CCC_UPDOWN_CLASS = 9
    CCC_WC_HEADER = 10
    CCC_WC_LISTVIEW = 11
    CCC_WC_TABCONTROL = 12
    CCC_WC_TREEVIEW = 13
End Enum

Private Function CommonControlClassname(ByVal ClassType As enCommonControlClasses) As String

Select Case ClassType
Case CCC_BUTTONSLISTBOX
    CommonControlClassname = "BUTTONSLISTBOX"
Case CCC_HOTKEY_CLASS
    CommonControlClassname = "HOTKEY_CLASS"
Case CCC_PROGRESS_CLASS
    CommonControlClassname = "PROGRESS_CLASS"
Case CCC_STATUSCLASSNAME
    CommonControlClassname = "STATUSCLASSNAME"
Case CCC_RICHEDIT
    CommonControlClassname = "RICHEDIT"
Case CCC_TOOLBARCLASSNAME
    CommonControlClassname = "TOOLBARCLASSNAME"
Case CCC_TOOLTIPS_CLASS
    CommonControlClassname = "TOOLTIPS_CLASS"
Case CCC_TRACKBARCLASS
    CommonControlClassname = "TRACKBARCLASS"
Case CCC_UPDOWN_CLASS
    CommonControlClassname = "UPDOWN_CLASS"
Case CCC_WC_HEADER
    CommonControlClassname = "WC_HEADER"
Case CCC_WC_LISTVIEW
    CommonControlClassname = "WC_LISTVIEW"
Case CCC_WC_TABCONTROL
    CommonControlClassname = "WC_TABCONTROL"
Case CCC_WC_TREEVIEW
    CommonControlClassname = "WC_TREEVIEW"
End Select

End Function


'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this WNDCLASS object from the class name identified to by
'\\ the instance and class name combination
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function CreateFromClassname(ByVal hInst As Long, ByVal lpClassName As String) As Boolean

Dim lpClass As WndClass
Dim lRet As Long

lRet = GetClassInfoApi(hInst, lpClassName, lpClass)
If Err.LastDllError = 0 And lRet > 0 Then
    CreateFromClassname = CreateFromPointer(VarPtr(lpClass))
End If

End Function

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this WNDCLASS object from the location poiunted to by
'\\ lpWndClass
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function CreateFromPointer(lpWndClass As Long) As Boolean

Dim wcThis As WndClass

CreatedOK = False

If Not IsBadReadPtrWndclass(lpWndClass, Len(wcThis)) Then
    Call CopyMemoryWndClass(wcThis, lpWndClass, Len(wcThis))
    If Err.LastDllError = 0 Then
        With wcThis
            style = .style
            lpfnWndProc = .lpfnWndProc
            cbClsExtra = .cbClsExtra
            cbWndExtra2 = .cbWndExtra2
            hInstance = .hInstance
            Set Icon = New ApiIcon
            Icon.hIcon = .hIcon
            hCursor = .hCursor
            hbrBackground = .hbrBackground
            lpszClassName = StringFromPointer(.lpszClassName, 1024)
            lpszMenuName = StringFromPointer(.lpszMenuName, 1024)
            If Err.LastDllError = 0 Then
                CreatedOK = True
            End If
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function

Public Property Set Icon(ByVal newIcon As ApiIcon)

If newIcon Is Nothing Then
    Set mIcon = Nothing
Else
    If newIcon.hIcon <> Icon.hIcon Then
        Set mIcon = newIcon
    End If
End If

End Property

Public Property Get Icon() As ApiIcon

    If mIcon Is Nothing Then
        Set mIcon = New ApiIcon
    End If
    
    Set Icon = mIcon
    
End Property

Public Property Let lpfnWndProc(ByVal newLPfnwndproc As Long)

    m_lpfnwndproc = newLPfnwndproc
    
End Property

Public Property Get lpfnWndProc() As Long

    lpfnWndProc = m_lpfnwndproc
    
End Property


Public Function RegisterClassEx() As Boolean

Dim lRet As Long

Dim wndclassThis As WNDCLASSEX

With wndclassThis
    .cbClsExtra = Me.cbClsExtra
    .cbWndExtra = Me.cbWndExtra2
    .cbSize = LenB(wndclassThis)
    .hbrBackground = Me.hbrBackground
    .hCursor = Me.hCursor
    .hInstance = Me.hInstance
    .hIcon = Me.Icon.hIcon
    '.hIconSm = Me.IconSmall.hIcon
    .lpfnWndProc = Me.lpfnWndProc
    .lpszClassName = Me.lpszClassName
    .lpszMenuName = Me.lpszMenuName
    .style = Me.style
End With

lRet = RegisterClassExApi(wndclassThis)
If Err.LastDllError = 0 Then
    RegisterClassEx = True
Else
    Call ReportError(Err.LastDllError, "ApiWndClass:RegisterClassEx", APIDispenser.LastSystemError)
End If

End Function

Public Sub SetDefaultProcAddress()

   Call SetProcAddress(AddressOf VB_WindowProc)
    
End Sub

Private Sub SetProcAddress(ByVal lpAddress As Long)

    Me.lpfnWndProc = lpAddress
    
End Sub


'\\ --[StandardClassName]-------------------------------------------
'\\ Converts between the exported StandarWindowClass type
'\\ and the actual string class name used by windows
'\\ This is done because VB doesn't allow fixed length string
'\\ constants to be exported from a class.
'\\ ----------------------------------------------------------------
Private Function StandardClassName(ByVal ClassType As enStandardWindowClasses) As String

Select Case ClassType
Case SWC_BUTTON
    StandardClassName = "BUTTON"
Case SWC_COMBOBOX
    StandardClassName = "COMBOBOX"
Case SWC_EDIT
    StandardClassName = "EDIT"
Case SWC_LISTBOX
    StandardClassName = "LISTBOX"
Case SWC_SCROLLBAR
    StandardClassName = "SCROLLBAR"
End Select

End Function

Private Sub Class_Terminate()

Set mIcon = Nothing

End Sub


