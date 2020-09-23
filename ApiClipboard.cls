VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'\\ --[ApiClipboard]-----------------------------------------------------------
'\\ Extends the Visual basic clipboard object by use of the Api
'\\ ---------------------------------------------------------------------------

Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long

Private Declare Function GetClipboardOwner Lib "user32" () As Long

Private Declare Function GetClipboardViewer Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function SetClipboardViewer Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long

Private Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long

Public Enum enClipboardFormats
    CF_BITMAP = 2
    CF_DIB = 8
    CF_DIF = 5
    CF_ENHMETAFILE = 14
    CF_METAFILEPICT = 3
    CF_OEMTEXT = 7
    CF_PALETTE = 9
    CF_PENDATA = 10
    CF_RIFF = 11
    CF_SYLK = 4
    CF_TEXT = 1
    CF_TIFF = 6
    CF_UNICODETEXT = 13
    CF_WAVE = 12
End Enum

Public ParenthWnd As Long

Private myMemory As ApiGlobalmemory
Private mLastFormat As Long

Public Property Get BackedUp() As Boolean

    BackedUp = Not (myMemory Is Nothing)
    
End Property

'\\ --[Backup]------------------------------------------------------
'\\ Makes an in-memory copy of the clipboard's contents so that they
'\\ can be restored easily
'\\ ----------------------------------------------------------------
Public Sub Backup()

Dim lRet As Long
Dim AllFormats As Collection
Dim lFormat As Long

'\\ Need to get all the formats first...
Set AllFormats = Me.ClipboardFormats

lRet = OpenClipboard(ParenthWnd)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiClipboard:Backup", APIDispenser.LastSystemError)
End If
If lRet Then
    If AllFormats.Count > 0 Then
        '\\ Get the first format that holds any data
        For lFormat = 0 To AllFormats.Count - 1
            lRet = GetClipboardData(lFormat)
            If lRet > 0 Then
                Set myMemory = New ApiGlobalmemory
                Call myMemory.CopyFromHandle(lRet)
                '\\ Keep a note of this format
                mLastFormat = lFormat
                Exit For
            End If
            'clipboard
        Next lFormat
    End If
    lRet = CloseClipboard()
End If


End Sub

Public Property Get ClipboardFormats() As Collection

Dim lRet As Long
Dim colFormats As Collection

lRet = OpenClipboard(ParenthWnd)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiClipboard:Backup", APIDispenser.LastSystemError)
End If
If lRet > 0 Then
    Set colFormats = New Collection
    '\\ Get the first available format
    lRet = EnumClipboardFormats(0)
    If Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, "ApiClipboard:Backup", APIDispenser.LastSystemError)
    End If
    While lRet > 0
        colFormats.Add lRet
        '\\ Get the next available format
        lRet = EnumClipboardFormats(lRet)
        If Err.LastDllError > 0 Then
            Call ReportError(Err.LastDllError, "ApiClipboard:Backup", APIDispenser.LastSystemError)
        End If
    Wend
    '\\ Close the clipboard object to make it available to other apps.
    lRet = CloseClipboard()
End If

Set ClipboardFormats = colFormats

End Property

'\\ --[Restore]-----------------------------------------------------
'\\ Takes the in-memory copy of the clipboard object and restores it
'\\ to the clipboard.
'\\ ----------------------------------------------------------------
Public Sub Restore()

Dim lRet As Long

If Me.BackedUp Then
    lRet = OpenClipboard(ParenthWnd)
    If Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, "ApiClipboard:Restore", APIDispenser.LastSystemError)
    End If
    If lRet Then
        myMemory.AllocationType = GMEM_FIXED
        lRet = SetClipboardData(mLastFormat, myMemory.Handle)
        myMemory.Free
        If Err.LastDllError > 0 Then
            Call ReportError(Err.LastDllError, "ApiClipboard:Backup", APIDispenser.LastSystemError)
        End If
        lRet = CloseClipboard()
        If Err.LastDllError > 0 Then
            Call ReportError(Err.LastDllError, "ApiClipboard:Backup", APIDispenser.LastSystemError)
        End If
    End If
End If

End Sub

Public Property Get Text() As String

Dim sRet As String
If Clipboard.GetFormat(vbCFText) Then
    sRet = Clipboard.GetText()
End If

End Property

Private Sub Class_Terminate()

Set myMemory = Nothing

End Sub


