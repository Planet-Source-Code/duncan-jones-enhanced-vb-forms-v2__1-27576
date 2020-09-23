VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'\\ --[EnumHandler]-----------------------------------------------------
'\\ Provides an 'event' interface to APICallbackProcs
'\\ --------------------------------------------------------------------

Public Enum CallbackProcTypes
    DLGPROC = 1 'typedef BOOL (CALLBACK* DLGPROC)(HWND, UINT, WPARAM, LPARAM);
    TIMERPROC = 2 'typedef VOID (CALLBACK* TIMERPROC)(HWND, UINT, UINT, DWORD);
    GRAYSTRINGPROC = 3 'typedef BOOL (CALLBACK* GRAYSTRINGPROC)(HDC, LPARAM, int);
    HOOKPROC = 4 'typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
    SENDASYNCPROC = 5 'typedef VOID (CALLBACK* SENDASYNCPROC)(HWND, UINT, DWORD, LRESULT);
    PROPENUMPROC = 6 'typedef BOOL (CALLBACK* PROPENUMPROCA)(HWND, LPCSTR, HANDLE);
    PROPENUMPROCEX = 7 'typedef BOOL (CALLBACK* PROPENUMPROCEXA)(HWND, LPSTR, HANDLE, DWORD);
    EDITWORDBREAKPROC = 8 'typedef int (CALLBACK* EDITWORDBREAKPROCA)(LPSTR lpch, int ichCurrent, int cch, int code);
    WNDENUMPROC = 9 'typedef BOOL (CALLBACK* WNDENUMPROC)(HWND, LPARAM);
    WNDPROC = 10 'typedef LRESULT (CALLBACK* WNDPROC)(HWND, UINT, WPARAM, LPARAM);
    WINSTATIONENUMPROC = 11
    DESKTOPENUMPROC = 12
    ENUMRESTYPEPROC = 13
    ENUMRESNAMEPROC = 14
    ENUMRESLANGPROC = 15
    '\\ DEJ 01/03/2001 - Added hook proc subtypes
    HOOKPROC_CALLWNDPROC = 16 'CallWndProc
    HOOKPROC_CALLWNDPROCRET = 17 'CallWndProcRet
    HOOKPROC_CBT = 18 'CBTProc
    HOOKPROC_DEBUG = 19 'DebugProc
    HOOKPROC_GETMESSAGE = 20 'GetMsgProc
    HOOKPROC_JOURNALPLAYBACK = 21 'JournalPlaybackProc
    HOOKPROC_JOURNALRECORD = 22 'JournalRecordProc
    HOOKPROC_KEYBOARD = 23 'KeyboardProc
    HOOKPROC_MOUSE = 24 'MouseProc
    HOOKPROC_MESSAGEFILTER = 25 'MessageProc
    HOOKPROC_SHELL = 26 'ShellProc
    HOOKPROC_SYSMESSAGEFILTER = 27 'SysMsgProc
    '\\ DEJ 22/03/2001 - Win NT/2000
    HOOKPROC_MOUSE_LL = 28
    HOOKPROC_KEYBOARD_LL = 29
    '\\ 26/03/2001 - More new hookproc types added
    HOOKPROC_HARDWARE = 30
    HOOKPROC_FOREGROUNDIDLE = 31
End Enum

Event DLGPROCFired(hwnd As Long, uint As Long, wParam As Long, lParam As Long, bDiscardMessage As Boolean, lDlgProcRet As Long)
Event TIMERPROCFired(hwnd As Long, uint As Long, nEventId As Long, dwTime As Long)
Event GRAYSTRINGPROCfired(hdc As Long, lParam As Long, nInt As Long)
Event HOOKPROCFired(code As Long, wParam As Long, lParam As Long, lMsgRet As Long)
Event SENDASYNCPROCFired(hwnd As Long, uint As Long, dWord As Long, lResult As Long)
Event PROPENUMPROCFired(hwnd As Long, LPCSTR As String, Handle As Long)
Event PROPENUMPROCEXFired(hwnd As Long, lpStr As String, Handle As Long, dWord As Long)
Event EDITWORDBREAKPROCFired(lpch As String, ichCurrent As Long, cch As Long, code As Long)
Event WNDENUMPROCFired(hwnd As Long, lParam As Long)
Event WNDPROCFired(hwnd As Long, wMsg As Long, wParam As Long, lParam As Long, bDiscardMessage As Boolean, lWndProcRet As Long)
Event WINSTATIONENUMPROCFired(lpstrName As String, lParam As Long)
Event DESKTOPENUMPROCFired(lpstrName As String, lParam As Long)
Event ENUMRESTYPEPROCFired(hModule As Long, lpType As String, lParam As Long)
Event ENUMRESNAMEPROCFired(hModule As Long, lpType As String, lpName As String, lParam As Long)
Event ENUMRESLANGPROCFired(hModule As Long, lpType As String, lpName As String, wLanguage As Long, lParam As Long)
'\\ New events for better hook procs...
Event HOOKPROCCALLWNDPROC(Action As enHookCode, FromCurrentProcess As Boolean, wParam As Long, lMsgRet As Long)   'wParam As CWPSTRUCT
Event HOOKPROCCALLWNDPROCRET(Action As enHookCode, FromCurrentProcess As Boolean, wParam As Long, lMsgRet As Long) 'wParam As CWPRETSTRUCT
Event HOOKPROCCBT(CBTCode As enHookCBTCodes, lParam As Long, wParam As Long, lMsgRet As Long)
Event HOOKPROCDEBUG(Action As enHookCode, HookType As enHookTypes, DebugInfo As ApiDEBUGHOOKINFO, lMsgRet As Long)  'DebugInfo As DEBUGHOOKINFO
Event HOOKPROCMESSAGE(Action As enHookCode, RemoveFlag As enPeekMessage, MessageInfo As ApiMSG, lMsgRet As Long)   'Message As MSG
Event HOOKPROCJOURNALPLAYBACK(Action As enHookCode, wNull As Long, wParam As Long, lMsgRet As Long)  'EventMessage As EVENTMSG
Event HOOKPROCJOURNALRECORD(Action As enHookCode, wNull As Long, wParam As Long, lMsgRet As Long) 'EventMessage As EVENTMSG
Event HOOKPROCKEYBOARD(Action As enHookCode, VirtualKey As Long, KeyStrokeInfo As Long, lMsgRet As Long)
Event HOOKPROCMOUSE(Action As enHookCode, wParam As WindowMessages, lParam As ApiMOUSEHOOKSTRUCT, lMsgRet As Long)  'wParam as MOUSEHOOKSTRUCT
Event HOOKPROCMOUSELL(Action As enHookCode, wParam As WindowMessages, lParam As ApiLLMOUSEHOOKSTRUCT, lMsgRet As Long)  'wParam as MOUSEHOOKSTRUCT
Event HOOKPROCMESSAGEFILTER(Filter As enMessageFilter, MessageInfo As ApiMSG, lMsgRet As Long)   'Message As MSG
'Event HOOKPROCSHELL(Message As enShellMessage, hwnd As Long, lParam As Long, lMsgRet As Long)
Event HOOKPROCSYSMESSAGE(Filter As enMessageFilter, ByVal wMsg As ApiMSG, lMsgRet As Long)   'Message As MSG
Event HOOKPROCKEYBOARDLL(Action As enHookCode, VirtualKey As Long, KeyStrokeInfo As Long, lMsgRet As Long)
Event HOOKPROCFOREGROUNDIDLE(Action As enHookCode)



'\\ API Calls to pass on WM_ message to previous windows proc...
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function DefDlgProc Lib "user32" Alias "DefDlgProcA" (ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'\\ Existing HOOK id...
'\\ Note: current design is max one or each type of hook proc per EnumHandler
Private CallWndProcHookId As Long
Private CallWndProcRetHookId As Long
Private CBTHookId As Long
Private DebugHookId As Long
Private ForegroundidleHookId As Long
Private GetMessageHookid As Long
Private HardwareHookId As Long
Private MessageHookId As Long
Private JournalRecordHookId As Long
Private JournalPlaybackHookId As Long
Private MouseHookId As Long
Private MsgFilterHookId As Long
Private KeyboardHookId As Long
Private ShellHookId As Long
Private SysmessageHookId As Long
Private LowlevelMouseHookId As Long
Private LowLevelKeyboardHookId As Long

'\\ Hook codes....
Public Enum enHookCode
    HC_ACTION = 0
    HC_GETNEXT = 1
    HC_NOREMOVE = 3
    HC_SKIP = 2
    HC_SYSMODALOFF = 5
    HC_SYSMODALON = 4
End Enum

Public Enum enHookTypes
    WH_CALLWNDPROC = 4
    WH_CBT = 5
    WH_DEBUG = 9
    WH_FOREGROUNDIDLE = 11
    WH_GETMESSAGE = 3
    WH_HARDWARE = 8
    WH_JOURNALPLAYBACK = 1
    WH_JOURNALRECORD = 0
    WH_MOUSE = 7
    WH_MSGFILTER = (-1)
    WH_SHELL = 10
    WH_SYSMSGFILTER = 6
    WH_KEYBOARD_LL = 13
    WH_MOUSE_LL = 14
    WH_KEYBOARD = 2
End Enum

Public Enum enHookCBTCodes
    HCBT_ACTIVATE = 5  '\\ The system is about to activate a window....
    HCBT_CLICKSKIPPED = 6 '\\ A mouse click has been removed from the queue for processing
    HCBT_CREATEWND = 3 '\\ A window is being created
    HCBT_DESTROYWND = 4
    HCBT_KEYSKIPPED = 7 '\\ A keystroke has been removed from the queue
    HCBT_MINMAX = 1 '\\ A window is going to be minimised or maximised
    HCBT_MOVESIZE = 0 '\\ A window is being moved or resized
    HCBT_QS = 2 '\\ A WM_QUEUESYNC message has been received
    HCBT_SETFOCUS = 9 '\\ A window is about to get a setfocus
    HCBT_SYSCOMMAND = 8
End Enum

Public Enum enPeekMessage
    PM_NOREMOVE = &H0
    PM_NOYIELD = &H2
    PM_REMOVE = &H1
End Enum

Public Enum enMessageFilter
    MSGF_DDEMGR = &H8001
    MSGF_DIALOGBOX = 0
    MSGF_MAINLOOP = 8
    MSGF_MENU = 2
    MSGF_MESSAGEBOX = 1
    MSGF_MOVE = 3
    MSGF_NEXTWINDOW = 6
    MSGF_SCROLLBAR = 5
    MSGF_SIZE = 4
    MSGF_USER = 4096
End Enum

Public Enum enShellMessage
    HSHELL_ACTIVATESHELLWINDOW = 3
    HSHELL_WINDOWCREATED = 1
    HSHELL_WINDOWDESTROYED = 2
End Enum

'\\ Windows hooks...
'SetWindowsHookEx
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public CurrentHookType As enHookTypes

Public Event ApiError(ByVal Source As String, ByVal Description As String)

'\\ New events for the different CBT messages that could happen....
Public Event HOOKPROCCBTACTIVATE(ByVal wnd As ApiWindow, ByVal ActivateSytruct As ApiCBTACTIVATESTRUCT, Cancel As Boolean)
Public Event HOOKPROCCBTCREATEWND(ByVal wnd As ApiWindow, ByVal CreateWndStruct As ApiCBT_CREATEWND, Cancel As Boolean)
Public Event HOOKPROCCBTDESTROYWND(ByVal wnd As ApiWindow, Cancel As Boolean)
Public Event HOOKPROCCBTMINMAX(ByVal wnd As ApiWindow, ByVal ShowWindow As enShowWindow, Cancel As Boolean)
Public Event HOOKPROCCBTMOVESIZE(ByVal wnd As ApiWindow, ByVal NewRect As APIRect, Cancel As Boolean)
Public Event HOOKPROCCBTSETFOCUS(ByVal hwndNewFocus As ApiWindow, ByVal hwndOldFocus As ApiWindow, Cancel As Boolean)
Public Event HOOKPROCCBTSYSCOMMAND(ByVal SysCommand As enSystemCommands, ByVal lParam As Long, Cancel As Boolean)

'\\ New events for hookprocshell
Public Event HOOKPROCSHELLCREATEWINDOW(ByVal wnd As ApiWindow)
Public Event HOOKPROCSHELLDESTROYWINDOW(ByVal wnd As ApiWindow)

Public Property Let HookIdByType(ByVal HookType As enHookTypes, ByVal newHookId As Long)

Select Case HookType
Case WH_CALLWNDPROC
    If newHookId <> CallWndProcHookId Then
        CallWndProcHookId = newHookId
    End If

Case WH_CBT
    If newHookId <> CBTHookId Then
        CBTHookId = newHookId
    End If
    
Case WH_DEBUG
    If newHookId <> DebugHookId Then
        DebugHookId = newHookId
    End If
    
Case WH_FOREGROUNDIDLE
    If newHookId <> ForegroundidleHookId Then
        ForegroundidleHookId = newHookId
    End If
    
Case WH_GETMESSAGE
    If newHookId <> GetMessageHookid Then
        GetMessageHookid = newHookId
    End If
    
Case WH_HARDWARE
    If newHookId <> HardwareHookId Then
        HardwareHookId = newHookId
    End If
    
Case WH_JOURNALPLAYBACK
    If newHookId <> JournalPlaybackHookId Then
        JournalPlaybackHookId = newHookId
    End If
    
Case WH_JOURNALRECORD
    If newHookId <> JournalRecordHookId Then
        JournalRecordHookId = newHookId
    End If
    
Case WH_MOUSE
    If newHookId <> MouseHookId Then
        MouseHookId = newHookId
    End If
    
Case WH_MSGFILTER
    If newHookId <> MsgFilterHookId Then
        MsgFilterHookId = newHookId
    End If
    
Case WH_SHELL
    If newHookId <> ShellHookId Then
        ShellHookId = newHookId
    End If
    
Case WH_SYSMSGFILTER
    If newHookId <> SysmessageHookId Then
        SysmessageHookId = newHookId
    End If
    
Case WH_KEYBOARD_LL
    If newHookId <> LowLevelKeyboardHookId Then
        LowLevelKeyboardHookId = newHookId
    End If
    
Case WH_MOUSE_LL
    If newHookId <> LowlevelMouseHookId Then
        LowlevelMouseHookId = newHookId
    End If
    
End Select

End Property

Public Property Get HookIdByType(ByVal HookType As enHookTypes) As Long

Select Case HookType
Case WH_CALLWNDPROC
    HookIdByType = CallWndProcHookId
Case WH_CBT
    HookIdByType = CBTHookId
Case WH_DEBUG
    HookIdByType = DebugHookId
Case WH_FOREGROUNDIDLE
    HookIdByType = ForegroundidleHookId
Case WH_GETMESSAGE
    HookIdByType = GetMessageHookid
Case WH_HARDWARE
    HookIdByType = HardwareHookId
Case WH_JOURNALPLAYBACK
    HookIdByType = JournalPlaybackHookId
Case WH_JOURNALRECORD
    HookIdByType = JournalRecordHookId
Case WH_MOUSE
    HookIdByType = MouseHookId
Case WH_MSGFILTER
    HookIdByType = MsgFilterHookId
Case WH_SHELL
    HookIdByType = ShellHookId
Case WH_SYSMSGFILTER
    HookIdByType = SysmessageHookId
Case WH_KEYBOARD_LL
    HookIdByType = LowLevelKeyboardHookId
Case WH_MOUSE_LL
    HookIdByType = LowlevelMouseHookId
End Select

End Property

Public Sub StartHook(ByVal HookType As enHookTypes, ByVal ModuleHandle As Long, ByVal ThreadId As Long)

Dim lRet As Long
Dim hMod As Long

'\\ If a hook of this type is already set, unhook this first
If HookIdByType(HookType) > 0 Then
    Call UnhookWindowsHookEx(HookIdByType(HookType))
End If

'\\ Start the appropriate hook procedure
Select Case HookType
Case WH_CALLWNDPROC
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKCALLWNDPROC, ModuleHandle, ThreadId)
Case WH_CBT
    If ModuleHandle > 0 Then
        lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKCBTPROC, ModuleHandle, ThreadId)
    Else
        lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKCBTPROC, vbNull, ThreadId)
    End If
Case WH_DEBUG
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKDEBUGPROC, ModuleHandle, ThreadId)
Case WH_FOREGROUNDIDLE
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKFOREGROUNDIDLEPROC, ModuleHandle, ThreadId)
Case WH_GETMESSAGE
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKGETMESSAGEPROC, ModuleHandle, ThreadId)
Case WH_HARDWARE
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKHARDWAREPROC, ModuleHandle, ThreadId)
Case WH_JOURNALPLAYBACK
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKJOURNALPLAYBACKPROC, ModuleHandle, ThreadId)
Case WH_JOURNALRECORD
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKJOURNALRECORDPROC, ModuleHandle, ThreadId)
Case WH_MOUSE
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKMOUSEPROC, ModuleHandle, ThreadId)
Case WH_MSGFILTER
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKMESSAGEFILTERPROC, ModuleHandle, ThreadId)
Case WH_SHELL
    If ModuleHandle > 0 Then
        lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKSHELLPROC, ModuleHandle, ThreadId)
    Else
        lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKSHELLPROC, vbNull, ThreadId)
    End If
Case WH_SYSMSGFILTER
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKSYSMESSAGEFILTERPROC, ModuleHandle, ThreadId)
Case WH_KEYBOARD_LL
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKLOWLEVELKEYBOARDPROC, ModuleHandle, ThreadId)
Case WH_MOUSE_LL
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKLOWLEVELMOUSEPROC, ModuleHandle, ThreadId)
Case WH_KEYBOARD
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKKEYBOARDPROC, ModuleHandle, ThreadId)
Case Else
    '\\ Non specific hook type
    lRet = SetWindowsHookEx(HookType, AddressOf VB_HOOKPROC, ModuleHandle, ThreadId)
End Select


If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "EnumHandler:StartHook", APIDispenser.LastSystemError)
End If

If lRet > 0 Then
    HookIdByType(HookType) = lRet
    CurrentHookType = HookType
End If

End Sub

Public Sub StopHook(ByVal HookType As enHookTypes)

Dim lRet As Long

'\\ If a hook of this type is already set, unhook this first
If HookIdByType(HookType) > 0 Then
    lRet = UnhookWindowsHookEx(HookIdByType(HookType))
    If Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, "EnumHandler:StopHook", APIDispenser.LastSystemError)
    End If
End If

End Sub


Public Sub TriggerEvent(ByVal ProcType As CallbackProcTypes, Arguments() As Variant)

Dim hwnd As Long, wMsg As Long, wParam As Long, lParam As Long
Dim nEventId As Long, dWord As Long, dwTime As Long
Dim hdc As Long, nInt As Long
Dim code As Long
Dim lResult As Long
Dim LPCSTR As String, Handle As Long
Dim lpch As String, ichCurrent As Long, cch As Long
Dim lRet As Long, lWndProcRet As Long
Dim bDiscardMsg As Boolean
Dim ftTime As APIFileTime
Dim stTime As APISystemTime
Dim wndThis As ApiWindow
Dim dtTime As Date
Dim lpstrName As String
Dim hModule As Long, lpType As String, lpName As String, wLanguage As Long
Dim bCurrentProc As Boolean
Dim MsgInfo As ApiMSG

On Error Resume Next

Select Case ProcType
Case DLGPROC

    'hwnd = Arguments(1)
    'wMsg = Arguments(2)
    'wParam = Arguments(3)
    'lParam = Arguments(4)
    'RaiseEvent DLGPROCFired(hwnd, wMsg, wParam, lParam, bDiscardMsg, lWndProcRet)
    'If Not bDiscardMsg Then
    '    If m_OldDlgProc = 0 Then
    '        lRet = DefDlgProc(hwnd, wMsg, wParam, lParam)
    '    Else
    '        lRet = CallWindowProc(m_OldDlgProc, hwnd, wMsg, wParam, lParam)
    '    End If
    'End If
    'If lWndProcRet = 0 Then
    '    Arguments(5) = lRet
    'Else
    '    Arguments(5) = lWndProcRet
    'End If

Case EDITWORDBREAKPROC

    lpch = Arguments(1)
    ichCurrent = Arguments(2)
    cch = Arguments(3)
    code = Arguments(4)
    RaiseEvent EDITWORDBREAKPROCFired(lpch, ichCurrent, cch, code)

Case GRAYSTRINGPROC

    RaiseEvent GRAYSTRINGPROCfired(hdc, lParam, nInt)
    
Case HOOKPROC

    code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    RaiseEvent HOOKPROCFired(code, wParam, lParam, lResult)
    '\\ Pass lResult back...
    Arguments(4) = lResult

Case PROPENUMPROC

    hwnd = Arguments(1)
    LPCSTR = Arguments(2)
    Handle = Arguments(3)
    RaiseEvent PROPENUMPROCFired(hwnd, LPCSTR, Handle)
    
Case PROPENUMPROCEX

    hwnd = Arguments(1)
    LPCSTR = Arguments(2)
    Handle = Arguments(3)
    dWord = Arguments(4)
    RaiseEvent PROPENUMPROCEXFired(hwnd, LPCSTR, Handle, dWord)
    
Case SENDASYNCPROC

    RaiseEvent SENDASYNCPROCFired(hwnd, wMsg, dWord, lResult)
    
Case TIMERPROC
    
    hwnd = Arguments(1)
    wMsg = Arguments(2)
    nEventId = Arguments(3)
    dwTime = Arguments(4)
    RaiseEvent TIMERPROCFired(hwnd, wMsg, nEventId, dwTime)
    
Case WNDENUMPROC

    hwnd = Arguments(1)
    lParam = Arguments(2)
    RaiseEvent WNDENUMPROCFired(hwnd, lParam)
    
Case WNDPROC

    hwnd = Arguments(1)
    wMsg = Arguments(2)
    wParam = Arguments(3)
    lParam = Arguments(4)
    
    If Not AllSubclassedWindows Is Nothing Then
        Call AllSubclassedWindows.Item(hwnd).TriggerEvent(wMsg, wParam, lParam, bDiscardMsg, lWndProcRet)
    End If
    
    If Not bDiscardMsg Then
        If AllSubclassedWindows.Item(hwnd).OldProcAddress > 0 Then
            lRet = CallWindowProc(AllSubclassedWindows.Item(hwnd).OldProcAddress, hwnd, wMsg, wParam, lParam)
        Else
            lRet = DefWindowProc(hwnd, wMsg, wParam, lParam)
        End If
    End If
    If lWndProcRet = 0 Then
        Arguments(5) = lRet
    Else
        Arguments(5) = lWndProcRet
    End If
    
Case WINSTATIONENUMPROC
    lpstrName = Arguments(1)
    lParam = Arguments(2)
    RaiseEvent WINSTATIONENUMPROCFired(lpstrName, lParam)

Case DESKTOPENUMPROC
    lpstrName = Arguments(1)
    lParam = Arguments(2)
    RaiseEvent DESKTOPENUMPROCFired(lpstrName, lParam)

Case ENUMRESTYPEPROC
    hModule = Arguments(1)
    lpType = Arguments(2)
    lParam = Arguments(3)
    RaiseEvent ENUMRESTYPEPROCFired(hModule, lpType, lParam)

Case ENUMRESNAMEPROC
    hModule = Arguments(1)
    lpType = Arguments(2)
    lpName = Arguments(3)
    lParam = Arguments(4)
    RaiseEvent ENUMRESNAMEPROCFired(hModule, lpType, lpName, lParam)
    
Case ENUMRESLANGPROC
    hModule = Arguments(1)
    lpType = Arguments(2)
    lpName = Arguments(3)
    wLanguage = Arguments(4)
    lParam = Arguments(5)
    RaiseEvent ENUMRESLANGPROCFired(hModule, lpType, lpName, wLanguage, lParam)

'\\ DEJ 01/03/2001 - New specific hook procedures
Case HOOKPROC_CALLWNDPROC
    code = Arguments(1)
    bCurrentProc = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    RaiseEvent HOOKPROCCALLWNDPROC(code, bCurrentProc, lParam, lResult)
    '\\ Pass lResult back...
    Arguments(4) = lResult
    
Case HOOKPROC_CALLWNDPROCRET
    code = Arguments(1)
    bCurrentProc = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    RaiseEvent HOOKPROCCALLWNDPROCRET(code, bCurrentProc, lParam, lResult)
    '\\ Pass lResult back...
    Arguments(4) = lResult

Case HOOKPROC_CBT
    code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    '\\ Create a new window for these to use...
    Set wndThis = New ApiWindow
    
    '\\ Different events raised according to which message this is...
    Select Case True
    Case code = HCBT_ACTIVATE
        Dim ActivateStruct As ApiCBTACTIVATESTRUCT
        Set ActivateStruct = New ApiCBTACTIVATESTRUCT
        ActivateStruct.CreateFromPointer lParam
        wndThis.hwnd = wParam
        RaiseEvent HOOKPROCCBTACTIVATE(wndThis, ActivateStruct, bDiscardMsg)
        Set ActivateStruct = Nothing
    
    Case code = HCBT_CREATEWND
        Dim CreateWndStruct As ApiCBT_CREATEWND
        Set CreateWndStruct = New ApiCBT_CREATEWND
        CreateWndStruct.CreateFromPointer lParam
        wndThis.hwnd = wParam
        RaiseEvent HOOKPROCCBTCREATEWND(wndThis, CreateWndStruct, bDiscardMsg)
        Set CreateWndStruct = Nothing
    
    Case code = HCBT_DESTROYWND
        wndThis.hwnd = wParam
        RaiseEvent HOOKPROCCBTDESTROYWND(wndThis, bDiscardMsg)
        
    Case code = HCBT_MINMAX
        wndThis.hwnd = wParam
        RaiseEvent HOOKPROCCBTMINMAX(wndThis, APIDispenser.LoWord(lParam), bDiscardMsg)
        
    Case code = HCBT_MOVESIZE
        Dim SizeRect As APIRect
        Set SizeRect = New APIRect
        SizeRect.CreateFromPointer lParam
        wndThis.hwnd = wParam
        RaiseEvent HOOKPROCCBTMOVESIZE(wndThis, SizeRect, bDiscardMsg)
        '\\ The sizerect can be changed :. save it back..
        SizeRect.SaveToPointer lParam
        Set SizeRect = Nothing
        
    Case code = HCBT_SETFOCUS
        Dim wndLostFocus As ApiWindow
        Set wndLostFocus = New ApiWindow
        wndThis.hwnd = wParam
        wndLostFocus.hwnd = lParam
        RaiseEvent HOOKPROCCBTSETFOCUS(wndThis, wndLostFocus, bDiscardMsg)
        
    Case code = HCBT_SYSCOMMAND
        RaiseEvent HOOKPROCCBTSYSCOMMAND(wParam, lParam, bDiscardMsg)
        
    Case Else
        RaiseEvent HOOKPROCCBT(code, wParam, lParam, lResult)
    End Select
    '\\ Pass lResult back...
    Arguments(4) = lResult + CLng(bDiscardMsg)

Case HOOKPROC_DEBUG
    code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    Dim DebugInfo As ApiDEBUGHOOKINFO
    Set DebugInfo = New ApiDEBUGHOOKINFO
    DebugInfo.CreateFromPointer lParam
    RaiseEvent HOOKPROCDEBUG(code, wParam, DebugInfo, lResult)
    '\\ Pass lResult back...
    Arguments(4) = lResult
    
Case HOOKPROC_GETMESSAGE
    code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    Set MsgInfo = New ApiMSG
    MsgInfo.CreateFromPointer lParam
    RaiseEvent HOOKPROCMESSAGE(code, wParam, MsgInfo, lResult)
    '\\ Pass lResult back...
    Arguments(4) = lResult

Case HOOKPROC_JOURNALPLAYBACK
    code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    RaiseEvent HOOKPROCJOURNALPLAYBACK(code, wParam, lParam, lResult)
    '\\ Pass lResult back...
    Arguments(4) = lResult

Case HOOKPROC_JOURNALRECORD
    code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    RaiseEvent HOOKPROCJOURNALRECORD(code, wParam, lParam, lResult)
    '\\ Pass lResult back...
    Arguments(4) = lResult

Case HOOKPROC_MOUSE
    code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    '\\ Conver the pointer into a mouse hook structure object
    Dim MouseHookData As ApiMOUSEHOOKSTRUCT
    Set MouseHookData = New ApiMOUSEHOOKSTRUCT
    MouseHookData.CreateFromPointer lParam
    
    RaiseEvent HOOKPROCMOUSE(code, wParam, MouseHookData, lResult)
    '\\ Pass lResult back...
    Arguments(4) = lResult

Case HOOKPROC_MOUSE_LL
    code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    '\\ Conver the pointer into a mouse hook structure object
    Dim LLMouseHookData As ApiLLMOUSEHOOKSTRUCT
    Set LLMouseHookData = New ApiLLMOUSEHOOKSTRUCT
    LLMouseHookData.CreateFromPointer lParam
    
    RaiseEvent HOOKPROCMOUSELL(code, wParam, LLMouseHookData, lResult)
    '\\ Pass lResult back...
    Arguments(4) = lResult

Case HOOKPROC_MESSAGEFILTER
    code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    Set MsgInfo = New ApiMSG
    MsgInfo.CreateFromPointer lParam
    RaiseEvent HOOKPROCMESSAGEFILTER(code, MsgInfo, lResult)
    Set MsgInfo = Nothing
    '\\ Pass lResult back...
    Arguments(4) = lResult

Case HOOKPROC_SHELL
    code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    Set wndThis = New ApiWindow
    wndThis.hwnd = wParam
    If code = HSHELL_WINDOWCREATED Then
        RaiseEvent HOOKPROCSHELLCREATEWINDOW(wndThis)
    ElseIf code = HSHELL_WINDOWDESTROYED Then
        RaiseEvent HOOKPROCSHELLDESTROYWINDOW(wndThis)
    End If
    Set wndThis = Nothing
    '\\ Pass lResult back...
    Arguments(4) = lResult

Case HOOKPROC_SYSMESSAGEFILTER
    code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    Set MsgInfo = New ApiMSG
    MsgInfo.CreateFromPointer lParam
    RaiseEvent HOOKPROCSYSMESSAGE(code, MsgInfo, lResult)
    Set MsgInfo = Nothing
    '\\ Pass lResult back...
    Arguments(4) = lResult

Case HOOKPROC_KEYBOARD
    code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    RaiseEvent HOOKPROCKEYBOARD(code, wParam, lParam, lResult)
    '\\ Pass lResult back...
    Arguments(4) = lResult

Case HOOKPROC_KEYBOARD_LL
    code = Arguments(1)
    wParam = Arguments(2)
    lParam = Arguments(3)
    lResult = Arguments(4)
    RaiseEvent HOOKPROCKEYBOARDLL(code, wParam, lParam, lResult)
    '\\ Pass lResult back...
    Arguments(4) = lResult
    
Case HOOKPROC_FOREGROUNDIDLE
    code = Arguments(1)
    '\\ Other argumenst aren't used
    RaiseEvent HOOKPROCFOREGROUNDIDLE(code)
    Arguments(4) = code
    
End Select


End Sub

Private Sub Class_Terminate()

Dim lRet As Long

'\\ If ANY hook id is set, it must be closed
Call StopHook(WH_CALLWNDPROC)
Call StopHook(WH_CBT)
Call StopHook(WH_DEBUG)
Call StopHook(WH_FOREGROUNDIDLE)
Call StopHook(WH_GETMESSAGE)
Call StopHook(WH_HARDWARE)
Call StopHook(WH_JOURNALPLAYBACK)
Call StopHook(WH_JOURNALRECORD)
Call StopHook(WH_KEYBOARD_LL)
Call StopHook(WH_MOUSE)
Call StopHook(WH_MOUSE_LL)
Call StopHook(WH_MSGFILTER)
Call StopHook(WH_SHELL)
Call StopHook(WH_SYSMSGFILTER)
Call StopHook(WH_KEYBOARD)


End Sub


