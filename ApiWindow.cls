VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_hwnd As Long

'\\ API Declarations -------------------------------------------------------------------------------------------------------------------------------------------------------
Private Declare Function IsWindowApi Lib "user32" Alias "IsWindow" (ByVal hwnd As Long) As Long
Private Declare Function GetParentApi Lib "user32" Alias "GetParent" (ByVal hwnd As Long) As Long
Private Declare Function WindowFromPointApi Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowApi Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function IsWindowEnabledApi Lib "user32" Alias "IsWindowEnabled" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowUnicodeApi Lib "user32" Alias "IsWindowUnicode" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisibleApi Lib "user32" Alias "IsWindowVisible" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomedApi Lib "user32" Alias "IsZoomed" (ByVal hwnd As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
'\\ Window specific information
Private Declare Function GetWindowLongApi Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongApi Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'\\ Menu specific....
Private Declare Function GetSystemMenuApi Lib "user32" Alias "GetSystemMenu" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuApi Lib "user32" Alias "GetMenu" (ByVal hwnd As Long) As Long
'\\ Device context specific....
Private Declare Function GetDCApi Lib "user32" Alias "GetDC" (ByVal hwnd As Long) As Long
'\\ Window class specific...
Private Declare Function GetClassNameApi Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'\\ Send message....
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function InvalidateRectByPointerApi Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
'\\ Scrolling....
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetWindowRectApi Lib "user32" Alias "GetWindowRect" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ScrollWindowExAPI Lib "user32" Alias "ScrollWindowEx" (ByVal hwnd As Long, ByVal dX As Long, ByVal dY As Long, ByVal lprcScroll As Long, ByVal lprcClip As Long, ByVal hrgnUpdate As Long, lprcUpdate As RECT, ByVal fuScroll As Long) As Long
'\\ Mouse message capture
Private Declare Function SetCaptureApi Lib "user32" Alias "SetCapture" (ByVal hwnd As Long) As Long
Private Declare Function GetCaptureApi Lib "user32" Alias "GetCapture" () As Long
'\\ Windows regions.....
Private Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal HRgn As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal HRgn As Long, ByVal bRedraw As Boolean) As Long
'\\ Refreshing the window states....
Private Declare Function DrawMenuBarApi Lib "user32" Alias "DrawMenuBar" (ByVal hwnd As Long) As Long
'\\ Window text...
Private Declare Function GetWindowTextLengthApi Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextApi Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'\\ Selected text....


'\\ New window events --------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Event WindowMessageFired(ByVal msg As WindowMessages, ByVal wParam As Long, ByVal lParam As Long, Cancel As Boolean, ProcRet As Long)
Public Event ActiveApplicationChanged(ByVal ActivatingThisApp As Boolean, ByVal hThread As Long, Cancel As Boolean)
Public Event LostCapture(ByVal hwndNewCapture As Long, Cancel As Boolean)
Public Event KeyPressed(ByVal VKey As Long, ByVal Repetition As Long, ByVal ScanCode As Long, ByVal ExtendedKey As Boolean, ByVal AltDown As Boolean, ByVal AlreadyDown As Boolean, ByVal BeingPressed As Boolean, Cancel As Boolean)
Public Event LowMemory(ByVal TimeSpentCompacting As Long)
Public Event Move(ByVal x As Long, ByVal y As Long, Cancel As Boolean)

Public Event VerticalScroll(ByVal Message As enScrollMessages, ByVal Position As Long, Cancel As Boolean)
Public Event HorizontalScroll(ByVal Message As enScrollMessages, ByVal Position As Long, Cancel As Boolean)

Public Event WindowsSettingsChanged(ByVal NewSetting As enSystemParametersInfo)
Public Event WindowsINIChanged(ByVal Section As String)

Public Event NonClientMouseMove(ByVal Location As enHitTestResult, ByVal x As Single, ByVal y As Single)
Public Event NonClientMouseDown(ByVal Location As enHitTestResult, ByVal Button As Integer, ByVal x As Single, ByVal y As Single)
Public Event NonClientMouseUp(ByVal Location As enHitTestResult, ByVal Button As Integer, ByVal x As Single, ByVal y As Single)
Public Event NonClientDblClick(ByVal Location As enHitTestResult, ByVal Button As Integer, ByVal x As Single, ByVal y As Single)

Public Event MinMaxSize(MaxHeight As Long, MaxWidth As Long, MaxPositionTop As Long, MaxPositionLeft As Long, MinTrackWidth As Long, MinTrackHeight As Long, MaxTrackWidth As Long, MaxTrackHeight As Long)

Public Event MouseOverMenu(ByVal Caption As String)


'\\ Private member variables ----------------------------------------------------------------------------------------------------------------------------------------------------------
Private mOldProcAddress As Long
Private mPrevCapture As Long
'\\ Objects....
Private mSystemMenu As ApiMenu
Private mMenu As ApiMenu
Private mWndClass As APIWndClass
Private mRegion As ApiRegion
Private mDeviceContext As ApiDeviceContext

'\\ Setting of minmaxinfo....
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type


'\\ Memory manipulation routines
Private Declare Sub CopyMemoryMinMaxInfo Lib "kernel32" Alias "RtlMoveMemory" (Destination As MINMAXINFO, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMinMaxToMemoryInfo Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As MINMAXINFO, ByVal Length As Long)

'\\ message passing
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

'\\ Threads...
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

'
Public Property Get ClassName() As String

Dim sClassname As String
Dim lRet As Long

'\\ Get the class name
sClassname = String$(1024, 0)
lRet = GetClassNameApi(m_hwnd, sClassname, 1024)
If (Err.LastDllError = 0) And (lRet > 0) Then
    ClassName = Left$(sClassname, lRet)
Else
    ReportError Err.LastDllError, "ApiWindow:Classname", APIDispenser.LastSystemError
End If

End Property

Public Property Get AttachedClipboard() As ApiClipboard

Dim mClipboard As ApiClipboard

Set mClipboard = New ApiClipboard
mClipboard.ParenthWnd = Me.hwnd
Set AttachedClipboard = mClipboard

End Property

'\\ --[CopyText]----------------------------------------------------
'\\ Copies the text from a window (using the clipboard)
'\\ and returns the result as a string.
'\\ ----------------------------------------------------------------
Public Function CopyText() As String

Dim lRet As Long
Dim sText As String

'\\ First copy the text into the clipboard buffer...
lRet = SendMessage(WM_COPY, 0, 0)

'\\ Now get that data out of the clipboard data area...
If Clipboard.GetFormat(vbCFText) Then
    sText = Clipboard.GetText()
End If

CopyText = sText

End Function

Public Property Get DeviceContext() As ApiDeviceContext

Dim lRet As Long

If mDeviceContext Is Nothing Then
    Set mDeviceContext = New ApiDeviceContext
End If
lRet = GetDCApi(m_hwnd)
If Err.LastDllError = 0 And lRet <> 0 Then
    mDeviceContext.hdc = lRet
End If
Set DeviceContext = mDeviceContext

End Property

Public Sub DrawMenuBar()

Dim lRet As Long

lRet = DrawMenuBarApi(m_hwnd)
If (lRet = 0) Or (Err.LastDllError > 0) Then
    Call ReportError(Err.LastDllError, "ApiWindow:DrawMenuBar:", APIDispenser.LastSystemError)
End If

End Sub



Public Function GetParent() As ApiWindow

Dim hwndParent As Long
Dim ParentWindow As ApiWindow

On Error Resume Next

If IsWindowApi(m_hwnd) Then
    hwndParent = GetParentApi(m_hwnd)
    If IsWindowApi(hwndParent) Then
        '\\ If its already subclassed, return that instance...
        Set ParentWindow = AllSubclassedWindows.Item(hwndParent)
        If ParentWindow Is Nothing Then
            '\\ Otherwise create a new instance
            Set ParentWindow = New ApiWindow
            ParentWindow.hwnd = hwndParent
        End If
        Set GetParent = ParentWindow
    End If
End If

End Function


Public Function GetWindow(wCmd As enGetWindow) As ApiWindow

Dim hwnd As Long
Dim window As ApiWindow

On Error Resume Next

hwnd = GetWindowApi(m_hwnd, wCmd)
If IsWindowApi(hwnd) Then
    '\\ If its already subclassed, return that instance...
    Set window = AllSubclassedWindows.Item(hwnd)
    If window Is Nothing Then
        Set window = New ApiWindow
        window.hwnd = hwnd
    End If
    Set GetWindow = window
End If

End Function

Public Function GetWindowLong(Index As enGetWindowLong) As Long

Dim lRet As Long

lRet = GetWindowLongApi(m_hwnd, Index)
If Err.LastDllError = 0 Then
    GetWindowLong = lRet
End If

End Function

Public Property Get hdc() As Long

hdc = GetDCApi(m_hwnd)

End Property

Public Property Get hwnd() As Long

    hwnd = m_hwnd
    
End Property

Public Property Let hwnd(ByVal newHwnd As Long)

Dim lRet As Long

    If newHwnd <> m_hwnd Then
        '\\ Must reset the proc address before closing
        If mOldProcAddress > 0 Then
            lRet = SetWindowLong(GWL_WNDPROC, mOldProcAddress)
            mOldProcAddress = 0
        End If
        '\\ Must unset the menus...
        If Not mSystemMenu Is Nothing Then
            Set mSystemMenu = Nothing
        End If
        
        If Not mMenu Is Nothing Then
            Set mMenu = Nothing
        End If
        
        m_hwnd = newHwnd
    End If
    
End Property

Public Sub InvalidateRect(ByVal RectIn As APIRect)

Dim rcThis As RECT
Dim lRet As Long

With rcThis
    .Bottom = RectIn.Bottom
    .Left = RectIn.Left
    .Top = RectIn.Top
    .Right = RectIn.Right
End With

lRet = InvalidateRectByPointerApi(m_hwnd, VarPtr(rcThis), True)

End Sub

'\\ --[IsWindowStyleSet]----------------------------------------------------------------------
'\\ Checks for the style bit specified in the window specified.
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function IsWindowStyleSet(ByVal TheStyle As Long) As Boolean

Dim lStyle As Long
Dim lWSIndex As Long

If TheStyle = WS_EX_ACCEPTFILES Or TheStyle = WS_EX_DLGMODALFRAME Or TheStyle = WS_EX_NOPARENTNOTIFY Or TheStyle = WS_EX_TOPMOST Or TheStyle = WS_EX_TRANSPARENT Then
    lWSIndex = GWL_EXSTYLE
Else
    lWSIndex = GWL_STYLE
End If

'\\ Get the current setting of that style bit
lStyle = GetWindowLongApi(m_hwnd, lWSIndex)
IsWindowStyleSet = (lStyle And TheStyle)

End Function





Public Property Get Menu() As ApiMenu

Dim lRet As Long

lRet = GetMenuApi(m_hwnd)
If Err.LastDllError = 0 And lRet > 0 Then
    If mMenu Is Nothing Then
        Set mMenu = New ApiMenu
    End If
    mMenu.hMenu = lRet
    Set Menu = mMenu
End If

End Property





Public Property Let MouseCapture(ByVal CaptureMouse As Boolean)

Dim lRet As Long

If CaptureMouse Then
    mPrevCapture = SetCaptureApi(m_hwnd)
Else
    lRet = SetCaptureApi(mPrevCapture)
End If

End Property

Public Property Get MouseCapture() As Boolean

MouseCapture = (GetCaptureApi() = m_hwnd)

End Property

Friend Property Get OldProcAddress() As Long

    OldProcAddress = mOldProcAddress
    
End Property

Public Property Get RECT() As APIRect

Dim mRect As APIRect
Dim lpRect As RECT
Dim lRet As Long

Set mRect = New APIRect
lRet = GetWindowRectApi(m_hwnd, lpRect)
If Err.LastDllError = 0 Then
    Call mRect.CreateFromPointer(VarPtr(lpRect))
End If
Set RECT = mRect

End Property

Public Sub Refresh()

Dim lRet As Long

lRet = InvalidateRectByPointerApi(m_hwnd, 0, True)
lRet = UpdateWindow(m_hwnd)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiWindow:Refresh", APIDispenser.LastSystemError
End If


End Sub


Public Property Set Region(ByVal newRegion As ApiRegion)

Dim lRet As Long

If newRegion Is Nothing Then
    Set mRegion = Nothing
Else
    If newRegion.HRgn <> Region.HRgn Then
        Set mRegion = newRegion
        lRet = SetWindowRgn(m_hwnd, mRegion.HRgn, True)
    End If
End If

End Property

Public Property Get Region() As ApiRegion

Dim lRet As Long

If mRegion Is Nothing Then
    Set mRegion = New ApiRegion
End If
lRet = GetWindowRgn(m_hwnd, mRegion.HRgn)
If Err.LastDllError = 0 And lRet > 0 Then
    mRegion.HRgn = lRet
End If
Set Region = mRegion

End Property

Public Sub ScrollWindow(ByVal dX As Long, ByVal dY As Long, ScrollMethod As enScrollWindow)

Dim lRet As Long
Dim lpUpdate As RECT

lRet = ScrollWindowExAPI(m_hwnd, dX, dY, vbNull, vbNull, vbNull, lpUpdate, ScrollMethod)
If Err.LastDllError = 0 Then

Else
    Call ReportError(Err.LastDllError, "ApiWindow:ScrollWindow", APIDispenser.LastSystemError)
End If

End Sub

'\\ --[SelectedText]----------------------------------------------------------------------------
'\\ Returns the text that is in the currently selected part of this window
'\\ (Allows the imlementation of .SelText on controls that don't have that member)
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Property Get SelectedText() As String

Dim sSelectedText As String
Dim lSelStart As Integer
Dim lSelEnd As Integer
Dim lRet As Long

Dim lSelMsg As Long

'\\ Select the appropriate message to query the selection according to the window type
If Me.WindowBaseClass = "EDIT" Then
    lSelMsg = EM_GETSEL
ElseIf Me.WindowBaseClass = "COMBOBOX" Then
    lSelMsg = CB_GETEDITSEL
Else
    '\\ If a control supports selected items, it should use the same message
    lSelMsg = EM_GETSEL
End If

lRet = SendMessageByLong(Me.hwnd, lSelMsg, 0, 0)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiWindow:SelectedText", APIDispenser.LastSystemError
Else
    lSelEnd = APIDispenser.HiWord(lRet)
    lSelStart = APIDispenser.LoWord(lRet)
    sSelectedText = Me.WindowText
    If lSelEnd > 0 Then
        sSelectedText = Left$(sSelectedText, lSelEnd)
    End If
    If lSelStart > 0 And lSelStart < lSelEnd Then
        sSelectedText = Mid$(sSelectedText, lSelStart)
    End If
End If

SelectedText = sSelectedText

End Property

Public Function SendMessage(ByVal wMsg As WindowMessages, ByVal wParam As Long, lParam As Long) As Long

Dim lRet As Long
Dim lThis As Long
Dim sThis As String

lRet = SendMessageLong(m_hwnd, wMsg, wParam, lParam)

If Err.LastDllError = 0 Then
    SendMessage = lRet
End If

End Function

'\\ --[SetWindowStyle]----------------------------------------------------------------------
'\\ Sets the style bit specified to the window specified.  Note that many window style
'\\ bits cannot be used at run time :. use this with caution
'\\ Returns true if it succeeded
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function SetWindowStyle(ByVal NewStyle As enWindowStyles) As Boolean

Dim lStyle As Long
Dim lWSIndex As Long
Dim lRet As Long

If NewStyle = WS_EX_ACCEPTFILES Or NewStyle = WS_EX_DLGMODALFRAME Or NewStyle = WS_EX_NOPARENTNOTIFY Or NewStyle = WS_EX_TOPMOST Or NewStyle = WS_EX_TRANSPARENT Then
    lWSIndex = GWL_EXSTYLE
Else
    lWSIndex = GWL_STYLE
End If

'\\ Get the current setting of that style bit
lStyle = GetWindowLongApi(m_hwnd, lWSIndex)

'\\ Add the new style bit to it
lStyle = lStyle Or NewStyle

'\\ Set it to the window
lRet = SetWindowLongApi(m_hwnd, lWSIndex, lStyle)

'\\ For some styles to take effect, the window must be redrawn...


SetWindowStyle = Me.IsWindowStyleSet(lStyle)

End Function

Public Property Get SystemMenu() As ApiMenu

Dim lRet As Long

lRet = GetSystemMenuApi(m_hwnd, False)
If Err.LastDllError = 0 And lRet > 0 Then
    If mSystemMenu Is Nothing Then
        Set mSystemMenu = New ApiMenu
    End If
    mSystemMenu.hMenu = lRet
    Set SystemMenu = mSystemMenu
End If

End Property

Public Property Get Thread() As ApiThread

Dim lRet As Long
Dim lProcess As Long

Dim threadThis As ApiThread

Set threadThis = New ApiThread
lRet = GetWindowThreadProcessId(Me.hwnd, lProcess)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiWindow:Thread", APIDispenser.LastSystemError
Else
    threadThis.ThreadId = lRet
End If

Set Thread = threadThis

End Property

Friend Sub TriggerEvent(wMsg As WindowMessages, wParam As Long, lParam As Long, bDiscardMsg As Boolean, lWndProcRet As Long)

Dim Cancel As Boolean

Dim lRet As Long

'\\ Used in wm_vscroll and wm_hscroll processing
Dim ScrollMessage As enScrollMessages
Dim ScrollPosition As Long
Static LastHScroll As Long
Static LastVScroll As Long

'\\ Used in WM_MOVE processing
Dim x As Long
Dim y As Long


    If wMsg = WM_ACTIVATEAPP Then
        '\\ This app is being activetd or deactivated...
        RaiseEvent ActiveApplicationChanged(CBool(wParam), lParam, Cancel)
    
    ElseIf wMsg = WM_CAPTURECHANGED Then
        '\\This window lost mouse capture
        RaiseEvent LostCapture(lParam, Cancel)
    
    ElseIf wMsg = WM_CHAR Then
        '\\ Keypress was intercepted by this window...
        Dim lRepetition As Long
        Dim lScanCode As Long
        Dim bExtendedKey As Boolean
        Dim bAltDown As Boolean
        Dim bAlreadyPressed As Boolean
        Dim bBeingPressed As Boolean
        With APIDispenser
            lRepetition = .GetBits(lParam, 0, 15)
            lScanCode = .GetBits(lParam, 16, 23)
            bExtendedKey = .GetBits(lParam, 24, 24)
            bAltDown = .GetBits(lParam, 29, 22)
            bAlreadyPressed = .GetBits(lParam, 30, 30)
            bBeingPressed = .GetBits(lParam, 31, 31)
        End With
        RaiseEvent KeyPressed(wParam, lRepetition, lScanCode, bExtendedKey, bAltDown, bAlreadyPressed, bBeingPressed, Cancel)
    
    ElseIf wMsg = WM_COMPACTING Then
        '\\ Windows is low on memory and is compacting memory...
        RaiseEvent LowMemory(wParam)
        '\\ It is not sensible to have a cancel option on this
    
    ElseIf wMsg = WM_MOVE Then
        x = APIDispenser.LoWord(lParam)
        y = APIDispenser.HiWord(lParam)
        RaiseEvent Move(x, y, Cancel)
    
    ElseIf wMsg = WM_HSCROLL Then
        '\\ Horizontal scrollbar event has occured
        ScrollMessage = APIDispenser.LoWord(wParam)
        ScrollPosition = APIDispenser.HiWord(wParam)
        RaiseEvent HorizontalScroll(ScrollMessage, ScrollPosition, Cancel)
        Call Me.ScrollWindow(ScrollPosition - LastHScroll, 0, SW_SCROLLCHILDREN)
        LastHScroll = ScrollPosition
        
    ElseIf wMsg = WM_VSCROLL Then
        '\\ Vertical scrollbar event has occured
        ScrollMessage = APIDispenser.LoWord(wParam)
        ScrollPosition = APIDispenser.HiWord(wParam)
        RaiseEvent VerticalScroll(ScrollMessage, ScrollPosition, Cancel)
        Call Me.ScrollWindow(0, ScrollPosition - LastVScroll, SW_SCROLLCHILDREN)
        LastVScroll = ScrollPosition
        
    ElseIf wMsg = WM_SETTINGCHANGE Then
        RaiseEvent WindowsSettingsChanged(wParam)
    
    ElseIf wMsg = WM_WININICHANGE Then
        Dim sSection As String
        sSection = StringFromPointer(lParam, 1024)
        RaiseEvent WindowsINIChanged(sSection)
        
    '\\ Mouse button down
    ElseIf wMsg = WM_NCLBUTTONDOWN Then
        x = APIDispenser.LoWord(lParam)
        y = APIDispenser.HiWord(lParam)
        RaiseEvent NonClientMouseDown(wParam, vbLeftButton, x, y)
        
    ElseIf wMsg = WM_NCMBUTTONDOWN Then
        x = APIDispenser.LoWord(lParam)
        y = APIDispenser.HiWord(lParam)
        RaiseEvent NonClientMouseDown(wParam, vbMiddleButton, x, y)
        
    ElseIf wMsg = WM_NCRBUTTONDOWN Then
        x = APIDispenser.LoWord(lParam)
        y = APIDispenser.HiWord(lParam)
        RaiseEvent NonClientMouseDown(wParam, vbRightButton, x, y)
        
    '\\ Mouse button up
    ElseIf wMsg = WM_NCLBUTTONUP Then
        x = APIDispenser.LoWord(lParam)
        y = APIDispenser.HiWord(lParam)
        RaiseEvent NonClientMouseUp(wParam, vbLeftButton, x, y)
        
    ElseIf wMsg = WM_NCMBUTTONUP Then
        x = APIDispenser.LoWord(lParam)
        y = APIDispenser.HiWord(lParam)
        RaiseEvent NonClientMouseUp(wParam, vbMiddleButton, x, y)
        
    ElseIf wMsg = WM_NCRBUTTONUP Then
        x = APIDispenser.LoWord(lParam)
        y = APIDispenser.HiWord(lParam)
        RaiseEvent NonClientMouseUp(wParam, vbRightButton, x, y)

    '\\ Mouse move
    ElseIf wMsg = WM_NCMOUSEMOVE Then
        x = APIDispenser.LoWord(lParam)
        y = APIDispenser.HiWord(lParam)
        RaiseEvent NonClientMouseMove(wParam, x, y)
        '\\ If the mouse moved over a menu...
        If wParam = HTMENU Then
            Dim sCaption As String
            Dim ptThis As APIPoint
            Set ptThis = New APIPoint
            With ptThis
                .x = x
                .y = y
            End With
            sCaption = Me.Menu.CaptionFromPoint(Me, ptThis)
            If sCaption <> "" Then
                RaiseEvent MouseOverMenu(sCaption)
            End If
        End If

    '\\ Minimum and maximum info.....
    ElseIf wMsg = WM_GETMINMAXINFO Then
        '\\ wParam is empty, lParam is a pointer to a minmaxinfo....
        Dim myMinMaxInfo As MINMAXINFO
        Dim lMaxHeight As Long, lMaxWidth As Long
        Dim lMaxPositionTop As Long
        Dim lMaxPositionLeft As Long
        Dim lMinTrackWidth As Long, lMinTrackheight As Long
        Dim lMaxTrackWidth As Long, lMaxTrackHeight As Long
        Call CopyMemoryMinMaxInfo(myMinMaxInfo, lParam, LenB(myMinMaxInfo))
        If Err.LastDllError <> 0 Then
            ReportError Err.LastDllError, "ApiWindow:WM_GETMINMAXINFO", APIDispenser.LastSystemError
        Else
            With myMinMaxInfo
                '\\ Maximum height and width
                lMaxWidth = .ptMaxSize.x
                lMaxHeight = .ptMaxSize.y
                '\\ Maximum left and top
                lMaxPositionLeft = .ptMaxPosition.x
                lMaxPositionTop = .ptMaxPosition.y
                '\\ Maxiumum track size
                lMaxTrackWidth = .ptMaxTrackSize.x
                lMaxTrackHeight = .ptMaxTrackSize.y
                '\\ Minimum track size
                lMinTrackWidth = .ptMinTrackSize.x
                lMinTrackheight = .ptMinTrackSize.y
            End With
            RaiseEvent MinMaxSize(lMaxHeight, lMaxWidth, lMaxPositionTop, lMaxPositionLeft, lMinTrackWidth, lMinTrackheight, lMaxTrackWidth, lMaxTrackHeight)
            With myMinMaxInfo
                '\\ Maximum height and width
                 .ptMaxSize.x = lMaxWidth
                 .ptMaxSize.y = lMaxHeight
                '\\ Maximum left and top
                 .ptMaxPosition.x = lMaxPositionLeft
                 .ptMaxPosition.y = lMaxPositionTop
                '\\ Maxiumum track size
                 .ptMaxTrackSize.x = lMaxTrackWidth
                 .ptMaxTrackSize.y = lMaxTrackHeight
                '\\ Minimum track size
                 .ptMinTrackSize.x = lMinTrackWidth
                 .ptMinTrackSize.y = lMinTrackheight
            End With
            Call CopyMinMaxToMemoryInfo(lParam, myMinMaxInfo, LenB(myMinMaxInfo))
        End If

    Else
        '\\ For all other events
        RaiseEvent WindowMessageFired(wMsg, wParam, lParam, bDiscardMsg, lWndProcRet)
    End If
    
    lWndProcRet = Cancel
    
End Sub

'\\ --[UnSetWindowStyle]----------------------------------------------------------------------
'\\ Unets the style bit specified to the window specified.  Note that many window style
'\\ bits cannot be used at run time :. use this with caution
'\\ Returns true if it succeeded
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function UnSetWindowStyle(ByVal NewStyle As enWindowStyles) As Boolean

Dim lStyle As Long
Dim lWSIndex As Long
Dim lRet As Long

If NewStyle = WS_EX_ACCEPTFILES Or NewStyle = WS_EX_DLGMODALFRAME Or NewStyle = WS_EX_NOPARENTNOTIFY Or NewStyle = WS_EX_TOPMOST Or NewStyle = WS_EX_TRANSPARENT Then
    lWSIndex = GWL_EXSTYLE
Else
    lWSIndex = GWL_STYLE
End If

'\\ Get the current setting of that style bit
lStyle = GetWindowLongApi(m_hwnd, lWSIndex)

'\\ Add the new style bit to it
If IsWindowStyleSet(NewStyle) Then
    lStyle = lStyle Xor NewStyle
End If

'\\ Set it to the window
lRet = SetWindowLongApi(m_hwnd, lWSIndex, lStyle)

UnSetWindowStyle = Not (IsWindowStyleSet(lStyle))

End Function
Public Property Get Enabled() As Boolean

    Enabled = IsWindowEnabledApi(m_hwnd)
    
End Property

Public Property Get Unicode() As Boolean

    Unicode = IsWindowUnicodeApi(m_hwnd)
    
End Property

Public Property Get Visible() As Boolean

    Visible = IsWindowVisibleApi(m_hwnd)
    
End Property

Public Property Get IsWindow() As Boolean

IsWindow = IsWindowApi(m_hwnd)

End Property

Public Property Get IsZoomed() As Boolean

    IsZoomed = IsZoomedApi(m_hwnd)
    
End Property

Public Property Get Key() As String

    Key = "HWND:" & m_hwnd

End Property

Public Function SetWindowLong(ByVal Index As enGetWindowLong, ByVal newValue As Long) As Long

Dim lRet As Long

lRet = SetWindowLongApi(m_hwnd, Index, newValue)
If Err.LastDllError = 0 Then
    If Index = GWL_WNDPROC Then
        '\\ We must store old proc address...
        mOldProcAddress = lRet
    End If
    SetWindowLong = lRet
End If

End Function

'\\ --[WindowBaseClass]----------------------------------------------------
'\\ Returns the base class of this window
'\\ -----------------------------------------------------------------------
Public Property Get WindowBaseClass() As String

Dim sClass As String

sClass = Me.ClassName

Select Case sClass
Case "BUTTON", "ThunderCheckBox", "ThunderRTCheckBox", _
     "ThunderCommandButton", "ThunderRTCommandButton", _
     "ThunderFrame", "ThunderRTFrame", _
     "ThunderOptionButton", "ThunderRTOptionButton"
    '\\ Classes derived from the BUTTON base class
    WindowBaseClass = "BUTTON"
Case "COMBOBOX", "ThunderComboBox", "ThunderRTComboBox", _
     "ThunderDriveListBox", "ThunderRTDriveListBox"
    
    '\\ Classes derived from the COMBOBOX base class
    WindowBaseClass = "COMBOBOX"
Case "LISTBOX", _
     "ThunderDirListBox", "ThunderRTDirListBox", _
     "ThunderListBox", "ThunderRTListBox", _
     "ThunderFileListBox", "ThunderRTFileListBox"
     
    '\\ Classes derived from the LISTBOX base class
    WindowBaseClass = "LISTBOX"
Case "SCROLLBAR", _
    "ThunderVScrollBar", "ThunderRTVScrollBar", _
    "ThunderHScrollBar", "ThunderRTHScrollBar"
    
    '\\ Classes derived from the SCROLLBAR base class
    WindowBaseClass = "SCROLLBAR"
Case "EDIT", _
     "ThunderTextBox", "ThunderRTTextBox"
    '\\ Classes derived from the SCROLLBAR base class
    WindowBaseClass = "EDIT"
Case Else
    WindowBaseClass = "STATIC"
End Select

End Property

Public Function WindowFromPoint(ByVal xPoint As Long, ByVal yPoint As Long) As ApiWindow

Dim hwnd As Long
Dim window As ApiWindow

hwnd = WindowFromPointApi(xPoint, yPoint)
If IsWindowApi(hwnd) Then
    Set window = New ApiWindow
    window.hwnd = hwnd
    Set WindowFromPoint = window
End If

End Function


Public Property Get WindowText() As String

Dim lRet As Long

Dim sRet As String

lRet = Me.SendMessage(WM_GETTEXTLENGTH, 0, 0)
If lRet > 0 Then
    sRet = String$(lRet + 1, 0)
    lRet = SendMessageByString(m_hwnd, WM_GETTEXT, Len(sRet), sRet)
    '\\ Returns length up to NULL terminator
    If lRet > 0 Then
        sRet = Left$(sRet, lRet)
    End If
ElseIf Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiWindow:WindowText", APIDispenser.LastSystemError
Else
    sRet = String$(1024, 0)
    lRet = GetWindowTextApi(m_hwnd, sRet, Len(sRet))
    sRet = Left$(sRet, lRet)
End If

WindowText = sRet

End Property

Public Property Get WndClass() As APIWndClass

If mWndClass Is Nothing Then
    Set mWndClass = New APIWndClass
End If
If mWndClass.CreateFromClassname(vbNull, Me.ClassName) Then
    Set WndClass = mWndClass
End If

End Property

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

    '\\ If the old proc address is set, unhook this...
    If mOldProcAddress > 0 Then
        Call SetWindowLong(GWL_WNDPROC, mOldProcAddress)
    End If

    
    Set mMenu = Nothing
    Set mSystemMenu = Nothing
    Set mWndClass = Nothing
    Set mRegion = Nothing
    Set mDeviceContext = Nothing
    
End Sub

