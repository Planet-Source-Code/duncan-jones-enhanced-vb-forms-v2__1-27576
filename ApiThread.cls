VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Type LDT_BYTES  ' Defined for use in LDT_ENTRY Type
    BaseMid As Byte
    Flags1 As Byte
    Flags2 As Byte
    BaseHi As Byte
End Type
Private Type LDT_ENTRY
    LimitLow As Integer
    BaseLow As Integer
    HighWord As LDT_BYTES
End Type

Private Type CONTEXT
    FltF0 As Double
    FltF1 As Double
    FltF2 As Double
    FltF3 As Double
    FltF4 As Double
    FltF5 As Double
    FltF6 As Double
    FltF7 As Double
    FltF8 As Double
    FltF9 As Double
    FltF10 As Double
    FltF11 As Double
    FltF12 As Double
    FltF13 As Double
    FltF14 As Double
    FltF15 As Double
    FltF16 As Double
    FltF17 As Double
    FltF18 As Double
    FltF19 As Double
    FltF20 As Double
    FltF21 As Double
    FltF22 As Double
    FltF23 As Double
    FltF24 As Double
    FltF25 As Double
    FltF26 As Double
    FltF27 As Double
    FltF28 As Double
    FltF29 As Double
    FltF30 As Double
    FltF31 As Double

    IntV0 As Double
    IntT0 As Double
    IntT1 As Double
    IntT2 As Double
    IntT3 As Double
    IntT4 As Double
    IntT5 As Double
    IntT6 As Double
    IntT7 As Double
    IntS0 As Double
    IntS1 As Double
    IntS2 As Double
    IntS3 As Double
    IntS4 As Double
    IntS5 As Double
    IntFp As Double
    IntA0 As Double
    IntA1 As Double
    IntA2 As Double
    IntA3 As Double
    IntA4 As Double
    IntA5 As Double
    IntT8 As Double
    IntT9 As Double
    IntT10 As Double
    IntT11 As Double
    IntRa As Double
    IntT12 As Double
    IntAt As Double
    IntGp As Double
    IntSp As Double
    IntZero As Double

    Fpcr As Double
    SoftFpcr As Double

    Fir As Double
    Psr As Long

    ContextFlags As Long
    Fill(4) As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function SetThreadAffinityMask Lib "kernel32" (ByVal hThread As Long, ByVal dwThreadAffinityMask As Long) As Long

Private Declare Function SetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As CONTEXT) As Long
Private Declare Function GetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As CONTEXT) As Long

Private Declare Function SetThreadDesktop Lib "user32" (ByVal hDesktop As Long) As Long
Private Declare Function GetThreadDesktop Lib "user32" (ByVal dwThread As Long) As Long

Private Declare Function SetThreadLocale Lib "kernel32" (ByVal Locale As Long) As Long
Private Declare Function GetThreadLocale Lib "kernel32" () As Long

Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long

Private Declare Function SetThreadToken Lib "advapi32" (Thread As Long, ByVal Token As Long) As Long
Private Declare Function GetThreadSelectorEntry Lib "kernel32" (ByVal hThread As Long, ByVal dwSelector As Long, lpSelectorEntry As LDT_ENTRY) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Declare Function GetThreadTimes Lib "kernel32" (ByVal hThread As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long

Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long

Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long

'\\ Member variables
Private mHThreadId

Private mBaseAddress

Public Enum enThreadPriorities
    THREAD_BASE_PRIORITY_IDLE = -15
    THREAD_BASE_PRIORITY_LOWRT = 15
    THREAD_BASE_PRIORITY_MAX = 2
    THREAD_BASE_PRIORITY_MIN = -2
    THREAD_PRIORITY_NORMAL = 0
    THREAD_PRIORITY_ABOVE_NORMAL = 1
    THREAD_PRIORITY_BELOW_NORMAL = -1
End Enum

Public Property Let BaseAddress(ByVal lProcAddress As Long)

If Not IsBadCodePtr(lProcAddress) Then
    mBaseAddress = lProcAddress
Else
    ReportError vbError + 100, "ApiThread:baseAddress", "Invalid code address"
End If
    
End Property


Public Property Get EndTime() As APIFileTime
Dim TimeStart As FILETIME
Dim TimeEnd As FILETIME
Dim TimeKernel As FILETIME
Dim TimeUser As FILETIME

Dim lRet As Long
Dim timeThis As APIFileTime
Set timeThis = New APIFileTime

lRet = GetThreadTimes(mHThreadId, TimeStart, TimeEnd, TimeKernel, TimeUser)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiThread:StartTime", APIDispenser.LastSystemError
Else
    timeThis.CreateFromPointer (VarPtr(TimeEnd))
End If

Set EndTime = timeThis

End Property

Public Property Get KernelTime() As APIFileTime
Dim TimeStart As FILETIME
Dim TimeEnd As FILETIME
Dim TimeKernel As FILETIME
Dim TimeUser As FILETIME

Dim lRet As Long
Dim timeThis As APIFileTime
Set timeThis = New APIFileTime

lRet = GetThreadTimes(mHThreadId, TimeStart, TimeEnd, TimeKernel, TimeUser)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiThread:StartTime", APIDispenser.LastSystemError
Else
    timeThis.CreateFromPointer (VarPtr(TimeKernel))
End If

Set KernelTime = timeThis

End Property

Public Property Let Priority(ByVal newPriority As enThreadPriorities)

Dim lRet As Long

lRet = SetThreadPriority(mHThreadId, newPriority)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiThread:Priority", APIDispenser.LastSystemError
End If

End Property

Public Property Get Priority() As enThreadPriorities

Dim lRet As Long

lRet = GetThreadPriority(mHThreadId)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiThread:Priority", APIDispenser.LastSystemError
Else
    Priority = lRet
End If

End Property

Public Property Get StartTime() As APIFileTime

Dim TimeStart As FILETIME
Dim TimeEnd As FILETIME
Dim TimeKernel As FILETIME
Dim TimeUser As FILETIME

Dim lRet As Long
Dim timeThis As APIFileTime
Set timeThis = New APIFileTime

lRet = GetThreadTimes(mHThreadId, TimeStart, TimeEnd, TimeKernel, TimeUser)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiThread:StartTime", APIDispenser.LastSystemError
Else
    timeThis.CreateFromPointer (VarPtr(TimeStart))
End If

Set StartTime = timeThis

End Property

Public Property Let ThreadId(ByVal newId As Long)

    mHThreadId = newId
    
End Property

Public Property Get ThreadId() As Long

    ThreadId = mHThreadId
    
End Property
Public Property Get UserTime() As APIFileTime

Dim TimeStart As FILETIME
Dim TimeEnd As FILETIME
Dim TimeKernel As FILETIME
Dim TimeUser As FILETIME

Dim lRet As Long
Dim timeThis As APIFileTime
Set timeThis = New APIFileTime

lRet = GetThreadTimes(mHThreadId, TimeStart, TimeEnd, TimeKernel, TimeUser)
If Err.LastDllError > 0 Then
    ReportError Err.LastDllError, "ApiThread:StartTime", APIDispenser.LastSystemError
Else
    timeThis.CreateFromPointer (VarPtr(TimeUser))
End If

Set UserTime = timeThis

End Property


