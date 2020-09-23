VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiColourAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Declare Function GetColorAdjustment Lib "gdi32" (ByVal hdc As Long, lpca As COLORADJUSTMENT) As Long
Private Type COLORADJUSTMENT
    caSize As Integer
    caFlags As Integer
    caIlluminantIndex As Integer
    caRedGamma As Integer
    caGreenGamma As Integer
    caBlueGamma As Integer
    caReferenceBlack As Integer
    caReferenceWhite As Integer
    caContrast As Integer
    caBrightness As Integer
    caColorfulness As Integer
    caRedGreenTint As Integer
End Type
Private Declare Function SetColorAdjustment Lib "gdi32" (ByVal hdc As Long, lpca As COLORADJUSTMENT) As Long

Private mParentDC As Long

Private mColourAdjustment As COLORADJUSTMENT

Public Enum enColourAdjustFlags
     CA_NEGATIVE = &H1
     CA_LOG_FILTER = &H2
End Enum

Public Enum enIlluninantIndexes
    ILLUMINANT_DEVICE_DEFAULT = 0
    ILLUMINANT_TUNGSTEN = 1
    ILLUMINANT_NOON_SUNLIGHT = 2
    ILLUMINANT_NTSC_DAYLIGHT = 3
    ILLUMINANT_NORMAL_PRINT = 4
    ILLUMINANT_BOND_PRINT = 5
    ILLUMINANT_STANDARD_DAYLIGHT = 6
    ILLUMINANT_NORTHEN_DAYLIGHT = 7
    ILLUMINANT_FLOURESCENT_LIGHT = 8
End Enum

Private Const RGB_GAMMA_MIN    As Long = 2500
Private Const RGB_GAMMA_MAX    As Long = 65000

'/* Min and max for ReferenceBlack and ReferenceWhite */
Private Const REFERENCE_WHITE_MIN  As Long = 6000
Private Const REFERENCE_WHITE_MAX  As Long = 10000
Private Const REFERENCE_BLACK_MIN  As Long = 0
Private Const REFERENCE_BLACK_MAX   As Long = 4000
Private Const COLOR_ADJ_MIN    As Long = -100
Private Const COLOR_ADJ_MAX    As Long = 100

Public Property Get BlueGamma() As Integer

    BlueGamma = mColourAdjustment.caBlueGamma
    
End Property

Public Property Let BlueGamma(ByVal newGamma As Integer)

    If newGamma < RGB_GAMMA_MIN Then
        newGamma = RGB_GAMMA_MIN
    ElseIf newGamma > RGB_GAMMA_MAX Then
        newGamma = RGB_GAMMA_MAX
    End If
    
    If mColourAdjustment.caBlueGamma <> newGamma Then
       mColourAdjustment.caBlueGamma = newGamma
       Call RefreshColourAdjustment
    End If
    
End Property

Public Property Let Brightness(ByVal newValue As Integer)

    If newValue < COLOR_ADJ_MIN Then
        newValue = COLOR_ADJ_MIN
    ElseIf newValue > COLOR_ADJ_MAX Then
        newValue = COLOR_ADJ_MAX
    End If
    
    If newValue <> mColourAdjustment.caBrightness Then
        mColourAdjustment.caBrightness = newValue
        Call RefreshColourAdjustment
    End If
    
End Property

Public Property Get Brightness() As Integer

    Brightness = mColourAdjustment.caBrightness
    
End Property

Public Property Let Colourfulness(ByVal newValue As Integer)

    If newValue < COLOR_ADJ_MIN Then
        newValue = COLOR_ADJ_MIN
    ElseIf newValue > COLOR_ADJ_MAX Then
        newValue = COLOR_ADJ_MAX
    End If
    
    If newValue <> mColourAdjustment.caColorfulness Then
        mColourAdjustment.caColorfulness = newValue
        Call RefreshColourAdjustment
    End If

End Property

Public Property Get Colourfulness() As Integer

Colourfulness = mColourAdjustment.caColorfulness

End Property

Public Property Let Contrast(ByVal newContrast As Integer)

    If newContrast < COLOR_ADJ_MIN Then
        newContrast = COLOR_ADJ_MIN
    ElseIf newContrast > COLOR_ADJ_MAX Then
        newContrast = COLOR_ADJ_MAX
    End If
    
    If newContrast <> mColourAdjustment.caContrast Then
        mColourAdjustment.caContrast = newContrast
        Call RefreshColourAdjustment
    End If
    
End Property

Public Property Get Contrast() As Integer

    Contrast = mColourAdjustment.caContrast
    
End Property

Public Property Let GreenGamma(ByVal newGamma As Integer)

    If newGamma < RGB_GAMMA_MIN Then
        newGamma = RGB_GAMMA_MIN
    ElseIf newGamma > RGB_GAMMA_MAX Then
        newGamma = RGB_GAMMA_MAX
    End If
    
    If mColourAdjustment.caGreenGamma <> newGamma Then
       mColourAdjustment.caGreenGamma = newGamma
       Call RefreshColourAdjustment
    End If
    
End Property

Public Property Get GreenGamma() As Integer

    GreenGamma = mColourAdjustment.caGreenGamma
    
End Property


Public Property Let IlluminantIndex(ByVal newIndex As enIlluninantIndexes)

    If mColourAdjustment.caIlluminantIndex <> newIndex Then
        mColourAdjustment.caIlluminantIndex = newIndex
        Call RefreshColourAdjustment
    End If
    
End Property

Public Property Get IlluminantIndex() As enIlluninantIndexes

IlluminantIndex = mColourAdjustment.caIlluminantIndex

End Property

Public Property Set ParentDC(ByVal newDC As ApiDeviceContext)

Dim lRet As Long

If newDC.hdc <> mParentDC Then
    mParentDC = newDC.hdc
    lRet = GetColorAdjustment(mParentDC, mColourAdjustment)
    If Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, "ApiColourAdjustment:ParentDC", APIDispenser.LastSystemError)
    End If
End If

End Property
Public Property Get RedGamma() As Integer

    RedGamma = mColourAdjustment.caRedGamma
    
End Property

Public Property Let RedGamma(ByVal newGamma As Integer)

    If newGamma < RGB_GAMMA_MIN Then
        newGamma = RGB_GAMMA_MIN
    ElseIf newGamma > RGB_GAMMA_MAX Then
        newGamma = RGB_GAMMA_MAX
    End If
    
    If mColourAdjustment.caRedGamma <> newGamma Then
       mColourAdjustment.caRedGamma = newGamma
       Call RefreshColourAdjustment
    End If
    
End Property

Public Property Let RedGreenTint(ByVal newTint As Integer)

    If newTint < COLOR_ADJ_MIN Then
        newTint = COLOR_ADJ_MIN
    ElseIf newTint > COLOR_ADJ_MAX Then
        newTint = COLOR_ADJ_MAX
    End If
    
    If newTint <> mColourAdjustment.caRedGreenTint Then
        mColourAdjustment.caRedGreenTint = newTint
        Call RefreshColourAdjustment
    End If
    
End Property

Public Property Get RedGreenTint() As Integer

    RedGreenTint = mColourAdjustment.caRedGreenTint
    
End Property

Private Function RefreshColourAdjustment()

Dim lRet As Long

If mParentDC > 0 Then
    lRet = SetColorAdjustment(mParentDC, mColourAdjustment)
    If Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, "ApiColourAdjustment:RefreshColourAdjustment", APIDispenser.LastSystemError)
    End If
End If

End Function


