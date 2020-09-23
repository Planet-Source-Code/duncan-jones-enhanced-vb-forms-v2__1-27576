VERSION 5.00
Object = "*\AEventCtl.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin EventVB.VBEventWindow VBEventWindow1 
      Left            =   600
      Top             =   480
      _ExtentX        =   1720
      _ExtentY        =   1508
      MaxHeight       =   100
      MaxWidth        =   200
      MinTrackHeight  =   100
      MaxTrackWidth   =   300
      MaxTrackHeight  =   200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    VBEventWindow1.ParentForm = Me.hWnd

End Sub

