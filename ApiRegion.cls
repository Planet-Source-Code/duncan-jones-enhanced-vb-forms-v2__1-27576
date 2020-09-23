VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mhRgn As Long


Public Property Let HRgn(ByVal newHRgn As Long)

If newHRgn <> mhRgn Then
    mhRgn = newHRgn
End If

End Property

Public Property Get HRgn() As Long

    HRgn = mhRgn
    
End Property

