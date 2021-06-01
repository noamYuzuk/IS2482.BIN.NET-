VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFigCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Using: mMth_CoordSys.bas, mMth_Arr.bas, c_Draw.cls

Public _
   p_dXCen As Double, p_dYCen As Double
   
'Rem: Added in order to standardize with other figures
Public _
   p_dARotRad As Double
   
Private _
   m_dD As Double


Public Property Let p_dD(d As Double)
   
   m_dD = d
End Property

Public Property Get p_dD() As Double

   p_dD = m_dD
End Property


Public Sub uDraw(pic As PictureBox, lColor As Long, _
                 Optional lColorFill As Long = -1)
   
   If lColorFill <> -1 Then
      pic.FillColor = lColorFill
      pic.FillStyle = 0    'solid
   End If
   pic.Circle (p_dXCen, p_dYCen), m_dD / 2, lColor
   pic.FillStyle = 1       'transparent
End Sub
