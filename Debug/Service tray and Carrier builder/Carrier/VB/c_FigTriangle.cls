VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFigTriangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Using: mMth_CoordSys.bas, mMth_Arr.bas, c_Draw.cls

Public _
   p_dXCen As Double, p_dYCen As Double, _
   p_dARotRad As Double
   
'Private m_oDraw As New clsDraw
'Private m_oDraw As New clsDrawGDI


Private _
   m_dH As Double, _
   m_dX(2) As Double, m_dY(2) As Double


Public Property Let p_dH(d As Double)
   
   m_dH = d
   m_dX(0) = -2# / 3 * m_dH * gc_dCos30
   m_dY(0) = -1# / 3 * m_dH
   m_dX(1) = -m_dX(0)
   m_dY(1) = m_dY(0)
   m_dX(2) = 0
   m_dY(2) = 2# / 3 * m_dH
End Property

Public Property Get p_dH() As Double

   p_dH = m_dH
End Property



Public Sub uDraw(pic As PictureBox, lColour As Long, _
                 Optional lColorFill As Long = -1)
   
   Const c_iIUPts As Integer = 2
   Dim _
   dXO As Double, dYO As Double, _
   dX(c_iIUPts) As Double, dY(c_iIUPts) As Double, _
   i As Integer, oDraw As New clsDrawGDI
   
   For i = 0 To c_iIUPts
      uRotatePtAboutAxis m_dX(i), m_dY(i), p_dARotRad, dXO, dYO
      dX(i) = p_dXCen + dXO
      dY(i) = p_dYCen - dYO
   Next i
   
   With oDraw
   .uSetObject pic
   .uSetContour 1, vbBlue
   .uSetFillColor lColorFill   'light yellow
   .uCurveC dX(), dY()
   End With
End Sub


