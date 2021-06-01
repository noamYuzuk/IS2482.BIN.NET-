Attribute VB_Name = "mMthCoordSys"
Option Explicit


'ROTATE POINT ABOUT CS AXIS

Public Sub uRotatePtAboutAxis(d1 As Double, d2 As Double, _
                              dARad As Double, _
                              d1O As Double, d2O As Double)
   'about Z axis: d1 - x, d2 - y
   'about Y axis: d1 - z, d2 - x
   'about X axis: d1 - y, d2 - z
   'dARad in radians

   d1O = d1 * Cos(dARad) - d2 * Sin(dARad)
   d2O = d1 * Sin(dARad) + d2 * Cos(dARad)
End Sub


Public Sub uRotatePtAboutAxis1( _
                             ByVal d1 As Double, ByVal d2 As Double, _
                             dA As Double, _
                             d1O As Double, d2O As Double)
   'about Z axis: d1 - x, d2 - y
   'about Y axis: d1 - z, d2 - x
   'about X axis: d1 - y, d2 - z
   'dA in radians

   Dim _
   dHypotenuse As Double, dAHyp As Double, dASum As Double
'dd
   dHypotenuse = Sqr(d1 * d1 + d2 * d2)
   dAHyp = fdAtn2(d1, d2)
   dASum = dAHyp + dA

   d1O = dHypotenuse * Cos(dASum)
   d2O = dHypotenuse * Sin(dASum)
End Sub


