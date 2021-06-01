Attribute VB_Name = "mMth_Trigonom"
Option Explicit

'Using m_Msg.bas

Public Const _
   gc_dPi2 As Double = 1.5707963267949, _
   gc_d2Pi3 As Double = 2.0943951023932, _
   gc_dPi As Double = 3.14159265358979, _
   gc_d3Pi2 As Double = 4.71238898038469, _
   gc_d2Pi As Double = 6.28318530717959
Public Const _
   gc_dToRad As Double = 1.74532925199433E-02, _
   gc_dToDeg As Double = 57.2957795130823
Public Const _
   gc_dTan30 As Double = 0.577350269189626, _
   gc_dCos30 As Double = 0.866025403784439
   
Public Const _
   gc_dA30 As Double = 30#, _
   gc_dA90 As Double = 90#, _
   gc_dA120 As Double = 120#, _
   gc_dA270 As Double = 270#


'TRIGONOMETRIC FUNCTIONS

Public Function fdAtn2(dX As Double, dY As Double, _
                       Optional bDegrees As Boolean) As Double
'Calculate an arctangent in range from 0 to 2Pi (or 360°)
   Dim dRes As Double
'dd
   On Error GoTo Error
   
   If dX = 0# Then
      If dY = 0# Then
         g_uMsgW "X=0 and Y=0", "mMth_Trigonom.fdAtn2"
         dRes = 0#
      ElseIf dY > 0# Then
         dRes = gc_dPi2
      ElseIf dY < 0# Then
         dRes = gc_d3Pi2
      End If
   ElseIf dX > 0# And dY >= 0# Then
      dRes = Atn(dY / dX)
   ElseIf dX > 0# And dY < 0# Then
      dRes = gc_d2Pi + Atn(dY / dX)
   ElseIf dX < 0# And dY > 0# Then
      dRes = gc_dPi + Atn(dY / dX)
   Else
      dRes = gc_dPi + Atn(dY / dX)
   End If
   If bDegrees Then dRes = gc_dToDeg * dRes
   fdAtn2 = dRes
   Exit Function
Error:
   g_uMsgE "mMth_Trigonom.fdAtn2"
End Function


Public Function fdACos(dCosA As Double) As Double
   Dim dTanA As Double, dRes As Double
'dd
   If Round(dCosA, 14) = 1# Then
      dRes = 0
   ElseIf Round(dCosA, 14) = -1# Then
      dRes = gc_dPi
   ElseIf Round(dCosA, 14) = 0# Then
      dRes = gc_dPi2
   Else
      dTanA = Sqr(1 / dCosA / dCosA - 1)
      dRes = Atn(dTanA)
      If dCosA < 0 Then dRes = gc_dPi - dRes
   End If
   
   fdACos = dRes
End Function


'MISCELANEOUS

Public Function fdSub2Pi(dARad As Double)

   'dARad - angle in radians
   If dARad > gc_d2Pi Then
      fdSub2Pi = dARad - gc_d2Pi
   Else
      fdSub2Pi = dARad
   End If
End Function


'ANGLES OF ROTATION

Public Function fbA0GtA1(dA0 As Double, dA1 As Double) As Boolean
   'Check if first angle grater then second.
   Dim _
   dAQu1Dn As Double, dAQu1Up As Double, _
   dAQu4Dn As Double, dAQu4Up As Double, _
   bA0In1 As Boolean, bA1In1 As Boolean, _
   bA0In4 As Boolean, bA1In4 As Boolean
'dd
   dAQu1Dn = 0:         dAQu1Up = gc_dPi2
   dAQu4Dn = gc_d3Pi2:   dAQu4Up = gc_d2Pi
   
   bA0In1 = (dA0 > dAQu1Dn) And (dA0 < dAQu1Up)
   bA0In4 = (dA0 > dAQu4Dn) And (dA0 < dAQu4Up)
   bA1In1 = (dA1 > dAQu1Dn) And (dA1 < dAQu1Up)
   bA1In4 = (dA1 > dAQu4Dn) And (dA1 < dAQu4Up)
   
   If bA0In1 And bA1In4 Then
      fbA0GtA1 = True
   ElseIf bA0In4 And bA1In1 Then
      fbA0GtA1 = False
   ElseIf dA0 > dA1 Then
      fbA0GtA1 = True
   End If
End Function


Public Function fdA02Pi(ByVal dA As Double) As Double
   If dA > gc_d2Pi Then
      dA = dA - gc_d2Pi
   ElseIf dA < 0 Then
      dA = gc_d2Pi + dA
   End If
   fdA02Pi = dA
End Function


Public Function fdHypotenuse(dX As Double, dY As Double) As Double
   fdHypotenuse = Sqr(dX * dX + dY * dY)
End Function


Private Function fdSubA(dA0, dA1) As Double
   Dim _
   dRes As Double
'dd
   dRes = dA0 - dA1
   If dRes < -gc_dPi Then
      dRes = gc_d2Pi + dRes
   ElseIf dRes > gc_dPi Then
      dRes = dRes - gc_d2Pi
   End If
   fdSubA = dRes
End Function
   

Private Function fdAddA(dA0, dA1) As Double
   Dim _
   dRes As Double
'dd
   dRes = dA0 + dA1
   If dRes > gc_d2Pi Then
      dRes = dRes - gc_d2Pi
   ElseIf dRes < 0# Then
      dRes = gc_d2Pi + dRes
   End If
   
   fdAddA = dRes
End Function
