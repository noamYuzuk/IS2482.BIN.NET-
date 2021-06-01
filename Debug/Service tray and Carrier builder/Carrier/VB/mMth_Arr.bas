Attribute VB_Name = "m_Arr"
Option Explicit

Public Function fdMax(dX() As Double, Optional iUp As Integer, _
                      Optional iLow As Integer) As Double
'Find Max element of array
   Dim iLower As Integer
   Dim iUpper As Integer
   Dim iI As Integer
   
   iLower = LBound(dX)
   iUpper = UBound(dX)
   If iUp <> 0 Then _
      If iUpper > iUp Then iUpper = iUp
   If iLow <> 0 Then _
      If iLower < iLow Then iLower = iLow
   
   fdMax = dX(iLower)
   For iI = iLower + 1 To iUpper
      If fdMax < dX(iI) Then fdMax = dX(iI)
   Next iI
End Function


Public Function fdMin(dX() As Double, Optional iUp As Integer, _
                      Optional iLow As Integer) As Double
'Find Min element of array
   Dim iLower As Integer
   Dim iUpper As Integer
   Dim iI As Integer
   
   iLower = LBound(dX)
   iUpper = UBound(dX)
   If iUp <> 0 Then _
      If iUpper > iUp Then iUpper = iUp
   If iLow <> 0 Then _
      If iLower < iLow Then iLower = iLow
   
   fdMin = dX(iLower)
   For iI = iLower + 1 To iUpper
      If fdMin > dX(iI) Then fdMin = dX(iI)
   Next iI
End Function


Public Function fdMax2(dX() As Double, iDim As Integer, _
                       iInd As Integer, Optional iUpper As Integer, _
                       Optional iLower As Integer) As Double
'Find Max element of array
   Dim iLower1 As Integer, iUpper1 As Integer, _
       iLower2 As Integer, iUpper2 As Integer
   Dim iI As Integer, iJ As Integer
   Dim dRes As Double
   
   If iDim = 1 Then
      iLower2 = LBound(dX, 2):   iUpper2 = UBound(dX, 2)
      If iUpper <> 0 Then _
         If iUpper2 > iUpper Then iUpper2 = iUpper
      If iLower <> 0 Then _
         If iLower2 < iLower Then iLower2 = iLower
      
      dRes = dX(iInd, iLower2)
      For iI = iLower2 + 1 To iUpper2
         If dRes < dX(iInd, iI) Then dRes = dX(iInd, iI)
      Next iI
   
   ElseIf iDim = 2 Then
      iLower1 = LBound(dX, 1):   iUpper1 = UBound(dX, 1)
      If iUpper <> 0 Then _
         If iUpper1 > iUpper Then iUpper1 = iUpper
      If iLower <> 0 Then _
         If iLower1 < iLower Then iLower1 = iLower
      
      dRes = dX(iLower1, iInd)
      For iI = iLower1 + 1 To iUpper1
         If dRes < dX(iI, iInd) Then dRes = dX(iI, iInd)
      Next iI
   End If
   fdMax2 = dRes
End Function


Public Function fdMin2(dX() As Double, iDim As Integer, _
                       iInd As Integer, Optional iUpper As Integer, _
                       Optional iLower As Integer) As Double
'Find Max element of array
   Dim iLower1 As Integer, iUpper1 As Integer, _
       iLower2 As Integer, iUpper2 As Integer
   Dim iI As Integer, iJ As Integer
   Dim dRes As Double
   
   If iDim = 1 Then
      iLower2 = LBound(dX, 2):   iUpper2 = UBound(dX, 2)
      If iUpper <> 0 Then _
         If iUpper2 > iUpper Then iUpper2 = iUpper
      If iLower <> 0 Then _
         If iLower2 < iLower Then iLower2 = iLower
      
      dRes = dX(iInd, iLower2)
      For iI = iLower2 + 1 To iUpper2
         If dRes > dX(iInd, iI) Then dRes = dX(iInd, iI)
      Next iI
   
   ElseIf iDim = 2 Then
      iLower1 = LBound(dX, 1):   iUpper1 = UBound(dX, 1)
      If iUpper <> 0 Then _
         If iUpper1 > iUpper Then iUpper1 = iUpper
      If iLower <> 0 Then _
         If iLower1 < iLower Then iLower1 = iLower
      
      dRes = dX(iLower1, iInd)
      For iI = iLower1 + 1 To iUpper1
         If dRes > dX(iI, iInd) Then dRes = dX(iI, iInd)
      Next iI
   End If
   fdMin2 = dRes
End Function


Public Function fdMax2Tot( _
                           dX() As Double) As Double
'Find Max element of array
   Dim _
   lIL1 As Long, lIU1 As Long, _
   lIL2 As Long, lIU2 As Long, _
   lI As Long, lJ As Long, dRes As Double
'dd
   lIL1 = LBound(dX(), 1):    lIU1 = UBound(dX(), 1)
   lIL2 = LBound(dX(), 2):    lIU2 = UBound(dX(), 2)
   dRes = dX(lIL1, lIL2)
   
   For lJ = lIL1 To lIU1
      For lI = lIL2 To lIU2
         If dRes < dX(lJ, lI) Then dRes = dX(lJ, lI)
      Next lI
   Next lJ
   
   fdMax2Tot = dRes
End Function


Public Function fdMin2Tot( _
                           dX() As Double) As Double
'Find Max element of array
   Dim _
   lIL1 As Long, lIU1 As Long, _
   lIL2 As Long, lIU2 As Long, _
   lI As Long, lJ As Long, dRes As Double
'dd
   lIL1 = LBound(dX(), 1):    lIU1 = UBound(dX(), 1)
   lIL2 = LBound(dX(), 2):    lIU2 = UBound(dX(), 2)
   
   dRes = dX(lIL1, lIL2)
   For lJ = lIL1 To lIU1
      For lI = lIL2 To lIU2
         If dRes > dX(lJ, lI) Then dRes = dX(lJ, lI)
      Next lI
   Next lJ
   
   fdMin2Tot = dRes
End Function


Public Function fvMax(vX As Variant) As Variant
'Find Max element of array
   Dim iLower As Integer
   Dim iUpper As Integer
   Dim iI As Integer
   
   iLower = LBound(vX)
   iUpper = UBound(vX)
   fvMax = vX(iLower)
   
   For iI = iLower + 1 To iUpper
      If fvMax < vX(iI) Then fvMax = vX(iI)
   Next iI
End Function


Public Function fvMin(vX As Variant) As Variant
'Find Min element of array
   Dim iLower As Integer
   Dim iUpper As Integer
   Dim iI As Integer
   
   iLower = LBound(vX)
   iUpper = UBound(vX)
   fvMin = vX(iLower)
   
   For iI = iLower + 1 To iUpper
      If fvMin > vX(iI) Then fvMin = vX(iI)
   Next iI
End Function

'*************************************************************************


Public Sub uArr3CopySn( _
                        snArr() As Single, iDim As Integer, _
                        iIFrom As Integer, iITo As Integer, _
                        iILTar As Integer, iIUTar As Integer, _
                        snArrO() As Single)
   Dim _
   iL1 As Integer, iU1 As Integer, _
   iL2 As Integer, iU2 As Integer, _
   iI As Integer, iJ As Integer, iK As Integer, _
   iITar As Integer
   
   iITar = iILTar - 1
   
   Select Case iDim
   Case 1
      iL1 = LBound(snArr, 2):     iU1 = UBound(snArr, 2)
      iL2 = LBound(snArr, 3):     iU2 = UBound(snArr, 3)
      ReDim snArrO(iILTar To iIUTar, iL1 To iU1, iL2 To iU2)
      For iI = iIFrom To iITo
         iITar = iITar + 1
         For iJ = iL1 To iU1
            For iK = iL2 To iU2
               snArrO(iITar, iJ, iK) = snArr(iI, iJ, iK)
            Next iK
         Next iJ
      Next iI
   Case 2
   Case 3
   Case Else
      MsgBox "Wrong iDim = " & iDim, vbCritical, "uArr3CopySn"
      Stop
   End Select
End Sub


Public Sub uArr3CopyB( _
                        bArr() As Boolean, iDim As Integer, _
                        iIFrom As Integer, iITo As Integer, _
                        iILTar As Integer, iIUTar As Integer, _
                        bArrO() As Boolean)
   Dim _
   iL1 As Integer, iU1 As Integer, _
   iL2 As Integer, iU2 As Integer, _
   iI As Integer, iJ As Integer, iK As Integer, _
   iITar As Integer
   
   iITar = iILTar - 1
   
   Select Case iDim
   Case 1
      iL1 = LBound(bArr, 2):     iU1 = UBound(bArr, 2)
      iL2 = LBound(bArr, 3):     iU2 = UBound(bArr, 3)
      ReDim bArrO(iILTar To iIUTar, iL1 To iU1, iL2 To iU2)
      For iI = iIFrom To iITo
         iITar = iITar + 1
         For iJ = iL1 To iU1
            For iK = iL2 To iU2
               bArrO(iITar, iJ, iK) = bArr(iI, iJ, iK)
            Next iK
         Next iJ
      Next iI
   Case 2
   Case 3
   Case Else
      MsgBox "Wrong iDim = " & iDim, vbCritical, "uArr3CopyB"
      Stop
   End Select
End Sub


Public Sub uArrCopyToSegB( _
                     bArrFrom() As Boolean, bArrTo() As Boolean, _
                     iIBeg As Integer)
   Dim _
   iIU As Integer, iI As Integer
   
   iIU = flUBoundB(bArrFrom())
   
   For iI = 0 To iIU
      bArrTo(iIBeg + iI) = bArrFrom(iI)
   Next iI
End Sub


Public Function flFindInArS( _
                              sArr() As String, sMemb As String, _
                              Optional lIUO As Long, _
                              Optional lILO As Long) As Long
   Dim _
   lI As Long, lRes As Long
'dd
   On Error Resume Next
   lIUO = UBound(sArr)
   If Err.Number <> 0 Then
      flFindInArS = -1
      Exit Function
   End If
   
   lILO = LBound(sArr)
   
   For lI = lILO To lIUO
      If sArr(lI) = sMemb Then Exit For
   Next lI
   If lI > lIUO Then
      lRes = -1
   Else
      lRes = lI
   End If
   
   flFindInArS = lRes
End Function


Public Function flFindInAr2S( _
                              sArr() As String, sMemb As String, _
                              lDimChanging As Long, lIndFixed As Long, _
                              Optional lIFrom As Long, _
                              Optional lITo As Long) As Long
   Dim _
   lIL As Long, lIU As Long, lIL2 As Long, lIU2 As Long, _
   lI As Long, lRes As Long
'dd
   lIL = LBound(sArr(), lDimChanging):   lIU = UBound(sArr(), lDimChanging)
   If lIFrom <> 0 Then _
      If lIL < lIFrom Then lIL = lIFrom
   If lITo <> 0 Then _
      If lIU > lITo Then lIU = lITo

   lRes = -1
   If lDimChanging = 2 Then
'      lIL2 = LBound(sArr, 2):   lIU2 = UBound(sArr, 2)
'      If lIU <> 0 Then _
'         If lIU2 > lIU Then lIU2 = lIU
'      If lIL <> 0 Then _
'         If lIL2 < lIL Then lIL2 = lIL
      
      For lI = lIL To lIU
         If sArr(lIndFixed, lI) = sMemb Then
            lRes = lI:        Exit For
         End If
      Next lI
   
   ElseIf lDimChanging = 1 Then
'      lIL1 = LBound(sArr, 1):   lIU1 = UBound(sArr, 1)
'      If lIU <> 0 Then _
'         If lIU1 > lIU Then lIU1 = lIU
'      If lIL <> 0 Then _
'         If lIL1 < lIL Then lIL1 = lIL
      
      For lI = lIL To lIU
         If sArr(lI, lIndFixed) = sMemb Then
            lRes = lI:        Exit For
         End If
      Next lI
   End If
   
   flFindInAr2S = lRes
End Function


Public Sub uArrayShift(dArray() As Double, iIBeg As Integer)

   Dim iI As Integer
   Dim iIMax As Integer, iIMin As Integer
   Dim vTemp() As Variant

   iIMin = LBound(dArray)
   iIMax = UBound(dArray)
   ReDim vTemp(iIMin To iIMax)

   For iI = iIMin To iIMax - iIBeg
      vTemp(iI) = dArray(iI + iIBeg)
   Next iI
   For iI = iIMax - iIBeg + 1 To iIMax
      vTemp(iI) = dArray(iI - (iIMax - iIBeg + 1))
   Next iI
   For iI = iIMin To iIMax
      dArray(iI) = vTemp(iI)
   Next iI
End Sub


Public Sub uArray2DShift(dArray() As Double, iIBeg As Integer, _
                         iDim As Integer)
   Dim iI As Integer, iK As Integer
   Dim iIMax(1 To 2) As Integer, iIMin(1 To 2) As Integer
   Dim vTemp() As Variant

   For iI = 1 To 2
      iIMin(iI) = LBound(dArray, iI)
      iIMax(iI) = UBound(dArray, iI)
   Next iI
   Select Case iDim
      Case 1
'                        For iK = iIMin(1) To iIMax(1)
'                                For iI = iIMin(2) To iIMax(2) - iIBeg
'                                                       vTemp(iI) = dArray(iK, iI + iIBeg)
'                                Next iI
'                                For iI = iIMax(2) - iIBeg + 1 To iIMax(2)
'                                                       vTemp(iI) = dArray(iK, iI - (iIMax - iIBeg + 1))
'                                Next iI
'                                For iI = iIMin(2) To iIMax(2)
'                                                       dArray(iK, iI) = vTemp(iI)
'                                Next iI
'                        Next iK
   Case 2
      ReDim vTemp(iIMin(2) To iIMax(2))
      For iK = iIMin(1) To iIMax(1)
         For iI = iIMin(2) To iIMax(2) - iIBeg
            vTemp(iI) = dArray(iK, iI + iIBeg)
         Next iI
         For iI = iIMax(2) - iIBeg + 1 To iIMax(2)
            vTemp(iI) = dArray(iK, iI - (iIMax(2) - iIBeg + 1))
         Next iI
         For iI = iIMin(2) To iIMax(2)
            dArray(iK, iI) = vTemp(iI)
         Next iI
      Next iK
   End Select
End Sub


Public Function fsArrAddItem(sArray() As String, _
                             sItem As String) As String()
   Dim _
   iIMax As Integer, iIMin As Integer, _
   sTemp() As String, lErr As Long

   On Error Resume Next 'if sArray is Nothing
   iIMin = LBound(sArray)
   lErr = Err.Number
   On Error GoTo 0
   
   If lErr = 0 Then
      iIMax = UBound(sArray)
      
      sTemp() = sArray()
      ReDim Preserve sTemp(iIMin To iIMax + 1)
   
      sTemp(iIMax + 1) = sItem
   Else
      ReDim sTemp(0)
      sTemp(0) = sItem
   End If
   
   fsArrAddItem = sTemp()
End Function


Public Function fsArrInsertItem(sArray() As String, _
                                sItem As String, _
                                iIIns As Integer) As String()
   Dim _
   iI As Integer, _
   iIMax As Integer, iIMin As Integer, _
   sTemp() As String, lErr As Long

   On Error Resume Next 'if sArray is Nothing
   iIMin = LBound(sArray)
   lErr = Err.Number
   On Error GoTo 0
   
   If lErr = 0 Then
      
      iIMax = UBound(sArray)
      ReDim sTemp(iIMin To iIMax + 1)
   
      For iI = iIMin To iIIns - 1
         sTemp(iI) = sArray(iI)
      Next iI
      
      sTemp(iIIns) = sItem
      
      For iI = iIIns To iIMax
         sTemp(iI + 1) = sArray(iI)
      Next iI
   Else
      ReDim sTemp(0)
      sTemp(0) = sItem
   End If
   
   fsArrInsertItem = sTemp()
End Function


Public Function fsArrDeleteItem(sArray() As String, _
                                iIDel As Integer) As String()
   Dim _
   iI As Integer, _
   iIMax As Integer, iIMin As Integer, _
   sTemp() As String

   On Error Resume Next 'if sArray is Nothing
   iIMin = LBound(sArray)
   If Err.Number <> 0 Then Exit Function
   On Error GoTo 0
   
   iIMin = LBound(sArray)
   iIMax = UBound(sArray)
   If iIMin > iIMax Or iIMin = iIMax Then
      Erase sTemp
      Exit Function
   End If
   ReDim sTemp(iIMin To iIMax - 1)

   For iI = iIMin To iIDel - 1
      sTemp(iI) = sArray(iI)
   Next iI
   
   For iI = iIDel + 1 To iIMax
      sTemp(iI - 1) = sArray(iI)
   Next iI
   
   fsArrDeleteItem = sTemp()
End Function


Public Function flUBoundS(s() As String) As Long

   Dim iIMin As Integer
   
   On Error Resume Next 'if sArray is Nothing
   flUBoundS = UBound(s())
   If Err.Number <> 0 Then flUBoundS = -1
End Function


Public Function flUBoundD(d() As Double) As Long

   Dim iIMin As Integer
   
   On Error Resume Next 'if sArray is Nothing
   flUBoundD = UBound(d())
   If Err.Number <> 0 Then flUBoundD = -1
End Function


Public Function flUBoundB(b() As Boolean) As Long

   Dim iIMin As Integer
   
   On Error Resume Next 'if sArray is Nothing
   flUBoundB = UBound(b())
   If Err.Number <> 0 Then flUBoundB = -1
End Function
