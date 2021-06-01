VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTraySectors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type udtPlace
   dX As Double
   dY As Double
   dA As Double
   dARad As Double
End Type

Private Type udtCircleP
   iNPoss As Integer
   iNPossSec As Integer
   dR As Double
   dAPlace0 As Double
   dAStep As Double
   dARot As Double
   utPlace(150) As udtPlace '/// was utPlace(60) As udtPlace
End Type

Private Type udtTrayP
   utCirc(5) As udtCircleP
   iNCircles As Integer
   iNPos As Integer
End Type

Private Type udtCircleI
   iNPossWithIns As Integer
   bInsert() As Boolean
End Type

Private Type udtTrayI
   utCirc(5) As udtCircleI
   iNCirclesWithIns As Integer
   iNPosWithIns As Integer
   iNPosWithInsRef As Integer
End Type


Const _
mc_dDOuter As Double = 245#, mc_dDInner As Double = 75#, _
mc_dASecBeg As Double = 5.5


Private _
   m_utTrayP As udtTrayP, m_utTrayI() As udtTrayI

Private _
   m_iIUPos As Integer, m_sInsertsPos As String, _
   m_iIUCircles As Integer, m_iNInsInOrder As Integer, _
   m_iNTraysInOrder As Integer, m_dPicScale As Double, _
   m_dXCen As Double, m_dYCen As Double
Private _
   m_oFigure As Object
Private _
   m_pic As PictureBox
Private _
sInsertType As String
   


'********************************************************************
'***   PROPERTIES   *************************************************


Public Property Get p_iNTraysInOrder() As Integer

   p_iNTraysInOrder = m_iNTraysInOrder
End Property

'/// added to access circle diameter
Public Property Get d_Diameter() As String
Dim dSearchDiameter As Double
   Select Case sInsertType
   Case "Circle"
        dSearchDiameter = m_oFigure.p_dD
   Case "Triangle"
        dSearchDiameter = m_oFigure.p_dH
   Case "Rectangle"
        dSearchDiameter = m_oFigure.p_dW
   Case "Rhombus"
        dSearchDiameter = m_oFigure.p_dDiagL
   End Select



    d_Diameter = dSearchDiameter / m_dPicScale
'   d_Diameter = m_oFigure.p_dD / m_dPicScale
End Property




'********************************************************************
'***   METHODS   ****************************************************

Public Sub uPictureSet(pic As PictureBox)

   Set m_pic = pic
   m_dPicScale = m_pic.Height / mc_dDOuter * 0.98
End Sub


Public Function fbCalc(sTrayName As String, _
                          iInsInOrder As Integer) As Boolean

'   On Error GoTo ErrHandler
   m_iNInsInOrder = iInsInOrder
   
   If Not fbGetTrayData(sTrayName) Then Exit Function
      
   uCalcTrayPlaces
   uInsertsSpread
'   m_sInsertsPos = fsInsertrsPos()
   
   fbCalc = True
   Exit Function
ErrHandler:
   g_uMsgE "clsTrayCircles.fbCalc"
End Function


Public Function fsInsertrsPos() As String
   Dim _
   sTemp As String, _
   iIT As Integer, iIC As Integer, iIP As Integer, _
   iI As Integer, iNInsertsInOrder As Integer
   
   iNInsertsInOrder = 0
   
   sTemp = "Number of trays: " & m_iNTraysInOrder
   
   For iIT = 0 To m_iNTraysInOrder - 1
      
      sTemp = sTemp & vbCrLf & vbCrLf & "Tray " & iIT + 1 & vbCrLf
      
      sTemp = sTemp & Format("No", "!@@@@") & _
                      Format("Cir", "!@@@@") & _
                      Format("Pos", "!@@@@@@@@") & _
                      Format("X", "!@@@@@@@@@") & _
                      Format("Y", "!@@@@@@") & _
                      Format("Angle", "!@@@@@@")
      iI = 0
      
      For iIC = 0 To m_utTrayP.iNCircles - 1
         For iIP = 0 To m_utTrayP.utCirc(iIC).iNPoss - 1
            If m_utTrayI(iIT).utCirc(iIC).bInsert(iIP) Then
               With m_utTrayP.utCirc(iIC).utPlace(iIP)
               sTemp = sTemp & vbCrLf & _
                        Format(iI + 1, "!@@@@@") & _
                        Format(iIC + 1, "!@@@@") & _
                        Format(iIP + 1, "!@@@@@") & _
                        Format(.dX, "!@@@@@@@@@") & _
                        Format(.dY, "!@@@@@@@@@") & _
                        Format(.dA, "!@@@@@")
               End With
               iI = iI + 1
               iNInsertsInOrder = iNInsertsInOrder + 1
            End If
         Next iIP
      Next iIC
   Next iIT
   
   sTemp = iNInsertsInOrder & vbCrLf & vbCrLf & sTemp
   sTemp = "Number of robot passes: " & sTemp
   fsInsertrsPos = sTemp
End Function


Public Sub uDraw(iITray As Integer)
   Dim _
   iIC As Integer, iIP As Integer
   
   With m_pic
   .Cls
   m_dXCen = .Width / 2 - 1.5
   m_dYCen = .Height / 2 - 1.5
   End With
   
   uDrawTrayContour
   
   'm_pic.FillStyle = 1
   
   For iIC = 0 To m_utTrayP.iNCircles - 1
      For iIP = 0 To m_utTrayP.utCirc(iIC).iNPoss - 1
         If m_utTrayI(iITray).utCirc(iIC).bInsert(iIP) Then
            uDrawPlace iIC, iIP, True
         Else
            uDrawPlace iIC, iIP, False
         End If
      Next iIP
   Next iIC
   
   'Rem: refresh after single figure drawing was removed
   '     because of picture blinking.
   m_pic.Refresh
End Sub


Public Function fbApplyEdit(iICurTray As Integer) As Boolean

   
   If m_utTrayI(iICurTray).iNPosWithIns <> _
      m_utTrayI(iICurTray).iNPosWithInsRef Then
      g_uMsgI _
         "Nuber of inserts " & fiInsertsInOrder() & _
         " is not equal to requested number " & m_iNInsInOrder & "."
      Exit Function
   End If

   Dim iIT As Integer
   For iIT = 0 To m_iNTraysInOrder - 1
      With m_utTrayI(iIT)
      If .iNPosWithIns = m_utTrayI(iICurTray).iNPosWithIns Then _
         m_utTrayI(iIT) = m_utTrayI(iICurTray)
      End With
   Next iIT
   
   fbApplyEdit = True
End Function


'********************************************************************
'***   PRIVATE PROCEDURES   *****************************************

Private Function fbGetTrayData(sTrayName As String) As Boolean

   Dim _
   rst As ADODB.Recordset, iIC As Integer
   
'   On Error GoTo ErrHandler
   
   Set rst = g_oTraysDB.foGetTrayTbl(sTrayName)
   
   m_utTrayP.iNCircles = rst.RecordCount
   
   rst.MoveFirst
   For iIC = 0 To m_utTrayP.iNCircles - 1
       With m_utTrayP.utCirc(iIC)
       .iNPossSec = rst.Fields("NumberOfInserts")
      .iNPoss = .iNPossSec * 3
'''      .dR = rst.Fields("CircleRadius") ' commented - there is no such property as CircleRadius
      .dR = rst.Fields("CircleDiameter")
      .dAPlace0 = rst.Fields("FirstPlaceAngle")
      .dARot = rst.Fields("RotShiftAngle")
'      .dAStep = rst.Fields("StepAngle")
      m_utTrayP.iNPos = m_utTrayP.iNPos + .iNPoss
      End With
      rst.MoveNext
   Next iIC
   
'   Dim sInsertType As String '///commented and changed to public to enable type enquary from property
   
   sInsertType = g_oTraysDB.fsGetInsertType(sTrayName)
   Set rst = g_oTraysDB.foGetInsertTbl(sTrayName)
   
   Select Case sInsertType
   
   Case "Circle"
      Dim _
      oCircle As New clsFigCircle
      oCircle.p_dD = rst.Fields("Diameter") * m_dPicScale
      Set m_oFigure = oCircle
      
   Case "Triangle"
      Dim _
      oTriangle As New clsFigTriangle
      oTriangle.p_dH = rst.Fields("Height") * m_dPicScale
      Set m_oFigure = oTriangle
   
   Case "Rectangle"
      Dim _
      oRectangle As New clsFigRectangle
      oRectangle.p_dH = rst.Fields("ShortEdge") * m_dPicScale
      oRectangle.p_dW = rst.Fields("LongEdge") * m_dPicScale
      Set m_oFigure = oRectangle
   
   Case "Rhombus"
      Dim _
      oRhombus As New clsFigRhombus
      oRhombus.p_dDiagL = rst.Fields("LongDiagonal") * m_dPicScale
      oRhombus.p_dDiagS = rst.Fields("ShortDiagonal") * m_dPicScale
      Set m_oFigure = oRhombus
   End Select
   
   fbGetTrayData = True
   Exit Function
ErrHandler:
   g_uMsgE "clsTrayCircles.fbGetTrayData"
End Function


Private Sub uCalcTrayPlaces()

   Dim _
   iIC As Integer, iIS As Integer, _
   iIPSec As Integer, iIP As Integer, _
   dA As Double, dASec As Double, _
   dARad As Double, dARotRad As Double
   
   For iIC = 0 To m_utTrayP.iNCircles - 1 'Circles count
      
      With m_utTrayP.utCirc(iIC)
      
      dARotRad = .dARot * gc_dToRad
      For iIS = 0 To 2
         dASec = iIS * 120 + mc_dASecBeg
         For iIPSec = 0 To .iNPossSec - 1
            
            'dA = 360# / .iNPoss * iIP
            dA = dASec + .dAPlace0 + iIPSec * .dAStep
            dARad = dA * gc_dToRad
            iIP = iIPSec + iIS * .iNPossSec
            .utPlace(iIP).dX = Round(.dR * Cos(dARad), 2)
            .utPlace(iIP).dY = Round(.dR * Sin(dARad), 2)
            .utPlace(iIP).dA = Round(dA + .dARot, 1)
            .utPlace(iIP).dARad = dARad + dARotRad
         Next iIPSec
      Next iIS
      End With
   Next iIC
End Sub


Private Sub uInsertsSpread()
   Dim _
   iNPosInProcess As Integer, iMod As Integer, _
   iIT As Integer, iIC As Integer, iNInsInTrayMin As Integer
   
   iNPosInProcess = 6 * m_utTrayP.iNPos
   m_iNTraysInOrder = -Int(-m_iNInsInOrder / iNPosInProcess) * 6
   
   ReDim m_utTrayI(m_iNTraysInOrder - 1)
   
   iMod = m_iNInsInOrder Mod m_iNTraysInOrder
   iNInsInTrayMin = Fix(m_iNInsInOrder / m_iNTraysInOrder)
   
   For iIT = 0 To iMod - 1
      m_utTrayI(iIT).iNPosWithIns = iNInsInTrayMin + 1
   Next iIT
   For iIT = iMod To m_iNTraysInOrder - 1
      m_utTrayI(iIT).iNPosWithIns = iNInsInTrayMin
   Next iIT

   Dim iNPosCur As Integer
   
   For iIT = 0 To m_iNTraysInOrder - 1
   
      With m_utTrayI(iIT)
      .iNPosWithInsRef = .iNPosWithIns
      End With
      
      iNPosCur = 0
      For iIC = 0 To m_utTrayP.iNCircles - 1
         iNPosCur = iNPosCur + m_utTrayP.utCirc(iIC).iNPoss
         If iNPosCur < m_utTrayI(iIT).iNPosWithIns Then
            m_utTrayI(iIT).utCirc(iIC).iNPossWithIns = _
                                 m_utTrayP.utCirc(iIC).iNPoss
         Else
            m_utTrayI(iIT).utCirc(iIC).iNPossWithIns = _
                      m_utTrayI(iIT).iNPosWithIns - _
                     (iNPosCur - m_utTrayP.utCirc(iIC).iNPoss)
            m_utTrayI(iIT).iNCirclesWithIns = iIC + 1
            Exit For
         End If
      Next iIC
   Next iIT

   Dim _
   oInsSpreadSec As New clsInsertsSpreadSector

   For iIT = 0 To m_iNTraysInOrder - 1
      For iIC = 0 To m_utTrayP.iNCircles - 1
         With m_utTrayI(iIT).utCirc(iIC)
         
         If .iNPossWithIns <> 0 Then
            .bInsert = _
               oInsSpreadSec.fbIns( _
                           m_utTrayP.utCirc(iIC).iNPoss, _
                           m_utTrayP.utCirc(iIC).iNPossSec, _
                           .iNPossWithIns)
         Else
            ReDim .bInsert(m_utTrayP.utCirc(iIC).iNPoss - 1)
         End If
         End With
      Next iIC
   Next iIT
End Sub


Private Sub uDrawPlace(iIC As Integer, iIP As Integer, _
                       bFill As Boolean)
   With m_oFigure
   .p_dARotRad = m_utTrayP.utCirc(iIC).utPlace(iIP).dARad
   'Rem: Relative to the drawings X go from tray center
   'to up, Y from tray center to left so next statements
   'are built accordingly
   .p_dXCen = m_dXCen - _
         m_utTrayP.utCirc(iIC).utPlace(iIP).dY * m_dPicScale
   .p_dYCen = m_dYCen - _
         m_utTrayP.utCirc(iIC).utPlace(iIP).dX * m_dPicScale
   If bFill Then
      .uDraw m_pic, &HD00000, RGB(255, 255, 180) '&HE0E0E0
   Else
      'Rem: m_pic.BackColor if system give black filling
      '     if using API
      .uDraw m_pic, &HD00000, m_pic.BackColor
   End If
   End With
End Sub


Public Sub uEdit(iIT As Integer, snX As Single, snY As Single)
   Dim _
   iIC As Integer, iIP As Integer
   
   uGetPosInd snX, snY, iIC, iIP
   
   With m_utTrayI(iIT).utCirc(iIC)
   .bInsert(iIP) = Not .bInsert(iIP)
   If .bInsert(iIP) Then
      .iNPossWithIns = .iNPossWithIns + 1
      m_utTrayI(iIT).iNPosWithIns = m_utTrayI(iIT).iNPosWithIns + 1
      uDrawPlace iIC, iIP, True
   Else
      .iNPossWithIns = .iNPossWithIns - 1
      m_utTrayI(iIT).iNPosWithIns = m_utTrayI(iIT).iNPosWithIns - 1
      uDrawPlace iIC, iIP, False
   End If
   End With
   
   'Rem: refresh after single figure drawing was removed
   '     because of picture blinking.
   m_pic.Refresh
End Sub


Private Sub uGetPosInd(snX As Single, snY As Single, _
                       iICOut As Integer, iIPOut As Integer)
   Dim _
   dXP As Double, dYP As Double, dDlARad As Double, _
   dARad As Double, dR As Double, dDlR As Double
      
   'Rem: Relative to the drawings X go from tray center
   'to up, Y from tray center to left so next statements
   'are built accordingly
   dYP = (m_dXCen - snX) / m_dPicScale
   dXP = (m_dYCen - snY) / m_dPicScale
   dARad = fdAtn2(dXP, dYP)
   dR = fdHypotenuse(dXP, dYP)
   
   With m_utTrayP
   If .iNCircles = 1 Then
      iICOut = 0
   Else
      dDlR = (.utCirc(0).dR - .utCirc(1).dR) / 2
      Dim iIC As Integer
      For iIC = 0 To .iNCircles - 1
         If iIC = 0 Then
            If dR > .utCirc(iIC).dR - dDlR Then
               iICOut = iIC
               Exit For
            End If
         ElseIf iIC = .iNCircles - 1 Then
            If dR < .utCirc(iIC).dR + dDlR Then
               iICOut = iIC
               Exit For
            End If
         Else
            If dR < .utCirc(iIC).dR + dDlR And _
               dR > .utCirc(iIC).dR - dDlR Then
               iICOut = iIC
               Exit For
            End If
         End If
      Next iIC
   End If
   
   Dim _
   iIS As Integer, iIPInSec As Integer, dASubst As Double, _
   dAStepRad As Double, dAInSecRad As Double

   iIS = Fix(dARad / gc_d2Pi3)
   dAStepRad = .utCirc(iIC).dAStep * gc_dToRad
   dAInSecRad = dARad - iIS * gc_d2Pi3 - _
              (.utCirc(iIC).dAPlace0 + mc_dASecBeg) * gc_dToRad
   
   iIPInSec = Fix(dAInSecRad / dAStepRad + 0.5)
   If iIPInSec > .utCirc(iIC).iNPossSec - 1 Then
      iIPOut = (iIS + 1) * .utCirc(iIC).iNPossSec - 1
   Else
      iIPOut = iIS * .utCirc(iIC).iNPossSec + _
               Fix(dAInSecRad / dAStepRad + 0.5)
   End If
   End With
End Sub


Private Function fiInsertsInOrder() As Integer

   Dim iIT As Integer
   
   For iIT = 0 To m_iNTraysInOrder - 1
      fiInsertsInOrder = fiInsertsInOrder + _
                         m_utTrayI(iIT).iNPosWithIns
   Next iIT
End Function


Private Sub uDrawTrayContour()

   Dim _
   dRad As Double, iICon As Integer
   
   m_pic.Cls
   
   'Rem: Line and Circle methods can't be used with With...End With
   m_pic.DrawStyle = 1
   m_pic.Line (m_dXCen, m_dYCen)-(m_dXCen, 0), &HFF00FF
   m_pic.Line (m_dXCen, m_dYCen)-(m_dXCen, m_pic.Height), &H17000
   m_pic.Line (0, m_dYCen)-(m_pic.Width, m_dYCen), &H17000
   m_pic.DrawStyle = 0
   dRad = mc_dDOuter / 2 * m_dPicScale
   m_pic.Circle (m_dXCen, m_dYCen), dRad, &HD00000
   dRad = mc_dDInner / 2 * m_dPicScale
   m_pic.Circle (m_dXCen, m_dYCen), dRad, &HD00000
   
   For iICon = 0 To 2
      uDrawSector 120# * iICon * gc_dToRad
   Next iICon
End Sub


Private Sub uDrawSector(dABegRad As Double)

   Const _
   c_dROut As Double = 107.5, _
   c_dRIn As Double = 47.5, _
   c_dRCorn As Double = 5#, _
   c_dDlA1 As Double = 2.8, _
   c_dDlA2 As Double = 5.45, _
   c_dRL1 As Double = 102.23, _
   c_dRL2 As Double = 52.26 - 0.5
   
   Dim _
   dRad As Double, dAStart As Double, dAEnd As Double, _
   dALn As Double, dACen As Double, _
   dX1 As Double, dY1 As Double, dX2 As Double, dY2 As Double, _
   snX1 As Single, snY1 As Single, snX2 As Single, snY2 As Single
      
     
   dRad = c_dROut * m_dPicScale
   dAStart = (gc_dA90 + mc_dASecBeg + c_dDlA1) * gc_dToRad
   dAEnd = (gc_dA90 + gc_dA120 - mc_dASecBeg - c_dDlA1) * gc_dToRad
   dAStart = fdSub2Pi(dAStart + dABegRad)
   dAEnd = fdSub2Pi(dAEnd + dABegRad)
   m_pic.Circle (m_dXCen, m_dYCen), dRad, &HD00000, dAStart, dAEnd
   
   dRad = c_dRIn * m_dPicScale
   dAStart = (gc_dA90 + mc_dASecBeg + c_dDlA2) * gc_dToRad
   dAEnd = (gc_dA90 + gc_dA120 - mc_dASecBeg - c_dDlA2) * gc_dToRad
   dAStart = fdSub2Pi(dAStart + dABegRad)
   dAEnd = fdSub2Pi(dAEnd + dABegRad)
   m_pic.Circle (m_dXCen, m_dYCen), dRad, &HD00000, dAStart, dAEnd
   
   '  Line 1
   dRad = c_dRL1 * m_dPicScale
   dALn = (gc_dA90 + mc_dASecBeg) * gc_dToRad
   dALn = fdSub2Pi(dALn + dABegRad)
   uRotatePtAboutAxis dRad, 0, dALn, dX1, dY1
   dRad = c_dRL2 * m_dPicScale
   uRotatePtAboutAxis dRad, 0, dALn, dX2, dY2
   m_pic.Line (m_dXCen + dX1, m_dYCen - dY1)- _
              (m_dXCen + dX2, m_dYCen - dY2), &HD00000
   
   '     Fillet 1
   dRad = (c_dROut - c_dRCorn) * m_dPicScale
   dACen = (gc_dA90 + mc_dASecBeg + c_dDlA1) * gc_dToRad
   dACen = fdSub2Pi(dACen + dABegRad)
   uRotatePtAboutAxis dRad, 0, dACen, dX1, dY1
   dRad = c_dRCorn * m_dPicScale
   dAStart = mc_dASecBeg * gc_dToRad
   dAEnd = dAStart + (gc_dA90 + c_dDlA1) * gc_dToRad
   dAStart = fdSub2Pi(dAStart + dABegRad)
   dAEnd = fdSub2Pi(dAEnd + dABegRad)
   m_pic.Circle (m_dXCen + dX1, m_dYCen - dY1), _
                 dRad, &HD00000, dAStart, dAEnd
                 
   '     Fillet 2
   dRad = (c_dRIn + c_dRCorn) * m_dPicScale
   dACen = (gc_dA90 + mc_dASecBeg + c_dDlA2) * gc_dToRad
   dACen = fdSub2Pi(dACen + dABegRad)
   uRotatePtAboutAxis dRad, 0, dACen, dX1, dY1
   dRad = c_dRCorn * m_dPicScale
   dAStart = (gc_dA270 + mc_dASecBeg + c_dDlA2) * gc_dToRad
   dAEnd = dAStart + (gc_dA90 - c_dDlA2) * gc_dToRad
   dAStart = fdSub2Pi(dAStart + dABegRad)
   dAEnd = fdSub2Pi(dAEnd + dABegRad)
   m_pic.Circle (m_dXCen + dX1, m_dYCen - dY1), _
                 dRad, &HD00000, dAStart, dAEnd
   
   '  Line 2
   dRad = c_dRL1 * m_dPicScale
   dALn = (gc_dA90 + gc_dA120 - mc_dASecBeg) * gc_dToRad
   dALn = fdSub2Pi(dALn + dABegRad)
   uRotatePtAboutAxis dRad, 0, dALn, dX1, dY1
   dRad = c_dRL2 * m_dPicScale
   uRotatePtAboutAxis dRad, 0, dALn, dX2, dY2
   m_pic.Line (m_dXCen + dX1, m_dYCen - dY1)- _
              (m_dXCen + dX2, m_dYCen - dY2), &HD00000
   
   '     Fillet 1
   dRad = (c_dROut - c_dRCorn) * m_dPicScale
   dACen = (gc_dA90 + gc_dA120 - mc_dASecBeg - c_dDlA1) * gc_dToRad
   dACen = fdSub2Pi(dACen + dABegRad)
   uRotatePtAboutAxis dRad, 0, dACen, dX1, dY1
   dRad = c_dRCorn * m_dPicScale
   dAStart = (gc_dA90 + gc_dA120 - mc_dASecBeg - c_dDlA1) * gc_dToRad
   dAEnd = dAStart + (gc_dA90 + c_dDlA1) * gc_dToRad
   dAStart = fdSub2Pi(dAStart + dABegRad)
   dAEnd = fdSub2Pi(dAEnd + dABegRad)
   m_pic.Circle (m_dXCen + dX1, m_dYCen - dY1), _
                 dRad, &HD00000, dAStart, dAEnd
                 
   '     Fillet 2
   dRad = (c_dRIn + c_dRCorn) * m_dPicScale
   dACen = (gc_dA90 + gc_dA120 - mc_dASecBeg - c_dDlA2) * gc_dToRad
   dACen = fdSub2Pi(dACen + dABegRad)
   uRotatePtAboutAxis dRad, 0, dACen, dX1, dY1
   dRad = c_dRCorn * m_dPicScale
   dAStart = (gc_dA270 + gc_dA30 - mc_dASecBeg) * gc_dToRad
   dAEnd = dAStart + (gc_dA90 - c_dDlA2) * gc_dToRad
   dAStart = fdSub2Pi(dAStart + dABegRad)
   dAEnd = fdSub2Pi(dAEnd + dABegRad)
   m_pic.Circle (m_dXCen + dX1, m_dYCen - dY1), _
                 dRad, &HD00000, dAStart, dAEnd
End Sub

