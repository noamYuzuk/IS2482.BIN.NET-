VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrawGDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Using: mMth_Arr.bas

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, ByRef lplb As LOGBRUSH, ByVal dwStyleCount As Long, ByVal lpStyle As Long) As Long

Private Const PS_GEOMETRIC = &H10000
Private Const PS_SOLID = 0
Private Const PS_ENDCAP_FLAT = &H200
Private Const PS_JOIN_BEVEL = &H1000

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long

Private Const HS_BDIAGONAL = 3               '  /////
Private Const HS_CROSS = 4                   '  +++++
Private Const HS_DIAGCROSS = 5               '  xxxxx
Private Const HS_FDIAGONAL = 2               '  \\\\\
Private Const HS_HORIZONTAL = 0              '  -----
Private Const HS_SOLID = 8
Private Const BS_SOLID = 0

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare Function SelectObject Lib "gdi32" ( _
                                 ByVal hdc As Long, _
                                 ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" ( _
                                 ByVal hObject As Long) As Long

Private Declare Function PolyBezier Lib "gdi32.dll" (ByVal hdc As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
Private Declare Function PolyBezierTo Lib "gdi32.dll" (ByVal hdc As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Private Declare Function PolyPolygon Lib "gdi32.dll" (ByVal hdc As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
Private Declare Function RectangleX Lib "gdi32" Alias "Rectangle" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long


Private Declare Sub CopyMemory Lib "kernel32" _
                                  Alias "RtlMoveMemory" _
                                 (ByRef Destination As Any, _
                                  ByRef Source As Any, _
                                  ByVal Bytes As Long)

Private Declare Function OleTranslateColor Lib "olepro32.dll" _
                                 (ByVal OLE_COLOR As Long, _
                                  ByVal HPALETTE As Long, _
                                  pccolorref As Long) As Long



Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type


Dim blueBrush As Long
Dim htchBrush As Long
Dim activeFlag As Boolean


Private _
m_o As Object, _
m_lHDC As Long, _
m_lPen As Long, m_lPenSv As Long, _
m_lBrush As Long, m_lBrushSv As Long



Private Sub Class_Terminate()

   DeleteObject SelectObject(m_lHDC, m_lPenSv)
   DeleteObject SelectObject(m_lHDC, m_lBrushSv)
End Sub


'********************************************************************
'***   METHODS   ****************************************************


Public Sub uSetObject(oDrawOnIt As Object)
   
   Set m_o = oDrawOnIt
   m_lHDC = oDrawOnIt.hdc
End Sub


Public Sub uSetContour(lWidth As Long, lColor As Long)
   Dim _
   btC() As Byte
   
   'Rem: fbtGetRGBFromOLEColour used because APIs don't work
   '     with System Colors such as &H8000000F& etc.
   btC() = fbtGetRGBFromOLEColour(lColor)
   
   m_lPen = CreatePen(PS_SOLID, lWidth, RGB(btC(0), btC(1), btC(2)))
   'Rem: SelectObject return handle of the object being replaced
   'Rem: deletion of original pen may lead to problems with
   '     subsequent using of standard VB methods.
   If m_lPenSv = 0 Then
      m_lPenSv = SelectObject(m_lHDC, m_lPen)
   Else
      DeleteObject SelectObject(m_lHDC, m_lPen)
   End If
End Sub


Public Sub uSetFillColor(lColor As Long)
   Dim _
   btC() As Byte
   
   'Rem: fbtGetRGBFromOLEColour used because APIs don't work
   '     with System Colors such as &H8000000F& etc.
   btC() = fbtGetRGBFromOLEColour(lColor)
   
   m_lBrush = CreateSolidBrush(RGB(btC(0), btC(1), btC(2)))
   'Rem: SelectObject return handle of the object being replaced
   'Rem: deletion of original pen may lead to problems with
   '     subsequent using of standard VB methods.
   If m_lPenSv = 0 Then
      m_lBrushSv = SelectObject(m_lHDC, m_lBrush)
   Else
      DeleteObject SelectObject(m_lHDC, m_lBrush)
   End If
End Sub


Public Sub uCurveO(dX() As Double, dY() As Double)
   
   'Rem: dX() and dY() is zero based arrays
   Dim _
   lIU As Long, iI As Integer
   
   'Rem: Very strange problem -
   '     after MoveToEx lIU variable is being changed!!!!!!
   'lIU = flUBoundD(dX())
   
   BeginPath m_lHDC
   MoveToEx m_lHDC, dX(0), dY(0), 0
   lIU = flUBoundD(dX())
   For iI = 1 To lIU
      LineTo m_lHDC, dX(iI), dY(iI)
   Next iI
   EndPath m_lHDC
   StrokePath m_lHDC
   'm_o.Refresh
End Sub


Public Sub uCurveC(dX() As Double, dY() As Double)
   
   'Rem: dX() and dY() is zero based arrays
   Dim _
   lIU As Long, iI As Integer
   
   'Rem: Very strange problem -
   '     after MoveToEx lIU variable is being changed!!!!!!
   'lIU = flUBoundD(dX())
   
   BeginPath m_lHDC
   MoveToEx m_lHDC, dX(0), dY(0), 0
   lIU = flUBoundD(dX())
   For iI = 1 To lIU
      LineTo m_lHDC, dX(iI), dY(iI)
   Next iI
   EndPath m_lHDC
   StrokeAndFillPath m_lHDC
   'm_o.Refresh
End Sub


Public Sub uRestoreOrigProp()

End Sub


'********************************************************************
'***   PRIVATE PROCEDURES   *****************************************


Private Sub GetRGBFromOLEColour(ByVal dwOleColour As Long, r As Long, g As Long, b As Long)
    
  'pass a hex colour, return the rgb components
   Dim clrref As Long
   
  'translate OLE color to valid color if passed
   OleTranslateColor dwOleColour, 0, clrref
  
   b = (clrref \ 65536) And &HFF
   g = (clrref \ 256) And &HFF
   r = clrref And &HFF
End Sub


Private Function fbtGetRGBFromOLEColour( _
                           ByVal dwOleColour As Long) As Byte()
    
  'pass a hex colour, return the rgb components
   Dim clrref As Long
   Dim bt(0 To 3) As Byte
   
  'translate OLE color to valid color if passed
   OleTranslateColor dwOleColour, 0, clrref

  'copy the color into a byte array
  'the size of the long (4 bytes)
   CopyMemory bt(0), clrref, 4
   
  'return the byte array. Note there
  'is no value in byte(4).
   fbtGetRGBFromOLEColour = bt()
End Function
