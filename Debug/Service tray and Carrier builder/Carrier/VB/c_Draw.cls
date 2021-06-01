VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Using: mMth_Arr.bas

Public Sub uCurveO(dX() As Double, dY() As Double, _
                   pic As PictureBox, lColour As Long)
   'Rem: dX() and dY() is zero based arrays
   
   Dim iIU As Integer, iI As Integer
   iIU = flUBoundD(dX())
   
   For iI = 0 To iIU - 1
      pic.Line (dX(iI), dY(iI))- _
               (dX(iI + 1), dY(iI + 1)), lColour
   Next iI
End Sub


Public Sub uCurveC(dX() As Double, dY() As Double, _
                    pic As PictureBox, lColour As Long)
   'Rem: dX() and dY() is zero based arrays
   
   Dim iIU As Integer, iI As Integer
   iIU = flUBoundD(dX())
   
   For iI = 0 To iIU - 1
      pic.Line (dX(iI), dY(iI))- _
               (dX(iI + 1), dY(iI + 1)), lColour
   Next iI
   
   If iIU > 0 Then
      pic.Line (dX(iIU), dY(iIU))- _
               (dX(0), dY(0)), lColour
   End If
End Sub
