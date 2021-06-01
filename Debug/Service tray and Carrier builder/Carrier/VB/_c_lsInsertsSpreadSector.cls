VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsertsSpreadSector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_iNPos As Integer, m_iNIns As Integer, _
        m_bIns() As Boolean


Public Function fbIns(iNPos As Integer, _
                      iNPosSec As Integer, _
                      iNIns As Integer) As Boolean()
   ReDim m_bIns(iNPos - 1)
   
   Dim iIS As Integer, iNInsSec(2) As Integer
   
   iNInsSec(0) = iNIns / 3
   iNInsSec(1) = iNInsSec(0)
   iNInsSec(2) = iNIns - iNInsSec(0) - iNInsSec(1)
   
   For iIS = 0 To 2
      u1Group iNPosSec, iNInsSec(iIS), m_bIns(), iNPosSec * iIS
   Next iIS
   
   fbIns = m_bIns()
End Function


Private Function u1Group(iNPos As Integer, _
                         iNIns As Integer, _
                         b() As Boolean, _
                         iIBeg As Integer)
   Dim _
   iStep As Integer, i As Integer, _
   iIBeforeLast As Integer, iInterval As Integer
   
   If iNIns = 0 Then Exit Function
   
   iStep = fiStep(iNPos, iNIns)
   
   If iNIns > 1 Then
      For i = 0 To iNIns - 2
         b(iIBeg + i * iStep) = True
      Next i
   
      iIBeforeLast = (iNIns - 2) * iStep
      iInterval = iNPos - 1 - iIBeforeLast
      If iInterval > 1 Then
         b(iIBeg + iIBeforeLast + CInt(iInterval / 1.999)) = True
      Else
         b(iIBeg + iIBeforeLast + 1) = True
      End If
   Else
      b(iIBeg) = True
   End If
End Function


Private Function fiStep(iNPos As Integer, _
                        iNIns As Integer) As Integer
   fiStep = iNPos / iNIns
   If fiStep * (iNIns - 1) > iNPos - 1 Then fiStep = fiStep - 1
End Function
