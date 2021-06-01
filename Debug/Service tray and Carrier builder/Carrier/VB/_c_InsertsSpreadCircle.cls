VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsertsSpreadCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_iNPos As Integer, m_iNIns As Integer, _
        m_bIns() As Boolean


Public Function fbIns(iNPos As Integer, _
                      iNIns As Integer) As Boolean()
   m_iNPos = iNPos
   m_iNIns = iNIns
   ReDim m_bIns(iNPos - 1)
   
   If m_iNPos Mod m_iNIns = 0 Then
      'uEqually
      u1Group m_iNPos, m_iNIns, m_bIns(), 0  'equally
   Else
      If m_iNPos Mod 2 = 0 Then
         u2Poles
      Else
         uOdd
      End If
   End If
   fbIns = m_bIns()
End Function



Private Sub u2Poles()
   Dim _
   iNPos1 As Integer, iNPos2 As Integer, _
   iNIns1 As Integer, iNIns2 As Integer
   
   iNPos1 = Fix(m_iNPos / 2)
   iNPos2 = m_iNPos - iNPos1
   iNIns1 = Fix(m_iNIns / 2)
   iNIns2 = m_iNIns - iNIns1
   
   'If Fix(m_iNIns / m_iNPos) = 1 Then
   If CInt(m_iNPos / m_iNIns) < 3 Then       'dense

      u2Groups iNPos1, iNIns1, m_bIns(), 0
      u2Groups iNPos2, iNIns2, m_bIns(), iNPos1
   Else                                      'rarefied
      If m_iNIns Mod 2 = 1 Then              'Odd
         u1Group m_iNPos, m_iNIns, m_bIns(), 0     'equally
      Else                                   'Even
         u1Group iNPos1, iNIns1, m_bIns(), 0
         u1Group iNPos2, iNIns2, m_bIns(), iNPos1
      End If
   End If
End Sub


Private Sub uOdd()
   
   Dim _
   iStep As Integer, iInterval As Integer

   iStep = fiStep(m_iNPos, m_iNIns)
   iInterval = m_iNPos - 1 - (m_iNIns - 1) * iStep

   If Fix(iStep) < 2 Or iInterval > 1.5 * iStep Then    'dense
      Dim _
      iNPos1 As Integer, iNPos2 As Integer, _
      iNIns1 As Integer, iNIns2 As Integer

      iNPos1 = Fix(m_iNPos / 2)
      iNPos2 = m_iNPos - iNPos1
      iNIns1 = Fix(m_iNIns / 2)
      iNIns2 = m_iNIns - iNIns1
   
      u2Groups iNPos1, iNIns1, m_bIns(), 0
      u2Groups iNPos2, iNIns2, m_bIns(), iNPos1
   
   Else                                                  'rarefied
      u1Group m_iNPos, m_iNIns, m_bIns(), 0
   End If
End Sub


Private Sub u2Groups(iNPos As Integer, _
                     iNIns As Integer, _
                     b() As Boolean, _
                     iIBeg As Integer)
   Dim _
   iNPos1 As Integer, iNPos2 As Integer, _
   iNIns1 As Integer, iNIns2 As Integer
   
   iNPos1 = Fix(iNPos / 2)
   iNPos2 = iNPos - iNPos1       '>= iNPos1
   iNIns1 = Fix(iNIns / 2)
   iNIns2 = iNIns - iNIns1       '>= iNIns1
   
   u1Group iNPos1, iNIns1, b, iIBeg
   u1Group iNPos2, iNIns2, b, iIBeg + iNPos1
End Sub


Private Function u1Group(iNPos As Integer, _
                         iNIns As Integer, _
                         b() As Boolean, _
                         iIBeg As Integer)
   Dim _
   iStep As Integer, i As Integer, _
   iIBeforeLast As Integer, iInterval As Integer
   
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


Private Sub uEqually()
   
   Dim iStep As Integer, i As Integer
   
   iStep = m_iNPos / m_iNIns
   
   For i = 0 To m_iNPos - 1 Step iStep
      m_bIns(i) = True
   Next i
End Sub

