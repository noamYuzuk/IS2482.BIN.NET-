Attribute VB_Name = "m_Msg"
  Option Explicit
   
  
   Public Sub g_uMsgE(Optional sTitle As String)
   
      If sTitle = vbNullString Then sTitle = App.Title
      
      MsgBox sAddSpaces("Error number " & Err.Number & "." & _
                         vbCrLf & Err.Description), vbCritical, sTitle
   End Sub


   Public Function g_fbMsgQ(sMsg As String, _
                           Optional sTitle As String) As Boolean
      Dim iRet As Integer

      If sTitle = vbNullString Then sTitle = App.Title

      iRet = MsgBox(sAddSpaces(sMsg), _
                    vbQuestion + vbDefaultButton2 + vbYesNo, sTitle)
      
      If iRet = vbYes Then g_fbMsgQ = True
   End Function


   Public Sub g_uMsgI(sMsg As String, Optional sTitle As String)

      If sTitle = vbNullString Then sTitle = App.Title

      MsgBox sAddSpaces(sMsg), vbInformation, sTitle
   End Sub


   Public Sub g_uMsgW(sMsg As String, Optional sTitle As String)

      If sTitle = vbNullString Then sTitle = App.Title

      MsgBox sAddSpaces(sMsg), vbExclamation, sTitle
   End Sub


   Public Sub g_MsgC(sMsg As String, Optional sTitle As String)

      If sTitle = vbNullString Then sTitle = App.Title

      MsgBox sAddSpaces(sMsg), vbCritical, sTitle
   End Sub


   Public Sub g_uLog(sMsg As String, Optional sRemark As String)
       
      Dim _
      iNFile As Integer, s As String
      
      iNFile = FreeFile
      Open App.path & "\" & App.EXEName & ".log" For Append As #iNFile
      
      s = Now & _
          vbCrLf & Space(5) & sMsg & _
          vbCrLf & Space(5) & sRemark & vbCrLf
          
      Print #iNFile, s
      Close iNFile
   End Sub




   Private Function sAddSpaces(ByVal s As String) As String
   
      If s = vbNullString Then Exit Function
      
      Dim sAr() As String, i As Integer
      
      sAr = Split(s, vbCrLf)
      sAddSpaces = sAr(0) + Space(7)
      For i = 1 To UBound(sAr)
         sAddSpaces = sAddSpaces & vbCrLf & sAr(i) & Space(7)
      Next i
   End Function
