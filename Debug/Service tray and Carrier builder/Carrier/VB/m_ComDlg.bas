Attribute VB_Name = "m_ComDlg"
Option Explicit



Public Function fsChooseFilePath(dlgOpnSv As CommonDialog, _
                                 sTitle As String, _
                                 sFilters As String, _
                                 bOpenFlg As Boolean, _
                                 sInitFileName As String) As String
' Get the text filename from the user.
' Filter example:
' "Text Files (*.txt;*.dat)|*.txt;*.dat|All Files (*.*)|*.*"
   
   Dim _
   sCurDir As String, sDirTmp As String, sFileNameTmp As String, _
   iRet As Integer, sRet As String
'dd
   On Error Resume Next
   
   dlgOpnSv.DialogTitle = sTitle
   dlgOpnSv.Filter = sFilters
   dlgOpnSv.Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or _
                    cdlOFNNoChangeDir Or cdlOFNHideReadOnly
   Do
      dlgOpnSv.CancelError = True ' Why within the loop?
      dlgOpnSv.FileName = sInitFileName ' sFileNameTmp
      If bOpenFlg Then
         dlgOpnSv.ShowOpen
      Else
         dlgOpnSv.Flags = dlgOpnSv.Flags Or cdlOFNHideReadOnly
         dlgOpnSv.ShowSave
      End If
      
      If Err = cdlCancel Then
         DoEvents
         Exit Function
      End If
      If Err Then
         MsgBox Error$, 48
         Exit Function
      End If
      sRet = dlgOpnSv.FileName

      If Not bOpenFlg Then
         Exit Do
      Else
         ' If the file doesn't exist, go back.
         iRet = Len(Dir$(sRet))
   
         If iRet Then
            Exit Do
         Else
            MsgBox sRet + " not found!" + Space(5), 48
         End If ' iRet
      End If ' Not  bOpenFlg
   Loop
   DoEvents
   fsChooseFilePath = sRet
End Function



Public Function fsOpenFiles(dlgOpnSv As CommonDialog, _
                            sTitle As String, _
                            sFilters As String, _
                            sInitFileName As String) As String()
'Get the text filename from the user.
       
   Dim _
      sCurDir As String, sDirTmp As String, sFileNameTmp As String, _
      sFls() As String, iIU As Integer, iRet As Integer, _
      sFls1() As String, i As Integer
   
   On Error Resume Next
   
   dlgOpnSv.DialogTitle = sTitle
   'Filter example: "Text Files (*.TXT)|*.txt|All Files (*.*)|*.*"
   dlgOpnSv.Filter = sFilters
   dlgOpnSv.Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist _
                    Or cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or _
                    cdlOFNExplorer
   
   Do
      dlgOpnSv.CancelError = True ' Why within the loop?
      dlgOpnSv.FileName = sInitFileName ' sFileNameTmp
      dlgOpnSv.ShowOpen
      
      If Err = cdlCancel Then
         DoEvents
         Exit Function
      End If
      
      If Err Then
         MsgBox Error$, vbExclamation, "m_ComDlg.fsChooseFilePathM"
         Exit Function
      End If
      
      sFls() = Split(dlgOpnSv.FileName, Chr(0))
      iIU = UBound(sFls())
      
      If iIU = 0 Then
         iRet = Len(Dir$(sFls(0)))
         If Not iRet Then _
            MsgBox "File " & sFls(0) & " is missing." & Space(7), _
                   vbExclamation, "m_ComDlg.fsOpenFiles"
            Erase sFls()
      Else
         ReDim sFls1(iIU - 1)
         
         For i = 0 To iIU
            sFls1(i) = sFls(0) & "\" & sFls(i + 1)
         Next i
         
         iRet = 1
      End If
      
   Loop Until iRet
   
   If iIU = 0 Then
      fsOpenFiles = sFls()
   Else
      fsOpenFiles = sFls1()
   End If
   
   DoEvents
End Function

