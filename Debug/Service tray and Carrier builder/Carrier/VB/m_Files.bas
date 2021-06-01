Attribute VB_Name = "m_Files"
Option Explicit
Private Declare Function SHFileOperation Lib "Shell32.dll" _
                  Alias "SHFileOperationA" ( _
                         lpFileOp As SHFILEOPSTRUCT) As Long

Private Type SHFILEOPSTRUCT
   hWnd As Long
   wFunc As Long
   pFrom As String
   pTo As String
   fFlags As Integer
   fAnyOperationsAborted As Long
   hNameMappings As Long
   lpszProgressTitle As String
End Type

Private Const FO_COPY = &H2&
Private Const FO_DELETE = &H3&
Private Const FO_MOVE = &H1&
Private Const FO_RENAME = &H4&
Private Const FOF_SILENT = &H4
Private Const FOF_SIMPLEPROGRESS = &H100
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_NOCONFIRMMKDIR = &H200
Private Const FOF_RENAMEONCOLLISION = &H8
Private Const FOF_ALLOWUNDO = &H40


Declare Function GetShortPathName Lib "kernel32.dll" _
          Alias "GetShortPathNameA" _
            (ByVal lpszLongPath As String, _
             ByVal lpszShortPath As String, _
             ByVal cchBuffer As Long) As Long


Public Function fiReplaceLine(sSegmentOfReplaced As String, _
                              sReplacing As String, _
                              sFilePath As String) As String

'Replace a lines in the file according to given line segment
'and return a number of last replaced line.
   
   Dim iFNum As Integer, sTemp As String
   Dim sOut As String, iI As Integer
      
   iFNum = FreeFile
   Open sFilePath For Input As #iFNum
   
   iI = 1
   Line Input #iFNum, sTemp
   If InStr(sTemp, sSegmentOfReplaced) = 0 Then
      sOut = sTemp
   Else
      If sReplacing <> "" Then sOut = sReplacing
      fiReplaceLine = iI
      iI = iI - 1
   End If
   
   While Not EOF(iFNum)
      iI = iI + 1
      Line Input #iFNum, sTemp
      If InStr(sTemp, sSegmentOfReplaced) = 0 Then
         If sOut <> "" Then
            sOut = sOut & vbCrLf & sTemp
         Else: sOut = sTemp
         End If
      Else
         If sReplacing <> "" Then sOut = sOut & vbCrLf & sReplacing
         fiReplaceLine = iI
         iI = iI - 1
      End If
   Wend
   Close #iFNum
   
   iFNum = FreeFile
   Open sFilePath For Output As #iFNum
   Print #iFNum, sOut;
   Close #iFNum

End Function


Public Sub uInsertLine(iInsLineNum As Integer, _
                              sLine As String, _
                              sFilePath As String)

'Insert line into the file according to the line number.

   Dim iFNum As Integer, sTemp As String
   Dim sOut As String, iI As Integer
      
   iFNum = FreeFile
   Open sFilePath For Input As #iFNum
   
   If iInsLineNum = 1 Then
      sOut = iInsLineNum
   Else
      Line Input #iFNum, sTemp
      sOut = sTemp
      For iI = 2 To iInsLineNum - 1
         Line Input #iFNum, sTemp
         sOut = sOut & vbCrLf & sTemp
      Next iI
      sOut = sOut & vbCrLf & sLine
   End If
   
   While Not EOF(iFNum)
      Line Input #iFNum, sTemp
      sOut = sOut & vbCrLf & sTemp
   Wend
   Close #iFNum
   
   iFNum = FreeFile
   Open sFilePath For Output As #iFNum
   Print #iFNum, sOut;
   Close #iFNum

End Sub


Public Sub uDelLine(iDelLineNum As Integer, _
                              sFilePath As String)

'Delete line from the file according to the line number.

   Dim iFNum As Integer, sTemp As String
   Dim sOut As String, iI As Integer
      
   iFNum = FreeFile
   Open sFilePath For Input As #iFNum
   
   If iDelLineNum = 1 Then
      Line Input #iFNum, sTemp
   Else
      Line Input #iFNum, sTemp
      sOut = sTemp
      For iI = 2 To iDelLineNum - 1
         Line Input #iFNum, sTemp
         sOut = sOut & vbCrLf & sTemp
      Next iI
      Line Input #iFNum, sTemp
   End If
   
   While Not EOF(iFNum)
      Line Input #iFNum, sTemp
      sOut = sOut & vbCrLf & sTemp
   Wend
   Close #iFNum
   
   iFNum = FreeFile
   Open sFilePath For Output As #iFNum
   Print #iFNum, sOut;
   Close #iFNum

End Sub


Public Function fbStringToFile(sStr As String, sFileName As String) As Boolean
'Write string to file (no append).
   
   Dim iFNum As Integer
   
   On Error Resume Next
   If Dir(sFileName) <> "" Then Kill sFileName
   iFNum = FreeFile
   Open sFileName For Binary As #iFNum
   If Err Then
      MsgBox Error$ & vbCrLf & "No saving done." & Space(5), 48, _
             "Error in fbStringToFile function"
      Exit Function
   End If
   Put #iFNum, , sStr
   Close #iFNum
   fbStringToFile = True
End Function


Public Function fbDataToFile(vData As Variant, sFileName As String) As Boolean
'Write data to file (no append).
   
   Dim iFNum As Integer
   
   On Error Resume Next
   Kill sFileName
   iFNum = FreeFile
   Open sFileName For Binary As #iFNum
   If Err And Err.Number <> 53 Then
      MsgBox Error$ & Space(5), 48, "Error in fbDataToFile function"
      Exit Function
   End If
   Put #iFNum, , vData
   Close #iFNum
   fbDataToFile = True
End Function

   
Public Sub uStringToFile(sStr As String, sFileName As String)
'Write string to file (no append).
'Using: m_ErrMsg.bas
   Dim _
   iFNum As Integer
'dd
   iFNum = FreeFile
   On Error Resume Next
   Open sFileName For Output As #iFNum
   If Err Then
      g_uMsgE ("m_Files.uPutStrToFile"):    On Error GoTo 0
      Exit Sub
   End If
   Print #iFNum, sStr;
   Close #iFNum
End Sub


Public Sub uPutStrToFile( _
                          sStr As String, sPathF As String)
'Append string to file.
'Using: m_ErrMsg.bas
   Dim _
   iFNum As Integer
'dd
   iFNum = FreeFile
   On Error Resume Next
   Open sPathF For Binary As #iFNum
   If Err Then
      g_uMsgE ("uPutStrToFile"):    On Error GoTo 0
      Exit Sub
   End If
   Put #iFNum, , sStr
   Close #iFNum
End Sub


Public Function fbStringFromFile( _
                                 sFileName As String, _
                                 sStr_O As String) As Boolean
   Dim _
   iFNum As Integer, lFileLen As Long
'dd
   On Error GoTo ErrHandle
   
   iFNum = FreeFile
   Open sFileName For Binary As #iFNum
   
   lFileLen = LOF(iFNum)
   sStr_O = Space(lFileLen)
   Get #iFNum, , sStr_O
   Close #iFNum
   
   fbStringFromFile = True
   Exit Function
ErrHandle:
   g_uMsgE ("m_Files.fbStringFromFile")
End Function


Public Function fbStringsFromFile( _
                                   sFileName As String, _
                                   sStr_O() As String) As Boolean
   'Using m_Str.bas
   Dim _
   iFNum As Integer, lFileLen As Long, sTemp As String
'dd
   On Error GoTo ErrHandle
   
   iFNum = FreeFile
   Open sFileName For Binary As #iFNum
   
   lFileLen = LOF(iFNum)
   sTemp = Space(lFileLen)
   Get #iFNum, , sTemp
   Close #iFNum
   
   sTemp = Replace(sTemp, vbCr, vbNullString)
   sStr_O = Split(sTemp, vbLf)
   
   fbStringsFromFile = True
   Exit Function
ErrHandle:
   g_uMsgE ("m_Files.fbStringFromFile")
End Function


Public Function fvDataFromFile(sFileName As String) As Variant
'Read file to variant.

   Dim iFNum As Integer, lFileLen As Long, sTemp As String
   
   On Error Resume Next
   iFNum = FreeFile
   Open sFileName For Binary As #iFNum
   If Err Then
      MsgBox Error$, 48
      Exit Function
   End If
      lFileLen = LOF(iFNum)
      Get #iFNum, , fvDataFromFile
   Close #iFNum
End Function


Public Function fbDoublesFromFile(d() As Double, _
                                  sFileName As String) As Boolean
'Read file to array of doubles.
'Using: m_ErrMsg.bas

   Dim _
   iFNum As Integer, lFileLen As Long, _
   iIU As Integer, i As Integer
'dd
   If Dir(sFileName) = vbNullString Then
      MsgBox "File  " & sFileName & "  does not exist." & _
              Space(7), vbExclamation, "m_Files.fbDoublesFromFile"
      Exit Function
   End If
   On Error GoTo ErrHandle1
   iFNum = FreeFile
   Open sFileName For Binary As #iFNum
   On Error GoTo ErrHandle
   
   lFileLen = LOF(iFNum)
   iIU = Fix(lFileLen / 8) - 1
   ReDim d(iIU)
   For i = 0 To iIU
      Get #iFNum, , d(i)
   Next i
   On Error GoTo ErrHandle1
   Close #iFNum
   fbDoublesFromFile = True
   Exit Function
ErrHandle:
   Close #iFNum
ErrHandle1:
   g_uMsgE ("m_Files.fbDoublesFromFile")
End Function


Public Function fbDoublesToFile(d() As Double, _
                                sFileName As String) As Boolean
'Write array of doubles to file (no append).
'Using: m_ErrMsg.bas
   
   Dim _
   iFNum As Integer, _
   iIU As Integer, i As Integer
'dd
   On Error GoTo ErrHandle1
   If Dir(sFileName) <> vbNullString Then Kill sFileName
   iFNum = FreeFile
   Open sFileName For Binary As #iFNum
   On Error GoTo ErrHandle
   iIU = UBound(d())
   For i = 0 To iIU
      Put #iFNum, , d(i)
   Next i
   On Error GoTo ErrHandle
   Close #iFNum
   fbDoublesToFile = True
   Exit Function
ErrHandle:
   Close #iFNum
ErrHandle1:
   g_uMsgE ("m_Files.fbDoublesToFile")
End Function


Public Function fsCutPath(ByVal sEntierFileName As String) As String
' Truncate the path from an entier file name

   Dim iFl As Integer
   Do
      iFl = InStr(sEntierFileName, "\")
      sEntierFileName = Mid(sEntierFileName, iFl + 1)
   Loop Until iFl = 0
   
   fsCutPath = sEntierFileName
End Function


Public Function fsCutName(sFileName As String) As String
' Truncate the file name from an entier file name
   
   Dim iFl As Integer, sTemp As String
   
   sTemp = sFileName
   
   Do
      iFl = InStr(sTemp, "\")
     sTemp = Mid(sTemp, iFl + 1)
   Loop Until iFl = 0

   fsCutName = Mid(sFileName, 1, Len(sFileName) - Len(sTemp))
End Function


Public Function fsFileExtension(sFileName As String)
   Dim iDotPos As Integer
'dd
   iDotPos = InStrRev(sFileName, ".")
   If iDotPos = 0 Then Exit Function
   fsFileExtension = Mid(sFileName, iDotPos + 1)
End Function


Public Function fsChangeExtension(sFileName As String, _
                                  sExtension As String) As String
   Dim _
   iNameHeadLength As Integer
   
   iNameHeadLength = InStrRev(sFileName, ".")
   fsChangeExtension = Mid(sFileName, 1, iNameHeadLength) & sExtension
End Function


Public Function fsCutExtension(sFileName As String) As String
   Dim iDotPos As Integer
'dd
   iDotPos = InStrRev(sFileName, ".")
   If iDotPos = 0 Then
      fsCutExtension = sFileName
   Else
      fsCutExtension = Left$(sFileName, iDotPos - 1)
   End If
End Function


Public Function fbDelFile(sNameF As String) As Boolean
   If Dir(sNameF) = vbNullString Then
      fbDelFile = True
      Exit Function
   End If
   On Error GoTo ErrorMsg
      Kill sNameF
      fbDelFile = True
   Exit Function
ErrorMsg:
   MsgBox sNameF & Space(5) & vbCrLf & Err.Description & Space(5), _
          vbExclamation, _
         "Error in fbDelFile (No " & Err.Number & ")"
End Function


Public Function fbCopyFile(sPathFrom As String, sPathTo As String)
   Dim _
   fso As New FileSystemObject, f As File
   
   On Error GoTo ErrHandler
   Set f = fso.GetFile(sPathFrom)
   f.Copy sPathTo
   fbCopyFile = True
   Exit Function
ErrHandler:
   g_uMsgE "m_Files.fbCopyFile"
End Function


Public Function fbCopyFiles(sPathFrom() As String, sDirTo As String, _
                            hWnd As Long) As Boolean
Dim _
   iIU As Integer, i As Integer, sPathsFrom As String, _
   udOper As SHFILEOPSTRUCT, lRet As Long
'dd
   iIU = UBound(sPathFrom)
   For i = 0 To iIU
      sPathsFrom = sPathsFrom & sPathFrom(i) & Chr(0)
   Next i
   udOper.wFunc = FO_COPY
   udOper.pFrom = sPathsFrom
   udOper.hWnd = hWnd
   udOper.fFlags = 0   'FOF_SIMPLEPROGRESS
   udOper.pTo = sDirTo
   
   DoEvents
   lRet = SHFileOperation(udOper)
   If lRet = 0 Then fbCopyFiles = True
End Function


Public Function fbCopyFilesM(sPathFrom() As String, sDirTo() As String, _
                      hWnd As Long) As Boolean
Dim _
   iIU As Integer, i As Integer, sPathsFrom As String, _
   udOper As SHFILEOPSTRUCT, lRet As Long, bRet As Boolean
'dd
   iIU = UBound(sPathFrom)
   For i = 0 To iIU
      sPathsFrom = sPathsFrom & sPathFrom(i) & Chr(0)
   Next i
   udOper.wFunc = FO_COPY
   udOper.pFrom = sPathsFrom
   udOper.hWnd = hWnd
   udOper.fFlags = 0 'FOF_SIMPLEPROGRESS
   
   bRet = True
   iIU = UBound(sDirTo())
   For i = 0 To iIU
      If Dir(sDirTo(i) & "\", vbDirectory) = vbNullString Then
         MsgBox "Destination directory " & sDirTo(i) & " does not exist." & _
                Space(7), vbExclamation, "m_Files.fbCopyFilesM"
         bRet = False
      Else
         DoEvents
         udOper.pTo = sDirTo(i)
         lRet = SHFileOperation(udOper)
         If lRet <> 0 Then bRet = False
      End If
   Next i
   fbCopyFilesM = bRet
End Function


Public Function fbMoveFiles(sPathFrom() As String, _
                            sDirTo As String, _
                            hWnd As Long) As Boolean
Dim _
   iIU As Integer, i As Integer, sPathsFrom As String, _
   udOper As SHFILEOPSTRUCT, lRet As Long
'dd
   iIU = UBound(sPathFrom)
   For i = 0 To iIU
      sPathsFrom = sPathsFrom & sPathFrom(i) & Chr(0)
   Next i
   udOper.wFunc = FO_MOVE
   udOper.pFrom = sPathsFrom
   udOper.pTo = sDirTo
   udOper.hWnd = hWnd
   udOper.fFlags = 0 'FOF_SIMPLEPROGRESS
   DoEvents
   lRet = SHFileOperation(udOper)
   If lRet = 0 Then fbMoveFiles = True
End Function



Public Function fbDelFilesToRecycleBin(sPath() As String, _
                              hWnd As Long) As Boolean
Dim _
   iIU As Integer, i As Integer, sPaths As String, _
   udOper As SHFILEOPSTRUCT, lRet As Long
'dd
   iIU = UBound(sPath())
   For i = 0 To iIU
      sPaths = sPaths & sPath(i) & Chr(0)
   Next i
   udOper.wFunc = FO_DELETE
   udOper.pFrom = sPaths
   udOper.hWnd = hWnd
   udOper.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
   DoEvents
   lRet = SHFileOperation(udOper)
   If lRet = 0 Then fbDelFilesToRecycleBin = True
End Function



Public Function fsShortPath(sPathLong As String) As String
   Dim _
   sRet As String, lRet As Long
'dd
   lRet = GetShortPathName(sPathLong, sRet, 0)
   If lRet = 0 Then
      MsgBox "GetShortPathName failed" & Space(7), _
              vbCritical, "frm.MainfsShortPath"
      Exit Function
   End If
   sRet = Space(lRet - 1)
   GetShortPathName sPathLong, sRet, lRet
   fsShortPath = sRet
End Function


Public Function fbFileExist(s As String) As Boolean
   Dim fso As New FileSystemObject
'dd
   If s = vbNullString Or Not fso.FileExists(s) Then
      MsgBox "File " & s & " does not exist." & Space(7), _
             vbExclamation, "m_Files.fbFileExist"
      Exit Function
   End If
   
   fbFileExist = True
End Function


Public Function fbFolderExist(s As String) As Boolean
   Dim fso As New FileSystemObject
'dd
   If s = vbNullString Or Not fso.FolderExists(s) Then
      MsgBox "Folder " & s & " does not exist." & Space(7), _
             vbExclamation, "m_Files.fbFolderExist"
      Exit Function
   End If
   
   fbFolderExist = True
End Function


Public Function bfCreateFolder(path$)

Dim fso As New FileSystemObject
Dim i%, s$, a
 
  
  a = Split(path, "\")
  s = a(0)
  For i = 1 To UBound(a)
    If a(i) <> "" Then
      s = s & "\" & a(i)
      If Not fbFolderExist(s) Then
        fso.CreateFolder s
      End If
    End If
  Next i
  
End Function

