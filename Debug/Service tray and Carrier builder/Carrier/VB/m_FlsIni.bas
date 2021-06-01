Attribute VB_Name = "m_FlsIni"
Option Explicit

'Using: m_Msg.bas

Private Const _
mcsError As String = "<error>", mcsEmpty As String = ""

Public Declare Function GetPrivateProfileString _
                     Lib "kernel32.dll" _
                   Alias "GetPrivateProfileStringA" ( _
            ByVal lpApplicationName As String, _
            ByVal lpKeyName As Any, ByVal lpDefault As String, _
            ByVal lpReturnedString As String, ByVal nSize As Long, _
            ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString _
                    Lib "kernel32.dll" _
                  Alias "WritePrivateProfileStringA" ( _
            ByVal lpApplicationName As String, _
            ByVal lpKeyName As String, _
            ByVal lpString As Any, _
            ByVal lpFileName As String) As Long


Private Declare Function GetPrivateProfileSection _
                    Lib "kernel32" _
                  Alias "GetPrivateProfileSectionA" ( _
            ByVal lpAppName As String, _
            ByVal lpReturnedString As String, _
            ByVal nSize As Long, _
            ByVal lpFileName As String) As Long



Private Declare Function WritePrivateProfileSection _
                  Lib "kernel32" _
                  Alias "WritePrivateProfileSectionA" ( _
            ByVal lpAppName As String, _
            ByVal lpString As Any, _
            ByVal lpFileName As String) As Long



Public Function fsGetIniItem(sNameF As String, sNameGrp As String, _
                             sNameItem As String) As String
   Dim objFile As New FileSystemObject
'dd
   If Not objFile.FileExists(sNameF) Then
      MsgBox "File '" & sNameF & "' does not exist." & Space(5), _
              vbExclamation, "Error in fsGetIniItem"
      Exit Function
   End If
   
   Dim sRet As String, lLenStr As Long
'dd
   sRet = Space(255&)
   lLenStr = GetPrivateProfileString(sNameGrp, sNameItem, _
                                     mcsError, sRet, 255, sNameF)
   sRet = Left(sRet, lLenStr)
   If sRet = mcsError Or sNameItem = mcsEmpty Then
      MsgBox "Items group '" & sNameGrp & "' or item '" & _
              sNameItem & "' does not exist." & Space(5), _
              vbExclamation, "Error in fsGetIniItem"
      sRet = fsSetIniValue(sNameF, sNameGrp, sNameItem)
   End If
   
   fsGetIniItem = sRet
End Function


Public Function fsGetIniItem1(sNameF As String, sNameGrp As String, _
                              sNameItem As String) As String
   Dim objFile As New FileSystemObject
'dd
   If Not objFile.FileExists(sNameF) Then
      MsgBox "File '" & sNameF & "' does not exist." & Space(5), _
              vbExclamation, "Error in fsGetIniItem1"
      Exit Function
   End If
   
   Dim sRet As String, lLenStr As Long
'dd
   sRet = Space(255&)
   lLenStr = GetPrivateProfileString(sNameGrp, sNameItem, _
                                     vbNullString, sRet, 255, sNameF)
   sRet = Left(sRet, lLenStr)
   
   fsGetIniItem1 = sRet
End Function


Public Function fsGetIniItem2(sNameGrp As String, _
                              sNameItem As String) As String
   Dim _
   sNameF As String, _
   objFile As New FileSystemObject
'dd
   sNameF = App.path & "\" & App.EXEName & ".ini"
   If Not objFile.FileExists(sNameF) Then
      MsgBox "File '" & sNameF & "' does not exist." & Space(5), _
              vbExclamation, "Error in fsGetIniItem"
'      Exit Function
   End If
   
   Dim sRet As String, lLenStr As Long
'dd
   sRet = Space(255&)
   lLenStr = GetPrivateProfileString(sNameGrp, sNameItem, _
                                     mcsError, sRet, 255, sNameF)
   sRet = Left(sRet, lLenStr)
   If sRet = mcsError Or sNameItem = mcsEmpty Then
      MsgBox "Items group '" & sNameGrp & "' or item '" & _
              sNameItem & "' does not exist." & Space(5), _
              vbExclamation, "Error in fsGetIniItem"
      sRet = fsSetIniValue(sNameF, sNameGrp, sNameItem)
   End If
   
   fsGetIniItem2 = sRet
End Function


Public Function fbSetIniItem(sNameF As String, sNameGrp As String, _
                             sNameItem As String, sStr As String) _
                             As Boolean
   Dim lRet As Long
'dd
   lRet = WritePrivateProfileString(sNameGrp, sNameItem, sStr, sNameF)
   If lRet = 0 Then
      MsgBox "Item does not saved." & Space(5), vbExclamation, _
             "Error in fsSetIniItem"
      Exit Function
   End If
   
   fbSetIniItem = True
End Function


Public Function fbSetIniItem2(sNameGrp As String, _
                              sNameItem As String, sStr As String) _
                              As Boolean
   Dim _
   sNameF As String, _
   lRet As Long
'dd
   sNameF = App.path & "\" & App.EXEName & ".ini"
   If sStr <> vbNullString Then
      lRet = WritePrivateProfileString(sNameGrp, sNameItem, _
                                       sStr, sNameF)
   Else
      lRet = WritePrivateProfileString(sNameGrp, sNameItem, _
                                       0&, sNameF)
   End If
   
   If lRet = 0 Then
      MsgBox "Item does not saved." & Space(5), vbExclamation, _
             "Error in fsSetIniItem"
      Exit Function
   End If
   
   fbSetIniItem2 = True
End Function


Public Function fsSetIniValue( _
                     sPathIni As String, sDigGrpIni As String, _
                     sNameItem As String, _
                     Optional sTypeFl As String) As String
   
   Dim iRet As Integer, sRet As String
'dd
   While iRet <> vbYes
      sRet = InputBox("Input value for '" & sNameItem & "' item", _
                          "Change '" & sNameItem & "' item", _
                           mcsEmpty)
      If sRet = mcsEmpty Then Exit Function
      
      Select Case sTypeFl
      Case "d"
         sRet = Val(sRet)
      Case "i"
         sRet = CInt(sRet)
      Case "b"
         sRet = CBool(sRet)
      End Select
      
      iRet = MsgBox(sNameItem & " is   " & sRet & Space(7) & vbCrLf & _
                    "Is it correct?", vbYesNo + vbQuestion, _
                     sNameItem & " setting")
   Wend
   
   fbSetIniItem sPathIni, sDigGrpIni, sNameItem, sRet
   fsSetIniValue = sRet
End Function


Public Function fbReadInitPars( _
                                sPathIni As String, _
                                sNmGrp As String, _
                                sNmPar() As String, _
                                sValPar() As String) As Boolean
   Dim _
   iIU As Integer, iI As Integer, sTemp As String
'dd
   iIU = UBound(sNmPar)
   ReDim sValPar(iIU)
   
   For iI = 0 To iIU
      sTemp = fsGetIniItem(sPathIni, sNmGrp, sNmPar(iI))
      If sTemp <> mcsEmpty Then
         sValPar(iI) = sTemp
      Else
         MsgBox "In group '" & sNmGrp & " '" & sNmPar(iI) & "'" & _
                " not defined." & Space(7), vbCritical, _
                "Error in fbReadInitPars"
         Exit Function
      End If
   Next iI
   
   fbReadInitPars = True
End Function


Public Function fbSaveInitPars( _
                                 sPathIni As String, _
                                 sNmGrp As String, _
                                 sNmPar() As String, _
                                 sValPar() As String) As Boolean
   Dim _
   iIU As Integer, iI As Integer, iRet As Integer
'dd
   iIU = UBound(sNmPar)
   For iI = 0 To iIU
      iRet = WritePrivateProfileString(sNmGrp, sNmPar(iI), _
                                       sValPar(iI), sPathIni)
      If iRet = 0 Then
         MsgBox "In group '" & sNmGrp & " '" & sNmPar(iI) & "'" & _
                " not saved." & Space(7), vbCritical, _
                "Error in fbSaveInitPars"
         Exit Function
      End If
   Next iI
   fbSaveInitPars = True
End Function



Public Function fsGetItemsVal( _
                              sNameF As String, _
                              sNameGrp As String) As String()
'   Dim objFile As New FileSystemObject
'dd
'   If Not objFile.FileExists(sNameF) Then
'      MsgBox "File '" & sNameF & "' does not exist." & Space(5), _
'              vbExclamation, "Error in fsGetIniItem"
'      Exit Function
'   End If
   
   If Dir(sNameF) = vbNullString Then
      MsgBox "File '" & sNameF & "' does not exist." & Space(5), _
              vbExclamation, "Error in fsGetIniItem"
      Exit Function
   End If
   
   Const _
   ciLenS As Integer = 10000
   Dim _
   sRet As String, lLenStr As Long, sItemsVal() As String
'dd
   sRet = Space(ciLenS)
   On Error GoTo ErrHandler
   lLenStr = GetPrivateProfileSection(sNameGrp, sRet, _
                                      ciLenS, sNameF)
   If lLenStr = 0 Then Exit Function
   
   On Error GoTo 0
   sRet = Left(sRet, lLenStr - 1)
   sItemsVal() = Split(sRet, Chr(0))
   
   Dim _
   iIU As Integer, iIUPart As Integer, _
   iI As Integer, sPart() As String
'dd
   iIU = UBound(sItemsVal())
   For iI = 0 To iIU
      sPart() = Split(sItemsVal(iI), "=")
      iIUPart = UBound(sPart())
      sItemsVal(iI) = sPart(iIUPart)
   Next iI
   
   fsGetItemsVal = sItemsVal()
   Exit Function
ErrHandler:
   g_uMsgE ("m_FlsIni.fsGetItemsVal")
End Function



Public Function fsGetItemsVal2( _
                                sNameGrp As String) As String()
   
   Const _
   ciLenS As Integer = 1000
   Dim _
   sNameF As String, _
   sRet As String, lLenStr As Long, sItemsVal() As String
'dd
   On Error GoTo ErrHandler
   
   sNameF = App.path & "\" & App.EXEName & ".ini"
   sRet = Space(ciLenS)
   
   lLenStr = GetPrivateProfileSection(sNameGrp, sRet, _
                                      ciLenS, sNameF)
   If lLenStr = 0 Then Exit Function
   
   On Error GoTo 0
   sRet = Left(sRet, lLenStr - 1)
   sItemsVal() = Split(sRet, Chr(0))
   
   Dim _
   iIU As Integer, iIUPart As Integer, _
   iI As Integer, sPart() As String
'dd
   iIU = UBound(sItemsVal())
   For iI = 0 To iIU
      sPart() = Split(sItemsVal(iI), "=")
      iIUPart = UBound(sPart())
      sItemsVal(iI) = sPart(iIUPart)
   Next iI
   
   fsGetItemsVal2 = sItemsVal()
   
   Exit Function
ErrHandler:
   g_uMsgE ("m_FlsIni.fsGetItemsVal")
End Function



Public Function fbGetItemsVal2( _
                                sNameGrp As String, _
                                sItemsO() As String) As Boolean
   Const _
   ciLenS As Integer = 1000
   Dim _
   sNameF As String, _
   sRet As String, lLenStr As Long
'dd
   On Error GoTo ErrHandler
   
   sNameF = App.path & "\" & App.EXEName & ".ini"
   sRet = Space(ciLenS)
   
   lLenStr = GetPrivateProfileSection(sNameGrp, sRet, _
                                      ciLenS, sNameF)
   If lLenStr = 0 Then Exit Function
   
   
   On Error GoTo 0
   sRet = Left(sRet, lLenStr - 1)
   sItemsO() = Split(sRet, Chr(0))
   
   Dim _
   iIU As Integer, iIUPart As Integer, _
   iI As Integer, sPart() As String
'dd
   iIU = UBound(sItemsO())
   For iI = 0 To iIU
      sPart() = Split(sItemsO(iI), "=")
      iIUPart = UBound(sPart())
      sItemsO(iI) = sPart(iIUPart)
   Next iI
   
   fbGetItemsVal2 = True
   
   Exit Function
ErrHandler:
   g_uMsgE ("m_FlsIni.fsGetItemsVal")
End Function



Public Function fbSetItemsVal2( _
                                 sNmGrp As String, _
                                 sValPar() As String) As Boolean
   'If return 0 - failed, else - succeeded
   Dim _
   sNameF As String, iIU As Integer, iI As Integer, _
   iRet As Integer, sBuffer As String
'dd
   sNameF = App.path & "\" & App.EXEName & ".ini"
   
   iIU = UBound(sValPar())
'dd
   For iI = 0 To iIU
      sBuffer = sBuffer & iI & "=" & sValPar(iI) & Chr(0)
   Next iI
   
   On Error GoTo ErrHandler
   fbSetItemsVal2 = (WritePrivateProfileSection( _
                                 sNmGrp, sBuffer, sNameF) <> 0)
   Exit Function
ErrHandler:
   g_uMsgE ("m_FlsIni.fsGetItemsVal")
End Function


