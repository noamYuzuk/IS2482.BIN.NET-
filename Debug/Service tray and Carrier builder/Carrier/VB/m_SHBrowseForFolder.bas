Attribute VB_Name = "m_SHBrowseForFolder"
Option Explicit

'Using: m_Files.bat, m_Msg.bas


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2003 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'common to both methods

Public Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Public Declare Function SHBrowseForFolder Lib "shell32" _
   Alias "SHBrowseForFolderA" _
   (lpBrowseInfo As BROWSEINFO) As Long

Public Declare Function SHGetPathFromIDList Lib "shell32" _
   Alias "SHGetPathFromIDListA" _
   (ByVal pidl As Long, _
   ByVal pszPath As String) As Long

Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
   (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
   
Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
   (pDest As Any, _
    pSource As Any, _
    ByVal dwLength As Long)
    
Public Const MAX_PATH = 260
Public Const WM_USER = &H400
Public Const BFFM_INITIALIZED = 1

'Constants ending in 'A' are for Win95 ANSI
'calls; those ending in 'W' are the wide Unicode
'calls for NT.

'Sets the status text to the null-terminated
'string specified by the lParam parameter.
'wParam is ignored and should be set to 0.
Public Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Public Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)

'If the lParam  parameter is non-zero, enables the
'OK button, or disables it if lParam is zero.
'(docs erroneously said wParam!)
'wParam is ignored and should be set to 0.
Public Const BFFM_ENABLEOK As Long = (WM_USER + 101)

'Selects the specified folder. If the wParam
'parameter is FALSE, the lParam parameter is the
'PIDL of the folder to select , or it is the path
'of the folder if wParam is the C value TRUE (or 1).
'Note that after this message is sent, the browse
'dialog receives a subsequent BFFM_SELECTIONCHANGED
'message.
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
   

'specific to the PIDL method
'Undocumented call for the example. IShellFolder's
'ParseDisplayName member function should be used instead.
Public Declare Function SHSimpleIDListFromPath Lib _
   "shell32" Alias "#162" _
   (ByVal szPath As String) As Long


'specific to the STRING method
Public Declare Function LocalAlloc Lib "kernel32" _
   (ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long
    
Public Declare Function LocalFree Lib "kernel32" _
   (ByVal hMem As Long) As Long

Public Declare Function lstrcpyA Lib "kernel32" _
   (lpString1 As Any, lpString2 As Any) As Long

Public Declare Function lstrlenA Lib "kernel32" _
   (lpString As Any) As Long

Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

'windows-defined type OSVERSIONINFO
Public Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type
        
Public Const VER_PLATFORM_WIN32_NT = 2

Public Declare Function GetVersionEx Lib "kernel32" _
   Alias "GetVersionExA" _
  (lpVersionInformation As OSVERSIONINFO) As Long



Public Function BrowseCallbackProcStr(ByVal hWnd As Long, _
                                      ByVal uMsg As Long, _
                                      ByVal lParam As Long, _
                                      ByVal lpData As Long) As Long
                                       
  'Callback for the Browse STRING method.
 
  'On initialization, set the dialog's
  'pre-selected folder from the pointer
  'to the path allocated as bi.lParam,
  'passed back to the callback as lpData param.
 
   Select Case uMsg
      Case BFFM_INITIALIZED
      
         Call SendMessage(hWnd, BFFM_SETSELECTIONA, _
                          True, ByVal lpData)
                          
         Case Else:
         
   End Select
          
End Function
          


Public Function BrowseCallbackProc(ByVal hWnd As Long, _
                                   ByVal uMsg As Long, _
                                   ByVal lParam As Long, _
                                   ByVal lpData As Long) As Long
 
  'Callback for the Browse PIDL method.
 
  'On initialization, set the dialog's
  'pre-selected folder using the pidl
  'set as the bi.lParam, and passed back
  'to the callback as lpData param.
   Select Case uMsg
      Case BFFM_INITIALIZED
      
         Call SendMessage(hWnd, BFFM_SETSELECTIONA, _
                          False, ByVal lpData)
                          
         Case Else:
         
   End Select

End Function


Public Function FARPROC(pfn As Long) As Long
  
  'A dummy procedure that receives and returns
  'the value of the AddressOf operator.
 
  'This workaround is needed as you can't assign
  'AddressOf directly to a member of a user-
  'defined type, but you can assign it to another
  'long and use that (as returned here)
  FARPROC = pfn

End Function



Public Function BrowseForFolderByPIDL(sSelPath As String, _
                                      sTitle As String, _
                                      frmOwner As Form) As String

   Dim BI As BROWSEINFO
   Dim pidl As Long
   Dim sPath As String * MAX_PATH
          
   With BI
      'owner of the dialog. Pass 0 for the desktop.
      .hOwner = frmOwner.hWnd
     
      'The desktop folder will be the dialog's
      'root folder. SHSimpleIDListFromPath return
      'values can also be used to set this. This
      'member determines the 'root' point of the
      'Browse display.
      .pidlRoot = 0
     
      'Set the dialog's prompt string, if desired
      .lpszTitle = sTitle
     
      'Obtain and set the address of the callback
      'function. We need this workaround as you can't
      'assign the AddressOf directly to a member of
      'a user-defined type, but you can set assign it
      'to another long and use that (as returned in
      'the FARPROC call!!)
      .lpfn = FARPROC(AddressOf BrowseCallbackProc)
     
      'Obtain and set the pidl of the pre-selected folder
      .lParam = GetPIDLFromPath(sSelPath)
   End With
          
  'Shows the browse dialog and doesn't return until the
  'dialog is closed. The BrowseCallbackProc below will
  'receive all browse dialog specific messages while
  'the dialog is open. pidl will contain the pidl of
  'the selected folder if the dialog is not cancelled.
   pidl = SHBrowseForFolder(BI)
          
   If pidl Then
          
     'Get the path from the selected folder's pidl returned
     'from the SHBrowseForFolder call. Returns True on success.
     'Note: sPath must be pre-allocated!)
      If SHGetPathFromIDList(pidl, sPath) Then
     
        'Return the path
         BrowseForFolderByPIDL = Left$(sPath, InStr(sPath, vbNullChar) - 1)
        
      End If
     
     'Free the memory allocated for the pidl.
      Call CoTaskMemFree(pidl)
     
   End If
          
  'Free the memory allocated for
  'the pre-selected folder.
   Call CoTaskMemFree(BI.lParam)
          
End Function
         
         
Public Function fsBrowseForFolderByPath(sSelPath As String, _
                                        sTitle As String, _
                                        frmOwner As Form) As String
         
Dim BI As BROWSEINFO
Dim pidl As Long
Dim lpSelPath As Long
Dim sPath As String * MAX_PATH
          
   With BI
      'owner of the dialog. Pass 0 for the desktop.
      .hOwner = frmOwner.hWnd
           
      'The desktop folder will be the dialog's root folder.
      'SHSimpleIDListFromPath can also be used to set this value.
      .pidlRoot = 0
   
      'Set the dialog's prompt string
      .lpszTitle = sTitle
   
      'Obtain and set the address of the callback function
      .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
   
      'Now the fun part. Allocate some memory for the dialog's
      'selected folder path (sSelPath), blast the string into
      'the allocated memory, and set the value of the returned
      'pointer to lParam (checking LocalAlloc's success is
      'omitted for brevity). Note: VB's StrPtr function won't
      'work here because a variable's memory address goes out
      'of scope when passed to SHBrowseForFolder.
      '
      If Len(sSelPath) <> 0 Then '<added by MPh
         'Note: Win2000 requires that the memory block
         'include extra space for the string's terminating null.
         lpSelPath = LocalAlloc(LPTR, Len(sSelPath)) + 1
         CopyMemory ByVal lpSelPath, ByVal sSelPath, Len(sSelPath) + 1
         .lParam = lpSelPath
      Else '<added by MPh
         .lParam = 0
      End If
      
   End With
   
  'Shows the browse dialog and doesn't return until the
  'dialog is closed. The BrowseCallbackProcStr will
  'receive all browse dialog specific messages while
  'the dialog is open. pidl will contain the pidl of the
  'selected folder if the dialog is not cancelled.
   pidl = SHBrowseForFolder(BI)
           
   If pidl Then
           
     'Get the path from the selected folder's pidl returned
     'from the SHBrowseForFolder call (rtns True on success,
     'sPath must be pre-allocated!)
          
      If SHGetPathFromIDList(pidl, sPath) Then
     
        'Return the path
         fsBrowseForFolderByPath = Left$(sPath, InStr(sPath, vbNullChar) - 1)
        
      End If
     
     'Free the memory allocated for the pidl.
      Call CoTaskMemFree(pidl)
           
   End If
           
  'Free the allocated pointer
   Call LocalFree(lpSelPath)
         
End Function


Public Function GetPIDLFromPath(sPath As String) As Long

  'return the pidl to the path supplied by calling the
  'undocumented API #162 (our name SHSimpleIDListFromPath).
  'This function is necessary as, unlike documented APIs,
  'the API is not implemented in 'A' or 'W' versions.

  If IsWinNT Then
    GetPIDLFromPath = SHSimpleIDListFromPath(StrConv(sPath, vbUnicode))
  Else
    GetPIDLFromPath = SHSimpleIDListFromPath(sPath)
  End If

End Function


Public Function IsWinNT() As Boolean

   #If Win32 Then
  
      Dim OSV As OSVERSIONINFO
   
      OSV.OSVSize = Len(OSV)
   
     'API returns 1 if a successful call
      If GetVersionEx(OSV) = 1 Then
   
        'PlatformId contains a value representing
        'the OS, so if it's VER_PLATFORM_WIN32_NT,
        'return true
         IsWinNT = OSV.PlatformID = VER_PLATFORM_WIN32_NT
      End If

   #End If

End Function


Public Function UnqualifyPath(sPath As String) As String

  'Qualifying a path involves assuring that its format
  'is valid, including a trailing slash, ready for a
  'filename. Since SHBrowseForFolder will not pre-select
  'the path if it contains the trailing slash, it must be
  'removed, hence 'unqualifying' the path.
   If Len(sPath) > 0 Then
   
      If Right$(sPath, 1) = "\" Then
      
         UnqualifyPath = Left$(sPath, Len(sPath) - 1)
         Exit Function
      
      End If
   
   End If
   
   UnqualifyPath = sPath
   
End Function


Public Function fsBrowseForFolder(ByVal sFolder As String, _
                                 sTitle As String, frm As Form)
   Dim sDirInit As String, sDirTmp As String
'dd
   On Error GoTo ErrHndl
   If Not fbFolderExist(sFolder) Then     'from m_Files.bat
      sDirInit = "c:\"
   Else
      sDirInit = sFolder
   End If
   
   'fsBrowseForFolder = fsBrowseForFolderByPath(sDirInit, sTitle, frm)
   fsBrowseForFolder = BrowseForFolderByPIDL(sDirInit, sTitle, frm)
 
   Exit Function
ErrHndl:
    g_uMsgE "m_SHBrowseForFolder.fsBrowseForFolder"
End Function
