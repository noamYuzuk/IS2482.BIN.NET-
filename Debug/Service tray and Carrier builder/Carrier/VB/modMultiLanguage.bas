Attribute VB_Name = "modMultiLanguage"
Option Explicit


'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'NAME:      modMultiLanguage
'
'AUTHOR:    Amit Levi
'
'ABSTRACT:  This module holds the relevent actions to perform control language change
'
'DATE CREATED:  02/06/10
'
'DATE UPDATED:  date of update
'
'* Copyright (c) 2009 by SHAFIR PRODUCTION SYSTEMS
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------


'------------------------------ DECLARATIONS --------------------------------------
'----------------------------------------------------------------------------------

Public Const _
   mc_sSecSettings As String = "Application Settings", _
   mc_sItmStoreLang As String = "Laguage"
   
Public m_sSelectedLang As String




'------------------------------ PROCEDURES --------------------------------------
'----------------------------------------------------------------------------------

Public Sub uSetLanguage(frmSourceForm As Form, sLanguage As String)
On Error GoTo ErrHandler
Dim i As Integer
Dim vCaption As Variant
Dim cntrl As Control
   
    With frmSourceForm
        vCaption = Split(.Tag, ",")
        If UBound(vCaption) > 0 Then
            If sLanguage = "Hebrew" Then
                .RightToLeft = True
                .Caption = vCaption(1)
            Else
                .RightToLeft = False
                .Caption = vCaption(0)
            End If
        End If
    End With
    
    '/// changing all controls caption (any control configured to support multi language)
    For Each cntrl In frmSourceForm.Controls
        vCaption = Split(cntrl.Tag, ",")
        
        If UBound(vCaption) > 0 Then
            If TypeOf cntrl Is CommandButton Then
                If sLanguage = "Hebrew" Then
                    cntrl.Caption = vCaption(1)
                Else
                    cntrl.Caption = vCaption(0)
                End If
            ElseIf TypeOf cntrl Is DataGrid Then
                If sLanguage = "Hebrew" Then
                    cntrl.Caption = vCaption(1)
                Else
                    cntrl.Caption = vCaption(0)
                End If
            ElseIf TypeOf cntrl Is SSTab Then
                For i = 0 To cntrl.Tabs - 1
                    cntrl.Tab = i
                    If sLanguage = "Hebrew" Then
                        cntrl.Caption = vCaption(i * 2 + 1)
                    Else
                        cntrl.Caption = vCaption(i * 2)
                    End If
                Next
            Else
                If sLanguage = "Hebrew" Then
                    cntrl.Caption = vCaption(1)
                Else
                    cntrl.Caption = vCaption(0)
                End If
            End If
            '/// setting the right to left property according to the selected lenguage
            If TypeOf cntrl Is CommandButton Or _
            TypeOf cntrl Is Frame Or _
            TypeOf cntrl Is Label Or _
            TypeOf cntrl Is DataGrid Then
                If sLanguage = "Hebrew" Then
                    cntrl.RightToLeft = True
                Else
                    cntrl.RightToLeft = False
                End If
            End If
        End If
        

        
     Next cntrl
     
   
    Exit Sub
ErrHandler:
   g_uMsgE "m__Declarations.uSetLanguage"
End Sub


