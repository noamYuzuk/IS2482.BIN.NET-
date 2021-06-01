VERSION 5.00
Begin VB.Form frmTraysList 
   Caption         =   "Trays list"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   177
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_f_TraysList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   471
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Trays List,רשימת לולים"
   Begin VB.ListBox lstTrays 
      Height          =   6960
      IntegralHeight  =   0   'False
      Left            =   96
      TabIndex        =   6
      Top             =   120
      Width           =   3108
   End
   Begin VB.Frame Frame1 
      Height          =   7092
      Left            =   3312
      TabIndex        =   0
      Top             =   0
      Width           =   3660
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   348
         Left            =   1224
         TabIndex        =   5
         Tag             =   "Exit,יציאה"
         Top             =   3336
         Width           =   1140
      End
      Begin VB.TextBox txtTrayName 
         Height          =   360
         Left            =   288
         TabIndex        =   4
         Top             =   408
         Width           =   3108
      End
      Begin VB.CommandButton cmdDlete 
         Caption         =   "Delete"
         Height          =   348
         Left            =   1224
         TabIndex        =   3
         Tag             =   "Delete,מחק"
         Top             =   2328
         Width           =   1140
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   348
         Left            =   1224
         TabIndex        =   2
         Tag             =   "New,חדש"
         Top             =   1824
         Width           =   1140
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   348
         Left            =   1224
         TabIndex        =   1
         Tag             =   "Edit,ערוך"
         Top             =   1320
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmTraysList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oLst As New clsDataList


Private Sub Form_Load()
   
   m_oLst.uInit lstTrays, g_oTraysDB.p_rstTraysList, _
                g_oTraysDB.p_rstTraysList.Fields(0).Name, _
                txtTrayName
                
    Call uSetLanguage(Me, m_sSelectedLang)
    
End Sub


Private Sub cmdNew_Click()

   Dim sTrayName As String
   sTrayName = Trim(txtTrayName)
   If sTrayName = vbNullString Then
      g_uMsgI ("Tray name field is empty.")
      Exit Sub
   End If
   frmTrayData.p_sTrayName(True) = txtTrayName
   If frmTrayData.p_bTrayExist Then
      g_uMsgI ("Tray already exist.")
      Exit Sub
   Else
      frmTrayData.Show 1
   End If
   m_oLst.uUpdate
End Sub


Private Sub cmdEdit_Click()
   Dim sTrayName As String
   sTrayName = m_oLst.fsSelectedItem()
   If sTrayName = vbNullString Then
      g_uMsgI "No tray selected."
      Exit Sub
   End If

   frmTrayData.p_sTrayName(False) = sTrayName
   frmTrayData.Show 1
   m_oLst.uUpdate
End Sub


Private Sub cmdDlete_Click()

   Dim sTrayName As String
   sTrayName = m_oLst.fsSelectedItem()
   If sTrayName = vbNullString Then
      g_uMsgI "No tray selected."
      Exit Sub
   End If
   
   g_oTraysDB.fbDeleteTray sTrayName
   m_oLst.uUpdate
   txtTrayName = vbNullString
End Sub


'Private Sub lstTrays_Click()
'
'   txtTrayName = lstTrays.Text
'End Sub


Private Sub cmdExit_Click()
   Unload Me
End Sub


