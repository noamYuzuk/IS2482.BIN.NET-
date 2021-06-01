VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTrayData 
   Caption         =   "Tray Data"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7740
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   177
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_f_TrayData.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5145
   ScaleWidth      =   7740
   StartUpPosition =   1  'CenterOwner
   Tag             =   "נתוני לולים"
   Begin MSDataListLib.DataCombo cboInsertType 
      Height          =   336
      Left            =   1344
      TabIndex        =   5
      Top             =   3744
      Width           =   1668
      _ExtentX        =   2937
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   "cboInsertType"
   End
   Begin MSDataGridLib.DataGrid grdTray 
      Height          =   2892
      Left            =   24
      TabIndex        =   3
      Tag             =   "Location Data,נתוני מיקום"
      Top             =   96
      Width           =   7692
      _ExtentX        =   13573
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   24
      TabAction       =   2
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Location Data"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   348
      Left            =   5724
      TabIndex        =   1
      Tag             =   "Cancel,ביטול"
      Top             =   3480
      Width           =   1140
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   348
      Left            =   4044
      TabIndex        =   0
      Tag             =   "OK,אישור"
      Top             =   3480
      Width           =   1140
   End
   Begin MSDataGridLib.DataGrid grdInsert 
      Height          =   804
      Left            =   24
      TabIndex        =   2
      Tag             =   "Insert Data,נתוני שימות"
      Top             =   4296
      Width           =   7692
      _ExtentX        =   13573
      _ExtentY        =   1429
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   24
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Insert data"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo cboTrayType 
      Height          =   336
      Left            =   1344
      TabIndex        =   6
      Top             =   3216
      Width           =   1668
      _ExtentX        =   2937
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   "cboTrayType"
   End
   Begin VB.Label Label1 
      Caption         =   "Tray type"
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   216
      TabIndex        =   7
      Tag             =   "Tray Type,סוג לול"
      Top             =   3240
      Width           =   1068
   End
   Begin VB.Label Label2 
      Caption         =   "Insert type"
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   216
      TabIndex        =   4
      Tag             =   "Insert Type,סוג שימה"
      Top             =   3768
      Width           =   1068
   End
End
Attribute VB_Name = "frmTrayData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mc_sIsertType As String = "InsertType"

Private m_sTrayName As String, m_bNew As Boolean, _
        m_sInsertTypeEdit As String, m_bTrayExist As Boolean


Private Sub Form_Load()
   
   If m_bNew Then
      Me.Caption = "New tray '" & m_sTrayName & "'"
   Else
      Me.Caption = "Edit tray '" & m_sTrayName & "'"
   End If
   
   Call uSetLanguage(Me, m_sSelectedLang)
   
   g_oTraysDB.uSetCrlsTrayDataFrm grdTray, cboTrayType, _
                                  cboInsertType, grdInsert, _
                                  cmdOk, cmdCancel
   g_oTraysDB.uInitTraysDataFrm m_sTrayName, m_bNew
End Sub


Private Sub Form_Unload(Cancel As Integer)

'   If m_bNew Then g_oTraysDB.uCloseRstsOfTemplates
End Sub


Private Sub cmdOk_Click()
   
   'Rem: processing is done in g_oTraysDB in m_cmdOk_Click
   Unload Me
End Sub


Private Sub cmdCancel_Click()
   
   Unload Me
End Sub


Public Property Let p_sTrayName(bNew As Boolean, s As String)

   m_sTrayName = s
   m_bNew = bNew
   If m_bNew Then
      m_bTrayExist = g_oTraysDB.fbTrayExist(m_sTrayName)
   End If
End Property


Public Property Get p_bTrayExist() As Boolean
   p_bTrayExist = m_bTrayExist
End Property
