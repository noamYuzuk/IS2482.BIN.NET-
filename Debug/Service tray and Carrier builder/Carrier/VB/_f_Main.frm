VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Carrier Trays"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8010
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   177
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_f_Main.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   457
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   534
   StartUpPosition =   1  'CenterOwner
   Tag             =   " Carrier Trays,מגשי לולים"
   Begin MSComDlg.CommonDialog dlg 
      Left            =   1896
      Top             =   6108
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6756
      Left            =   48
      TabIndex        =   0
      Tag             =   "Carrier Select,בחירת מגשים,Output,מוצא"
      Top             =   48
      Width           =   7908
      _ExtentX        =   13944
      _ExtentY        =   11906
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   635
      TabCaption(0)   =   "Control"
      TabPicture(0)   =   "_f_Main.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frames"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Output"
      TabPicture(1)   =   "_f_Main.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTxtPic"
      Tab(1).Control(1)=   "txtOut"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fraTxtPic 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6300
         Left            =   -68592
         TabIndex        =   7
         Top             =   324
         Width           =   1380
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   396
            Left            =   216
            TabIndex        =   18
            Tag             =   "Save,שמור"
            Top             =   5352
            Width           =   972
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "Apply"
            Enabled         =   0   'False
            Height          =   396
            Left            =   216
            TabIndex        =   16
            Tag             =   "Apply,החל"
            Top             =   4416
            Width           =   972
         End
         Begin VB.CommandButton cmdEditPlaces 
            Caption         =   "Edit"
            Enabled         =   0   'False
            Height          =   396
            Left            =   216
            TabIndex        =   15
            Tag             =   "Edit,ערוך"
            Top             =   3864
            Width           =   972
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000005&
            Height          =   516
            Left            =   264
            ScaleHeight     =   450
            ScaleWidth      =   810
            TabIndex        =   10
            Top             =   1320
            Width           =   876
            Begin VB.TextBox txtNTray 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   48
               TabIndex        =   11
               Text            =   "1"
               Top             =   96
               Width           =   480
            End
            Begin ComCtl2.UpDown udcNTray 
               Height          =   468
               Left            =   552
               TabIndex        =   12
               Top             =   0
               Width           =   276
               _ExtentX        =   450
               _ExtentY        =   820
               _Version        =   327681
               Value           =   1
               BuddyControl    =   "txtNTray"
               BuddyDispid     =   196614
               OrigRight       =   276
               OrigBottom      =   372
               Max             =   100
               Min             =   1
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   0
               Enabled         =   0   'False
            End
         End
         Begin VB.OptionButton optPicture 
            Caption         =   "Picture"
            Height          =   264
            Left            =   240
            TabIndex        =   9
            Tag             =   "Picture,סכמה"
            Top             =   2580
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optText 
            Caption         =   "Text"
            Height          =   264
            Left            =   240
            TabIndex        =   8
            Tag             =   "Text,טקסט"
            Top             =   2940
            Width           =   1044
         End
         Begin VB.Label lables 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Carrier No"
            ForeColor       =   &H00800000&
            Height          =   312
            Index           =   2
            Left            =   144
            TabIndex        =   20
            Tag             =   "Carrier No,מס' לול"
            Top             =   1032
            Width           =   1068
         End
         Begin VB.Label lblNTrays 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   312
            TabIndex        =   19
            Top             =   600
            Width           =   732
         End
         Begin VB.Label lables 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Carriers Count"
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   1
            Left            =   30
            TabIndex        =   17
            Tag             =   "Carrier Count,כמות לולים"
            Top             =   315
            Width           =   1305
         End
      End
      Begin VB.Frame frames 
         Caption         =   " Tray "
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   10.5
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   6216
         Left            =   120
         TabIndex        =   2
         Tag             =   "Tray,סוג מגש"
         Top             =   408
         Width           =   7668
         Begin VB.CommandButton cmdProcess 
            Caption         =   "Process"
            Height          =   396
            Left            =   936
            TabIndex        =   13
            Tag             =   " Process,יצירת לול"
            Top             =   1656
            Width           =   2700
         End
         Begin VB.TextBox txtNInserts 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2784
            TabIndex        =   5
            Text            =   "60"
            Top             =   624
            Width           =   732
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit trays data"
            Height          =   396
            Left            =   912
            TabIndex        =   4
            Tag             =   "Edit Trays Data,ערוך נתוני לולים"
            Top             =   2712
            Width           =   2724
         End
         Begin VB.ListBox lstTrays 
            Height          =   5784
            IntegralHeight  =   0   'False
            Left            =   4536
            TabIndex        =   3
            Top             =   288
            Width           =   2988
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00808080&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   900
            Left            =   168
            Shape           =   4  'Rounded Rectangle
            Top             =   1416
            Width           =   4212
         End
         Begin VB.Label lables 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Inserts Count"
            ForeColor       =   &H00800000&
            Height          =   312
            Index           =   0
            Left            =   456
            TabIndex        =   6
            Tag             =   "Insert Count, מספר שימות"
            Top             =   672
            Width           =   1956
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00808080&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   900
            Left            =   168
            Shape           =   4  'Rounded Rectangle
            Top             =   2472
            Width           =   4212
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00808080&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   900
            Left            =   168
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   4212
         End
      End
      Begin VB.TextBox txtOut 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6144
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   492
         Width           =   6195
      End
   End
   Begin VB.PictureBox picOut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6144
      Left            =   168
      ScaleHeight     =   408
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   411
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   540
      Width           =   6195
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileBackupDB 
         Caption         =   "&Backup DB"
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSet 
      Caption         =   "&Settings"
      Begin VB.Menu mnuLanguage 
         Caption         =   "Language Select"
         Begin VB.Menu mnuEnglish 
            Caption         =   "English"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuHebrew 
            Caption         =   "Hebrew"
         End
      End
      Begin VB.Menu mnuSetDir 
         Caption         =   "Store folder"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const _
   mc_sSecFolders As String = "Folders", _
   mc_sItmStoreFld As String = "Store"


Private _
   m_sPathDB As String, m_sStoreFld As String
   
Private _
   m_bEditTray As Boolean
   
Private _
   m_oLst As New clsDataList, _
   m_oOrder As New clsOrder



'********************************************************************
'***   EVENTS   *****************************************************

Private Sub Form_Load()

   m_sPathDB = App.path & "\TraysStahli.mdb"
   
   With g_oTraysDB
   If Not .fbInit(m_sPathDB) Then
   End If
   
   m_oLst.uInit lstTrays, .p_rstTraysList, _
               .p_rstTraysList.Fields(0).Name
   End With
   If Not m_oOrder.fbInit(m_sPathDB, m_oLst, picOut, _
                         txtOut, txtNInserts, lblNTrays) Then
   End If
   
   m_sStoreFld = fsGetIniItem2(mc_sSecFolders, mc_sItmStoreFld)
   
   m_sSelectedLang = fsGetIniItem2(mc_sSecSettings, mc_sItmStoreLang)
   
    If m_sSelectedLang = "Hebrew" Then
        mnuHebrew.Checked = True   ' the click event toggles the language selection (this is way now it is the opposite language)
        mnuEnglish.Checked = False
    Else
        mnuHebrew.Checked = False
        mnuEnglish.Checked = True
    End If
    Call uSetLanguage(Me, m_sSelectedLang)
'    Call mnuHebrew_Click
   SSTab1.Tab = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
'   If giChildFormFl = 1 Then Unload frmTraySchem
   'End
End Sub


Private Sub mnuAbout_Click()
Dim str As String
    
    str = String(Len(App.FileDescription), "-")
    MsgBox App.ProductName & vbCrLf & _
            str & vbCrLf & _
            "Version: " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
            str & vbCrLf & _
            App.FileDescription & vbCrLf & _
             vbCrLf & _
             vbCrLf & _
            "*" & str & "*" _
            & vbCrLf & _
            "*          " & App.LegalCopyright & "                       *" & _
            vbCrLf & _
            "*" & str & "*"

End Sub

Private Sub mnuEnglish_Click()
Dim sLanguage As String
    mnuEnglish.Checked = Not mnuEnglish.Checked
    mnuHebrew.Checked = Not mnuEnglish.Checked
    sLanguage = IIf(mnuHebrew.Checked, "Hebrew", "English")
    
    Call uSetLanguage(Me, sLanguage)
    Call uSetLanguage(frmTrayData, sLanguage)
    Call uSetLanguage(frmTraysList, sLanguage)
    
    '/// save selected language to the setting sectio no the app. ini file
    fbSetIniItem2 mc_sSecSettings, mc_sItmStoreLang, m_sSelectedLang
End Sub
Private Sub mnuHebrew_Click()
Dim sLanguage As String
    mnuHebrew.Checked = Not mnuHebrew.Checked
    mnuEnglish.Checked = Not mnuHebrew.Checked
    sLanguage = IIf(mnuHebrew.Checked, "Hebrew", "English")
    
    Call uSetLanguage(Me, sLanguage)
    Call uSetLanguage(frmTrayData, sLanguage)
    Call uSetLanguage(frmTraysList, sLanguage)

    '/// save selected language to the setting sectio no the app. ini file
    fbSetIniItem2 mc_sSecSettings, mc_sItmStoreLang, m_sSelectedLang
End Sub
Private Sub mnuFileBackupDB_Click()
   Dim sPathBU As String
   sPathBU = fsChangeExtension(m_sPathDB, "bak")
   If fbCopyFile(m_sPathDB, sPathBU) Then
      g_uMsgI "Saved to " & sPathBU
   End If
End Sub


Private Sub mnuFileExit_Click()
   Unload Me
End Sub




Private Sub mnuSetDir_Click()

   Dim s As String
   s = fsBrowseForFolder(m_sStoreFld, _
                        "Directory for trays data saving", Me)
   If s <> vbNullString Then
      m_sStoreFld = s
      fbSetIniItem2 mc_sSecFolders, mc_sItmStoreFld, m_sStoreFld
   End If
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
   'Rem: there is a problem to draw on PicBox within SSTab.
   'So picOut is forced to be placed out of SSTab
   If SSTab1.Tab = 1 Then
      If optPicture.Value Then picOut.ZOrder
   Else
      picOut.ZOrder 1
   End If
End Sub


Private Sub cmdProcess_Click()

   If m_oOrder.fbCalc Then
      cmdSave.Enabled = True
      udcNTray.Enabled = True
      txtNTray.Enabled = True
      cmdEditPlaces.Enabled = True
      cmdApply.Enabled = False
      udcNTray.Max = m_oOrder.p_iNTrays
      udcNTray.Value = 1   'This will raise udcNTray_Change
      SSTab1.Tab = 1
      optPicture.Value = True
      optText.Enabled = True
      picOut.ZOrder
      DoEvents
   Else
      cmdSave.Enabled = False
      udcNTray.Enabled = False
      cmdEditPlaces.Enabled = False
      cmdApply.Enabled = False
   End If
End Sub


Private Sub cmdSave_Click()

   Dim sFilePath As String
   
   sFilePath = fsChooseFilePath( _
      dlg, "Save file", _
      "Text Files (*.txt)|*.txt|All Files (*.*)|*.*", _
      False, m_sStoreFld & "\" & m_oOrder.p_sTrayName & ".txt")
      
   If sFilePath <> vbNullString Then
      uStringToFile m_oOrder.oTrays.fsInsertrsPos, sFilePath
'''      uStringToFile txtOut.Text, sFilePath
      '/// added to prevent another DB tables in stahli project
      uStringToFile m_oOrder.d_InsDiameter, Replace(sFilePath, ".txt", ".phw")
   End If
End Sub


Private Sub optPicture_Click()
   picOut.ZOrder
End Sub


Private Sub optText_Click()
   picOut.ZOrder 1
End Sub


Private Sub lstTrays_Click()
   Me.Caption = " Stahli Trays  [" & m_oLst.fsSelectedItem & "]"
End Sub


Private Sub cmdEdit_Click()
   
   frmTraysList.Show 1
   m_oLst.uUpdate
End Sub


Private Sub txtNTray_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
      Dim i As Integer
      i = Val(txtNTray)
      If i > udcNTray.Max Then
         udcNTray.Value = udcNTray.Max
         txtNTray = udcNTray.Max
      Else
         udcNTray.Value = i
      End If
      
      txtNTray.ForeColor = vbBlack
      optPicture.Value = True
   Else
      txtNTray.ForeColor = vbRed
   End If
End Sub


Private Sub udcNTray_Change()

   m_oOrder.uDraw Val(txtNTray) - 1
   optPicture.Value = True
End Sub


Private Sub cmdEditPlaces_Click()

   m_bEditTray = True
   uSetCrlsOnEdit m_bEditTray
End Sub


Private Sub cmdApply_Click()

   If Not m_oOrder.fbApplyEdit(Val(txtNTray) - 1) Then _
      Exit Sub
   
   m_bEditTray = False
   uSetCrlsOnEdit m_bEditTray
End Sub

Private Sub picOut_MouseUp(Button As Integer, Shift As Integer, _
                           x As Single, y As Single)
   If m_bEditTray Then _
      m_oOrder.uTraysEdit Val(txtNTray) - 1, x, y
End Sub


'********************************************************************
'***   FUNCTIONS   **************************************************

Private Sub uSetCrlsOnEdit(bEdit As Boolean)

   If bEdit Then
      optPicture.Value = True
   Else
      optText.Value = True
   End If
   cmdApply.Enabled = bEdit
   cmdEditPlaces.Enabled = Not bEdit
   optText.Enabled = Not bEdit
   udcNTray.Enabled = Not bEdit
   txtNTray.Enabled = Not bEdit
End Sub




