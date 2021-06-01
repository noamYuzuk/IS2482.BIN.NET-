VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "W"
Option Explicit

Private Const _
mc_sCircles As String = "Circles", _
mc_sSectors As String = "Sectors"
'mc_sTraysList = "TraysList", _

Private _
m_oDB As New clsAccessDB, _
m_oLst As clsDataList

Private _
m_pic As PictureBox, m_txtOut As TextBox, _
m_txtNInsInOrder As TextBox, m_lblNTrays As Label

Private _
m_sTrayName As String, m_sTrayType As String
', _
'm_sNTrays As Integer

Private _
m_oTrays As Object
Private _
m_oCircles As clsTraySectors

', _
'm_oCircles As clsTrayCircles, _
'm_oSectors As clsTraySectors



'********************************************************************
'***   PROPERTIES   *************************************************

Property Get p_sTrayName() As String

   p_sTrayName = m_sTrayName
End Property

Property Get d_InsDiameter() As Double

   d_InsDiameter = m_oCircles.d_Diameter
End Property

Property Get p_iNTrays() As Integer

   'p_iNTrays = m_oCircles.p_iNTraysInOrder
   p_iNTrays = m_oTrays.p_iNTraysInOrder
End Property

'/// added to allow access to all insert coordinates due to TextBox maximal string limitation
Property Get oTrays() As Object

   Set oTrays = m_oTrays
End Property

'********************************************************************
'***   METHODS   ****************************************************

Public Function fbInit(sPathDB As String, oLst As clsDataList, _
                       pic As PictureBox, txtOut As TextBox, _
                       txtNInsInOrder As TextBox, _
                       lblNTrays As Label) As Boolean
   
   If Not m_oDB.fbCnnOpen(sPathDB) Then Exit Function
   If Not m_oDB.fbRSAddOpenTable(g_oTraysDB.p_sCnTraysList) Then _
                                        Exit Function
   Set m_oLst = oLst
   Set m_pic = pic
   Set m_txtOut = txtOut
   Set m_txtNInsInOrder = txtNInsInOrder
   Set m_lblNTrays = lblNTrays
   fbInit = True
End Function


Public Function fbCalc() As Boolean
   
   m_sTrayName = m_oLst.fsSelectedItem
   If m_sTrayName = vbNullString Then
      g_uMsgI "No tray selected."
      Exit Function
   End If
   
   Dim iNIns As Integer
   iNIns = Val(m_txtNInsInOrder)
   If iNIns < 6 Then
      g_uMsgI "Inserts number is to small."
      Exit Function
   End If
   
   With g_oTraysDB
   m_sTrayType = m_oDB.fsGetValue(.p_sCnTraysList, .p_sCnName, _
                                   m_sTrayName, .p_sCnTypeTray)
   End With
   
   Select Case m_sTrayType
   Case mc_sCircles
   '---------------------------------------------------------------------------------------------------------
      Set m_oCircles = New clsTraySectors   '/// was commented - uncommented to allow search diameter access
      m_oCircles.uPictureSet m_pic
      If m_oCircles.fbCalc(m_sTrayName, iNIns) Then
         m_txtOut = m_oCircles.fsInsertrsPos
         m_lblNTrays = m_oCircles.p_iNTraysInOrder
         fbCalc = True
      End If
   '---------------------------------------------------------------------------------------------------------
      Set m_oTrays = New clsTrayCircles
   Case mc_sSectors
      Set m_oTrays = New clsTraySectors
   End Select
   
   m_oTrays.uPictureSet m_pic
   
   If m_oTrays.fbCalc(m_sTrayName, iNIns) Then
      m_txtOut = m_oTrays.fsInsertrsPos
      m_lblNTrays = m_oTrays.p_iNTraysInOrder
      fbCalc = True
   End If
End Function


Public Sub uDraw(iITray As Integer)
   
'   Select Case m_sTrayType
'   Case "Circles"
'      m_oCircles.uDraw iITray
'   Case "Circles"
'      m_oCircles.uDraw iITray
'   End Select
   m_oTrays.uDraw iITray
End Sub


Public Sub uTraysEdit(iITray As Integer, snX As Single, snY As Single)

'   Select Case m_sTrayType
'   Case "Circles"
'      m_oCircles.uEdit iITray, snX, snY
'   End Select
   
   m_oTrays.uEdit iITray, snX, snY
End Sub


Public Function fbApplyEdit(iITray As Integer) As Boolean

   'If Not m_oCircles.fbApplyEdit(iITray) Then Exit Function
   If Not m_oTrays.fbApplyEdit(iITray) Then Exit Function
   m_txtOut = m_oTrays.fsInsertrsPos
   fbApplyEdit = True
End Function


'********************************************************************
'***   FUNCTIONS   **************************************************

Private Function fsTrayTbl(sTrayName As String) As String
   fsTrayTbl = "Tray" & sTrayName
End Function

Private Function fsInsertTbl(sTrayName As String) As String
   fsInsertTbl = "Insert" & sTrayName
End Function

