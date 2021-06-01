VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTraysData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'References:
'Microsoft ActiveX data objects 2.7 library

'Warning: DataList control does not view sorted recordset!!!

Private Const _
mc_sTraysList = "TraysList", _
mc_sName = "Name", _
mc_sTypeTray = "TypeOfTray", _
mc_sTypeInsert = "TypeOfInsert", _
mc_sTypesTray = "TypesTray", _
mc_sTypesInsert = "TypesInsert"
'mc_sCirclesTemplate = "CirclesTemplate", _


Private _
m_oDB As New clsAccessDB
   
Private _
m_colInsertTemplates As New Collection, _
m_colTrayTemplates As New Collection
   
Private _
m_rstInsertData As ADODB.Recordset, _
m_rstTrayData As ADODB.Recordset
'm_rstCirclesData As ADODB.Recordset

Private _
m_sInsertTypeBeginning As String, _
m_sTrayTypeBeginning As String, _
m_sTrayType As String, m_sInsertType As String, _
m_sTrayName As String, m_bNewTray As Boolean

Private _
   m_grdTray As DataGrid, _
   m_grdInsert As DataGrid
Private _
   WithEvents m_cboTrayType As DataCombo, _
   WithEvents m_cboInsertType As DataCombo, _
   WithEvents m_cmdOk As CommandButton, _
   WithEvents m_cmdCancel As CommandButton
Attribute m_cboTrayType.VB_VarHelpID = -1
Attribute m_cboInsertType.VB_VarHelpID = -1
Attribute m_cmdOk.VB_VarHelpID = -1
Attribute m_cmdCancel.VB_VarHelpID = -1



'********************************************************************
'***   PROPERTIES   *************************************************


Property Get p_sCnTraysList() As String
   p_sCnTraysList = mc_sTraysList
End Property

Property Get p_sCnName() As String
   p_sCnName = mc_sName
End Property

Property Get p_sCnTypeTray() As String
   p_sCnTypeTray = mc_sTypeTray
End Property

Property Get p_sCnTypeInsert() As String
   p_sCnTypeInsert = mc_sTypeInsert
End Property

Property Get p_sCnTypesInsert() As String
   p_sCnTypesInsert = mc_sTypesInsert
End Property

Property Get p_sCnTypesTray() As String
   p_sCnTypesTray = mc_sTypesTray
End Property



Property Get p_rstTraysList() As ADODB.Recordset

   Set p_rstTraysList = m_oDB.p_rst(mc_sTraysList)
End Property


Property Get p_rstTrayTypes() As ADODB.Recordset

   Set p_rstTrayTypes = m_oDB.p_rst(mc_sTypesTray)
End Property


Property Get p_rstInsertTypes() As ADODB.Recordset

   Set p_rstInsertTypes = m_oDB.p_rst(mc_sTypesInsert)
End Property


Private Property Get p_rstInsertData( _
   Optional sTrayName As String) As ADODB.Recordset
   
   If sTrayName <> vbNullString Then
      Set m_rstInsertData = _
         m_oDB.foGetTable(fsInsertTbl(sTrayName))
   End If
   Set p_rstInsertData = m_rstInsertData
End Property


'Private Property Get p_rstCirclesData(
Private Property Get p_rstTrayData( _
   Optional sTrayName As String) As ADODB.Recordset

   If sTrayName <> vbNullString Then
'      Set m_rstCirclesData = _
'         m_oDB.foGetTable(fsTrayTbl(sTrayName))
      Set m_rstTrayData = m_oDB.foGetTable(fsTrayTbl(sTrayName))
   End If
   'Set p_rstCirclesData = m_rstCirclesData
   Set p_rstTrayData = m_rstTrayData
End Property



'********************************************************************
'***   METHODS   ****************************************************


Public Function fbInit(sPathDB As String) As Boolean
   
   If Not m_oDB.fbCnnOpen(sPathDB) Then Exit Function
   
   If Not m_oDB.fbRSAddOpenTable(mc_sTraysList, False, _
                                 mc_sName) Then Exit Function
   If Not m_oDB.fbRSAddOpenTable(mc_sTypesInsert, _
                                 True) Then Exit Function
   If Not m_oDB.fbRSAddOpenTable(mc_sTypesTray, _
                                 True) Then Exit Function
   Dim _
   sType As String, sTable As String
   
'   If Not m_oDB.fbRSAddOpenTable(mc_sCirclesTemplate, _
'                                 True) Then Exit Function
   m_oDB.p_rst(mc_sTypesTray).MoveFirst
   While Not m_oDB.p_rst(mc_sTypesTray).EOF
      sType = m_oDB.p_rst(mc_sTypesTray).Fields(mc_sTypeTray)
      sTable = fsTrayTemplTbl(sType)
      If Not m_oDB.fbRSAddOpenTable(sTable, True) Then Exit Function
      m_colTrayTemplates.Add m_oDB.p_rst(sTable), sType
      m_oDB.p_rst(mc_sTypesTray).MoveNext
   Wend
   
   m_oDB.p_rst(mc_sTypesInsert).MoveFirst
   While Not m_oDB.p_rst(mc_sTypesInsert).EOF
      sType = m_oDB.p_rst(mc_sTypesInsert).Fields(mc_sTypeInsert)
      sTable = fsInsertTemplTbl(sType)
      If Not m_oDB.fbRSAddOpenTable(sTable, True) Then Exit Function
      m_colInsertTemplates.Add m_oDB.p_rst(sTable), sType
      m_oDB.p_rst(mc_sTypesInsert).MoveNext
   Wend
                                 
   'Rem: Currently there is only tray type only.
   'The field was created for further development.
   'In case of several tray types TrayTypes table should be created
   'and all the functions and controls as for insert type.
   m_sTrayType = "Circles"
   
   fbInit = True
End Function


Public Sub uSetCrlsTrayDataFrm( _
                           grdTray As DataGrid, _
                           cboTrayType As DataCombo, _
                           cboInsertType As DataCombo, _
                           grdInsert As DataGrid, _
                           cmdOk As CommandButton, _
                           cmdCancel As CommandButton)
   Set m_grdTray = grdTray
   Set m_cboTrayType = cboTrayType
   Set m_cboInsertType = cboInsertType
   Set m_grdInsert = grdInsert
   Set m_cmdOk = cmdOk
   Set m_cmdCancel = cmdCancel
End Sub


Public Sub uInitTraysDataFrm(sTrayName As String, _
                             bNewTray As Boolean)
   m_sTrayName = sTrayName
   m_bNewTray = bNewTray
   
   '< Clear values from previous loadings of frmTrayData >
   Dim _
   rst As ADODB.Recordset, fld As ADODB.Field, _
   i As Integer
   
   For Each rst In m_colInsertTemplates
      rst.MoveFirst
      For Each fld In rst.Fields
         fld.Value = 0
      Next fld
   Next rst
   
   For Each rst In m_colTrayTemplates
      With rst
      .MoveFirst
      For Each fld In .Fields
         fld.Value = 0
      Next fld
      For i = 2 To .RecordCount
         .MoveNext
         .Delete
      Next i
      'Rem: without .MoveFirst m_grdTray.DataSource
      '     will not set.
      .MoveFirst
      End With
   Next rst
   '/>
   
   m_cboTrayType.ListField = mc_sTypeTray
   Set m_cboTrayType.RowSource = p_rstTrayTypes
   m_cboInsertType.ListField = mc_sTypeInsert
   Set m_cboInsertType.RowSource = p_rstInsertTypes
   
   If m_bNewTray Then
      p_rstTrayTypes.MoveFirst
      m_sTrayType = p_rstTrayTypes.Fields(mc_sTypeTray)
      'Rem: will rise m_cboTrayType_Change
      m_cboTrayType.Text = m_sTrayType
      
      p_rstInsertTypes.MoveFirst
      m_sInsertType = p_rstInsertTypes.Fields(mc_sTypeInsert)
      'Rem: will rise m_cboInsertType_Change
      m_cboInsertType.Text = m_sInsertType
   Else
      m_sTrayTypeBeginning = fsGetTrayType(m_sTrayName)
      'Rem: will rise m_cboInsertType_Change
      m_cboTrayType.Text = m_sTrayTypeBeginning
   
      m_sInsertTypeBeginning = fsGetInsertType(m_sTrayName)
      'Rem: will rise m_cboInsertType_Change
      m_cboInsertType.Text = m_sInsertTypeBeginning
   End If
End Sub


Public Function fbSaveNewTray() As Boolean
   'Const sDQ As String = """"
   Dim sField(2) As String, vValue(2) As Variant
   
   sField(0) = mc_sName
   sField(1) = mc_sTypeTray
   sField(2) = mc_sTypeInsert
   vValue(0) = m_sTrayName
   vValue(1) = m_sTrayType
   vValue(2) = m_sInsertType
   If Not m_oDB.fbRecordAddV(mc_sTraysList, _
                            sField(), vValue()) Then Exit Function
   m_oDB.p_rst(mc_sTraysList).Requery
'   sValue(0) = m_sTrayName
'   sValue(1) = sInsertType
'   With m_oDB.p_rst(mc_sTraysList)
'   .AddNew sField(), sValue()
'   .Fields(0) = m_sTrayName
'   .Fields(1) = sInsertType
'   .UpdateBatch
'   End With
   
   Dim sNameTbl As String, sNameTemplTbl As String
   
   sNameTbl = fsTrayTbl(m_sTrayName)
   sNameTemplTbl = fsTrayTemplTbl(m_sTrayType)
   If Not m_oDB.fbTableCopy(sNameTemplTbl, _
                            sNameTbl, "*") Then Exit Function
   If Not m_oDB.fbRSTableUpdate(sNameTemplTbl, _
                                sNameTbl) Then Exit Function
                                
   sNameTbl = fsInsertTbl(m_sTrayName)
   sNameTemplTbl = fsInsertTemplTbl(m_sInsertType)
   If Not m_oDB.fbTableCopy(sNameTemplTbl, _
                            sNameTbl, "*") Then Exit Function
   If Not m_oDB.fbRSTableUpdate(sNameTemplTbl, _
                                sNameTbl) Then Exit Function
   fbSaveNewTray = True
End Function


Public Function fbSaveEditedTray() As Boolean
   
   Dim _
   sNameTbl As String, sNameTemplTbl As String
   
   If m_sTrayTypeBeginning = m_sTrayType Then
      p_rstTrayData().UpdateBatch
   Else
      If Not m_oDB.fbRecodsUpdate(mc_sTraysList, _
         mc_sTypeTray, m_sTrayType, _
         mc_sName, m_sTrayName) Then Exit Function
      m_oDB.p_rst(mc_sTraysList).Requery
      
      sNameTbl = fsTrayTbl(m_sTrayName)
      sNameTemplTbl = fsTrayTemplTbl(m_sInsertType)
      If Not m_oDB.fbTableCopy(sNameTemplTbl, _
                               sNameTbl, "*") Then Exit Function
      If Not m_oDB.fbRSTableUpdate(sNameTemplTbl, _
                                   sNameTbl) Then Exit Function
   End If
   
   If m_sInsertTypeBeginning = m_sInsertType Then
      p_rstInsertData().UpdateBatch
   Else
      If Not m_oDB.fbRecodsUpdate(mc_sTraysList, _
         mc_sTypeInsert, m_sInsertType, _
         mc_sName, m_sTrayName) Then Exit Function
      m_oDB.p_rst(mc_sTraysList).Requery
      
      sNameTbl = fsInsertTbl(m_sTrayName)
      sNameTemplTbl = fsInsertTemplTbl(m_sInsertType)
      If Not m_oDB.fbTableCopy(sNameTemplTbl, _
                               sNameTbl, "*") Then Exit Function
      If Not m_oDB.fbRSTableUpdate(sNameTemplTbl, _
                                   sNameTbl) Then Exit Function
   End If

   fbSaveEditedTray = True
End Function


Public Function fbDeleteTray(sTrayName As String) As Boolean

   If Not m_oDB.fbTableDelete(fsTrayTbl(sTrayName)) Then _
         Exit Function
   If Not m_oDB.fbTableDelete(fsInsertTbl(sTrayName)) Then _
         Exit Function
   m_oDB.p_rst(mc_sTraysList).Delete
   m_oDB.p_rst(mc_sTraysList).Update
End Function


Public Function fbTrayExist(sTrayName As String) As Boolean

   fbTrayExist = m_oDB.foRecodsFind( _
      mc_sTraysList, mc_sName, sTrayName).RecordCount <> 0
End Function


Public Function foGetTrayTbl(sTrayName As String) As ADODB.Recordset

   Set foGetTrayTbl = m_oDB.foGetTable(fsTrayTbl(sTrayName), True)
End Function


Public Function foGetInsertTbl(sInsertType As String) As ADODB.Recordset

   Set foGetInsertTbl = m_oDB.foGetTable(fsInsertTbl(sInsertType), True)
End Function


Public Function fsGetTrayType(Optional sTrayName As String) As String

   If sTrayName <> vbNullString Then
      m_oDB.p_rst(mc_sTraysList).MoveFirst
      m_oDB.p_rst(mc_sTraysList).Find _
                  mc_sName & "=" & "'" & sTrayName & "'"
      m_sTrayTypeBeginning = _
                  m_oDB.p_rst(mc_sTraysList).Fields(mc_sTypeTray)
      m_sTrayType = m_sTrayTypeBeginning
   End If
   fsGetTrayType = m_sTrayType
End Function


Public Function fsGetInsertType(Optional sTrayName As String) As String

   If sTrayName <> vbNullString Then
      m_oDB.p_rst(mc_sTraysList).MoveFirst
      m_oDB.p_rst(mc_sTraysList).Find _
                  mc_sName & "=" & "'" & sTrayName & "'"
      m_sInsertTypeBeginning = _
                  m_oDB.p_rst(mc_sTraysList).Fields(mc_sTypeInsert)
      m_sInsertType = m_sInsertTypeBeginning
   End If
   fsGetInsertType = m_sInsertType
End Function



'********************************************************************
'***   FUNCTIONS   **************************************************


Private Function fsTrayTbl(sTrayName As String) As String
   fsTrayTbl = "Tray " & sTrayName
End Function

Private Function fsInsertTbl(sTrayName As String) As String
   fsInsertTbl = "Insert " & sTrayName
End Function

Private Function fsInsertTemplTbl(sInsertType As String) As String
      fsInsertTemplTbl = "TemplateInsert" & sInsertType
End Function

Private Function fsTrayTemplTbl(sTrayType As String) As String
      fsTrayTemplTbl = "TemplateTray" & sTrayType
End Function



'********************************************************************
'***  EVENTS  *******************************************************


Private Sub m_cboTrayType_Change()

   m_sTrayType = m_cboTrayType.Text
   If m_sTrayTypeBeginning <> m_sTrayType Or m_bNewTray Then
      Set m_grdTray.DataSource = m_colTrayTemplates(m_sTrayType)
   Else
      Set m_grdTray.DataSource = p_rstTrayData(m_sTrayName)
   End If
End Sub


Private Sub m_cboInsertType_Change()

   m_sInsertType = m_cboInsertType.Text
   If m_sInsertTypeBeginning <> m_sInsertType Or m_bNewTray Then
      Set m_grdInsert.DataSource = m_colInsertTemplates(m_sInsertType)
   Else
      Set m_grdInsert.DataSource = p_rstInsertData(m_sTrayName)
   End If
End Sub


Private Sub m_cmdCancel_Click()

   If m_bNewTray Then
   Else
      If Not p_rstTrayData Is Nothing Then _
         p_rstTrayData().CancelBatch
      If Not p_rstInsertData Is Nothing Then _
         p_rstInsertData().CancelBatch
   End If
End Sub


Private Sub m_cmdOk_Click()

   Dim b As Boolean
   
   If m_bNewTray Then
      b = fbSaveNewTray
   Else
      b = fbSaveEditedTray
   End If
   
   If Not b Then g_uMsgW "Tray data was not saved.", _
                         "clsTrayData.m_cmdOk_Click"
End Sub
