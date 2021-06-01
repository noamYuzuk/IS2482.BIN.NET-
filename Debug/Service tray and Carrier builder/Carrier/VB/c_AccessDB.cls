VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccessDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================================================
'   USING: m_Msg.bas, c_AccessDB_FieldStruct.bas
'====================================================================

Option Explicit

Private _
m_cn As New ADODB.Connection ', m_rst As New ADODB.Recordset

Private _
m_iDW As Integer, m_iDH As Integer, m_sSQL As String

Private m_colRst As New Collection


'====================================================================
'   DESTRUCTOR
'====================================================================

Private Sub Class_Terminate()

   Dim rst As ADODB.Recordset
   On Error GoTo ErrHandler
   For Each rst In m_colRst
      If rst.State = adStateOpen Then
         rst.ActiveConnection = Nothing
         If rst.RecordCount <> 0 Then rst.MoveFirst
         If Not (rst.BOF Or rst.EOF) Then rst.Update
         rst.Close
      End If
      'Rem: Next line seems does not work (see Locals Window)
      'Set rst = Nothing
   Next rst
   
   If Not m_cn Is Nothing Then
      m_cn.Close
      Set m_cn = Nothing
   End If
   Exit Sub
ErrHandler:
   g_uMsgE "clsAccessDB.Class_Terminate"
End Sub


'====================================================================
'   PROPERTIES
'====================================================================

Public Property Get p_rst(sKey As String) As ADODB.Recordset
   On Error Resume Next
   Set p_rst = m_colRst.Item(sKey)
End Property


'====================================================================
'   METHODS
'====================================================================

Public Function fbCnnOpen(sDBPath As String) As Boolean
   
   If m_cn.State = adStateOpen Then
      g_uMsgW "Connection already opened.", _
              "clsAccessDB.fbCnnOpen"
      Exit Function
   End If
   
   Dim _
   sProvider As String, sDataSource As String
   
   sProvider = "Provider = Microsoft.Jet.OLEDB.4.0; "
   sDataSource = "Data Source = " & sDBPath
   m_cn.ConnectionString = sProvider & sDataSource
   On Error GoTo ErrHandler
   m_cn.CursorLocation = adUseClient
   m_cn.Open
   fbCnnOpen = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbCnnOpen"
End Function


Public Function fbCnnClose(sDBPath As String) As Boolean
   
   On Error GoTo ErrHandler
   If m_cn.State <> adStateClosed Then m_cn.Close
   If m_cn.State = adStateClosed Then fbCnnClose = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbCnnClose"
End Function


Public Function fbCnnReopen() As Boolean
   
   On Error GoTo ErrHandler
   If m_cn.State <> adStateClosed Then m_cn.Close
   m_cn.Open

   Dim rs As ADODB.Recordset, b As Boolean
   For Each rs In m_colRst
      'Set rs = m_colRst.Item(sKey)
      rs.ActiveConnection = m_cn
      rs.Open
   Next rs
   fbCnnReopen = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbCnnReopen"
End Function


Public Function fbRSAddOpen(sTable As String, _
                            sFields As String, _
                            Optional sKey As String, _
                            Optional sSortField As String) As Boolean
   Dim sKeyTmp As String
   
   On Error GoTo ErrHandler
   If sKey = vbNullString Then
      sKeyTmp = sTable
   Else
      sKeyTmp = sKey
   End If
   m_colRst.Add foFieldsGet(sTable, sFields, sSortField), sKeyTmp
              
   fbRSAddOpen = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbRSAddOpen"
End Function


Public Function fbRSAddOpenTable(sTable As String, _
                                 Optional bDisconnect As Boolean, _
                                 Optional sSortField As String) As Boolean
   Dim rst As New ADODB.Recordset
   
   On Error GoTo ErrHandler
   With rst
   '.CursorType = adOpenDynamic  'this does not work beacause
                                 'm_cn.CursorLocation = adUseClient
   .LockType = adLockOptimistic
   .Open sTable, m_cn, , , adCmdTable
   .Sort = sSortField
   End With

   If bDisconnect Then rst.ActiveConnection = Nothing
   m_colRst.Add rst, sTable
              
   fbRSAddOpenTable = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbRSAddOpenTable"
End Function


Public Function fbRSRemove(sKey As String) As Boolean

   If Not fbRSClose(sKey) Then Exit Function
   
   On Error GoTo ErrHandler
   m_colRst.Remove sKey

   fbRSRemove = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbRSRemove"
End Function


Public Function fbRSClose(sKey As String) As Boolean

   On Error GoTo ErrHandler
   
   Dim rs As ADODB.Recordset
   Set rs = m_colRst.Item(sKey)
   With rs
   If .State <> adStateClosed Then
      If Not (.EOF Or .BOF) Then .Update
      'Rem:
      'Update does not work if rs in EOF of BOF.
      'If not updated Close method will fail.
      'Among others Update copy Value to Initial value of rs.
      .Close
      Set rs = Nothing
   End If
   End With

   fbRSClose = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbRSClose"
End Function


Public Function fbRSTableUpdate(sKeyRS As String, _
                                sNameTbl As String) As Boolean
   Dim rstFrom As ADODB.Recordset, _
       rstTo As New ADODB.Recordset, _
       i As Integer
   
   On Error GoTo ErrHandler
   With rstTo
   .LockType = adLockOptimistic
   .Open "[" & sNameTbl & "]", m_cn, , , adCmdTable
   End With
   If Not fbRSCopy(m_colRst(sKeyRS), rstTo) Then Exit Function
                 
   fbRSTableUpdate = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbRSTableUpdate"
End Function


Public Function fbRSCopy(rstFrom As ADODB.Recordset, _
                         rstTo As ADODB.Recordset) As Boolean
   Dim i As Integer
   
   On Error GoTo ErrHandler
   
   rstFrom.MoveFirst
   rstTo.MoveFirst
   rstTo.MovePrevious
   Do Until rstFrom.EOF
      If rstFrom.AbsolutePosition > rstTo.RecordCount Then
         rstTo.AddNew
      Else
         rstTo.MoveNext
      End If
      For i = 0 To rstFrom.Fields.Count - 1
         rstTo.Fields(i) = rstFrom.Fields(i)
      Next i
      rstFrom.MoveNext
   Loop
   rstTo.UpdateBatch
                 
   fbRSCopy = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbRSCopy"
End Function


Public Function foGetTable(sTable As String, _
                             Optional bDisconnect As Boolean, _
                             Optional sSortField As String) _
                             As ADODB.Recordset
   Dim rst As New ADODB.Recordset
   
   On Error GoTo ErrHandler
   With rst
   '.CursorType = adOpenDynamic  'this does not work beacause
                                 'm_cn.CursorLocation = adUseClient
   .LockType = adLockOptimistic
   .Open "[" & sTable & "]", m_cn, , , adCmdTable
   .Sort = sSortField
   End With

   If bDisconnect Then rst.ActiveConnection = Nothing
              
   Set foGetTable = rst
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.foGetTable"
End Function


Public Function foFieldsGet(sTable As String, _
                            sFields As String, _
                            Optional sSortField As String) _
                            As ADODB.Recordset
   'sFields -
   '  string with comma delimited names of fields.
   'Rem:
   'If sFields contain *,all fields will be selected.
   'If field name consist of space delimited words then
   'it must be included in square brackets.
   Dim _
   sSQL As String, sVal As String, lRecords As Long, _
   rst As New ADODB.Recordset
   
   On Error GoTo ErrHandler
      
   sSQL = "SELECT " & sFields & " FROM " & "[" & sTable & "]"
   'Debug.Print sSQL
   Set rst = m_cn.Execute(sSQL, lRecords)
   rst.LockType = adLockOptimistic
   rst.Sort = sSortField
   Set foFieldsGet = rst
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.foFieldsGet"
End Function


Public Function fbFieldAdd(sTable As String, _
                           sField As String, _
                           sFieldType As String, _
                           sDefault As String) As Boolean
   'sFieldType may be:
   '   BOOLEAN, BYTE, CURRENCY, DATETIME, DOUBLE, LONG,
   '   LONGTEXT(Access 'Memo', VB 'String'), SINGLE,
   '   SHORT(Integer in VB), TEXT(VB 'String')
   
   On Error GoTo ErrHandler
   
   '--- Close recordset ---
   If m_colRst.Item(sTable).State = adStateOpen Then
      m_colRst.Item(sTable).Close
   End If
   
   '--- Add field to database ---
   m_cn.Execute "ALTER TABLE " & "[" & sTable & "] " & _
                "ADD COLUMN " & "[" & sField & "] " & _
                sFieldType & " NOT NULL DEFAULT " & sDefault
   '--- Rem: If another fields alredy filled with data
   '        the added column sells will be empty in spite of
   '        setting of default value and NOT NULL constraint.
   '        The next statement will fill all the collumn
   m_cn.Execute ("UPDATE " & "[" & sTable & "] " & "SET " & _
                 "[" & sField & "] " & " = " & sDefault)
   '--- Rem: For fill a field in the sertain row use
   '        UPDATE table SET column1_name = new_value
   '               WHERE column2_name = some_value
   fbFieldAdd = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbFieldAdd"
End Function


Public Function fbFieldDelete(sTable As String, _
                              sField As String) As Boolean

   On Error GoTo ErrHandler
   
   '--- Close recordset ---
   If m_colRst.Item(sTable).State = adStateOpen Then
      m_colRst.Item(sTable).Close
   End If
   
   '--- Delete field from database ---
   m_cn.Execute "ALTER TABLE " & "[" & sTable & "] " & _
                "DROP COLUMN " & "[" & sField & "]"
   
   fbFieldDelete = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbFieldDelete"
End Function


Public Function flFieldsCopy(sTableFrom As String, _
                             sTableTo As String, _
                             sFieldsFrom() As String, _
                             sFieldsTo() As String) As Long
   'Function return number of affected records.
   'Rem: designated fiels must already exist
   '     in destination table.
   
   On Error GoTo ErrHandler
   
   '--- Create the table.
   Dim _
   iIUFields As Integer, i As Integer, sSQL As String
   
   iIUFields = UBound(sFieldsTo())
   
   sSQL = "INSERT INTO " & "[" & sTableTo & "] ("
   
   For i = 0 To iIUFields
      sSQL = sSQL & "[" & sFieldsTo(i) & "]"
      If i < iIUFields Then
         sSQL = sSQL & ", "
      Else
         sSQL = sSQL & ")"
      End If
   Next i
   
   sSQL = sSQL & " SELECT "
   
   For i = 0 To iIUFields
      sSQL = sSQL & "[" & sFieldsFrom(i) & "]"
      If i < iIUFields Then sSQL = sSQL & ", "
   Next i
   
   sSQL = sSQL & " FROM " & "[" & sTableFrom & "]"
   
   'Debug.Print sSQL
   
   Dim lNRec As Long
   m_cn.Execute sSQL, lNRec
   
   flFieldsCopy = lNRec
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.flFieldsCopy"
End Function


Public Function fbFieldsUpdate(sTable As String, _
                               sFields As String, _
                               sValues As String) As Boolean
   'sFields -
   '  string with comma delimited names of fields.
   'sValues - string with "|" delimited values.
   
   Dim _
   sFieldsAr() As String, sValsAr() As String, _
   iIUFields As Integer, i As Integer
   Dim _
   sSQL As String, vVal As Variant, sVal As String, _
   sFieldsTmp As String
   
   On Error GoTo ErrHandler
   
   sFieldsAr() = Split(sFields, ",")
   sValsAr() = Split(sValues, "|")
   iIUFields = UBound(sFieldsAr)
   
   sSQL = "UPDATE " & "[" & sTable & "]" & " SET "
   
   For i = 0 To iIUFields
      sSQL = sSQL & "[" & sFieldsAr(i) & "] = "
      
      vVal = sValsAr(i)
      If IsNumeric(vVal) Then
         sVal = sValsAr(i)
      Else
         sVal = """" & CStr(sValsAr(i)) & """"
      End If
      sSQL = sSQL & sVal
      
      If i < iIUFields Then sSQL = sSQL & ", "
   Next i
   
   'Debug.Print sSQL
   m_cn.Execute sSQL
   fbFieldsUpdate = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.bRecodsUpdate"
End Function


'Rem: deprecated - use fbRecordAddV
Public Function fbRecordAdd(sTable As String, _
                            sFields() As String, _
                            sValues() As String) As Boolean
   'If value item corresponding to the text field in DB
   'it must be enclosed in double quotes or in quotes if
   'it have no several words.
   
   On Error GoTo ErrHandler
   
   Dim _
   iIUFields As Integer, i As Integer, sSQL As String
   
   iIUFields = UBound(sFields())
   
   sSQL = "INSERT INTO " & "[" & sTable & "] ("
   
   For i = 0 To iIUFields
      sSQL = sSQL & "[" & sFields(i) & "]"
      
      If i < iIUFields Then
         sSQL = sSQL & ", "
      Else
         sSQL = sSQL & ")"
      End If
   Next i
         
   sSQL = sSQL & " VALUES ("

   For i = 0 To iIUFields
      'sSQL = sSQL & "[" & sValues(i) & "]"
      sSQL = sSQL & sValues(i)
      
      If i < iIUFields Then
         sSQL = sSQL & ", "
      Else
         sSQL = sSQL & ")"
      End If
   Next i
       
   'Debug.Print sSQL
   m_cn.Execute sSQL
                               
   fbRecordAdd = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbRecordAdd"
End Function


Public Function fbRecordAddV(sTable As String, _
                             sFields() As String, _
                             vValues() As Variant) As Boolean
   On Error GoTo ErrHandler
   
   Dim _
   iIUFields As Integer, i As Integer, sSQL As String
   
   iIUFields = UBound(sFields())
   
   sSQL = "INSERT INTO " & "[" & sTable & "] ("
   
   For i = 0 To iIUFields
      sSQL = sSQL & "[" & sFields(i) & "]"
      
      If i < iIUFields Then
         sSQL = sSQL & ", "
      Else
         sSQL = sSQL & ")"
      End If
   Next i
         
   sSQL = sSQL & " VALUES ("

   For i = 0 To iIUFields
      If IsNumeric(vValues(i)) Then
         sSQL = sSQL & vValues(i)
      Else
         sSQL = sSQL & """" & CStr(vValues(i)) & """"
      End If
      
      If i < iIUFields Then
         sSQL = sSQL & ", "
      Else
         sSQL = sSQL & ")"
      End If
   Next i
       
   'Debug.Print sSQL
   m_cn.Execute sSQL
                               
   fbRecordAddV = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbRecordAddV"
End Function


Public Function foRecodsFind(sTable As String, _
                             sField As String, _
                             vValue As Variant, _
                             Optional sFields As String) _
                             As ADODB.Recordset
   'sFields -
   '  string with comma delimited names of fields.
   'If field name consist of space delimited words then
   'it must be included in square brackets.
   
   Dim sSQL As String, sVal As String, sFieldsTmp As String
   
   If IsNumeric(vValue) Then
      sVal = vValue
   Else
      sVal = """" & CStr(vValue) & """"
   End If
      
   If sFields = vbNullString Then
      sFieldsTmp = "*"
   Else
      sFieldsTmp = sFields
   End If
   
   sSQL = "SELECT " & sFieldsTmp & _
          " FROM " & "[" & sTable & "]" & _
          " WHERE " & "[" & sField & "] = "  'sVal
   
   'Debug.Print sSQL & sVal
   On Error Resume Next
   Set foRecodsFind = m_cn.Execute(sSQL & sVal)
   If Err.Number = -2147217913 Then
      Err.Clear
      sVal = """" & CStr(vValue) & """"
      Set foRecodsFind = m_cn.Execute(sSQL & sVal)
   End If
   If Err.Number <> 0 Then g_uMsgE "clsAccessDB.foRecodsFind"
End Function


Public Function fsGetValue(sTable As String, _
                           sFieldFind As String, _
                           vValueFind As Variant, _
                           sFieldGet As String) _
                           As String
   Dim rst As ADODB.Recordset
   
   Set rst = foRecodsFind(sTable, sFieldFind, _
                          vValueFind, sFieldGet)
   fsGetValue = rst.Fields(0)
End Function


Public Function fbRecodsUpdate(sTable As String, _
                              sFields As String, _
                              sValues As String, _
                              sWhereField As String, _
                              vWhereValue As Variant) _
                              As Boolean
   'sFields -
   '  string with comma delimited names of fields.
   'sValues - string with "|" delimited names of fields.
   
   Dim _
   sFieldsAr() As String, sValsAr() As String, _
   iIUFields As Integer, i As Integer
   Dim _
   sSQL As String, vVal As Variant, sVal As String, _
   sFieldsTmp As String
   
   On Error GoTo ErrHandler
   
   sFieldsAr() = Split(sFields, ",")
   sValsAr() = Split(sValues, "|")
   iIUFields = UBound(sFieldsAr)
   
   sSQL = "UPDATE " & "[" & sTable & "]" & " SET "
   
   For i = 0 To iIUFields
      sSQL = sSQL & "[" & sFieldsAr(i) & "] = "
      
      vVal = sValsAr(i)
      If IsNumeric(vVal) Then
         sVal = sValsAr(i)
      Else
         sVal = """" & CStr(sValsAr(i)) & """"
      End If
      sSQL = sSQL & sVal
      
      If i < iIUFields Then sSQL = sSQL & ", "
   Next i
   
   sSQL = sSQL & _
      " WHERE " & "[" & sWhereField & "] = "
   'Debug.Print sSQL & vWhereValue
   
   On Error Resume Next
   m_cn.Execute sSQL & vWhereValue
   If Err.Number <> 0 Then
     Err.Clear
     'Debug.Print sSQL & """" & vWhereValue & """"
     m_cn.Execute sSQL & _
               """" & vWhereValue & """"
   End If
   If Err.Number <> 0 Then GoTo ErrHandler

   fbRecodsUpdate = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.bRecodsUpdate"
End Function


Public Function fbRecodsDelete(sTable As String, _
                               sWhereField As String, _
                               vWhereValue As Variant) As Boolean
   Dim _
   sSQL As String, vVal As Variant, sVal As String
   
   On Error GoTo ErrHandler
   
   sSQL = "DELETE FROM " & "[" & sTable & "]" & _
          " WHERE " & "[" & sWhereField & "] = "
   'Debug.Print sSQL
   
   On Error Resume Next
   m_cn.Execute sSQL & vWhereValue
   If Err.Number = -2147217913 Then
      Err.Clear
      m_cn.Execute sSQL & _
               """" & vWhereValue & """"
   End If
   If Err.Number <> 0 Then
      g_uMsgE "clsAccessDB.fbRecodsDelete"
      Exit Function
   End If

   fbRecodsDelete = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbRecodsDelete"
End Function


Public Function fiTableCreate(sTable As String, _
                           oField() As clsAccessDB_FieldStruct) _
                           As Integer
   'Function return number of fields in created table.
   
   '--- Drop the table if it already exists.
   On Error Resume Next
   m_cn.Execute "DROP TABLE " & "[" & sTable & "]"
   On Error GoTo ErrHandler
   
   '--- Create the table.
   Dim _
   iIUFields As Integer, i As Integer, sSQL As String
   
   iIUFields = UBound(oField())
   
   sSQL = "CREATE TABLE " & "[" & sTable & "] ("
   
   For i = 0 To iIUFields
      With oField(i)
      sSQL = sSQL & "[" & .sName & "] "
      If .bAutoincrement Then
         sSQL = sSQL & "AUTOINCREMENT"
      Else
         sSQL = sSQL & .sType
         If .iTypeSize <> 0 Then _
            sSQL = sSQL & "(" & .iTypeSize & ")"
         If .bNotNull Then _
            sSQL = sSQL & " NOT NULL"
      End If
      If .bPrimaryKey Then
         sSQL = sSQL & " CONSTRAINT " & _
                Replace(sTable, " ", "_") & "_" & "PK " & _
                "PRIMARY KEY"
      ElseIf .bUnique Then
         sSQL = sSQL & " CONSTRAINT " & _
                Replace(sTable, " ", "_") & "_" & _
                "UN" & i & " " & "UNIQUE"
      End If
      End With
      If i < iIUFields Then
         sSQL = sSQL & ", "
      Else
         sSQL = sSQL & ")"
      End If
   Next i
   
   'Debug.Print sSQL
   m_cn.Execute sSQL
   
   fiTableCreate = iIUFields + 1
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fiTableCreate"
End Function


Public Function fbTableCopy(sTableFrom As String, _
                            sTableTo As String, _
                            sFields As String) As Boolean
   'Rem: Creates sTableTo and copy the fields to it.
   '     Use "*" for copy all the fields.
   
   '--- Drop the table if it already exists.
   On Error Resume Next
   m_cn.Execute "DROP TABLE " & "[" & sTableTo & "]"
   On Error GoTo ErrHandler
   
   Dim _
   sFieldsAr() As String, iIUFields As Integer, _
   i As Integer, sSQL As String
   
   '--- Create the table.
   If sFields = "*" Then
         sSQL = "SELECT *"
   Else
      sFieldsAr() = Split(sFields, ",")
      iIUFields = UBound(sFieldsAr())
      
      sSQL = "SELECT "
      
      For i = 0 To iIUFields
         sSQL = sSQL & "[" & sFieldsAr(i) & "]"
         If i < iIUFields Then sSQL = sSQL & ", "
      Next i
   End If
   
   sSQL = sSQL & " INTO " & "[" & sTableTo & "]" & _
                 " FROM " & "[" & sTableFrom & "]"
   'Debug.Print sSQL
   m_cn.Execute sSQL
   
   fbTableCopy = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbTableCopy"
End Function


Public Function fbTableDelete(sTable As String) As Boolean

   On Error GoTo ErrHandler
   m_cn.Execute "DROP TABLE " & "[" & sTable & "]"
   
   fbTableDelete = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbTableDelete"
End Function


Public Function fbTableRename(sTableFrom As String, _
                              sTableTo As String) As Boolean
   
   If Not fbTableCopy(sTableFrom, sTableTo, _
                      sTableFrom & ".*") Then Exit Function
   On Error GoTo ErrHandler
   m_cn.Execute "DROP TABLE " & "[" & sTableFrom & "]"
   
   fbTableRename = True
   Exit Function
ErrHandler:
   g_uMsgE "clsAccessDB.fbTableRename"
End Function

