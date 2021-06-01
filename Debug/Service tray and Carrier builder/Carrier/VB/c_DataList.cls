VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private WithEvents m_lst As ListBox, _
        WithEvents m_rst As ADODB.Recordset, _
        WithEvents m_txt As TextBox
Attribute m_lst.VB_VarHelpID = -1
Attribute m_rst.VB_VarHelpID = -1
Attribute m_txt.VB_VarHelpID = -1
        
Private m_sField As String


Public Sub uInit(lst As ListBox, rst As ADODB.Recordset, _
                 sField As String, _
                 Optional txt As TextBox)

'Warning: set lst.Multiselect = false at design time!!!

   Set m_lst = lst
   Set m_rst = rst
   m_sField = sField  '"[" & sField & "]"
   Set m_txt = txt
   uUpdate
End Sub


Public Sub uUpdate()
   
   m_lst.Clear
   If m_rst.RecordCount = 0 Then Exit Sub
   m_rst.MoveFirst
   While Not m_rst.EOF
      m_lst.AddItem m_rst.Fields(m_sField)
      m_rst.MoveNext
   Wend
End Sub


Public Sub uDelete()
   If m_rst.BOF Or m_rst.EOF Then Exit Sub
   m_rst.Delete
   m_rst.UpdateBatch
   uUpdate
End Sub


Public Function fsSelectedItem() As String

   If m_lst.SelCount = 0 Then Exit Function
   fsSelectedItem = m_lst.List(m_lst.ListIndex)
End Function


Private Sub m_lst_Click()

   If Not m_txt Is Nothing Then _
      m_txt = m_lst.List(m_lst.ListIndex)
   m_rst.AbsolutePosition = m_lst.ListIndex + 1
End Sub
