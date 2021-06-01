VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccessDB_FieldStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public _
   sName As String, _
   sType As String, _
   iTypeSize As Integer, _
   bAutoincrement As Boolean, _
   bNotNull As Boolean, _
   bPrimaryKey As Boolean, _
   bUnique As Boolean
   

Private Sub Class_Initialize()

   sType = "SINGLE"
   bNotNull = True
End Sub

