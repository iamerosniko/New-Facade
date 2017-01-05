VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DL_DA_Generic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'--------------------------------------------------NOTES--------------------------------------------------------------
'
'Project: Facade Framework 1.0
'Date Created: 08/04/2016
'Created By: Eros Niko Cas Alvarez
'
'1.This Class is used to Communicate the DataUtility Classes
'2.Create this Class once only.
'3.set the instancing to publicNotCreatable
'
'METHODS:
'1. ManipulateData(ByRef SQLString As String) As Boolean - it is used for manipulate data.
'       it is designed for UPDATE, DELETE and INSERT commands only
'       It returns a boolean, TRUE if the command is successfully implemented. FALSE if not
'2. GetData(ByVal query As String, ByRef rs As ADODB.Recordset)
'       it is designed for SELECT STATEMENT only
'       it returns a resultset after accessing database
'
'all methods requires SQLString as its command query when accessing the database.
'--------------------------------------------------NOTES--------------------------------------------------------------

Private con As New DL_UT_ACCESSDB

Function ManipulateData(ByRef SQLString As String) As Boolean
    ManipulateData = con.ModifyData(SQLString)
End Function

Function GetData(ByVal SQLString As String) As ADODB.Recordset
    Set GetData = con.GetData(SQLString)
End Function
