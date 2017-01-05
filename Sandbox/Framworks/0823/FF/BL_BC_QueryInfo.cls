VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BL_BC_QueryInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'--------------------------------------------------NOTES--------------------------------------------------------------
'
'Project: Facade Framework 1.0
'Date Created: 08/05/2016
'Created By: Eros Niko Cas Alvarez
'
'This Class is a BUSINESS COMPONENT to standardize the set up for SQL Query
'
'NOTE: It can only use for simple query. do not use this for joins
'    : Do Not Modify
'
'--------------------------------------------------NOTES--------------------------------------------------------------

Public tableName As String      'table target ex: SELECT * from TableName
Public UpdateSets As String     'use only in update statement
Public ColumnSets As String     'columns included in Insert EX: insert from students (col1,col2)
Public ValuesSets As String     'values included in insert EX: values (' ',' ',' ')
Public WhereCondition As String 'can be used in delete, update, or select

