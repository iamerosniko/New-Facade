VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TEMPLATE_BE"
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
'This Class Sample is used to indicate what specific Entity contains
'NOTE: Before saving: under class properties, set instancing from "1 - private" to "2 - PublicNotCreatable".
'STANDARD NAMING: BL_BE_NAMEOFENTITY
'
'--------------------------------------------------NOTES--------------------------------------------------------------

'TODO: Declare a variable that is equal to its database table's columns
'EX:
'   SQL: TABLE_NAME = SUBJECTS
'   Columns : subj_no (integer), subj_name (text/varchar), subj_desc (text/varchar)
'
'   Public subj_no As Integer
'   Public subj_name As String
'   Public subj_desc As String
'NOTE: use datatype according to DB Table's columns definition
