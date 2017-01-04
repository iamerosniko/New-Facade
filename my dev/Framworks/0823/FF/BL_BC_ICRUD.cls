VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BL_BC_ICRUD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'--------------------------------------------------NOTES--------------------------------------------------------------
'                                                                                                                    '
'Project: Facade Framework 1.0                                                                                       '
'Date Created: 08/04/2016                                                                                            '
'Created By: Eros Niko Cas Alvarez                                                                                   '
'                                                                                                                    '
'This Class is the blueprint methods designed for business workflows                                                 '
'It is also a standardized methods that is commonly used for CRUD                                                    '
'                                                                                                                    '
'--------------------------------------------------NOTES--------------------------------------------------------------

'CRUD FUNCTIONALITY
Public Function ReadData() As ADODB.Recordset
End Function

Public Function CreateData() As Boolean
End Function

Public Function UpdateData() As Boolean
End Function

Public Function DeleteData() As Boolean
End Function

'QUERYINFO HELPER
Public Function ColumnSets(ByVal Mode As String) As String
End Function

Public Function TableSets() As String
End Function

Public Function UpdateSets() As String
End Function

Public Function ValueSets() As String
End Function

Public Function WhereSets(ByVal Mode As String) As String
End Function


