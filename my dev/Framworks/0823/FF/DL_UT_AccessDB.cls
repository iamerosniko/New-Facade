VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DL_UT_ACCESSDB"
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
'This Class is used to access the database or the data repository (MSACCESS)
'
'METHODS:
'   OpenCon() - is used to open the database connectivity
'   CloseCon()- is used to Close the database connectivity
'   ModifyData() - is used to modify data (INSERT, UPDATE and DELETE commands). It returns a boolean to know if the
'       command query works (TRUE) or not (FALSE).
'                - it accepts string as its command for database
'   GetData() - is used to get data from repository.
'             - it accepts string as its command for database.
'
'--------------------------------------------------NOTES--------------------------------------------------------------

Private con As New ADODB.Connection

Private Sub OpenCon()
    Set con = CurrentProject.Connection
End Sub

Private Sub CloseCon()
    con.Close
    Set con = Nothing
End Sub

'INSERT, UPDATE and DELETE Method
Function ModifyData(ByVal SQLCmd As String) As Boolean
On Error GoTo ErrFrom_ModifyData:
    
    OpenCon
    con.Execute SQLCmd
    ModifyData = True
    Exit Function

ErrFrom_ModifyData:

    ModifyData = False
End Function

'get Data / View
Public Function GetData(ByVal SQLCmd As String) As ADODB.Recordset
On Error GoTo ErrFrom_GetData:
    Dim rs As New ADODB.Recordset
    
    OpenCon
    rs.Open SQLCmd, con, adOpenDynamic
    Set GetData = rs
GetDataExit:
    Exit Function
ErrFrom_GetData:
    Dim errList As Error
    
    MsgBox Error$ & ":" & SQLCmd
    For Each errList In DBEngine.Errors
        MsgBox "(" & errList.Number & "): " & errList.Description
    Next
    
    Resume GetDataExit
End Function
