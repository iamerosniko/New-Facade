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
'    for debugging
'    Dim errList As Error
'
'    MsgBox Error$ & ":" & SQLcmd
'    For Each errList In DBEngine.Errors
'        MsgBox "(" & errList.Number & "): " & errList.Description
'    Next

    ModifyData = False
End Function

'get Data / View
Function GetData(ByVal SQLCmd As String) As ADODB.Recordset
On Error GoTo ErrFrom_GetData:
    Dim rs As New ADODB.Recordset
    
    OpenCon
    rs.Open SQLCmd, con, adOpenDynamic
    Set GetData = rs
GetDataExit:
    Exit Function
ErrFrom_GetData:
'    Dim errList As Error
'
'    MsgBox Error$ & ":" & SQLCmd
'    For Each errList In DBEngine.Errors
'        MsgBox "(" & errList.Number & "): " & errList.Description
'    Next
'
'    Set GetData = Nothing
'    Resume GetDataExit
End Function

'--------------------------------------------------Advanced  ADO
Function ExecCMDParam(ByRef params As Collection, ByRef strCMD As String) As ADODB.Recordset
On Error GoTo err:
    Dim param As New BL_BE_Parameters
    Dim gt As New BL_BC_GetDataTypeEnum
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim cmd As New ADODB.Command
    
    cmd.ActiveConnection = con
    cmd.CommandText = strCMD
                
    For i = 1 To params.Count
        For Each param In params(i)
            cmd.Parameters.Append cmd.CreateParameter(param.paramName, gt.gDataType(param.dataType), adParamInput, param.dataLength, param.paramValue)
        Next
    Next i
    
    Set rs = cmd.Execute
    'Set cmd = Nothing
    If Not rs.State = 1 Then
        Exit Function
    End If
    
    ExecCMDParam = rs
    Exit Function
err:
    Exit Function
End Function
