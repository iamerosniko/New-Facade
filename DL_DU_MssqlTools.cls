VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DL_DU_MssqlTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private blbcConfigs As New BL_BC_AppCons
Private dbObject As DAO.Database
Private queryDef As DAO.queryDef
Private tbleDef As DAO.tableDef
Public isError As Boolean
Public errorMessage As String
'Dependencies: DL_DA_LTC, BL_BC_Configs

Private Sub OpenConn()
On Error GoTo OpenConnError
  Set dbObject = Application.CurrentDb
  Set queryDef = dbObject.CreateQueryDef("")
  queryDef.Connect = blbcConfigs.CONN_STRING
  queryDef.ODBCTimeout = 0
 
OpenConnExit:
    Exit Sub

OpenConnError:
    MsgBox Error$
    Dim dbeError As Error
    For Each dbeError In DBEngine.Errors
        MsgBox "(" & dbeError.Number & "): " & dbeError.Description
    Next
    Resume OpenConnExit
End Sub

Private Sub CloseConn()
On Error GoTo CloseConnError
    
    dbObject.Close
    queryDef.Close

CloseConnExit:
    Set dbObject = Nothing
    Set queryDef = Nothing
    Exit Sub

CloseConnError:
    MsgBox Error$
    Dim dbeError As Error
    For Each dbeError In DBEngine.Errors
        MsgBox "(" & dbeError.Number & "): " & dbeError.Description
    Next
    Resume CloseConnExit
    
End Sub




Public Function ModifyData(ByVal query As String)
On Error GoTo ModifyDataError
    
    OpenConn
    isError = False
    queryDef.sql = query
    queryDef.ReturnsRecords = False
    queryDef.Execute

ModifyDataExit:
    CloseConn
    Exit Function

ModifyDataError:
    MsgBox Error$ & query
    Dim errList As Error
    For Each errList In DBEngine.Errors
        MsgBox "(" & errList.Number & "): " & errList.Description
    Next
    isError = True
    Resume ModifyDataExit
    
End Function

Public Function GetData(ByVal query As String) As Collection
On Error GoTo GetDataError

    Dim resultset As DAO.Recordset
    Dim record As Variant
    Dim recordList As Collection
    
    isError = False
    Set recordList = New Collection
    
    OpenConn
    queryDef.sql = query
    Set resultset = queryDef.OpenRecordset(DB_OPEN_SNAPSHOT)
        
    If resultset.RecordCount > 0 Then
           
        resultset.MoveFirst
        Do Until resultset.EOF
        
            record = resultset.GetRows(1)
            recordList.Add record
        
        Loop
          
    End If
     
    Set GetData = recordList
      
GetDataExit:
    If Not resultset Is Nothing Then
        resultset.Close
    End If
    Set resultset = Nothing
    CloseConn
    Exit Function
    
GetDataError:
    errorMessage = Error$ & " " & query & " "
    Dim errList As Error
    For Each errList In DBEngine.Errors
        errorMessage = errorMessage & "(" & errList.Number & "): " & errList.Description & " "
    Next
    MsgBox errorMessage
    isError = True
    Resume GetDataExit
End Function

Public Function GetPermission(ByVal query As String) As Collection
On Error GoTo GetDataError

    Dim resultset As DAO.Recordset
    Dim record As New BL_BE_BTSSModules
    Dim recordList As Collection
    
    isError = False
    Set recordList = New Collection
    
    OpenConn
    queryDef.sql = query
    Set resultset = queryDef.OpenRecordset(DB_OPEN_SNAPSHOT)
        
    If resultset.RecordCount > 0 Then
           
        resultset.MoveFirst
        Do Until resultset.EOF
        
            record.BTSS_CanView = resultset(0)
            record.BTSS_CanAdd = resultset(1)
            record.BTSS_CanEdit = resultset(2)
            record.BTSS_CanDelete = resultset(3)
            recordList.Add record
            resultset.MoveNext
        Loop
    Else
        
    End If
     
    Set GetPermission = recordList
      
GetDataExit:
    If Not resultset Is Nothing Then
        resultset.Close
    End If
    Set resultset = Nothing
    CloseConn
    Exit Function
    
GetDataError:
    errorMessage = Error$ & " " & query & " "
    Dim errList As Error
    For Each errList In DBEngine.Errors
        errorMessage = errorMessage & "(" & errList.Number & "): " & errList.Description & " "
    Next
    MsgBox errorMessage
    isError = True
    Resume GetDataExit
End Function

