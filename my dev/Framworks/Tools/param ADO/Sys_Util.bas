Attribute VB_Name = "Sys_Util"
 Option Compare Database

Private con As New DL_DA_Generic
Function RecordValueToCollection(ByRef rs As ADODB.Recordset, ByRef index As Integer) As Collection
    Dim col As New Collection
    
    Do Until rs.EOF
        col.Add rs.Fields(index).value
        rs.MoveNext
    Loop
    Set RecordValueToCollection = col
End Function

Function RecordNameToCollection(ByRef rs As ADODB.Recordset) As Collection
    Dim col As New Collection
    Dim cols As Long
    Dim en As Long
    
    en = rs.Fields.Count
    For cols = 0 To en - 1
        col.Add rs.Fields(cols).Name
    Next cols

    Set RecordNameToCollection = col
End Function

Function GetSQLFromQuery(ByRef QueryName As String) As String
On Error GoTo err:
    Dim qdf As QueryDef
    
    Set qdf = CurrentDb.QueryDefs(QueryName)
    GetSQLFromQuery = qdf.SQL
    qdf.Close
err:
    Sys_Messages.msg "Query name " & QueryName & "not found", "Error", Warning
End Function

Function addParam(ByVal parameterName As String, ByVal dataType As getDatatype, _
    ByVal datasize As Integer, ByVal parameterValue As String) As Collection
    
    Dim col As New Collection
    Dim par As New BL_BE_Parameters
    'add a parameter
    par.dataLength = datasize
    par.dataType = dataType
    par.paramName = parameterName
    par.paramValue = parameterValue
    col.Add par
    
    Set addParam = col
End Function

