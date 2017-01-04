VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DL_UT_MSSQLTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private con As New DL_DA_Generic
Private AppCons As New BL_BC_AppCons

Public Sub InitGetTables(ByRef colTables As Collection)
    CreateTableLink colTables
End Sub

Private Function CreateTableLink(ByRef tbls As Collection)
    Dim i As Integer
    Dim ctr As Integer
    Dim temptable As String
    
    ctr = tbls.Count

    For i = 1 To ctr
        temptable = CStr(tbls(i))
        If TableExists(temptable) = 1 Then
            DoCmd.DeleteObject acTable, temptable
        End If
        GetTables (temptable)
    Next i
End Function

Private Function TableExists(ByVal TableName As String) As Integer
    Dim col As New Collection
    
    Set col = RecordValueToCollection(con.GetData("SELECT [Name] FROM MsysObjects WHERE [Name] = '" & TableName & "'"), 0)
    
    TableExists = col.Count
End Function

Private Sub GetTables(ByRef TableName As String)
    DoCmd.TransferDatabase acLink, "ODBC", AppCons.CONN_STRING, acTable, TableName, TableName
    Exit Sub
End Sub


