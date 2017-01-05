VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BL_BC_Dispatcher"
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
'This Class is used to implement as interface for Business Workflow
'It is also a standardized methods that is commonly used for CRUD
'NOTE: DO NOT MODIFY
'
'--------------------------------------------------NOTES--------------------------------------------------------------

Function GetSQLCMD(ByVal Mode As String, ByRef queryInfo As BL_BC_QueryInfo) As String
    queryInfo.ColumnSets = Trim(queryInfo.ColumnSets)
    queryInfo.tableName = Trim(queryInfo.tableName)
    queryInfo.UpdateSets = Trim(queryInfo.UpdateSets)
    queryInfo.ValuesSets = Trim(queryInfo.ValuesSets)
    queryInfo.WhereCondition = Trim(queryInfo.WhereCondition)
    
    If Mode = "UPDATE" Then
         GetSQLCMD = "UPDATE " & queryInfo.tableName & _
                        " SET " & queryInfo.UpdateSets & _
                        " WHERE (" & queryInfo.WhereCondition & ")"
                        
    ElseIf Mode = "DELETE" Then
        GetSQLCMD = "DELETE FROM " & queryInfo.tableName & _
                        " WHERE (" & queryInfo.WhereCondition & ")"
                        
    ElseIf Mode = "INSERT" Then
        GetSQLCMD = "INSERT INTO " & queryInfo.tableName & " " & _
                    "(" & queryInfo.ColumnSets & ") " & _
                    "VALUES (" & queryInfo.ValuesSets & ")"
                    
    ElseIf Mode = "SELECT" Then
        GetSQLCMD = "SELECT " & queryInfo.ColumnSets & " " & _
                    "FROM " & queryInfo.tableName & " "
                    
        If Not (Len(queryInfo.WhereCondition) = 0) Then
            GetSQLCMD = GetSQLCMD & "WHERE " & queryInfo.WhereCondition

        End If
        
    End If
End Function

