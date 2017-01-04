VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TEMPLATE_BW"
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
'INSTRUCTIONS:
'
'NOTE: Before saving: under class properties, set instancing from "1 - private" to "2 - PublicNotCreatable".
'STANDARD NAMING: BL_BW_NAMEOFWORKFLOW
'
'1.IMPLEMENT BL_BC_ICRUD
'2.Instantiate the following:
'   Private QueryInfo As BL_BC_QueryInfo
'   Private Dispatcher As New BL_BC_Dispatcher
'   Private DataAccess As New DL_DA_Generic
'3. Private S_Model as TEMPLATE_BE : change TEMPLATE_BE according to its corresponding ENTITY
'4.Copy CRUD Methods, Business Workflow Methods, Query Info Section
'5.EDIT THE FOLLOWING:
'BUSINESS WORKFLOW SECTION
'   VIEW - For viewing data. Returns Collection
'       1.From (ByVal Model As TEMPLATE_BE), Change TEMPLATE_BE to its Corresponding ENTITY
'   COMPILE
'       1.From (Dim Model As TEMPLATE_BE), Change TEMPLATE_BE to its Corresponding ENTITY
'       2.On (Do Until Record.EOF), assign model's variable from record's respective fields
'   MODIFYDATA
'       1.From (ByVal Mode As String, ByVal Model As TEMPLATE_BE):
'           1.1. Mode: Use (INSERT,DELETE,UPDATE) only
'           1.2. change TEMPLATE_BE to its Corresponding ENTITY
'QUERY INFO SECTION
'   BL_BC_ICRUD_ColumnSets(ByVal Mode As String) As String - USE FOR INSERT AND SELECT ONLY
'       1.if Mode = "INSERT": Display all Columns according to its assigned Table.
'   BL_BC_ICRUD_TableSets() : Set BL_BC_ICRUD_TableSets = to its table
'   BL_BC_ICRUD_UpdateSets(): For Update Only
'        NOTE: use the "Private Model" or the S_Model
'       1.assign BL_BC_ICRUD_UpdateSets to its corresponding value
'           NOTE: if the column contains VARCHAR use like this
'           EX: "columnName = '" & S_Teachers.varName & "'"
'                 if the column contains integer or Float, do not insert ' on before and after value
'           EX: "columnName = " & S_Teachers.varName &
'   BL_BC_ICRUD_ValueSets() : For INSERT only
'        NOTE: use the "Private Model" or the S_Model
'   BL_BC_ICRUD_WhereSets(ByVal Mode As String) : For SELECT, DELETE, and UPDATE
'        accepts mode for selection, delete or update
'        THIS METHOD DEPENDS ON BW's REQUIREMENTS
'
'--------------------------------------------------NOTES--------------------------------------------------------------

Implements BL_BC_ICRUD

Private queryInfo As BL_BC_QueryInfo
Private dispatcher As New BL_BC_Dispatcher
Private DataAccess As New DL_DA_Generic

Private S_Model As TEMPLATE_BE
'----------------------------------------------CRUD METHODS----------------------------------------------------------
'TODO: DO NOT CHANGE CRUD METHODS. JUST COPY AND PASTE THIS SECTION
Private Function BL_BC_ICRUD_CreateData() As Boolean
    Dim SQLCmd As String
    
    Set queryInfo = New BL_BC_QueryInfo
    queryInfo.tableName = BL_BC_ICRUD_TableSets
    queryInfo.ColumnSets = BL_BC_ICRUD_ColumnSets("INSERT")
    queryInfo.ValuesSets = BL_BC_ICRUD_ValueSets
    SQLCmd = dispatcher.GetSQLCMD("INSERT", queryInfo)
    BL_BC_ICRUD_CreateData = DataAccess.ManipulateData(SQLCmd)
End Function

Private Function BL_BC_ICRUD_ReadData() As ADODB.Recordset
    Dim SQLCmd As String
    Dim rs As New ADODB.Recordset

    Set queryInfo = New BL_BC_QueryInfo
    queryInfo.tableName = BL_BC_ICRUD_TableSets
    queryInfo.ColumnSets = BL_BC_ICRUD_ColumnSets("SELECT")
    queryInfo.WhereCondition = BL_BC_ICRUD_WhereSets("SELECT")
    SQLCmd = dispatcher.GetSQLCMD("SELECT", queryInfo)
    Call DataAccess.GetData(SQLCmd, rs)
    Set BL_BC_ICRUD_ReadData = rs
End Function

Private Function BL_BC_ICRUD_UpdateData() As Boolean
    Dim SQLCmd As String
    Set queryInfo = New BL_BC_QueryInfo
    queryInfo.tableName = BL_BC_ICRUD_TableSets
    queryInfo.WhereCondition = BL_BC_ICRUD_WhereSets("UPDATE")
    queryInfo.UpdateSets = BL_BC_ICRUD_UpdateSets
    SQLCmd = dispatcher.GetSQLCMD("UPDATE", queryInfo)
    BL_BC_ICRUD_UpdateData = DataAccess.ManipulateData(SQLCmd)
End Function

Private Function BL_BC_ICRUD_DeleteData() As Boolean
    Dim SQLCmd As String
    Set queryInfo = New BL_BC_QueryInfo
    queryInfo.tableName = BL_BC_ICRUD_TableSets
    queryInfo.WhereCondition = BL_BC_ICRUD_WhereSets("DELETE")
    SQLCmd = dispatcher.GetSQLCMD("DELETE", queryInfo)
    BL_BC_ICRUD_DeleteData = DataAccess.ManipulateData(SQLCmd)
End Function
'--------------------------------------------END CRUD SECTION---------------------------------------------------------

'--------------------------------------------BUSINESS WORKFLOW--------------------------------------------------------
'TODO: Change "TEMPLATE_BE" to its Corresponding ENTITY
Function View(ByVal Model As TEMPLATE_BE) As Collection
    'DO NOT CHANGE THE CODE BELOW
    Dim Record As ADODB.Recordset
    Set S_Model = Model
    Set Record = BL_BC_ICRUD_ReadData()
    Set View = Compile(Record)
End Function

Private Function Compile(ByRef Record As ADODB.Recordset) As Collection
    Dim Collect As New Collection
    'TODO: CHANGE TEMPLATE_BE TO ITS CORRESPONDING BUSINESS ENTITY
    Dim Model As TEMPLATE_BE
    
    Do Until Record.EOF
        'TODO: INSTANTIATE TO ITS CORRESPONDING BUSINESS ENTITY
        Set Model = New TEMPLATE_BE
        'TODO: FILL THE FIELDS ACCORDING TO BUSINESS ENTITY's CONTENTS and RECORD'S FIELDS
'        Model.child1 = Record.Fields(0)
'        Model.child2 = Record.Fields(1)
        Collect.Add Model
        Record.MoveNext
    Loop
    Set Compile = Collect
End Function

Function ModifyData(ByVal Mode As String, ByVal Model As TEMPLATE_BE) As Boolean
    'TODO: DO NOT CHANGE CODES BELOW
    Set S_Model = Model
    If (Mode = "INSERT") Then
        ModifyData = BL_BC_ICRUD_CreateData
    ElseIf (Mode = "DELETE") Then
        ModifyData = BL_BC_ICRUD_DeleteData
    ElseIf (Mode = "UPDATE") Then
        ModifyData = BL_BC_ICRUD_UpdateData
    End If
End Function
'--------------------------------------END OF BUSINESS WORKFLOW SECTION-----------------------------------------------

'-------------------------------------------------QUERY INFO----------------------------------------------------------

Private Function BL_BC_ICRUD_ColumnSets(ByVal Mode As String) As String 'SELECT AND INSERT
    If Mode = "SELECT" Then
        'IT IS RECOMMENDED TO GATHER ALL DATA TO AVOID ERRORS
        BL_BC_ICRUD_ColumnSets = "*"
    ElseIf Mode = "INSERT" Then
        'TODO: CHANGE "InsertColumnsHere" to COLUMNFORMAT like this
        'EX: "Teacher_id,Teacher_fname,teacher_mname,Teacher_lname,Teacher_address"
        BL_BC_ICRUD_ColumnSets = "InsertColumnsHere"
    End If
End Function

Private Function BL_BC_ICRUD_TableSets() As String 'CRUD
    'TODO: CHANGE "InsertTableNameHere" TO ITS CORRESPONDING TABLE
    BL_BC_ICRUD_TableSets = "InsertTableNameHere"
End Function


Private Function BL_BC_ICRUD_UpdateSets() As String 'UPDATE
'    BL_BC_ICRUD_UpdateSets = "columnname1 = '" & S_Model.child1 & "'," & _
'                             "columnname2 = '" & S_Model.child2 & "'"
End Function

Private Function BL_BC_ICRUD_ValueSets() As String 'INSERT
'    BL_BC_ICRUD_ValueSets = CStr(S_Model.child1) & "," & _
'                           "'" & S_Model.child2 & "'"
End Function

'THIS METHOD DEPENDS ON BW's REQUIREMENTS
Private Function BL_BC_ICRUD_WhereSets(ByVal Mode As String) As String 'MODE (SELECT, UPDATE and DELETE)
    Dim whereClause As String
    Mode = UCase(Mode)
    If (Mode = "SELECT") Then
        'whereClause = "whereClause for select"
    ElseIf (Mode = "UPDATE") Then
        'whereClause = "whereClause for Update"
    ElseIf (Mode = "DELETE") Then
        'whereClause = "whereClause for Delete"
    Else
        MsgBox Mode & " is not allowed. Mode must be SELECT, UPDATE and DELETE only"
    End If
    BL_BC_ICRUD_WhereSets = whereClause
End Function

'-----------------------------------------END OF QUERY INFO SECTION---------------------------------------------------
