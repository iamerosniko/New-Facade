Option Compare Database

Implements BL_BC_ICRUD

Private queryinfo As BL_BC_QueryInfo
Private dispatcher As New BL_BC_Dispatcher
Private con As New DL_DA_Generic

'--------------------------------------------BUSINESS WORKFLOW--------------------------------------------------------
'SELECT

Private Function BL_BC_ICRUD_SelectAll() As Collection
    Set BL_BC_ICRUD_SelectAll = BL_BC_ICRUD_CompileAsCollection(BL_BC_ICRUD_ReadData)
End Function

Private Function BL_BC_ICRUD_CompileAsCollection(record As ADODB.Recordset) As Collection
    Dim col As New Collection
    Dim model As Variant

    Do Until record.EOF
        'TODO: INSTANTIATE TO ITS CORRESPONDING BUSINESS ENTITY
        'model = New BL_BE_Products
        'TODO: Match Entity.Fields to recordset.Fields
        'model.prod_name = record.Fields(0)
        'model.prod_desc = record.Fields(1)
        'model.prod_price = record.Fields(2)
        'after matching add it to temporary collection
        col.Add model
        record.MoveNext
    Loop

    Set BL_BC_ICRUD_CompileAsCollection = col
End Function

Private Function BL_BC_ICRUD_ModifyData(CR As CRUD, param As Collection) As Boolean
    'TODO: DO NOT CHANGE CODES BELOW
    Dim SQLCMD As String
    If CR = CInsert Then
        SQLCMD = BL_BC_ICRUD_CreateData
    ElseIf CR = CDelete Then
        SQLCMD = BL_BC_ICRUD_DeleteData
    ElseIf CR = CUpdate Then
        SQLCMD = BL_BC_ICRUD_UpdateData
    End If
    BL_BC_ICRUD_ModifyData = con.execCMDParam(param, SQLCMD)
End Function
'--------------------------------------END OF BUSINESS WORKFLOW SECTION-----------------------------------------------

'----------------------------------------------CRUD METHODS-----------------------------------------------------------
'TODO: DO NOT CHANGE CRUD METHODS. JUST COPY AND PASTE THIS SECTION
Private Function BL_BC_ICRUD_ReadData() As ADODB.Recordset
    Dim SQLCMD As String
    Dim rs As New ADODB.Recordset

    Set queryinfo = New BL_BC_QueryInfo
    queryinfo.tableName = BL_BC_ICRUD_TableSets
    queryinfo.ColumnSets = BL_BC_ICRUD_ColumnSets(CSelect)
    queryinfo.WhereCondition = BL_BC_ICRUD_WhereSets(CSelect)
    SQLCMD = dispatcher.GetSQLCMD(CSelect, queryinfo)
    Call con.getData(SQLCMD, rs)
    Set BL_BC_ICRUD_ReadData = rs
End Function

Private Function BL_BC_ICRUD_CreateData() As String
    Set queryinfo = New BL_BC_QueryInfo

    queryinfo.tableName = BL_BC_ICRUD_TableSets
    queryinfo.ColumnSets = BL_BC_ICRUD_ColumnSets(CInsert)
    queryinfo.ValuesSets = BL_BC_ICRUD_ValueSets
    BL_BC_ICRUD_CreateData = dispatcher.GetSQLCMD(CInsert, queryinfo)
End Function

Private Function BL_BC_ICRUD_UpdateData() As String
    Set queryinfo = New BL_BC_QueryInfo

    queryinfo.tableName = BL_BC_ICRUD_TableSets
    queryinfo.WhereCondition = BL_BC_ICRUD_WhereSets(CUpdate)
    queryinfo.UpdateSets = BL_BC_ICRUD_UpdateSets
    BL_BC_ICRUD_UpdateData = dispatcher.GetSQLCMD(CUpdate, queryinfo)
End Function

Private Function BL_BC_ICRUD_DeleteData() As String
    Set queryinfo = New BL_BC_QueryInfo

    queryinfo.tableName = BL_BC_ICRUD_TableSets
    queryinfo.WhereCondition = BL_BC_ICRUD_WhereSets(CDelete)
    BL_BC_ICRUD_DeleteData = dispatcher.GetSQLCMD(CDelete, queryinfo)
End Function
'--------------------------------------------END CRUD SECTION---------------------------------------------------------

'-------------------------------------------------QUERY INFO----------------------------------------------------------
Private Function BL_BC_ICRUD_ColumnSets(ByVal CR As CRUD) As String
    'Enclosed Columns in square brackets []
    If CR = CSelect Then
        BL_BC_ICRUD_ColumnSets = ""
    ElseIf CR = CInsert Then
        BL_BC_ICRUD_ColumnSets = ""
    End If
End Function

Private Function BL_BC_ICRUD_TableSets() As String 'CRUD
    BL_BC_ICRUD_TableSets = "" 'table name
End Function

Private Function BL_BC_ICRUD_UpdateSets() As String 'UPDATE
'    BL_BC_ICRUD_UpdateSets = "price = [param1], prod_desc = [param2]"
End Function

Private Function BL_BC_ICRUD_ValueSets() As String 'INSERT
'    BL_BC_ICRUD_ValueSets = [param1],[param2],[param3]
End Function

Private Function BL_BC_ICRUD_WhereSets(ByRef CR As CRUD) As String
    If CR = CDelete Then
        BL_BC_ICRUD_WhereSets = "" 'delete condition
    ElseIf CR = CSelect Then
        BL_BC_ICRUD_WhereSets = "" 'select condition
    ElseIf CR = CUpdate Then
        BL_BC_ICRUD_WhereSets = "" 'Update condition
    End If
End Function

'-----------------------------------------END OF QUERY INFO SECTION---------------------------------------------------

