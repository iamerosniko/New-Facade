VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BL_BC_GetDataTypeEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database

Function gDataType(ByRef rs As getDatatype) As Integer
    Select Case rs
        Case getDatatype.gArray
            gDataType = 8192
        Case getDatatype.gBigInt
            gDataType = 20
        Case getDatatype.gBinary
            gDataType = 128
        Case getDatatype.gBoolean
            gDataType = 11
        Case getDatatype.gChar
            gDataType = 129
        Case getDatatype.gCurrency
            gDataType = 6
        Case getDatatype.gDate
            gDataType = 7
        Case getDatatype.gDecimal
            gDataType = 14
        Case getDatatype.gDouble
            gDataType = 5
        Case getDatatype.gInteger
            gDataType = 3
        Case getDatatype.gTime
            gDataType = 134
        Case getDatatype.gVarchar
            gDataType = 200
    End Select
End Function
