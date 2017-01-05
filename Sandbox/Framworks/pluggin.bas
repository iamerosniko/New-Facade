Attribute VB_Name = "Module1"
Option Compare Database
Private pluggin As New BL_BC_ImportPluggin


Sub test()
    Dim source As String, map As String
    source = "C:\Users\alvaero\Desktop\test\20160829_103811_sampl.xls"
    map = "C:\Users\alvaero\Desktop\MAP.xlsx"
    'true or false
    'if false then there is some errors
    'if true it can proceed to importation
    'NOTE: the validator only depends on its map and its source.
    Call pluggin.StartValidation(source, map)
End Sub

