Attribute VB_Name = "Sys_FileDialog"
Option Compare Database

Private fDialog As Office.FileDialog

Function saveAs() As String
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .title = "Select Folder"
        If .Show = True Then
            For Each varFile In .SelectedItems
                saveAs = varFile
            Next
        End If
    End With
End Function

Function browseExcel(ByRef title As String, ByRef filterCollection As Collection) As String
    Dim FilterField As New BL_BE_FilterField
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    
    With fDialog
        .AllowMultiSelect = False
        .title = title
        .Filters.clear
        For Each FilterField In filterCollection
            If Not (Len(FilterField.Filters) = 0 Or Len(FilterField.FilterName) = 0) Then
                .Filters.Add FilterField.FilterName, FilterField.Filters
            End If
        Next
        If .Show = True Then
            For Each varFile In .SelectedItems
                browseExcel = varFile
            Next
        End If
    End With
End Function

Sub MakeDirectory(ByRef path As String)
    If Len(Dir(path, vbDirectory)) = 0 Then
       MkDir path
    End If
End Sub



'ADDING FILTERS
'Sub test()
'    Dim col As New Collection
'    Dim entity As New BL_BE_FilterField
'    entity.FilterName = "excel"
'    entity.Filters = "*.xls"
'    col.Add entity
'    Set entity = New BL_BE_FilterField
'    entity.FilterName = "excel"
'    entity.Filters = "*.xlsx"
'    col.Add entity
'    Call browseExcel("select", col)
'End Sub
