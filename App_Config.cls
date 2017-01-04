VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database

Private Const DEV_Location As String = "DEV\Apps"
Private Const UAT_Location As String = "UAT\Apps"
Private Const OPS_Location As String = "OPS\Apps"
Private Const APP_Title As String = "TITLEHERE"

Public Function GetConnection()
    Dim Str As String
    Str = Application.CurrentProject.path
    If InStr(LCase(Str), LCase(DEV_Location)) > 0 Then
        
        MsgBox "under dev"
    ElseIf InStr(LCase(Str), LCase(UAT_Location)) > 0 Then
        MsgBox "under stage"
    ElseIf InStr(LCase(Str), LCase(OPS_Location)) > 0 Then
        MsgBox "under production"
    Else
        MsgBox "default: under dev"
    End If
End Function
   
Private Function GetEnvironmentMode()
    Dim Str As String
    GetEnvironmentMode = APP_Title & "--"
    Str = Application.CurrentProject.path
    If InStr(LCase(Str), LCase(DEV_Location)) > 0 Then
        GetEnvironmentMode = GetEnvironmentMode & "DEV"
    ElseIf InStr(LCase(Str), LCase(UAT_Location)) > 0 Then
        GetEnvironmentMode = GetEnvironmentMode & "MOD"
    ElseIf InStr(LCase(Str), LCase(OPS_Location)) > 0 Then
        GetEnvironmentMode = GetEnvironmentMode & "OPS"
    Else
        GetEnvironmentMode = GetEnvironmentMode & "DEV"
    End If
End Function

Public Sub GetTitle()
    Dim obj As Object
    Const conPropNotFoundError = 3270
    
    On Error GoTo ErrorHandler
    Set dbs = CurrentDb
    ' Change title bar.
    dbs.Properties!AppTitle = GetEnvironmentMode
    ' Update title bar on screen.
    Application.RefreshTitleBar
    Exit Sub
 
ErrorHandler:
    If Err.Number = conPropNotFoundError Then
        Set obj = dbs.CreateProperty("AppTitle", dbText, GetEnvironmentMode)
        dbs.Properties.Append obj
    Else
        MsgBox "Error: " & Err.Number & vbCrLf & Err.Description
    End If
    
    Resume Next
End Sub
