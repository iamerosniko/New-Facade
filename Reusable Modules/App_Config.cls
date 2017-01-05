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

'declare your Environment Location Here (MOD and OPS only)
Private Const UAT_Location As String = "UAT\Apps"
Private Const OPS_Location As String = "OPS\Apps"

'declare your Environment ConnectionStrings here
'for Dev Server
Const DEV_ConStr As String = "ODBC;DRIVER={sql server};" & _
                                    "DATABASE=dbname;" & _
                                    "SERVER=serverhere" & _
                                    "Trusted_Connection=YES;"
'for MOD Server
Const UAT_ConStr As String = "ODBC;DRIVER={sql server};" & _
                                    "DATABASE=dbname;" & _
                                    "SERVER=serverhere" & _
                                    "Trusted_Connection=YES;"
'for PRODUCTION Server
Const OPS_ConStr As String = "ODBC;DRIVER={sql server};" & _
                                    "DATABASE=dbname;" & _
                                    "SERVER=serverhere" & _
                                    "Trusted_Connection=YES;"
                                    
'declare title of the application here
Private Const APP_Title As String = "TITLEHERE"

'this will provide a Connection string depending on its location
Public Function GetConnection() As String
    Dim Str As String
    Str = Application.CurrentProject.path
    If InStr(LCase(Str), LCase(UAT_Location)) > 0 Then
        GetConnection = UAT_ConStr
    ElseIf InStr(LCase(Str), LCase(OPS_Location)) > 0 Then
        GetConnection = OPS_ConStr
    Else
        GetConnection = DEV_ConStr
    End If
End Function

'Determines application's Environment Mode
Public Sub GetEnvironmentMode()
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
    'if AppTitle Property not found
    If Err.Number = conPropNotFoundError Then
        Set obj = dbs.CreateProperty("AppTitle", dbText, GetEnvironmentMode)
        dbs.Properties.Append obj
    Else
        MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, , "App_Config/GetTitle"
    End If
    
    Resume Next
End Sub

'This determines the application's title
Private Function GetTitle()
    Dim Str As String
    
    GetTitle = APP_Title & "  "
    Str = Application.CurrentProject.path
    If InStr(LCase(Str), LCase(UAT_Location)) > 0 Then
        GetTitle = GetTitle & "UAT Mode"
    ElseIf InStr(LCase(Str), LCase(OPS_Location)) > 0 Then
        GetTitle = Trim(GetTitle)
    Else
        GetTitle = GetTitle & "Developers Mode"
    End If
End Function

