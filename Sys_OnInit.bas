Attribute VB_Name = "Sys_OnInit"
Option Compare Database

Private app_conf As New App_Config
Sub GetTitle()
    app_conf.GetTitle
End Sub

Sub sampleloc()
    app_conf.GetConnected
End Sub
