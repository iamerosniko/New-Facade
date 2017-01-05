Attribute VB_Name = "Sys_Settings"
Option Compare Database

Private Const Password As String = "Enca0907!"
Private pass As String

Function ap_DisableShift()
    'This function disable the shift at startup. This action causes
    'the Autoexec macro and Startup properties to always be executed.
    
    On Error GoTo errDisableShift
    
    Dim db As DAO.Database
    Dim prop As DAO.Property
    Const conPropNotFound = 3270
    
    Set db = CurrentDb()
    
    pass = Sys_Messages.TEST
    If pass = Password Then
    'This next line disables the shift key on startup.
        db.Properties("AllowByPassKey") = False
        MsgBox "Developer's Mode Disabled", vbInformation
        'The function is successful.
    Else
        MsgBox "Password is Incorrect", vbExclamation
    End If
    Exit Function
    
errDisableShift:
    'The first part of this error routine creates the "AllowByPassKey
    'property if it does not exist.
    If err = conPropNotFound Then
    Set prop = db.CreateProperty("AllowByPassKey", _
    dbBoolean, False)
    db.Properties.Append prop
    Resume Next
    Else
    MsgBox "Function 'ap_DisableShift' did not complete successfully."
    Exit Function
    End If

End Function

Function ap_EnableShift()
    'This function enables the SHIFT key at startup. This action causes
    'the Autoexec macro and the Startup properties to be bypassed
    'if the user holds down the SHIFT key when the user opens the database.
    
    On Error GoTo errEnableShift
    
    Dim db As DAO.Database
    Dim prop As DAO.Property
    Const conPropNotFound = 3270
    
    Set db = CurrentDb()
    
    pass = Sys_Messages.TEST
    If pass = Password Then
    'This next line of code disables the SHIFT key on startup.
        db.Properties("AllowByPassKey") = True
        MsgBox "Developer's Mode Enabled", vbInformation
    Else
        MsgBox "Password is Incorrect", vbExclamation
    End If
    'function successful
    Exit Function
    
errEnableShift:
    'The first part of this error routine creates the "AllowByPassKey
    'property if it does not exist.
    If err = conPropNotFound Then
    Set prop = db.CreateProperty("AllowByPassKey", _
    dbBoolean, True)
    db.Properties.Append prop
    Resume Next
    Else
    MsgBox "Function 'ap_DisableShift' did not complete successfully."
    Exit Function
    End If

End Function

