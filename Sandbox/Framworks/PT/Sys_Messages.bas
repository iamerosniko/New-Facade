Attribute VB_Name = "Sys_Messages"
Option Compare Database
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, _
                                                      ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
 
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
 
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
                                          (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, _
                                          ByVal dwThreadId As Long) As Long
 
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
 
Private Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" _
                                            (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, _
                                            ByVal wParam As Long, ByVal lParam As Long) As Long
 
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                                                                          ByVal lpClassName As String, _
                                                                          ByVal nMaxCount As Long) As Long
 
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
 
'Constants to be used in our API functions
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
Private Const HC_ACTION = 0
 
Private hHook As Long

Public Enum messageType
    PlainMessage = 0
    AbortRetryIgnore = 1
    Warning = 2
    Information = 3
    Help = 4
    OKCancel = 5
    YesNo = 6
End Enum
 
Private Function NewProc(ByVal lngCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim RetVal
    Dim strClassName As String, lngBuffer As Long
 
    If lngCode < HC_ACTION Then
        NewProc = CallNextHookEx(hHook, lngCode, wParam, lParam)
        Exit Function
    End If
 
    strClassName = String$(256, " ")
    lngBuffer = 255
 
    If lngCode = HCBT_ACTIVATE Then    'A window has been activated
 
        RetVal = GetClassName(wParam, strClassName, lngBuffer)
 
        If Left$(strClassName, RetVal) = "#32770" Then  'Class name of the Inputbox
 
            'This changes the edit control so that it display the password character *.
            'You can change the Asc("*") as you please.
            SendDlgItemMessage wParam, &H1324, EM_SETPASSWORDCHAR, Asc("*"), &H0
        End If
 
    End If
 
    'This line will ensure that any other hooks that may be in place are
    'called correctly.
    CallNextHookEx hHook, lngCode, wParam, lParam
 
End Function
 
Private Function InputBoxDK(Prompt, Optional Title, Optional Default, Optional XPos, _
                        Optional YPos, Optional HelpFile, Optional Context) As String
 
    Dim lngModHwnd As Long, lngThreadID As Long
 
    lngThreadID = GetCurrentThreadId
    lngModHwnd = GetModuleHandle(vbNullString)
 
    hHook = SetWindowsHookEx(WH_CBT, AddressOf NewProc, lngModHwnd, lngThreadID)
 
    InputBoxDK = InputBox(Prompt, Title, Default, XPos, YPos, HelpFile, Context)
    UnhookWindowsHookEx hHook
 
End Function  'Hope someone can use it!

Function TEST() As String
    Dim strAdminPWord As String
    strAdminPWord = InputBoxDK("Password required to proceed.", "Enter Developer's code")
    TEST = strAdminPWord
End Function



Function msg(ByVal message As String, ByVal msgTitle As String, msgType As messageType)
    Select Case msgType
        Case messageType.PlainMessage
            msg = MsgBox(message, , msgTitle)
        Case messageType.AbortRetryIgnore
            msg = MsgBox(message, vbAbortRetryIgnore, msgTitle)
        Case messageType.Warning
            msg = MsgBox(message, vbExclamation, msgTitle)
        Case messageType.Information
            msg = MsgBox(message, vbInformation, msgTitle)
        Case messageType.Help
            msg = MsgBox(message, vbMsgBoxHelpButton, msgTitle)
        Case messageType.OKCancel
            msg = MsgBox(message, vbOKCancel, msgTitle)
        Case messageType.YesNo
            msg = MsgBox(message, vbYesNo, msgTitle)
    End Select
End Function

'----------RETURN VALUES---------'
' Constant    Value   Description'
' vbOK        1       OK         '
' vbCancel    2       Cancel     '
' vbAbort     3       Abort      '
' vbRetry     4       Retry      '
' vbIgnore    5       Ignore     '
' vbYes       6       Yes        '
' vbNo        7       No         '
'--------------------------------'

