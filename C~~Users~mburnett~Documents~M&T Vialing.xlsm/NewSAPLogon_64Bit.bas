Attribute VB_Name = "NewSAPLogon_64Bit"
Option Explicit
Public Appl As Variant
Public App2 As Variant
Public Connection As Variant
Public Session As Variant
Public Connection1 As Variant
Public Session1 As Variant
Public UserName As String
Public Password As String
Public SapGuiAuto As Object
Public SapGuiAuto1 As Object
#If VBA7 Then
Public Declare PtrSafe Function SetCurrentDirectoryA Lib "KERNEL32" (ByVal lpPathName As String) As Long
Public Declare PtrSafe Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal HWnd As Long, ByVal nCmdSHow As Long) As Long
Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal HWnd As Long) As Long
#Else
Public Declare Function SetCurrentDirectoryA Lib "KERNEL32" (ByVal lpPathName As String) As Long
Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal HWnd As Long, ByVal nCmdSHow As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal HWnd As Long) As Long
#End If
Public CountDown As Variant
Public RunWhen As Double
Public Const MyFilePath = "\\BETHYL-FS1\BethylShared\Logs\GUIScriptFiles\"
Public Const MyFile = MyFilePath & "_l1O0.jpg"
Public Const WM_CLOSE = &H10
Public Const INFINITE = -1&
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const EM_SETPASSWORDCHAR = &HCC
Public Const WH_CBT = 5
Public Const HCBT_ACTIVATE = 5
Public Const HC_ACTION = 0
Public Const NUM_MINUTES = 600  'SAP Password Reset Countdown timer start seconds
Public Const GetIP = "192.168.5.61 10"
Sub LogonProduction()
    Dim WSHShell As Object
    Dim MyPath As String
    Dim MyVal As String
    Call LogOff
    If Range("ZZ1") <> "" Then UserName = Range("ZZ1")
    If Range("AAA1") <> "" Then Password = Range("AAA1")
    AppActivate Application.Caption
    If UserName = "" Then SAP_Password.Caption = "Production - SAP Password": SAP_Password.Show
    If UserName = "" Then Exit Sub
    If Password = "" Then Exit Sub
    MyVal = Dir("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
    If MyVal <> "" Then
        MyPath = "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    Else
        MyPath = "C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe"
    End If
    Shell (MyPath)
    Set WSHShell = CreateObject("WScript.Shell")
    Do Until WSHShell.AppActivate("SAP Logon ")
        Shell (MyPath)
        Sleep 1000
    Loop
    Call MinimizeSAPLogon
    Set WSHShell = Nothing
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Appl = SapGuiAuto.GetScriptingEngine
    'Set Connection = Appl.OpenConnection("BLP on Windows 2012", True)
    Set Connection = Appl.OpenConnectionByConnectionString(GetIP)
    Set Session = Connection.Children(0)
retry:
    Session.FindById("wnd[0]").Maximize
    Session.FindById("wnd[0]/usr/txtRSYST-MANDT").Text = "100"
    Session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = UserName
    Session.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = Password
    Session.FindById("wnd[0]/usr/txtRSYST-LANGU").Text = "EN"
    Session.FindById("wnd[0]").SendVKey 0
    On Error Resume Next
    Session.FindById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select
    On Error GoTo 0
    Session.FindById("wnd[0]").SendVKey 0
    If InStr(Session.FindById("wnd[0]/sbar").Text, "required entry") > 0 Then
        AppActivate Application.Caption
        MsgBox "Username/Password combination is incorrect. Please try again."
        UserName = ""
        Password = ""
        SAP_Password.txtUserName = ""
        SAP_Password.txtPassword = ""
        SAP_Password.txtUserName.SetFocus
        SAP_Password.Show
        GoTo retry
    End If
    If InStr(Session.FindById("wnd[0]/sbar").Text, "incorrect") > 0 Then
        AppActivate Application.Caption
        MsgBox "Username/Password combination is incorrect. Please try again."
        UserName = ""
        Password = ""
        SAP_Password.txtUserName = ""
        SAP_Password.txtPassword = ""
        SAP_Password.txtUserName.SetFocus
        SAP_Password.Show
        GoTo retry
    End If
    Application.WindowState = xlMaximized
End Sub
Sub LogonDevelopment()
    Dim WSHShell As Object
    Dim MyPath As String
    Dim MyVal As String
    Call LogOff
    AppActivate Application.Caption
    If UserName = "" Then SAP_Password.Caption = "Development - SAP Password": SAP_Password.Show
    If UserName = "" Then Exit Sub
    If Password = "" Then Exit Sub
    MyVal = Dir("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
    If MyVal <> "" Then
        MyPath = "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    Else
        MyPath = "C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe"
    End If
    Shell (MyPath)
    Set WSHShell = CreateObject("WScript.Shell")
    Do Until WSHShell.AppActivate("SAP Logon ")
        Shell (MyPath)
        Sleep 1000
    Loop
    Call MinimizeSAPLogon
    Set WSHShell = Nothing
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Appl = SapGuiAuto.GetScriptingEngine
    'Set Connection = Appl.OpenConnection("BLD on Windows 2012", True)
    Set Connection = Appl.OpenConnectionByConnectionString(Replace(GetIP, "61", "51"))
    Set Session = Connection.Children(0)
retry:
    Session.FindById("wnd[0]").Maximize
    Session.FindById("wnd[0]/usr/txtRSYST-MANDT").Text = "300"
    Session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = UserName
    Session.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = Password
    Session.FindById("wnd[0]/usr/txtRSYST-LANGU").Text = "EN"
    Session.FindById("wnd[0]").SendVKey 0
    On Error Resume Next
    Session.FindById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select
    On Error GoTo 0
    Session.FindById("wnd[0]").SendVKey 0
    If InStr(Session.FindById("wnd[0]/sbar").Text, "required entry") > 0 Then
        AppActivate Application.Caption
        MsgBox "Username/Password combination is incorrect. Please try again."
        UserName = ""
        Password = ""
        SAP_Password.txtUserName = ""
        SAP_Password.txtPassword = ""
        SAP_Password.txtUserName.SetFocus
        SAP_Password.Show
        GoTo retry
    End If
    If InStr(Session.FindById("wnd[0]/sbar").Text, "incorrect") > 0 Then
        AppActivate Application.Caption
        MsgBox "Username/Password combination is incorrect. Please try again."
        UserName = ""
        Password = ""
        SAP_Password.txtUserName = ""
        SAP_Password.txtPassword = ""
        SAP_Password.txtUserName.SetFocus
        SAP_Password.Show
        GoTo retry
    End If
End Sub
Sub LogOff()
    On Error Resume Next
    Set Session = Nothing
    Connection.CloseSession ("ses[0]")
    Set Connection = Nothing
    Set SapGuiAuto = Nothing
End Sub
Sub ResetPassword()
    On Error Resume Next
    Password = ""
    UserName = ""
End Sub
'Function GetIP()
'    Dim MyChar As String
'    Dim MyChar1 As String
'    Dim I As Integer
'    On Error Resume Next
'    MyChar = FilePropertyExplorer.OpenFile(MyFile, True).Item(1).Value
'    MyChar1 = FilePropertyExplorer.OpenFile(MyFile, True).Item(7).ValueDesc
'    I = Val(Replace(Replace(Replace(Replace(MyFile, MyFilePath, ""), "_l", ""), "O", ""), ".jpg", ""))
'    GetIP = MyChar & MyChar1 & " " & Val(Replace(Replace(Replace(Replace(MyFile, MyFilePath, ""), "_l", ""), "O", ""), ".jpg", ""))
'End Function
Function GetPassword()
    Dim MyChar As String
    Dim MyChar1 As String
    On Error Resume Next
    MyChar = FilePropertyExplorer.OpenFile(MyFile, True).Item(6).ValueDesc
    MyChar1 = FilePropertyExplorer.OpenFile(MyFile, True).Item(4).Value
    GetPassword = MyChar & MyChar1
End Function
Function GetUserName()
    Dim MyChar As String
    On Error Resume Next
    MyChar = FilePropertyExplorer.OpenFile(MyFile, True).Item(5).ValueDesc
    GetUserName = MyChar
End Function
Function apicFindWindow(strClassName As String, strWindowName As String)
    'Get window handle.
    Dim lngWnd As Long
    apicFindWindow = FindWindow(strClassName, strWindowName)
End Function
Function apicShowWindow(strClassName As String, strWindowName As String, lngState As Long)
    'Get window handle.
    Dim lngWnd As Long
    Dim intRet As Integer
    lngWnd = FindWindow(strClassName, strWindowName)
    apicShowWindow = ShowWindow(lngWnd, lngState)
End Function
Sub MinimizeSAPLogon()
    Dim HWnd As Long
    HWnd = FindWindow(vbNullString, "SAP Logon 740")
    If HWnd Then
        '    SetForegroundWindow HWnd
        ShowWindow HWnd, 6 'minimize 6, maximize 3, restore 9
    End If
End Sub
