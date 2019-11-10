Attribute VB_Name = "LoggingCode"
Option Explicit
Public Const LogPath = "\\BETHYL-FS1\BethylShared\Logs"
Public DateTime As String
#If VBA7 Then
Private Type HOSTENT
    hName As LongPtr
    hAliases As LongPtr
    hAddrType As Integer
    hLength As Integer
    hAddrList As LongPtr
End Type
Private Declare PtrSafe Function gethostbyname Lib "WSOCK32.DLL" (ByVal HostName As String) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "KERNEL32" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
#Else
Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type
Private Declare Function gethostbyname Lib "WSOCK32.DLL" Alias "GetHostByName" (ByVal HostName$) As Long
Private Declare Sub RtlMoveMemory Lib "KERNEL32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
#End If
Sub LogAll()
    On Error Resume Next
    Logger "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
    Logger "The Program's Name: " & ThisWorkbook.Name
    Logger "The Program's path: " & ThisWorkbook.Path
    Logger "Date Program was Created: " & GetDateCreated(ThisWorkbook.Path & "\" & ThisWorkbook.Name)
    Logger "Date Program was LastModified: " & GetDateLastModified(ThisWorkbook.Path & "\" & ThisWorkbook.Name)
    Logger GetPCName
    Logger GetPC_IP
    Logger GetOS
    Logger GetPersonLoggedInToPC
    Logger GetRAM
    Logger GetFreeMemory
    Logger GetNumProcessors
    Logger GetPCsSpeed
    Logger GetPCType
    Logger GetPCSerialNumAssetTag
    '    Logger PingPC(GetIPFromHostName("Bethyl-Server2"))
    '    Logger PingPC("192.168.0.109") & "   C2191 (Dropbox PC)"
    Logger PingPC("192.168.5.11") & "   Bethyl-Server2"
    Logger GetRunningPrograms("outlook.exe")
    Logger GetRunningPrograms("sap.exe")
    Logger GetRunningPrograms("excel.exe")
    Logger GetRunningPrograms("word.exe")
    Logger GetRunningPrograms("access.exe")
    Logger GetAllEnvironmentalVariables
    Logger GetPCsDevicesNotWorking
    Logger GetUsedDiskSpace
    Logger GetAllServices
    Logger AllRunningApps
    Logger "###################################################################################################################################################################################"
End Sub
Function GetDateLastModified(strFilename As String)
    Dim oFS As Object
    'This creates an instance of the MS Scripting Runtime FileSystemObject class
    Set oFS = CreateObject("Scripting.FileSystemObject")
    GetDateLastModified = oFS.GetFile(strFilename).DateLastModified
    Set oFS = Nothing
End Function
Function GetDateCreated(strFilename As String)
    Dim oFS As Object
    'This creates an instance of the MS Scripting Runtime FileSystemObject class
    Set oFS = CreateObject("Scripting.FileSystemObject")
    GetDateCreated = oFS.GetFile(strFilename).DateCreated
    Set oFS = Nothing
End Function
Function GetIPFromHostName(HostName As String)
    Dim HostEnt_Addr As Long
    Dim Host As HOSTENT
    Dim HostIP_Addr As Long
    Dim Temp_IP_Address() As Byte
    Dim i As Integer
    Dim IP_Address As String
    On Error Resume Next
    HostEnt_Addr = gethostbyname(HostName)
    If HostEnt_Addr = 0 Then
        GetIPFromHostName = "Can't resolve the host name: " & HostName
        Exit Function
    End If
    RtlMoveMemory Host, HostEnt_Addr, LenB(Host)
    RtlMoveMemory HostIP_Addr, Host.hAddrList, 4
    ReDim Temp_IP_Address(1 To Host.hLength)
    RtlMoveMemory Temp_IP_Address(1), HostIP_Addr, Host.hLength
    For i = 1 To Host.hLength
        IP_Address = IP_Address & Temp_IP_Address(i) & "."
    Next
    IP_Address = Mid$(IP_Address, 1, Len(IP_Address) - 1)
    GetIPFromHostName = IP_Address
End Function
Function GetAllEnvironmentalVariables()
    Dim i As Integer
    On Error Resume Next
    GetAllEnvironmentalVariables = "Get All Environmental Variables" & vbNewLine & "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv" & " Beginning of Environmental Variables"
    For i = 1 To 27
        GetAllEnvironmentalVariables = GetAllEnvironmentalVariables & vbNewLine & i & "  " & (Environ$(i))
    Next i
    GetAllEnvironmentalVariables = GetAllEnvironmentalVariables & vbNewLine & "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^" & " End of Environmental Variables"
End Function
Sub Logger(Message As String)
    On Error Resume Next
    If DateTime = "" Then DateTime = Format(Now, "yyyymmddhhmmss")
    Open LogPath & "\" & ThisWorkbook.Name & "ErrorLog.txt" For Append As #1
    Print #1, DateTime & " _" & Now & "_:  " & Message
    Close #1
End Sub
Public Function AllRunningApps() As String
    Dim strComputer As String
    Dim objServices As Object, objProcessSet As Object, Process As Object
    Dim oDic As Object, a() As Variant
    Dim i As Integer
    On Error Resume Next
    Set oDic = CreateObject("Scripting.Dictionary")
    strComputer = "."
    Set objServices = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set objProcessSet = objServices.ExecQuery("SELECT Name FROM Win32_Process", , 48)
    For Each Process In objProcessSet
        If Not oDic.exists(Process.Name) Then oDic.Add Process.Name, Process.Name
    Next
    a() = oDic.keys
    Set objProcessSet = Nothing
    Set oDic = Nothing
    AllRunningApps = "Get All Running Tasks" & vbNewLine & "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv" & " Beginning Running Tasks"
    For i = 0 To UBound(a)
        AllRunningApps = AllRunningApps & vbNewLine & a(i)
    Next i
    AllRunningApps = AllRunningApps & vbNewLine & "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^" & "End of Running Tasks"
End Function
Function GetAllServices()
    Dim colItems  As Object
    Dim objItem As Object
    Dim strComputer As String
    Dim objWMIService As Object
    On Error Resume Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Service", , 48)
    GetAllServices = "Services" & vbNewLine & "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv" & "Beginning of Services"
    For Each objItem In colItems
        GetAllServices = GetAllServices & vbNewLine & objItem.Name & "_" & "State: " & objItem.State
    Next
    GetAllServices = GetAllServices & vbNewLine & "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^" & "End of Services"
End Function
Function GetRunningPrograms(ProgramName As String)
    Dim ProgramPath As String
    Dim MyProgramName As String
    Dim strComputer As String
    Dim colItems  As Object
    Dim objItem As Object
    Dim objWMIService As Object
    On Error Resume Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Process" & " WHERE Name = '" & ProgramName & "'" & " OR Name = 'wscript.exe'", , 48)
    For Each objItem In colItems
        GetRunningPrograms = GetRunningPrograms & vbNewLine & "-------------------------------------------"
        ProgramPath = objItem.CommandLine
        MyProgramName = objItem.Name
    Next
    If GetRunningPrograms = "" Then GetRunningPrograms = "No instances of " & ProgramName & " are running" Else GetRunningPrograms = ProgramName & " is running. " & MyProgramName & "'s path is: " & ProgramPath
End Function
Function GetUsedDiskSpace()
    Dim colDisks As Object
    Dim objDisk As Object
    Dim strComputer As String
    Dim objWMIService As Object
    On Error Resume Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")
    GetUsedDiskSpace = "Get ALL Local Drive Info" & vbNewLine & "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv" & " Beginning of Local Drives Info"
    For Each objDisk In colDisks
        If objDisk.Size <> 0 Then
            GetUsedDiskSpace = GetUsedDiskSpace & vbNewLine & "DeviceID: " & objDisk.DeviceID
            GetUsedDiskSpace = GetUsedDiskSpace & vbNewLine & "Free Disk Space: " & objDisk.FreeSpace / 1000000000 & "GB"
            GetUsedDiskSpace = GetUsedDiskSpace & vbNewLine & "Disk Size: " & objDisk.Size / 1000000000 & "GB"
        End If
    Next
    GetUsedDiskSpace = GetUsedDiskSpace & vbNewLine & "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^" & " End of Local Drives Info"
End Function
Function GetPersonLoggedInToPC()
    Dim colComputer As Object
    Dim objComputer As Object
    Dim strComputer As String
    Dim objWMIService As Object
    On Error Resume Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colComputer = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
    For Each objComputer In colComputer
        GetPersonLoggedInToPC = objComputer.UserName
    Next
End Function
Function GetFreeMemory()
    Dim objOperatingSystem As Object
    Dim colSettings As Object
    Dim strComputer As String
    Dim objWMIService As Object
    On Error Resume Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colSettings = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    For Each objOperatingSystem In colSettings
        GetFreeMemory = "Available Physical Memory: " & objOperatingSystem.FreePhysicalMemory / 1000000 & "MB"
    Next
End Function
Function GetRAM()
    Dim colComputer As Object
    Dim objComputer As Object
    Dim colSettings As Object
    Dim strComputer As String
    Dim objWMIService As Object
    On Error Resume Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
    For Each objComputer In colSettings
        GetRAM = "Total Physical Memory: " & objComputer.TotalPhysicalMemory / 1000000000 & "GB"
    Next
End Function
Function GetNumProcessors()
    Dim colComputer As Object
    Dim objComputer As Object
    Dim colSettings As Object
    Dim strComputer As String
    Dim objWMIService As Object
    On Error Resume Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
    For Each objComputer In colSettings
        GetNumProcessors = "Number of Processors: " & objComputer.NumberOfProcessors & "     Number of LogicalProcessors: " & objComputer.NumberOfLogicalProcessors
    Next
End Function
Function GetPCsDevicesNotWorking()
    Dim colItems  As Object
    Dim objItem As Object
    Dim strComputer As String
    Dim objWMIService As Object
    On Error Resume Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPEntity " & "WHERE ConfigManagerErrorCode <> 0")
    GetPCsDevicesNotWorking = "PC Device Check" & vbNewLine & "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv" & " Beginning of PC Device Check"
    For Each objItem In colItems
        GetPCsDevicesNotWorking = GetPCsDevicesNotWorking & vbNewLine & "Class GUID: " & objItem.ClassGuid
        GetPCsDevicesNotWorking = GetPCsDevicesNotWorking & vbNewLine & "Description: " & objItem.Description
        GetPCsDevicesNotWorking = GetPCsDevicesNotWorking & vbNewLine & "Device ID: " & objItem.DeviceID
        GetPCsDevicesNotWorking = GetPCsDevicesNotWorking & vbNewLine & "Manufacturer: " & objItem.Manufacturer
        GetPCsDevicesNotWorking = GetPCsDevicesNotWorking & vbNewLine & "Name: " & objItem.Name
        GetPCsDevicesNotWorking = GetPCsDevicesNotWorking & vbNewLine & "PNP Device ID: " & objItem.PNPDeviceID
        GetPCsDevicesNotWorking = GetPCsDevicesNotWorking & vbNewLine & "Service: " & objItem.Service
    Next
    If GetPCsDevicesNotWorking = "PC Device Check" & vbNewLine & "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv" & " Beginning of PC Device Check" Then
        GetPCsDevicesNotWorking = GetPCsDevicesNotWorking & vbNewLine & "All PC devices working"
    End If
    GetPCsDevicesNotWorking = GetPCsDevicesNotWorking & vbNewLine & "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^" & "End of Device Check"
End Function
Function GetPCsSpeed()
    Dim colItems  As Object
    Dim objItem As Object
    Dim strComputer As String
    Dim objWMIService As Object
    On Error Resume Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
    For Each objItem In colItems
        GetPCsSpeed = "Processor Id: " & objItem.ProcessorId & "        Maximum Clock Speed: " & objItem.MaxClockSpeed
    Next
End Function
Function GetPCType()
    Dim colChassis As Object
    Dim objChassis As Object
    Dim colItems  As Object
    Dim objItem As Variant
    Dim strComputer As String
    Dim objWMIService As Object
    On Error Resume Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colChassis = objWMIService.ExecQuery("Select * from Win32_SystemEnclosure")
    For Each objChassis In colChassis
        For Each objItem In objChassis.ChassisTypes
            GetPCType = "Chassis Type: " & objItem
        Next
    Next
End Function
Function GetPCName()
    Dim colItems  As Object
    Dim objItem As Object
    Dim strComputer As String
    Dim objWMIService As Object
    On Error Resume Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
    For Each objItem In colItems
        GetPCName = "Computer Name: " & objItem.Name
    Next
End Function
Function GetPCSerialNumAssetTag()
    Dim colSMBIOS As Object
    Dim objSMBIOS As Object
    Dim strComputer As String
    Dim objWMIService As Object
    On Error Resume Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colSMBIOS = objWMIService.ExecQuery("Select * from Win32_SystemEnclosure")
    For Each objSMBIOS In colSMBIOS
        GetPCSerialNumAssetTag = "PartNumber: " & objSMBIOS.PartNumber & "     SerialNumber: " & objSMBIOS.SerialNumber & "     AssetTag: " & objSMBIOS.SMBIOSAssetTag
    Next
End Function
Function GetPC_IP()
    Dim i As Integer
    Dim IPConfigSet As Object
    Dim IPConfig As Object
    Dim strComputer As String
    Dim objWMIService As Object
    On Error Resume Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set IPConfigSet = objWMIService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration ")
    For Each IPConfig In IPConfigSet
        If Not IsNull(IPConfig.IPAddress) Then
            If GetPC_IP = "" Then GetPC_IP = "IP address: "
            For i = LBound(IPConfig.IPAddress) To UBound(IPConfig.IPAddress)
                GetPC_IP = GetPC_IP & IPConfig.IPAddress(i)
                If i = LBound(IPConfig.IPAddress) Then GetPC_IP = GetPC_IP & "        MAC_Address: "
            Next
        End If
    Next
End Function
Function PingPC(IP As String) As String
    Dim objStatus As Object
    Dim colPings As Object
    Dim strComputer As String
    Dim objWMIService As Object
    On Error Resume Next
    If Left(UCase(IP), 3) <> "CAN" Then
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colPings = objWMIService.ExecQuery("Select * From Win32_PingStatus where Address = '" & IP & "'")
        For Each objStatus In colPings
            If IsNull(objStatus.StatusCode) Or objStatus.StatusCode <> 0 Then
                PingPC = "Unable to ping: " & IP
            Else
                PingPC = "Was able to ping computer at IP address: " & IP
            End If
        Next
    Else
        PingPC = IP
    End If
End Function
Function GetOS()
    Dim colOperatingSystems As Object
    Dim objOperatingSystem As Object
    Dim colSettings As Object
    Dim strComputer As String
    Dim objWMIService As Object
    On Error Resume Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    For Each objOperatingSystem In colOperatingSystems
        GetOS = objOperatingSystem.Caption & "  " & objOperatingSystem.Version
    Next
End Function
Sub testGetPCName()
    MsgBox (GetPCNameFromIP("192.168.0.109"))
End Sub
Function GetPCNameFromIP(IP As String)
    Dim obj As clsIPRESOLVE
    On Error Resume Next
    Set obj = New clsIPRESOLVE
    'To get the remote computer name from IP address
    GetPCNameFromIP = obj.AddressToName(IP)
    Set obj = Nothing
End Function
