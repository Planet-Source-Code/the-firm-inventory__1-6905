Attribute VB_Name = "CPU"

Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string For PSS usage
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const PROCESSOR_INTEL_386 = 386
Private Const PROCESSOR_INTEL_486 = 486
Private Const PROCESSOR_INTEL_PENTIUM = 586
Private Const PROCESSOR_LEVEL_80386 As Long = 3
Private Const PROCESSOR_LEVEL_80486 As Long = 4
Private Const PROCESSOR_LEVEL_PENTIUM As Long = 5
Private Const PROCESSOR_LEVEL_PENTIUMII As Long = 6
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Type udtCPU
    lClockSpeed As Variant
    lProcType As Integer
    strProcLevel As String
    strProcRevision As String
    lNumberOfProcessors As Long
End Type
Public Enum eVersion
    eWindowsNT = 1
    eWindows95_98 = 2
    eUnknown = 3
End Enum

Public Function GetCPUInfo(ptCPUInfo As udtCPU)
    Dim tSYS As SYSTEM_INFO
    Dim intProcType As Integer
    Dim strProcLevel As String
    Dim strProcRevision As String
    Call GetSystemInfo(tSYS)
    Select Case tSYS.dwProcessorType
        Case PROCESSOR_INTEL_386: intProcType = 386
        Case PROCESSOR_INTEL_486: intProcType = 486
        Case PROCESSOR_INTEL_PENTIUM: intProcType = 586
    End Select
    Select Case tSYS.wProcessorLevel
        Case PROCESSOR_LEVEL_80386: strProcLevel = "Intel 80386"
        Case PROCESSOR_LEVEL_80486: strProcLevel = "Intel 80486"
        Case PROCESSOR_LEVEL_PENTIUM: strProcLevel = "Intel Pentium"
        Case PROCESSOR_LEVEL_PENTIUMII: strProcLevel = "Intel Pentium Pro or Pentium II"
    End Select
    strProcRevision = "Model " & HiByte(tSYS.wProcessorRevision) & ", Stepping " & LoByte(tSYS.wProcessorRevision)
    With ptCPUInfo
        .lClockSpeed = GetCPUSpeed
        .lNumberOfProcessors = tSYS.dwNumberOfProcessors
        .lProcType = intProcType
        .strProcLevel = IIf(strProcLevel = "", "None", strProcLevel)
        .strProcRevision = IIf(strProcRevision = "", "None", strProcRevision)
    End With
End Function

Public Function HiByte(ByVal wParam As Integer) As Byte
    HiByte = (wParam And &HFF00&) \ (&H100)
End Function
Public Function LoByte(ByVal wParam As Integer) As Byte
    LoByte = wParam And &HFF&
End Function
Private Function GetCPUSpeed() As Variant
    Dim hKey As Long
    Dim lClockSpeed As Long
    Dim strKey As String
    Dim Maj As Integer
    Dim Min As Integer
    Dim Version As String
    Dim systeem As New system
    systeem.WinVer Maj, Min, Version
    If Version = "Windows 2000" Or Version = "Windows NT" Then  ' tnx to pt@sonic.net
        strKey = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
        Call RegOpenKey(HKEY_LOCAL_MACHINE, strKey, hKey)
        Call RegQueryValueEx(hKey, "~MHz", 0, 0, lClockSpeed, 4)
        Call RegCloseKey(hKey)
        GetCPUSpeed = lClockSpeed & " MHZ"
    Else
        GetCPUSpeed = "Could Not be determined"
    End If
End Function

