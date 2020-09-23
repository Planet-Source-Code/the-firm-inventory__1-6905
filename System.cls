VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------
' API declarations in order to obtain names (computer and user names)
'---------------------------------------------------------------------------
Private Declare Function apiUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function apiCompName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


'---------------------------------------------------------------------------
' API declarations in order to obtain the Windows / System and Temporary
' directories.
'---------------------------------------------------------------------------
Private Declare Function apiWindDir Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function apiSysDir Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function apiTempDir Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


'---------------------------------------------------------------------------
' API declarations in order to obtain Memory and System information.
'---------------------------------------------------------------------------
Private Declare Sub apiMemStatus Lib "kernel32" Alias "GlobalMemoryStatus" (lpBuffer As MEMORYSTATUS)
Private Type MEMORYSTATUS                   ' size of 'Type' = 8 x 4 bytes = 32 (a Long is 4 Bytes)
    dwLength As Long                    ' This need to be set at the size of this 'Type'  = 32
    dwMemoryLoad As Long                ' Gives global indication of used RAM (in %)
    dwTotalPhys As Long                 ' Gives total RAM of the computer
    dwAvailPhys As Long                 ' Gives the amount of free RAM
    dwTotalPageFile As Long             ' I don't use this (don't know what it means)
    dwAvailPageFile As Long             ' I don't use this (don't know what it means)
    dwTotalVirtual As Long              ' I don't use this (don't know what it means)
    dwAvailVirtual As Long              ' I don't use this (don't know what it means)
End Type

Private Declare Sub apiSystemInfo Lib "kernel32" Alias "GetSystemInfo" (lpSystemInfo As SYSTEM_INFO)
Private Type SYSTEM_INFO                    ' size of 'Type' = 9 x 4 bytes = 36
    dwOemID As Long
    dwPageSize As Long                  ' Must be set at the size of this 'Type'
    lpMinimumApplicationAddress As Long ' ?
    lpMaximumApplicationAddress As Long ' ?
    dwActiveProcessorMask As Long       ' Gives the active processor number
    dwNumberOfProcessors As Long        ' Gives number of processors
    dwProcessorType As Long             ' Gives the processor type (386,486,586)
    dwAllocationGranularity As Long     ' ?
    dwReserved As Long                  ' ?
End Type


'---------------------------------------------------------------------------
' API declaration to get information about the drives.
'---------------------------------------------------------------------------
Private Declare Function apiDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function apiDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Private Declare Function apiFastFreeSpace Lib "STKIT432.DLL" Alias "DISKSPACEFREE" () As Long
Private Declare Function apiGetDrives Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function apiSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Private Declare Function apiMonitors Lib "winspool.drv" Alias "EnumMonitorsA" (ByVal pName As String, ByVal Level As Long, pMonitors As Byte, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Type MONITOR_INFO_1
    pName As String
End Type
Private Type MONITOR_INFO_2
    pName As String
    pEnvironment As String
    pDLLName As String
End Type

Dim r As Long, BytesFreeToCalller As Currency, TotalBytes As Currency
Dim TotalFreeBytes As Currency, TotalBytesUsed As Currency
Dim TNB As Double
Dim TFB As Double
Dim FreeBytes As Long
Dim DriveLetter As String
Dim DLetter As String
Dim spaceInt As Integer

Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
        Alias "GetDiskFreeSpaceExA" _
        (ByVal lpRootPathName As String, _
        lpFreeBytesAvailableToCaller As Currency, _
        lpTotalNumberOfBytes As Currency, _
        lpTotalNumberOfFreeBytes As Currency) As Long


Private Declare Function GetDiskFreeSpace Lib "kernel32" _
        Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
        lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
        lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters _
        As Long) As Long

'---------------------------------------------------------------------------
' API declarations to obtain the Windows version and the type of
' keyboard.
'---------------------------------------------------------------------------
Private Declare Function apiGetVersion Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    ' size of 'Type' = (5 x 4 bytes) =  20 bytes (the 5 Longs)
    '                                  128 bytes (fixed-length string)
    '                                 ----- +
    '                                  148 bytes

    dwOSVersionInfoSize As Long         ' Has to be set to size of 'type'= 148
    dwMajorVersion As Long              ' Gives the Major version
    dwMinorVersion As Long              ' Gives the Minor version
    dwBuildNumber As Long               ' Gives the buildnumber (I don't use it)
    dwPlatformId As Long                ' Gives the operating system.
    szCSDVersion As String * 128        ' ?
End Type

Private Declare Function apiKeyboardType Lib "user32" Alias "GetKeyboardType" (ByVal nTypeFlag As Long) As Long
Public Sub DriveInfo(ByVal strRoot As String, ByRef lngTotalSpace As Long, ByRef lngFreeSpace As Long)

    Dim TotalClusters As Long
    Dim FreeClusters As Long
    Dim SectorsPerCluster As Long
    Dim BytesPerSector As Long


    'If apiDiskFreeSpace(strRoot, SectorsPerCluster, BytesPerSector, FreeClusters, TotalClusters) Then

    lngTotalSpace = SectorsPerCluster * BytesPerSector * (TotalClusters \ 1024) \ 1024
    lngFreeSpace = SectorsPerCluster * BytesPerSector * (FreeClusters \ 1024) \ 1024

    'End If

End Sub


Public Sub Drives(ByRef intRemovable As Integer, ByRef intNotRemovable As Integer, ByRef intCD As Integer, ByRef intRAM As Integer, ByRef intNetwork As Integer)

    '
    Dim Retrn As Long
    Dim Buffer As Long
    Dim Temp As String
    Dim intI As Integer
    Dim Read As String
    Dim Counter As Integer
    Buffer = 10

Again:
    Temp = Space$(Buffer)
    Retrn = apiGetDrives(Buffer, Temp)
    ' Call the API function.

    If Retrn > Buffer Then ' If the API returned a value that is bigger than Buffer,
        Buffer = Retrn     ' than the Buffer isn't big enough to hold the information.
        GoTo Again         ' In that case adjust the Buffer to the right size (returned by
    End If                 ' the API) and try again.


    ' The API returns something like :
    ' A:\*B:\*C:\*D:\**  , with  * = NULL character
    ' 1234123412341234
    ' \ 1 \ 2 \ 3 \ 4 \
    '
    ' So we start reading three characters, we step 4 further (the three we read + the
    ' NULL-character), and we read again three characters, step 4, ect.

    Counter = 0
    For intI = 1 To (Buffer - 4) Step 4

        Counter = Counter + 1
        Read = Mid$(Temp, intI, 3)

        Select Case DriveType(Read)
            Case "Removable drive"
                intRemovable = intRemovable + 1
            Case "Fixed drive"
                intNotRemovable = intNotRemovable + 1
            Case "Network drive"
                intNetwork = intNetwork + 1
            Case "CD-ROM drive"
                intCD = intCD + 1
            Case "RAM-disk"
                intRAM = intRAM + 1
        End Select

    Next

End Sub

Public Function DriveType(ByVal strRoot As String) As String
    '---------------------------------------------------------------------------
    ' FUNCTION: DriveType
    '
    ' This function returns information about the drive you asked for. It will
    ' return whether the drive is a Removable drive, a non-removable (fixed)
    ' drive, a CD-ROM drive, a RAM drive or a Network drive.
    '
    ' IN:  strRoot      - String containing the root of a drive. (e.g. "C:\")
    '
    ' OUT: DriveType    - String containing type of drive.
    '
    ' If the function fails a empty string is returned.
    '
    ' You can also re-program this Function so that it doens't return a string,
    ' but it returns the value. That can be easier if you want to work with
    ' the returned information. I let it return a string, so that I can print
    ' it.
    '
    ' THE DRIVES FUNCTION USES THIS FUNCTION, SO IF YOU MODIFY THIS FUNCTION,
    ' YOU ALSO HAVE TO MODIFY THAT FUNCTION!
    '
    '---------------------------------------------------------------------------
    '
    Dim lngType As Long
    Const DRIVE_CDROM = 5       ' Some API constants required to
    Const DRIVE_FIXED = 3       ' get the difference between the
    Const DRIVE_RAMDISK = 6     ' drive types.
    Const DRIVE_REMOTE = 4
    Const DRIVE_REMOVABLE = 2

    lngType = apiDriveType(strRoot)
    ' The API returns a value in lngType. Use the Constants to
    ' make the strings.

    Select Case lngType
        Case DRIVE_REMOVABLE
            DriveType = "Removable drive"
        Case DRIVE_FIXED
            DriveType = "Fixed drive"
        Case DRIVE_REMOTE
            DriveType = "Network drive"
        Case DRIVE_CDROM
            DriveType = "CD-ROM drive"
        Case DRIVE_RAMDISK
            DriveType = "RAM-disk"
        Case Else
            DriveType = ""   ' If the API returns an error, we return a empty string
    End Select

End Function



Public Function FastDiskSpace(ByVal strRoot As String) As Long

    '
    Dim strCurrent As String
    On Error GoTo Bliep              '(Dutch variation of Beep, means an error)

    strCurrent = CurDir              ' Save the current drive
    ChDrive strRoot                  ' Change to the requested drive

    FastDiskSpace = apiFastFreeSpace 'Get the free space

    ChDrive Left$(strCurrent, 2)     ' Return to the saved drive
    ChDir strCurrent                 ' Return to the saved directory

    Exit Function

Bliep:
    ' If the drive wasn't ready or something
    FastDiskSpace = 0            ' Return zero as free disk space
    ChDrive Left$(strCurrent, 2) ' Retur to the saved drive
    ChDir strCurrent             ' and directory

End Function

Public Sub FreeMemory(ByRef btePercentUsed As Byte, ByRef lngTotalRam As Long, ByRef lngFreeRam As Long)

    Dim Memory As MEMORYSTATUS

    Memory.dwLength = 32
    ' This must be set to the size of the structure before the call

    apiMemStatus Memory
    ' Call the API. This function fills the Memory structure (Type) with
    ' a lot of information. I only use three parts of it.

    ' Fill the variables with the desired values.
    btePercentUsed = Memory.dwMemoryLoad
    lngTotalRam = Memory.dwTotalPhys / 1024
    lngFreeRam = Memory.dwAvailPhys / 1024

End Sub
Public Function FunctionKeys() As Byte

    FunctionKeys = apiKeyboardType(2)

End Function
Public Function KeyboardType() As String
    '---------------------------------------------------------------------------
    ' FUNCTION: KeyboardType
    '
    ' Returns a string containing the type of Keyboard you use.
    '
    '---------------------------------------------------------------------------
    '
    Dim intBuffer As Long

    intBuffer = apiKeyboardType(0)
    ' Call the API. The zero specifies that I want to get information
    ' about the keyboard type. The FunctionKeys function uses the
    ' same API call, only with a "2" specified.


    ' These are all constants I've found somewhere, so I can't explain
    ' it. It's just true.
    Select Case intBuffer
        Case 1
            KeyboardType = "IBM PC/XT or compatible (83 key)"
        Case 2
            KeyboardType = "Olivetti ""ico"" (102 key)"
        Case 3
            KeyboardType = "IBM PC/AT or compatible (84 key)"
        Case 4
            KeyboardType = "IBM enhanced (101 or 102 key)"
        Case 5
            KeyboardType = "Nokia 1050 or compatible"
        Case 6
            KeyboardType = "Nokia 9140 or compatible"
        Case 7
            KeyboardType = "Japanese"
    End Select

End Function




Private Function StripTerminator(ByVal strString As String) As String

    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If

End Function
Public Function Username() As String

    Dim Bufstr As String
    Bufstr = Space$(50)

    If apiUserName(Bufstr, 50) > 0 Then
        Username = Bufstr
        Username = RTrim(Username)
        Username = StripTerminator(Username)
    Else
        Username = ""
    End If

End Function
Public Function SerialNumber(ByVal strRoot As String) As String

    Dim VolLabel As String
    Dim VolSize As Long
    Dim SerNum As Long
    Dim MaxLen As Long
    Dim FLAGS As Long
    Dim Name As String
    Dim NameSize As Long
    Dim Check As String

    If apiSerialNumber(strRoot, VolLabel, VolSize, SerNum, MaxLen, FLAGS, Name, NameSize) Then


        Check = Format(Hex(SerNum), "00000000")


        SerialNumber = Left$(Check, 4) + "-" + Right$(Check, 4)

    Else

        SerialNumber = "0000-0000"

    End If

End Function

Public Function VolumeLabel(ByVal strRoot As String) As String

    On Error GoTo Further

    VolumeLabel = Dir(strRoot, vbVolume)
    VolumeLabel = StripTerminator(VolumeLabel)


    If VolumeLabel = "" Then VolumeLabel = "NoName"

    Exit Function


Further:
    VolumeLabel = ""


End Function

Public Function WinVer(ByRef intMajor As Integer, ByRef intMinor As Integer, ByRef strPlatform As String)

    Dim OSystem As OSVERSIONINFO

    OSystem.dwOSVersionInfoSize = 148


    If apiGetVersion(OSystem) Then


        intMajor = OSystem.dwMajorVersion
        intMinor = OSystem.dwMinorVersion
        Select Case OSystem.dwPlatformId
            Case 0
                strPlatform = "Windows " + CStr(intMajor) + "." + CStr(intMinor)
            Case 1
                If CStr(intMajor) + "." + CStr(intMinor) = "4.10" Then
                    strPlatform = "Windows 98"
                Else
                    strPlatform = "Windows 95"
                End If
            Case 2
                If CStr(intMajor) + "." + CStr(intMinor) = "5.0" Then
                    strPlatform = "Windows 2000"
                Else
                    strPlatform = "Windows NT"
                End If

        End Select

    Else

        intMajor = 0
        intMinor = 0
        strPlatform = ""

    End If

End Function

Public Function Full(DriveLetter As String) As String
    Dim Maj As Integer
    Dim Min As Integer
    Dim Version As String
    Dim systeem As New system
    Dim lAns As Long
    Dim lSectorsPerCluster As Long
    Dim lBytesPerSector As Long

    Dim lFreeClusters As Long
    Dim lTotalClusters As Long
    Dim lBytesPerCluster As Long
    Dim lFreeBytes As Double

    systeem.WinVer Maj, Min, Version

    If Version = "Windows 95" Then

        If Len(DriveLetter) = 1 Then DriveLetter = DriveLetter & ":\"
        If Len(DriveLetter) = 2 And Right$(DriveLetter, 1) = ":" _
                Then DriveLetter = DriveLetter & "\"
        lAns = GetDiskFreeSpace(DriveLetter, lSectorsPerCluster, _
                lBytesPerSector, lFreeClusters, lTotalClusters)
        lBytesPerCluster = lSectorsPerCluster * lBytesPerSector
        lTotalClusters = lBytesPerCluster * lTotalClusters
        Full = Format$(lTotalClusters / 1024, "###,###,###,##0") & " bytes"

    Else

        spaceInt = InStr(DriveLetter, " ")
        If spaceInt > 0 Then DriveLetter = Left$(DriveLetter, spaceInt - 1)
        If Right$(DriveLetter, 1) <> "\" Then DriveLetter = DriveLetter & "\"
        DLetter = Left(UCase(DriveLetter), 1)
        Call GetDiskFreeSpaceEx(DriveLetter, BytesFreeToCalller, TotalBytes, TotalFreeBytes)
        TNB = TotalBytes * 10000
        TFB = (TotalBytes - TotalFreeBytes) * 10000
        Full = Format$(TotalBytes * 10000, "###,###,###,##0") & " bytes"

    End If

End Function

Public Function Free(DriveLetter As String) As String
    Dim Maj As Integer
    Dim Min As Integer
    Dim Version As String
    Dim systeem As New system
    Dim lAns As Long
    Dim lSectorsPerCluster As Long
    Dim lBytesPerSector As Long

    Dim lFreeClusters As Long
    Dim lTotalClusters As Long
    Dim lBytesPerCluster As Long
    Dim lFreeBytes As Double

    systeem.WinVer Maj, Min, Version

    If Version = "Windows 95" Then

        If Len(DriveLetter) = 1 Then DriveLetter = DriveLetter & ":\"
        If Len(DriveLetter) = 2 And Right$(DriveLetter, 1) = ":" _
                Then DriveLetter = DriveLetter & "\"

        lAns = GetDiskFreeSpace(DriveLetter, lSectorsPerCluster, _
                lBytesPerSector, lFreeClusters, lTotalClusters)

        lBytesPerCluster = lSectorsPerCluster * lBytesPerSector
        lFreeBytes = lBytesPerCluster * lFreeClusters

        Free = Format$(lFreeBytes / 1024, "###,###,###,##0") & " bytes"


    Else
        spaceInt = InStr(DriveLetter, " ")
        If spaceInt > 0 Then DriveLetter = Left$(DriveLetter, spaceInt - 1)
        If Right$(DriveLetter, 1) <> "\" Then DriveLetter = DriveLetter & "\"
        DLetter = Left(UCase(DriveLetter), 1)
        Call GetDiskFreeSpaceEx(DriveLetter, BytesFreeToCalller, TotalBytes, TotalFreeBytes)
        TNB = TotalBytes * 10000
        TFB = (TotalBytes - TotalFreeBytes) * 10000
        Free = Format$(BytesFreeToCalller * 10000, "###,###,###,##0") & " bytes"
    End If
End Function

Public Function Used(DriveLetter As String) As String
    Dim Maj As Integer
    Dim Min As Integer
    Dim Version As String
    Dim systeem As New system


    systeem.WinVer Maj, Min, Version

    If Version = "Windows 95" Then

        TNB = Full(DriveLetter) * 10000
        TFB = Full(DriveLetter) - Free(DriveLetter) * 10000
        Used = Format(TFB / TNB * 100, "###.#0") & " %"
    Else

        spaceInt = InStr(DriveLetter, " ")
        If spaceInt > 0 Then DriveLetter = Left$(DriveLetter, spaceInt - 1)
        If Right$(DriveLetter, 1) <> "\" Then DriveLetter = DriveLetter & "\"
        DLetter = Left(UCase(DriveLetter), 1)
        Call GetDiskFreeSpaceEx(DriveLetter, BytesFreeToCalller, TotalBytes, TotalFreeBytes)
        TNB = TotalBytes * 10000
        TFB = (TotalBytes - TotalFreeBytes) * 10000
        Used = Format(TFB / TNB * 100, "###.#0") & " %"
        'Form1.Label4.Caption = "Disk space available:  " & Format(100 - TFB / TNB * 100, "###.#0") & " %"
        'Label5.Caption = Format(100 - TFB / TNB * 100, "###.#0") & " % of free space"
        'Picture1.Width = Format(100 - TFB / TNB * 100, "###.#0") * 50
    End If
End Function








