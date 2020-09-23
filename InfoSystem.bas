Attribute VB_Name = "InfoSystem"
Option Explicit

' Some variables to obtain the system information

Dim systeem As New system   ' Systeem is Dutch for System
Dim Drive As String
Dim Percent As Byte
Dim Free As Long
Dim Total As Long
Dim Processor As String
Dim Number As Long
Dim Active As Long
Dim Maj As Integer
Dim Min As Integer
Dim Version As String
Dim TotalDiskSpace As Long
Dim FreeDiskSpace As Long
Dim ENTER As String

Dim Removable As Integer
Dim Fixed As Integer
Dim Ram As Integer
Dim Network As Integer
Dim CDrom As Integer
Dim C$
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
        Alias "GetDiskFreeSpaceExA" _
        (ByVal lpRootPathName As String, _
        lpFreeBytesAvailableToCaller As Currency, _
        lpTotalNumberOfBytes As Currency, _
        lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function apiGetDrives Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Function Drives()

    Dim Retrn As Long
    Dim Buffer As Long
    Dim Temp As String
    Dim intI As Integer
    Dim Read(1 To 100) As String
    Dim Counter As Integer
    Dim Hnode As Node
    Invent.Hardware.ImageList = Invent.ImageList3
    Buffer = 10


Again:
    Temp = Space$(Buffer)
    Retrn = apiGetDrives(Buffer, Temp)

    If Retrn > Buffer Then
        Buffer = Retrn
        GoTo Again
    End If

    Counter = 0
    For intI = 1 To (Buffer - 4) Step 4
        Counter = Counter + 1
        Read(Counter) = Mid$(Temp, intI, 3)
    Next


    ENTER = Chr$(13) + Chr$(10)


    C = "Disks"
    Set Hnode = Invent.Hardware.Nodes.Add("Hardware", 4, C, C, C)


    For intI = 1 To Counter

        systeem.DriveInfo Drive, TotalDiskSpace, FreeDiskSpace
        Drive = Read(intI)
        If TotalDiskSpace = Empty Then

            If systeem.DriveType(Drive) = "CD-ROM drive" Then
                Set Hnode = Invent.Hardware.Nodes.Add("Disks", 4, "Disks" & intI, Drive, "cdrom")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "DriveType - " & _
                        systeem.DriveType(Drive), "cdrom")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "TotalDiskspace - " & _
                        systeem.Full(Drive), "cdrom")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "FreeDiskSpace - " & _
                        systeem.Free(Drive), "cdrom")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "VolumeName - " & _
                        systeem.VolumeLabel(Drive), "cdrom")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "SerialNumber - " & _
                        systeem.SerialNumber(Drive), "cdrom")
            End If

            If systeem.DriveType(Drive) = "Fixed drive" Then
                Set Hnode = Invent.Hardware.Nodes.Add("Disks", 4, "Disks" & intI, Drive, "Disks")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "DriveType - " & _
                        systeem.DriveType(Drive), "Disks")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "TotalDiskspace - " & _
                        systeem.Full(Drive), "Disks")
                If Drive = "C:\" Then: Call Dodrives(Drive)

                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "FreeDiskSpace - " & _
                        systeem.Free(Drive), "Disks")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "VolumeName - " & _
                        systeem.VolumeLabel(Drive), "Disks")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "SerialNumber - " & _
                        systeem.SerialNumber(Drive), "Disks")
            End If

            If systeem.DriveType(Drive) = "Removable drive" Then
                Set Hnode = Invent.Hardware.Nodes.Add("Disks", 4, "Disks" & intI, Drive, "Diskettestation")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "DriveType - " & _
                        systeem.DriveType(Drive), "Diskettestation")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "TotalDiskspace - " & _
                        systeem.Full(Drive), "Diskettestation")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "FreeDiskSpace - " & _
                        systeem.Free(Drive), "Diskettestation")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "VolumeName - " & _
                        systeem.VolumeLabel(Drive), "Diskettestation")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "SerialNumber - " & _
                        systeem.SerialNumber(Drive), "Diskettestation")

            End If
            If systeem.DriveType(Drive) = "Network drive" Then
                Set Hnode = Invent.Hardware.Nodes.Add("Disks", 4, "Disks" & intI, Drive, "Networkdrive")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "DriveType - " & _
                        systeem.DriveType(Drive), "Networkdrive")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "TotalDiskspace - " & _
                        systeem.Full(Drive), "Networkdrive")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "FreeDiskSpace - " & _
                        systeem.Free(Drive), "Networkdrive")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "VolumeName - " & _
                        systeem.VolumeLabel(Drive), "Networkdrive")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "SerialNumber - " & _
                        systeem.SerialNumber(Drive), "Networkdrive")
            End If

        Else

            If systeem.DriveType(Drive) = "CD-ROM drive" Then
                Set Hnode = Invent.Hardware.Nodes.Add("Disks", 4, "Disks" & intI, Drive, "cdrom")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "DriveType - " & _
                        systeem.DriveType(Drive), "cdrom")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "TotalDiskspace - " & _
                        systeem.Full(Drive), "cdrom")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "FreeDiskSpace - " & _
                        systeem.Free(Drive), "cdrom")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "VolumeName - " & _
                        systeem.VolumeLabel(Drive), "cdrom")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "SerialNumber - " & _
                        systeem.SerialNumber(Drive), "cdrom")
            End If

            If systeem.DriveType(Drive) = "Fixed drive" Then
                Set Hnode = Invent.Hardware.Nodes.Add("Disks", 4, "Disks" & intI, Drive, "Disks")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "DriveType - " & _
                        systeem.DriveType(Drive), "Disks")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "TotalDiskspace - " & _
                        systeem.Full(Drive), "Disks")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "FreeDiskSpace - " & _
                        systeem.Free(Drive), "Disks")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "VolumeName - " & _
                        systeem.VolumeLabel(Drive), "Disks")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "SerialNumber - " & _
                        systeem.SerialNumber(Drive), "Disks")
            End If

            If systeem.DriveType(Drive) = "Removable drive" Then
                Set Hnode = Invent.Hardware.Nodes.Add("Disks", 4, "Disks" & intI, Drive, "Diskettestation")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "DriveType - " & _
                        systeem.DriveType(Drive), "Diskettestation")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "TotalDiskspace - " & _
                        systeem.Full(Drive), "Diskettestation")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "FreeDiskSpace - " & _
                        systeem.Free(Drive), "Diskettestation")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "VolumeName - " & _
                        systeem.VolumeLabel(Drive), "Diskettestation")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "SerialNumber - " & _
                        systeem.SerialNumber(Drive), "Diskettestation")

            End If
            If systeem.DriveType(Drive) = "Network drive" Then
                Set Hnode = Invent.Hardware.Nodes.Add("Disks", 4, "Disks" & intI, Drive, "Networkdrive")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "DriveType - " & _
                        systeem.DriveType(Drive), "Networkdrive")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "TotalDiskspace - " & _
                        systeem.Full(Drive), "Networkdrive")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "FreeDiskSpace - " & _
                        systeem.Free(Drive), "Networkdrive")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "VolumeName - " & _
                        systeem.VolumeLabel(Drive), "Networkdrive")
                Set Hnode = Invent.Hardware.Nodes.Add("Disks" & intI, 4, , "SerialNumber - " & _
                        systeem.SerialNumber(Drive), "Networkdrive")
            End If
        End If

    Next intI

End Function
Public Sub Sysinfo()
    ENTER = Chr$(13) + Chr$(10)
    systeem.FreeMemory Percent, Total, Free
    systeem.WinVer Maj, Min, Version

    Dim tCPU As udtCPU
    Call GetCPUInfo(tCPU)
    Invent.Label1(0).Caption = "Operating System": Invent.Label2(0).Caption = Version
    Invent.Label1(1).Caption = "Windows version": Invent.Label2(1).Caption = CStr(Maj) + "." + CStr(Min)
    Invent.Label1(2).Caption = "User name": Invent.Label2(2).Caption = systeem.Username
    Invent.Label1(3).Caption = "Number of Processors": Invent.Label2(3).Caption = "#" & tCPU.lNumberOfProcessors
    Invent.Label1(4).Caption = "Processor Type": Invent.Label2(4).Caption = tCPU.lProcType
    Invent.Label1(5).Caption = "Processor Model": Invent.Label2(5).Caption = tCPU.strProcRevision
    Invent.Label1(6).Caption = "Processor speed": Invent.Label2(6).Caption = tCPU.lClockSpeed
    Invent.Label1(7).Caption = "Total RAM": Invent.Label2(7).Caption = CStr(Total) + " Kb"
    Invent.Label1(8).Caption = "Free  RAM": Invent.Label2(8).Caption = CStr(Free) + " Kb"
    Invent.Label2(11).Caption = GetNetConnectString
    Invent.Label1(10).Caption = "Computername": Invent.Label2(10).Caption = Invent.Winsock1.LocalHostName
End Sub
Public Sub Dodrives(DriveLetter As String)
    Dim Maj As Integer
    Dim Min As Integer
    Dim Version As String
    Dim systeem As New system
    Dim bIsNTFS As Boolean
    systeem.WinVer Maj, Min, Version
    Version = Version


    If Version <> "Windows 95" Then

        ' On Error Resume Next
        Dim r As Long, BytesFreeToCalller As Currency, TotalBytes As Currency
        Dim TotalFreeBytes As Currency, TotalBytesUsed As Currency
        Dim TNB As String
        Dim TFB As String
        Dim FreeBytes As Long
        Dim DLetter As String
        Dim spaceInt As Integer
        spaceInt = InStr(DriveLetter, " ")
        If spaceInt > 0 Then DriveLetter = Left$(DriveLetter, spaceInt - 1)
        If Right$(DriveLetter, 1) <> "\" Then DriveLetter = DriveLetter & "\"
        DLetter = Left(UCase(DriveLetter), 1)
        Call GetDiskFreeSpaceEx(DriveLetter, BytesFreeToCalller, TotalBytes, TotalFreeBytes)
        TNB = TotalBytes * 10000
        TFB = (TotalBytes - TotalFreeBytes) * 10000
        Invent.Label3.Caption = Format$(BytesFreeToCalller * 10000, "###,###,###,##0") & " Bytes"
        Invent.Label4.Caption = Format$(TotalBytes * 10000, "###,###,###,##0") & " Bytes"
        Invent.Label5.Caption = Format(100 - TFB / TNB * 100, "###.#0") & " %"
        Invent.Label6.Caption = "Drive C:\"
        Invent.Picture6.Width = Format(100 - TFB / TNB * 100, "###.#0") * 50

        bIsNTFS = FileSystemName("C:\") = "NTFS"
        If bIsNTFS = True Then
            Invent.Label7.Caption = "NTFS"
        Else
            Invent.Label7.Caption = "FAT"
        End If


    Else

        On Error Resume Next
        TNB = Left(systeem.Full("c:\"), Len(systeem.Full("c:\")) - 6)
        TFB = Left(systeem.Free("c:\"), Len(systeem.Free("c:\")) - 6)

        Invent.Label3.Caption = Format$(TFB, "###,###,###,##0") & " Bytes"
        Invent.Label4.Caption = Format$(TNB, "###,###,###,##0") & " Bytes"
        Invent.Label5.Caption = Format(100 - TFB / TNB * 100, "###.#0") & " %"
        Invent.Label6.Caption = "Drive C:\"
        Invent.Picture6.Width = Format(100 - TFB / TNB * 100, "###.#0") * 50


        bIsNTFS = FileSystemName("C:\") = "NTFS"

        If bIsNTFS = True Then
            Invent.Label7.Caption = "NTFS"
        Else
            Invent.Label7.Caption = "FAT"
        End If


    End If
End Sub
