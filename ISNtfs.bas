Attribute VB_Name = "ISNtfs"
Option Explicit

Public Declare Function GetVolumeInformation Lib _
        "kernel32.dll" Alias "GetVolumeInformationA" _
        (ByVal lpRootPathName As String, _
        ByVal lpVolumeNameBuffer As String, _
        ByVal nVolumeNameSize As Integer, _
        lpVolumeSerialNumber As Long, _
        lpMaximumComponentLength As Long, _
        lpFileSystemFlags As Long, _
        ByVal lpFileSystemNameBuffer As String, _
        ByVal nFileSystemNameSize As Long) As Long

Public Function FileSystemName(ByVal Drive As String) As String

    'usage:
    Dim lAns As Long
    Dim lRet As Long
    Dim sVolumeName As String, sDriveType As String
    Dim sDrive As String
    Dim iPos As Integer

    'Deal with one and two character input values
    'Dim bIsNTFS As Boolean
    'bISNTFS = FileSystemName("C:\") = "NTFS"

    sDrive = Drive
    If Len(sDrive) = 1 Then
        sDrive = sDrive & ":\"
    ElseIf Len(sDrive) = 2 And Right(sDrive, 1) = ":" Then
        sDrive = sDrive & "\"
    End If

    sVolumeName = String$(255, Chr$(0))
    sDriveType = String$(255, Chr$(0))

    lRet = GetVolumeInformation(sDrive, sVolumeName, _
            255, lAns, 0, 0, sDriveType, 255)
    iPos = InStr(sDriveType, Chr$(0))
    If iPos > 0 Then sDriveType = Left(sDriveType, iPos - 1)

    FileSystemName = sDriveType
End Function


