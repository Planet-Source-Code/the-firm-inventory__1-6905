Attribute VB_Name = "Drivemapper"
Option Explicit

Private Const CONNECT_UPDATE_PROFILE = &H1

Private Const RESOURCE_CONNECTED As Long = &H1&
Private Const RESOURCE_GLOBALNET As Long = &H2&
Private Const RESOURCETYPE_DISK As Long = &H1&
Private Const RESOURCEDISPLAYTYPE_SHARE& = &H3
Private Const RESOURCEUSAGE_CONNECTABLE As Long = &H1&

Private Declare Function WNetAddConnection2 Lib "mpr.dll" _
        Alias "WNetAddConnection2A" (lpNetResource As NETCONNECT, _
        ByVal lpPassword As String, ByVal lpUserName As String, _
        ByVal dwflags As Long) As Long

Private Declare Function WNetCancelConnection2 Lib "mpr.dll" _
        Alias "WNetCancelConnection2A" (ByVal lpName As String, _
        ByVal dwflags As Long, ByVal fForce As Long) As Long



Private Type NETCONNECT
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Private Declare Function GetSystemMetrics Lib "user32" (ByVal _
        nIndex As Long) As Long

Private Const SM_NETWORK = 63

Public Function MapDrive(LocalDrive As String, _
            RemoteDrive As String, Optional Username As String, _
            Optional Password As String) As Boolean

    'Example:
    'MapDrive "Q:", "\\RemoteMachine\RemoteDirectory", _
     '"MyLoginName", "MyPassword"

    Dim NetR As NETCONNECT

    NetR.dwScope = RESOURCE_GLOBALNET
    NetR.dwType = RESOURCETYPE_DISK
    NetR.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
    NetR.dwUsage = RESOURCEUSAGE_CONNECTABLE
    NetR.lpLocalName = Left$(LocalDrive, 1) & ":"
    NetR.lpRemoteName = RemoteDrive

    MapDrive = (WNetAddConnection2(NetR, Username, Password, _
            CONNECT_UPDATE_PROFILE) = 0)


End Function

Public Function UnmapDrive(LocalDrive As String) As Boolean

    UnmapDrive = WNetCancelConnection2(Left$(LocalDrive, 1) & ":", _
            CONNECT_UPDATE_PROFILE, False) = 0

End Function
Public Function IsNetInstalled() As Boolean
    Dim IsNetworkInstalled

    IsNetworkInstalled = GetSystemMetrics(SM_NETWORK)
End Function

