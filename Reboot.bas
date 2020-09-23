Attribute VB_Name = "Reboot"

Private Declare Function ExitWindowsEx Lib "user32" _
        (ByVal dwOptions As Long, _
        ByVal dwReserved As Long) As Long

Private Const EWX_LOGOFF As Long = 0
Private Const EWX_SHUTDOWN As Long = 1
Private Const EWX_REBOOT As Long = 2

Public Sub ShutDown()
    Dim llResult As Long
    llResult = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Sub
Public Sub Reboots()
    Dim llResult As Long
    llResult = ExitWindowsEx(EWX_REBOOT, 0&)
End Sub

Public Sub LogOff()
    Dim llResult As Long
    llResult = ExitWindowsEx(EWX_LOGOFF, 0&)
End Sub
