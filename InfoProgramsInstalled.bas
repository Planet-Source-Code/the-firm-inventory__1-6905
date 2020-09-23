Attribute VB_Name = "InfoProgramsInstalled"
Dim Count As Integer
Dim returnName As Collection
Dim returnSubs As Collection
Dim DisplayName As String
Dim UninstallString As String
Dim Version As String

Public Function Programs()


    Invent.lsprograms.Clear
    Call EnumRegKeys(returnName, returnSubs, "HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
    If returnName.Count > 0 Then
        For Count = 1 To returnName.Count
            DisplayName = GetSetting("", returnName(Count), "DisplayName", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
            UninstallString = GetSetting("", returnName(Count), "UninstallString", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
            Version = GetSetting("", returnName(Count), "Version", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
            If DisplayName <> "" And (UninstallString <> "" Or Version <> "") Then
                Call Invent.lsprograms.AddItem(DisplayName)
            End If
        Next Count
    End If
End Function
