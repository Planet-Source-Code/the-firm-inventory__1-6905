Attribute VB_Name = "InfoHardware"

Option Explicit
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Dim m_clsRegAccess As Registry

Dim p_vntRtn As Variant
Dim p_vntAdapters As Variant
Dim p_strSubKey As String
Dim p_strValueName As String
Dim p_lngNumAdapters As Long
Dim p_lngLoop As Long
Dim p_lngPos As Long
Dim p_strAdapterName As String
Dim p_strTmp As String
Dim p_blnFirstTime As Boolean
Dim Result As String
Public Enum Hardware
    Printer
    Modem
    GraphicCard
    Soundcard
    Mouse
    Display
    Monitor
End Enum
Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Private Function sOund() As String

    Dim I As Integer
    I = waveOutGetNumDevs()
    If I > 0 Then
        sOund = "Yes"
    Else
        sOund = "No"
    End If
End Function
Public Function HardInfo(Hardware As Hardware) As String
    Dim Maj As Integer
    Dim Min As Integer
    Dim Version As String
    Dim systeem As New system

    systeem.WinVer Maj, Min, Version
    Version = Version
    If Version = "Windows 98" Then Hardinfo98 (Hardware)
    HardInfo = Result
    If Version = "Windows 95" Then Hardinfo95 (Hardware)
    HardInfo = Result
    If Version = "Windows NT" Then HardinfoNT (Hardware)
    HardInfo = Result
    If Version = "Windows 2000" Then Hardinfo2000 (Hardware)
    HardInfo = Result
End Function
Private Function HardinfoNT(Hardware As Hardware) As String
    Result = ""
    Select Case Hardware
        Case Printer

            Set m_clsRegAccess = New Registry
            'p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            On Error Resume Next
            p_strSubKey = "SYSTEM\ControlSet001\Control\Print\Printers\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "Printer Driver"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\ControlSet001\Control\Print\Printers\" & p_strAdapterName
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "|" & CStr(p_vntRtn)

                End If
            Next p_lngLoop
        Case Modem
            Dim test
            Set m_clsRegAccess = New Registry
            On Error Resume Next
            p_strSubKey = "SYSTEM\CurrentControlset\Control\class\{4D36E96D-E325-11CE-BFC1-08002BE10318}\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "Model"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlset\Control\class\{4D36E96D-E325-11CE-BFC1-08002BE10318}\" _
                        & p_strAdapterName
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & ", " & CStr(p_vntRtn)
                End If
            Next p_lngLoop

            Dim X As Integer
            Dim tel As Integer
            Dim A As String
            Dim v As String
            v = Result
            Result = ""
            For X = 1 To Len(v)

                A = Mid(v, X, 1)
                If A = "," Or X = Len(v) Then
                    If Result = "" Then
                        Result = Mid(v, 1, tel)

                        tel = 0
                    Else
                        test = Mid(v, X + 1, 1)
                        If test = " " Then
                            Result = Result & "|" & Mid(v, X - tel + 2, tel - 2)
                        Else
                            test = Mid(v, X - tel + 1, 1)
                            If test = " " Then
                                Result = Result & "|" & Mid(v, X - tel + 2, tel - 1)
                            End If
                        End If
                        tel = 0
                    End If
                End If
                tel = tel + 1
            Next X
        Case GraphicCard
            Set m_clsRegAccess = New Registry
            On Error Resume Next
            p_strSubKey = "SYSTEM\CurrentControlset\Control\class\{4D36E968-E325-11CE-BFC1-08002BE10318}\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "DriverDesc"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlset\Control\class\{4D36E968-E325-11CE-BFC1-08002BE10318}\" _
                        & p_strAdapterName
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "  |  " & CStr(p_vntRtn)
                End If
            Next p_lngLoop
            Result = CStr(p_vntRtn)
        Case Soundcard
            p_strValueName = "DisplayName"
            p_strSubKey = "SYSTEM\CurrentControlset\services\sbpcint4\"
            p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            Result = CStr(p_vntRtn)
        Case Mouse
            On Error Resume Next
            p_strSubKey = "SYSTEM\CurrentControlset\Control\class\{4D36E96F-E325-11CE-BFC1-08002BE10318}\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "DriverDesc"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlset\Control\class\{4D36E96F-E325-11CE-BFC1-08002BE10318}\" _
                        & p_strAdapterName
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "  |  " & CStr(p_vntRtn)
                End If

            Next p_lngLoop
            Result = CStr(p_vntRtn)
        Case Monitor
    End Select
End Function

Private Function Hardinfo98(Hardware As Hardware) As String
    Result = ""
    Select Case Hardware
        Case Printer

            Set m_clsRegAccess = New Registry
            'p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            On Error Resume Next
            p_strSubKey = "SYSTEM\currentControlSet\Control\Print\Printers\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "Printer Driver"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\currentControlSet\Control\Print\Printers\" & p_strAdapterName & "\"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "|" & CStr(p_vntRtn)
                End If
            Next p_lngLoop
        Case Modem
            Set m_clsRegAccess = New Registry
            'p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            On Error Resume Next
            p_strSubKey = "SYSTEM\Currentcontrolset\Services\class\Modem\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strValueName = "Model"
                p_strSubKey = "SYSTEM\CurrentControlSet\Services\Class\Modem\" & p_strAdapterName & "\"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    If p_vntRtn = "" Then
                        Result = Result
                    Else
                        Result = CStr(p_vntRtn)
                    End If
                Else
                    If p_vntRtn = "" Then
                        Result = Result
                    Else
                        Result = Result & "|" & CStr(p_vntRtn)
                        p_vntRtn = ""
                    End If
                End If
            Next p_lngLoop
        Case GraphicCard
            Set m_clsRegAccess = New Registry
            'p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            On Error Resume Next
            p_strSubKey = "SYSTEM\Currentcontrolset\Services\class\Display\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "DriverDesc"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlSet\Services\Class\Display\" & p_strAdapterName & "\"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "|" & CStr(p_vntRtn)
                End If
            Next p_lngLoop
        Case Soundcard
            Set m_clsRegAccess = New Registry

            On Error Resume Next
            p_strSubKey = "SYSTEM\Currentcontrolset\Control\MediaResources\aux\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "Description"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\Currentcontrolset\Control\MediaResources\aux\" & p_strAdapterName & "\"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "|" & CStr(p_vntRtn)
                End If
            Next p_lngLoop

        Case Mouse
            Set m_clsRegAccess = New Registry
            'p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            On Error Resume Next
            p_strSubKey = "SYSTEM\Currentcontrolset\Services\class\mouse\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "DriverDesc"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlSet\Services\Class\mouse\" & p_strAdapterName & "\"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "|" & CStr(p_vntRtn)
                End If
            Next p_lngLoop
        Case Monitor
            Set m_clsRegAccess = New Registry

            On Error Resume Next
            p_strSubKey = "SYSTEM\currentControlSet\Services\Class\Monitor"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "DriverDesc"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\currentControlSet\Services\Class\Monitor\" & p_strAdapterName & "\"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "|" & CStr(p_vntRtn)
                End If
            Next p_lngLoop
    End Select
End Function

Private Function Hardinfo95(Hardware As Hardware) As String
    Result = ""
    On Error Resume Next
    Select Case Hardware
        Case Printer

            Set m_clsRegAccess = New Registry

            On Error Resume Next
            p_strSubKey = "System\CurrentControlSet\Control\Print\Printers\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strValueName = "Printer Driver"

                p_strSubKey = "System\CurrentControlSet\Control\Print\Printers\" & p_strAdapterName & "\"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)

                Else
                    Result = Result & "|" & CStr(p_vntRtn)

                End If
            Next p_lngLoop


        Case Soundcard
            Set m_clsRegAccess = New Registry


            p_strSubKey = "SYSTEM\Currentcontrolset\Control\MediaResources\aux\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "Description"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\Currentcontrolset\Control\MediaResources\aux\" & p_strAdapterName & "\"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "|" & CStr(p_vntRtn)
                End If
            Next p_lngLoop
        Case Mouse
            Set m_clsRegAccess = New Registry

            p_strSubKey = "SYSTEM\Currentcontrolset\Services\class\mouse\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "DriverDesc"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlSet\Services\Class\mouse\" & p_strAdapterName
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "|" & CStr(p_vntRtn)
                End If
            Next p_lngLoop
        Case Modem
            Set m_clsRegAccess = New Registry

            p_strSubKey = "SYSTEM\Currentcontrolset\Services\class\Modem\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "Model"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlSet\Services\Class\Modem\" & p_strAdapterName

                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)

                Else
                    Result = Result & "|" & CStr(p_vntRtn)

                End If
            Next p_lngLoop
        Case GraphicCard
            Set m_clsRegAccess = New Registry
            'p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            On Error Resume Next
            p_strSubKey = "SYSTEM\Currentcontrolset\Services\class\Display\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "DriverDesc"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlSet\Services\Class\Display\" & p_strAdapterName & "\"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "|" & CStr(p_vntRtn)
                End If
            Next p_lngLoop
        Case Monitor
            Set m_clsRegAccess = New Registry

            On Error Resume Next
            p_strSubKey = "SYSTEM\currentControlSet\Services\Class\Monitor"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "DriverDesc"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\currentControlSet\Services\Class\Monitor\" & p_strAdapterName & "\"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "|" & CStr(p_vntRtn)
                End If
            Next p_lngLoop

    End Select
End Function


Private Function Hardinfo2000(Hardware As Hardware) As String
    Result = ""
    Select Case Hardware
        Case Printer

            Set m_clsRegAccess = New Registry
            'p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            On Error Resume Next
            p_strSubKey = "SYSTEM\currentControlSet\Control\Print\Printers\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "Printer Driver"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\currentControlSet\Control\Print\Printers\" & p_strAdapterName & "\"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "  |  " & CStr(p_vntRtn)
                End If
            Next p_lngLoop
        Case Modem
            Set m_clsRegAccess = New Registry
            'p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            On Error Resume Next
            p_strSubKey = "SYSTEM\Currentcontrolset\Services\class\Modem\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "Model"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlSet\Services\Class\Modem\" & p_strAdapterName & "\"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "  |  " & CStr(p_vntRtn)
                End If
            Next p_lngLoop
        Case GraphicCard
            Set m_clsRegAccess = New Registry
            'p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            On Error Resume Next
            p_strSubKey = "SYSTEM\Currentcontrolset\Services\class\Display\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "DriverDesc"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlSet\Services\Class\Display\" & p_strAdapterName & "\"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "  |  " & CStr(p_vntRtn)
                End If
            Next p_lngLoop
        Case Soundcard
            Set m_clsRegAccess = New Registry

            On Error Resume Next
            p_strSubKey = "SYSTEM\Currentcontrolset\Control\MediaResources\aux\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "Description"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\Currentcontrolset\Control\MediaResources\aux\" & p_strAdapterName & "\"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "  |  " & CStr(p_vntRtn)
                End If
            Next p_lngLoop

        Case Mouse
            Set m_clsRegAccess = New Registry
            'p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            On Error Resume Next
            p_strSubKey = "SYSTEM\Currentcontrolset\Services\class\mouse\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strValueName = "DriverDesc"
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlSet\Services\Class\mouse\" & p_strAdapterName & "\"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_lngLoop = 1 Then
                    Result = CStr(p_vntRtn)
                Else
                    Result = Result & "  |  " & CStr(p_vntRtn)
                End If
            Next p_lngLoop
        Case Monitor
    End Select
End Function
