Attribute VB_Name = "InfoNet"

Option Explicit
'MacAddress
Private Const NAME_FLAGS_MASK = &H87
Private Const GROUP_NAME = &H80
Private Const UNIQUE_NAME = &H0
Private Const REGISTERING = &H0
Private Const REGISTERED = &H4
Private Const DEREGISTERED = &H5
Private Const DUPLICATE = &H6
Private Const DUPLICATE_DEREG = &H7
Private Type NTWRKCNTRLBLCK
    ncb_command As Byte
    ncb_retcode As Byte
    ncb_lsn As Byte
    ncb_num As Byte
    ncb_buffer As Long
    ncb_length As Integer
    ncb_callname(0 To 15) As Byte
    ncb_name(0 To 15) As Byte
    ncb_rto As Byte
    ncb_sto As Byte
    lpFunc As Long
    ncb_lana_num As Byte
    ncb_cmd_cplt As Byte
    ncb_reserve(0 To 9) As Byte
    ncb_event As Long
End Type
Private Type LANA_ENUM
    length As Byte
    lana(0 To 256) As Byte
End Type
Private Type ADAPTER_STATUS
    adapter_address(0 To 5) As Byte

    rev_major As Byte
    reserved0 As Byte
    adapter_type As Byte
    rev_minor As Byte
    duration As Integer
    frmr_recv As Integer
    frmr_xmit As Integer
    iframe_recv_err As Integer
    xmit_aborts As Integer
    xmit_success As Long
    recv_success As Long
    iframe_xmit_err As Integer
    recv_buff_unavail As Integer
    t1_timeouts As Integer
    ti_timeouts As Integer
    reserved1 As Long
    free_ncbs As Integer
    max_cfg_ncbs As Integer
    max_ncbs As Integer
    xmit_buf_unavail As Integer
    max_dgram_size As Integer
    pending_sess As Integer
    max_cfg_sess As Integer
    max_sess As Integer
    max_sess_pkt_size As Integer
    name_count As Integer
End Type
Private Type NAME_BUFFER
    name_(0 To 15) As Byte
    name_num As Byte
    name_flags As Byte
End Type
Private Type NET_STATUS
    Adapter As ADAPTER_STATUS
    NameBuffer(30) As NAME_BUFFER
End Type
Private Const NCBENUM = &H37
Private Const NCBRESET = &H32
Private Const NCBASTAT = &H33
Private Declare Function NetBios Lib "netapi32.dll" Alias "Netbios" (ByRef pncb As NTWRKCNTRLBLCK) As Byte
Private Declare Function VarPtr Lib "MSVBVM60.DLL" (pVoid As Any) As Long

' Animation
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Private Const ICC_ANIMATE_CLASS = &H80
Private Const ANIMATE_CLASS = "SysAnimate32"

Private Const ACS_CENTER = &H1&
Private Const ACS_TRANSPARENT = &H2&
Private Const ACS_AUTOPLAY = &H4&
Private Const ACS_TIMER = &H8&

Private Const WM_PAINT = &HF
Private Const WM_USER = &H400&
Private Const ACM_OPEN = WM_USER + 100
Private Const ACM_PLAY = WM_USER + 101
Private Const ACM_STOP = WM_USER + 102

Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_BORDER = &H800000
Private Const WS_CLIPSIBLINGS = &H4000000

Private Const SW_HIDE = 0
Private Const SW_SHOW = 5

Private m_hWnd As Long
Private m_hWndParent As Long
Private m_AutoPlay As Boolean
Private m_Center As Boolean
Private m_Transparent As Boolean
Private m_Visible As Boolean
Private m_Playing As Boolean
Private m_AniResID As Long
Private m_AniFile As String
Private m_Left As Long
Private m_Top As Long
Private m_Width As Long
Private m_Height As Long

' Registry
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
Public Enum what
    SubNetmask
    gateway
    wins
    DNS
    AdapterName
    networkcomment
    MacAdrress
End Enum



Public Function NetInfo(what As what) As String
    Dim Maj As Integer
    Dim Min As Integer
    Dim Version As String
    Dim systeem As New system

    systeem.WinVer Maj, Min, Version
    Version = Version
    If Version = "Windows 98" Then NetInfo98 (what)
    NetInfo = Result
    If Version = "Windows 95" Then NetInfo95 (what)
    NetInfo = Result
    If Version = "Windows NT" Then NetInfoNT (what)
    NetInfo = Result
    If Version = "Windows 2000" Then NetInfo2000 (what)



End Function
Public Function NetInfo95(what As what) As String
    Result = ""
    On Error GoTo ErrorHandler


    Select Case what

        Case SubNetmask

            Set m_clsRegAccess = New Registry

            p_strSubKey = "System\CurrentControlSet\Services\Class\NetTrans\0006"
            p_strValueName = "IPMask"
            p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            Result = CStr(p_vntRtn)

        Case gateway

            Set m_clsRegAccess = New Registry

            p_strSubKey = "System\CurrentControlSet\Services\Class\NetTrans\0006"
            p_strValueName = "DefaultGateway"
            p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            Result = CStr(p_vntRtn)

        Case wins

            Set m_clsRegAccess = New Registry
            p_strValueName = "NameServer1"
            p_strSubKey = "System\CurrentControlSet\Services\VxD\MSTCP"
            p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            Result = CStr(p_vntRtn)

        Case DNS

            Set m_clsRegAccess = New Registry
            p_strValueName = "NameServer"
            p_strSubKey = "System\CurrentControlSet\Services\VxD\MSTCP"
            p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            p_strTmp = CStr(p_vntRtn)
            p_blnFirstTime = True

            Dim X As Integer
            Dim tel As Integer
            Dim A As String
            For X = 1 To Len(p_vntRtn)
                A = Mid(p_vntRtn, X, 1)
                If A = "," Or X = Len(p_vntRtn) Then
                    If Result = "" Then
                        Result = Mid(p_vntRtn, 1, tel)
                        tel = 0
                    Else
                        If X = Len(p_vntRtn) Then
                            Result = Result & "|" & Mid(p_vntRtn, X - tel + 1, tel)
                        Else
                            Result = Result & "|" & Mid(p_vntRtn, X - tel + 1, tel - 1)
                        End If
                        Debug.Print Result
                        tel = 0
                    End If
                End If
                tel = tel + 1

            Next X

        Case AdapterName

            Set m_clsRegAccess = New Registry
            p_strSubKey = "System\CurrentControlSet\Services\Class\Net\0000"
            p_strValueName = "DriverDesc"

            p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)

            Result = CStr(p_vntRtn)

        Case networkcomment

            Set m_clsRegAccess = New Registry
            p_strSubKey = "System\CurrentControlSet\Services\VxD\VNETSUP"
            p_strValueName = "Comment"
            p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            Result = CStr(p_vntRtn)

        Case MacAdrress

            Dim NCB As NTWRKCNTRLBLCK, Status As NET_STATUS, LanEnum As LANA_ENUM
            Dim bReturn As Byte, sMacAddress As String, I As Integer, sHex As String, l%
            Dim k%, iNumNames%, j%, m%, sName$, iPos%, nFlags%
            Dim sBuff As String

            NCB.ncb_command = NCBENUM
            NCB.ncb_buffer = VarPtr(LanEnum)
            NCB.ncb_length = LenB(LanEnum)
            bReturn = NetBios(NCB)
            sBuff = ""
            l = LanEnum.length

            If l > 0 Then
                NCB.ncb_command = NCBRESET
                NCB.ncb_lana_num = LanEnum.lana(k)
                bReturn = NetBios(NCB)
                NCB.ncb_command = NCBASTAT
                NCB.ncb_lana_num = LanEnum.lana(k)
                NCB.ncb_callname(0) = 42 'Max number of sessions            42
                NCB.ncb_buffer = VarPtr(Status)
                bReturn = NetBios(NCB)

                For I = 0 To 5
                    sHex = Hex(Status.Adapter.adapter_address(I))
                    If Len(sHex) = 1 Then sHex = "0" & sHex
                    sMacAddress = sMacAddress & sHex
                    If I <> 5 Then sMacAddress = sMacAddress + "-"
                Next I

                sBuff = sMacAddress
            End If
            Result = sBuff


    End Select

    GoTo noError

ErrorHandler:

    If Err.Number = 13 Then

    Else
        Call MsgBox(Err.Description & "!" & Chr(13) _
                , vbCritical + vbMsgBoxHelpButton + vbDefaultButton1 _
                , "Error #" & Err.Number, Err.HelpFile, 5000)

    End If
noError:


End Function


Public Function NetInfoNT(what As what) As String
    Result = ""
    On Error GoTo ErrorHandler

    Select Case what

        Case SubNetmask

            Set m_clsRegAccess = New Registry
            p_strSubKey = "SYSTEM\CurrentControlSet\Services\NetBT\Adapters"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            On Error Resume Next
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlSet\Services\" & p_strAdapterName & "\Parameters\Tcpip"
                p_strValueName = "SubnetMask"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                Result = CStr(p_vntRtn)
            Next p_lngLoop

        Case gateway

            Set m_clsRegAccess = New Registry
            p_strSubKey = "SYSTEM\CurrentControlSet\Services\NetBT\Adapters"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlSet\Services\" & p_strAdapterName & "\Parameters\Tcpip"
                p_strValueName = "DefaultGateway"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                Result = CStr(p_vntRtn)
            Next p_lngLoop

        Case wins

            Set m_clsRegAccess = New Registry
            p_strSubKey = "SYSTEM\CurrentControlSet\Services\NetBT\Adapters"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlSet\Services\NetBT\Adapters\" & p_strAdapterName
                p_strValueName = "NameServer"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                Result = CStr(p_vntRtn)
            Next p_lngLoop


        Case DNS

            Set m_clsRegAccess = New Registry
            p_strSubKey = "SYSTEM\CurrentControlSet\Services\NetBT\Adapters"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            p_strValueName = "NameServer"
            p_strSubKey = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
            p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            p_strTmp = CStr(p_vntRtn)

            Dim X As Integer
            Dim tel As Integer
            Dim A As String
            For X = 1 To Len(p_vntRtn)
                A = Mid(p_vntRtn, X, 1)
                If A = " " Or X = Len(p_vntRtn) Then
                    If Result = "" Then
                        Result = Mid(p_vntRtn, 1, tel)
                        tel = 0
                    Else
                        If X = Len(p_vntRtn) Then
                            Result = Result & "|" & Mid(p_vntRtn, X - tel + 1, tel)
                        Else
                            Result = Result & "|" & Mid(p_vntRtn, X - tel + 1, tel - 1)
                        End If
                        Debug.Print Result
                        tel = 0
                    End If
                End If
                tel = tel + 1

            Next X

        Case AdapterName

            Set m_clsRegAccess = New Registry
            p_strSubKey = "SYSTEM\CurrentControlSet\Services\El90x"
            p_strValueName = "DisplayName"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)

            p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)

            Result = CStr(p_vntRtn)

        Case networkcomment

            Set m_clsRegAccess = New Registry
            p_strSubKey = "SYSTEM\ControlSet001\Services\LanmanServer\Parameters\"
            p_strValueName = "srvcomment"
            p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            If Result = "" Then
                Result = CStr(p_vntRtn)
            Else
                If p_vntRtn = "" Then
                Else
                    Result = Result & " | " & CStr(p_vntRtn)
                End If
            End If
            Result = p_vntRtn

        Case MacAdrress
            Dim NCB As NTWRKCNTRLBLCK, Status As NET_STATUS, LanEnum As LANA_ENUM
            Dim bReturn As Byte, sMacAddress As String, I As Integer, sHex As String, l%
            Dim k%, iNumNames%, j%, m%, sName$, iPos%, nFlags%
            Dim sBuff As String

            NCB.ncb_command = NCBENUM
            NCB.ncb_buffer = VarPtr(LanEnum)
            NCB.ncb_length = LenB(LanEnum)
            bReturn = NetBios(NCB)
            sBuff = ""
            l = LanEnum.length

            If l > 0 Then
                NCB.ncb_command = NCBRESET
                NCB.ncb_lana_num = LanEnum.lana(k)
                bReturn = NetBios(NCB)
                NCB.ncb_command = NCBASTAT
                NCB.ncb_lana_num = LanEnum.lana(k)
                NCB.ncb_callname(0) = 42 'Max number of sessions            42
                NCB.ncb_buffer = VarPtr(Status)
                bReturn = NetBios(NCB)

                For I = 0 To 5
                    sHex = Hex(Status.Adapter.adapter_address(I))
                    If Len(sHex) = 1 Then sHex = "0" & sHex
                    sMacAddress = sMacAddress & sHex
                    If I <> 5 Then sMacAddress = sMacAddress + "-"
                Next I

                sBuff = sMacAddress
            End If
            Result = sBuff


    End Select

    GoTo noError

ErrorHandler:

    If Err.Number = 13 Then

    Else
        Call MsgBox(Err.Description & "!" & Chr(13) _
                , vbCritical + vbMsgBoxHelpButton + vbDefaultButton1 _
                , "Error #" & Err.Number, Err.HelpFile, 5000)

    End If
noError:


End Function

Public Function NetInfo98(what As what) As String
    Result = ""
    On Error GoTo ErrorHandler

    Select Case what

        Case SubNetmask

            Set m_clsRegAccess = New Registry
            p_strSubKey = "System\CurrentControlSet\Services\Class\NetTrans\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            On Error Resume Next
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "System\CurrentControlSet\Services\Class\NetTrans\" & p_strAdapterName & "\"
                p_strValueName = "IPMask"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If p_vntRtn = "0.0.0.0" Or p_vntRtn = "" Then
                Else
                    Result = CStr(p_vntRtn)
                End If
            Next p_lngLoop



        Case gateway

            Set m_clsRegAccess = New Registry
            p_strSubKey = "System\CurrentControlSet\Services\Class\NetTrans\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "System\CurrentControlSet\Services\Class\NetTrans\" & p_strAdapterName & "\"
                p_strValueName = "DefaultGateway"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)

                If Result = "" Then
                    Result = CStr(p_vntRtn)
                Else
                    If p_vntRtn = "" Then
                    Else
                        Result = Result & " | " & CStr(p_vntRtn)
                    End If
                End If
            Next p_lngLoop

        Case wins

            Set m_clsRegAccess = New Registry
            p_strSubKey = "System\CurrentControlSet\Services\Class\NetTrans\"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "System\CurrentControlSet\Services\Class\NetTrans\" & p_strAdapterName & "\"
                p_strValueName = "NameServer1"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                If Result = "" Then
                    Result = CStr(p_vntRtn)
                Else
                    If p_lngLoop = 1 Then
                        Result = CStr(p_vntRtn)
                    Else
                        If p_vntRtn = "" Then
                            Result = Result
                        Else
                            Result = Result & "|" & CStr(p_vntRtn)
                        End If
                    End If
                End If

            Next p_lngLoop


        Case DNS
            Set m_clsRegAccess = New Registry
            p_strValueName = "NameServer"
            p_strSubKey = "System\CurrentControlSet\Services\Vxd\MSTCP\"
            p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            p_strTmp = CStr(p_vntRtn)
            p_blnFirstTime = True
            p_lngPos = InStr(1, p_strTmp, ",", vbTextCompare)
            Do While p_lngPos > 0
                If p_blnFirstTime = True Then
                    Result = Result & Trim$(Mid$(p_strTmp, 1, p_lngPos - 1))
                    p_strTmp = Mid$(p_strTmp, p_lngPos + 1)
                    p_blnFirstTime = False
                Else
                    Result = Result & "|" & Trim$(Mid$(p_strTmp, 1, p_lngPos - 1))
                    p_strTmp = Mid$(p_strTmp, p_lngPos + 1)
                End If
                p_lngPos = InStr(1, p_strTmp, ",", vbTextCompare)
            Loop
            If p_lngLoop = 1 Then
                Result = CStr(p_strTmp)
            Else
                If Result = "" Then
                    Result = CStr(p_strTmp)
                Else
                    Result = Result & "|" & CStr(p_strTmp)
                End If
            End If

        Case AdapterName

            Set m_clsRegAccess = New Registry
            p_strSubKey = "SYSTEM\CurrentControlSet\Services\El90x"
            p_strValueName = "DisplayName"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)

            p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)

            Result = p_vntRtn

        Case networkcomment

            Set m_clsRegAccess = New Registry
            p_strSubKey = "SYSTEM\CurrentControlSet\Services\VxD\VNETSUP\"
            p_strValueName = "comment"
            p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            Result = p_vntRtn

        Case MacAdrress
            Dim NCB As NTWRKCNTRLBLCK, Status As NET_STATUS, LanEnum As LANA_ENUM
            Dim bReturn As Byte, sMacAddress As String, I As Integer, sHex As String, l%
            Dim k%, iNumNames%, j%, m%, sName$, iPos%, nFlags%
            Dim sBuff As String

            NCB.ncb_command = NCBENUM
            NCB.ncb_buffer = VarPtr(LanEnum)
            NCB.ncb_length = LenB(LanEnum)
            bReturn = NetBios(NCB)
            sBuff = ""
            l = LanEnum.length

            If l > 0 Then
                NCB.ncb_command = NCBRESET
                NCB.ncb_lana_num = LanEnum.lana(k)
                bReturn = NetBios(NCB)
                NCB.ncb_command = NCBASTAT
                NCB.ncb_lana_num = LanEnum.lana(k)
                NCB.ncb_callname(0) = 42 'Max number of sessions            42
                NCB.ncb_buffer = VarPtr(Status)
                bReturn = NetBios(NCB)

                For I = 0 To 5
                    sHex = Hex(Status.Adapter.adapter_address(I))
                    If Len(sHex) = 1 Then sHex = "0" & sHex
                    sMacAddress = sMacAddress & sHex
                    If I <> 5 Then sMacAddress = sMacAddress + "-"
                Next I

                sBuff = sMacAddress
            End If
            Result = sBuff


    End Select

    GoTo noError

ErrorHandler:

    If Err.Number = 13 Then

    Else
        Call MsgBox(Err.Description & "!" & Chr(13) _
                , vbCritical + vbMsgBoxHelpButton + vbDefaultButton1 _
                , "Error #" & Err.Number, Err.HelpFile, 5000)

    End If
noError:


End Function

Public Function NetInfo2000(what As what) As String
    Result = ""
    On Error GoTo ErrorHandler

    Select Case what

        Case SubNetmask

            Set m_clsRegAccess = New Registry
            p_strSubKey = "SYSTEM\CurrentControlSet\Services\NetBT\Adapters"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            On Error Resume Next
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlSet\Services\" & p_strAdapterName & "\Parameters\Tcpip"
                p_strValueName = "SubnetMask"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                Result = CStr(p_vntRtn)
            Next p_lngLoop

        Case gateway

            Set m_clsRegAccess = New Registry
            p_strSubKey = "SYSTEM\CurrentControlSet\Services\NetBT\Adapters"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlSet\Services\" & p_strAdapterName & "\Parameters\Tcpip"
                p_strValueName = "DefaultGateway"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                Result = CStr(p_vntRtn)
            Next p_lngLoop

        Case wins

            Set m_clsRegAccess = New Registry
            p_strSubKey = "SYSTEM\CurrentControlSet\Services\NetBT\Adapters"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            For p_lngLoop = 1 To p_lngNumAdapters
                p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
                p_strSubKey = "SYSTEM\CurrentControlSet\Services\NetBT\Adapters\" & p_strAdapterName
                p_strValueName = "NameServer"
                p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
                Result = CStr(p_vntRtn)
            Next p_lngLoop


        Case DNS

            Set m_clsRegAccess = New Registry
            p_strSubKey = "SYSTEM\CurrentControlSet\Services\NetBT\Adapters"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
            p_lngNumAdapters = UBound(p_vntAdapters)

            p_strValueName = "NameServer"
            p_strSubKey = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
            p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            p_strTmp = CStr(p_vntRtn)
            p_blnFirstTime = True
            p_lngPos = InStr(1, p_strTmp, " ", vbTextCompare)
            Do While p_lngPos > 0
                If p_blnFirstTime = True Then
                    Result = Result & Trim$(Mid$(p_strTmp, 1, p_lngPos - 1))
                    p_strTmp = Mid$(p_strTmp, p_lngPos + 1)
                    p_blnFirstTime = False
                Else
                    Result = Result & " | " & Trim$(Mid$(p_strTmp, 1, p_lngPos - 1))
                    p_strTmp = Mid$(p_strTmp, p_lngPos + 1)
                End If
                p_lngPos = InStr(1, p_strTmp, " ", vbTextCompare)
            Loop
            If Len(p_strTmp) > 0 Then
                Result = Result & " | " & Trim$(p_strTmp)
            End If

        Case AdapterName

            Set m_clsRegAccess = New Registry
            p_strSubKey = "SYSTEM\CurrentControlSet\Services\El90x"
            p_strValueName = "DisplayName"
            p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)

            p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)

            Result = p_vntRtn

        Case networkcomment

            Set m_clsRegAccess = New Registry
            p_strSubKey = "SYSTEM\ControlSet001\Services\LanmanServer\Parameters\"
            p_strValueName = "srvcomment"
            p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
            Result = p_vntRtn

        Case MacAdrress
            Dim NCB As NTWRKCNTRLBLCK, Status As NET_STATUS, LanEnum As LANA_ENUM
            Dim bReturn As Byte, sMacAddress As String, I As Integer, sHex As String, l%
            Dim k%, iNumNames%, j%, m%, sName$, iPos%, nFlags%
            Dim sBuff As String

            NCB.ncb_command = NCBENUM
            NCB.ncb_buffer = VarPtr(LanEnum)
            NCB.ncb_length = LenB(LanEnum)
            bReturn = NetBios(NCB)
            sBuff = ""
            l = LanEnum.length

            If l > 0 Then
                NCB.ncb_command = NCBRESET
                NCB.ncb_lana_num = LanEnum.lana(k)
                bReturn = NetBios(NCB)
                NCB.ncb_command = NCBASTAT
                NCB.ncb_lana_num = LanEnum.lana(k)
                NCB.ncb_callname(0) = 42 'Max number of sessions            42
                NCB.ncb_buffer = VarPtr(Status)
                bReturn = NetBios(NCB)

                For I = 0 To 5
                    sHex = Hex(Status.Adapter.adapter_address(I))
                    If Len(sHex) = 1 Then sHex = "0" & sHex
                    sMacAddress = sMacAddress & sHex
                    If I <> 5 Then sMacAddress = sMacAddress + "-"
                Next I

                sBuff = sMacAddress
            End If
            Result = sBuff


    End Select

    GoTo noError

ErrorHandler:

    If Err.Number = 13 Then

    Else
        Call MsgBox(Err.Description & "!" & Chr(13) _
                , vbCritical + vbMsgBoxHelpButton + vbDefaultButton1 _
                , "Error #" & Err.Number, Err.HelpFile, 5000)

    End If
noError:


End Function

