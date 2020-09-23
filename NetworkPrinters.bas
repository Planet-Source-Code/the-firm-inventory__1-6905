Attribute VB_Name = "NetworkPrinters"
Public Const HWND_BROADCAST As Long = &HFFFF
Public Const WM_WININICHANGE As Long = &H1A

Declare Function GetProfileString Lib "kernel32" _
        Alias "GetProfileStringA" _
        (ByVal lpAppName As String, _
        ByVal lpKeyName As String, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long) As Long

Declare Function WriteProfileString Lib "kernel32" _
        Alias "WriteProfileStringA" _
        (ByVal lpszSection As String, _
        ByVal lpszKeyName As String, _
        ByVal lpszString As String) As Long

Declare Function SendMessage Lib "user32" _
        Alias "SendMessageA" _
        (ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lparam As Any) As Long


Public Function ProfileLoadWinIniList() As String
    lpSectionName = "PrinterPorts"

    Dim success As Long
    Dim nSize As Long
    Dim lpKeyName As String
    Dim ret As String

    ret = Space$(8102)
    nSize = Len(ret)
    success = GetProfileString(lpSectionName, vbNullString, "", ret, nSize)

    If success Then

        ret = Left$(ret, success)

        Do Until ret = ""

            lpKeyName = StripNulls(ret)
            If Mid(lpKeyName, 1, 2) = "\\" Then
                If ProfileLoadWinIniList = "" Then
                    ProfileLoadWinIniList = lpKeyName
                Else
                    ProfileLoadWinIniList = ProfileLoadWinIniList & "|" & lpKeyName
                End If
            End If
            Debug.Print ProfileLoadWinIniList
        Loop

    End If


End Function
Private Function StripNulls(startstr As String) As String

    Dim pos As Long

    pos = InStr(startstr$, Chr$(0))

    If pos Then

        StripNulls = Mid$(startstr, 1, pos - 1)
        startstr = Mid$(startstr, pos + 1, Len(startstr))

    End If

End Function
