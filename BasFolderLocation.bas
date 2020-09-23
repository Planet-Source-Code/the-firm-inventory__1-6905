Attribute VB_Name = "InfoFolderLocation"
Option Explicit

Private Type ITEMIDLIST
    mkid As Long
End Type

Private Declare Function SHGetSpecialFolderLocation _
        Lib "shell32.dll" _
        (ByVal hwndOwner As Long, ByVal nFolder As SHFolders, _
        ppidl As ITEMIDLIST) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" _
        (ByVal pv As Long)
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
        Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
        ByVal pszPath As String) As Long

Private Declare Function apiWindDir Lib "kernel32" _
        Alias "GetWindowsDirectoryA" _
        (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function apiSysDir Lib "kernel32" _
        Alias "GetSystemDirectoryA" _
        (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function apiTempDir Lib "kernel32" _
        Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


Public Enum SHFolders
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_ALTSTARTUP = &H1D '// DBCS
    CSIDL_COMMON_ALTSTARTUP = &H1E '// DBCS
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
End Enum


Public Function FolderLocation(lFolder As SHFolders, hWnd As Long) As String

    Dim lp As ITEMIDLIST
    Dim tmpStr As String
    'Get the PIDL for this folder
    SHGetSpecialFolderLocation hWnd, lFolder, lp
    'Convert it to a string path
    tmpStr = Space$(255)
    SHGetPathFromIDList lp.mkid, tmpStr
    If InStr(tmpStr, Chr$(0)) > 0 Then
        'Strip nulls from the string
        tmpStr = Left$(tmpStr, InStr(tmpStr, Chr$(0)) - 1)
    End If
    'Free the PIDL
    CoTaskMemFree lp.mkid
    'Return
    If tmpStr = "" Then tmpStr = "Could not be determined"
    FolderLocation = tmpStr

End Function

Public Function SystemDir() As String
    '---------------------------------------------------------------------------
    ' FUNCTION: SystemDir
    '
    ' Gets the WINDOWS\SYSTEM directory.
    '
    ' Returns a string containing the full path, ends with a "\". If the
    ' call fails a empty string is returned.
    '---------------------------------------------------------------------------
    '
    Dim Bufstr As String
    Bufstr = Space$(50)


    '---------------------------------------------------------------------------
    ' Call the API and remove the spaces using RTrim. Remove the terminating
    ' character and add a backslash when it isn't already there.
    '---------------------------------------------------------------------------
    If apiSysDir(Bufstr, 50) > 0 Then
        SystemDir = Bufstr
        SystemDir = RTrim(SystemDir)
        SystemDir = StripTerminator(SystemDir)

        If Right$(SystemDir, 1) <> "\" Then
            SystemDir = SystemDir + "\"
        End If

    Else
        SystemDir = ""
    End If

End Function

Public Function TempDir() As String
    '---------------------------------------------------------------------------
    ' FUNCTION: TempDir
    '
    ' Get the Temporary directory windows uses.
    '
    ' OUT: TempDir  - String containing the directory.
    '
    ' If the function fails a empty string is returned.
    '---------------------------------------------------------------------------
    '
    Dim Bufstr As String
    Bufstr = Space$(50)


    '---------------------------------------------------------------------------
    ' Call the API and remove the spaces using RTrim. Next, remove the terminating
    ' character using StripTerminator, and add a backslash, when if it wasn't
    ' already there.
    '---------------------------------------------------------------------------
    If apiTempDir(50, Bufstr) > 0 Then
        TempDir = Bufstr
        TempDir = RTrim(TempDir)
        TempDir = StripTerminator(TempDir)

        If Right$(TempDir, 1) <> "\" Then
            TempDir = TempDir + "\"
        End If

    Else
        TempDir = ""
    End If

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
