Attribute VB_Name = "NetRecources"
' * Programmer Name  : Waty Thierry
' * Web Site : http://smalig.tripod.com/vb/119902.html
' * E-Mail           : waty.thierry@usa.net
' * Date             : 08/11/1999
' * Time             : 12:40
' ******************************************
' * Comments         : Display networked computers in a list box
' *
' *
' *****************************************

Option Explicit

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As Long
    lpRemoteName As Long
    lpComment As Long
    lpProvider As Long
End Type

Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias _
        "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, _
        ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long

Private Declare Function WNetEnumResource Lib "mpr.dll" Alias _
        "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, _
        ByVal lpBuffer As Long, lpBufferSize As Long) As Long

Private Declare Function WNetCloseEnum Lib "mpr.dll" _
        (ByVal hEnum As Long) As Long

Private Const RESOURCE_CONNECTED = &H1
Private Const RESOURCE_GLOBALNET = &H2
Private Const RESOURCE_REMEMBERED = &H3

Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCETYPE_DISK = &H1
Private Const RESOURCETYPE_PRINT = &H2
Private Const RESOURCETYPE_UNKNOWN = &HFFFF

Private Const RESOURCEUSAGE_CONNECTABLE = &H1
Private Const RESOURCEUSAGE_CONTAINER = &H2
Private Const RESOURCEUSAGE_RESERVED = &H80000000

Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Declare Function GlobalAlloc Lib "kernel32" _
        (ByVal wFlags As Long, ByVal dwBytes As Long) As Long

Private Declare Function GlobalFree Lib "kernel32" _
        (ByVal hMem As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias _
        "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, _
        ByVal cbCopy As Long)

Private Declare Function CopyPointer2String Lib _
        "kernel32" Alias "lstrcpyA" (ByVal NewString As _
        String, ByVal OldString As Long) As Long

Public Function DoNetEnum(list As Object) As Boolean

    'PURPOSE: DISPLAYS NETWORK NAME AND
    'ALL COMPUTERS ON THE NETWORK IN A LIST
    'BOX

    'PARAMETER: ListBox (or any control with similar
    'interface, such as ComboBox) in which to display
    'list of computers

    'RETURNS: True if successful, false otherwise

    Dim hEnum As Long, lpBuff As Long, NR As NETRESOURCE
    Dim cbBuff As Long, cCount As Long
    Dim p As Long, res As Long, I As Long

    On Error Resume Next
    'test to see if list is a
    'list box type control
    list.AddItem " "
    list.Clear
    If Err.Number > 0 Then Exit Function

    On Error GoTo ErrorHandler

    'Setup the NETRESOURCE input structure.
    NR.lpRemoteName = 0
    cbBuff = 10000
    cCount = &HFFFFFFFF

    'Open a Net enumeration operation handle: hEnum.
    res = WNetOpenEnum(RESOURCE_GLOBALNET, _
            RESOURCETYPE_ANY, 0, NR, hEnum)

    If res = 0 Then

        'Create a buffer large enough for the results.
        '10000 bytes should be sufficient.
        lpBuff = GlobalAlloc(GPTR, cbBuff)
        'Call the enumeration function.
        res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)
        If res = 0 Then
            p = lpBuff
            'WNetEnumResource fills the buffer with an array of
            'NETRESOURCE structures. Walk through the list and print
            'each local and remote name.
            For I = 1 To cCount
                ' All we get back are the Global Network Containers --- Enumerate each of these
                CopyMemory NR, ByVal p, LenB(NR)
                list.AddItem "Network Name " & _
                        PointerToString(NR.lpRemoteName)

                DoNetEnum2 NR, list
                p = p + LenB(NR)
            Next I
        End If
        DoNetEnum = True

ErrorHandler:
        On Error Resume Next
        If lpBuff <> 0 Then GlobalFree (lpBuff)
        WNetCloseEnum (hEnum) 'Close the enumeration

    End If

End Function

Private Function PointerToString(p As Long) As String

    'The values returned in the NETRESOURCE structures are pointers to
    'ANSI strings so they need to be converted to Visual Basic Strings.

    Dim s As String
    s = String(255, Chr$(0))
    CopyPointer2String s, p
    PointerToString = Left(s, InStr(s, Chr$(0)) - 1)

End Function

Public Sub DoNetEnum2(NR As NETRESOURCE, list As Object)

    Dim hEnum As Long, lpBuff As Long
    Dim cbBuff As Long, cCount As Long
    Dim p As Long, res As Long, I As Long

    'Setup the NETRESOURCE input structure.
    cbBuff = 10000
    cCount = &HFFFFFFFF

    'Open a Net enumeration operation handle: hEnum.
    res = WNetOpenEnum(RESOURCE_GLOBALNET, _
            RESOURCETYPE_ANY, 0, NR, hEnum)

    If res = 0 Then

        'Create a buffer large enough for the results.
        '10000 bytes should be sufficient.
        lpBuff = GlobalAlloc(GPTR, cbBuff)
        'Call the enumeration function.
        res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)

        If res = 0 Then
            p = lpBuff
            'WNetEnumResource fills the buffer with an array of
            'NETRESOURCE structures. Walk through the list and print
            'each remote name.
            For I = 1 To cCount
                CopyMemory NR, ByVal p, LenB(NR)
                list.AddItem " " & PointerToString(NR.lpRemoteName)
                p = p + LenB(NR)
            Next I
        End If

        If lpBuff <> 0 Then GlobalFree (lpBuff)
        WNetCloseEnum (hEnum) 'Close the enumeration

    End If

End Sub




