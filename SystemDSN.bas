Attribute VB_Name = "SystemDSN"
Option Explicit
Private Declare Function SQLAllocEnv Lib "odbc32.dll" _
        (phenv As Long) As Integer

Private Declare Function SQLDataSources Lib "odbc32.dll" _
        (ByVal hEnv As Long, ByVal fDirection As Integer, _
        ByVal szDSN$, ByVal cbDSNMax%, pcbDSN As Integer, _
        ByVal szDescription As String, _
        ByVal cbDescriptionMax As Integer, _
        pcbDescription As Integer) As Integer

Private Declare Function SQLFreeEnv Lib "odbc32.dll" _
        (ByVal hEnv As Long) As Integer

Private Const SQL_SUCCESS As Long = 0
Private Const SQL_FETCH_NEXT = 1
Private Const SQL_FETCH_FIRST_SYSTEM = 32
Private Const KEY_QUERY_VALUE = &H1

Private Const ERROR_SUCCESS = 0&

Private Const REG_SZ = 1

Private Const HKEY_LOCAL_MACHINE = &H80000002

Private Const REG_DWORD = 4

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long ' Note that If you declare the lpData parameter as String, you must pass it By Value.

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Private Sub SystemDSNs(DSNArray() As String)
    Dim iRet As Integer
    Dim sDSN As String
    Dim sDriver As String
    Dim iDSNLen As Integer
    Dim iDriverLen As Integer
    ReDim DSNArray(0) As String
    Dim lEnvHandle As Long

    iRet = SQLAllocEnv(lEnvHandle)
    sDSN = Space(1024)
    sDriver = Space(1024)
    iRet = SQLDataSources(lEnvHandle, SQL_FETCH_FIRST_SYSTEM, _
            sDSN, 1024, iDSNLen, sDriver, 1024, iDriverLen)

    If iRet = SQL_SUCCESS Then

        sDSN = Mid(sDSN, 1, iDSNLen)
        sDriver = Mid(sDriver, 1, iDriverLen)
        DSNArray(0) = sDSN & "|" & sDriver

        Do Until iRet <> SQL_SUCCESS
            sDSN = Space(1024)
            sDriver = Space(1024)
            iRet = SQLDataSources(lEnvHandle, SQL_FETCH_NEXT, _
                    sDSN, 1024, iDSNLen, sDriver, 1024, iDriverLen)

            If Trim(sDSN) <> "" Then
                sDSN = Mid(sDSN, 1, iDSNLen)
                sDriver = Mid(sDriver, 1, iDriverLen)
                ReDim Preserve DSNArray(UBound(DSNArray) + 1)
                DSNArray(UBound(DSNArray)) = sDSN & "|" & sDriver
            End If
        Loop

    End If

    iRet = SQLFreeEnv(lEnvHandle)

End Sub

Public Function dSN() As String
    Dim asDSNArray() As String
    Dim iCtr As Integer

    SystemDSNs asDSNArray
    For iCtr = 0 To UBound(asDSNArray)
        dSN = asDSNArray(iCtr)
    Next

End Function


Public Function isSZKeyExist(szKeyPath As String, szKeyName As String, _
            ByRef szKeyValue As String) As Boolean

    Dim bRes As Boolean
    Dim lRes As Long
    Dim hKey As Long

    lRes = RegOpenKeyEx(HKEY_LOCAL_MACHINE, szKeyPath, 0&, KEY_QUERY_VALUE, hKey)

    If lRes <> ERROR_SUCCESS Then

        isSZKeyExist = False
        Exit Function

    End If

    lRes = RegQueryValueEx(hKey, szKeyName, 0&, REG_SZ, ByVal szKeyValue, Len(szKeyValue))
    RegCloseKey (hKey)

    If lRes <> ERROR_SUCCESS Then

        isSZKeyExist = False
        Exit Function

    End If

    isSZKeyExist = True

End Function

Public Function checkAccessDriver(ByRef Path As String) As Boolean

    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim bRes As Boolean
    bRes = False

    szKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\Microsoft Access Driver (*.mdb)"

    szKeyName = "Driver"

    szKeyValue = String(255, Chr(32))

    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then

        Path = szKeyValue
        bRes = True

    Else

        bRes = False

    End If

    checkAccessDriver = bRes

End Function

Public Function checkWantedAccessDSN(szWantedDSN As String) As Boolean

    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim bRes As Boolean

    szKeyPath = "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"

    szKeyName = szWantedDSN

    szKeyValue = String(255, Chr(32))

    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then

        bRes = True

    Else
        bRes = False

    End If

    checkWantedAccessDSN = bRes

End Function

Public Function createAccessDSN(Path As String, _
            szWantedDSN As String) As Boolean

    Dim hKey As Long
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim lKeyValue As Long
    Dim lRes As Long
    Dim lSize As Long
    Dim szEmpty As String

    szEmpty = Chr(0)

    lSize = 4

    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & szWantedDSN, hKey)

    If lRes <> ERROR_SUCCESS Then

        createAccessDSN = False

        Exit Function

    End If

    lRes = RegSetValueExString(hKey, "UID", 0&, REG_SZ, _
            szEmpty, Len(szEmpty))

    szKeyValue = Path

    lRes = RegSetValueExString(hKey, "DBQ", 0&, REG_SZ, _
            szKeyValue, Len(szKeyValue))

    szKeyValue = Path

    lRes = RegSetValueExString(hKey, "Driver", 0&, REG_SZ, _
            szKeyValue, Len(szKeyValue))

    szKeyValue = "MS Access;"

    lRes = RegSetValueExString(hKey, "FIL", 0&, REG_SZ, _
            szKeyValue, Len(szKeyValue))

    lKeyValue = 25

    lRes = RegSetValueExLong(hKey, "DriverId", 0&, REG_DWORD, _
            lKeyValue, 4)

    lKeyValue = 0

    lRes = RegSetValueExLong(hKey, "SafeTransactions", 0&, REG_DWORD, _
            lKeyValue, 4)

    lRes = RegCloseKey(hKey)

    szKeyPath = "SOFTWARE\ODBC\ODBC.INI\" & szWantedDSN & "\Engines\Jet"
    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, szKeyPath, hKey)

    If lRes <> ERROR_SUCCESS Then

        createAccessDSN = False
        Exit Function

    End If

    lRes = RegSetValueExString(hKey, "ImplicitCommitSync", 0&, REG_SZ, _
            szEmpty, Len(szEmpty))

    szKeyValue = "Yes"

    lRes = RegSetValueExString(hKey, "UserCommitSync", 0&, REG_SZ, _
            szKeyValue, Len(szKeyValue))

    lKeyValue = 2048

    lRes = RegSetValueExLong(hKey, "MaxBufferSize", 0&, REG_DWORD, lKeyValue, 4)

    lKeyValue = 5

    lRes = RegSetValueExLong(hKey, "PageTimeout", 0&, REG_DWORD, lKeyValue, 4)

    lKeyValue = 3

    lRes = RegSetValueExLong(hKey, "Threads", 0&, REG_DWORD, lKeyValue, 4)

    lRes = RegCloseKey(hKey)

    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, _
            "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKey)

    If lRes <> ERROR_SUCCESS Then
        createAccessDSN = False
        Exit Function
    End If

    szKeyValue = "Microsoft Access Driver (*.mdb)"

    lRes = RegSetValueExString(hKey, szWantedDSN, 0&, REG_SZ, _
            szKeyValue, Len(szKeyValue))

    lRes = RegCloseKey(hKey)
    createAccessDSN = True

End Function

