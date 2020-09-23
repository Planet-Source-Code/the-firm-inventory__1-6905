Attribute VB_Name = "Save"
Dim cOnn As ADODB.Connection
Dim cOmd As ADODB.Command
Dim rS As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim MysQl As String
' end
'*******************************
' String ------>>
Dim A As String
Dim B As String
Dim dSN As Boolean
Dim stRConn As String
'*******************************
' Integer ------>>
Dim I As Integer
Public Function save()

    DoEvents

    ' Check for existing dsn_connection
    A = inIgEt("inventaris_dsn_connection", "inventaris_dsn_connection_key", _
            App.Path & "\ini\inventaris.ini")

    dSN = checkWantedAccessDSN(A) ' Dns_connection_module

    If dSN = True Then

        ' If exist then goto Connection_information

    Else

        ' If Not, Make it....!!
        A = inIgEt("inventaris_mdb_location", "inventaris_mdb_location_key", _
                App.Path & "\ini\inventaris.ini") ' Look for value in inifile


        B = inIgEt("inventaris_dsn_connection", "inventaris_dsn_connection_key", _
                App.Path & "\ini\inventaris.ini") ' Look for value in inifile


        Call createAccessDSN(A, B) 'Create DSN_connection
        ' Goto Connection_information
    End If


    'Connection_information

    B = inIgEt("inventaris_dsn_connection", "inventaris_dsn_connection_key", _
            App.Path & "\ini\inventaris.ini") ' Look for value in inifile

    stRConn = "Data Source=" & B

    'Set Connection

    Set cOnn = New ADODB.Connection
    Set cOmd = New ADODB.Command
    Set rS = New ADODB.Recordset

    'Open connection

    cOnn.ConnectionString = stRConn
    cOnn.Open stRConn
    cOmd.ActiveConnection = cOnn
    MysQl = "SELECT * FROM [Afdeling] ORDER BY [Veld1]"
    rS.Open MysQl, cOnn, adOpenDynamic, adLockOptimistic
    rS.MoveFirst
    I = 1

End Function

Public Function FileExists(ByVal sFileName As String) As Boolean
    Dim sFile As String
    On Error Resume Next
    FileExists = False
    sFile = Dir$(sFileName)
    If (Len(sFile) > 0) And (Err = 0) Then
        FileExists = True
    End If
End Function
