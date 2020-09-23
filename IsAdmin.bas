Attribute VB_Name = "IsAdmin"
' Developed for you by Elvio Serrao
' Elvio.Serrao@nrma.com.au

'=====================================================================================
' This function will determine whether or not a thread is running in the user context
' of the local Administrator account, you need to examine the access token associated
' with that thread using the GetTokenInformation() API, since this access token
' represents the user under which the thread is running. By default the token
' associated with a thread is that of its containing process, but this user context
' will be superceded by any token attached directly to the thread. So to determine a
' thread’s user context, first attempt to obtain any token attached directly to the
' thread with OpenThreadToken(). If this fails, and it reports an ERROR_NO_TOKEN,
' then obtain the token of the thread’s containing process with OpenProcessToken().
'=====================================================================================

Option Explicit
Option Base 0     ' Important assumption for this code

Private Const ANYSIZE_ARRAY = 20 'Fixed at this size for comfort. Could be bigger or made dynamic.

' Security APIs
Private Const TokenUser = 1
Private Const TokenGroups = 2
Private Const TokenPrivileges = 3
Private Const TokenOwner = 4
Private Const TokenPrimaryGroup = 5
Private Const TokenDefaultDacl = 6
Private Const TokenSource = 7
Private Const TokenType = 8
Private Const TokenImpersonationLevel = 9
Private Const TokenStatistics = 10

' Token Specific Access Rights
Private Const TOKEN_ASSIGN_PRIMARY = &H1
Private Const TOKEN_DUPLICATE = &H2
Private Const TOKEN_IMPERSONATE = &H4
Private Const TOKEN_QUERY = &H8
Private Const TOKEN_QUERY_SOURCE = &H10
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_ADJUST_GROUPS = &H40
Private Const TOKEN_ADJUST_DEFAULT = &H80

' NT well-known SIDs
Private Const SECURITY_DIALUP_RID = &H1
Private Const SECURITY_NETWORK_RID = &H2
Private Const SECURITY_BATCH_RID = &H3
Private Const SECURITY_INTERACTIVE_RID = &H4
Private Const SECURITY_SERVICE_RID = &H6
Private Const SECURITY_ANONYMOUS_LOGON_RID = &H7
Private Const SECURITY_LOGON_IDS_RID = &H5
Private Const SECURITY_LOCAL_SYSTEM_RID = &H12
Private Const SECURITY_NT_NON_UNIQUE = &H15
Private Const SECURITY_BUILTIN_DOMAIN_RID = &H20

' Well-known domain relative sub-authority values (RIDs)
Private Const DOMAIN_ALIAS_RID_ADMINS = &H220
Private Const DOMAIN_ALIAS_RID_USERS = &H221
Private Const DOMAIN_ALIAS_RID_GUESTS = &H222
Private Const DOMAIN_ALIAS_RID_POWER_USERS = &H223
Private Const DOMAIN_ALIAS_RID_ACCOUNT_OPS = &H224
Private Const DOMAIN_ALIAS_RID_SYSTEM_OPS = &H225
Private Const DOMAIN_ALIAS_RID_PRINT_OPS = &H226
Private Const DOMAIN_ALIAS_RID_BACKUP_OPS = &H227
Private Const DOMAIN_ALIAS_RID_REPLICATOR = &H228

Private Const SECURITY_NT_AUTHORITY = &H5

Type SID_AND_ATTRIBUTES
    Sid As Long
    Attributes As Long
End Type

Type TOKEN_GROUPS
    GroupCount As Long
    Groups(ANYSIZE_ARRAY) As SID_AND_ATTRIBUTES
End Type

Type SID_IDENTIFIER_AUTHORITY
    value(0 To 5) As Byte
End Type

Declare Function GetCurrentProcess Lib "kernel32" () As Long

Declare Function GetCurrentThread Lib "kernel32" () As Long

Declare Function OpenProcessToken Lib "Advapi32" ( _
        ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, _
        TokenHandle As Long) As Long

Declare Function OpenThreadToken Lib "Advapi32" ( _
        ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, _
        ByVal OpenAsSelf As Long, TokenHandle As Long) As Long

Declare Function GetTokenInformation Lib "Advapi32" ( _
        ByVal TokenHandle As Long, TokenInformationClass As Integer, _
        TokenInformation As Any, ByVal TokenInformationLength As Long, _
        ReturnLength As Long) As Long

Declare Function AllocateAndInitializeSid Lib "Advapi32" ( _
        pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, _
        ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, _
        ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, _
        ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, _
        ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, _
        ByVal nSubAuthority7 As Long, lpPSid As Long) As Long

Declare Function RtlMoveMemory Lib "kernel32" ( _
        Dest As Any, Source As Any, ByVal lSize As Long) As Long

Declare Function IsValidSid Lib "Advapi32" (ByVal pSid As Long) As Long

Declare Function EqualSid Lib "Advapi32" (pSid1 As Any, pSid2 As Any) As Long

Declare Sub FreeSid Lib "Advapi32" (pSid As Any)

Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Function IsAdmins() As Boolean
    Dim Maj As Integer
    Dim Min As Integer
    Dim Version As String
    Dim systeem As New system

    systeem.WinVer Maj, Min, Version
    Version = Version
    If Version = "Windows 98" Then
        IsAdmins = True
        Exit Function
    End If
    If Version = "Windows 95" Then
        IsAdmins = True
        Exit Function
    End If



    Dim hProcessToken       As Long
    Dim BufferSize          As Long
    Dim psidAdmin           As Long
    Dim lResult             As Long
    Dim X                   As Integer
    Dim tpTokens            As TOKEN_GROUPS
    Dim tpSidAuth           As SID_IDENTIFIER_AUTHORITY

    IsAdmins = False
    tpSidAuth.value(5) = SECURITY_NT_AUTHORITY

    ' Obtain current process token
    If Not OpenThreadToken(GetCurrentThread(), TOKEN_QUERY, True, hProcessToken) Then
        Call OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hProcessToken)
    End If
    If hProcessToken Then

        ' Deternine the buffer size required
        Call GetTokenInformation(hProcessToken, ByVal TokenGroups, 0, 0, BufferSize) ' Determine required buffer size
        If BufferSize Then
            ReDim InfoBuffer((BufferSize \ 4) - 1) As Long

            ' Retrieve your token information
            lResult = GetTokenInformation(hProcessToken, ByVal TokenGroups, InfoBuffer(0), BufferSize, BufferSize)
            If lResult <> 1 Then Exit Function

            ' Move it from memory into the token structure
            Call RtlMoveMemory(tpTokens, InfoBuffer(0), Len(tpTokens))

            ' Retreive the admins sid pointer
            lResult = AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, _
                    DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, psidAdmin)
            If lResult <> 1 Then Exit Function
            If IsValidSid(psidAdmin) Then
                For X = 0 To tpTokens.GroupCount

                    ' Run through your token sid pointers
                    If IsValidSid(tpTokens.Groups(X).Sid) Then

                        ' Test for a match between the admin sid equalling your sid's
                        If EqualSid(ByVal tpTokens.Groups(X).Sid, ByVal psidAdmin) Then
                            IsAdmins = True
                            Exit For
                        End If
                    End If
                Next
            End If
            If psidAdmin Then Call FreeSid(psidAdmin)
        End If
        Call CloseHandle(hProcessToken)
    End If
End Function


