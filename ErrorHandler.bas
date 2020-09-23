Attribute VB_Name = "ErrorHandler"
Option Explicit

Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
        (ByVal dwflags As Long, lpSource As Any, _
        ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
        ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Public Const FORMAT_MESSAGE_FROM_STRING = &H400
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Public Const LANG_USER_DEFAULT = &H400&


Public Function GetLastErrorStr(dwErrCode As Long) As String

    Static sMsgBuf As String * 257, dwLen As Long


    dwLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM _
            Or FORMAT_MESSAGE_IGNORE_INSERTS _
            Or FORMAT_MESSAGE_MAX_WIDTH_MASK, ByVal 0&, _
            dwErrCode, LANG_USER_DEFAULT, _
            ByVal sMsgBuf, 256&, 0&)

    If dwLen Then GetLastErrorStr = Left$(sMsgBuf, dwLen)


End Function






