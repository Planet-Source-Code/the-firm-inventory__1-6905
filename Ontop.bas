Attribute VB_Name = "Ontop"
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Public Sub Form_Ontop(hwnd As Long, value As Boolean)

    If value = True Then
        res% = SetWindowPos(hwnd, HWND_TOPMOST, _
                0, 0, 0, 0, FLAGS)
    Else
        res% = SetWindowPos(hwnd, HWND_NOTOPMOST, _
                0, 0, 0, 0, FLAGS)
    End If

End Sub

Sub WindowHandle(win, cas As Long)

    Select Case cas
        Case 0:
        Dim X%
        X% = SendMessage(win, WM_CLOSE, 0, 0)
        Case 1:
        X = ShowWindow(win, SW_SHOW)
        Case 2:
        X = ShowWindow(win, SW_HIDE)
        Case 3:
        X = ShowWindow(win, SW_MAXIMIZE)
        Case 4:
        X = ShowWindow(win, SW_MINIMIZE)
    End Select




End Sub

