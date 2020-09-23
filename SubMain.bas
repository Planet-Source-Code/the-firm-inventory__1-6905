Attribute VB_Name = "SubMain"
Dim Os As String
Sub Main()
    Dim OsVersion As New system
    On Error GoTo ErrorHandler



    Invent.Show

    GoTo noError

ErrorHandler:

    Call MsgBox(Err.Description & "!" & Chr(13) _
            , vbCritical + vbDefaultButton1 _
            , "Error #" & Err.Number, Err.HelpFile, 5000)

noError:

End Sub

