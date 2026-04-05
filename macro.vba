' Klein Finance - Monthly Update Button
' Install ONCE. Never touch again.
' Assign to button: RunMonthlyUpdate

Sub RunMonthlyUpdate()
    Dim wsh      As Object
    Dim exitCode As Long

    If Dir("C:\KleinFinance\launcher.py") = "" Then
        MsgBox "launcher.py not found in C:\KleinFinance", vbCritical, "Klein Finance"
        Exit Sub
    End If

    Set wsh = CreateObject("WScript.Shell")
    exitCode = wsh.Run("cmd /c python ""C:\KleinFinance\launcher.py""", 1, True)

    Application.CalculateFull

    If exitCode = 0 Then
        MsgBox "Update complete!", vbInformation, "Klein Finance"
    Else
        MsgBox "Finished with warnings.", vbExclamation, "Klein Finance"
    End If

    Set wsh = Nothing
End Sub