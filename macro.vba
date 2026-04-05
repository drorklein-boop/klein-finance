Sub RunMonthlyUpdate()
    Dim wsh As Object
    Dim exitCode As Long
    If Dir("C:\KleinFinance\launcher.py") = "" Then
        MsgBox "launcher.py not found", vbCritical
        Exit Sub
    End If
    Set wsh = CreateObject("WScript.Shell")
    ' Use 1 (visible window) so user can see output and press Enter
    exitCode = wsh.Run("cmd /k python ""C:\KleinFinance\launcher.py""", 1, False)
    Set wsh = Nothing
End Sub