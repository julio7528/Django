Function StartConfigurations(iParameter)

    vParameter = Split(iParameter, "|")
    vDir = vParameter(0)
    vDelRequirements = vParameter(1)
    vCreateRequirements = vParameter(2)

    On Error Resume Next

    vCommand = Array(cDir, vDelRequirements, vCreateRequirements)

    For Each vCommand In vParameter

    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("cmd /c " & vCommand)

        WScript.Echo "Start: " & vCommand
        Do While Not objExec.StdOut.AtEndOfStream
            strLine = objExec.StdOut.ReadLine
            WScript.Echo strLine
        Loop
        WScript.Echo "End: " & vCommand
        WScript.Echo "----------------------------------------"

    Next
    'Check Error
    If Err.Number <> 0 Then
        WScript.Echo "Error: " & Err.Number & " " & Err.Description
        Err.Clear
    End If

End Function

vParameter = "dir|del requirements.txt|pip freeze > requirements.txt"
Call StartConfigurations(vParameter)