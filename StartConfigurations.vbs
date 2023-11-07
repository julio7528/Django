Sub StartConfigurations(iParameter)
    '["""]
    'Name: Function StartConfigurations
    'Functionality: This function performs a series of configurations based on the input parameters.

    'Input Values:
    '    - iParameter (String): A string containing three parameters separated by '|'. The parameters are directory path, delete requirements, and create requirements.

    'Output Values:
    '   - None

    'Returns:
    '    - None

    'Description of Process:
    '    - Step 1. Split the input parameter into individual values and assign them to variables.
    '    - Step 2. Attempt to execute a command for each value in vParameter.
    '    - Step 3. Create a shell object and execute the command using it.
    '    - Step 4. Print the output of the command line by line.
    '    - Step 5. Handle any errors that may occur during execution.

    'Exception Description:
    '    - If an error occurs during execution, it is caught using "On Error Resume Next" and the details are printed.

    'Exception Return:
    '    - None
    
    'IMPORTANT: ENABLE WSCRIPT.ECHO TO SEE THE OUTPUT OF THE COMMANDS
    '["""]

    vParameter = Split(iParameter, "|")
    vDir = vParameter(0)
    vDelRequirements = vParameter(1)
    vCreateRequirements = vParameter(2)

    On Error Resume Next

    RecordLog("----------------------------------------")
    RecordLog("Initalized: StartConfigurations")

    vCommand = Array(cDir, vDelRequirements, vCreateRequirements)
    For Each vCommand In vParameter
        Set objShell = CreateObject("WScript.Shell")
        ' Execute the command and get an Exec object
        Set objExec = objShell.Exec("cmd /c " & vCommand)
        'WScript.Echo "Start: " & vCommand
        RecordLog(vCommand)
        ' Loop over each line of the command's output
        Do While Not objExec.StdOut.AtEndOfStream
            strLine = objExec.StdOut.ReadLine
            'WScript.Echo strLine
            RecordLog(strLine)
        Loop
        'WScript.Echo "End: " & vCommand
        'WScript.Echo "----------------------------------------"
    Next
    ' If there was an error, print it
    If Err.Number <> 0 Then
        'WScript.Echo "Error: " & Err.Number & " " & Err.Description
        MsgBox "Error: " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        'WScript.Echo "No Errors"
        MsgBox "Executed Successfully"
    End If

    RecordLog("Finished: StartConfigurations")
    RecordLog("----------------------------------------")
End Sub

Sub RecordLog(iMessage)
    Dim objFSO, objArquivo
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objArquivo = objFSO.OpenTextFile("StartConfigurationsLog.txt", 8, True)

    vCurrentTime = FormatDateTime(Now, vbShortDate) & " - " & FormatDateTime(Now, vbLongTime)
    'Example Output vCurrentTime Variable: 11/7/2023 - 2:04:49 PM

    vLog = vCurrentTime & " - " & iMessage
    objArquivo.WriteLine vLog
    objArquivo.Close
End Sub

vParameter = "dir|del requirements.txt|pip freeze > requirements.txt"
Call StartConfigurations(vParameter)