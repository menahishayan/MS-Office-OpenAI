' Menahi Shayan, 2025.
'
' This project attempts to create an open source implementation of OpenAI and Gemini to 
' work with MS Office 365 Apps such as Excel, Word, and Powerpoint.
'
' For questions, comments, or concerns, please raise a PR or Issue on the GitHub repository. (https://github.com/menahishayan/MS-Office-OpenAI)
'
' Licensed under GNU General Public License v3.0. See LICENSE for more information.

Function AI(venvPath As String, argument As String) As String
    Dim pythonExe As String
    Dim pythonScript As String
    Dim command As String
    Dim result As String

    On Error GoTo ErrorHandler

    ' Determine the Python executable based on the platform
    If InStr(Application.OperatingSystem, "Mac") > 0 Then
        ' macOS: Use Python in the "bin" folder
        pythonExe = venvPath & Application.PathSeparator & "bin" & Application.PathSeparator & "python3"
    Else
        ' Windows: Use Python in the "Scripts" folder
        pythonExe = venvPath & Application.PathSeparator & "Scripts" & Application.PathSeparator & "python.exe"
    End If

    ' Path to the Python script
    pythonScript = venvPath & Application.PathSeparator & "__main__.py"

    ' Construct the shell command
    command = """" & pythonExe & """ """ & pythonScript & """ """ & argument & """"

    ' Execute the shell command
    Application.Cursor = xlWait
    If InStr(Application.OperatingSystem, "Mac") > 0 Then
        ' macOS: Use AppleScriptTask to execute
        result = ExecuteShellCommandMac(command)
    Else
        ' Windows: Use WScript.Shell to execute
        result = ExecuteShellCommandWindows(command)
    End If
    Application.Cursor = xlDefault

    ' Return the result
    RunPythonFromVenv = result
    Exit Function

ErrorHandler:
    RunPythonFromVenv = "Error: " & Err.Description
End Function

Function ExecuteShellCommandWindows(command As String) As String
    Dim wsh As Object
    Dim wExec As Object
    Dim stdOut As Object
    Dim result As String

    On Error GoTo ErrorHandler

    ' Create WScript.Shell object
    Set wsh = CreateObject("WScript.Shell")
    Set wExec = wsh.Exec(command)
    Set stdOut = wExec.stdOut

    ' Capture standard output
    result = ""
    Do While Not stdOut.AtEndOfStream
        result = result & stdOut.ReadLine() & vbCrLf
    Loop

    ExecuteShellCommandWindows = result
    Exit Function

ErrorHandler:
    ExecuteShellCommandWindows = "Error: " & Err.Description
End Function

Function ExecuteShellCommandMac(command As String) As String
    On Error GoTo ErrorHandler

    ' Use AppleScriptTask to run the shell command
    ExecuteShellCommandMac = AppleScriptTask("ShellExecutor.scpt", "runShellCommand", command)
    Exit Function

ErrorHandler:
    ExecuteShellCommandMac = "Error: " & Err.Description
End Function

