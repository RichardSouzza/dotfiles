Option Explicit

Dim fso, shell, logPath, logFile

Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("wscript.shell")

logPath = shell.ExpandEnvironmentStrings("%USERPROFILE%\.startup.log")

Set logFile = fso.OpenTextFile(logPath, 2, True)


' =====================================
'  Logging functions
' =====================================

Sub LogMessage(msg)
    logFile.WriteLine Now & " | " & msg
End Sub

Sub LogError(msg)
    logFile.WriteLine Now & " | [ERROR]" & msg
End Sub

' =====================================
'  Execution functions
' =====================================

Function RunAndLog(cmd)
    LogMessage "Running: " & cmd
    Dim execObj, line
    Set execObj = shell.Exec("cmd /c " & cmd)

    Do While Not execObj.StdOut.AtEndOfStream
        line = execObj.StdOut.ReadLine
        LogMessage line
    Loop

    Do While Not execObj.StdErr.AtEndOfStream
        line = execObj.StdErr.ReadLine
        LogMessage line
    Loop

    RunAndLog = execObj.ExitCode
End Function

' =====================================
'  Main execution
' =====================================

LogMessage "===== Starting System Setup ====="

RunAndLog "komorebic start --whkd"
RunAndLog "scoop cache rm *"
RunAndLog "scoop cleanup *"
RunAndLog "scoop update *"
RunAndLog "windhawk -tray-only"

LogMessage "===== Ending System Setup ====="

logFile.Close
