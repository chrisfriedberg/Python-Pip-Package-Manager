On Error Resume Next
Set WshShell = CreateObject("WScript.Shell")

' Get current directory
strPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' Check if the Python executable exists
pythonPath = strPath & "\venv\Scripts\pythonw.exe"
launcherPath = strPath & "\launcher.py"

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(pythonPath) Then
    WScript.Echo "Error: Python executable not found at:" & vbCrLf & pythonPath
    WScript.Quit(1)
End If

If Not fso.FileExists(launcherPath) Then
    WScript.Echo "Error: Launcher script not found at:" & vbCrLf & launcherPath
    WScript.Quit(1)
End If

' Run the Python script
WshShell.CurrentDirectory = strPath
returnCode = WshShell.Run("""" & pythonPath & """ """ & launcherPath & """", 0, False)

' Check for errors
If Err.Number <> 0 Then
    WScript.Echo "Error launching Python application:" & vbCrLf & Err.Description
    WScript.Quit(1)
End If 