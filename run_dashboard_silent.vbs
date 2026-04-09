Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
' Ensure the script runs in its own directory
currentDir = fso.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = currentDir
WshShell.Run "node server.js", 0, False
MsgBox "MPS Dashboard Server started in background." & vbCrLf & "http://localhost:8889/dashboard.html", vbInformation, "MPS Dashboard"
