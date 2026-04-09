Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
' Automatically find current folder
currentDir = fso.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = currentDir

' Run node server.js in background (0 = hide window)
WshShell.Run "node server.js", 0, False

MsgBox "MPS Dashboard Server (Port: 8890) started in background." & vbCrLf & "You can access it at: http://localhost:8890/dashboard.html", vbInformation, "MPS Silent Startup"
