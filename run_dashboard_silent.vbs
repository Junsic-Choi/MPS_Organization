Set shell = CreateObject("Wscript.Shell")
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists("mps_dashboard_app.exe") Then
    shell.Run "mps_dashboard_app.exe", 0, False
Else
    shell.Run "node server.js", 0, False
End If
