Option Explicit
Dim SapGuiAuto, application, connection, session
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Or SapGuiAuto Is Nothing Then
    WScript.Echo "ERROR_NO_SAPGUI: SAP Logon is not running. Please launch SAP Logon first."
    WScript.Quit 1
End If

Set application = SapGuiAuto.GetScriptingEngine
If Err.Number <> 0 Or application Is Nothing Then
    WScript.Echo "ERROR_NO_SCRIPTING: SAP GUI Scripting is disabled. Please enable scripting in SAP options."
    WScript.Quit 2
End If

If application.Children.Count = 0 Then
    WScript.Echo "ERROR_NO_CONNECTION: No active SAP connection. Please log in to your SAP system."
    WScript.Quit 3
End If

Set connection = application.Children(0)
If connection.Children.Count = 0 Then
    WScript.Echo "ERROR_NO_SESSION: No open SAP session window found."
    WScript.Quit 4
End If

Set session = connection.Children(0)
If IsObject(WScript) Then
    WScript.ConnectObject session,     "on"
    WScript.ConnectObject application, "on"
End If
On Error GoTo 0

session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.Echo "SUCCESS: Returned to SAP Home"
WScript.Quit 0