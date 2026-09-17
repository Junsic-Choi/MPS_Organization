Option Explicit

Dim SapGuiAuto, application, connection, session
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Or SapGuiAuto Is Nothing Then WScript.Quit 0

Set application = SapGuiAuto.GetScriptingEngine
If Err.Number <> 0 Or application Is Nothing Then WScript.Quit 0

If application.Children.Count = 0 Then WScript.Quit 0
Set connection = application.Children(0)

If connection.Children.Count = 0 Then WScript.Quit 0
Set session = connection.Children(0)

' If any modal popup window wnd[1] is open, dismiss it
session.findById("wnd[1]/tbar[0]/btn[12]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press

' Return to home
session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
session.findById("wnd[0]").sendVKey 0
On Error GoTo 0
WScript.Echo "SUCCESS: Returned to SAP Home"
WScript.Quit 0