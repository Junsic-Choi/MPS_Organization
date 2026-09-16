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

Dim plant, startMonth, endMonth
plant = "1840"
startMonth = "2026.08"
endMonth = "2027.01"

If WScript.Arguments.Count >= 1 Then plant = Trim(WScript.Arguments(0))
If WScript.Arguments.Count >= 2 Then startMonth = Trim(WScript.Arguments(1))
If WScript.Arguments.Count >= 3 Then endMonth = Trim(WScript.Arguments(2))

session.findById("wnd[0]/tbar[0]/okcd").text = "/nZPPM6680"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/ctxtPA_WERKS").text = plant
session.findById("wnd[0]/usr/txtSO_EMONU-LOW").text = startMonth
session.findById("wnd[0]/usr/txtSO_EMONU-HIGH").text = endMonth
session.findById("wnd[0]").sendVKey 8

session.findById("wnd[0]/usr/cntlCC_CONTAINER_0100/shellcont/shell/shellcont/shell").currentCellRow = 1
session.findById("wnd[0]/usr/cntlCC_CONTAINER_0100/shellcont/shell/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlCC_CONTAINER_0100/shellcont/shell/shellcont/shell").selectContextMenuItem "&XXL"

session.findById("wnd[1]/tbar[0]/btn[8]").press
On Error Resume Next
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press
On Error GoTo 0

WScript.Echo "SUCCESS: ZPPM6680 completed for Plant " & plant
WScript.Quit 0