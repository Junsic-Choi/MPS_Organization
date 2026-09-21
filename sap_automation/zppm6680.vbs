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

On Error Resume Next
session.findById("wnd[0]").maximize
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
On Error Resume Next
session.findById("wnd[0]/usr/ctxtSO_VERID-LOW").text = ""
session.findById("wnd[0]/usr/ctxtSO_VERID-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtSO_MATNR-LOW").text = ""
session.findById("wnd[0]/usr/ctxtSO_MATNR-HIGH").text = ""
On Error GoTo 0
session.findById("wnd[0]").sendVKey 8

Dim grid, retry
Set grid = Nothing
For retry = 1 To 30
    Err.Clear
    On Error Resume Next
    Set grid = session.findById("wnd[0]/usr/cntlCC_CONTAINER_0100/shellcont/shell/shellcont/shell")
    If Err.Number = 0 And Not grid Is Nothing Then
        On Error GoTo 0
        Exit For
    End If
    WScript.Sleep 1000
Next
On Error GoTo 0
If grid Is Nothing Then
    WScript.Echo "ERROR_NO_GRID: ZPPM6680 ALV Grid did not load within 30 seconds."
    WScript.Quit 5
End If

If grid.rowCount = 0 Then
    WScript.Echo "ERROR_ZERO_ROWS: ZPPM6680 Plant " & plant & " (" & startMonth & " ~ " & endMonth & ") has 0 planned orders."
    WScript.Quit 6
End If

grid.currentCellRow = 1
grid.contextMenu
grid.selectContextMenuItem "&XXL"

HandleExportPopups

WScript.Echo "SUCCESS: ZPPM6680 completed for Plant " & plant
WScript.Quit 0

Sub HandleExportPopups()
    Dim loopCount, wnd
    For loopCount = 1 To 25
        WScript.Sleep 1000
        On Error Resume Next
        Err.Clear
        Set wnd = session.findById("wnd[1]")
        If Err.Number = 0 And Not wnd Is Nothing Then
            Err.Clear
            wnd.findById("tbar[0]/btn[11]").press
            If Err.Number <> 0 Then
                Err.Clear
                wnd.findById("tbar[0]/btn[8]").press
                If Err.Number <> 0 Then
                    Err.Clear
                    wnd.findById("tbar[0]/btn[0]").press
                End If
            End If
        Else
            If loopCount >= 3 Then
                On Error GoTo 0
                Exit Sub
            End If
        End If
        On Error GoTo 0
    Next
End Sub