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

Dim plant, beskz, sobsl, layoutRow
plant = "1840"
beskz = "e"
sobsl = ""
layoutRow = 18

If WScript.Arguments.Count >= 1 Then plant = Trim(WScript.Arguments(0))
If WScript.Arguments.Count >= 2 Then beskz = Trim(WScript.Arguments(1))
If WScript.Arguments.Count >= 3 Then sobsl = Trim(WScript.Arguments(2))
If WScript.Arguments.Count >= 4 Then layoutRow = Trim(WScript.Arguments(3))

session.findById("wnd[0]/tbar[0]/okcd").text = "/nZPPR6470"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/btn%_SO_VBELN_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press

If plant <> "" Then
    session.findById("wnd[0]/usr/ctxtSO_WERKS-LOW").text = plant
Else
    session.findById("wnd[0]/usr/ctxtSO_WERKS-LOW").text = ""
End If

If beskz <> "" Then
    session.findById("wnd[0]/usr/ctxtSO_BESKZ-LOW").text = beskz
Else
    session.findById("wnd[0]/usr/ctxtSO_BESKZ-LOW").text = ""
End If

If sobsl <> "" Then
    session.findById("wnd[0]/usr/ctxtSO_SOBSL-LOW").text = sobsl
Else
    session.findById("wnd[0]/usr/ctxtSO_SOBSL-LOW").text = ""
End If

session.findById("wnd[0]").sendVKey 8

Dim conShell, retryCon
Set conShell = Nothing
For retryCon = 1 To 30
    Err.Clear
    On Error Resume Next
    Set conShell = session.findById("wnd[0]/usr/cntlCON_S100/shellcont/shell")
    If Err.Number = 0 And Not conShell Is Nothing Then
        On Error GoTo 0
        Exit For
    End If
    WScript.Sleep 1000
Next
On Error GoTo 0
If conShell Is Nothing Then
    WScript.Echo "ERROR_NO_GRID: ZPPR6470 Grid container did not load within 30 seconds."
    WScript.Quit 5
End If

conShell.pressToolbarButton "&MB_VARIANT"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell CInt(layoutRow), "TEXT"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = CStr(layoutRow)
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell

conShell.currentCellRow = 1
conShell.contextMenu
conShell.selectContextMenuItem "&XXL"

HandleExportPopups

WScript.Echo "SUCCESS: ZPPR6470 completed for Plant " & plant
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