' ==============================================================================
' ZPPM6680 Automation Worker
' Usage: cscript //Nologo zppm6680.vbs <Plant> <StartMonth> <EndMonth>
' Example: cscript //Nologo zppm6680.vbs 1840 2026.08 2027.01
' ==============================================================================
Option Explicit

Dim plant, startMonth, endMonth
If WScript.Arguments.Count < 3 Then
    WScript.Echo ERROR: Arguments missing. Usage: zppm6680.vbs <Plant> <StartMonth> <EndMonth>
    WScript.Quit 1
End If

plant = Trim(WScript.Arguments(0))
startMonth = Trim(WScript.Arguments(1))
endMonth = Trim(WScript.Arguments(2))

On Error Resume Next
Dim SapGuiAuto, application, connection, session
Set SapGuiAuto = GetObject(SAPGUI)
If Err.Number <> 0 Then
    WScript.Echo ERROR: SAP GUI is not running. Please launch and login to SAP Logon first.
    WScript.Quit 2
End If

Set application = SapGuiAuto.GetScriptingEngine
If Err.Number <> 0 Then
    WScript.Echo ERROR: SAP GUI Scripting is not enabled. Please check SAP GUI Options.
    WScript.Quit 3
End If

Set connection = application.Children(0)
If Err.Number <> 0 Then
    WScript.Echo ERROR: No active SAP connection found.
    WScript.Quit 4
End If

Set session = connection.Children(0)
If Err.Number <> 0 Then
    WScript.Echo ERROR: No active SAP session found.
    WScript.Quit 5
End If
On Error GoTo 0

' Navigate to ZPPM6680
session.findById(wnd[0]/tbar[0]/okcd).text = /nZPPM6680
session.findById(wnd[0]).sendVKey 0

' Fill selection criteria
session.findById(wnd[0]/usr/ctxtPA_WERKS).text = plant
session.findById(wnd[0]/usr/txtSO_EMONU-LOW).text = startMonth
session.findById(wnd[0]/usr/txtSO_EMONU-HIGH).text = endMonth
On Error Resume Next
session.findById(wnd[0]/usr/ctxtSO_VERID-LOW).text = "
On Error GoTo 0

' Execute (F8)
session.findById(wnd[0]).sendVKey 8

' Export to MHTML via ALV Grid Context Menu (&XXL)
session.findById(wnd[0]/usr/cntlCC_CONTAINER_0100/shellcont/shell/shellcont/shell).currentCellRow = 1
session.findById(wnd[0]/usr/cntlCC_CONTAINER_0100/shellcont/shell/shellcont/shell).contextMenu
session.findById(wnd[0]/usr/cntlCC_CONTAINER_0100/shellcont/shell/shellcont/shell).selectContextMenuItem &XXL

' Handle export popup dialogs
session.findById(wnd[1]/tbar[0]/btn[8]).press
On Error Resume Next
session.findById(wnd[1]/tbar[0]/btn[0]).press
session.findById(wnd[1]/tbar[0]/btn[11]).press
On Error GoTo 0

WScript.Echo SUCCESS: ZPPM6680 completed for Plant  & plant
WScript.Quit 0