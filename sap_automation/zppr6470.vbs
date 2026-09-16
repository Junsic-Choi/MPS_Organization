' ==============================================================================
' ZPPR6470 Automation Worker
' Reads Sales Orders from Clipboard and Exports Component Requirements
' Usage: cscript //Nologo zppr6470.vbs <Plant> <BESKZ> <SOBSL> <LayoutRow>
' Example 1840: cscript //Nologo zppr6470.vbs 1840 e " 18
' Example 1842: cscript //Nologo zppr6470.vbs  F 44 18
' ==============================================================================
Option Explicit

Dim plant, beskz, sobsl, layoutRow
plant = 
beskz = 
sobsl = 
layoutRow = 18

If WScript.Arguments.Count >= 1 Then plant = Trim(WScript.Arguments(0))
If WScript.Arguments.Count >= 2 Then beskz = Trim(WScript.Arguments(1))
If WScript.Arguments.Count >= 3 Then sobsl = Trim(WScript.Arguments(2))
If WScript.Arguments.Count >= 4 Then layoutRow = Trim(WScript.Arguments(3))

On Error Resume Next
Dim SapGuiAuto, application, connection, session
Set SapGuiAuto = GetObject(SAPGUI)
If Err.Number <> 0 Then
 WScript.Echo ERROR: SAP GUI is not running. Please launch and login to SAP Logon first.
 WScript.Quit 2
End If

Set application = SapGuiAuto.GetScriptingEngine
If Err.Number <> 0 Then
 WScript.Echo ERROR: SAP GUI Scripting is not enabled.
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

' Navigate to ZPPR6470
session.findById(wnd[0]/tbar[0]/okcd).text = /nZPPR6470
session.findById(wnd[0]).sendVKey 0

' Click Sales Doc Multiple Selection Button
session.findById(wnd[0]/usr/btn%_SO_VBELN_%_APP_%-VALU_PUSH).press

' Upload from Clipboard (btn[24]) and Confirm (btn[8])
session.findById(wnd[1]/tbar[0]/btn[24]).press
session.findById(wnd[1]/tbar[0]/btn[8]).press

' Set Plant (if specified)
If plant <>  Then
 session.findById(wnd[0]/usr/ctxtSO_WERKS-LOW).text = plant
Else
 session.findById(wnd[0]/usr/ctxtSO_WERKS-LOW).text = 
End If

' Set BESKZ (조달구분)
If beskz <>  Then
 session.findById(wnd[0]/usr/ctxtSO_BESKZ-LOW).text = beskz
End If

' Set SOBSL (특별조달)
If sobsl <>  Then
 session.findById(wnd[0]/usr/ctxtSO_SOBSL-LOW).text = sobsl
Else
 session.findById(wnd[0]/usr/ctxtSO_SOBSL-LOW).text = 
End If

' Execute (F8)
session.findById(wnd[0]).sendVKey 8

' Select Layout Variant
session.findById(wnd[0]/usr/cntlCON_S100/shellcont/shell).pressToolbarButton &MB_VARIANT
session.findById(wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell).setCurrentCell CInt(layoutRow), TEXT
session.findById(wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell).selectedRows = CStr(layoutRow)
session.findById(wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell).clickCurrentCell

' Export to XXL
session.findById(wnd[0]/usr/cntlCON_S100/shellcont/shell).currentCellRow = 1
session.findById(wnd[0]/usr/cntlCON_S100/shellcont/shell).contextMenu
session.findById(wnd[0]/usr/cntlCON_S100/shellcont/shell).selectContextMenuItem &XXL

' Handle export popup dialogs
session.findById(wnd[1]/tbar[0]/btn[8]).press
On Error Resume Next
session.findById(wnd[1]/tbar[0]/btn[0]).press
session.findById(wnd[1]/tbar[0]/btn[11]).press
On Error GoTo 0

WScript.Echo SUCCESS: ZPPR6470 completed
WScript.Quit 0