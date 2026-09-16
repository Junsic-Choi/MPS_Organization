If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").resizeWorkingPane 262,40,false
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00029"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00029"
session.findById("wnd[0]/usr/btn%_SO_VBELN_%_APP_%-VALU_PUSH").press
session.findById("wnd[0]/usr/ctxtSO_WERKS-LOW").text = "1840"
session.findById("wnd[0]/usr/ctxtSO_VBELN-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtSO_VBELN-HIGH").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtSO_BESKZ-LOW").text = "e"
session.findById("wnd[0]/usr/ctxtSO_SOBSL-LOW").setFocus
session.findById("wnd[0]/usr/ctxtSO_SOBSL-LOW").caretPosition = 0
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlCON_S100/shellcont/shell").pressToolbarButton "&MB_VARIANT"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell 18,"TEXT"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "18"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/usr/cntlCON_S100/shellcont/shell").setCurrentCell 1,"KDAUF"
session.findById("wnd[0]/usr/cntlCON_S100/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlCON_S100/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press
