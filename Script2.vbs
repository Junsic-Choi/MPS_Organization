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
session.findById("wnd[0]/usr/cntlCC_CONTAINER_0100/shellcont/shell/shellcont/shell").currentCellRow = 6
session.findById("wnd[0]/usr/cntlCC_CONTAINER_0100/shellcont/shell/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlCC_CONTAINER_0100/shellcont/shell/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/usr/cntlCC_CONTAINER_0100/shellcont/shell/shellcont/shell").setCurrentCell 24,"MATNR"
