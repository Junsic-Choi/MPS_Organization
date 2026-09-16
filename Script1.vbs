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
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00030"
session.findById("wnd[0]/usr/ctxtPA_WERKS").text = "1840"
session.findById("wnd[0]/usr/txtSO_EMONU-LOW").text = "2026.08"
session.findById("wnd[0]/usr/txtSO_EMONU-HIGH").text = "2027.01"
session.findById("wnd[0]/usr/ctxtSO_VERID-LOW").setFocus
session.findById("wnd[0]/usr/ctxtSO_VERID-LOW").caretPosition = 0
session.findById("wnd[0]").sendVKey 8
