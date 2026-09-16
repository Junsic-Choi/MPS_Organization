On Error Resume Next
Set SapGuiAuto = GetObject(SAPGUI)
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children(0)
Set session = connection.Children(0)
session.findById(wnd[0]/tbar[0]/okcd).text = /n
session.findById(wnd[0]).sendVKey 0