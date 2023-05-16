Sub Auto_Open()
Application.DisplayAlerts = False
Dim SapGui, Applic, Connection, Session, WSHShell

'Open SAP GUI
Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus

'Set object SapGui
Set WSHShell = CreateObject("Wscript.Shell")


Do Until WSHShell.AppActivate("SAP Logon")
    Application.Wait Now + TimeValue("0:00:01")
Loop

'Clear object WSHShell
Set WSHShell = Nothing

Set SapGui = GetObject("SAPGUI")


Set Applic = SapGui.GetScriptingEngine

Set Connection = Applic.OpenConnection("# -E05 - ECC - Produção / Producción / Production", True)
Set Session = Connection.Children(0)

Session.findById("wnd[0]").maximize
Session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "150"
Session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = user
Session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = password
Session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "PT"
Session.findById("wnd[0]").sendVKey 0




'_


Set Connection = Nothing
End Sub
