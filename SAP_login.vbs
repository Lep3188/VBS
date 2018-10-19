

BoxName_Example = "ECC - Q10 - Quality Test System" 'This is the name or description of the SAP Enviroment to access.

set WshShell = CreateObject("WScript.Shell")
 Set proc = WshShell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\SAPLGPAD.EXE") 'Executable name might change depending on your SAP LogonPad
           
   Set SapGui = GetObject("SAPGUI")
Set Appl = SapGui.GetScriptingEngine

Set Connection = Appl.Openconnection(BoxName, True)
Set session = Connection.Children(0)

'Remember that the GUI scripting below will run depending the SAP enviroment(box setup), do your own login recording and paste below

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "900" 'Client
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "*******" 'UserID
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "*******" 'Password
session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus
session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 8
session.findById("wnd[0]").sendVKey 0
