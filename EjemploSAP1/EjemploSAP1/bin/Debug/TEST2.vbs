dim aTest
aTest = qParams

dim forImpre

Dim args
Set args = WScript.Arguments

'msgbox("Hello " & args.Item(0)) 

    Set SapGuiAuto = GetObject("SAPGUI")
    Set sapApplication = SapGuiAuto.GetScriptingEngine
    Set Connection = sapApplication.Children(0)
    Set session = Connection.Children(0)
       
   


session.findById("wnd[0]").resizeWorkingPane 256,37,false
session.findById("wnd[0]/tbar[0]/okcd").text = "ZMFP11"
session.findById("wnd[0]").sendVKey 0


session.findById("wnd[0]/usr/subSPLITTED_SCREEN:SAPLVHUDIAL2:0053/subAREA1:SAPLVHUDIAL2:0110/tabsTABSTRIP1/tabpTAB1/ssubSUB1:SAPLVHUDIAL2:0200/ctxtPLAPPLDATA-MATNR").text = args.Item(0)
session.findById("wnd[0]/usr/subSPLITTED_SCREEN:SAPLVHUDIAL2:0053/subAREA1:SAPLVHUDIAL2:0110/tabsTABSTRIP1/tabpTAB1/ssubSUB1:SAPLVHUDIAL2:0200/ctxtPLAPPLDATA-WERKS").text = "5210"
session.findById("wnd[0]/usr/subSPLITTED_SCREEN:SAPLVHUDIAL2:0053/subAREA1:SAPLVHUDIAL2:0110/tabsTABSTRIP1/tabpTAB1/ssubSUB1:SAPLVHUDIAL2:0200/txtPLAPPLDATA-MAXQUA").text = args.Item(1)
session.findById("wnd[0]/usr/subSPLITTED_SCREEN:SAPLVHUDIAL2:0053/subAREA1:SAPLVHUDIAL2:0110/tabsTABSTRIP1/tabpTAB1/ssubSUB1:SAPLVHUDIAL2:0200/txtPLAPPLDATA-MAXQUA").setFocus
session.findById("wnd[0]/usr/subSPLITTED_SCREEN:SAPLVHUDIAL2:0053/subAREA1:SAPLVHUDIAL2:0110/tabsTABSTRIP1/tabpTAB1/ssubSUB1:SAPLVHUDIAL2:0200/txtPLAPPLDATA-MAXQUA").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subSPLITTED_SCREEN:SAPLVHUDIAL2:0052/subAREA2:SAPLVHUDIAL2:0061/subPACKDIALOG:SAPLVHUSUBSC:0100/btnCREATE_HUS").press

'forImpre = session.findById("wnd[0]/usr/subSPLITTED_SCREEN:SAPLVHUDIAL2:0053/subAREA3:SAPLVHUDIAL2:0072/cntlHUCONTAINER/shellcont/shell").selectItem "1","3".Text

'msgbox("Hello " + ) 

session.findById("wnd[0]").resizeWorkingPane 256,37,false
session.findById("wnd[0]/mbar/menu[0]/menu[10]").select




'session.findById("wnd[0]").resizeWorkingPane 256,37,false
'session.findById("wnd[0]/tbar[0]/okcd").text = "z_uc_del"
'session.findById("wnd[0]").sendVKey 0


'session.findById("wnd[0]").resizeWorkingPane 256,37,false
'session.findById("wnd[0]/usr/chkCOPY").selected = true
'session.findById("wnd[0]/usr/ctxtV_DISPO").text = "001"
'session.findById("wnd[0]/usr/ctxtB_DISPO").text = "999"
'session.findById("wnd[0]/usr/ctxtP_LDEST").text = "S528"
'session.findById("wnd[0]/usr/chkCOPY").setFocus


'session.findById("wnd[0]/tbar[1]/btn[8]").press









'session.findById("wnd[0]").resizeWorkingPane 256,37,false
'session.findById("wnd[0]/mbar/menu[0]/menu[10]").select

'msgbox("Hello " + session.findById("wnd[1]/usr/lbl%_AUTOTEXT003").Text + " ... " + session.findById("wnd[1]/usr/lbl%_AUTOTEXT013").Text) 