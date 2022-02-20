Attribute VB_Name = "SAP_Script_SO_PO"
Sub PO_Automation_SAP_Script()
    
    ' Created by Renato Santos
    ' Macro to run SAP Script based on specific customer project template for Sales Orders & Purchase Orders
    ' Macro allows to run multiple templates
    ' Specific templates already have automatic fields based on dependent dropdowns
    
    Dim Appl As Object
    Dim Connection As Object
    Dim session As Object
    Dim WshShell As Object
    Dim SapGui As Object
    
    Dim FileToOpen As Variant
    Dim SelectedBook As Workbook
    Dim TempSheet As Worksheet
    Dim FileCnt As Byte
    Dim StartRow As Byte: StartRow = 13
    Dim i As Byte, z As Byte
    Dim TempLR As Byte
    Dim SavePath As String
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .StatusBar = "Macro is running, please wait..."
    End With
    
    ' getting connection to SAP Server
    
    'Replace directory if needed
    Shell "C:\WINCOR-NIXDORF\SAP\FrontEnd\SAPgui\saplogon.exe", 4
    Set WshShell = CreateObject("WScript.Shell")
    
    Do Until WshShell.AppActivate("SAP Logon ")
        Application.Wait Now + TimeValue("0:00:01")
    Loop
    
    Set WshShell = Nothing
    Set SapGui = GetObject("SAPGUI")
    Set Appl = SapGui.GetScriptingEngine
    'Choose SAP Server
    Set Connection = Appl.Openconnection("220. ERP Europe/America      - Login without password", _
    True)
    Set session = Connection.Children(0)
    
    'if You need to use password and username
    'session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "900"
    'session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "user"
    'session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "password"
    'session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "EN"
    
    ' if you already have existing SAP session
    If session.Children.Count > 1 Then
        ' to click in "Continue with this logon, whithout ending any other logons in system..."
        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select
        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").SetFocus
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If

    session.findById("wnd[0]").maximize
    session.findById("wnd[0]").sendVKey 0
    
    
    ' import all templates
    FileToOpen = Application.GetOpenFilename(Filefilter:="Excel Files (*.xlsx),*xlsx", Title:="Select PO Template to import", MultiSelect:=True)
    If IsArray(FileToOpen) Then
        For FileCnt = 1 To UBound(FileToOpen)
            Set SelectedBook = Workbooks.Open(FileToOpen(FileCnt))
            Set TempSheet = SelectedBook.Sheets(1)
            TempLR = TempSheet.Range("A" & Rows.Count).End(xlUp).Row
            'setting item number in SAP starting with Z=0
            z = 0
            ' SAP script goes here
            ' Sales Order Header
            'session.findById("wnd[0]").maximize
            ' this makes the session to run minimized so that user don't see it (in background)
            session.findById("wnd[0]").iconify
            ' Go to t-code VA01
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva01"
            session.findById("wnd[0]").sendVKey 0
            ' Fill in sales org data
            session.findById("wnd[0]/usr/ctxtVBAK-AUART").Text = "zxav"
            session.findById("wnd[0]/usr/ctxtVBAK-VKORG").Text = "50z1"
            session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").Text = "04"
            session.findById("wnd[0]/usr/ctxtVBAK-SPART").Text = "sr"
            session.findById("wnd[0]/usr/ctxtVBAK-VKBUR").Text = "z101"
            session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").Text = "xbp"
            session.findById("wnd[0]").sendVKey 0
            ' PO #
            session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").Text = TempSheet.Range("B8").Value
            ' Customer
            session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").Text = TempSheet.Range("B1").Value
            ' Consignee
            session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").Text = TempSheet.Range("B1").Value
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
            ' Order Reason
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-AUGRU").Key = "AV"
            session.findById("wnd[0]").sendVKey 0
            ' Go to header level
            session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13").Select
            ' Business Line
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBKD-Y3SBL").Text = "imac"
            session.findById("wnd[0]").sendVKey 0
            ' Header Text
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09").Select
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = TempSheet.Range("B9").Value + vbCr + "" + vbCr + ""
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select
            ' Setting partners
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,4]").Text = "70019174"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,7]").Key = "Z1"
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,7]").Text = "70019153"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            ' Sales Order Item
                For i = StartRow To TempLR
                    ' Material number
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1," & z & "]").Text = TempSheet.Range("B" & i).Value
                    ' Quantity
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2," & z & "]").Text = TempSheet.Range("D" & i).Value
                    ' Text
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-ARKTX[5," & z & "]").Text = "."
                    session.findById("wnd[0]").sendVKey 0
                    ' to select correct line to go to item level
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").getAbsoluteRow(z).Selected = True
                    session.findById("wnd[0]").sendVKey 2
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\14").Select
                    ' business line item level
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\14/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/ctxtVBKD-Y3SBL").Text = TempSheet.Range("F" & i).Value
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09").Select
                    ' item text
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/btnTP_DELETE").press
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = TempSheet.Range("C" & i).Value + vbCr + "" + vbCr + ""
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05").Select
                    ' choose price condition ZAVL
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[0,8]").Text = "zavl"
                    ' amount in price
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[2,8]").Text = TempSheet.Range("E" & i).Value
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]/tbar[0]/btn[3]").press
                    ' next item
                    z = z + 1
            Next i
            ' Save Sales Order
            session.findById("wnd[0]/tbar[0]/btn[11]").press
            ' Get Sales Order number to excel template
            TempSheet.Range("C34").Value = Mid(session.findById("wnd[0]/sbar").Text, 17, 7)
            ' fill in PO table the SO to have connection PO<->SO in ME21N
            TempSheet.Range("Q13:Q" & TempLR) = TempSheet.Range("C34").Value
            
            ' Script for SAP ME21N
            ' Header PO
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme21n"
            session.findById("wnd[0]").sendVKey 0
            ' Vendor
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").Text = TempSheet.Range("I1").Value
            session.findById("wnd[0]").sendVKey 0
            ' Org Data
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG").Text = TempSheet.Range("I2").Value
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP").Text = TempSheet.Range("I3").Value
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").Text = TempSheet.Range("I4").Value
            session.findById("wnd[0]").sendVKey 0
            ' Header Text
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3").Select
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell").Text = TempSheet.Range("I4").Value + vbCr + "" + vbCr + ""
            ' Item PO
            z = 0
            For i = StartRow To TempLR
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[3," & z & "]").Text = "s"
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EPSTP[2," & z & "]").Text = "s"
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[5," & z & "]").Text = TempSheet.Range("K" & i).Value
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-TXZ01[4," & z & "]").Text = "."
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6," & z & "]").Text = TempSheet.Range("M" & i).Value
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[9," & z & "]").Text = TempSheet.Range("N" & i).Value
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[13," & z & "]").Text = TempSheet.Range("O" & i).Value
                session.findById("wnd[0]").sendVKey 0
                ' Change to account assignment tab
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT12").Select
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT12").Select
                ' Choose G/L Account
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT12/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/ctxtMEACCT1100-SAKTO").Text = TempSheet.Range("P" & i).Value
                ' Choose Sales Order that we created before
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT12/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KDAUF").Text = TempSheet.Range("Q" & i).Value
                ' Choose Item in SO
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT12/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KDPOS").Text = TempSheet.Range("R" & i).Value
                ' Change to text tab
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14").Select
                ' Change to item text
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell").selectedNode = "F03"
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell").Text = TempSheet.Range("L" & i).Value + vbCr + ""
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]").sendVKey 0
                ' Colapse item view for next line to have SAPLMEGUI:0013
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON").press
                z = z + 1
            Next i
            ' loop to fix prices that change automatically due to automatic price conditions
            z = 0
            For i = StartRow To TempLR
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[9," & z & "]").Text = TempSheet.Range("N" & i).Value
                z = z + 1
            Next i
            ' setting up the output for PDF release
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/tbar[1]/btn[21]").press
            session.findById("wnd[0]/tbar[1]/btn[5]").press
            session.findById("wnd[0]/usr/cmbNAST-VSZTP").Key = "4"
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            session.findById("wnd[0]/tbar[1]/btn[2]").press
            session.findById("wnd[0]/usr/chkNAST-DELET").Selected = True
            session.findById("wnd[0]/usr/ctxtNAST-LDEST").Text = "MAIL_PDF"
            session.findById("wnd[0]/usr/txtNAST-ANZAL").Text = "1"
            session.findById("wnd[0]/usr/cmbNAST-TDARMOD").Key = "2"
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            ' Save
            session.findById("wnd[0]/tbar[0]/btn[11]").press
            ' Get PO number to excel
            TempSheet.Range("J34").Value = Right(session.findById("wnd[0]/sbar").Text, 10)
            ' Closes session
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nex"
            session.findById("wnd[0]").sendVKey 0
            SavePath = ThisWorkbook.Path
            SelectedBook.SaveAs SavePath & "\" & TempSheet.Range("C34").Value & "_" & TempSheet.Range("J34").Value & ".xlsx"
            SelectedBook.Close False
            'next template
        Next FileCnt
    End If
    
    MsgBox "SO's and PO's created!", vbInformation, "Task Successful"
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .StatusBar = "Macro is running, please wait..."
    End With
    
End Sub
