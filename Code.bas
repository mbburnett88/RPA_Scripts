Attribute VB_Name = "Code"
Option Explicit
Public CountOfPlannedOrder As Integer
Public SumQtyOfOrders As Long
Public ProdSup As String
Public UnRStock As Long
Public RStock As Long
Public BStock As Long
Sub btnSaveMasterRouting()
    Dim I As Integer
    Dim Z As Integer
    Application.DisplayAlerts = False
    'add code to replace material routing with the one on the MaterialRouting tab
    If Sheets("Settings").Range("B13") = "Production" Then Call LogonProduction Else Call LogonDevelopment
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nca02"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ctxtRC27M-MATNR").Text = Range("Q1") '"XA20-101A"
    Session.FindById("wnd[0]/usr/ctxtRC27M-WERKS").Text = "3000"
    Session.FindById("wnd[0]/usr/ctxtRC271-PLNNR").Text = ""
    Session.FindById("wnd[0]").SendVKey 8
    Session.FindById("wnd[0]/tbar[1]/btn[26]").press
    Session.FindById("wnd[0]/tbar[1]/btn[14]").press
    Session.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
    I = 0
    Z = 0
    While Range("A" & Z + 2) <> ""
        Session.FindById("wnd[0]").ResizeWorkingPane 177, 42, False
        Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VORNR[0," & I & "]").Text = Range("A" & Z + 2) '"0001"
        Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-ARBPL[2," & I & "]").Text = Range("B" & Z + 2) '"3500"
        Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-WERKS[3," & I & "]").Text = "3000"
        Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-STEUS[4," & I & "]").Text = Range("C" & Z + 2) '"ZP01"
        Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-KTSCH[5," & I & "]").Text = Range("M" & Z + 2) '"AB"
        Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6," & I & "]").Text = Range("D" & Z + 2) '"Start"
        'add note to step
        If Range("O" & Z + 2) <> "" Then
            While Len(Range("O" & Z + 2)) > 72
                Range("O" & Z + 2) = InputBox("Shorten your note.", "Note is too long.", Range("O" & Z + 2))
                DoEvents
            Wend
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/chkRC270-TXTKZ[7," & I & "]").SetFocus
            Session.FindById("wnd[0]").SendVKey 2
            Session.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select
            Session.FindById("wnd[0]").SendVKey 0
            Session.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,2]").Text = Range("O" & Z + 2) '"test"
            Session.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,2]").CaretPosition = 72
            Session.FindById("wnd[0]").SendVKey 0
            Session.FindById("wnd[0]/tbar[0]/btn[3]").press
        End If
        'add IS to step
        If InStr(UCase(Range("D" & Z + 2)), "ADSORB") <> 0 Or InStr(UCase(Range("D" & Z + 2)), "AFFINITY") <> 0 Then
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/chkPLPOD-PRTKZ[8," & I & "]").SetFocus
            Session.FindById("wnd[0]").SendVKey 2
            Session.FindById("wnd[1]/usr/ctxtPLFHD-MATNR").Text = Range("N" & Z + 2) '"ISAD20-101"
            Session.FindById("wnd[1]/usr/ctxtPLFHD-STEUF").Text = "1"
            Session.FindById("wnd[1]/tbar[0]/btn[3]").press
            Session.FindById("wnd[0]/usr/tblSAPLCFDITCTRL_0100").GetAbsoluteRow(0).Selected = True
            Session.FindById("wnd[0]/tbar[1]/btn[28]").press
            Session.FindById("wnd[0]/usr/ctxtPLFHD-MGFORM").Text = "ZAP005"
            Session.FindById("wnd[0]/tbar[0]/btn[3]").press
            Session.FindById("wnd[0]/tbar[0]/btn[3]").press
        End If
        Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-BMSCH[14," & I & "]").Text = Range("E" & Z + 2) '"150"
        Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[16," & I & "]").SetFocus
        Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[16," & I & "]").CaretPosition = 0
        Session.FindById("wnd[0]").SendVKey 2
        Session.FindById("wnd[0]/usr/subDEFAULTVAL:SAPLCPDO:1211/txtPLPOD-VGW01").Text = Range("G" & Z + 2) '""
        Session.FindById("wnd[0]/usr/subDEFAULTVAL:SAPLCPDO:1211/ctxtPLPOD-VGE01").Text = Range("H" & Z + 2) '"H"
        Session.FindById("wnd[0]/usr/subDEFAULTVAL:SAPLCPDO:1211/ctxtPLPOD-LAR01").Text = "" 'Clear setup activity type
        Session.FindById("wnd[0]/usr/subDEFAULTVAL:SAPLCPDO:1211/txtPLPOD-VGW02").Text = Range("I" & Z + 2) '""
        Session.FindById("wnd[0]/usr/subDEFAULTVAL:SAPLCPDO:1211/ctxtPLPOD-VGE02").Text = Range("J" & Z + 2) '"H"
        Session.FindById("wnd[0]/usr/subDEFAULTVAL:SAPLCPDO:1211/ctxtPLPOD-LAR02").Text = "" 'Clear machine activity type
        Session.FindById("wnd[0]/usr/subDEFAULTVAL:SAPLCPDO:1211/txtPLPOD-VGW03").Text = Range("K" & Z + 2) '""
        Session.FindById("wnd[0]/usr/subDEFAULTVAL:SAPLCPDO:1211/ctxtPLPOD-VGE03").Text = Range("L" & Z + 2) '"H"
        Session.FindById("wnd[0]/usr/subDEFAULTVAL:SAPLCPDO:1211/ctxtPLPOD-LAR03").Text = "L" & Range("B" & Z + 2)
        Session.FindById("wnd[0]/tbar[0]/btn[3]").press
        DoEvents
        If I = 33 Then
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400").VerticalScrollbar.Position = Z
            I = 1
            Z = Z + 1
        Else
            I = I + 1
            Z = Z + 1
        End If
    Wend
'    Session.FindById("wnd[0]/tbar[0]/btn[11]").press
'    Sleep 1000
'    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nca02"
'    Session.FindById("wnd[0]").SendVKey 0
'    Session.FindById("wnd[0]/usr/ctxtRC27M-MATNR").Text = Range("Q1")
'    Session.FindById("wnd[0]/usr/ctxtRC27M-WERKS").Text = "3000"
'    Session.FindById("wnd[0]/usr/ctxtRC271-PLNNR").Text = ""
'    Session.FindById("wnd[0]/usr/ctxtRC271-PLNNR").SetFocus
'    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/tbar[1]/btn[7]").press
    Session.FindById("wnd[0]/usr/tblSAPLCMDITCTRL_1000").GetAbsoluteRow(0).Selected = True
    Session.FindById("wnd[0]/tbar[1]/btn[5]").press
    Session.FindById("wnd[1]/usr/txtRCM01-VORNR").Text = "0010"
    Session.FindById("wnd[1]").SendVKey 0
    Session.FindById("wnd[0]/tbar[0]/btn[11]").press
    AppActivate Application.Caption
    If InStr(Session.FindById("wnd[0]/sbar").Text, "save") = 0 Then MsgBox (Session.FindById("wnd[0]/sbar").Text)
    Call LogOff
End Sub
'Sub SaveMasterRouting(MatNum As String)
'    Dim R_Rows As Integer
'    Dim D_Rows As Long
'    Dim A_Rows As Long
'    Dim MatRow As Long
'    Dim Grouping As String
'    Workbooks.Open Filename:=ThisWorkbook.Path & "\MaterialRoutingDatabase.xlsx"
'    MatRow = WorksheetFunction.Match(MatNum, Workbooks("MaterialRoutingDatabase.xlsx").Sheets("Routing2MatNum").Range("A:A"), 0)
'    Grouping = Workbooks("MaterialRoutingDatabase.xlsx").Sheets("Routing2MatNum").Range("B" & MatRow)
'    R_Rows = WorksheetFunction.CountA(ThisWorkbook.Sheets("MaterialRouting").Range("A:A"))
'    Call SortRoutings("RoutingDatabase")
'    Call SortRoutings("Archive")
'    A_Rows = WorksheetFunction.CountA(Workbooks("MaterialRoutingDatabase.xlsx").Sheets("Archive").Range("A:A"))
'    D_Rows = WorksheetFunction.Match(Grouping, Workbooks("MaterialRoutingDatabase.xlsx").Sheets("RoutingDatabase").Range("A:A"), 0)
'    While Workbooks("MaterialRoutingDatabase.xlsx").Sheets("RoutingDatabase").Range("A" & D_Rows) = Grouping
'        Sheets("Archive").Range("A" & A_Rows + 1 & ":Q" & A_Rows + 1) = Sheets("RoutingDatabase").Range("A" & D_Rows & ":Q" & D_Rows).Value
'        Workbooks("MaterialRoutingDatabase.xlsx").Sheets("RoutingDatabase").Range("A" & D_Rows).EntireRow.Delete
'        DoEvents
'        A_Rows = A_Rows + 1
'    Wend
'    D_Rows = WorksheetFunction.CountA(Workbooks("MaterialRoutingDatabase.xlsx").Sheets("RoutingDatabase").Range("A:A"))
'    Sheets("RoutingDatabase").Range("A" & D_Rows + 1 & ":A" & D_Rows + R_Rows - 1) = Grouping
'    Sheets("RoutingDatabase").Range("B" & D_Rows + 1 & ":E" & D_Rows + R_Rows - 1) = ThisWorkbook.Sheets("MaterialRouting").Range("A2:D" & R_Rows).Value
'    Sheets("RoutingDatabase").Range("H" & D_Rows + 1 & ":M" & D_Rows + R_Rows - 1) = ThisWorkbook.Sheets("MaterialRouting").Range("G2:L" & R_Rows).Value
'    Sheets("RoutingDatabase").Range("P" & D_Rows + 1 & ":P" & D_Rows + R_Rows - 1) = ThisWorkbook.Sheets("MaterialRouting").Range("O2:O" & R_Rows).Value
'    Sheets("RoutingDatabase").Range("Q" & D_Rows + 1 & ":Q" & D_Rows + R_Rows - 1) = Now
'    Workbooks("MaterialRoutingDatabase.xlsx").Close True
'    Call GetRoutingDatabase
'End Sub
Sub SortRoutings(MySheet As String)
    Sheets(MySheet).Select
    Columns("A:Q").Select
    ActiveWorkbook.Worksheets("RoutingDatabase").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("RoutingDatabase").Sort.SortFields.Add Key:=Range( _
        "A2:A1000000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("RoutingDatabase").Sort.SortFields.Add Key:=Range( _
        "B2:B1000000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("RoutingDatabase").Sort
        .SetRange Range("A1:Q1000000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub SaveAllOperations()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Sheets("AllOperations").Copy
    ActiveSheet.Shapes.Range(Array("Rounded Rectangle 1")).Delete
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\AllOperations.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Windows("AllOperations.xlsx").Close True
End Sub
Sub LoadAllOperations()
    Workbooks.Open Filename:=ThisWorkbook.Path & "\AllOperations.xlsx", ReadOnly:=True
    Workbooks("MakeList&RouteMaker.xlsm").Sheets("AllOperations").Range("A1:AY10000") = Workbooks("AllOperations.xlsx").Sheets("AllOperations").Range("A1:AY10000").Value
    Windows("AllOperations.xlsx").Close False
End Sub
Sub RoundedRectangle2_Click()
    If Len(Range("B2")) > 40 Then MsgBox ("Description must be less than 40 characters. Try Again."): Range("B2").Select: End
    If Sheets("Settings").Range("B13") = "Production" Then Call LogonProduction Else Call LogonDevelopment
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmm02"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = Range("A2") '"ISAD20-115"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").GetAbsoluteRow(0).Selected = True
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    Session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB1:SAPLMGD1:1002/txtMAKT-MAKTX").Text = Range("B2") '"RABBIT X-CAT IGG-H+L, IS, ADS"
    Session.FindById("wnd[0]/tbar[0]/btn[11]").press
    Application.Visible = True
    AppActivate Application.Caption
    MsgBox (Session.FindById("wnd[0]/sbar").Text)
    Call LogOff
End Sub
Sub RoundedRectangle4_Click()
    If Sheets("Settings").Range("B13") = "Production" Then Call LogonProduction Else Call LogonDevelopment
    On Error Resume Next
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmigo_tr"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").SetFocus
    Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A08"
    Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").Text = "311"
    Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").SetFocus
    Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").CaretPosition = 3
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").Text = Range("A2") '"ISAD20-115"
    Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-NAME1").Text = "3000"
    Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-LGOBE").Text = Range("D2") '"3003"
    AppActivate Application.Caption
    Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGOITEM-UMLGOBE").Text = InputBox("Enter the new Storage Location", "IS Storage", Range("D2")) '"3003"
    Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/txtGODYNPRO-ERFMG").Text = "1"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-CHARG").Text = Range("C2") '"L100601"
    Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-UMCHA").Text = Range("E2") '"CAT_10"
    Session.FindById("wnd[0]/tbar[0]/btn[11]").press
    Application.Visible = True
    MsgBox (Session.FindById("wnd[0]/sbar").Text)
    Call LogOff
End Sub
Sub RoundedRectangle6_Click()
    If Range("E2") = "" Then MsgBox ("Must enter a batch number. Try Again."): Range("E2").Select: End
    If Range("A2") = "" Then MsgBox ("Must enter a Material number. Try Again."): Range("A2").Select: End
    If Len(Range("E2")) > 10 Then MsgBox ("Batch must be less than 10 characters. Try Again."): Range("E2").Select: End
    If Sheets("Settings").Range("B13") = "Production" Then Call LogonProduction Else Call LogonDevelopment
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nco01"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-MATNR").Text = Range("A2") '"ISAD20-115"
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-WERKS").Text = "3000"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/txtCAUFVD-GAMNG").Text = "1"
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/ctxtCAUFVD-GSTRP").Text = Date '"7/30/2014"
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/cmbCAUFVD-TERKZ").Key = "1"
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOWE").Select
    Dim MyText As String
    On Error Resume Next
    MyText = Session.FindById("wnd[1]/usr/txtSPOP-DIAGNOSE").Text
    AppActivate Application.Caption
    If InStr(MyText, "o routing") <> 0 Then Session.FindById("wnd[1]/usr/btnSPOP-VAROPTION3").press: MsgBox (MyText): End
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOWE/ssubSUBSCR_0115:SAPLCOKO1:0190/ctxtAFPOD-CHARG").Text = Range("E2") '"Cat_140730"
    Session.FindById("wnd[0]/tbar[0]/btn[11]").press
    Session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").press
    Session.FindById("wnd[0]/tbar[0]/btn[11]").press
    If InStr(Session.FindById("wnd[0]/sbar").Text, "saved") = 0 Then
        Session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").press
        Session.FindById("wnd[0]").SendVKey 0
        Session.FindById("wnd[0]/tbar[0]/btn[11]").press
    End If
    AppActivate Application.Caption
    MsgBox (Session.FindById("wnd[0]/sbar").Text)
    Call LogOff
End Sub
Function GetStepRange(MySearchVal As String)
    Dim MyRow As Long
    MyRow = WorksheetFunction.Match(MySearchVal, Sheets("AllOperations").Range("A:A"), 0)
    GetStepRange = "B" & MyRow & ":" & Chr(64 + Sheets("AllOperations").Cells(MyRow, Sheets("AllOperations").Columns.Count).End(xlToLeft).Column) & MyRow
End Function
Sub CreateEditIS()
    Range("A1000000").End(xlUp).Select
    If ActiveCell.Row <> 1 Then
        Sheets("AddImmunosorbent").Range("A2:D2") = Range("A" & ActiveCell.Row & ":D" & ActiveCell.Row).Value
        Sheets("AddImmunosorbent").Range("E2") = ""
    Else
        Sheets("AddImmunosorbent").Range("A2:E2") = ""
    End If
    Sheets("AddImmunosorbent").Select
End Sub
Sub LookupIS()
    Application.ScreenUpdating = False
    Sheets("Settings").Select
    Columns("X:Y").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range("AI1:AI2"), CopyToRange:=Range("AJ1:AK1"), Unique:=False
    Sheets("MaterialRouting").Range("AD1:AE100000") = Sheets("Settings").Range("AJ1:AK100000").Value
    Sheets("MaterialRouting").Select
End Sub
Sub LookupISNum()
    Application.ScreenUpdating = False
    Sheets("Settings").Select
    Columns("X:Y").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range("AL1:AL2"), CopyToRange:=Range("AJ1:AK1"), Unique:=False
    Sheets("MaterialRouting").Range("AD1:AE100000") = Sheets("Settings").Range("AJ1:AK100000").Value
    Sheets("MaterialRouting").Select
End Sub
Sub LookupLotInfo()
    Dim MyMaterial As String
    If ActiveCell.Column = 14 Then
        MyMaterial = Range("N" & ActiveCell.Row)
    Else
        MyMaterial = Range("AD" & ActiveCell.Row)
    End If
    If MyMaterial = "" Then MsgBox ("Select a cell with an IS before clicking this button."): End
    Sheets("AllMaterialStock").Select
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$J$1000000").AutoFilter Field:=1, Criteria1:="=" & MyMaterial, Operator:=xlAnd
End Sub
Sub AddOperations()
        Dim MyRange As String
        Dim MyVal As String
        Dim MySearchVal As String
        Sheets("MaterialRouting").Select
        Range("D:D").Validation.Delete
        Range("D2").Select
        While ActiveCell.Row < 100
            On Error Resume Next
            MyVal = ""
            MySearchVal = WorksheetFunction.Proper(Range("D" & ActiveCell.Row - 1))
            MyVal = WorksheetFunction.VLookup(MySearchVal, Sheets("AllOperations").Range("A:B"), 2, False)
            If MyVal = "" Then
                MySearchVal = WorksheetFunction.UCase(Range("D" & ActiveCell.Row - 1))
                MyVal = WorksheetFunction.VLookup(MySearchVal, Sheets("AllOperations").Range("A:B"), 2, False)
            End If
            If MyVal = "" Then
                MySearchVal = WorksheetFunction.LCase(Range("D" & ActiveCell.Row - 1))
                MyVal = WorksheetFunction.VLookup(MySearchVal, Sheets("AllOperations").Range("A:B"), 2, False)
            End If
            If MyVal = "" Then
                MySearchVal = Range("D" & ActiveCell.Row - 1)
                MyVal = WorksheetFunction.VLookup(MySearchVal, Sheets("AllOperations").Range("A:B"), 2, False)
            End If
            On Error GoTo 0
            If MyVal <> "" Then
                MyRange = "=indirect(""AllOperations!" & GetStepRange(MySearchVal) & """" & ")"
            Else
                MyRange = "=indirect(""AllOperations!A:A" & """" & ")"
            End If
            Range("D" & ActiveCell.Row).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=MyRange
            DoEvents
            ActiveCell.Offset(1, 0).Select
        Wend
        Range("D2").Select
End Sub
Sub ChangeIS()
    Dim Rng As Range
    Dim Cell As Range
    Dim MyIS As String
    Dim MyCnt As Integer
    Set Rng = Application.Selection
    MyCnt = Rng.Count
    If MyCnt <> 2 Then MsgBox ("Select one cell in Column N and one from Column AD and then hit this button to change the IS for the Operation step."): End
    For Each Cell In Rng
        If Cell.Row = 1 Then Cell.Select: MsgBox ("Select one cell in Column N and one from Column AD (and not Row 1) and then hit this button to change the IS for the Operation step."): End
    Next
    For Each Cell In Rng
        If Cell.Column = 30 Then MyIS = Trim(Cell.Value)
    Next
    For Each Cell In Rng
        If Cell.Column = 14 Then Cell.Value = Trim(MyIS)
    Next
End Sub
Sub GotoSettings()
    If Sheets("Settings").Range("B13") = "Production" Then Sheets("Settings").Range("B13") = "Development" Else Sheets("Settings").Range("B13") = "Production"
End Sub
Function GetPlannedOrder(MatNum As String)
    Dim MyCurSheet As String
    Dim MyMatQty As Long
    On Error Resume Next
    MyCurSheet = ActiveSheet.Name
    Sheets("OrderHeaders").Range("S2:V100000") = ""
    ProdSup = ""
    SumQtyOfOrders = 0
    CountOfPlannedOrder = 0
    Sheets("OrderHeaders").Range("R1") = Sheets("OrderHeaders").Range("F1").Value
    Sheets("OrderHeaders").Range("S1") = Sheets("OrderHeaders").Range("A1").Value
    Sheets("OrderHeaders").Range("T1") = Sheets("OrderHeaders").Range("B1").Value
    Sheets("OrderHeaders").Range("U1") = Sheets("OrderHeaders").Range("D1").Value
    Sheets("OrderHeaders").Range("U2") = "<>*P01"
    Sheets("OrderHeaders").Range("T2") = MatNum
    Sheets("OrderHeaders").Select
    Sheets("OrderHeaders").Range("A:P").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range("T1:U2"), CopyToRange:=Range("R1:S1"), Unique:=False
    CountOfPlannedOrder = WorksheetFunction.CountA(Sheets("OrderHeaders").Range("S:S")) - 1
    MyMatQty = WorksheetFunction.VLookup(Sheets("OrderHeaders").Range("S2"), Sheets("OrderHeaders").Range("A:H"), 8, False)
    SumQtyOfOrders = SumQtyOfOrders + (CountOfPlannedOrder * MyMatQty)
    ProdSup = Sheets("OrderHeaders").Range("R2")
    Sheets(MyCurSheet).Select
    GetPlannedOrder = Sheets("OrderHeaders").Range("S2")
End Function
Function GetProductionOrder(MatNum As String)
    Dim MyCurSheet As String
    On Error Resume Next
    ProdSup = ""
    SumQtyOfOrders = ""
    MyCurSheet = ActiveSheet.Name
    Sheets("OrderHeaders").Select
    Sheets("OrderHeaders").Range("S2:V100000") = ""
    Sheets("OrderHeaders").Range("S1") = Sheets("OrderHeaders").Range("A1").Value
    Sheets("OrderHeaders").Range("T1") = Sheets("OrderHeaders").Range("B1").Value
    Sheets("OrderHeaders").Range("U1") = Sheets("OrderHeaders").Range("D1").Value
    Sheets("OrderHeaders").Range("R1") = Sheets("OrderHeaders").Range("F1").Value
    Sheets("OrderHeaders").Range("U2") = "*P01"
    Sheets("OrderHeaders").Range("T2") = MatNum
    Sheets("OrderHeaders").Range("A:P").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range("T1:U2"), CopyToRange:=Range("R1:S1"), Unique:=False
    ProdSup = Sheets("OrderHeaders").Range("R2")
    SumQtyOfOrders = WorksheetFunction.VLookup(Sheets("OrderHeaders").Range("S2"), Sheets("OrderHeaders").Range("A:H"), 8, False)
    Sheets(MyCurSheet).Select
    GetProductionOrder = Sheets("OrderHeaders").Range("S2")
End Function
Sub MakeBOM()
    Dim MyMat As String
    Range("A2:J100000").Interior.Pattern = xlNone
    MyMat = Trim(Range("B" & ActiveCell.Row))
    Range("A" & ActiveCell.Row & ":J" & ActiveCell.Row).Interior.Color = RGB(175, 120, 215)
    Sheets("MakeTree").Select
    Range("A2") = MyMat
    Range("D2:Y2") = ""
    Range("A4:AI1000") = ""
    Call MakeBomTree
End Sub
Sub GotoMakeList()
    Sheets("ModifyBOM").Range("A4:L1000") = ""
    Sheets("MaterialRouting").Range("A2:O1000") = ""
    Sheets("MakeList").Select
End Sub
Sub GetMaterialBatchCount(MyMatNum As String)
    'Sheets("AllMaterialStock").Select
    Sheets("AllMaterialStock").Range("R2,R4,R6") = "'=" & MyMatNum
    Sheets("AllMaterialStock").Range("R1,R3,R5") = Sheets("AllMaterialStock").Range("A1")
    Sheets("AllMaterialStock").Range("S1") = Sheets("AllMaterialStock").Range("F1")
    Sheets("AllMaterialStock").Range("S3") = Sheets("AllMaterialStock").Range("G1")
    Sheets("AllMaterialStock").Range("S5") = Sheets("AllMaterialStock").Range("H1")
    Sheets("AllMaterialStock").Range("S2,S4,S6").FormulaR1C1 = ">0"
    Sheets("AllMaterialStock").Range("T2,T4,T6").FormulaR1C1 = "=DCOUNTA(C1:C10,R[-1]C[-1],R[-1]C[-2]:RC[-1])"
    UnRStock = Sheets("AllMaterialStock").Range("T2")
    RStock = Sheets("AllMaterialStock").Range("T4")
    BStock = Sheets("AllMaterialStock").Range("T6")
End Sub
Sub MakeBomTree()
    Dim Lev(99), Item(99), PPh(99), Object(99), Component(99), ObjDec(99), I As Integer, MaxI As Integer, MaxPPh As Integer, MaxLev As Integer
    Dim Qty(99) As String
    Dim X As Integer, Y As Integer
    Dim MyMat As String
    Dim BOM As Long
    Dim MyCol As Integer
    Dim Tree As Boolean
    Dim MyRGB As String
    Dim RGBArray() As String
    Dim MyRed As Integer, MyGreen As Integer, MyBlue As Integer
    Dim MyStartCell As String
    Application.ScreenUpdating = False
    On Error Resume Next
    Sheets("MakeTree").Select
    Sheets("MakeTree").Range("D2:AJ2") = ""
    Sheets("MakeTree").Range("A3:AJ1000") = ""
    Sheets("MakeTree").Range("A4:AJ1000").Interior.Color = xlNone
    Sheets("MakeTree").Range("A1") = "MatNum"
    Sheets("MakeTree").Range("D1") = "Target Qty"
    Sheets("MakeTree").Range("J1") = "Unrestricted Inventory" & vbLf & "[# Batches]"
    Sheets("MakeTree").Range("N1") = "Restricted Inventory" & vbLf & "[# Batches]"
    Sheets("MakeTree").Range("O1") = "Blocked Inventory" & vbLf & "[# Batches]"
    Sheets("MakeTree").Range("P1") = "Base Qty"
    Sheets("MakeTree").Range("Q1") = "Base Units"
    Sheets("MakeTree").Range("R1") = "Production Supervisor"
    Sheets("MakeTree").Range("S1") = "Planned Order"
    Sheets("MakeTree").Range("T1") = "Production Order"
    Sheets("MakeTree").Range("U1") = "# Planned Orders"
    Sheets("MakeTree").Range("V1") = "Sum Qty Orders"
    Sheets("MakeTree").Range("W1") = "Safety Stock"
    Sheets("MakeTree").Range("X1") = "Order Start"
    Sheets("MakeTree").Range("Y1") = "Est. Order Finish"
    Sheets("MakeTree").Range("AA1") = "BOM"
    ActiveSheet.Shapes.Range(Array("BOMTree")).Delete
    MyMat = Trim(Range("A2"))
    If MyMat = "" Then End
    BOM = WorksheetFunction.VLookup(Range("A2"), Sheets("MAST").Range("A:B"), 2, False)
    Sheets("STPO").Range("J2") = BOM
    Range("J2") = WorksheetFunction.VLookup(MyMat, Sheets("AllMaterialStock").Range("A:J"), 6, False)
    If Range("J2") = "" Then Range("J2") = 0
    Range("N2") = WorksheetFunction.VLookup(MyMat, Sheets("AllMaterialStock").Range("A:J"), 7, False)
    If Range("N2") = "" Then Range("N2") = 0
    Range("O2") = WorksheetFunction.VLookup(MyMat, Sheets("AllMaterialStock").Range("A:J"), 8, False)
    If Range("O2") = "" Then Range("O2") = 0
    Call GetMaterialBatchCount(MyMat)
    Range("J2") = Range("J2") & " [" & UnRStock & "]"
    Range("N2") = Range("N2") & " [" & RStock & "]"
    Range("O2") = Range("O2") & " [" & BStock & "]"
    Range("P2") = WorksheetFunction.VLookup(BOM, Sheets("STKO").Range("A:C"), 3, False)
    Range("Q2") = WorksheetFunction.VLookup(BOM, Sheets("STKO").Range("A:C"), 2, False)
    Range("T2") = GetProductionOrder(MyMat)
    If Range("T2") = "" And MyMat <> "" Then Range("S2") = GetPlannedOrder(MyMat)
    If Range("S2") <> "" And ProdSup <> "" Then Range("R2") = ProdSup
    If Range("T2") <> "" And ProdSup <> "" Then Range("R2") = ProdSup
    Range("U2") = CountOfPlannedOrder
    Range("V2") = SumQtyOfOrders
    Range("D2") = SumQtyOfOrders
    Range("W2") = WorksheetFunction.VLookup(MyMat, Sheets("MARC").Range("A:E"), 4, False)
    If Range("T2") <> "" Then
        Range("X2").FormulaR1C1 = "=VLOOKUP(TEXT(RC[-4],""0""),OrderHeaders!C1:C16,10,FALSE)"
        Range("X2") = Range("X2").Value
        Range("Y2").FormulaR1C1 = "=VLOOKUP(TEXT(RC[-5],""0""),OrderHeaders!C1:C16,11,FALSE)"
        Range("Y2") = Range("Y2").Value
    End If
    Range("AA2") = BOM
    Call GetBOMForMaterial(MyMat)
    If Range("B4") <> "No BOM" Then
        I = 0
        Sheets("BOMData").Select
        Range("A2").Select
        While Range("A" & ActiveCell.Row) <> ""
            MaxI = I
            Lev(I) = Range("L" & ActiveCell.Row)
            If Lev(I) >= MaxLev Then MaxLev = Lev(I)
            PPh(I) = Range("M" & ActiveCell.Row)
            Qty(I) = Range("F" & ActiveCell.Row)
            Item(I) = Range("B" & ActiveCell.Row)
            If PPh(I) >= MaxPPh Then MaxPPh = PPh(I)
            Object(I) = Range("E" & ActiveCell.Row)
            ObjDec(I) = Range("H" & ActiveCell.Row)
            ActiveCell.Offset(1, 0).Select
            I = I + 1
            DoEvents
        Wend
        Sheets("MakeTree").Select
        Range("B4").Select
        If Sheets("Settings").Range("B13") = "Production" Then Call LogonProduction Else Call LogonDevelopment
        Range("L2") = GetFixedLotSize(Range("A2"))
        Call LogOff
        For I = 0 To MaxI
            If Lev(I) = 1 And Item(I) <> "" And Object(I) <> "" And InStr(Qty(I), "-") = 0 Then
                ActiveCell.Offset(0, 1) = Object(I) & "_" & ObjDec(I)
                ActiveCell.Offset(0, 0) = "|----"
                MyCol = ActiveCell.Column
                MyRGB = WorksheetFunction.VLookup(MyCol, Sheets("Settings").Range("A2:C9"), 3, False)
                RGBArray = Split(MyRGB, ",", -1)
                Range(Cells(ActiveCell.Row, MyCol), Cells(ActiveCell.Row, 35)).Interior.Color = RGB(Val(RGBArray(0)), Val(RGBArray(1)), Val(RGBArray(2)))
                Range("T" & ActiveCell.Row) = Object(I)
                Range("V" & ActiveCell.Row) = WorksheetFunction.SumIf(Sheets("AllMaterialStock").Range("A:A"), Range("T" & ActiveCell.Row), Sheets("AllMaterialStock").Range("F:F"))
                Range("W" & ActiveCell.Row) = WorksheetFunction.SumIf(Sheets("AllMaterialStock").Range("A:A"), Range("T" & ActiveCell.Row), Sheets("AllMaterialStock").Range("G:G"))
                'If Range("W" & ActiveCell.Row) <> 0 Then Range("W" & ActiveCell.Row) = Range("W" & ActiveCell.Row) & " [" & WorksheetFunction.CountIf(Sheets("AllMaterialStock").Range("A:A"), Range("T" & ActiveCell.Row)) & "]"
                Range("X" & ActiveCell.Row) = WorksheetFunction.SumIf(Sheets("AllMaterialStock").Range("A:A"), Range("T" & ActiveCell.Row), Sheets("AllMaterialStock").Range("H:H"))
                Call GetMaterialBatchCount(Range("T" & ActiveCell.Row))
                Range("V" & ActiveCell.Row) = Range("V" & ActiveCell.Row) & " [" & UnRStock & "]"
                Range("W" & ActiveCell.Row) = Range("W" & ActiveCell.Row) & " [" & RStock & "]"
                Range("X" & ActiveCell.Row) = Range("X" & ActiveCell.Row) & " [" & BStock & "]"
                Range("AD" & ActiveCell.Row) = GetProductionOrder((Object(I)))
                If Range("AD" & ActiveCell.Row) = "" And Object(I) <> "" Then Range("AC" & ActiveCell.Row) = GetPlannedOrder((Object(I)))
                If Range("AC" & ActiveCell.Row) <> "" And ProdSup <> "" Then Range("AB" & ActiveCell.Row) = ProdSup
                If Range("AD" & ActiveCell.Row) <> "" And ProdSup <> "" Then Range("AB" & ActiveCell.Row) = ProdSup
                Range("AE" & ActiveCell.Row) = CountOfPlannedOrder
                Range("AF" & ActiveCell.Row) = SumQtyOfOrders
                Range("U" & ActiveCell.Row) = SumQtyOfOrders
                Range("AG" & ActiveCell.Row) = WorksheetFunction.VLookup(Object(I), Sheets("MARC").Range("A:E"), 4, False)
                If Range("AD" & ActiveCell.Row) <> "" Then
                    Range("AH" & ActiveCell.Row).FormulaR1C1 = "=VLOOKUP(TEXT(RC[-4],""0""),OrderHeaders!C1:C16,10,FALSE)"
                    Range("AH" & ActiveCell.Row) = Range("AH" & ActiveCell.Row).Value
                    Range("AI" & ActiveCell.Row).FormulaR1C1 = "=VLOOKUP(TEXT(RC[-5],""0""),OrderHeaders!C1:C16,11,FALSE)"
                    Range("AI" & ActiveCell.Row) = Range("AI" & ActiveCell.Row).Value
                End If
                Sheets("STPO").Range("K2") = Object(I)
                Range("Y" & ActiveCell.Row) = WorksheetFunction.DGet(Sheets("STPO").Range("A:D"), "Quantity", Sheets("STPO").Range("J1:K2"))
                Range("Z" & ActiveCell.Row) = WorksheetFunction.DGet(Sheets("STPO").Range("A:D"), "Un", Sheets("STPO").Range("J1:K2"))
                BOM = WorksheetFunction.VLookup(Object(I), Sheets("MAST").Range("A:B"), 2, False)
                Range("AA" & ActiveCell.Row) = BOM
                Sheets("STPO").Range("J2") = BOM
                ActiveCell.Offset(1, 0).Select
            Else
                If Item(I) = "" Then
                    MyStartCell = ActiveCell.Address
                    Range("A4:Z99").Select
                    If Object(I) <> "" Then
                        Selection.Find(What:=Object(I), After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
                        If ActiveCell <> "" Then
                            ActiveCell.Offset(1, 0).Select
                            Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
                            Range(Cells(ActiveCell.Row, 2), Cells(ActiveCell.Row, 35)).Interior.Color = xlNone
                        Else
                            Range(MyStartCell).Select
                        End If
                    End If
                Else
                    If Object(I) <> "" And InStr(Qty(I), "-") = 0 Then
                        ActiveCell.Offset(0, 1) = Object(I) & "_" & ObjDec(I)
                        ActiveCell.Offset(0, 0) = "|----"
                        MyCol = ActiveCell.Column
                        MyRGB = WorksheetFunction.VLookup(MyCol, Sheets("Settings").Range("A2:C9"), 3, False)
                        RGBArray = Split(MyRGB, ",", -1)
                        Range(Cells(ActiveCell.Row, MyCol), Cells(ActiveCell.Row, 35)).Interior.Color = RGB(Val(RGBArray(0)), Val(RGBArray(1)), Val(RGBArray(2)))
                        Range("T" & ActiveCell.Row) = Object(I)
                        Range("V" & ActiveCell.Row) = WorksheetFunction.SumIf(Sheets("AllMaterialStock").Range("A:A"), Range("T" & ActiveCell.Row), Sheets("AllMaterialStock").Range("F:F"))
                        Range("W" & ActiveCell.Row) = WorksheetFunction.SumIf(Sheets("AllMaterialStock").Range("A:A"), Range("T" & ActiveCell.Row), Sheets("AllMaterialStock").Range("G:G"))
                        'If Range("W" & ActiveCell.Row) <> 0 Then Range("W" & ActiveCell.Row) = Range("W" & ActiveCell.Row) & " [" & WorksheetFunction.CountIf(Sheets("AllMaterialStock").Range("A:A"), Range("T" & ActiveCell.Row)) & "]"
                        Range("X" & ActiveCell.Row) = WorksheetFunction.SumIf(Sheets("AllMaterialStock").Range("A:A"), Range("T" & ActiveCell.Row), Sheets("AllMaterialStock").Range("H:H"))
                        Call GetMaterialBatchCount(Range("T" & ActiveCell.Row))
                        Range("V" & ActiveCell.Row) = Range("V" & ActiveCell.Row) & " [" & UnRStock & "]"
                        Range("W" & ActiveCell.Row) = Range("W" & ActiveCell.Row) & " [" & RStock & "]"
                        Range("X" & ActiveCell.Row) = Range("X" & ActiveCell.Row) & " [" & BStock & "]"
                        Range("AD" & ActiveCell.Row) = GetProductionOrder((Object(I)))
                        If Range("AD" & ActiveCell.Row) = "" Then Range("AC" & ActiveCell.Row) = GetPlannedOrder((Object(I)))
                        Range("AE" & ActiveCell.Row) = CountOfPlannedOrder
                        Range("AF" & ActiveCell.Row) = SumQtyOfOrders
                        Range("U" & ActiveCell.Row) = SumQtyOfOrders
                        Range("AG" & ActiveCell.Row) = WorksheetFunction.VLookup(Object(I), Sheets("MARC").Range("A:E"), 4, False)
                        If Range("AD" & ActiveCell.Row) <> "" Then
                            Range("AH" & ActiveCell.Row).FormulaR1C1 = "=VLOOKUP(TEXT(RC[-4],""0""),OrderHeaders!C1:C16,10,FALSE)"
                            Range("AH" & ActiveCell.Row) = Range("AH" & ActiveCell.Row).Value
                            Range("AI" & ActiveCell.Row).FormulaR1C1 = "=VLOOKUP(TEXT(RC[-5],""0""),OrderHeaders!C1:C16,11,FALSE)"
                            Range("AI" & ActiveCell.Row) = Range("AI" & ActiveCell.Row).Value
                        End If
                        If Range("AC" & ActiveCell.Row) <> "" And ProdSup <> "" Then Range("AB" & ActiveCell.Row) = ProdSup
                        If Range("AD" & ActiveCell.Row) <> "" And ProdSup <> "" Then Range("AB" & ActiveCell.Row) = ProdSup
                        Sheets("STPO").Range("K2") = Object(I)
                        Sheets("STPO").Range("J2") = Range("AA" & ActiveCell.Row - 1)
                        Range("Y" & ActiveCell.Row) = WorksheetFunction.DGet(Sheets("STPO").Range("A:D"), "Quantity", Sheets("STPO").Range("J1:K2"))
                        Range("Z" & ActiveCell.Row) = WorksheetFunction.DGet(Sheets("STPO").Range("A:D"), "Un", Sheets("STPO").Range("J1:K2"))
                        BOM = WorksheetFunction.VLookup(Object(I), Sheets("MAST").Range("A:B"), 2, False)
                        Range("AA" & ActiveCell.Row) = BOM
                        Sheets("STPO").Range("J2") = BOM
                        ActiveCell.Offset(1, 0).Select
                        Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
                        Range(Cells(ActiveCell.Row, 2), Cells(ActiveCell.Row, 34)).Interior.Color = xlNone
                    End If
                End If
            End If
        Next I
        Tree = False
        MyCol = ActiveCell.Column
        For I = MyCol To 2 Step -1
            Cells(99, I).Select
            While ActiveCell.Row <> 2
                If ActiveCell = "|----" Then
                    If ActiveCell.Offset(-1, -1) <> "" Then
                        Tree = False
                        ActiveCell.Offset(-2, 0).Select
                    Else
                        Tree = True
                        ActiveCell.Offset(-1, 0).Select
                    End If
                End If
                If Tree = False And ActiveCell = "" Then
                    ActiveCell = ""
                Else
                    If ActiveCell = "" Then
                        If ActiveCell.Row <> 3 Then ActiveCell = "|"
                        If ActiveCell.Row = 3 Then
                            Tree = False
                            If ActiveCell.Column = 2 Then ActiveCell = "|" & vbLf & "|" & vbLf & "|"
                        End If
                    End If
                End If
                ActiveCell.Offset(-1, 0).Select
                DoEvents
            Wend
        Next I
        Application.CutCopyMode = False
        Range("T3:AI99").Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        Range("T3") = "MatNum"
        Range("U3") = "Target Qty"
        Range("V3") = "Unrestricted Inventory" & vbLf & "[# Batches]"
        Range("W3") = "Restricted Inventory" & vbLf & "[# Batches]"
        Range("X3") = "Blocked Inventory" & vbLf & "[# Batches]"
        Range("Y3") = "Component Qty"
        Range("Z3") = "Component Units"
        Range("AA3") = "BOM"
        Range("AB3") = "Production Supervisor"
        Range("AC3") = "Planned Order"
        Range("AD3") = "Production Order"
        Range("AE3") = "# Planned Orders"
        Range("AF3") = "Sum Qty Orders"
        Range("AG3") = "Safety Stock"
        Range("AH3") = "Order Start"
        Range("AI3") = "Est. Order Finish"
        Columns("U:AG").EntireColumn.AutoFit
        Call GetMaterialRouting(MyMat)
        Range("AA1").ColumnWidth = 0
    End If
    If Sheets("MaterialRouting").Range("A2") = "" Then
        Stop
        'missing routing need to verify the code below
        Call CA01_CreateRouting(Range("A2"), WorksheetFunction.VLookup(Range("A2"), Sheets("MARA").Range("A:F"), 6, False), WorksheetFunction.VLookup(Range("A2"), Sheets("MARA").Range("A:F"), 4, False))
    End If
    Sheets("MakeTree").AutoFilterMode = False
End Sub
Sub CA01_CreateRouting(MatNum As String, MatDescription As String, Units As String)
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nca01"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ctxtRC27M-MATNR").Text = MatNum '"XA80-108"
    Session.FindById("wnd[0]/usr/ctxtRC27M-MATNR").CaretPosition = 8
    Session.FindById("wnd[0]/tbar[1]/btn[5]").press
    Session.FindById("wnd[1]/usr/sub:SAPLCPCO:0101/radTYP[1,0]").Select
    Session.FindById("wnd[1]/usr/sub:SAPLCPCO:0101/radTYP[1,0]").SetFocus
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    Session.FindById("wnd[1]/usr/ctxtRC271-PLNNR").Text = "default"
    Session.FindById("wnd[1]/usr/ctxtRC271-WERKS").Text = "3000"
    Session.FindById("wnd[1]/usr/ctxtRC271-WERKS").SetFocus
    Session.FindById("wnd[1]/usr/ctxtRC271-WERKS").CaretPosition = 4
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    Session.FindById("wnd[0]/usr/subGENERAL:SAPLCPDA:1210/txtPLKOD-KTEXT").Text = MatDescription '"[ ] GOAT X-HUMAN IGE"
    Session.FindById("wnd[0]/usr/subGENERALVW:SAPLCPDA:1211/ctxtPLKOD-STATU").Text = "4"
    Session.FindById("wnd[0]/usr/subGENERALVW:SAPLCPDA:1211/ctxtPLKOD-PLNME").Text = Units '"ML"
    Session.FindById("wnd[0]").SendVKey 0
    Application.Visible = True
    MsgBox (Session.FindById("wnd[0]/tbar[0]/btn[11]").Text)
End Sub
Sub GetMakeList()
    Sheets("MakeList").Select
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Range("A2:J100000") = ""
    Range("A2:J100000").Interior.Pattern = xlNone
    'Sheets("MakeList").Range("B1") = "Material"
    'Sheets("MakeList").Range("B1") = Sheets("Components").Range("P1").Value
    Sheets("OrderHeaders").Range("A1:P1000000").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Sheets("MakeList").Range("K1:M2"), CopyToRange:=Sheets("MakeList").Range("B1:C1"), Unique:=True
    'Sheets("MakeList").Range("B1") = "Material"
    If WorksheetFunction.CountA(Range("B:B")) > 1 Then
        Range("A2:A" & WorksheetFunction.CountA(Range("B:B"))).FormulaR1C1 = "=IFERROR(if(VLOOKUP(RC[1],ZVIALORDERS!C1,1,FALSE)=RC[1],-1),IF(SUM(RC[4]:RC[6])=0,0,iferror(VALUE(TEXT(RC[4]/RC[7],""0.0"")),"""")))"
'        Range("A2:A" & WorksheetFunction.CountA(Range("B:B"))).FormulaR1C1 = "=IFERROR(if(VLOOKUP(RC[1],ZVIALORDERS!C1,1,FALSE)=RC[1],-1),IF(SUM(RC[4]:RC[6])=0,0,VALUE(TEXT(RC[4]/MakeList!RC[7],""0.0""))))"
        Range("D2:D" & WorksheetFunction.CountA(Range("B:B"))).FormulaR1C1 = "=IF(RC[-3]=-1,""Backorder"",RC[-3]*100 & ""% Safety Stock"")"
        Range("E2:E" & WorksheetFunction.CountA(Range("B:B"))).FormulaR1C1 = "=SUMIF(AllMaterialStock!C1,RC2,AllMaterialStock!C[1])"
        Range("F2:F" & WorksheetFunction.CountA(Range("B:B"))).FormulaR1C1 = "=SUMIF(AllMaterialStock!C1,RC2,AllMaterialStock!C[1])"
        Range("G2:G" & WorksheetFunction.CountA(Range("B:B"))).FormulaR1C1 = "=SUMIF(AllMaterialStock!C1,RC2,AllMaterialStock!C[1])"
        Range("H2:H" & WorksheetFunction.CountA(Range("B:B"))).FormulaR1C1 = "=VLOOKUP(RC[-6],MARC!C1:C4,4,FALSE)"
        Range("J2:J" & WorksheetFunction.CountA(Range("B:B"))).FormulaR1C1 = "=VLOOKUP(RC[-8],MARA!C1:C6,4,FALSE)"
        Range("I2:I" & WorksheetFunction.CountA(Range("B:B"))).FormulaR1C1 = "=iferror(INDIRECT(""MCSI!C"" & MATCH(RC[-7],MCSI!C1,0)),"""")"
        Range("A2:A" & WorksheetFunction.CountA(Range("B:B"))) = Range("A2:A" & WorksheetFunction.CountA(Range("B:B"))).Value
        Range("D2:D" & WorksheetFunction.CountA(Range("B:B"))) = Range("D2:D" & WorksheetFunction.CountA(Range("B:B"))).Value
        Range("E2:E" & WorksheetFunction.CountA(Range("B:B"))) = Range("E2:E" & WorksheetFunction.CountA(Range("B:B"))).Value
        Range("F2:F" & WorksheetFunction.CountA(Range("B:B"))) = Range("F2:F" & WorksheetFunction.CountA(Range("B:B"))).Value
        Range("G2:G" & WorksheetFunction.CountA(Range("B:B"))) = Range("G2:G" & WorksheetFunction.CountA(Range("B:B"))).Value
        Range("H2:H" & WorksheetFunction.CountA(Range("B:B"))) = Range("H2:H" & WorksheetFunction.CountA(Range("B:B"))).Value
        Range("I2:I" & WorksheetFunction.CountA(Range("B:B"))) = Range("I2:I" & WorksheetFunction.CountA(Range("B:B"))).Value
        Range("J2:J" & WorksheetFunction.CountA(Range("B:B"))) = Range("J2:J" & WorksheetFunction.CountA(Range("B:B"))).Value
    End If
    Range("A1") = "Score"
    Range("B1") = "Material"
    Range("C1") = "Material description"
    Range("D1") = "Status"
    Range("E1") = "Unrestricted"
    Range("F1") = "Restricted"
    Range("G1") = "Blocked"
    Range("H1") = "Safety Stock"
    Range("I1") = "Yearly Units Sold"
    Range("J1") = "BUn"
    Columns("A:J").Select
    ActiveWorkbook.Worksheets("MakeList").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MakeList").Sort.SortFields.Add Key:=Range("A2:A100000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("MakeList").Sort.SortFields.Add Key:=Range("H2:H100000"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("MakeList").Sort
        .SetRange Range("A1:J100000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A:J").Columns.AutoFit
    ActiveCell.Select
    Application.EnableEvents = True
End Sub
Sub SortByCurCol()
    Dim CurCol As String
    CurCol = ActiveCell.Address
    CurCol = Left(Replace(CurCol, "$", ""), 1)
    Columns("A:J").Select
    ActiveWorkbook.Worksheets("MakeList").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MakeList").Sort.SortFields.Add Key:=Range(CurCol & "2:" & CurCol & "100000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("MakeList").Sort
        .SetRange Range("A1:J100000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub SelectFieldFromSE16n(FieldName As String)
    Session.FindById("wnd[0]/tbar[0]/btn[71]").press
    Session.FindById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = FieldName '"maktx"
    Session.FindById("wnd[1]").SendVKey 0
    Session.FindById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
End Sub
Sub GetAllFERTs()
    Dim MySheet As String
    Dim MyUnits As String
    Dim MyUFreq As Integer
    Dim MyRow As Long
    Application.DisplayAlerts = False
    Sheets("MARA").Select
    MySheet = ActiveSheet.Name
    MyRow = WorksheetFunction.Match(MySheet, Sheets("Settings").Range("S:S"), 0)
    MyUnits = Sheets("Settings").Range("U" & MyRow)
    MyUFreq = Sheets("Settings").Range("T" & MyRow)
    'if getmoddate is more than 1 week old then run the code below otherwise call getcurrentsheet
    If UCase(Sheets("Settings").Range("G1")) = "FALSE" Then
        Sheets("MARA").Select
        Call GetCurrentSheet
    Else
        Cells = ""
        If Sheets("Settings").Range("B13") = "Production" Then Call LogonProduction Else Call LogonDevelopment
        Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nse16n"
        Session.FindById("wnd[0]").SendVKey 0
        Session.FindById("wnd[0]/usr/ctxtGD-TAB").Text = "mara"
        Session.FindById("wnd[0]/usr/txtGD-MAX_LINES").Text = ""
        Session.FindById("wnd[0]").SendVKey 0
        Session.FindById("wnd[0]").SendVKey 18
        Call SelectFieldFromSE16n("matnr")
        Call SelectFieldFromSE16n("mtart")
        Call SelectFieldFromSE16n("matkl")
        Call SelectFieldFromSE16n("meins")
        Call SelectFieldFromSE16n("normt")
        Call SelectFieldFromSE16n("maktx")
        Session.FindById("wnd[0]").SendVKey 8
        Session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
        Session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").SelectContextMenuItem "&PC"
        Session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
        Session.FindById("wnd[1]/tbar[0]/btn[0]").press
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "#"
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
            :="#", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
        ActiveSheet.PasteSpecial Format:="Unicode Text", Link:=False, _
            DisplayAsIcon:=False, NoHTMLFormatting:=True
        Call LogOff
        Range("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
            :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1)), TrailingMinusNumbers:=True
        Columns("A:A").Delete Shift:=xlToLeft
        Range("A1").Select: While Range("A" & ActiveCell.Row) = "": Range("A" & ActiveCell.Row).EntireRow.Delete: DoEvents: Wend: ActiveCell.Offset(1, 0).Select: If Range("A" & ActiveCell.Row) = "" Then Range("A" & ActiveCell.Row).EntireRow.Delete
        Range("1:1").Font.Bold = True
        Columns("A:F").Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        Range("A1:F1").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Columns("A:F").EntireColumn.AutoFit
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Call Trim_Clean_RemoveAllNonPrintableCharacters
        Call SaveCurrentSheet
    End If
End Sub
Sub Trim_Clean_RemoveAllNonPrintableCharacters()
    Dim Rng As Range
    Dim C As Range
    Dim lngMemoCalculation As Long
    Const csBLANK As String = " "
    Set Rng = Range("A1", Range("A1").SpecialCells(xlLastCell)).SpecialCells(xlCellTypeConstants, 2)
    Application.ScreenUpdating = False
    lngMemoCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    For Each C In Rng.Cells
        C.Value = Application.WorksheetFunction.Clean( _
        Trim(Replace(C.Value, Chr(160), csBLANK)))
    Next C
    Application.Calculation = lngMemoCalculation ' Restore original Calculation mode
End Sub
Sub GetComponentBatchInfo()
    Range("AA2") = "'=" & Range("B" & ActiveCell.Row)
    Sheets("AllMaterialStock").Columns("A:J").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range("AA1:AA2"), CopyToRange:=Range("AB1:AF1"), Unique:=False
End Sub
Function FillColorRGB(Target As Range) As Variant
    Dim N As Double
    N = Target.Interior.Color
    FillColorRGB = Trim(Str(N Mod 256)) & "," & Trim(Str(Int(N / 256) Mod 256)) & "," & Trim(Str(Int(N / 256 / 256) Mod 256))
End Function
Sub BackToTree()
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
    Sheets("MakeTree").Select
    Application.EnableEvents = True
End Sub
Sub BackToRouting()
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
    Sheets("MaterialRouting").Select
    'Sheets("MaterialRouting").Range("S:S") = Sheets("AllOperations").Range("A:A").Value
    Application.EnableEvents = True
End Sub
Sub BackToSettings()
    Sheets("Settings").Select
    Application.EnableEvents = True
End Sub
Sub FillModBOMSheet()
    Dim MatNum As String
    Dim OrderNum As String
    Sheets("MakeTree").Select
    If ActiveCell.Row > 3 And Range("T" & ActiveCell.Row) <> "" Then
        MatNum = Sheets("MakeTree").Range("T" & ActiveCell.Row)
        OrderNum = Sheets("MakeTree").Range("AD" & ActiveCell.Row)
        If OrderNum = "" Then MsgBox ("Must convert the planned order 1st"): End
        OrderNum = Sheets("MakeTree").Range("AD" & ActiveCell.Row)
        If InStr(OrderNum, "Choose") <> 0 Then MsgBox ("Error converting planned order."): End
        If OrderNum <> "" Then Call CO02_GetProductionOrderBOM(Range("AD" & ActiveCell.Row)) Else MsgBox ("Not able to convert planned order")
    Else
        If ActiveCell.Row = 2 And Range("A" & ActiveCell.Row) <> "" Then
            MatNum = Sheets("MakeTree").Range("A" & ActiveCell.Row)
            OrderNum = Sheets("MakeTree").Range("T" & ActiveCell.Row)
            If OrderNum = "" Then MsgBox ("Must convert the planned order 1st"): End
            OrderNum = Sheets("MakeTree").Range("T" & ActiveCell.Row)
            If OrderNum <> "" Then Call CO02_GetProductionOrderBOM(Sheets("MakeTree").Range("T2")) Else MsgBox ("Not able to convert planned order")
        Else
            MsgBox ("Please select a valid line before clicking the Modify BOM button")
        End If
    End If
    Sheets("ModifyBOM").Range("J1") = MatNum
    Sheets("ModifyBOM").Range("J2") = OrderNum
    Sheets("ModifyBOM").Range("AA2:AE1000") = ""
End Sub
Sub GetBOMDataFromSAP()
    If ActiveCell.Row > 3 And Range("AD" & ActiveCell.Row) <> "" Then
        Sheets("ModifyBOM").Select
        Call CO02_GetProductionOrderBOM(Range("AD" & ActiveCell.Row))
    End If
End Sub
Sub CO02_GetProductionOrderBOM(ProductionOrderNum As String)
    Dim I As Integer
    Dim MyText As String
    If Sheets("Settings").Range("B13") = "Production" Then Call LogonProduction Else Call LogonDevelopment
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nco02"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = ProductionOrderNum '"1043244"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/tbar[1]/btn[6]").press
    I = 0
    Sheets("ModifyBOM").Range("A4:L10000") = ""
    MyText = ""
    On Error Resume Next
    MyText = Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & I & "]").Text
    On Error GoTo 0
    While MyText <> ""
        If Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/chkRESBD-XLOEK[16," & I & "]").Selected = False Then
            Sheets("ModifyBOM").Range("A" & 4 + I) = Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-POSNR[0," & I & "]").Text ' = "0010"
            Sheets("ModifyBOM").Range("B" & 4 + I) = Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & I & "]").Text ' = "BL1289"
            Sheets("ModifyBOM").Range("C" & 4 + I) = Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MATXT[2," & I & "]").Text
            Sheets("ModifyBOM").Range("D" & 4 + I) = Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3," & I & "]").Text ' = "55 "
            Sheets("ModifyBOM").Range("E" & 4 + I) = Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-EINHEIT[4," & I & "]").Text ' = "MG"
            Sheets("ModifyBOM").Range("F" & 4 + I) = Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-CHARG[10," & I & "]").Text ' = "MG"
            Sheets("ModifyBOM").Range("G" & 4 + I) = Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/chkRESBD-RGEKZ[13," & I & "]").Selected ' = "MG"
            Sheets("ModifyBOM").Range("H" & 4 + I) = Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-POSTP[5," & I & "]").Text ' = "MG"
            Sheets("ModifyBOM").Range("I" & 4 + I) = Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-VORNR[6," & I & "]").Text ' = "MG"
            Sheets("ModifyBOM").Range("J" & 4 + I) = Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRCOLS-APLFL[7," & I & "]").Text ' = "MG"
            Sheets("ModifyBOM").Range("K" & 4 + I) = Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-WERKS[8," & I & "]").Text ' = "MG"
            Sheets("ModifyBOM").Range("L" & 4 + I) = Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-LGORT[9," & I & "]").Text ' = "MG"
        End If
        I = I + 1
        MyText = ""
        On Error Resume Next
        MyText = Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & I & "]").Text
        On Error GoTo 0
        DoEvents
    Wend
    Call LogOff
    Sheets("ModifyBOM").Select
End Sub
Sub GetRoutingDataFromSAP()
    Dim MatNum As String
    Dim ProdOrder As String
    Dim TargetQty As String
    Dim BaseUnits As String
    Dim MyButton As String
    MyButton = ActiveSheet.Shapes(Application.Caller).Name
    If MyButton = "MatRouting" Then Sheets("MaterialRouting").Select: End
    If ActiveCell.Row = 1 Then End
    If ActiveCell.Row = 3 Then End
    If ActiveCell.Row > 3 And Range("AD" & ActiveCell.Row) <> "" Then
        MatNum = Range("T" & ActiveCell.Row)
        ProdOrder = Range("AD" & ActiveCell.Row)
        If ProdOrder <> "" Then
            TargetQty = Range("U" & ActiveCell.Row)
            BaseUnits = Range("Z" & ActiveCell.Row)
            Call CO02_GetProductionOrderRouting(ProdOrder)
            If Range("T" & ActiveCell.Row) <> "" Then
'                If Sheets("MaterialRouting").Range("A2") = 1 And Sheets("MaterialRouting").Range("A3") = 9999 Then Call GetRoutingFromDatabase(ProdOrder, MatNum, TargetQty, BaseUnits)
            End If
        Else
            MsgBox ("No Production Order")
        End If
    Else
        'Stop
        If ActiveCell.Row > 3 Then
            If Range("T" & ActiveCell.Row) <> "" Then
                MsgBox ("Must convert planned order 1st"): End
            End If
        End If
    End If
    Sheets("MakeTree").Select
    If ActiveCell.Row = 2 And Range("T" & ActiveCell.Row) <> "" Then
        MatNum = Range("A" & ActiveCell.Row)
        ProdOrder = Range("T" & ActiveCell.Row)
        If ProdOrder <> "" Then
            TargetQty = Range("D" & ActiveCell.Row)
            BaseUnits = Range("Q" & ActiveCell.Row)
            Call CO02_GetProductionOrderRouting(ProdOrder)
'            If Sheets("MaterialRouting").Range("A2") = 1 And Sheets("MaterialRouting").Range("A3") = 9999 Then Call GetRoutingFromDatabase(ProdOrder, MatNum, TargetQty, BaseUnits)
        Else
            MsgBox ("No Production Order")
        End If
    Else
        'Stop
        If ActiveCell.Row = 2 Then
            If Range("A" & ActiveCell.Row) <> "" Then
                MsgBox ("Must convert planned order 1st"): End
            End If
        End If
    End If
    Sheets("MaterialRouting").Select
    Sheets("MaterialRouting").Range("Q1") = MatNum
    Sheets("MaterialRouting").Range("P1") = ProdOrder
    Sheets("MaterialRouting").Range("S1") = TargetQty
    Sheets("MaterialRouting").Range("T1") = BaseUnits
'    Sheets("MaterialRouting").Range("R1").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],Routing2MatNum!C[-17]:C[-16],2,FALSE),""Not in routing database"")"
'    Sheets("MaterialRouting").Range("R1") = Sheets("MaterialRouting").Range("R1").Value
End Sub
'Sub GetRoutingFromDatabase(ProductionOrderNum As String, MatNum As String, OrderQty As String, BaseUnits As String)
'    Dim MyVal As String
'    Dim MatGrouping As String
'    Dim CurOperation As String
'    Dim I As Integer
'    On Error Resume Next
'    Application.EnableEvents = False
'    Sheets("MaterialRouting").Range("Q1") = MatNum
'    Sheets("MaterialRouting").Range("P1") = ProductionOrderNum
'    Sheets("MaterialRouting").Range("R1").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],Routing2MatNum!C[-17]:C[-16],2,FALSE),""Not in routing database"")"
'    Sheets("MaterialRouting").Range("R1") = Sheets("MaterialRouting").Range("R1").Value
'    MyVal = WorksheetFunction.Match(MatNum, Sheets("Routing2MatNum").Range("A:A"), 0)
'    If MyVal <> "" Then
'        Call GetRoutingDatabase
'        Sheets("MaterialRouting").Range("A2:O10000") = ""
'        MatGrouping = Sheets("Routing2MatNum").Range("B" & MyVal)
'        MyVal = WorksheetFunction.Match(MatGrouping, Sheets("RoutingDatabase").Range("A:A"), 0)
'        'Stop
'        Sheets("MaterialRouting").Select
'        I = 0
'        Sheets("MaterialRouting").Range("A1:O1") = Sheets("RoutingDatabase").Range("B1:P1").Value
'        CurOperation = Sheets("RoutingDatabase").Range("B" & MyVal).Offset(I, 0)
'        While Sheets("RoutingDatabase").Range("B" & MyVal).Offset(I, 0) <> "9999"
'            Sheets("MaterialRouting").Range("A2").Offset(I, 0) = Sheets("RoutingDatabase").Range("B" & MyVal).Offset(I, 0)
'            Sheets("MaterialRouting").Range("B2").Offset(I, 0) = Sheets("RoutingDatabase").Range("C" & MyVal).Offset(I, 0)
'            Sheets("MaterialRouting").Range("C2").Offset(I, 0) = Sheets("RoutingDatabase").Range("D" & MyVal).Offset(I, 0)
'            Sheets("MaterialRouting").Range("D2").Offset(I, 0) = Sheets("RoutingDatabase").Range("E" & MyVal).Offset(I, 0)
'            Sheets("MaterialRouting").Range("E2").Offset(I, 0) = OrderQty
'            Sheets("MaterialRouting").Range("F2").Offset(I, 0) = BaseUnits
'            Sheets("MaterialRouting").Range("G2").Offset(I, 0) = Sheets("RoutingDatabase").Range("H" & MyVal).Offset(I, 0)
'            Sheets("MaterialRouting").Range("H2").Offset(I, 0) = Sheets("RoutingDatabase").Range("I" & MyVal).Offset(I, 0)
'            Sheets("MaterialRouting").Range("I2").Offset(I, 0) = Sheets("RoutingDatabase").Range("J" & MyVal).Offset(I, 0)
'            Sheets("MaterialRouting").Range("J2").Offset(I, 0) = Sheets("RoutingDatabase").Range("K" & MyVal).Offset(I, 0)
'            Sheets("MaterialRouting").Range("K2").Offset(I, 0) = Sheets("RoutingDatabase").Range("L" & MyVal).Offset(I, 0)
'            Sheets("MaterialRouting").Range("L2").Offset(I, 0) = Sheets("RoutingDatabase").Range("M" & MyVal).Offset(I, 0)
'            Sheets("MaterialRouting").Range("M2").Offset(I, 0) = Sheets("RoutingDatabase").Range("N" & MyVal).Offset(I, 0)
'            Sheets("MaterialRouting").Range("N2").Offset(I, 0) = Sheets("RoutingDatabase").Range("O" & MyVal).Offset(I, 0)
'            Sheets("MaterialRouting").Range("O2").Offset(I, 0) = Sheets("RoutingDatabase").Range("P" & MyVal).Offset(I, 0)
'            I = I + 1
'            CurOperation = Sheets("RoutingDatabase").Range("B" & MyVal).Offset(I, 0)
'            DoEvents
'        Wend
'        Sheets("MaterialRouting").Range("A2").Offset(I, 0) = Sheets("RoutingDatabase").Range("B" & MyVal).Offset(I, 0)
'        Sheets("MaterialRouting").Range("B2").Offset(I, 0) = Sheets("RoutingDatabase").Range("C" & MyVal).Offset(I, 0)
'        Sheets("MaterialRouting").Range("C2").Offset(I, 0) = Sheets("RoutingDatabase").Range("D" & MyVal).Offset(I, 0)
'        Sheets("MaterialRouting").Range("D2").Offset(I, 0) = Sheets("RoutingDatabase").Range("E" & MyVal).Offset(I, 0)
'        Sheets("MaterialRouting").Range("E2").Offset(I, 0) = OrderQty
'        Sheets("MaterialRouting").Range("F2").Offset(I, 0) = BaseUnits
'        Sheets("MaterialRouting").Range("G2").Offset(I, 0) = Sheets("RoutingDatabase").Range("H" & MyVal).Offset(I, 0)
'        Sheets("MaterialRouting").Range("H2").Offset(I, 0) = Sheets("RoutingDatabase").Range("I" & MyVal).Offset(I, 0)
'        Sheets("MaterialRouting").Range("I2").Offset(I, 0) = Sheets("RoutingDatabase").Range("J" & MyVal).Offset(I, 0)
'        Sheets("MaterialRouting").Range("J2").Offset(I, 0) = Sheets("RoutingDatabase").Range("K" & MyVal).Offset(I, 0)
'        Sheets("MaterialRouting").Range("K2").Offset(I, 0) = Sheets("RoutingDatabase").Range("L" & MyVal).Offset(I, 0)
'        Sheets("MaterialRouting").Range("L2").Offset(I, 0) = Sheets("RoutingDatabase").Range("M" & MyVal).Offset(I, 0)
'        Sheets("MaterialRouting").Range("M2").Offset(I, 0) = Sheets("RoutingDatabase").Range("N" & MyVal).Offset(I, 0)
'        Sheets("MaterialRouting").Range("N2").Offset(I, 0) = Sheets("RoutingDatabase").Range("O" & MyVal).Offset(I, 0)
'        Sheets("MaterialRouting").Range("O2").Offset(I, 0) = Sheets("RoutingDatabase").Range("P" & MyVal).Offset(I, 0)
'    End If
'    Application.EnableEvents = True
'End Sub
Sub OpenPDF()
    Dim MyPath As String
    Dim FileToOpen As String
    MyPath = "\\BETHYL-SERVER2\ProductionRecords\"
    SetCurrentDirectoryA (MyPath)
    FileToOpen = Application.GetOpenFilename(Title:="Please choose a file to open", FileFilter:="PDF Files *.PDF (*.PDF),")
    If FileToOpen <> "False" Then ActiveWorkbook.FollowHyperlink FileToOpen
End Sub
'Sub GetRoutingDatabase()
'    Dim MyPath As String
'    Dim MySheet As String
'    Application.ScreenUpdating = False
'    Application.DisplayAlerts = False
'    MySheet = ActiveSheet.Name
'    MyPath = ThisWorkbook.Path
'    On Error Resume Next
'    Sheets(Array("Routing2MatNum", "RoutingDatabase")).Delete
'    On Error GoTo 0
'    Workbooks.Open Filename:=MyPath & "\MaterialRoutingDatabase.xlsx", ReadOnly:=True
'    Sheets(Array("Routing2MatNum", "RoutingDatabase")).Copy After:=Workbooks(ThisWorkbook.Name).Sheets(21)
'    Workbooks("MaterialRoutingDatabase.xlsx").Close False
'    Sheets(MySheet).Select
'End Sub

Sub TestProdOrderRouting()
    Call CO02_GetProductionOrderRouting(1043704)
End Sub
Sub CO02_GetProductionOrderRouting(ProductionOrderNum As String)
    Dim I As Integer
    Dim X As Integer
    Dim Z As Integer
    Dim MyText As String
    Dim MyVal As String
    Dim MatNum As String
    Dim OrderQty As String
    Dim BaseUnits As String
    'Stop
    'add code to fill in new info on the routing sheet
    Application.EnableEvents = False
    Sheets("MaterialRouting").Range("A4:O2000") = ""
    If Sheets("Settings").Range("B13") = "Production" Then Call LogonProduction Else Call LogonDevelopment
    Session.FindById("wnd[0]").ResizeWorkingPane 177, 42, False
ReTry:
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nco02"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = ProductionOrderNum '"1043244"
    Session.FindById("wnd[0]").SendVKey 0
    On Error GoTo ReTry
    OrderQty = Trim(Session.FindById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/txtCAUFVD-GAMNG").Text)
    BaseUnits = Trim(Session.FindById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/ctxtCAUFVD-GMEIN").Text)
    MatNum = Session.FindById("wnd[0]/usr/ctxtCAUFVD-MATNR").Text
    Sheets("MaterialRouting").Select
    Session.FindById("wnd[0]/tbar[1]/btn[5]").press
    I = 0
    Z = 0
    Sheets("MaterialRouting").Range("A2:O10000") = ""
    Range("A2").Select
    MyVal = ""
    On Error Resume Next
    MyVal = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8," & I & "]").Text
    On Error GoTo 0
    While MyVal <> ""
        'Stop
        Range("A" & ActiveCell.Row).Offset(Z, 0) = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0," & I & "]").Text
        Range("B" & ActiveCell.Row).Offset(Z, 0) = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-ARBPL[4," & I & "]").Text ' = "3600"
        'Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-WERKS[5," & I & "]").Text = "3000"
        Range("C" & ActiveCell.Row).Offset(Z, 0) = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-STEUS[6," & I & "]").Text ' = "ZP01"
        Range("M" & ActiveCell.Row).Offset(Z, 0) = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-KTSCH[7," & I & "]").Text ' = "JDF"
        Range("D" & ActiveCell.Row).Offset(Z, 0) = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8," & I & "]").Text ' = "START"
        Range("E" & ActiveCell.Row).Offset(Z, 0) = OrderQty
        Range("F" & ActiveCell.Row).Offset(Z, 0) = BaseUnits
        If Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/chkAFVGD-FLG_FHM[13," & I & "]").Selected = True Then
            Range("N" & ActiveCell.Row).Offset(I, 0) = True
            Range("N" & ActiveCell.Row).Offset(I, 0) = GetIS(MatNum, Z)
        End If
        Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100").GetAbsoluteRow(Z).Selected = True
        Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnSHOWDETAIL").press
        Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW").Select
        Range("G" & ActiveCell.Row).Offset(Z, 0) = Val(Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW01").Text) ' = "1"
        Range("H" & ActiveCell.Row).Offset(Z, 0) = Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE01").Text ' = "H"
        Range("I" & ActiveCell.Row).Offset(Z, 0) = Val(Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW02").Text) ' = "2"
        Range("J" & ActiveCell.Row).Offset(Z, 0) = Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE02").Text ' = "MN"
        Range("K" & ActiveCell.Row).Offset(Z, 0) = Val(Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW03").Text) ' = "3"
        Range("L" & ActiveCell.Row).Offset(Z, 0) = Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE03").Text ' = "D"
        Session.FindById("wnd[0]/tbar[0]/btn[3]").press
        Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnDESELECT").press
        If Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/chkRC270-TXTKZ[9," & I & "]").Selected = True Then
            Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100").GetAbsoluteRow(Z).Selected = True
            Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnLONG_TEXT").press
'            Session.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select
            MyText = ""
            On Error Resume Next
            X = 2
            MyText = ""
            MyText = Session.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & X & "]").Text
            While MyText <> ""
                Range("O" & ActiveCell.Row).Offset(Z, 0) = Range("O" & ActiveCell.Row).Offset(I, 0) & MyText ' = "START"
                X = X + 1
                MyText = ""
                MyText = Session.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & X & "]").Text
                DoEvents
            Wend
            Session.FindById("wnd[0]/tbar[0]/btn[3]").press
            Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnDESELECT").press
        End If
        On Error Resume Next
        If I = 30 Then
            Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100").VerticalScrollbar.Position = Z + 1
            I = 0
            Z = Z + 1
        Else
            I = I + 1
            Z = Z + 1
        End If
        MyVal = ""
        MyVal = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0," & I & "]").Text
        If Val(MyVal) = Val(Range("A" & ActiveCell.Row).Offset(Z - 1, 0)) Then MyVal = ""
        If MyVal <> "" Then MyVal = "": MyVal = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8," & I & "]").Text
        On Error GoTo 0
        DoEvents
    Wend
    Call LogOff
    Sheets("MaterialRouting").Select
    Sheets("MaterialRouting").Range("P1") = ProductionOrderNum
    Sheets("MaterialRouting").Range("Q1") = MatNum
'    Sheets("MaterialRouting").Range("R1").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],Routing2MatNum!C[-17]:C[-16],2,FALSE),""Not in routing database"")"
'    Sheets("MaterialRouting").Range("R1") = Sheets("MaterialRouting").Range("R1").Value
    Application.EnableEvents = True
    Columns("O:O").ColumnWidth = 24
    Columns("N:N").ColumnWidth = 12
End Sub
Sub btnPullDownFilter_Click()
    Dim MyRange As String
    Dim MyVal As String
    Dim MySearchVal As String
    Dim I As Integer
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    If ActiveSheet.Shapes.Range(Array("btnPullDownFilter")).TextFrame2.TextRange.Characters.Text = "Enable Dropdown Filter" Then
        ActiveSheet.Shapes.Range(Array("btnPullDownFilter")).TextFrame2.TextRange.Characters.Text = "Disable Dropdown Filter"
        MyRange = "=indirect(""AllOperations!A:A" & """" & ")"
        Range("D:D").Validation.Delete
        I = 2
        While I < 100
            Range("D" & I).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=MyRange
            DoEvents
            I = I + 1
        Wend
    Else
        ActiveSheet.Shapes.Range(Array("btnPullDownFilter")).TextFrame2.TextRange.Characters.Text = "Enable Dropdown Filter"
        Range("D:D").Validation.Delete
        I = 2
        While I < 100
            On Error Resume Next
            MyVal = ""
            MySearchVal = WorksheetFunction.Proper(Range("D" & I - 1))
            MyVal = WorksheetFunction.VLookup(MySearchVal, Sheets("AllOperations").Range("A:B"), 2, False)
            If MyVal = "" Then
                MySearchVal = WorksheetFunction.UCase(Range("D" & I - 1))
                MyVal = WorksheetFunction.VLookup(MySearchVal, Sheets("AllOperations").Range("A:B"), 2, False)
            End If
            If MyVal = "" Then
                MySearchVal = WorksheetFunction.LCase(Range("D" & I - 1))
                MyVal = WorksheetFunction.VLookup(MySearchVal, Sheets("AllOperations").Range("A:B"), 2, False)
            End If
            If MyVal = "" Then
                MySearchVal = Range("D" & I - 1)
                MyVal = WorksheetFunction.VLookup(MySearchVal, Sheets("AllOperations").Range("A:B"), 2, False)
            End If
            On Error GoTo 0
            If MyVal <> "" Then
                MyRange = "=indirect(""AllOperations!" & GetStepRange(MySearchVal) & """" & ")"
            Else
                MyRange = "=indirect(""AllOperations!A:A" & """" & ")"
            End If
            Range("D" & I).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=MyRange
            DoEvents
            I = I + 1
        Wend
    End If
    Application.EnableEvents = True
End Sub
Sub SaveRouting_CO02()
    If Sheets("Settings").Range("B13") = "Production" Then Call LogonProduction Else Call LogonDevelopment
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nco02"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = Sheets("MaterialRouting").Range("N1") '"1042870"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/tbar[1]/btn[5]").press
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0,0]").Text = "0010"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0,1]").Text = "0020"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0,2]").Text = "0030"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0,3]").Text = "9010"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-ARBPL[4,0]").Text = "3600"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-ARBPL[4,1]").Text = "3600"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-ARBPL[4,2]").Text = "3700"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-ARBPL[4,3]").Text = "3600"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-WERKS[5,0]").Text = "3000"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-WERKS[5,1]").Text = "3000"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-WERKS[5,2]").Text = "3000"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-WERKS[5,3]").Text = "3000"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-STEUS[6,0]").Text = "ZP01"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-STEUS[6,1]").Text = "ZP01"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-STEUS[6,2]").Text = "ZP03"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-STEUS[6,3]").Text = "ZP01"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-KTSCH[7,0]").Text = "JDF"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-KTSCH[7,1]").Text = "JDF"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-KTSCH[7,2]").Text = "JDF"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-KTSCH[7,2]").SetFocus
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-KTSCH[7,2]").CaretPosition = 3
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-KTSCH[7,0]").Text = "JDFIII"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-KTSCH[7,1]").Text = "JDFIII"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-KTSCH[7,2]").Text = "JDFIII"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-KTSCH[7,3]").Text = "JDFIII"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8,0]").Text = "START"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8,1]").Text = "MAKE CONJUGATE"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8,2]").Text = "QUALITY CONTROL"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8,3]").Text = "START2"
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0,0]").SetFocus
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0,0]").CaretPosition = 1
    Session.FindById("wnd[0]").SendVKey 2
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    Session.FindById("wnd[1]").SendVKey 0
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW").Select
    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW01").Text = "1"
    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW03").Text = "20"
    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW03").SetFocus
    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW03").CaretPosition = 2
    Session.FindById("wnd[0]/tbar[0]/btn[11]").press
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/tbar[1]/btn[5]").press
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0,1]").SetFocus
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0,1]").CaretPosition = 3
    Session.FindById("wnd[0]").SendVKey 2
    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW").Select
    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW01").Text = "11"
    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW03").Text = "2"
    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW03").SetFocus
    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW03").CaretPosition = 1
    Session.FindById("wnd[0]/tbar[0]/btn[11]").press
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/tbar[1]/btn[5]").press
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0,2]").SetFocus
    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0,2]").CaretPosition = 2
    Session.FindById("wnd[0]").SendVKey 2
    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW").Select
    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW01").Text = "3"
    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW03").Text = "2"
    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW03").SetFocus
    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW03").CaretPosition = 1
    Session.FindById("wnd[0]/tbar[0]/btn[11]").press
    Call LogOff
End Sub
Sub HideUnusedCells()
    Dim I As Integer, X As Integer
    Range("A2:M10").EntireRow.Hidden = True
    Range("A1:M10").EntireColumn.Hidden = True
    Range("A1").EntireRow.Hidden = False
    Range("C1").EntireColumn.Hidden = False
    Range("D1").Select
    While ActiveCell.Row <> 10
        If Range("D" & ActiveCell.Row) <> "" Then Rows(ActiveCell.Row & ":" & ActiveCell.Row).EntireRow.Hidden = False
        ActiveCell.Offset(1, 0).Select
        DoEvents
    Wend
    For I = 2 To 20
        For X = 2 To 20
            If Cells(I, X) <> "" Then Cells(I, X).EntireColumn.Hidden = False
        Next X
    Next I
End Sub

Sub CO01_MakeProductionOrder()
    End
'    If Range("B" & ActiveCell.Row) = "" Then MsgBox ("Select a line wia Material Number before clicking this button."): End
'    If Range("G" & ActiveCell.Row) < Date Then MsgBox ("Enter a date in the future for this order."): End
'    If Range("H" & ActiveCell.Row).Interior.Color = 10498160 Then Range("H" & ActiveCell.Row) = CO01_MakeNewProductionOrder(Range("B" & ActiveCell.Row), Range("G" & ActiveCell.Row)) Else MsgBox ("Double click the order before clicking the TECO button.")
End Sub
Sub TECO_CurrentlySelectedOrder()
    End
'    If Range("H" & ActiveCell.Row).Interior.Color = 10498160 Then Range("H" & ActiveCell.Row) = CO02_TECO(Range("A" & ActiveCell.Row)) Else MsgBox ("Double click the order before clicking the TECO button.")
End Sub
Sub GetCurrentSheet()
    Dim MySheet As String
    MySheet = ActiveSheet.Name
    Workbooks.Open Filename:="\\BETHYL-FS1\BethylShared\SAPData" & "\SAP_Dashboard_" & MySheet & ".xlsx"
    Cells.Copy
    ThisWorkbook.Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Workbooks("SAP_Dashboard_" & MySheet & ".xlsx").Close False
End Sub
Sub EnableEventsNow()
    Application.EnableEvents = True
End Sub
Sub ChangeFinishDate_Email_CS()
    Dim MatNum As String
    Dim NewFDate As Date
    Dim MyOrderNum As String
    Dim MyEmail As String
    Dim MyCell As String
    Dim LastRow As String
    Dim RouteStep As Integer
    Range("A4").FormulaR1C1 = "=VLOOKUP(R[-2]C,Settings!C1:C2,2,FALSE)"
    Range("A4") = Range("A4").Value
    If Range("H" & ActiveCell.Row).Interior.Color = 10498160 Then
        MyOrderNum = Range("A" & ActiveCell.Row)
        MatNum = Range("B" & ActiveCell.Row)
        NewFDate = Range("G" & ActiveCell.Row)
        RouteStep = Val(Range("A4"))
        If RouteStep > 0 Then
            Range("I" & ActiveCell.Row) = "Step " & RouteStep & " Finish Date> " & NewFDate
            NewFDate = GetFinishDateFromRoutingStep(NewFDate, RouteStep, MatNum)
            Range("G" & ActiveCell.Row) = NewFDate
        End If
        'Stop
        If NewFDate <> 0 Then
            Call LogonProduction
            Range("H" & ActiveCell.Row) = ChangeFinishDate(Range("A" & ActiveCell.Row), NewFDate)
            Call LogOff
            Sheets("ZVialOrders").Select
            Sheets("ZVialOrders").Range("Q2:W" & 100000) = ""
            Sheets("ZVialOrders").Range("O2") = MatNum
            Sheets("ZVialOrders").Range("A:K").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Sheets("ZVialOrders").Range("O1:O2"), CopyToRange:=Sheets("ZVialOrders").Range("R1:V1"), Unique:=False
            Sheets("ZVialOrders").Range("W1") = "Entered By"
            Sheets("ZVialOrders").Range("X1") = "New Delivery Date"
            Sheets("ZVialOrders").Range("Q1") = "Email"
            Sheets("ZVialOrders").Columns("Q:Y").EntireColumn.AutoFit
            Sheets("ZVialOrders").Range("R1").Select
            Sheets("ZVialOrders").Range("R1000000").End(xlUp).Select
            LastRow = ActiveCell.Row
            If LastRow <> 1 Then
                Sheets("ZVialOrders").Range("W2:W" & 100000) = ""
                Sheets("ZVialOrders").Range("W2:W" & LastRow).FormulaR1C1 = "=VLOOKUP(RC[-1],SalesDoc!C1:C2,2,FALSE)"
                Sheets("ZVialOrders").Range("Q2:Q" & LastRow).FormulaR1C1 = "=VLOOKUP(RC[6],Settings!C30:C31,2,FALSE)"
                Sheets("ZVialOrders").Range("Q2:W" & LastRow).Copy
                Sheets("ZVialOrders").Range("Q2").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            End If
            Range("X2") = NewFDate
            Range("Q2").Select
            MyEmail = Range("Q" & ActiveCell.Row)
            While Range("Q" & ActiveCell.Row) <> ""
                MyCell = ActiveCell.Address
                Call SendCSEmail(MyEmail, "jfrost@bethyl.com")
                Range(MyCell).Select
                ActiveCell.Offset(1, 0).Select
                MyEmail = Range("Q" & ActiveCell.Row)
                DoEvents
            Wend
        End If
    Else
        MsgBox ("Before cicking on this button, double click on the row you want to change.")
    End If
    Sheets("Main").Select
    Application.EnableEvents = True
End Sub
Sub GetOrders()
Attribute GetOrders.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim MyRange As String
    Dim MyCNum As Integer
    Application.ScreenUpdating = False
    On Error Resume Next
    ActiveSheet.Shapes.Range("picOperations").Delete
    ActiveSheet.Shapes.Range("picComponents").Delete
    ActiveSheet.Shapes.Range("picZVial").Delete
    Range("AA2:AC100") = ""
    On Error GoTo 0
    Range("A6:I20006") = ""
    Range("A6:H20006").Font.Bold = False
    With Range("G6:H20006").Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Range("A6:H20006").Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Sheets("Main").Range("AA1") = "MRP ctrlr"
    Sheets("Main").Range("AB1") = "Order Type"
    Sheets("Main").Range("AG1") = "Act. start"
    Sheets("Main").Range("AH1") = "Order"
    Sheets("Main").Range("AG2") = "="
    Sheets("Main").Range("AI2").Formula = "=DMIN(Operations!C1:C15,""Oper./Act."",R[-1]C[-2]:RC[-1])"
    MyRange = WorksheetFunction.Match(Sheets("Main").Range("A2"), Sheets("Settings").Range("A:A"), 0)
    Sheets("Settings").Range("C" & MyRange & ":X" & MyRange).Copy
    Sheets("Main").Select
    Range("AA2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    MyCNum = WorksheetFunction.CountA(Sheets("Main").Range("AA:AA"))
    Application.CutCopyMode = False
    On Error GoTo 0
    If Sheets("Main").Range("A3") = "Production Only" Then Sheets("Main").Range("AB2:AB" & MyCNum) = "'=*1"
    If Sheets("Main").Range("A3") = "Planned Only" Then Sheets("Main").Range("AB2:AB" & MyCNum) = "'<>*1"
    If Sheets("Main").Range("A3") = "Planned & Production" Then Sheets("Main").Range("AB2:AB" & MyCNum) = ""
    Sheets("OrderHeaders").Range("A:O").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Sheets("Main").Range("AA1:AB" & MyCNum), CopyToRange:=Sheets("Main").Range("A5:E5"), Unique:=False
    Range("A6").Select
    While Range("A" & ActiveCell.Row) <> ""
        Range("AH2") = Range("A" & ActiveCell.Row)
        Application.StatusBar = ActiveCell.Row
        Range("F" & ActiveCell.Row) = Range("AI2")
        If Range("AI1") <> "All" Then
            If Range("F" & ActiveCell.Row) <> Range("AI1") Then ActiveCell.EntireRow.Delete Else ActiveCell.Offset(1, 0).Select
        End If
        If Range("F" & ActiveCell.Row) <> "" Then ActiveCell.Offset(1, 0).Select
        DoEvents
    Wend
    ActiveWindow.SplitRow = 21
    ActiveWindow.FreezePanes = True
    Application.ScreenUpdating = True
    Range("F1").Select
    ActiveCell.Offset(0, 1).Select
    If Range("A6") <> "" Then
        Range("A5:H1000000").Select
        ActiveWorkbook.Worksheets("Main").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("Main").Sort.SortFields.Add Key:=Range("E6:E20006"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Main").Sort
            .SetRange Range("A5:H20006")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
    Range("A2").Select
    Application.StatusBar = "Ready"
    Application.EnableEvents = True
End Sub
Sub TrimAllCells()
    Dim Cell As Range
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    For Each Cell In ActiveSheet.UsedRange.SpecialCells(xlCellTypeConstants)
        Cell = WorksheetFunction.Trim(Cell)
    Next Cell
    Application.Calculation = xlCalculationAutomatic
End Sub
Sub RunProgram()
    Dim MyEmail As String
    Dim MyReport As String
    Dim MyDay As String
    Dim test As Integer
    MyDay = Format(Date, "dddd")
    If Sheets("Settings").Range("B13") = "Production" Then Call LogonProduction Else Call LogonDevelopment
    Call RunMRP
    Call COHV_GetProductionHeadersOperationsComponents
    Call GetSalesDoc2CreatedBy
    Call ZVialOrders
    Call LogOff
    test = 0
    On Error Resume Next
    test = WorksheetFunction.Match(MyDay, Sheets("Settings").Range("AB29:AB35"))
    On Error GoTo 0
    If test <> 0 Then
        Sheets("Settings").Select
        Range("A2").Select
        While Range("A" & ActiveCell.Row) <> ""
            MyReport = Sheets("Settings").Range("A" & ActiveCell.Row)
            MyEmail = Sheets("Settings").Range("Z" & ActiveCell.Row)
            If MyEmail <> "" Then
                Sheets("Main").Select
                Sheets("Main").Range("A2") = MyReport
                Sheets("Main").Range("A3") = "Production Only"
                Call GetOrders
                'If Sheets("Main").Range("A6") <> "" Then Call SendImageInEmail("Weekly Finish Date Report-> " & Sheets("Main").Range("A2"), Sheets("Main").Range("A5:E" & 1 + Application.WorksheetFunction.CountA(Sheets("Main").Range("A:A"))), MyEmail, "jfrost@bethyl.com") 'SendProdEmail(MyEmail, "jfrost@bethyl.com")
            End If
            Sheets("Settings").Select
            ActiveCell.Offset(1, 0).Select
            DoEvents
        Wend
    End If
    Call GetAllSafetyStock
    test = 0
    On Error Resume Next
    test = WorksheetFunction.Match(MyDay, Sheets("Settings").Range("AC29:AC35"))
    On Error GoTo 0
    If test <> 0 Then Call SendZeroInventoryEmails
    Sheets("Settings").Select
    Call SaveCurrentSheet
    Sheets("SalesDoc").Select
    Call SaveCurrentSheet
    Sheets("OrderHeaders").Select
    Call SaveCurrentSheet
    Sheets("SafetyStock").Select
    Call SaveCurrentSheet
    Sheets("FERT_Inventory").Select
    Call SaveCurrentSheet
    Sheets("Operations").Select
    Call SaveCurrentSheet
    Sheets("Components").Select
    Call SaveCurrentSheet
    Sheets("ZVialOrders").Select
    Call SaveCurrentSheet
End Sub
Sub SendZeroInventoryEmails()
    Sheets("SafetyStock").Select
    Columns("J:M").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:G").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("G1").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Sheets("OrderHeaders").Select
    Range("A1:B1,K1,D1").Select
    Selection.Copy
    Sheets("SafetyStock").Select
    Range("J1").Select
    ActiveSheet.Paste
    Range("M1").Select
    Selection.Cut
    Range("L1").Select
    Selection.Insert Shift:=xlToRight
    Range("M2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "'=*1"
    Sheets("OrderHeaders").Columns("A:O").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range("M1:M2"), CopyToRange:=Range("J1:L1"), Unique:=False
    Columns("J:L").Select
    ActiveWorkbook.Worksheets("SafetyStock").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SafetyStock").Sort.SortFields.Add Key:=Range( _
        "M2:M1000000"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("SafetyStock").Sort
        .SetRange Range("K1:M1000000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("SafetyStock").Select
    Range("D1") = "Inventory"
    Range("E1") = "# of Prod Orders"
    Range("F1") = "1st Order#"
    Range("G1") = "Min Date"
    Range("D2").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],FERT_Inventory!C[-3]:C[2],6,FALSE),"""")"
    Range("E2").FormulaR1C1 = "=COUNTIF(C[6],RC[-4])"
    Range("F2").FormulaR1C1 = "=IFERROR(INDIRECT(""J"" & MATCH(RC[-5],C[5],0)),"""")"
    Range("G2").FormulaR1C1 = "=IFERROR(INDIRECT(""L"" & MATCH(RC1,C[4],0)),"""")"
    Range("D2:G2").AutoFill Destination:=Range("D2:G" & WorksheetFunction.CountA(Range("A:A"))), Type:=xlFillDefault
    Range("D2:G" & WorksheetFunction.CountA(Range("A:A"))).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range("A1:G" & WorksheetFunction.CountA(Range("A:A"))).Select
    ActiveWorkbook.Worksheets("SafetyStock").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SafetyStock").Sort.SortFields.Add Key:=Range( _
        "D2:D1000000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("SafetyStock").Sort.SortFields.Add Key:=Range( _
        "B2:B1000000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("SafetyStock").Sort
        .SetRange Range("A1:G1000000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Call TrimAllCells
    Range("D1").End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    While Range("A" & ActiveCell.Row) <> ""
        If Range("E" & ActiveCell.Row) > 0 And Range("G" & ActiveCell.Row) > Date Then Range("D" & ActiveCell.Row) = 0
        ActiveCell.Offset(1, 0).Select
        DoEvents
    Wend
    Range("A1:G" & WorksheetFunction.CountA(Range("A:A"))).Select
    ActiveWorkbook.Worksheets("SafetyStock").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SafetyStock").Sort.SortFields.Add Key:=Range( _
        "D2:D1000000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("SafetyStock").Sort.SortFields.Add Key:=Range( _
        "B2:B1000000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("SafetyStock").Sort
        .SetRange Range("A1:G1000000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Call TrimAllCells
    Range("D1").End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    Dim CurMRP As String
    Dim StartRange As String
    Dim EndRange As String
    Dim CurEmail As String
    Dim MyRange As String
    CurMRP = Range("B" & ActiveCell.Row)
    On Error Resume Next
    CurEmail = WorksheetFunction.VLookup(CurMRP, Sheets("Settings").Range("AB2:AC20"), 2, False)
    StartRange = ActiveCell.Row
    If StartRange <> "" And CurMRP <> "" Then
        While Range("A" & ActiveCell.Row) <> ""
            'Stop
            If Range("B" & ActiveCell.Row) = CurMRP Then
                ActiveCell.Offset(1, 0).Select
            Else
                EndRange = ActiveCell.Row - 1
                MyRange = "A1:G1,A" & StartRange & ":G" & EndRange
                Call SendProdEmail("Daily 0 Inventory Report-> " & CurMRP, Sheets("SafetyStock").Range(MyRange), CurEmail, "jfrost@bethyl.com")
                CurMRP = Range("B" & ActiveCell.Row)
                CurEmail = WorksheetFunction.VLookup(CurMRP, Sheets("Settings").Range("AB2:AC20"), 2, False)
                StartRange = ActiveCell.Row
            End If
            DoEvents
        Wend
        EndRange = ActiveCell.Row - 1
        MyRange = "A1:G1,A" & StartRange & ":G" & EndRange
        Call SendProdEmail("Daily 0 Inventory Report-> " & CurMRP, Sheets("SafetyStock").Range(MyRange), CurEmail, "jfrost@bethyl.com")
    End If
End Sub
Sub SaveCurrentSheet()
    Dim MyName As String
    Application.DisplayAlerts = False
    MyName = ActiveSheet.Name
    Sheets(MyName).Copy
    Range("ZZ1000000").Select
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
    ActiveWorkbook.SaveAs Filename:="\\BETHYL-FS1\BethylShared\SAPData" & "\SAP_Dashboard_" & MyName & ".xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks("SAP_Dashboard_" & MyName & ".xlsx").Close False
End Sub
Sub InsertBelow()
    Dim MyRow
    On Error GoTo 0
    Range("A" & ActiveCell.Row & ":O" & ActiveCell.Row).Insert Shift:=xlDown
End Sub
Sub DeleteRoutingStep()
    If ActiveCell.Row <> 1 Then Range("A" & ActiveCell.Row & ":O" & ActiveCell.Row).Delete Shift:=xlUp
End Sub
Sub ChangeRouting()
    Dim MyVal
    Dim I
        Dim MyMatNum
    Dim MyLineNum
    'Stop
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nca02"
    Session.FindById("wnd[0]").SendVKey 0
    Sheets("MatNums2Change").Select
    If Range("B" & ActiveCell.Row) <> "No Routing" Then
        Session.FindById("wnd[0]/usr/ctxtRC27M-MATNR").Text = ActiveCell
        MyMatNum = ActiveCell
        Session.FindById("wnd[0]/usr/ctxtRC27M-WERKS").Text = "3000"
        Session.FindById("wnd[0]/usr/ctxtRC271-PLNNR").Text = ""
        Session.FindById("wnd[0]/usr/ctxtRC271-PLNNR").SetFocus
        Session.FindById("wnd[0]/usr/ctxtRC271-PLNNR").CaretPosition = 0
        Session.FindById("wnd[0]").SendVKey 0
        Session.FindById("wnd[0]/tbar[1]/btn[26]").press
        Session.FindById("wnd[0]/tbar[1]/btn[14]").press
        Session.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
        Sheets("RoutingData").Select
        Range("A2").Select
        I = 0
        While Range("A" & ActiveCell.Row).Value <> ""
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VORNR[0," & I & "]").Text = Range("A" & ActiveCell.Row).Value 'Operat.
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-ARBPL[2," & I & "]").Text = Range("B" & ActiveCell.Row).Value 'Work Center
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-WERKS[3," & I & "]").Text = 3000 'Range("C" & ActiveCell.Row).Value 'PLNT
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-STEUS[4," & I & "]").Text = Range("C" & ActiveCell.Row).Value 'Control Key
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-KTSCH[5," & I & "]").Text = Range("I" & ActiveCell.Row).Value 'Step Operator
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6," & I & "]").Text = Range("D" & ActiveCell.Row).Value 'Description
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-BMSCH[14," & I & "]").Text = Range("F" & ActiveCell.Row).Value 'Base Quantity
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-MEINH[15," & I & "]").Text = Range("G" & ActiveCell.Row).Value 'Unit of Measure
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[16," & I & "]").Text = Range("H" & ActiveCell.Row).Value 'Setup
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-VGE01[17," & I & "]").Text = Range("I" & ActiveCell.Row).Value 'Setup Unit
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[19," & I & "]").Text = "" 'Range("J" & ActiveCell.Row).Value 'Machine
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-VGE02[20," & I & "]").Text = "" 'Range("K" & ActiveCell.Row).Value 'Machine Unit
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW03[22," & I & "]").Text = Range("L" & ActiveCell.Row).Value 'Labor
            Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-VGE03[23," & I & "]").Text = Range("M" & ActiveCell.Row).Value 'Labor Unit
            ActiveCell.Offset(1, 0).Select
            MyLineNum = Session.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VORNR[0,0]").Text
            I = I + 1
            On Error Resume Next
            MyVal = Range("A" & ActiveCell.Row).Value
            On Error GoTo 0
            DoEvents
        Wend
        Session.FindById("wnd[0]/tbar[0]/btn[11]").press
        MyVal = Session.FindById("wnd[0]/sbar").Text
        Sheets("MatNums2Change").Select
        ActiveCell.Offset(0, 1).Value = MyVal
    End If
End Sub
Sub GetMRPControllers()
    Dim MySheet As String
    Dim MyUnits As String
    Dim MyUFreq As Integer
    Dim MyRow As Long
    Dim MyDate As Date
    Dim MyCompDate As Date
    Application.DisplayAlerts = False
    Sheets("MD07").Select
    MySheet = ActiveSheet.Name
    MyRow = WorksheetFunction.Match(MySheet, Sheets("Settings").Range("S:S"), 0)
    MyUnits = Sheets("Settings").Range("U" & MyRow)
    MyUFreq = Sheets("Settings").Range("T" & MyRow)
    'if getmoddate is more than 1 week old then run the code below otherwise call getcurrentsheet
    MyDate = GetModDate("\\BETHYL-FS1\BethylShared\SAPData" & "\SAP_Dashboard_" & MySheet & ".xlsx")
    MyCompDate = DateAdd(MyUnits, MyUFreq, Now)
    If UCase(Sheets("Settings").Range("G1")) = "FALSE" Then
        Sheets("MD07").Select
        Call GetCurrentSheet
    Else
        If Sheets("Settings").Range("B13") = "Production" Then Call LogonProduction Else Call LogonDevelopment
        Sheets("MD07").Select
        Range("A1") = "#"
        Range("A1").Select
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:="#", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
        Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmd07"
        Session.FindById("wnd[0]").SendVKey 0
        Session.FindById("wnd[0]").SendVKey 4
        Session.FindById("wnd[1]").SendVKey 14
        Session.FindById("wnd[2]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
        Session.FindById("wnd[2]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
        Session.FindById("wnd[2]/tbar[0]/btn[0]").press
        Selection.PasteSpecial
        Call LogOff
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:="|", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        Columns("A:B").Select
        Selection.Delete Shift:=xlToLeft
        Columns("B:D").Select
        Selection.Delete Shift:=xlToLeft
        Rows("1:1").Select
        Selection.Delete Shift:=xlUp
        Rows("2:2").Select
        Selection.Delete Shift:=xlUp
        Call TrimAllCells
        Call SaveCurrentSheet
    End If
    Sheets("Settings").Select
    Sheets("Settings").Range("AG1:AG1000") = Sheets("MD07").Range("A1:A1000").Value
End Sub
Sub btnAddEditOperations_Click()
    Sheets("AllOperations").Select
    Range("A1").End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
End Sub
'Sub btnResetRouting_Click()
'    End
'    Dim MyRow As Long
'    Dim I As Long
'    Dim CurOperation As String
'    If Sheets("MakeTree").Range("A2") = Sheets("MaterialRouting").Range("Q1") Then
'        Sheets("MaterialRouting").Range("S1") = Sheets("MakeTree").Range("D2")
'        Sheets("MaterialRouting").Range("T1") = Sheets("MakeTree").Range("Q2")
'    Else
'        MyRow = WorksheetFunction.Match(Sheets("MaterialRouting").Range("P1"), Sheets("MakeTree").Range("AD:AD"), 0)
'        Sheets("MaterialRouting").Range("S1") = Sheets("MakeTree").Range("U" & MyRow)
'        Sheets("MaterialRouting").Range("T1") = Sheets("MakeTree").Range("Z" & MyRow)
'        MyRow = 0
'    End If
'    Application.EnableEvents = False
'    If Sheets("MaterialRouting").Range("R1") <> "Not in routing database" Then
'        Sheets("MaterialRouting").Range("A2:O5000") = ""
'        MyRow = WorksheetFunction.Match(Sheets("MaterialRouting").Range("R1"), Sheets("RoutingDatabase").Range("A:A"), 0)
'        While Sheets("RoutingDatabase").Range("B" & MyRow).Offset(I, 0) <> "9999"
'            Sheets("MaterialRouting").Range("A2").Offset(I, 0) = Sheets("RoutingDatabase").Range("B" & MyRow).Offset(I, 0)
'            Sheets("MaterialRouting").Range("B2").Offset(I, 0) = Sheets("RoutingDatabase").Range("C" & MyRow).Offset(I, 0)
'            Sheets("MaterialRouting").Range("C2").Offset(I, 0) = Sheets("RoutingDatabase").Range("D" & MyRow).Offset(I, 0)
'            Sheets("MaterialRouting").Range("D2").Offset(I, 0) = Sheets("RoutingDatabase").Range("E" & MyRow).Offset(I, 0)
'            Sheets("MaterialRouting").Range("E2").Offset(I, 0) = Sheets("MaterialRouting").Range("S1")
'            Sheets("MaterialRouting").Range("F2").Offset(I, 0) = Sheets("MaterialRouting").Range("T1")
'            Sheets("MaterialRouting").Range("G2").Offset(I, 0) = Sheets("RoutingDatabase").Range("H" & MyRow).Offset(I, 0)
'            Sheets("MaterialRouting").Range("H2").Offset(I, 0) = Sheets("RoutingDatabase").Range("I" & MyRow).Offset(I, 0)
'            Sheets("MaterialRouting").Range("I2").Offset(I, 0) = Sheets("RoutingDatabase").Range("J" & MyRow).Offset(I, 0)
'            Sheets("MaterialRouting").Range("J2").Offset(I, 0) = Sheets("RoutingDatabase").Range("K" & MyRow).Offset(I, 0)
'            Sheets("MaterialRouting").Range("K2").Offset(I, 0) = Sheets("RoutingDatabase").Range("L" & MyRow).Offset(I, 0)
'            Sheets("MaterialRouting").Range("L2").Offset(I, 0) = Sheets("RoutingDatabase").Range("M" & MyRow).Offset(I, 0)
'            Sheets("MaterialRouting").Range("M2").Offset(I, 0) = Sheets("RoutingDatabase").Range("N" & MyRow).Offset(I, 0)
'            Sheets("MaterialRouting").Range("N2").Offset(I, 0) = Sheets("RoutingDatabase").Range("O" & MyRow).Offset(I, 0)
'            Sheets("MaterialRouting").Range("O2").Offset(I, 0) = Sheets("RoutingDatabase").Range("P" & MyRow).Offset(I, 0)
'            I = I + 1
'            CurOperation = Sheets("RoutingDatabase").Range("B" & MyRow).Offset(I, 0)
'            DoEvents
'        Wend
'        Sheets("MaterialRouting").Range("A2").Offset(I, 0) = Sheets("RoutingDatabase").Range("B" & MyRow).Offset(I, 0)
'        Sheets("MaterialRouting").Range("B2").Offset(I, 0) = Sheets("RoutingDatabase").Range("C" & MyRow).Offset(I, 0)
'        Sheets("MaterialRouting").Range("C2").Offset(I, 0) = Sheets("RoutingDatabase").Range("D" & MyRow).Offset(I, 0)
'        Sheets("MaterialRouting").Range("D2").Offset(I, 0) = Sheets("RoutingDatabase").Range("E" & MyRow).Offset(I, 0)
'        Sheets("MaterialRouting").Range("E2").Offset(I, 0) = Sheets("MaterialRouting").Range("S1")
'        Sheets("MaterialRouting").Range("F2").Offset(I, 0) = Sheets("MaterialRouting").Range("T1")
'        Sheets("MaterialRouting").Range("G2").Offset(I, 0) = Sheets("RoutingDatabase").Range("H" & MyRow).Offset(I, 0)
'        Sheets("MaterialRouting").Range("H2").Offset(I, 0) = Sheets("RoutingDatabase").Range("I" & MyRow).Offset(I, 0)
'        Sheets("MaterialRouting").Range("I2").Offset(I, 0) = Sheets("RoutingDatabase").Range("J" & MyRow).Offset(I, 0)
'        Sheets("MaterialRouting").Range("J2").Offset(I, 0) = Sheets("RoutingDatabase").Range("K" & MyRow).Offset(I, 0)
'        Sheets("MaterialRouting").Range("K2").Offset(I, 0) = Sheets("RoutingDatabase").Range("L" & MyRow).Offset(I, 0)
'        Sheets("MaterialRouting").Range("L2").Offset(I, 0) = Sheets("RoutingDatabase").Range("M" & MyRow).Offset(I, 0)
'        Sheets("MaterialRouting").Range("M2").Offset(I, 0) = Sheets("RoutingDatabase").Range("N" & MyRow).Offset(I, 0)
'        Sheets("MaterialRouting").Range("N2").Offset(I, 0) = Sheets("RoutingDatabase").Range("O" & MyRow).Offset(I, 0)
'        Sheets("MaterialRouting").Range("O2").Offset(I, 0) = Sheets("RoutingDatabase").Range("P" & MyRow).Offset(I, 0)
'    End If
'    Application.EnableEvents = True
'End Sub
'Sub SaveMaterialRoutingDatabase()
'    Application.DisplayAlerts = False
'    Workbooks.Open ThisWorkbook.Path & "\MaterialRoutingDatabase.xlsx"
'    Sheets(Array("Routing2MatNum", "RoutingDatabase")).Delete
'    Windows("MakeList&RouteMaker.xlsm").Activate
'    Sheets(Array("Routing2MatNum", "RoutingDatabase")).Copy Before:=Workbooks("MaterialRoutingDatabase.xlsx").Sheets(1)
'    Windows("MaterialRoutingDatabase.xlsx").Close True
'End Sub
Sub ResetOrderRoutingToMaterialRouting(Optional ProductionOrderNumber As String)
    'Open order and run update to PP master
    If Sheets("Settings").Range("B13") = "Production" Then Call LogonProduction Else Call LogonDevelopment
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nco02"
    Session.FindById("wnd[0]").SendVKey 0
    'get production order number for the material displayed
    If Sheets("MaterialRouting").Range("P1") = "Displaying Material Routing" Then
        If Sheets("MakeTree").Range("A2") = Sheets("MaterialRouting").Range("Q1") Then Sheets("MaterialRouting").Range("P1") = Sheets("MakeTree").Range("T2")
    End If
    If ProductionOrderNumber = "" Then ProductionOrderNumber = Sheets("MaterialRouting").Range("P1")
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = ProductionOrderNumber 'Sheets("MaterialRouting").Range("P1") '"1043566"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/mbar/menu[1]/menu[6]").Select
    Session.FindById("wnd[1]/usr/btnSPOP-VAROPTION2").press
    Session.FindById("wnd[1]/usr/chkRC62F-NEW_BOM").Selected = False
    Session.FindById("wnd[1]/usr/chkRC62F-NEW_ROUT").Selected = True
    Session.FindById("wnd[1]/usr/ctxtRC62F-PLAUF").Text = Date '"10/14/2014"
    Session.FindById("wnd[1]/usr/ctxtRC62F-AUFLD").Text = Date '"10/14/2014"
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    Session.FindById("wnd[0]/tbar[0]/btn[11]").press
    Call LogOff
    Call CO02_GetProductionOrderRouting(ProductionOrderNumber)
End Sub
Sub ResetOrderBOMToMaterialBOM(Optional ProductionOrderNumber As String)
    If Sheets("Settings").Range("B13") = "Production" Then Call LogonProduction Else Call LogonDevelopment
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nco02"
    Session.FindById("wnd[0]").SendVKey 0
    If ProductionOrderNumber = "" Then ProductionOrderNumber = Sheets("ModifyBOM").Range("J2")
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = ProductionOrderNumber 'Sheets("MaterialRouting").Range("P1") '"1043566"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/mbar/menu[1]/menu[6]").Select
    Session.FindById("wnd[1]/usr/btnSPOP-VAROPTION2").press
    Session.FindById("wnd[1]/usr/ctxtRC62F-AUFLD").Text = Date '"10/14/2014"
    Session.FindById("wnd[1]/usr/chkRC62F-NEW_ROUT").Selected = False
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    Session.FindById("wnd[0]/tbar[0]/btn[11]").press
    Call LogOff
    Call CO02_GetProductionOrderBOM(Sheets("ModifyBOM").Range("J2"))
End Sub
Sub SaveOrderRouting()
    Dim MyOrderNum As String
    Dim CurOperation As String
    Dim NextOperation As String
    Dim SAPMessage As String
    Dim MyCnt As Integer
    Dim MyFoundRow As Integer
    Dim MyOShortText As String
    Dim MyString As String
    Dim MyVal As Integer
    Dim MyRow As Integer
    Dim MyMat As String
    Dim MyMat2 As String
    Dim I As Integer
    Application.DisplayAlerts = False
    'get production order number for the material displayed
    If Sheets("MaterialRouting").Range("P1") = "Displaying Material Routing" Then
        MyMat = Trim(Sheets("MakeTree").Range("A2"))
        MyMat2 = Trim(Sheets("MaterialRouting").Range("Q1"))
        If MyMat = MyMat2 Then Sheets("MaterialRouting").Range("P1") = Sheets("MakeTree").Range("T2")
    End If
    MyVal = WorksheetFunction.CountIf(Range("C:C"), "ZP03")
    If MyVal <> 1 Then MsgBox ("Make sure you only have one ZP03 specified in your routing."): End
    I = 2
    While Range("B" & I) <> ""
        On Error Resume Next
        MyVal = 0
        MyVal = WorksheetFunction.Match(Range("B" & I), Sheets("WorkCenters").Range("A:A"), 0)
        On Error GoTo 0
        If MyVal = 0 Then MsgBox ("Not a valid work center in SAP"): Range("B" & I).Select: End
        I = I + 1
        DoEvents
    Wend
    For I = 1 To 11
        MyCnt = WorksheetFunction.CountA(ThisWorkbook.Sheets("MaterialRouting").Range("A:A").Offset(0, I - 1))
        If MyCnt <> WorksheetFunction.CountA(ThisWorkbook.Sheets("MaterialRouting").Range("A:A").Offset(0, I)) Then MsgBox ("Blank in column " & I + 1): ThisWorkbook.Sheets("MaterialRouting").Range("A:A").Offset(0, I).Select: End
    Next I
    'check for existanse of IS on steps with adsorb and affinity
    Range("N2").Select
    While Range("A" & ActiveCell.Row) <> ""
        If InStr(Range("D" & ActiveCell.Row), "Affinity") <> 0 And Range("N" & ActiveCell.Row) = "" Then MsgBox ("Add IS"): Range("N" & ActiveCell.Row).Select: End
        If InStr(Range("D" & ActiveCell.Row), "Adsorb") <> 0 And Range("N" & ActiveCell.Row) = "" Then MsgBox ("Add IS"): Range("N" & ActiveCell.Row).Select: End
        ActiveCell.Offset(1, 0).Select
        DoEvents
    Wend
    MyOrderNum = Range("P1")
    If Range("A2") <> "" And MyOrderNum <> "" Then
        If Sheets("Settings").Range("B13") = "Production" Then Call LogonProduction Else Call LogonDevelopment
        Stop
        Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nco02"
        Session.FindById("wnd[0]").SendVKey 0
        Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = MyOrderNum
        Session.FindById("wnd[0]").SendVKey 0
        Session.FindById("wnd[0]/tbar[1]/btn[5]").press
        I = 0
        MyString = ""
        On Error Resume Next
        MyString = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8," & I & "]").Text
        On Error GoTo 0
        While MyString <> ""
            MyVal = Val(Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0," & I & "]").Text)
            MyRow = 0
            On Error Resume Next
            MyRow = WorksheetFunction.Match(MyVal, Sheets("MaterialRouting").Range("A:A"), 0)
            On Error GoTo 0
            If MyRow = 0 Then
                'Delete operations that are not in excel
                SAPMessage = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VSTTXT[10," & I & "]").Text
                If InStr(SAPMessage, "DLT") = 0 Then
                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100").GetAbsoluteRow(I).Selected = True
                    Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnDELETE").press
                    Session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").press
                    Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnDESELECT").press
                End If
            End If
            I = I + 1
            MyString = ""
            On Error Resume Next
            MyString = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8," & I & "]").Text
            On Error GoTo 0
            DoEvents
        Wend
        SAPMessage = ""
'        CurOperation = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0,0]").Text
'        NextOperation = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0,1]").Text
'        If CurOperation = "0001" And NextOperation = "9999" Then
'            Range("A3").Select
'            I = 1
'            While Range("A" & ActiveCell.Row) <> ""
'                If Range("A" & ActiveCell.Row) <> 9999 Then
'
'                    Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnINSERT").press
'                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0," & I & "]").Text = Format(Range("A" & ActiveCell.Row), "0000")
'                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-ARBPL[4," & I & "]").Text = Range("B" & ActiveCell.Row) '"3600"
'                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-STEUS[6," & I & "]").Text = Range("C" & ActiveCell.Row) '"ZP01"
'                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8," & I & "]").Text = Range("D" & ActiveCell.Row) '"IEP, operation testing"
'                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-KTSCH[7," & I & "]").Text = Trim(Range("M" & ActiveCell.Row)) '"AB"
'                    Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnSHOWDETAIL").press
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW").Select
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-BMSCH").Text = Range("E" & ActiveCell.Row) '"Base Qty"
''                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-MEINH").Text = Range("F" & ActiveCell.Row)
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW01").Text = Range("G" & ActiveCell.Row) '"1"
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE01").Text = Range("H" & ActiveCell.Row) '"D"
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW02").Text = Range("I" & ActiveCell.Row) '"0"
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE02").Text = Range("J" & ActiveCell.Row) '"H"
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW03").Text = Range("K" & ActiveCell.Row) '"0.5"
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE03").Text = Range("L" & ActiveCell.Row) '"H"
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE03").SetFocus
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE03").CaretPosition = 1
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpALAV").Select
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpALAV").Select
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpALAV").Select
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpALAV").Select
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpALAV").Select
'                    Session.FindById("wnd[0]/tbar[1]/btn[5]").press
'                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-KTSCH[7," & I & "]").Text = Trim(Range("M" & ActiveCell.Row)) '"AB"
'                    If Range("N" & ActiveCell.Row) <> "" Then
'                        Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/chkAFVGD-FLG_FHM[13," & I & "]").SetFocus
'                        Session.FindById("wnd[0]").SendVKey 2
'                        Session.FindById("wnd[1]/usr/ctxtAFFHD-MATNR").Text = Range("N" & ActiveCell.Row) '"ISAD100-205"
'                        Session.FindById("wnd[1]/usr/ctxtAFFHD-STEUF").Text = "1"
'                        Session.FindById("wnd[1]/tbar[0]/btn[3]").press
'                        Session.FindById("wnd[0]/tbar[0]/btn[3]").press
'                    End If
'                    If Range("O" & ActiveCell.Row) <> "" Then
'                        Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnSHOWDETAIL").press
'                        Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpTEXT").Select
'                        Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpTEXT/ssubSUBSCR_0101:SAPLCOVF:0230/cntlTEXTEDITOR_COVF/shell").Text = Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpTEXT/ssubSUBSCR_0101:SAPLCOVF:0230/cntlTEXTEDITOR_COVF/shell").Text + vbCr + Range("O" & ActiveCell.Row) + vbCr + ""
'                        Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpTEXT/ssubSUBSCR_0101:SAPLCOVF:0230/cntlTEXTEDITOR_COVF/shell").SetSelectionIndexes 11, 11
'                        Session.FindById("wnd[0]/tbar[0]/btn[3]").press
'                    End If
'                End If
'                DoEvents
'                I = I + 1
'                ActiveCell.Offset(1, 0).Select
'            Wend
'            Session.FindById("wnd[0]/tbar[0]/btn[11]").press
'            SAPMessage = Session.FindById("wnd[0]/sbar").Text
'            Call LogOff
'            MsgBox SAPMessage
'        Else
            Range("A2").Select
            CurOperation = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0,0]").Text
            While Range("A" & ActiveCell.Row) <> ""
                I = 0
                MyFoundRow = 0
                On Error Resume Next
                MyOShortText = ""
                MyOShortText = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8," & I & "]").Text
                If Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0," & I & "]").Text <> Format(Range("A" & ActiveCell.Row), "0000") Then
                    Call FindOperation(Format(Range("A" & ActiveCell.Row), "0000"))
'                    While MyOShortText <> ""
'                        If Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0," & I & "]").Text = Format(Range("A" & ActiveCell.Row), "0000") And MyFoundRow = 0 Then MyFoundRow = I
'                        I = I + 1
'                        MyOShortText = ""
'                        MyOShortText = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8," & I & "]").Text
'                        DoEvents
'                    Wend
                End If
                On Error GoTo 0
                Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnDESELECT").press
                If Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0," & MyFoundRow & "]").Text <> Format(Range("A" & ActiveCell.Row), "0000") Then
                    'Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100").GetAbsoluteRow(I).Selected = True
                    'Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0," & MyFoundRow & "]").Select
                    'Session.FindById("wnd[0]").SendVKey 2
                    Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnDESELECT").press
                    Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnINSERT").press
                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0," & I & "]").Text = Format(Range("A" & ActiveCell.Row), "0000")
                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-ARBPL[4," & I & "]").Text = Range("B" & ActiveCell.Row) '"3600"
                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-STEUS[6," & I & "]").Text = Range("C" & ActiveCell.Row) '"ZP01"
                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8," & I & "]").Text = Range("D" & ActiveCell.Row) '"IEP, operation testing"
                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-KTSCH[7," & I & "]").Text = Trim(Range("M" & ActiveCell.Row)) '"AB"
                    Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnSHOWDETAIL").press
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW").Select
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-BMSCH").Text = Range("E" & ActiveCell.Row) '"Base Qty"
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-MEINH").Text = Range("F" & ActiveCell.Row)
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW01").Text = Range("G" & ActiveCell.Row) '"1"
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE01").Text = Range("H" & ActiveCell.Row) '"D"
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW02").Text = Range("I" & ActiveCell.Row) '"0"
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE02").Text = Range("J" & ActiveCell.Row) '"H"
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW03").Text = Range("K" & ActiveCell.Row) '"0.5"
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE03").Text = Range("L" & ActiveCell.Row) '"H"
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE03").SetFocus
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE03").CaretPosition = 1
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpALAV").Select
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpALAV").Select
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpALAV").Select
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpALAV").Select
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpALAV").Select
                    Session.FindById("wnd[0]/tbar[1]/btn[5]").press
                    I = 0
                    MyFoundRow = 0
                    On Error Resume Next
                    MyOShortText = ""
                    MyOShortText = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8," & I & "]").Text
                    On Error GoTo 0
                    Call FindOperation(Format(Range("A" & ActiveCell.Row), "0000"))
'                    While MyOShortText <> ""
'                        If Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0," & I & "]").Text = Format(Range("A" & ActiveCell.Row), "0000") And MyFoundRow = 0 Then MyFoundRow = I
'                        I = I + 1
'                        MyOShortText = Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8," & I & "]").Text
'                        DoEvents
'                    Wend
                    Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnDESELECT").press
                    Stop
                    If Range("O" & ActiveCell.Row) <> "" Then
                        'Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnSHOWDETAIL").press
                        Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/chkRC270-TXTKZ[9,0]").SetFocus
                        Session.FindById("wnd[0]").SendVKey 2
'                        Session.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select
                        Session.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,2]").Text = Range("O" & ActiveCell.Row)
                        Session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    End If
                    If Range("N" & ActiveCell.Row) <> "" Then
                        Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/chkAFVGD-FLG_FHM[13," & MyFoundRow & "]").SetFocus
                        Session.FindById("wnd[0]").SendVKey 2
                        'Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnFHM").press
                        Session.FindById("wnd[1]/usr/ctxtAFFHD-MATNR").Text = Range("N" & ActiveCell.Row) '"ISAD100-205"
                        Session.FindById("wnd[1]/usr/ctxtAFFHD-STEUF").Text = "1"
                        Session.FindById("wnd[1]/tbar[0]/btn[3]").press
                        Session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    End If
                Else
                    'found row you need to edit
                    Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnDESELECT").press
                    'Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100").GetAbsoluteRow(MyFoundRow).Selected = True
                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VORNR[0," & MyFoundRow & "]").Text = Format(Range("A" & ActiveCell.Row), "0000")
                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-ARBPL[4," & MyFoundRow & "]").Text = Range("B" & ActiveCell.Row) '"3600"
                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-STEUS[6," & MyFoundRow & "]").Text = Range("C" & ActiveCell.Row) '"ZP01"
                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[8," & MyFoundRow & "]").Text = Range("D" & ActiveCell.Row) '"IEP, operation testing"
                    Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-KTSCH[7," & MyFoundRow & "]").Text = Trim(Range("M" & ActiveCell.Row)) '"Step Operator"
                    Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnSHOWDETAIL").press
                    On Error Resume Next
                    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
                    On Error GoTo 0
                    Session.FindById("wnd[0]").SendVKey 2
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW").Select
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-BMSCH").Text = Range("E" & ActiveCell.Row) '"Base Qty"
'                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-MEINH").Text = Range("F" & ActiveCell.Row)
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW01").Text = Range("G" & ActiveCell.Row) '"1"
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE01").Text = Range("H" & ActiveCell.Row) '"D"
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW02").Text = Range("I" & ActiveCell.Row) '"0"
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE02").Text = Range("J" & ActiveCell.Row) '"H"
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/txtAFVGD-VGW03").Text = Range("K" & ActiveCell.Row) '"0.5"
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpVOGW/ssubSUBSCR_0101:SAPLCOVF:0130/ctxtAFVGD-VGE03").Text = Range("L" & ActiveCell.Row) '"H"
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpALAV").Select
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpALAV").Select
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpALAV").Select
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpALAV").Select
                    Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpALAV").Select
                    Session.FindById("wnd[0]/tbar[1]/btn[5]").press
                    If Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/chkRC270-TXTKZ[9," & MyFoundRow & "]").Selected = True And Range("O" & ActiveCell.Row) = "" Then
                        'note has been deleted
                        Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnLONG_TEXT").press
                        Session.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select
                        Session.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,2]").Text = ""
                        Session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    End If
                    If Range("O" & ActiveCell.Row) <> "" Then
                        'note in routing step
                        Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnSHOWDETAIL").press
                        Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpTEXT").Select
                        Session.FindById("wnd[0]/usr/subTAB_SUB_SCREEN:SAPLCOVF:0101/tabsTABSTRIP_0100/tabpTEXT/ssubSUBSCR_0101:SAPLCOVF:0230/cntlTEXTEDITOR_COVF/shell").Text = Range("D" & ActiveCell.Row) + vbCr + Range("O" & ActiveCell.Row)
                        Session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    End If
                    If Range("N" & ActiveCell.Row) <> "" Then
                        Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnDESELECT").press
                        Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/chkAFVGD-FLG_FHM[13," & MyFoundRow & "]").SetFocus
                        Session.FindById("wnd[0]").SendVKey 2
                        'Session.FindById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100").GetAbsoluteRow(MyFoundRow).Selected = True
                        'Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOVG:0050/btnFHM").press
                        If Session.FindById("wnd[0]/usr/tblSAPLCOFUTCTRL_0100/txtAFFHD-FHMNR[2,0]").Text <> Range("N" & ActiveCell.Row) Then
                            'New IS delete old and add new
                            Dim NextOp As Integer
                            Dim MaxOperation As Integer
                            Dim J As Integer
                            NextOp = Val(Session.FindById("wnd[0]/usr/tblSAPLCOFUTCTRL_0100/txtAFFHD-PSNFH[0,0]").Text)
                            While NextOp <> 0
                                If NextOp > MaxOperation Then MaxOperation = NextOp
                                DoEvents
                                J = J + 1
                                NextOp = 0
                                NextOp = Val(Session.FindById("wnd[0]/usr/tblSAPLCOFUTCTRL_0100/txtAFFHD-PSNFH[0," & J & "]").Text)
                            Wend
                            MaxOperation = MaxOperation + 10
                            Session.FindById("wnd[0]/usr/tblSAPLCOFUTCTRL_0100/txtAFFHD-PSNFH[0,0]").Text = MaxOperation
                            Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOFU:0050/btnINSERT").press
                            Session.FindById("wnd[1]/usr/txtAFFHD-PSNFH").Text = "0010"
                            Session.FindById("wnd[1]/usr/ctxtAFFHD-MATNR").Text = Range("N" & ActiveCell.Row)
                            Session.FindById("wnd[1]/usr/ctxtAFFHD-STEUF").Text = "1"
                            Session.FindById("wnd[1]/tbar[0]/btn[3]").press
                            J = 0
                            NextOp = Val(Session.FindById("wnd[0]/usr/tblSAPLCOFUTCTRL_0100/txtAFFHD-PSNFH[0," & J & "]").Text)
                            While NextOp <> MaxOperation
                                'If NextOp > MaxOperation Then MaxOperation = NextOp
                                DoEvents
                                J = J + 1
                                NextOp = 0
                                NextOp = Val(Session.FindById("wnd[0]/usr/tblSAPLCOFUTCTRL_0100/txtAFFHD-PSNFH[0," & J & "]").Text)
                            Wend
                            Session.FindById("wnd[0]/usr/tblSAPLCOFUTCTRL_0100").GetAbsoluteRow(J).Selected = True
                            Session.FindById("wnd[0]/usr/subBUTTONS:SAPLCOFU:0050/btnDELETE").press
                            Session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").press
                        End If
                        Session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    End If
                End If
                DoEvents
                ActiveCell.Offset(1, 0).Select
            Wend
            Session.FindById("wnd[0]/tbar[0]/btn[11]").press
            On Error Resume Next
            Session.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
            On Error GoTo 0
            Call LogOff
'        End If
    End If
End Sub
Sub FindOperation(Operation As String)
    Session.FindById("wnd[0]/tbar[0]/btn[71]").press
    Session.FindById("wnd[1]/usr/txtRCOSU-VORNR").Text = Operation '"0010"
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    On Error Resume Next
    Session.FindById("wnd[2]/tbar[0]/btn[0]").press
    On Error GoTo 0
End Sub
