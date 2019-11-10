Attribute VB_Name = "Code"
Option Explicit
Public RunTwice As Boolean
Sub WipeData()
    Sheets("Main").Select
    Sheets("Main").Range("AA2:AQ10000") = ""
    Sheets("Main").Range("AA2:AQ10000").Borders(xlDiagonalDown).LineStyle = xlNone
    Sheets("Main").Range("AA2:AQ10000").Borders(xlDiagonalUp).LineStyle = xlNone
    Sheets("Main").Range("AA2:AQ10000").Borders(xlEdgeLeft).LineStyle = xlNone
    Sheets("Main").Range("AA2:AQ10000").Borders(xlEdgeTop).LineStyle = xlNone
    Sheets("Main").Range("AA2:AQ10000").Borders(xlEdgeBottom).LineStyle = xlNone
    Sheets("Main").Range("AA2:AQ10000").Borders(xlEdgeRight).LineStyle = xlNone
    Sheets("Main").Range("AA2:AQ10000").Borders(xlInsideVertical).LineStyle = xlNone
    Sheets("Main").Range("AA2:AQ10000").Borders(xlInsideHorizontal).LineStyle = xlNone
    With Sheets("Main").Range("AA2:AQ10000").Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Sheets("ZVialOrders").Range("A2:Z10000") = ""
    Sheets("BOMChild").Range("A2:Z10000") = ""
    Sheets("MB52").Range("A2:K10000") = ""
    Sheets("ZCharValues").Range("A2:K10000") = ""
    ThisWorkbook.Save
End Sub
Sub testZ4()
    Call ReleaseSalesOrder("111927")
End Sub
Function ReleaseSalesOrder(SalesOrderNum As String)
    Dim MyVal As String, i As Integer, j As Integer
    Dim SendAgain As Boolean
    SendAgain = True
    UserName = Range("ZZ1")
    Password = Range("AAA1")
    If Sheets("Main").Range("AA1") = "Production" Then Call LogonProduction Else Call LogonDevelopment
retry:
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nVA02"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ctxtVBAK-VBELN").Text = SalesOrderNum '"79134"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]").SendVKey 0 'probably not necessary
    If Len(SalesOrderNum) = 5 Or Len(SalesOrderNum) = 6 Then
        'handle when someone else is in the order
        If InStr(Session.FindById("wnd[0]/sbar").Text, "currently being processed") <> 0 Then
            If MsgBox(Session.FindById("wnd[0]/sbar").Text & vbCrLf & vbCrLf & "Please advise the user to exit the order and click OK to proceed." & vbCrLf & vbCrLf & "Click cancel to exit the order-release process (no changes to the order documentation will be made).", vbOKCancel) = vbOK Then
                'notify user that is viewing order to close out and proceed with the routine as usual
            Else
                ReleaseSalesOrder = "Order Locked"
                Exit Function
            End If
        End If
        'the following patch was added to try to get the shipping document to not show that item is backordered when it isn't (don't fully understand why this happens)
        'Bokanyi made some changes to tweak the layout of the report and now this needs to be done to prevent material from being labeled as backordered...mbb 20160614
        Session.FindById("wnd[0]/mbar/menu[1]/menu[10]").Select 'AV CHECK
        Do
            On Error Resume Next
            Session.FindById("wnd[0]/tbar[1]/btn[18]").press 'generates error 619
        Loop Until Err.Number <> 0
        Do
            On Error Resume Next
            Session.FindById("wnd[0]/usr/btnBUT1").press 'generates error 619
        Loop Until Err.Number <> 0
        On Error GoTo 0

        If InStr(Session.FindById("wnd[0]/sbar").Text, "saved") = 0 Then Session.FindById("wnd[0]/tbar[0]/btn[11]").press  'save
        'If InStr(Session.FindById("wnd[0]/sbar").Text, "saved") = 0 Then MsgBox ("There is a payment authorization issue with " & SalesOrderNum & ". Contact Mike to modify program.")

        MyVal = ""
        On Error Resume Next
        MyVal = Session.FindById("wnd[1]/usr/txtMESSTXT1").Text
        Application.Wait (Now + TimeValue("0:00:01"))
        If InStr(MyVal, "requires complete delivery") > 0 Then
            Session.FindById("wnd[1]").SendVKey 0 'hit the green check
            Application.Wait (Now + TimeValue("0:00:01"))
            'goes back to initial screen
        End If
        MyVal = ""
        On Error GoTo 0

        On Error Resume Next
        MyVal = Session.FindById("wnd[1]/usr/txtMESSTXT1").Text
        Application.Wait (Now + TimeValue("0:00:01"))
        If InStr(MyVal, "Oldest of open") > 0 Then
            Session.FindById("wnd[1]").SendVKey 0 'hit the green check
            Application.Wait (Now + TimeValue("0:00:01"))
            'goes back to initial screen
        End If
        MyVal = ""
        On Error GoTo 0

        On Error Resume Next
        MyVal = Session.FindById("wnd[1]/usr/txtMESSTXT1").Text
        Application.Wait (Now + TimeValue("0:00:01"))
        If InStr(MyVal, "The system has set the authorization block") > 0 Then
            Session.FindById("wnd[1]").SendVKey 0
        End If
        MyVal = ""
        On Error GoTo 0

        Application.Wait (Now + TimeValue("0:00:01"))
        If InStr(Session.FindById("wnd[0]/sbar").Text, "saved") <> 0 Then ReleaseSalesOrder = Session.FindById("wnd[0]/sbar").Text
        'If InStr(Session.FindById("wnd[0]/sbar").Text, "saved") <> 0 Then ReleaseSalesOrder = Session.FindById("wnd[0]/sbar").Text: Exit Function

        'handle finance messages
        'handle no response from clearing house
        On Error Resume Next
        MyVal = ""
        MyVal = Session.FindById("wnd[1]/usr/txtSPOP-TEXTLINE1").Text
        If InStr(MyVal, "No response from clearing house") <> 0 Then
            Session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").press
            Session.FindById("wnd[1]/tbar[0]/btn[0]").press
            'If InStr(Session.FindById("wnd[0]/sbar").Text, "saved") <> 0 Then ReleaseSalesOrder = Session.FindById("wnd[0]/sbar").Text: Exit Function
        End If
        'handle authorization was unsuccessful issue
        MyVal = ""
        MyVal = Session.FindById("wnd[1]/usr/txtSPOP-TEXTLINE1").Text
        If InStr(MyVal, "Authorization was unsuccessful") <> 0 Then
            Session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").press
            Session.FindById("wnd[1]/tbar[0]/btn[0]").press
            'If InStr(Session.FindById("wnd[0]/sbar").Text, "saved") <> 0 Then ReleaseSalesOrder = Session.FindById("wnd[0]/sbar").Text: Exit Function
        End If
        Session.FindById("wnd[1]/tbar[0]/btn[0]").press
        Session.FindById("wnd[1]/tbar[0]/btn[0]").press
        Application.Wait (Now + TimeValue("0:00:01"))
        'is this necessary?
        If InStr(Session.FindById("wnd[1]/usr/txtMESSTXT1").Text, "Oldest of open") > 0 Then
            If Err.Number = 0 Then Session.FindById("wnd[1]").SendVKey 0
        End If

        On Error GoTo 0
        Application.Wait (Now + TimeValue("0:00:01"))
        'enter order again and continue with removal of "waiting on inventory" statement (if it's there)
        Session.FindById("wnd[0]").SendVKey 0
        Session.FindById("wnd[0]").SendVKey 0 'to get rid of "Consider Subsequent documents" statement when it appears (does nothing if the dialog does not appear)
        MyVal = ""
        MyVal = Trim(Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").Text)
        'MyVal will be empty if there is nothing in the Delivery Block box
        'handle when SD is not in DB
        If InStr(MyVal, "Waiting on Inventory") > 0 Or InStr(MyVal, "Z4") > 0 Then
            'iterate through the items until an error occurs, indicating that the last item has been reached
            On Error Resume Next
            i = 0
            While Err.Number = 0
                Session.FindById("wnd[0]/tbar[1]/btn[18]").press 'click continue button
                i = i + 1 ' count if items in list??
                DoEvents
            Wend
            On Error GoTo 0
            'iterate through the items until an error occurs, indicating that the last item has been reached
            On Error Resume Next
            While Err.Number = 0
                Session.FindById("wnd[1]/tbar[0]/btn[0]").press 'click green check button
                DoEvents
            Wend
            On Error GoTo 0
            'if the "Authorization was unsuccessful" dialog appears, then click through it
            MyVal = ""
            On Error Resume Next
            MyVal = Session.FindById("wnd[1]/usr/txtSPOP-TEXTLINE1").Text
            If InStr(MyVal, "Authorization was unsuccessful") > 0 Then
                Session.FindById("wnd[1]").SendVKey 12
                Session.FindById("wnd[1]").SendVKey 0
            End If
            On Error GoTo 0
            MyVal = ""
            MyVal = Trim(Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").Text) 'grab value from field that contains "Waiting on Inventory"
            If MyVal <> "" And i < 2 Then
                While InStr(MyVal, "Waiting") > 0 Or InStr(MyVal, "Z4") > 0  'loop until Z4 is properly removed
                    Application.Wait (Now + TimeValue("0:00:02"))
                    Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").Key = " " 'clears Delivery Block
                    Application.Wait (Now + TimeValue("0:00:02"))
                    MyVal = ""
                    MyVal = Trim(Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").Text)
                    If InStr(MyVal, "Waiting") > 0 Then Stop
                Wend
                Session.FindById("wnd[0]/tbar[0]/btn[11]").press 'save
                If InStr(Session.FindById("wnd[0]/sbar").Text, "saved") <> 0 Then ReleaseSalesOrder = Session.FindById("wnd[0]/sbar").Text

                'save fails if there is an authorization block
                'handle "the system has set the authorization block"
                MyVal = ""
                On Error Resume Next
                MyVal = Session.FindById("wnd[1]/usr/txtMESSTXT1").Text
                If InStr(MyVal, "The system has set the authorization block") > 0 Then
                    'Session.FindById("wnd[1]").SendVKey 12
                    Session.FindById("wnd[1]").SendVKey 0
                End If
                On Error GoTo 0
                MyVal = ""
                'handle "oldest of open items overdue"
                On Error Resume Next
                MyVal = Session.FindById("wnd[1]/usr/txtMESSTXT1").Text
                If InStr(MyVal, "Oldest of open") > 0 Then
                    Session.FindById("wnd[1]").SendVKey 0
                End If
                On Error GoTo 0
                'try saving again
                If InStr(ReleaseSalesOrder, "saved") = 0 Then Session.FindById("wnd[0]/tbar[0]/btn[11]").press
                If InStr(Session.FindById("wnd[0]/sbar").Text, "saved") <> 0 Then ReleaseSalesOrder = Session.FindById("wnd[0]/sbar").Text
                If InStr(Session.FindById("wnd[0]/sbar").Text, "Choose a valid function") > 0 Then 'save was successful..reattempting generated the "choose a valid function" message...keep previous releasesalesorder string
                Else
                    If InStr(ReleaseSalesOrder, "saved") = 0 Then ReleaseSalesOrder = Session.FindById("wnd[0]/sbar").Text
                End If
                If InStr(ReleaseSalesOrder, "saved") = 0 Then
                    j = j + 1
                    If j = 10 Then MsgBox "After 10 tries, the order-release won't save. Call Mike."
                    GoTo retry
                End If
                j = 0
            ElseIf MyVal <> "" And i > 1 Then
                Session.FindById("wnd[0]/tbar[0]/btn[11]").press 'save
                If InStr(Session.FindById("wnd[0]/sbar").Text, "saved") <> 0 Then ReleaseSalesOrder = Session.FindById("wnd[0]/sbar").Text

                On Error Resume Next
                MyVal = ""
                MyVal = Session.FindById("wnd[1]/usr/txtMESSTXT1").Text
                If InStr(MyVal, "Oldest of open") > 0 Then
                    Session.FindById("wnd[1]").SendVKey 0
                End If
                On Error GoTo 0
                'try saving again
                If InStr(ReleaseSalesOrder, "saved") = 0 Then Session.FindById("wnd[0]/tbar[0]/btn[11]").press
                If InStr(Session.FindById("wnd[0]/sbar").Text, "saved") <> 0 Then ReleaseSalesOrder = Session.FindById("wnd[0]/sbar").Text
                If InStr(Session.FindById("wnd[0]/sbar").Text, "Choose a valid function") > 0 Then 'save was successful..reattempting generated the "choose a valid function" message...keep previous releasesalesorder string
                Else
                    If InStr(ReleaseSalesOrder, "saved") = 0 Then ReleaseSalesOrder = Session.FindById("wnd[0]/sbar").Text
                End If
                If InStr(ReleaseSalesOrder, "saved") = 0 Then
                    j = j + 1
                    If j = 10 Then MsgBox "After 10 tries, the order-release won't save. Call Mike."
                    GoTo retry
                End If
                j = 0
            Else
                If InStr(MyVal, "Waiting") > 0 Then ReleaseSalesOrder = "Z4"
                Session.FindById("wnd[0]/tbar[0]/btn[11]").press 'save
                If InStr(Session.FindById("wnd[0]/sbar").Text, "saved") = 0 Then Stop 'order not saved properly after Z4 removal
            End If
        ElseIf MyVal = "CSR Review" Then
            ReleaseSalesOrder = MyVal
        ElseIf MyVal = "" Then
            If InStr(ReleaseSalesOrder, "saved") = 0 Then ReleaseSalesOrder = "saved" 'probably can remove
        End If
        On Error Resume Next
    ElseIf Left(SalesOrderNum, 2) = 80 And Len(SalesOrderNum) = 8 Then
        ReleaseSalesOrder = "Delivery Num"
    Else
        ReleaseSalesOrder = "Order Dropped" 'if it's one of those sales order numbers with length greater than 5
    End If
    If InStr(MyVal, "Waiting") > 0 Or InStr(MyVal, "Z4") > 0 Then ReleaseSalesOrder = "Z4"
    Call LogOff
    On Error GoTo 0
End Function
Function GetMatNum(OrderNum As String)
    UserName = Range("ZZ1")
    Password = Range("AAA1")
    If Sheets("Main").Range("AA1") = "Production" Then Call LogonProduction Else Call LogonDevelopment
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nco03"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = OrderNum '"1047118"
    Session.FindById("wnd[0]").SendVKey 0
    GetMatNum = Session.FindById("wnd[0]/usr/ctxtCAUFVD-MATNR").Text
    Call LogOff
End Function
Sub ConfirmOrders()
    Dim MyBarcode As String, MaterialMoved As String, MySalesOrderNumber As String, MyMaterialNumber As String, MyProductionOrderNumber As String
    Dim MyRow As Long
    Dim MyTest As String
    '1064072
    '
    AppActivate Application.Caption
retry:
    MyBarcode = InputBox("Scan in the production order number to confirm the order.", "Confirm Order")
    MySalesOrderNumber = Replace(MyBarcode, "SO:", "")

    If Left(MyBarcode, 3) = "SO:" Then 'if there is already inventory in 3002 then move material from 3002 to 3001...then run VA02 and process accordingly
        '=============================Move Material First before releasing anything===================================
        If WorksheetFunction.CountIf(Range("AN:AN"), "*" & Val(MySalesOrderNumber)) = 1 Or WorksheetFunction.CountIf(Range("AN:AN"), "*" & Val(MySalesOrderNumber)) = 0 Then 'if Sales Order number is unique then move inventory into 3001
            If WorksheetFunction.CountIf(Range("AN:AN"), Val(MySalesOrderNumber)) = 1 Then
                MyRow = Application.WorksheetFunction.Match(Val(MySalesOrderNumber), Sheets("Main").Range("AN1:AN100000"), 0)
                Application.StatusBar = "Moving required quantity of " & Range("AC" & MyRow) & " batch " & Range("AH" & MyRow) & " from 3002 to 3001..."
                If InStr(NMIGO_TR(Range("AC" & MyRow), Range("AH" & MyRow), "3002", Range("AC" & MyRow), Range("AH" & MyRow), Range("AO" & MyRow), "3001"), "posted") > 0 Then
                    Range("AQ" & MyRow) = Range("AO" & MyRow) & " each(es) moved from 3002 to 3001."
                    Range("AM" & MyRow) = "Transfer Confirmed"
                End If
            End If
        End If
        If WorksheetFunction.CountIf(Range("AN:AN"), "*" & Val(MySalesOrderNumber)) > 1 Or WorksheetFunction.CountIf(Range("AN:AN"), Val(MySalesOrderNumber)) > 1 Then 'if there are multiple instances of the Sales Order number then you need to scan the barcode on the right
            MyBarcode = InputBox("The Sales Order Number is not unique. Please scan the Material Number barcode on the upper right of the label.", "Confirm Order")
            MyMaterialNumber = Left(MyBarcode, InStr(MyBarcode, "_") - 1)
            MyRow = Application.WorksheetFunction.Match(MyMaterialNumber, Sheets("Main").Range("AC1:AC10000"), 0)
            Application.StatusBar = "Moving required quantity of " & Range("AC" & MyRow) & " batch " & Range("AH" & MyRow) & " from 3002 to 3001..."
            If InStr(NMIGO_TR(Range("AC" & MyRow), Range("AH" & MyRow), "3002", Range("AC" & MyRow), Range("AH" & MyRow), Range("AO" & MyRow), "3001"), "posted") > 0 Then
                Range("AQ" & MyRow) = Range("AO" & MyRow) & " each(es) moved from 3002 to 3001."
                Range("AM" & MyRow) = "Transfer Confirmed"
            End If
        End If
        Call LogOff
        '=============================Inventory Transfer Complete===================================
        Application.StatusBar = "Going into Sales Order " & MySalesOrderNumber
        Logger "MySalesOrderNumber: " & MySalesOrderNumber
        'debugging
        If Len(MySalesOrderNumber) = 6 Or Len(MySalesOrderNumber) = 8 Then Else Stop
        If InStr(MySalesOrderNumber, "_") > 0 Then Stop
        'debugging
        MyTest = ReleaseSalesOrder(MySalesOrderNumber)
        If InStr(MyTest, "saved") <> 0 Then 'if "Waiting on Inventory" was cleared from the Sales Order in VA02
            If InStr(MyTest, "saved") <> 0 And WorksheetFunction.CountIf(Range("AN:AN"), Val(MySalesOrderNumber)) = 1 Then 'if Sales Order number is unique, mark it as released
                MyRow = Application.WorksheetFunction.Match(Val(MySalesOrderNumber), Sheets("Main").Range("AN1:AN100000"), 0)
                Range("AN" & MyRow) = "Released-" & Range("AN" & MyRow)
                Application.StatusBar = Range("AN" & MyRow)
                Else 'if there are multiple instances of the Sales Order number, find the row containing MyMaterialNumber scanned in above and mark that one as Unmodified
                'MyMaterialNumber was already assigned a value in the "multiple instance" material movement subroutine above
                MyRow = Application.WorksheetFunction.Match(MyMaterialNumber, Sheets("Main").Range("AC1:AC10000"), 0)
                Do
                    If Range("AN" & MyRow) = MySalesOrderNumber Then 'if the sales order in the same row as the material number matches the sales order number in play
                        Range("AN" & MyRow) = "Other items on SO-" & Range("AN" & MyRow)
                        Application.StatusBar = Range("AN" & MyRow)
                        Exit Do
                    End If
                    MyRow = MyRow + 1
                    If Range("AC" & MyRow) = "" Then
                        MsgBox MyMaterialNumber & "/" & MySalesOrderNumber & " combination not found."
                        Exit Do
                    End If
                Loop While Range("AN" & MyRow) <> MySalesOrderNumber
                MyRow = 0
            End If
        ElseIf InStr(MyTest, "CSR Review") <> 0 Then 'if the sales order is still blocked (CSR Review)...mark as "CSR Review"
            If WorksheetFunction.CountIf(Range("AN:AN"), "*" & Val(MySalesOrderNumber)) = 1 Or WorksheetFunction.CountIf(Range("AN:AN"), "*" & Val(MySalesOrderNumber)) = 0 Then 'if Sales Order number is unique
                If WorksheetFunction.CountIf(Range("AN:AN"), Val(MySalesOrderNumber)) = 1 Then
                    MyRow = Application.WorksheetFunction.Match(Val(MySalesOrderNumber), Sheets("Main").Range("AN1:AN100000"), 0)
                    Application.StatusBar = "Sales Order " & Range("AN" & MyRow) & " is under CSR Review and cannot be released."
                    Range("AN" & MyRow) = "CSR Review-" & Range("AN" & MyRow)
                    Range("AN" & MyRow).Interior.Color = RGB(0, 176, 240)
                End If
            End If
            If WorksheetFunction.CountIf(Range("AN:AN"), "*" & Val(MySalesOrderNumber)) > 1 Or WorksheetFunction.CountIf(Range("AN:AN"), Val(MySalesOrderNumber)) > 1 Then 'if sales order number is not unique
                '                MyBarcode = InputBox("The Sales Order Number is not unique. Please scan the Material Number barcode on the upper right of the label.", "Confirm Order")
                '                MyMaterialNumber = Left(MyBarcode, InStr(MyBarcode, "_") - 1)
                MyRow = Application.WorksheetFunction.Match(MyMaterialNumber, Sheets("Main").Range("AC1:AC10000"), 0)
                Do
                    If Range("AN" & MyRow) = MySalesOrderNumber Then
                        Range("AN" & MyRow) = "CSR Review-" & Range("AN" & MyRow)
                        Application.StatusBar = Range("AN" & MyRow)
                        Exit Do
                    End If
                    MyRow = MyRow + 1
                    If Range("AC" & MyRow) = "" Then
                        MsgBox MyMaterialNumber & "/" & MySalesOrderNumber & " combination not found."
                        Exit Do
                    End If
                Loop While Range("AN" & MyRow) <> MySalesOrderNumber
                MyRow = 0
            End If
        ElseIf InStr(MyTest, "Delivery") <> 0 Then
            If WorksheetFunction.CountIf(Range("AN:AN"), "*" & Val(MySalesOrderNumber)) = 1 Or WorksheetFunction.CountIf(Range("AN:AN"), "*" & Val(MySalesOrderNumber)) = 0 Then 'if Sales Order number is unique
                If WorksheetFunction.CountIf(Range("AN:AN"), Val(MySalesOrderNumber)) = 1 Then 'if Sales Order number is unique
                    MyRow = Application.WorksheetFunction.Match(Val(MySalesOrderNumber), Sheets("Main").Range("AN1:AN100000"), 0)
                    Application.StatusBar = Range("AN" & MyRow) & " is a delivery number, not a sales document number."
                    Range("AN" & MyRow) = "Delivery-" & Range("AN" & MyRow)
                End If
            End If
            If WorksheetFunction.CountIf(Range("AN:AN"), "*" & Val(MySalesOrderNumber)) > 1 Or WorksheetFunction.CountIf(Range("AN:AN"), Val(MySalesOrderNumber)) > 1 Then 'if sales order number is not unique
                MyBarcode = InputBox("The Sales Order Number is not unique. Please scan the Material Number barcode on the upper right of the label.", "Confirm Order")

                MyMaterialNumber = Left(MyBarcode, InStr(MyBarcode, "_") - 1)
                MyRow = Application.WorksheetFunction.Match(MyMaterialNumber, Sheets("Main").Range("AC1:AC10000"), 0)
                Do
                    If Range("AN" & MyRow) = MySalesOrderNumber Then
                        Range("AN" & MyRow) = "Delivery-" & Range("AN" & MyRow)
                        Application.StatusBar = Range("AN" & MyRow)
                        Exit Do
                    End If
                    MyRow = MyRow + 1
                    If Range("AC" & MyRow) = "" Then
                        MsgBox MyMaterialNumber & "/" & MySalesOrderNumber & " combination not found."
                        Exit Do
                    End If
                Loop While Range("AN" & MyRow) <> MySalesOrderNumber
                MyRow = 0
            End If
        ElseIf InStr(MyTest, "Dropped") <> 0 Then
            If WorksheetFunction.CountIf(Range("AN:AN"), "*" & Val(MySalesOrderNumber)) = 1 Or WorksheetFunction.CountIf(Range("AN:AN"), "*" & Val(MySalesOrderNumber)) = 0 Then
                If WorksheetFunction.CountIf(Range("AN:AN"), Val(MySalesOrderNumber)) = 1 Then 'if Sales Order number is unique
                    MyRow = Application.WorksheetFunction.Match(Val(MySalesOrderNumber), Sheets("Main").Range("AN1:AN100000"), 0)
                    Application.StatusBar = "Sales Order " & Range("AN" & MyRow) & " has dropped."
                    Range("AN" & MyRow) = "Dropped-" & Range("AN" & MyRow)
                End If
            End If
            If WorksheetFunction.CountIf(Range("AN:AN"), "*" & Val(MySalesOrderNumber)) > 1 Or WorksheetFunction.CountIf(Range("AN:AN"), Val(MySalesOrderNumber)) > 1 Then
                MyRow = Application.WorksheetFunction.Match(MyMaterialNumber, Sheets("Main").Range("AC1:AC10000"), 0)
                Do
                    If Range("AN" & MyRow) = MySalesOrderNumber Then
                        Range("AN" & MyRow) = "Dropped-" & Range("AN" & MyRow)
                        Application.StatusBar = Range("AN" & MyRow)
                        Exit Do
                    End If
                    MyRow = MyRow + 1
                    If Range("AC" & MyRow) = "" Then
                        MsgBox MyMaterialNumber & "/" & MySalesOrderNumber & " combination not found."
                        Exit Do
                    End If
                Loop While Range("AN" & MyRow) <> MySalesOrderNumber
                MyRow = 0
            End If
            Else 'VA02 had nothing in the field...assume this means released...or sales order was locked by another user
            If WorksheetFunction.CountIf(Range("AN:AN"), "*" & Val(MySalesOrderNumber)) = 1 Or WorksheetFunction.CountIf(Range("AN:AN"), "*" & Val(MySalesOrderNumber)) = 0 Then
                If WorksheetFunction.CountIf(Range("AN:AN"), Val(MySalesOrderNumber)) = 1 Then 'if Sales Order number is unique
                    MyRow = Application.WorksheetFunction.Match(Val(MySalesOrderNumber), Sheets("Main").Range("AN1:AN100000"), 0)
                    Application.StatusBar = "Sales Order " & Range("AN" & MyRow) & " was not changed."
                    Range("AN" & MyRow) = "Unmodified-" & Range("AN" & MyRow)
                End If
            End If
            If WorksheetFunction.CountIf(Range("AN:AN"), "*" & Val(MySalesOrderNumber)) > 1 Or WorksheetFunction.CountIf(Range("AN:AN"), Val(MySalesOrderNumber)) > 1 Then
                MyRow = Application.WorksheetFunction.Match(MyMaterialNumber, Sheets("Main").Range("AC1:AC10000"), 0)
                Do
                    If Range("AN" & MyRow) = MySalesOrderNumber Then
                        Range("AN" & MyRow) = "Released-" & Range("AN" & MyRow)
                        Application.StatusBar = Range("AN" & MyRow)
                        Exit Do
                    End If
                    MyRow = MyRow + 1
                    If Range("AC" & MyRow) = "" Then
                        MsgBox MyMaterialNumber & "/" & MySalesOrderNumber & " combination not found."
                        Exit Do
                    End If
                Loop While Range("AN" & MyRow) <> MySalesOrderNumber
                MyRow = 0
            End If
        End If
        DoEvents
        MyBarcode = ""
        AppActivate Application.Caption
        Call LogOff
        Call LogOff
        Call LogOff
        Call ConfirmOrders
        Else 'scanned barcode is the PRODUCTION order number, of which there is only one...the correct row will be found
        Application.StatusBar = "Confirming Production Order " & MyBarcode
        While MyBarcode <> "" And IsNumeric(MyBarcode) = True
            On Error Resume Next
            MyRow = 0
            MyRow = Application.WorksheetFunction.Match(Val(MyBarcode), Sheets("Main").Range("AM1:AM100000"), 0)
            If MyRow = 0 Then
                MyRow = Application.WorksheetFunction.Match(GetMatNum(MyBarcode), Sheets("Main").Range("AC1:AC100000"), 0)
                If MyRow = 0 Then
                    MsgBox ("Can't find order number: " & MyBarcode)
                    End
                Else
                    Range("AM" & MyRow) = MyBarcode
                End If
            End If
            MyRow = Application.WorksheetFunction.Match(Val(MyBarcode), Sheets("Main").Range("AM1:AM100000"), 0)
            If Right(Range("AC" & MyRow), 1) = "T" And Left(Range("AC" & MyRow - 1), Len(Range("AC" & MyRow - 1)) - 1) = Left(Range("AC" & MyRow), Len(Range("AC" & MyRow)) - 1) And Left(Range("AM" & MyRow - 1), 10) <> "Confirmed-" And Range("AF" & MyRow).Value = Range("AF" & MyRow - 1).Value Then
                AppActivate Application.Caption
                MsgBox ("Must confirm material " & Range("AC" & MyRow - 1) & " first.")
                End
            End If
            MyTest = ConfirmOrder(MyBarcode, Range("AH" & MyRow))
            If InStr(MyTest, "saved") <> 0 Then
                Range("AM" & MyRow) = "Confirmed-" & Range("AM" & MyRow)
                Dim New3002Inv As Double
                If InStr(Range("AC" & MyRow).Value2, "-T") = 0 Then
                    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmb52"
                    Session.FindById("wnd[0]").SendVKey 0
                    Session.FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = Range("AC" & MyRow)
                    Session.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = "3000"
                    Session.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = "3002"
                    Session.FindById("wnd[0]/usr/ctxtCHARG-LOW").Text = Range("AH" & MyRow)
                    Session.FindById("wnd[0]/tbar[1]/btn[8]").press
                    New3002Inv = Trim(Val(Session.FindById("wnd[0]/usr/lbl[81,3]").Text))
                    Range("AQ" & MyRow).Value2 = New3002Inv & " eaches created in 3002."
                End If
                Call LogOff
                If Range("AB" & MyRow) <> "Dilute 1st" Then
                    On Error GoTo 0
                    'if the material does not need to be diluted to make the M then move the required quantity from 3002 to 3001
                    If InStr(Range("AC" & MyRow).Value2, "-T") = 0 Then 'all of the Ts are already in 3001 by default with the material master change...i.e. only need to move Ms
                        'confirming the T order via CO11 creates inventory in 3001
                        Application.StatusBar = "Moving required quantity of " & Range("AC" & MyRow) & " batch " & Range("AH" & MyRow) & " from 3002 to 3001..."
                        If InStr(NMIGO_TR(Range("AC" & MyRow), Range("AH" & MyRow), "3002", Range("AC" & MyRow), Range("AH" & MyRow), Range("AO" & MyRow), "3001"), "posted") > 0 Then
                            Range("AQ" & MyRow) = Range("AO" & MyRow) & " each(es) moved from 3002 to 3001"
                        End If
                    End If
                    Application.StatusBar = "Releasing Sales Order..."
                    'Stop 'verify that sales order number is valid
                    MyTest = ReleaseSalesOrder(Range("AN" & MyRow))
                    If InStr(MyTest, "saved") <> 0 Then 'if the sales order was successfully released...
                        If InStr(MyTest, "saved") <> 0 Then
                            Range("AN" & MyRow) = "Released-" & Range("AN" & MyRow)
                            Application.StatusBar = Range("AN" & MyRow)
                        End If
                    ElseIf InStr(MyTest, "CSR Review") <> 0 Then 'if the sales order is still blocked (CSR Review)...
                        If InStr(MyTest, "CSR Review") <> 0 Then
                            Application.StatusBar = "Sales Order " & Range("AN" & MyRow) & " is under CSR Review and cannot be released."
                            Range("AN" & MyRow) = "CSR Review-" & Range("AN" & MyRow)
                            Range("AN" & MyRow).Interior.Color = RGB(0, 176, 240)
                        End If
                    ElseIf InStr(MyTest, "Z4") <> 0 Then
                        If InStr(MyTest, "Z4") <> 0 Then
                            Application.StatusBar = "Sales Order " & Range("AN" & MyRow) & " remains under Z4 block."
                            Range("AN" & MyRow) = "Z4-" & Range("AN" & MyRow)
                            Range("AN" & MyRow).Interior.Color = RGB(0, 176, 240)
                        End If
                    End If
                End If
            End If
            DoEvents
            MyBarcode = ""
            AppActivate Application.Caption
            MyBarcode = InputBox("Scan in the production order number to confirm the order.", "Confirm Order")
        Wend
    End If
    Call LogOff
    Application.StatusBar = False
End Sub
Function ConfirmOrder(ProdOrder As String, Batch As String)
    UserName = Range("ZZ1")
    Password = Range("AAA1")
    If Sheets("Main").Range("AA1") = "Production" Then Call LogonProduction Else Call LogonDevelopment
    On Error Resume Next
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nco11"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ctxtCORUF-AUFNR").Text = ProdOrder '"1047085"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[1]/usr/tblSAPLCORUTC_OPERATIONS/radRC27X-FLG_SEL[0,1]").Selected = True
    'Put projected setup & labor into step
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:0850/txtAFRUD-ISM01").Text = Session.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:0850/txtCORUF-SLM01").Text
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:0850/ctxtAFRUD-ILE01").Text = Session.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:0850/txtCORUF-SLE01").Text
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:0850/txtAFRUD-ISM03").Text = Session.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:0850/txtCORUF-SLM03").Text
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:0850/ctxtAFRUD-ILE03").Text = Session.FindById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:0850/txtCORUF-SLE03").Text
    Session.FindById("wnd[0]/tbar[1]/btn[18]").press
    Session.FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-CHARG[5,0]").Text = Batch '"1"
    If Session.FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-MATNR[0,2]").Text <> "" And Session.FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-CHARG[5,2]").Text = "" Then
        Session.FindById("wnd[0]/usr/subTABLE:SAPLCOWB:0500/tblSAPLCOWBTCTRL_0500/ctxtCOWB_COMP-CHARG[5,2]").Text = Sheets("MB52").Range("Q14")  '"141212"
    End If
    'Batch = 2 'don't know the purpose of this...was probably inserted to try to finish process
    Session.FindById("wnd[0]/tbar[0]/btn[11]").press
    'added to get around issue with financials
    On Error Resume Next
    Session.FindById("wnd[1]/usr/btnSPOP-OPTION2").press
    Session.FindById("wnd[1]/usr/btnBUTTON_1").press
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0

    ConfirmOrder = Session.FindById("wnd[0]/sbar").Text
    'go back in and confirm last step 9999
    On Error GoTo 0
    On Error GoTo retry
retry:
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nco11"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ctxtCORUF-AUFNR").Text = ProdOrder '"1047085"
    Session.FindById("wnd[0]").SendVKey 0
    'sometimes SAP says the order is being modified by another user
    If InStr(Session.FindById("wnd[0]/sbar").Text, "currently") > 0 Then GoTo retry 'SAP is locking the user out of the record...need to loop until it stops locking out
    If InStr(Session.FindById("wnd[0]/sbar").Text, "already confirmed") <> 0 Then
        Exit Function
    End If
    Session.FindById("wnd[0]/tbar[0]/btn[11]").press 'attempt to save 9999 confirmation
    If InStr(Session.FindById("wnd[0]/sbar").Text, "saved") <> 0 Then 'if the save operation worked, then the status bar should have "saved" in it...9999 confirmed...end function
    Else
        GoTo retry
    End If
End Function
Sub GetProdOrder()
    MsgBox (CheckForProdOrder("A303-830A-T"))
End Sub
Function CheckForProdOrder(MatNum As String)
    UserName = Range("ZZ1")
    Password = Range("AAA1")
    If Sheets("Main").Range("AA1") = "Production" Then Call LogonProduction Else Call LogonDevelopment
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/ncoois"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").Selected = True
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").Selected = True
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_MATNR-LOW").Text = MatNum '"A303-830A-T"
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").Text = "teco"
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").Text = "dlv"
    Session.FindById("wnd[0]").SendVKey 8
    On Error Resume Next
    CheckForProdOrder = Session.FindById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").GetCellValue(0, "AUFNR")
    Call LogOff
End Function
Sub CreateProductionOrder()
    UserName = Range("ZZ1")
    Password = Range("AAA1")
    If Range("AI" & ActiveCell.Row) = "" Then
        MsgBox ("To create an order you must have a Component Storage Location.")
        Range("AI" & ActiveCell.Row).Select
        Exit Sub
    End If
    If Range("AP" & ActiveCell.Row) >= Range("AO" & ActiveCell.Row) Then
        'dont make order, but print vialing label
        Call PrintSOLabel
        Range("AM" & ActiveCell.Row) = "Confirm to Move Inventory"
        ActiveCell.Offset(1, 0).Select
        End
    End If
    If Range("AM" & ActiveCell.Row) = "" Then
        Range("AM" & ActiveCell.Row) = CheckForProdOrder(Range("AC" & ActiveCell.Row))
        If WorksheetFunction.CountIf(Range("AM2:AM1000"), Range("AM" & ActiveCell.Row)) > 1 Then
            Dim tempProdOrder As String
            tempProdOrder = Range("AM" & ActiveCell.Row)
            Range("AM" & ActiveCell.Row) = ""
            MsgBox Range("AC" & ActiveCell.Row) & " appears more than once in the material list. Please confirm order " & tempProdOrder & " before creating the new order."
            Rows(WorksheetFunction.Match(Val(tempProdOrder), Range("AM2:AM1000"), 0) + 1).Select
            End
        End If
        If Range("AM" & ActiveCell.Row) <> "" Then
            AppActivate Application.Caption
            MsgBox ("Order already exists or 10-Blot not confirmed.")
            Range("AC" & ActiveCell.Row).Select
            ActiveCell.Offset(1, 0).Select
            End
        End If
    End If
    If Range("AB" & ActiveCell.Row - 1) = "Dilute 1st" And InStr(Range("AM" & ActiveCell.Row - 1), "Confirmed") = 0 Then
        MsgBox ("You need to confirm the order above first")
        ActiveCell.Offset(-1, 0).Select
        End
    End If
    UserName = Range("ZZ1")
    Password = Range("AAA1")
    If Sheets("Main").Range("AA1") = "Production" Then Call LogonProduction Else Call LogonDevelopment
    If Range("AM" & ActiveCell.Row) = "" Then
        If Right(Range("AC" & ActiveCell.Row), 1) = "M" Then
            Application.StatusBar = "Generating Production Order for " & Range("AC" & ActiveCell.Row)
            'create order and print both labels
            Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nco01"
            Session.FindById("wnd[0]").SendVKey 0
            Session.FindById("wnd[0]/usr/ctxtCAUFVD-MATNR").Text = Range("AC" & ActiveCell.Row) '"A300-110A-M"
            Session.FindById("wnd[0]/usr/ctxtCAUFVD-WERKS").Text = "3000"
            Session.FindById("wnd[0]/usr/ctxtAUFPAR-PP_AUFART").Text = "ZP01"
            Session.FindById("wnd[0]").SendVKey 0
            Session.FindById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/txtCAUFVD-GAMNG").Text = Range("AL" & ActiveCell.Row) / 110 '"20"
            Session.FindById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/ctxtCAUFVD-GMEIN").Text = "EA"
            Session.FindById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/ctxtCAUFVD-GLTRP").Text = Date '"12/18/2014"
            Session.FindById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/ctxtCAUFVD-GSTRP").Text = Date '"12/18/2014"
            Session.FindById("wnd[0]").SendVKey 0
            Session.FindById("wnd[0]/tbar[1]/btn[6]").press
            Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120").Columns.ElementAt(0).Selected = True
            Session.FindById("wnd[0]/usr/cmbSORT_BOX").Key = "ST_COU"
            Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-EINHEIT[4,0]").Text = "無"
            Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-EINHEIT[4,1]").Text = "無"
            If InStr(Range("AJ" & ActiveCell.Row), " ") = 0 Then
                Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3,0]").Text = Range("AJ" & ActiveCell.Row) '"220"
            Else
                Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3,0]").Text = Trim(Left(Range("AJ" & ActiveCell.Row), InStr(Range("AJ" & ActiveCell.Row), " ")))
            End If
            Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3,1]").Text = Range("AK" & ActiveCell.Row) '"1980"
            Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-LGORT[9,0]").Text = Range("AI" & ActiveCell.Row)  '"3001"
            Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-CHARG[10,0]").Text = Range("AH" & ActiveCell.Row) '"3"
            Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-CHARG[10,1]").Text = Sheets("MB52").Range("Q14") '"141212"
            Session.FindById("wnd[0]/tbar[0]/btn[11]").press
            'capture message here before enter to paste into column AM...enter will erase message with order number
            On Error Resume Next
            Range("AM" & ActiveCell.Row) = Replace(Replace(Session.FindById("wnd[0]/sbar").Text, "Order number ", ""), " saved", "")
            On Error GoTo 0
            Session.FindById("wnd[0]").SendVKey 0
            Session.FindById("wnd[0]").SendVKey 0
            If Len(Range("AM" & ActiveCell.Row)) <> 7 Then
                Range("AM" & ActiveCell.Row) = Replace(Replace(Session.FindById("wnd[0]/sbar").Text, "Order number ", ""), " saved", "")
            End If
            Call PrintLabel
            'If Range("AB" & ActiveCell.Row) = "Dilute 1st" Then Call PrintComboLabel_MStock
        End If
        If Right(Range("AC" & ActiveCell.Row), 1) = "T" Then
            Application.StatusBar = "Generating Production Order for " & Range("AC" & ActiveCell.Row)
            Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nco01"
            Session.FindById("wnd[0]").SendVKey 0
            Session.FindById("wnd[0]/usr/ctxtCAUFVD-MATNR").Text = Range("AC" & ActiveCell.Row)
            Session.FindById("wnd[0]/usr/ctxtCAUFVD-WERKS").Text = "3000"
            Session.FindById("wnd[0]/usr/ctxtAUFPAR-PP_AUFART").Text = "ZP01"
            Session.FindById("wnd[0]").SendVKey 0
            Session.FindById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/txtCAUFVD-GAMNG").Text = Range("AO" & ActiveCell.Row)
            Session.FindById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/ctxtCAUFVD-GMEIN").Text = "EA"
            Session.FindById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/ctxtCAUFVD-GLTRP").Text = Date
            Session.FindById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/ctxtCAUFVD-GSTRP").Text = Date
            Session.FindById("wnd[0]").SendVKey 0
            Session.FindById("wnd[0]/tbar[1]/btn[6]").press
            If Right(Range("AG" & ActiveCell.Row), 1) = "M" Then Range("AI" & ActiveCell.Row) = "3002"
            If Right(Range("AI" & ActiveCell.Row), 1) = "P" Then
                MsgBox "The program is about to create or use a partial vial. Notify IT development to step through and improve this process."
                'if there is material in 3001P, then create a mock order (consuming nothing) so shipping will see it.
                Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-LGORT[9,0]").Text = "3001"
                Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3,0]").Text = 0
                'Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-EINHEIT[4,0]").Text = "無"
                Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-CHARG[10,0]").Text = Range("AH" & ActiveCell.Row)
                Session.FindById("wnd[0]/tbar[0]/btn[11]").press
                Range("AM" & ActiveCell.Row) = Replace(Replace(Session.FindById("wnd[0]/sbar").Text, "Order number ", ""), " saved", "")
                Call PrintLabel
            Else
                If Range("AI" & ActiveCell.Row) = "3001" And InStr(Range("AG" & ActiveCell.Row), "-M") = 0 Then 'if classic trial will be taken from 3001
                    'if there is no material in 3001P and we're making a classic trial out of 3001 (because there is no 3002 bulk), then move the whole 3001 vial out of SAP into 3001P
                    '20161102...no more partial consumption of 3001 vials...pull the entire vial out of 3001 and move the vial into non-SAP storage location 3001P (exists only in M&T vialing program)
                    Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-LGORT[9,0]").Text = "3001"
                    Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3,0]").Text = Range("AO" & ActiveCell.Row)
                    'Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-EINHEIT[4,0]").Text = "無"
                    Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-CHARG[10,0]").Text = Range("AH" & ActiveCell.Row) '"3"
                    Session.FindById("wnd[0]/tbar[0]/btn[11]").press
                    Range("AM" & ActiveCell.Row) = Replace(Replace(Session.FindById("wnd[0]/sbar").Text, "Order number ", ""), " saved", "")
                    Call PrintLabel
                    Else 'all other scenarios besides classic trial being pulled from 3001
                    Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-LGORT[9,0]").Text = Range("AI" & ActiveCell.Row)
                    If Left(Range("AG" & ActiveCell.Row), 4) = "A500" Then
                        If WorksheetFunction.VLookup(Range("AG" & ActiveCell.Row), Sheets("LabelData").Range("A:F"), 6, False) = "100 痞 (10 blots)" Then
                            Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3,0]").Text = Range("AO" & ActiveCell.Row) * 22
                            Else 'then it's the IHC version
                            Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3,0]").Text = Range("AO" & ActiveCell.Row) * 11
                        End If
                    ElseIf Right(Range("AG" & ActiveCell.Row), 1) = "A" Then
                        Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3,0]").Text = Range("AO" & ActiveCell.Row) * 11
                    ElseIf Left(Range("AG" & ActiveCell.Row), 3) = "IHC" Then
                        Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3,0]").Text = Range("AO" & ActiveCell.Row) * 11
                    ElseIf Left(Range("AG" & ActiveCell.Row), 3) = "A70" Then
                        'mbb 07272016...multiply standard 22痞 A700 trial volume by quotient of WBconc / ABconc...e.g. if the WB conc is 2痢/ml and the Ab stock is 1痢/ml (i.e. A700-002), then 44痞 will be shipped
                        Dim BOMQuantity As String
                        BOMQuantity = WorksheetFunction.VLookup(Range("AC" & ActiveCell.Row), Sheets("BomChild").Range("A:E"), 5, False)
                        Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3,0]").Text = CInt(BOMQuantity) * 0.11
                        'Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3,0]").Text = Range("AO" & ActiveCell.Row) * 22 * (WorksheetFunction.VLookup(Range("AG" & ActiveCell.Row), Sheets("Zcharvalues").Range("A:F"), 4, False) / WorksheetFunction.VLookup(Range("AG" & ActiveCell.Row), Sheets("Zcharvalues").Range("A:F"), 6, False)) '"220"
                    Else
                        Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3,0]").Text = Range("AO" & ActiveCell.Row) * 22
                    End If
                    Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-EINHEIT[4,0]").Text = "無"
                    Session.FindById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-CHARG[10,0]").Text = Range("AH" & ActiveCell.Row)
                    Session.FindById("wnd[0]/tbar[0]/btn[11]").press
                    Range("AM" & ActiveCell.Row) = Replace(Replace(Session.FindById("wnd[0]/sbar").Text, "Order number ", ""), " saved", "")
                    Call PrintLabel
                End If
            End If
        End If
    End If
    Call LogOff
    AppActivate ("Microsoft Excel")
    ActiveCell.Offset(1, 0).Select
    Application.StatusBar = False
End Sub
Function GetMatBatch()
    Sheets("MB52").Range("M17") = "Material"
    Sheets("MB52").Range("N17") = "Sloc"
    Sheets("MB52").Range("O17") = "Unrestricted"
    Sheets("MB52").Range("M18") = Range("AC" & ActiveCell.Row)
    Sheets("MB52").Range("N18") = "3002"
    Sheets("MB52").Range("O18") = ">=" & Range("AO" & ActiveCell.Row)
    Sheets("MB52").Range("Q17").FormulaR1C1 = "=DMIN(C1:C10,""Batch"",R17C13:R18C15)"
    GetMatBatch = Sheets("MB52").Range("Q17")
End Function
Sub RunMain()
    Dim WBConc As Double
    Dim AbConc As Double
    Dim TAmt As Double
    Dim MinAmt As Double
    Dim MyRow As Long, NumericalA300 As Double
    Dim OrderAmt As Double, My3001InventorySum As Double, DilutionFactor As Double, AntibodyAmt As Double, VialCountToCreateFull As Double
    Sheets("Main").Range("AA2:AZ100000") = vbNullString
    With Sheets("Main").Range("AA2:AZ100000").Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Bold = False
        .Italic = False
    End With
    Rows("2:100000").Interior.Pattern = xlNone
    Application.StatusBar = "Getting ZVialOrders..."
    Application.ScreenUpdating = False
    Call ZVialOrders
    Application.StatusBar = "Getting Component Data..."
    Application.ScreenUpdating = False
    Call SQ01_BOMComponent
    Sheets("BOMChild").Range("O1") = "A400Diluent"
    Application.StatusBar = "Getting MB52 Inventory..."
    Application.ScreenUpdating = False
    Call GetMB52Inventory
    Application.StatusBar = "Getting ZCharValues..."
    Application.ScreenUpdating = False
    Call ZCharValues
    '    If WorksheetFunction.CountA(Sheets("ZCharValues").Range("H2:H100")) > 0 Then
    '        MsgBox "One of the ordered materials is not correctly configured in SAP." & vbCrLf & vbCrLf & "The classic product has WB concentration data in MSC2N, but the associated M material has not been created in SAP." & vbCrLf & vbCrLf & _
    '        "Please add the M material(s) marked on the Zcharvalues sheet to SAP and click Update List." & vbCrLf & vbCrLf & _
    '        "If the previous lot is not M-qualified (i.e trial is pulled from classic), consider scrapping the previous lot and using the new M-qualified lot to make the trial."
    '        End
    '    End If
    Sheets("Main").Select
    Sheets("Main").Range("AC2:AC10000") = Sheets("ZVialOrders").Range("J2:J10000").Value
    Sheets("Main").Range("AE2:AE10000") = Sheets("ZVialOrders").Range("N2:N10000").Value
    Sheets("Main").Range("AF2:AF10000") = Sheets("ZVialOrders").Range("I2:I10000").Value
    Sheets("Main").Range("AN2:AN10000") = Sheets("ZVialOrders").Range("A2:A10000").Value
    Sheets("Main").Range("AO2:AO10000") = Sheets("ZVialOrders").Range("L2:L10000").Value
    Sheets("Main").Range("AP2:AP10000") = Sheets("ZVialOrders").Range("T2:T10000").Value 'inventory as it appears in zvialorders
    Range("AD2:AD10000").FormulaR1C1 = "=IF(RC29="""","""",VLOOKUP(RC29,BOMChild!C1:C11,2,FALSE))"
    Application.StatusBar = "Finding Component and Batch Data..."
    Range("AG2:AG10000").FormulaR1C1 = "=IF(RC29="""","""",VLOOKUP(RC29,BOMChild!C1:C11,7,FALSE))"
    Range("AD2:AD10000") = Range("AD2:AD10000").Value
    Range("AG2:AG10000") = Range("AG2:AG10000").Value
    'get comp location and lowest batch
    Range("AC2").Select 'delete A711s
    While Range("AC" & ActiveCell.Row) <> ""
        If Left(Range("AC" & ActiveCell.Row), 4) = "A711" Then
            ActiveCell.EntireRow.Delete
        Else
            ActiveCell.Offset(1, 0).Activate
        End If
        DoEvents
    Wend
    Range("AG2").Select
    'run through list and see if there are any matching classic items on PartialVials or if any items will need to be created
    While Range("AG" & ActiveCell.Row) <> ""
        If Right(Range("AG" & ActiveCell.Row), 1) = "A" Or Left(Range("AG" & ActiveCell.Row), 3) = "IHC" Or Left(Range("AG" & ActiveCell.Row), 4) = "A700" Then
            Dim MyPartialClassicRow As Double, MyPartialClassicAmount As Double
            MyPartialClassicRow = 0
            On Error Resume Next
            MyPartialClassicRow = WorksheetFunction.Match(Range("AG" & ActiveCell.Row), Sheets("PartialVials").Range("A:A"), 0)
            MyPartialClassicAmount = Sheets("PartialVials").Range("C" & MyPartialClassicRow).Value
            On Error GoTo 0
            If MyPartialClassicRow <> 0 And MyPartialClassicAmount > 0 Then
                'MsgBox "The program is about to create or use a partial vial. Notify IT development to step through and improve this process."
                Range("AI" & ActiveCell.Row) = "3001P"
                Range("AH" & ActiveCell.Row) = Sheets("PartialVials").Range("B" & MyPartialClassicRow)
            End If
        End If
        ActiveCell.Offset(1, 0).Activate
        DoEvents
    Wend
    Range("AG2").Select
    While Range("AG" & ActiveCell.Row) <> ""
        MyRow = 0
        If Range("AH" & ActiveCell.Row) = "" Then 'if no materials were found during the PartialVials Search
            If Range("AP" & ActiveCell.Row) >= 1.1 And Range("AO" & ActiveCell.Row) <= Range("AP" & ActiveCell.Row) - 1 Then 'if inventory has more than one each and exceeds required amount by 1 each
                'then there is enough inventory available to cover this order...change the component in column AG to -M so the program will know to not make a new M stock
                Dim Existing3002M As Boolean
                Existing3002M = True
                If Left(ActiveCell.Value, 3) = "A30" Then
                    NumericalA300 = CDbl(Trim(Replace(Replace(Replace(Replace(Range("AC" & ActiveCell.Row), "A30", ""), "A-T", ""), "A-M", ""), "-", "")))
                    If NumericalA300 < 5576 Then Range("AG" & ActiveCell.Row) = Replace(Range("AG" & ActiveCell.Row), "985A100", "985A") & "-M" 'change the component to -M...order will be fulfilled using existing M stock
                End If
                If InStr(Range("AG" & ActiveCell.Row), "-M-M") > 0 Then MsgBox "The component has '-M-M'...call Mike to fix!"
                Range("AG" & ActiveCell.Row).Font.Color = RGB(0, 0, 250)
                Range("AG" & ActiveCell.Row).Font.Bold = True
                Range("AG" & ActiveCell.Row).Font.Italic = True
            Else
                MyRow = WorksheetFunction.Match(Range("AG" & ActiveCell.Row), Sheets("BOMChild").Range("G:G"), 0)
                If MyRow = 0 Then Stop
            End If
            Sheets("MB52").Range("N2").Value = Range("AG" & ActiveCell.Row)
            OrderAmt = Range("AO" & ActiveCell.Row)
            If Not (Existing3002M) Then
                MyRow = WorksheetFunction.Match(Range("AG" & ActiveCell.Row), Sheets("BOMChild").Range("G:G"), 0)
                If MyRow = 0 Then Stop
            End If
            'if a-t or m-t component UI >=1
            'if a-m-t min inventory is the amount to make the M qty/50
            If Right(Range("AG" & ActiveCell.Row), 1) = "A" And Right(Range("AC" & ActiveCell.Row), 1) = "T" Then
                Sheets("MB52").Range("M2").Value = ">=" & 1 * OrderAmt
                Sheets("MB52").Range("M5").Value = ">=" & 1 * OrderAmt
                Sheets("MB52").Range("M8").Value = ">=" & 1 * OrderAmt
                Sheets("MB52").Range("M11").Value = ">=" & 1 * OrderAmt
            End If
            'if IHC trial then treat it like the A300 classic trial
            If Left(Range("AG" & ActiveCell.Row), 3) = "IHC" And Right(Range("AC" & ActiveCell.Row), 1) = "T" Then
                Sheets("MB52").Range("M2").Value = ">=" & 1 * OrderAmt
                Sheets("MB52").Range("M5").Value = ">=" & 1 * OrderAmt
                Sheets("MB52").Range("M8").Value = ">=" & 1 * OrderAmt
                Sheets("MB52").Range("M11").Value = ">=" & 1 * OrderAmt
            End If
            If (Right(Range("AG" & ActiveCell.Row), 1) = "A" Or Right(Range("AG" & ActiveCell.Row), 1) = "A100") And Right(Range("AC" & ActiveCell.Row), 1) = "M" Then
                '                    If Range("AP" & ActiveCell.Row) >= 1.1 And Range("AO" & ActiveCell.Row) > Range("AP" & ActiveCell.Row) - 1 Then 'if there is not enough inventory to cover this order, then
                MinAmt = (Sheets("BOMChild").Range("E" & MyRow) / 50) * OrderAmt
                Sheets("MB52").Range("M2").Value = ">=" & MinAmt
                Sheets("MB52").Range("M5").Value = ">=" & MinAmt
                Sheets("MB52").Range("M8").Value = ">=" & MinAmt
                Sheets("MB52").Range("M11").Value = ">=" & MinAmt
            End If
            If Right(Range("AG" & ActiveCell.Row), 1) = "M" And Right(Range("AC" & ActiveCell.Row), 1) = "T" Then
                MinAmt = (Sheets("BOMChild").Range("E" & MyRow) / 50) * OrderAmt
                Sheets("MB52").Range("M2").Value = ">=" & 1 * OrderAmt
                Sheets("MB52").Range("M5").Value = ">=" & 1 * OrderAmt
                Sheets("MB52").Range("M8").Value = ">=" & MinAmt
                Sheets("MB52").Range("M11").Value = ">=" & MinAmt
            End If
            '============begin batch selection process
            If Sheets("MB52").Range("S2").Value <> 0 Then
                Range("AH" & ActiveCell.Row) = Sheets("MB52").Range("Q2").Value 'line that changes batch number
                'A300-245A has enough batch 4, but batch 5 keeps getting picked because batch 4 doesn't have 8 eaches in inventory...
                'need to make the program pick the older batch
                Range("AI" & ActiveCell.Row) = Sheets("MB52").Range("O2").Value

                ActiveCell.Offset(1, 0).Select
            Else
                If Sheets("MB52").Range("S5").Value <> 0 Then
                    Range("AH" & ActiveCell.Row) = Sheets("MB52").Range("Q5").Value 'line that changes batch number
                    Range("AI" & ActiveCell.Row) = Sheets("MB52").Range("O5").Value
                    ActiveCell.Offset(1, 0).Select
                Else
                    If Sheets("MB52").Range("S8").Value <> 0 And Range("AH" & ActiveCell.Row) = "" Then
                        Range("AH" & ActiveCell.Row) = Sheets("MB52").Range("Q8").Value 'line that changes batch number
                        Range("AI" & ActiveCell.Row) = Sheets("MB52").Range("O8").Value
                        If Range("AP" & ActiveCell.Row) >= 1.1 And Range("AO" & ActiveCell.Row) < Range("AP" & ActiveCell.Row) - 1 Then
                        Else
                            Rows(ActiveCell.Row & ":" & ActiveCell.Row).Copy
                            ActiveCell.EntireRow.Insert
                            Range("AC" & ActiveCell.Row) = Range("AG" & ActiveCell.Row)
                            If Range("AG" & ActiveCell.Row) = "A301-985A-M" Then
                                Range("AG" & ActiveCell.Row) = Replace(Range("AG" & ActiveCell.Row), "-M", "100")
                            Else
                                Range("AG" & ActiveCell.Row) = Replace(Range("AG" & ActiveCell.Row), "-M", "")
                            End If
                            Range("AD" & ActiveCell.Row) = Replace(Range("AD" & ActiveCell.Row), "TRIAL SIZE", "10-BLOTS")
                            Range("AB" & ActiveCell.Row) = "Dilute 1st"
                            Range("AC" & ActiveCell.Row & ":AO" & ActiveCell.Row + 1).Interior.Color = RGB(250, 190, 190)
                            ActiveCell.Offset(2, 0).Select
                        End If
                    Else
                        If Sheets("MB52").Range("S11").Value <> 0 And Range("AH" & ActiveCell.Row) = "" Then
                            Range("AH" & ActiveCell.Row) = Sheets("MB52").Range("Q11").Value 'line that changes batch number
                            Range("AI" & ActiveCell.Row) = Sheets("MB52").Range("O11").Value
                            Rows(ActiveCell.Row & ":" & ActiveCell.Row).Copy
                            ActiveCell.EntireRow.Insert
                            Range("AC" & ActiveCell.Row) = Range("AG" & ActiveCell.Row)
                            Range("AG" & ActiveCell.Row) = Replace(Range("AG" & ActiveCell.Row), "-M", "")
                            Range("AD" & ActiveCell.Row) = Replace(Range("AD" & ActiveCell.Row), "TRIAL SIZE", "10-BLOTS")
                            Range("AB" & ActiveCell.Row) = "Dilute 1st"
                            Range("AC" & ActiveCell.Row & ":AO" & ActiveCell.Row + 1).Interior.Color = RGB(250, 190, 190)
                            ActiveCell.Offset(2, 0).Select
                        ElseIf Range("AH" & ActiveCell.Row) = "" Then
                            ActiveCell.EntireRow.Delete 'no inventory to make vial
                        End If
                    End If
                End If
            End If
            DoEvents
        Else
            ActiveCell.Offset(1, 0).Activate
        End If
        Existing3002M = False
    Wend
    Range("AG2").Select
    Dim MultOrders As Boolean
    MultOrders = False
    Application.StatusBar = "Calculating amounts..."
    While Range("AG" & ActiveCell.Row) <> ""
        If Range("AB" & ActiveCell.Row) <> "" Then
            Sheets("MB52").Range("N2").Value = Range("AG" & ActiveCell.Row)
            MyRow = 0
            MyRow = WorksheetFunction.Match(Range("AG" & ActiveCell.Row), Sheets("BOMChild").Range("G:G"), 0)
            MinAmt = (Sheets("BOMChild").Range("E" & MyRow) / 50) * OrderAmt
            Sheets("MB52").Range("M2").Value = ">=" & 1 * OrderAmt
            Sheets("MB52").Range("M5").Value = ">=" & 1 * OrderAmt
            Sheets("MB52").Range("M8").Value = ">=" & MinAmt
            Sheets("MB52").Range("M11").Value = ">=" & MinAmt
            If Sheets("MB52").Range("S8") = 0 Then
                Range("AH" & ActiveCell.Row) = Sheets("MB52").Range("Q11").Value 'line that changes batch number
                Range("AI" & ActiveCell.Row) = Sheets("MB52").Range("O11").Value
            Else
                Range("AH" & ActiveCell.Row) = Sheets("MB52").Range("Q8").Value 'line that changes batch number
                Range("AI" & ActiveCell.Row) = Sheets("MB52").Range("O8").Value
            End If
        End If

        If Range("AO" & ActiveCell.Row) > 1 Then
            Range("AO" & ActiveCell.Row).Interior.Color = RGB(200, 250, 20)
            MultOrders = True
        End If
        ActiveCell.Offset(1, 0).Select
        DoEvents
    Wend
    'get Ab, Dil & total amount
    Range("AG2").Select
    While Range("AG" & ActiveCell.Row) <> ""
        If Right(Range("AC" & ActiveCell.Row), 2) = "-M" And Right(Range("AG" & ActiveCell.Row), 2) <> "-M" Then
            Sheets("ZCharValues").Range("O2").Value = "'=" & Range("AG" & ActiveCell.Row).Value
            Sheets("ZCharValues").Range("P2").Value = Range("AH" & ActiveCell.Row).Value
            Sheets("ZCharValues").Columns("A:F").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Sheets("ZCharValues").Range("O1:P2"), CopyToRange:=Sheets("ZCharValues").Range("R1:S1"), Unique:=False
            'r2 wb and s2 ab stock
            WBConc = Sheets("ZCharValues").Range("R2").Value
            TAmt = 1100 'max volume of M stock vial
            If WBConc <> 0 Then
                'If WBConc = 0.04 Then TAmt = 2750 Else TAmt = 2200
                If Range("AI" & ActiveCell.Row) = "3002" Then
                    AbConc = Sheets("ZCharValues").Range("S2").Value
                    Range("AJ" & ActiveCell.Row) = (WBConc * TAmt) / AbConc
                    Range("AK" & ActiveCell.Row) = TAmt - Range("AJ" & ActiveCell.Row)
                    Range("AL" & ActiveCell.Row) = Range("AK" & ActiveCell.Row) + Range("AJ" & ActiveCell.Row)
                ElseIf Range("AI" & ActiveCell.Row) = "3001" Then
                    My3001InventorySum = WorksheetFunction.SumIfs(Sheets("MB52").Range("F:F"), Sheets("MB52").Range("A:A"), Range("AG" & ActiveCell.Row), Sheets("MB52").Range("D:D"), "3001", Sheets("MB52").Range("C:C"), Range("AH" & ActiveCell.Row))
                    AbConc = Sheets("ZCharValues").Range("S2").Value
                    AntibodyAmt = (WBConc * TAmt) / AbConc
                    VialCountToCreateFull = WorksheetFunction.RoundUp(AntibodyAmt / 110, 0)
                    If My3001InventorySum - VialCountToCreateFull >= 1 Then 'if using the #vials required to make 1100痞 of M leaves 1 or more vials left in the classic bin, then proceed as normal (i.e. make the full volume)
                        Range("AJ" & ActiveCell.Row) = AntibodyAmt
                        Range("AL" & ActiveCell.Row) = TAmt 'total volume
                        Range("AK" & ActiveCell.Row) = Range("AL" & ActiveCell.Row) - Range("AJ" & ActiveCell.Row) 'diluent
                        If VialCountToCreateFull = 1 Then
                            Range("AJ" & ActiveCell.Row) = Range("AJ" & ActiveCell.Row) & " (" & VialCountToCreateFull & " vial)"
                        Else
                            Range("AJ" & ActiveCell.Row) = Range("AJ" & ActiveCell.Row) & " (" & VialCountToCreateFull & " vials)"
                        End If
                        'move vials out of 3001 to
                    ElseIf My3001InventorySum - (VialCountToCreateFull * 0.75) >= 1 Then '3/4 volume
                        Range("AJ" & ActiveCell.Row) = AntibodyAmt * 0.75 'stock volume
                        Range("AL" & ActiveCell.Row) = TAmt * 0.75 'total volume
                        Range("AK" & ActiveCell.Row) = Range("AL" & ActiveCell.Row) - Range("AJ" & ActiveCell.Row) 'diluent
                        VialCountToCreateFull = VialCountToCreateFull * 0.75
                        If VialCountToCreateFull = 1 Then
                            Range("AJ" & ActiveCell.Row) = Range("AJ" & ActiveCell.Row) & " (" & VialCountToCreateFull & " vial)"
                        Else
                            Range("AJ" & ActiveCell.Row) = Range("AJ" & ActiveCell.Row) & " (" & VialCountToCreateFull & " vials)"
                        End If
                    ElseIf My3001InventorySum - (VialCountToCreateFull * 0.5) >= 1 Then '1/2 volume
                        Range("AJ" & ActiveCell.Row) = AntibodyAmt * 0.5 'stock volume
                        Range("AL" & ActiveCell.Row) = TAmt * 0.5 'total volume
                        Range("AK" & ActiveCell.Row) = Range("AL" & ActiveCell.Row) - Range("AJ" & ActiveCell.Row) 'diluent
                        VialCountToCreateFull = VialCountToCreateFull * 0.5
                        If VialCountToCreateFull = 1 Then
                            Range("AJ" & ActiveCell.Row) = Range("AJ" & ActiveCell.Row) & " (" & VialCountToCreateFull & " vial)"
                        Else
                            Range("AJ" & ActiveCell.Row) = Range("AJ" & ActiveCell.Row) & " (" & VialCountToCreateFull & " vials)"
                        End If
                    ElseIf My3001InventorySum - (VialCountToCreateFull * 0.25) >= 1 Then '1/4 volume
                        Range("AJ" & ActiveCell.Row) = AntibodyAmt * 0.25 'stock volume
                        Range("AL" & ActiveCell.Row) = TAmt * 0.25 'total volume
                        Range("AK" & ActiveCell.Row) = Range("AL" & ActiveCell.Row) - Range("AJ" & ActiveCell.Row) 'diluent
                        VialCountToCreateFull = VialCountToCreateFull * 0.25
                        If VialCountToCreateFull = 1 Then
                            Range("AJ" & ActiveCell.Row) = Range("AJ" & ActiveCell.Row) & " (" & VialCountToCreateFull & " vial)"
                        Else
                            Range("AJ" & ActiveCell.Row) = Range("AJ" & ActiveCell.Row) & " (" & VialCountToCreateFull & " vials)"
                        End If
                    Else
                        Range("AJ" & ActiveCell.Row) = "not enough vials in bin"
                    End If
                End If
            Else
                Range("AJ" & ActiveCell.Row) = "Verify classic inventory or check characteristic values."
            End If
        End If
        If Right(Range("AC" & ActiveCell.Row), 2) = "-T" And Range("AI" & ActiveCell.Row) = "3001P" Then
            'MsgBox "The program is about to create or use a partial vial. Notify IT development to step through and improve this process."
            MyPartialClassicRow = 0
            On Error Resume Next
            MyPartialClassicRow = WorksheetFunction.Match(Range("AG" & ActiveCell.Row), Sheets("PartialVials").Range("A:A"), 0)
            On Error GoTo 0
            If MyPartialClassicRow <> 0 Then
                'MsgBox "The program is about to create or use a partial vial. Notify IT development to step through and improve this process."
                Range("AJ" & ActiveCell.Row) = "aliquot from 3001P"
                Sheets("PartialVials").Range("C" & MyPartialClassicRow) = Sheets("PartialVials").Range("C" & MyPartialClassicRow) - 0.1
                Dim LowPartialVials As String
                If Sheets("PartialVials").Range("C" & MyPartialClassicRow).Value < 0.1 Then LowPartialVials = LowPartialVials & ", " & LowPartialVials
            End If
        End If
        My3001InventorySum = 0
        If Right(Range("AC" & ActiveCell.Row), 2) = "-T" And Right(Range("AG" & ActiveCell.Row), 2) <> "-M" And Range("AI" & ActiveCell.Row) = "3001" Then
            'if there is more than one vial in the bin then
            My3001InventorySum = WorksheetFunction.SumIfs(Sheets("MB52").Range("F:F"), Sheets("MB52").Range("A:A"), Range("AG" & ActiveCell.Row), Sheets("MB52").Range("D:D"), "3001", Sheets("MB52").Range("C:C"), Range("AH" & ActiveCell.Row))
            If My3001InventorySum > 1 Then
                'create an instance of this vial on PartialVials
                'MsgBox "The program is about to create or use a partial vial. Notify IT development to step through and improve this process."
                Sheets("PartialVials").Range("A" & Sheets("PartialVials").Range("A10000").End(xlUp).Row + 1).Value = Range("AG" & ActiveCell.Row).Value
                Sheets("PartialVials").Range("B" & Sheets("PartialVials").Range("B10000").End(xlUp).Row + 1).Value = Range("AH" & ActiveCell.Row).Value
                Sheets("PartialVials").Range("C" & Sheets("PartialVials").Range("C10000").End(xlUp).Row + 1).Value = 0.9 - (0.1 * Range("AO" & ActiveCell.Row)) 'eaches in vial after trial removal(s)
                'when this order is created, (0.1 * order quantity) eaches will be removed from 3001...need to automatically remove the remainder of the vial as well
                'mark this record in an another column as one that should have an entire each removed from 3001
                Range("AJ" & ActiveCell.Row).Value = "Move vial 3001 > 3001P"
            Else
                Range("AJ" & ActiveCell.Row).Value = "Only 1 vial in bin"
            End If
        End If
        ActiveCell.Offset(1, 0).Select
    Wend
    ActiveWorkbook.Worksheets("Main").sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Main").sort.SortFields.Add Key:=Range("AC2:AC500"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Main").sort
        .SetRange Range("AB1:AP500")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Main").Range("AB2:AZ100000").Borders.LineStyle = xlNone
    Dim ListLength As Long
    ListLength = WorksheetFunction.CountA(Range("AC:AC"))
    With Range("AC2:" & "AP" & ListLength).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    'mark future dates
    Range("AE2").Select
    While ActiveCell.Row < Range("AE1000").End(xlUp).Row + 1
        If ActiveCell > Date Then ActiveCell.Font.Color = vbRed
        ActiveCell.Offset(1, 0).Select
        DoEvents
    Wend
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Range("A1").Select
    ActiveCell.Offset(1, 0).Activate
    Range("AC2").Select
    Range("ZZ1") = UserName
    Range("AAA1") = Password
    If MultOrders = True And RunTwice = True Then MsgBox "Some customers have purchased multiple vials of the same material!"
    Call LogOff
    Call VerifyThatComponentIsCorrect
    Exit Sub
NoZchrVals:
    MsgBox ("No zcharvalues for material: " & Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row))
    If LowPartialVials <> "" Then MsgBox LowPartialVials & " may have insufficient volume in 3001P. Please verify."
End Sub
Sub GetMB52Inventory()
    Dim MyRow As Long
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Sheets("MB52").Select
    Cells.Delete Shift:=xlUp
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "}"
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="}", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmb52"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/chkNOZERO").Selected = True
    Session.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = "3000"
    Session.FindById("wnd[0]/usr/chkXMCHB").Selected = True
    Session.FindById("wnd[0]/usr/ctxtP_VARI").Text = "/INVENTORY"
    Session.FindById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
    Session.FindById("wnd[1]/tbar[0]/btn[16]").press
    Sheets("BOMChild").Select
    Range("O2:O20000").ClearContents
    Range("G2:G10000").Copy
    Range("O2").Select
    ActiveSheet.Paste
    Range("A2:A10000").Copy
    Range("O20000").Select
    DoEvents
    Range("O20000").End(xlUp).Offset(1, 0).Select
    ActiveSheet.Paste
    Range("O2").Select
    While Range("O" & ActiveCell.Row) <> ""
        If Right(UCase(Range("O" & ActiveCell.Row)), 2) = "-M" Then
            MyRow = WorksheetFunction.CountA(Range("O:O")) + 1
            Range("O" & MyRow) = Left(Range("O" & ActiveCell.Row), Len(Range("O" & ActiveCell.Row)) - 2)
        End If
        ActiveCell.Offset(1, 0).Select
        DoEvents
    Wend
    Sheets("BOMChild").Range("O1:O" & ActiveCell.Row).Copy
    Sheets("MB52").Select
    Session.FindById("wnd[1]/tbar[0]/btn[24]").press
    Session.FindById("wnd[1]/tbar[0]/btn[8]").press
    Session.FindById("wnd[0]/tbar[1]/btn[8]").press
    Session.FindById("wnd[0]/tbar[1]/btn[45]").press
    Session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    ActiveSheet.PasteSpecial Format:="Unicode Text", Link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1)), TrailingMinusNumbers:=True
    Columns("A:A").Delete Shift:=xlToLeft
    Range("A1").Select: While Range("A" & ActiveCell.Row) = "": Range("A" & ActiveCell.Row).EntireRow.Delete: DoEvents: Wend: ActiveCell.Offset(1, 0).Select: If Range("A" & ActiveCell.Row) = "" Then Range("A" & ActiveCell.Row).EntireRow.Delete
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1:J1").Font.Bold = True
    Rows("1:1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Range("B1").ColumnWidth = 50
    Columns("A:J").Select
    ActiveWorkbook.Worksheets("MB52").sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MB52").sort.SortFields.Add Key:=Range( _
        "A2:A500000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MB52").sort
        .SetRange Range("A1:J500000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Call TrimAllCells
    'need to update to handle previous batch with sufficient inventory
    Range("M1") = "Unrestricted"
    Range("M2") = "<>0"
    Range("N1") = "Material"
    Range("O1") = "Sloc"
    Range("Q1") = "Min Batch"
    Range("N2") = "A301-672A"
    Range("O2") = "3002"
    Range("M4") = "Unrestricted"
    Range("M5").FormulaR1C1 = "=R2C13"
    Range("N4") = "Material"
    Range("O4") = "Sloc"
    Range("N5").FormulaR1C1 = "=IF(R2C14="""","""",R2C14)"
    Range("O5") = "3001"
    Range("M7") = "Unrestricted"
    Range("M8").FormulaR1C1 = "=R2C13"
    Range("N7") = "Material"
    Range("O7") = "Sloc"
    Range("N8").FormulaR1C1 = "=IF(R2C14="""","""",IF(RIGHT(R2C14,2)=""-M"",LEFT(R2C14,LEN(R2C14)-2),R2C14))"
    Range("O8") = "3002"
    Range("M10") = "Unrestricted"
    Range("M11").FormulaR1C1 = "=R2C13"
    Range("N10") = "Material"
    Range("O10") = "Sloc"
    Range("N11").FormulaR1C1 = "=IF(R2C14="""","""",IF(RIGHT(R2C14,2)=""-M"",LEFT(R2C14,LEN(R2C14)-2),R2C14))"
    Range("O11") = "3001"
    Range("M13") = "Unrestricted"
    Range("N13") = "Material"
    Range("O13") = "Sloc"
    Range("M14") = ">=3"
    Range("N14") = "A400DILUENT"
    Range("O14") = "3002"
    Range("Q14").FormulaR1C1 = "=DMIN(C1:C6,""Batch"",R13C13:R14C15)"
    Range("Q2").FormulaR1C1 = "=IF(DMIN(C1:C6,""Batch"",R1C13:R2C15)=0,"""",DMIN(C1:C6,""Batch"",R1C13:R2C15))"
    Range("Q5").FormulaR1C1 = "=IF(DMIN(C1:C6,""Batch"",R4C13:R5C15)=0,"""",DMIN(C1:C6,""Batch"",R4C13:R5C15))"
    Range("Q8").FormulaR1C1 = "=IF(DMIN(C1:C6,""Batch"",R7C13:R8C15)=0,"""",DMIN(C1:C6,""Batch"",R7C13:R8C15))"
    Range("Q11").FormulaR1C1 = "=IF(DMIN(C1:C6,""Batch"",R10C13:R11C15)=0,"""",DMIN(C1:C6,""Batch"",R10C13:R11C15))"
    Range("R2").FormulaR1C1 = "=IF(MIN(RC[-1],R[6]C[-1])=0,MIN(R[3]C[-1],R[9]C[-1]),MIN(RC[-1],R[6]C[-1]))"
    Range("S2").FormulaR1C1 = "=IF(R2C18=RC[-2],RC[-2],0)"
    Range("S5").FormulaR1C1 = "=IF(R2C18=RC[-2],RC[-2],0)"
    Range("S8").FormulaR1C1 = "=IF(R2C18=RC[-2],RC[-2],0)"
    Range("S11").FormulaR1C1 = "=IF(R2C18=RC[-2],RC[-2],0)"
End Sub
Sub ZVialOrders()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Call LogOff

    If Sheets("Main").Range("AA1") = "Production" Then Call LogonProduction Else Call LogonDevelopment
    Range("ZZ1") = UserName
    Range("AAA1") = Password
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nzvialorders"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/tbar[1]/btn[17]").press
    Session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").SelectedRows = "0"
    Session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").DoubleClickCurrentCell
    Session.FindById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press
    Session.FindById("wnd[1]/tbar[0]/btn[16]").press
    'old prod lines
    'Session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,0]").Text = "*-T"
    'Session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,1]").Text = "*A-M"
    'new prod lines...mbb 20151026...the 255 now causes an error in new prod
    Session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "*-T"
    Session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "*A-M"

    Session.FindById("wnd[1]/tbar[0]/btn[8]").press
    Session.FindById("wnd[0]/tbar[1]/btn[8]").press
    Session.FindById("wnd[0]/tbar[1]/btn[45]").press
    Session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    Session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
    Session.FindById("wnd[1]/tbar[0]/btn[0]").press
    Sheets("ZVialOrders").Select
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "#"
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="#", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Range("A1").Select
    ActiveSheet.PasteSpecial Format:="Unicode Text", Link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 9), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, _
        2), Array(6, 2), Array(7, 2), Array(8, 3), Array(9, 2), Array(10, 2), Array(11, 2), Array(12 _
        , 2), Array(13, 2), Array(14, 2), Array(15, 2), Array(16, 2), Array(17, 2)), _
        TrailingMinusNumbers:=True
    While ActiveCell = "": Rows("1:1").Delete Shift:=xlUp: DoEvents: Wend
        Rows("2:2").Delete Shift:=xlUp
        Call TrimAllCells
        Rows("1:1").Select
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.399975585192419
            .PatternTintAndShade = 0
        End With
        Cells.Select
        Cells.EntireColumn.AutoFit
        Range("A:U").Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
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
    Sub SQ01_BOMComponent()
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        If Trim(Sheets("ZVialOrders").Range("A2")) = "List contains no data" Then
            Sheets("Main").Activate
            MsgBox ("No data to display")
            End
        End If
        Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nsq01"
        Session.FindById("wnd[0]").SendVKey 0
        Session.FindById("wnd[0]/mbar/menu[5]/menu[0]").Select
        Session.FindById("wnd[1]/usr/radRAD1").Select
        Session.FindById("wnd[1]/tbar[0]/btn[2]").press
        Session.FindById("wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell").PressToolbarButton "&FIND"
        Session.FindById("wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "BOM-COMPONENT"
        Session.FindById("wnd[1]").SendVKey 0
        Session.FindById("wnd[1]/tbar[0]/btn[12]").press
        AppActivate Application.Caption
        If Session.FindById("wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell").Text = "BOM-COMPONENT" Then MsgBox ("BOM-COMPONENT Query is Missing"): Exit Sub
        Session.FindById("wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell").SelectedRows = "1"
        Session.FindById("wnd[0]").SendVKey 8
        Session.FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH").press
        Session.FindById("wnd[1]/tbar[0]/btn[16]").press
        Sheets("ZVialOrders").Range("J2:J10000").Copy
        Range("Z1").Select
        ActiveSheet.Paste
        Range("Z1").Select
        Selection.End(xlDown).Select
        If ActiveCell = "" Then Range("Z2").Select Else ActiveCell.Offset(1, 0).Select
        ActiveSheet.Paste
        Selection.Replace What:="-T", Replacement:="-M", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        Range("Z1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Application.CutCopyMode = False
        Selection.Copy
        Session.FindById("wnd[1]/tbar[0]/btn[24]").press
        Session.FindById("wnd[1]/tbar[0]/btn[8]").press
        Session.FindById("wnd[0]").SendVKey 8
        Session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
        Session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").SelectContextMenuItem "&PC"
        Session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
        Session.FindById("wnd[1]/tbar[0]/btn[0]").press
        Sheets("BOMChild").Select
        Cells.ClearContents
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "#"
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="#", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
        Range("A1").Select
        ActiveSheet.PasteSpecial Format:="Unicode Text", Link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True
        Columns("A:A").Select
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 9), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, _
        2), Array(6, 2), Array(7, 2), Array(8, 3), Array(9, 2), Array(10, 2), Array(11, 2), Array(12 _
        , 2), Array(13, 2), Array(14, 2), Array(15, 2), Array(16, 2), Array(17, 2)), _
        TrailingMinusNumbers:=True
        While ActiveCell = "": Rows("1:1").Delete Shift:=xlUp: DoEvents: Wend
            Rows("2:2").Delete Shift:=xlUp
            Call TrimAllCells
            Rows("1:1").Select
            Selection.Font.Bold = True
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0.399975585192419
                .PatternTintAndShade = 0
            End With
            Cells.Select
            Cells.EntireColumn.AutoFit
            Range("A:L").Select
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            Range("A" & Range("A2").End(xlDown).Row + 1 & ":L1000000").ClearContents
        End Sub
        Sub ZCharValues()
            Dim MyVal As String
            Dim HeaderName(30) As String
            Dim MyCount, i As Double
            Application.ScreenUpdating = False
            If Sheets("Main").Range("AA1") = "Production" Then Call LogonProduction Else Call LogonDevelopment
            Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nzcharvalues"
            Session.FindById("wnd[0]").SendVKey 0
            i = 0
            MyCount = 0
            HeaderName(0) = "Material Number"
            HeaderName(1) = "Batch"
            HeaderName(2) = "Material Description"
            HeaderName(3) = "V_WBCONC"
            HeaderName(4) = "V_DilutionFactor"
            HeaderName(5) = "Ab_Stock_Concentration"
            'Session.FindById("wnd[0]/usr/ctxtP_VARI").Text = "VIALING"
            'couldn't figure out how to edit existing variant...made a new one called VIALING, without the slash
            Session.FindById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = "A300-003A" 'must enter to be able enter list mode
            Session.FindById("wnd[0]/usr/radR_BCHAR").Select
            Session.FindById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press 'enter list mode
            Session.FindById("wnd[1]/tbar[0]/btn[16]").press '
            Sheets("MB52").Range("A2:A10000").Copy 'copy inventory
            Session.FindById("wnd[1]/tbar[0]/btn[24]").press 'paste inventory list
            Session.FindById("wnd[1]/tbar[0]/btn[8]").press
            Session.FindById("wnd[0]").SendVKey 8 'submit
            Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").CurrentCellRow = -1
            Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").SelectColumn "MATNR"
            Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").ContextMenu
            Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").SelectContextMenuItem "&COL0"
            Session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").SelectAll
            Session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_FL_SING").press
            For i = 0 To UBound(HeaderName)
                If HeaderName(i) <> "" Then MyCount = MyCount + 1
            Next i
            i = 0
            For i = 0 To MyCount - 1
                Session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").PressToolbarButton "&FIND"
                Session.FindById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").Selected = True
                Session.FindById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").SetFocus
                Session.FindById("wnd[2]/usr/cmbGS_SEARCH-SEARCH_ORDER").Key = "0"
                Session.FindById("wnd[2]/usr/txtGS_SEARCH-VALUE").Text = HeaderName(i)
                Session.FindById("wnd[2]/tbar[0]/btn[0]").press
                Session.FindById("wnd[2]/tbar[0]/btn[12]").press
                Session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press
            Next i
            Session.FindById("wnd[1]/tbar[0]/btn[0]").press
            Session.FindById("wnd[0]/tbar[1]/btn[45]").press 'clipboard
            Session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
            Session.FindById("wnd[1]/tbar[0]/btn[0]").press
            Sheets("ZCharValues").Select
            Cells.Select
            Selection.ClearContents
            Range("A1").Select
            ActiveCell.FormulaR1C1 = "#"
            Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="#", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
            Range("A1").Select
            ActiveSheet.PasteSpecial Format:="Unicode Text", Link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True
            Columns("A:A").Select
            Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 9), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, _
        2), Array(6, 2), Array(7, 2), Array(8, 3), Array(9, 2), Array(10, 2), Array(11, 2), Array(12 _
        , 2), Array(13, 2), Array(14, 2), Array(15, 2), Array(16, 2), Array(17, 2)), _
        TrailingMinusNumbers:=True
            While ActiveCell = "": Rows("1:1").Delete Shift:=xlUp: DoEvents: Wend
                Rows("2:2").Delete Shift:=xlUp
                Call TrimAllCells
                Rows("1:1").Select
                Selection.Font.Bold = True
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.399975585192419
                    .PatternTintAndShade = 0
                End With
                Cells.Select
                Cells.EntireColumn.AutoFit
                Range("A:F").Select
                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                With Selection.Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                Range("O1") = "Material Number"
                Range("P1") = "Batch"
                Range("R1") = "V_WBCONC"
                Range("S1") = "Ab_Stock_Concentration"
                Call VerifyM_inSAP
                Call CheckReformulationStatus
                'Call LogOff
            End Sub
            Function NMIGO_TR(OMat, OBNum, OSLoc, NMat, NBNum, NQty, NSLoc)
                'OMat=Material Number, OBNum=Batch Number, OSLoc = Previous SLoc, NMat=Material Number, NBNum=bathc number, NQty = quantity to move, NSLoc = New SLoc
                Dim MyVal As String
                UserName = Range("ZZ1")
                Password = Range("AAA1")
                If Sheets("Main").Range("AA1") = "Production" Then Call LogonProduction Else Call LogonDevelopment
                Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmigo_tr"
                Session.FindById("wnd[0]").SendVKey 0
                On Error Resume Next
                Session.FindById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").TopNode = "          5" 'what does this do exactly??...has to do with the tree
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0009/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A08"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0008/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A08"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A08"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A08"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0005/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A08"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A08"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A08"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A08"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A08"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0009/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R10"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0008/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R10"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R10"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R10"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0005/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R10"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R10"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R10"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R10"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R10"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0009/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").Text = "311"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0008/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").Text = "311"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").Text = "311"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").Text = "311"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0005/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").Text = "311"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").Text = "311"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").Text = "311"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").Text = "311"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").Text = "311"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0009/subSUB_ITEMDETAIL:SAPLMIGO:0303/btnBUTTON_ITEMDETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0008/subSUB_ITEMDETAIL:SAPLMIGO:0303/btnBUTTON_ITEMDETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/btnBUTTON_ITEMDETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_ITEMDETAIL:SAPLMIGO:0303/btnBUTTON_ITEMDETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0005/subSUB_ITEMDETAIL:SAPLMIGO:0303/btnBUTTON_ITEMDETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_ITEMDETAIL:SAPLMIGO:0303/btnBUTTON_ITEMDETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0303/btnBUTTON_ITEMDETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMDETAIL:SAPLMIGO:0303/btnBUTTON_ITEMDETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001/subSUB_ITEMDETAIL:SAPLMIGO:0303/btnBUTTON_ITEMDETAIL").press
                'opens detail window
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0009/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0008/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0005/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL").press
                'opens header window
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0009/subSUB_HEADER:SAPLMIGO:0102/btnOK_HEADER").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0008/subSUB_HEADER:SAPLMIGO:0102/btnOK_HEADER").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_HEADER:SAPLMIGO:0102/btnOK_HEADER").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_HEADER:SAPLMIGO:0102/btnOK_HEADER").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0005/subSUB_HEADER:SAPLMIGO:0102/btnOK_HEADER").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_HEADER:SAPLMIGO:0102/btnOK_HEADER").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_HEADER:SAPLMIGO:0102/btnOK_HEADER").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0102/btnOK_HEADER").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001/subSUB_HEADER:SAPLMIGO:0102/btnOK_HEADER").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-MAKTX[1,0]").SetFocus
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-MAKTX[1,0]").CaretPosition = 0
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL").press
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").SetFocus
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").CaretPosition = 0
                Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmigo_tr"
                Session.FindById("wnd[0]").SendVKey 0
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").SetFocus
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").Text = OMat '"AS1206/207/208"
                Session.FindById("wnd[0]").SendVKey 4
                Session.FindById("wnd[1]/tbar[0]/btn[12]").press 'make conditional here
                Session.FindById("wnd[1]/tbar[0]/btn[71]").press
                Session.FindById("wnd[0]").SendVKey 0
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-LGOBE").Text = OSLoc '"3005"
                Session.FindById("wnd[0]").SendVKey 0
                'destination material already filled now (field is greyed out)
                'Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGOITEM-UMMAKTX").Text = NMat '"AS1206/207/208"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGOITEM-UMLGOBE").Text = NSLoc
                Session.FindById("wnd[0]").SendVKey 0
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/txtGODYNPRO-ERFMG").Text = NQty '"165"
                Session.FindById("wnd[0]").SendVKey 0
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-CHARG").Text = NBNum '"L100601"
                Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-UMCHA").Text = OBNum '"05-07/04"
                Session.FindById("wnd[0]/tbar[1]/btn[23]").press
                NMIGO_TR = Session.FindById("wnd[1]/usr/lbl[10,3]").Text
                If NMIGO_TR = "" Then NMIGO_TR = Session.FindById("wnd[0]/sbar").Text
            End Function
            Sub RetrieveData()
                Dim BartenderListLength As Long, FullPath, SheetName, BookName, Rnge
                BartenderListLength = ExecuteExcel4Macro("COUNTA('\\BETHYL-SERVER2\Vialing\BarTenderLabels\[Bartenderlabels.xlsx]Sheet1'!R1C1:R50000C1)")
                ThisWorkbook.Sheets("LabelData").Cells.ClearContents
                FullPath = "\\BETHYL-SERVER2\Vialing\BarTenderLabels" 'location of file that you want to pull data from
                Rnge = "A1:N" & BartenderListLength 'range that you want to pull data from
                SheetName = "Sheet1" 'name of sheet that you want to pull data from
                BookName = "Bartenderlabels.xlsx" 'name of workbook that you want to pull data from
                GetValuesFromAClosedWorkbook FullPath, BookName, SheetName, Rnge
            End Sub
            Private Sub GetValuesFromAClosedWorkbook(fPath, fName, sName, cellRange)
                With ThisWorkbook.Sheets("LabelData").Range(cellRange)
                    .FormulaArray = "='" & fPath & "\[" & fName & "]" & sName & "'!" & cellRange
                    .Value = .Value
                End With
            End Sub
            Sub VerifyM_inSAP()
                Dim Mcheck As Integer, MProduct As String
                Range("A2").Select
                While Range("A" & ActiveCell.Row) <> "" 'populate the concatenated column
                    ActiveCell.Offset(0, 6).Value = ActiveCell.Value & "_" & ActiveCell.Offset(0, 1).Value
                    ActiveCell.Offset(1, 0).Select
                    DoEvents
                Wend
                Range("A2").Select
                While Range("A" & ActiveCell.Row) <> ""
                    Mcheck = 0
                    If InStr(ActiveCell, "-M") = 0 And ActiveCell.Offset(0, 3).Value <> "" Then 'if the classic has a zcharvalue, then an M for that batch must exist exist in this list
                        MProduct = Replace(Replace(ActiveCell, "A100", "A"), "A50", "A") & "-M"
                        On Error Resume Next
                        Mcheck = WorksheetFunction.Match(MProduct & "_" & ActiveCell.Offset(0, 1).Value, Range("G2:G" & Range("A10000").End(xlUp).Row), 0)
                        On Error GoTo 0
                        If Mcheck = 0 Then Range("H" & ActiveCell.Row).Value = MProduct & " batch " & ActiveCell.Offset(0, 1).Value & " does not have inventory in SAP."
                    End If
                    ActiveCell.Offset(1, 0).Select
                    DoEvents
                Wend
            End Sub
            Sub VerifyThatComponentIsCorrect()
                Dim ClassicNum_Batch As String, MNum_Batch As String, NumericalA300 As Double, NoWB As Boolean, YesWB As Boolean, CompCheck
                Range("AC2").Select
                While Range("AC" & ActiveCell.Row) <> ""
                    CompCheck = 0
                    ClassicNum_Batch = Left(ActiveCell.Value, 9) & "_" & Range("AH" & ActiveCell.Row).Value
                    MNum_Batch = Left(ActiveCell.Value, 9) & "-M_" & Range("AH" & ActiveCell.Row).Value
                    If Left(ActiveCell.Value, 3) = "A30" Then NumericalA300 = CDbl(Trim(Replace(Replace(Replace(Replace(ActiveCell.Value, "A30", ""), "A-T", ""), "A-M", ""), "-", "")))
                    If Left(ActiveCell.Value, 9) = "A301-985A" Then ClassicNum_Batch = "A301-985A100_" & Range("AH" & ActiveCell.Row).Value
                    If Right(Range("AG" & ActiveCell.Row).Value, 2) <> "-M" And Right(Range("AC" & ActiveCell.Row).Value, 2) = "-M" Then 'if component is classic and an M is being made
                        On Error Resume Next
                        CompCheck = WorksheetFunction.Index(Sheets("Zcharvalues").Range("D:D"), WorksheetFunction.Match(ClassicNum_Batch, Sheets("Zcharvalues").Range("G:G"), 0))
                        On Error GoTo 0
                        If CompCheck = "" Then CompCheck = 0
                        If CompCheck = 0 Then 'if the classic batch doesn't have a WBCONC in MSC2N then there is an issue
                            ActiveCell.Offset(0, -2).Value = "Batch " & Range("AH" & ActiveCell.Row).Value & " doesn't WB!"
                            NoWB = True
                        End If
                    ElseIf Right(Range("AG" & ActiveCell.Row).Value, 2) = "-M" And Right(Range("AC" & ActiveCell.Row).Value, 2) = "-T" And NumericalA300 >= 5576 Then 'if component is M and a trial is being made and material was created after the new M-only system
                        On Error Resume Next
                        CompCheck = WorksheetFunction.Index(Sheets("Zcharvalues").Range("D:D"), WorksheetFunction.Match(MNum_Batch, Sheets("Zcharvalues").Range("G:G"), 0))
                        On Error GoTo 0
                        If CompCheck = "" Then CompCheck = 0
                        If CompCheck = 0 And WorksheetFunction.CountIf(Sheets("Zcharvalues").Range("A:A"), Left(MNum_Batch, InStr(MNum_Batch, "_"))) > 0 Then 'if there is batch of this classic and it doesn't have a WBCONC in MSC2N then there is an issue
                            ActiveCell.Offset(0, -2).Value = "Batch " & Range("AH" & ActiveCell.Row).Value & " classic lacks WBConc value!"
                            NoWB = True
                        ElseIf CompCheck = 0 And WorksheetFunction.CountIf(Sheets("Zcharvalues").Range("A:A"), Left(MNum_Batch, InStr(MNum_Batch, "_"))) = 0 Then  'batch of this classic is depleted
                            ActiveCell.Offset(0, -2).Value = "Batch " & Range("AH" & ActiveCell.Row).Value & " classic out of stock"
                            NoWB = True
                        End If
                    ElseIf Right(Range("AG" & ActiveCell.Row).Value, 2) = "-M" And Right(Range("AC" & ActiveCell.Row).Value, 2) = "-T" Then 'if component is M and a trial is being made and material was created before the new M-only system
                        On Error Resume Next
                        CompCheck = WorksheetFunction.Index(Sheets("Zcharvalues").Range("D:D"), WorksheetFunction.Match(ClassicNum_Batch, Sheets("Zcharvalues").Range("G:G"), 0))
                        On Error GoTo 0
                        If CompCheck = "" Then CompCheck = 0
                        If CompCheck = 0 And WorksheetFunction.CountIf(Sheets("Zcharvalues").Range("A:A"), Left(ClassicNum_Batch, InStr(ClassicNum_Batch, "_"))) > 0 Then 'if there is batch of this classic and it doesn't have a WBCONC in MSC2N then there is an issue
                            ActiveCell.Offset(0, -2).Value = "Batch " & Range("AH" & ActiveCell.Row).Value & " classic lacks WBConc value!"
                            NoWB = True
                        ElseIf CompCheck = 0 And WorksheetFunction.CountIf(Sheets("Zcharvalues").Range("A:A"), Left(ClassicNum_Batch, InStr(ClassicNum_Batch, "_"))) = 0 Then  'batch of this classic is depleted
                            ActiveCell.Offset(0, -2).Value = "Batch " & Range("AH" & ActiveCell.Row).Value & " classic out of stock"
                            NoWB = True
                        End If
                    ElseIf Right(Range("AG" & ActiveCell.Row).Value, 2) <> "-M" And Right(Range("AC" & ActiveCell.Row).Value, 2) = "-T" Then 'if component is classic and a trial is being made
                        On Error Resume Next
                        CompCheck = WorksheetFunction.Index(Sheets("Zcharvalues").Range("D:D"), WorksheetFunction.Match(ClassicNum_Batch, Sheets("Zcharvalues").Range("G:G"), 0))
                        On Error GoTo 0
                        If CompCheck = "" Then CompCheck = 0
                        'need to change this to account for the fact that some classics are 0.2mg/ml...use the dilution factor instead
                        If CompCheck < 1 And CompCheck > 0 Then 'if the classic batch has a WBCONC < 1痢/ml in MSC2N then there is an issue
                            ActiveCell.Offset(0, -2).Value = "Batch " & Range("AH" & ActiveCell.Row).Value & " western blots @ < 1痢/ml!"
                            YesWB = True
                        End If
                    End If
                    ActiveCell.Offset(1, 0).Select
                    DoEvents
                Wend
                '    If NoWB Then MsgBox "Some of the M-sourced trials list a component batch that is not qualified for WB." & vbCrLf & vbCrLf & "This is likely because the previous batch does not qualify for WB and the new one does (and the BOM for the trial product now lists the M product as the component)." & vbCrLf & vbCrLf & _
                '    "Please verify that you are creating the trial using the appropriate stock vial (classic or M)."
                '    If YesWB Then MsgBox "Some of the classic-sourced trials list a component batch that is qualified for WB." & vbCrLf & vbCrLf & "This is likely because the previous batch qualifies for WB and the new one does not (and the BOM for the trial product now lists the classic product as the component)." & vbCrLf & vbCrLf & _
                '    "Please verify that you are creating the trial using the appropriate stock vial (classic or M)."

                If NoWB Or YesWB Then MsgBox "There may be issues to address with one or more of these orders (incorrect BOM, classic depleted, new batch applications changed)" & vbCrLf & vbCrLf & "Please view the comments in the left margin."
                Range("AC2").Select
            End Sub
            Sub CheckReformulationStatus()
                'called from zcharvalues routine...pulls short text (if present) for each material and marks the ones that have been reformulated (and which reformulation is active)
                Range("A2").Select
                While Range("A" & ActiveCell.Row) <> ""
                    If InStr(Range("A" & ActiveCell.Row), "A400DILUENT") > 0 Then Exit Sub
                    If InStr(Range("A" & ActiveCell.Row), "-M") = 0 Then
                        Range("I" & ActiveCell.Row) = CheckMSC2NShortText(Range("A" & ActiveCell.Row), Range("B" & ActiveCell.Row))
                        If Range("I" & ActiveCell.Row) = "" Then 'if no text, then never reformulated...leave  column J blank
                        ElseIf InStr(Range("I" & ActiveCell.Row), "_") = 0 Then 'if text but no underscore, then reformulated once
                            Range("J" & ActiveCell.Row) = 1
                        ElseIf Len(Range("I" & ActiveCell.Row)) - Len(Replace(Range("I" & ActiveCell.Row), "_", "")) = 1 Then 'if one underscore, then reformulated twice
                            Range("J" & ActiveCell.Row) = 2
                        ElseIf Len(Range("I" & ActiveCell.Row)) - Len(Replace(Range("I" & ActiveCell.Row), "_", "")) = 2 Then 'if two underscores, then reformulated three times (not likely)
                            Range("J" & ActiveCell.Row) = 3
                        End If
                    End If
                    DoEvents
                    ActiveCell.Offset(1, 0).Select
                Wend
            End Sub
            Function CheckMSC2NShortText(MatNum As String, BatchNum As String)
                Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmsc2n"
                Session.FindById("wnd[0]").SendVKey 0
                Session.FindById("wnd[0]/usr/subSUBSCR_BATCH_MASTER:SAPLCHRG:1111/subSUBSCR_HEADER:SAPLCHRG:1501/ctxtDFBATCH-MATNR").Text = MatNum
                Session.FindById("wnd[0]/usr/subSUBSCR_BATCH_MASTER:SAPLCHRG:1111/subSUBSCR_HEADER:SAPLCHRG:1501/ctxtDFBATCH-WERKS").Text = "3000"
                Session.FindById("wnd[0]/usr/subSUBSCR_BATCH_MASTER:SAPLCHRG:1111/subSUBSCR_HEADER:SAPLCHRG:1501/ctxtDFBATCH-CHARG").Text = BatchNum
                Session.FindById("wnd[0]").SendVKey 0
                Session.FindById("wnd[0]/usr/subSUBSCR_BATCH_MASTER:SAPLCHRG:1111/subSUBSCR_TABSTRIP:SAPLCHRG:2000/tabsTS_BODY/tabpSNST").Select 'basic data 2
                CheckMSC2NShortText = Session.FindById("wnd[0]/usr/subSUBSCR_BATCH_MASTER:SAPLCHRG:1111/subSUBSCR_TABSTRIP:SAPLCHRG:2000/tabsTS_BODY/tabpSNST/ssubSUBSCR_BODY:SAPLCHRG:2200/txtDFBATCH-KZTXT").Text
            End Function
