Attribute VB_Name = "Module2"
Sub Macro1()
    '
    ' Macro1 Macro
    '

    '
    ActiveWorkbook.Queries.Add Name:="KSSK_INOB (2)", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    KSSK_INOB_Sheet = Source{[Item=""KSSK_INOB"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(KSSK_INOB_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""Object"", Int64.Type}, {""MatNum""," & _
        " type text}, {""Batch"", type text}})," & Chr(10) & "    #""Added Custom"" = Table.AddColumn(#""Changed Type"", ""Custom"", each [MatNum]&""_""&[#""Batch""])," & Chr(10) & "    #""Reordered Columns"" = Table.ReorderColumns(#""Added Custom"",{""Object"", ""Custom"", ""MatNum"", ""Batch""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Reordered Columns"",{{""Custom"", ""MatNum_Batch""}})," & Chr(10) & "" & _
        "    #""Merged Queries"" = Table.NestedJoin(#""Renamed Columns"",{""Object""},#""Animals used to make batch"",{""Object""},""NewColumn"")," & Chr(10) & "    #""Expanded NewColumn"" = Table.ExpandTableColumn(#""Merged Queries"", ""NewColumn"", {""Animals used to make batch""}, {""Animals used to make batch""})," & Chr(10) & "    #""Merged Queries1"" = Table.NestedJoin(#""Expanded NewColumn"",{""" & _
        "Object""},Ab_Stock_Concentration,{""Object""},""NewColumn"")," & Chr(10) & "    #""Expanded NewColumn1"" = Table.ExpandTableColumn(#""Merged Queries1"", ""NewColumn"", {""Ab_Stock_Concentration""}, {""Ab_Stock_Concentration""})," & Chr(10) & "    #""Merged Queries2"" = Table.NestedJoin(#""Expanded NewColumn1"",{""Object""},V_WBCONC,{""Object""},""NewColumn"")," & Chr(10) & "    #""Expanded NewColumn2"" = Ta" & _
        "ble.ExpandTableColumn(#""Merged Queries2"", ""NewColumn"", {""V_WBCONC""}, {""V_WBCONC""})," & Chr(10) & "    #""Merged Queries3"" = Table.NestedJoin(#""Expanded NewColumn2"",{""Object""},V_DilutionFactor,{""Object""},""NewColumn"")," & Chr(10) & "    #""Expanded NewColumn3"" = Table.ExpandTableColumn(#""Merged Queries3"", ""NewColumn"", {""V_DilutionFactor""}, {""V_DilutionFactor""})," & Chr(10) & "    #""" & _
        "Merged Queries4"" = Table.NestedJoin(#""Expanded NewColumn3"",{""Object""},V_IPCONC,{""Object""},""NewColumn"")," & Chr(10) & "    #""Expanded NewColumn4"" = Table.ExpandTableColumn(#""Merged Queries4"", ""NewColumn"", {""V_IPCONC""}, {""V_IPCONC""})," & Chr(10) & "    #""Merged Queries5"" = Table.NestedJoin(#""Expanded NewColumn4"",{""Object""},QC_Last_WB_Inspection_Date,{""Object""},""NewCol" & _
        "umn"")," & Chr(10) & "    #""Expanded NewColumn5"" = Table.ExpandTableColumn(#""Merged Queries5"", ""NewColumn"", {""QC_Last_WB_Inspection_Date""}, {""QC_Last_WB_Inspection_Date""})," & Chr(10) & "    #""Merged Queries6"" = Table.NestedJoin(#""Expanded NewColumn5"",{""Object""},QC_Last_IHC_Inspection_Date,{""Object""},""NewColumn"")," & Chr(10) & "    #""Expanded NewColumn6"" = Table.ExpandTableColumn(#""Me" & _
        "rged Queries6"", ""NewColumn"", {""QC_Last_IHC_Inspection_Date""}, {""QC_Last_IHC_Inspection_Date""})," & Chr(10) & "    #""Merged Queries7"" = Table.NestedJoin(#""Expanded NewColumn6"",{""Object""},QC_TestingBrother,{""Object""},""NewColumn"")," & Chr(10) & "    #""Expanded NewColumn7"" = Table.ExpandTableColumn(#""Merged Queries7"", ""NewColumn"", {""QC_TestingBrother""}, {""QC_TestingBrothe" & _
        "r""})," & Chr(10) & "    #""Merged Queries8"" = Table.NestedJoin(#""Expanded NewColumn7"",{""Object""},QC_QualifiedApplications,{""Object""},""NewColumn"")," & Chr(10) & "    #""Expanded NewColumn8"" = Table.ExpandTableColumn(#""Merged Queries8"", ""NewColumn"", {""QC_QualifiedApplications""}, {""QC_QualifiedApplications""})," & Chr(10) & "    #""Merged Queries9"" = Table.NestedJoin(#""Expanded NewColumn8""" & _
        ",{""Object""},QC_TestResult,{""Object""},""NewColumn"")," & Chr(10) & "    #""Expanded NewColumn9"" = Table.ExpandTableColumn(#""Merged Queries9"", ""NewColumn"", {""QC_TestResult""}, {""QC_TestResult""})," & Chr(10) & "    #""Merged Queries10"" = Table.NestedJoin(#""Expanded NewColumn9"",{""Object""},QC_Tested_Cell_Line,{""Object""},""NewColumn"")," & Chr(10) & "    #""Expanded NewColumn10"" = Table.Expand" & _
        "TableColumn(#""Merged Queries10"", ""NewColumn"", {""QC_Tested_Cell_Line""}, {""QC_Tested_Cell_Line""})," & Chr(10) & "    #""Merged Queries11"" = Table.NestedJoin(#""Expanded NewColumn10"",{""Object""},QC_Cell_Treatment,{""Object""},""NewColumn"")," & Chr(10) & "    #""Expanded NewColumn11"" = Table.ExpandTableColumn(#""Merged Queries11"", ""NewColumn"", {""QC_Cell_Treatment""}, {""QC_Cell_Tr" & _
        "eatment""})," & Chr(10) & "    #""Merged Queries12"" = Table.NestedJoin(#""Expanded NewColumn11"",{""Object""},QC_Qualified_Cell_Lines,{""Object""},""NewColumn"")," & Chr(10) & "    #""Expanded NewColumn12"" = Table.ExpandTableColumn(#""Merged Queries12"", ""NewColumn"", {""QC_Qualified_Cell_Lines""}, {""QC_Qualified_Cell_Lines""})," & Chr(10) & "    #""Merged Queries13"" = Table.NestedJoin(#""Expanded NewC" & _
        "olumn12"",{""Object""},QC_TestedApplication,{""Object""},""NewColumn"")," & Chr(10) & "    #""Expanded NewColumn13"" = Table.ExpandTableColumn(#""Merged Queries13"", ""NewColumn"", {""QC_TestedApplication""}, {""QC_TestedApplication""})," & Chr(10) & "    #""Merged Queries14"" = Table.NestedJoin(#""Expanded NewColumn13"",{""Object""},QC_Datasheet_Low_Dilution,{""Object""},""NewColumn"")," & Chr(10) & "    #" & _
        """Expanded NewColumn14"" = Table.ExpandTableColumn(#""Merged Queries14"", ""NewColumn"", {""QC_Datasheet_Low_Dilution""}, {""QC_Datasheet_Low_Dilution""})," & Chr(10) & "    #""Merged Queries15"" = Table.NestedJoin(#""Expanded NewColumn14"",{""Object""},QC_Datasheet_High_Dilution,{""Object""},""NewColumn"")," & Chr(10) & "    #""Expanded NewColumn15"" = Table.ExpandTableColumn(#""Merged Querie" & _
        "s15"", ""NewColumn"", {""QC_Datasheet_High_Dilution""}, {""QC_Datasheet_High_Dilution""})" & Chr(10) & "in" & Chr(10) & "    #""Expanded NewColumn15"""
    Sheets.Add After:=ActiveSheet
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""KSSK_INOB (2)""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [KSSK_INOB (2)]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = False
        .ListObject.DisplayName = "KSSK_INOB__2"
        .Refresh BackgroundQuery:=False
    End With
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
End Sub
Sub Macro2()
    '
    ' Macro2 Macro
    '

    '
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
End Sub
