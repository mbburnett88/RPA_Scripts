Attribute VB_Name = "Module1"
Sub UpdateSourceQueries()
    On Error Resume Next
    ActiveWorkbook.Queries("Ab_Stock_Concentration").Delete
    ActiveWorkbook.Queries("V_WBCONC").Delete
    ActiveWorkbook.Queries("V_DilutionFactor").Delete
    ActiveWorkbook.Queries("V_IPCONC").Delete
    ActiveWorkbook.Queries("QC_Last_WB_Inspection_Date").Delete
    ActiveWorkbook.Queries("QC_Last_IHC_Inspection_Date").Delete
    ActiveWorkbook.Queries("QC_TestingBrother").Delete
    ActiveWorkbook.Queries("QC_QualifiedApplications").Delete
    ActiveWorkbook.Queries("QC_TestResult").Delete
    ActiveWorkbook.Queries("QC_Tested_Cell_Line").Delete
    ActiveWorkbook.Queries("QC_Molecular_Weight").Delete
    ActiveWorkbook.Queries("QC_Gel_Type").Delete
    ActiveWorkbook.Queries("QC_Cell_Treatment").Delete
    ActiveWorkbook.Queries("QC_Qualified_Cell_Lines").Delete
    ActiveWorkbook.Queries("QC_TestedApplication").Delete
    ActiveWorkbook.Queries("QC_Datasheet_Low_Dilution").Delete
    ActiveWorkbook.Queries("QC_Datasheet_High_Dilution").Delete
    ActiveWorkbook.Queries("Animals used to make batch").Delete
    ActiveWorkbook.Queries("Uniprot_Accession_Number").Delete
    On Error GoTo 0
    ActiveWorkbook.Queries.Add Name:="Animals used to make batch", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    #""Animals used to make batch_Sheet"" = Source{[Item=""Animals used to make batch"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(#""Animals used to make batch_Sheet"")," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(" & _
        "#""Promoted Headers"",{{""Object"", Int64.Type}, {""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", type text}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Trimmed Text"" = Table.TransformColumns(#""Changed Type"",{{""Char. Value"", Text.Trim}})," & Chr(10) & "    #""Grouped Rows"" = Table.Group(#""Trimmed Text"", {""Object""}," & _
        " {{""Count"", each _, type table}})," & Chr(10) & "    #""Added Custom"" = Table.AddColumn(#""Grouped Rows"", ""Custom"", each Table.ToList(" & Chr(10) & "         Table.Transpose(" & Chr(10) & "              Table.FromList(" & Chr(10) & "                   Table.Column([Count], ""Char. Value"")" & Chr(10) & "                  )" & Chr(10) & "                  )," & Chr(10) & "    Combiner.CombineTextByDelimiter("", "")" & Chr(10) & "    ))," & Chr(10) & "    #""Expanded Custom"" = Table.E" & _
        "xpandListColumn(#""Added Custom"", ""Custom"")," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Expanded Custom"",{{""Custom"", ""Animals used to make batch""}})," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Renamed Columns"",{""Count""})" & Chr(10) & "in" & Chr(10) & "    #""Removed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - Animals used to make batch", _
        "Connection to the 'Animals used to make batch' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Animals used to make batch""" _
        , "SELECT * FROM [Animals used to make batch]", 2

    ActiveWorkbook.Queries.Add Name:="Ab_Stock_Concentration", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    Ab_Stock_Concentration_Sheet = Source{[Item=""Ab_Stock_Concentration"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(Ab_Stock_Concentration_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers""," & _
        "{{""Object"", Int64.Type}, {""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", type text}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Changed Type"",{""Counter"", ""Class Type"", ""Internal char."", ""Value from""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Remov" & _
        "ed Columns"",{{""Char. Value"", ""Ab_Stock_Concentration""}})" & Chr(10) & "in" & Chr(10) & "    #""Renamed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - Ab_Stock_Concentration", _
        "Connection to the 'Ab_Stock_Concentration' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Ab_Stock_Concentration" _
        , "SELECT * FROM [Ab_Stock_Concentration]", 2

    ActiveWorkbook.Queries.Add Name:="V_WBCONC", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    V_WBCONC_Sheet = Source{[Item=""V_WBCONC"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(V_WBCONC_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers""," & _
        "{{""Object"", Int64.Type}, {""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", type text}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Changed Type"",{""Counter"", ""Class Type"", ""Internal char."", ""Value from""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Remov" & _
        "ed Columns"",{{""Char. Value"", ""V_WBCONC""}})" & Chr(10) & "in" & Chr(10) & "    #""Renamed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 "Query - V_WBCONC", _
        "Connection to the 'V_WBCONC' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=V_WBCONC" _
        , "SELECT * FROM [V_WBCONC]", 2

    ActiveWorkbook.Queries.Add Name:="V_DilutionFactor", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    V_DilutionFactor_Sheet = Source{[Item=""V_DilutionFactor"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(V_DilutionFactor_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""Object"", Int6" & _
        "4.Type}, {""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", type text}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Changed Type"",{""Counter"", ""Class Type"", ""Internal char."", ""Value from""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Removed Columns"",{{""C" & _
        "har. Value"", ""V_DilutionFactor""}})" & Chr(10) & "in" & Chr(10) & "    #""Renamed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - V_DilutionFactor", _
        "Connection to the 'V_DilutionFactor' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=V_DilutionFactor" _
        , "SELECT * FROM [V_DilutionFactor]", 2

    ActiveWorkbook.Queries.Add Name:="V_IPCONC", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    V_IPCONC_Sheet = Source{[Item=""V_IPCONC"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(V_IPCONC_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""Object"", Int64.Type}, {""Counter"", I" & _
        "nt64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", type number}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Changed Type"",{""Counter"", ""Class Type"", ""Internal char."", ""Value from""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Removed Columns"",{{""Char. Value"", ""V_IPCO" & _
        "NC""}})" & Chr(10) & "in" & Chr(10) & "    #""Renamed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 "Query - V_IPCONC", _
        "Connection to the 'V_IPCONC' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=V_IPCONC" _
        , "SELECT * FROM [V_IPCONC]", 2

    ActiveWorkbook.Queries.Add Name:="QC_Last_WB_Inspection_Date", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    QC_Last_WB_Inspection_Date_Sheet = Source{[Item=""QC_Last_WB_Inspection_Date"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(QC_Last_WB_Inspection_Date_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promote" & _
        "d Headers"",{{""Value from"", type text}, {""Object"", Int64.Type}})," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Changed Type"",{""Counter"", ""Class Type"", ""Internal char."", ""Char. Value""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Removed Columns"",{{""Value from"", ""QC_Last_WB_Inspection_Date""}})" & Chr(10) & "in" & Chr(10) & "    #""Renamed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - QC_Last_WB_Inspection_Date", _
        "Connection to the 'QC_Last_WB_Inspection_Date' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QC_Last_WB_Inspection_Date" _
        , "SELECT * FROM [QC_Last_WB_Inspection_Date]", 2

    ActiveWorkbook.Queries.Add Name:="QC_Last_IHC_Inspection_Date", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    QC_Last_IHC_Inspection_Date_Sheet = Source{[Item=""QC_Last_IHC_Inspection_Date"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(QC_Last_IHC_Inspection_Date_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Prom" & _
        "oted Headers"",{{""Object"", Int64.Type}, {""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", type text}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Changed Type"",{""Counter"", ""Class Type"", ""Internal char."", ""Char. Value""})," & Chr(10) & "    #""Renamed Columns"" = Table.Rename" & _
        "Columns(#""Removed Columns"",{{""Value from"", ""QC_Last_IHC_Inspection_Date""}})" & Chr(10) & "in" & Chr(10) & "    #""Renamed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - QC_Last_IHC_Inspection_Date", _
        "Connection to the 'QC_Last_IHC_Inspection_Date' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QC_Last_IHC_Inspection_Date" _
        , "SELECT * FROM [QC_Last_IHC_Inspection_Date]", 2

    ActiveWorkbook.Queries.Add Name:="QC_TestingBrother", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    QC_TestingBrother_Sheet = Source{[Item=""QC_TestingBrother"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(QC_TestingBrother_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""Object"", I" & _
        "nt64.Type}, {""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", type text}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Trimmed Text"" = Table.TransformColumns(#""Changed Type"",{{""Char. Value"", Text.Trim}})," & Chr(10) & "    #""Grouped Rows"" = Table.Group(#""Trimmed Text"", {""Object""}, {{""Count"", each _, type table}})," & Chr(10) & "" & _
        "    #""Added Custom"" = Table.AddColumn(#""Grouped Rows"", ""Custom"", each Table.ToList(" & Chr(10) & "         Table.Transpose(" & Chr(10) & "              Table.FromList(" & Chr(10) & "                   Table.Column([Count], ""Char. Value"")" & Chr(10) & "                  )" & Chr(10) & "                  )," & Chr(10) & "    Combiner.CombineTextByDelimiter("", "")" & Chr(10) & "    ))," & Chr(10) & "    #""Expanded Custom"" = Table.ExpandListColumn(#""Added Custom"", """ & _
        "Custom"")," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Expanded Custom"",{""Count""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Removed Columns"",{{""Custom"", ""QC_TestingBrother""}})" & Chr(10) & "in" & Chr(10) & "    #""Renamed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - QC_TestingBrother", _
        "Connection to the 'QC_TestingBrother' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QC_TestingBrother" _
        , "SELECT * FROM [QC_TestingBrother]", 2

    ActiveWorkbook.Queries.Add Name:="QC_QualifiedApplications", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    QC_QualifiedApplications_Sheet = Source{[Item=""QC_QualifiedApplications"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(QC_QualifiedApplications_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Head" & _
        "ers"",{{""Object"", Int64.Type}, {""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", type text}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Trimmed Text"" = Table.TransformColumns(#""Changed Type"",{{""Char. Value"", Text.Trim}})," & Chr(10) & "    #""Grouped Rows"" = Table.Group(#""Trimmed Text"", {""Object""}, {{""Count"", ea" & _
        "ch _, type table}})," & Chr(10) & "    #""Added Custom"" = Table.AddColumn(#""Grouped Rows"", ""Custom"", each Table.ToList(" & Chr(10) & "         Table.Transpose(" & Chr(10) & "              Table.FromList(" & Chr(10) & "                   Table.Column([Count], ""Char. Value"")" & Chr(10) & "                  )" & Chr(10) & "                  )," & Chr(10) & "    Combiner.CombineTextByDelimiter("", "")" & Chr(10) & "    ))," & Chr(10) & "    #""Expanded Custom"" = Table.ExpandListColumn(" & _
        "#""Added Custom"", ""Custom"")," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Expanded Custom"",{""Count""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Removed Columns"",{{""Custom"", ""QC_QualifiedApplications""}})" & Chr(10) & "in" & Chr(10) & "    #""Renamed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - QC_QualifiedApplications", _
        "Connection to the 'QC_QualifiedApplications' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QC_QualifiedApplications" _
        , "SELECT * FROM [QC_QualifiedApplications]", 2

    ActiveWorkbook.Queries.Add Name:="QC_TestResult", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    QC_TestResult_Sheet = Source{[Item=""QC_TestResult"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(QC_TestResult_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""Object"", Int64.Type}, " & _
        "{""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", type text}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Trimmed Text"" = Table.TransformColumns(#""Changed Type"",{{""Char. Value"", Text.Trim}})," & Chr(10) & "    #""Grouped Rows"" = Table.Group(#""Trimmed Text"", {""Object""}, {{""Count"", each _, type table}})," & Chr(10) & "    #""Added" & _
        " Custom"" = Table.AddColumn(#""Grouped Rows"", ""Custom"", each Table.ToList(" & Chr(10) & "         Table.Transpose(" & Chr(10) & "              Table.FromList(" & Chr(10) & "                   Table.Column([Count], ""Char. Value"")" & Chr(10) & "                  )" & Chr(10) & "                  )," & Chr(10) & "    Combiner.CombineTextByDelimiter("", "")" & Chr(10) & "    ))," & Chr(10) & "    #""Expanded Custom"" = Table.ExpandListColumn(#""Added Custom"", ""Custom"")," & Chr(10) & " " & _
        "   #""Removed Columns"" = Table.RemoveColumns(#""Expanded Custom"",{""Count""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Removed Columns"",{{""Custom"", ""QC_TestResult""}})" & Chr(10) & "in" & Chr(10) & "    #""Renamed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - QC_TestResult", _
        "Connection to the 'QC_TestResult' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QC_TestResult" _
        , "SELECT * FROM [QC_TestResult]", 2

    ActiveWorkbook.Queries.Add Name:="QC_Tested_Cell_Line", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    QC_Tested_Cell_Line_Sheet = Source{[Item=""QC_Tested_Cell_Line"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(QC_Tested_Cell_Line_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""Objec" & _
        "t"", Int64.Type}, {""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", type text}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Trimmed Text"" = Table.TransformColumns(#""Changed Type"",{{""Char. Value"", Text.Trim}})," & Chr(10) & "    #""Grouped Rows"" = Table.Group(#""Trimmed Text"", {""Object""}, {{""Count"", each _, type tabl" & _
        "e}})," & Chr(10) & "    #""Added Custom"" = Table.AddColumn(#""Grouped Rows"", ""Custom"", each Table.ToList(" & Chr(10) & "         Table.Transpose(" & Chr(10) & "              Table.FromList(" & Chr(10) & "                   Table.Column([Count], ""Char. Value"")" & Chr(10) & "                  )" & Chr(10) & "                  )," & Chr(10) & "    Combiner.CombineTextByDelimiter("", "")" & Chr(10) & "    ))," & Chr(10) & "    #""Expanded Custom"" = Table.ExpandListColumn(#""Added Custom" & _
        """, ""Custom"")," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Expanded Custom"",{""Count""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Removed Columns"",{{""Custom"", ""QC_Tested_Cell_Line""}})" & Chr(10) & "in" & Chr(10) & "    #""Renamed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - QC_Tested_Cell_Line", _
        "Connection to the 'QC_Tested_Cell_Line' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QC_Tested_Cell_Line" _
        , "SELECT * FROM [QC_Tested_Cell_Line]", 2

    ActiveWorkbook.Queries.Add Name:="QC_Cell_Treatment", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    QC_Cell_Treatment_Sheet = Source{[Item=""QC_Cell_Treatment"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(QC_Cell_Treatment_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""Object"", t" & _
        "ype text}, {""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", type text}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Trimmed Text"" = Table.TransformColumns(#""Changed Type"",{{""Char. Value"", Text.Trim}})," & Chr(10) & "    #""Grouped Rows"" = Table.Group(#""Trimmed Text"", {""Object""}, {{""Count"", each _, type table}})," & Chr(10) & " " & _
        "   #""Added Custom"" = Table.AddColumn(#""Grouped Rows"", ""Custom"", each Table.ToList(" & Chr(10) & "         Table.Transpose(" & Chr(10) & "              Table.FromList(" & Chr(10) & "                   Table.Column([Count], ""Char. Value"")" & Chr(10) & "                  )" & Chr(10) & "                  )," & Chr(10) & "    Combiner.CombineTextByDelimiter("", "")" & Chr(10) & "    ))," & Chr(10) & "    #""Changed Type1"" = Table.TransformColumnTypes(#""Added Custom"",{{" & _
        """Object"", Int64.Type}})," & Chr(10) & "    #""Expanded Custom"" = Table.ExpandListColumn(#""Changed Type1"", ""Custom"")," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Expanded Custom"",{""Count""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Removed Columns"",{{""Custom"", ""QC_Cell_Treatment""}})," & Chr(10) & "    #""Removed Errors"" = Table.RemoveRowsWithErrors(#""Renamed Colu" & _
        "mns"", {""Object""})" & Chr(10) & "in" & Chr(10) & "    #""Removed Errors"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - QC_Cell_Treatment", _
        "Connection to the 'QC_Cell_Treatment' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QC_Cell_Treatment" _
        , "SELECT * FROM [QC_Cell_Treatment]", 2

    ActiveWorkbook.Queries.Add Name:="QC_Qualified_Cell_Lines", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    QC_Qualified_Cell_Lines_Sheet = Source{[Item=""QC_Qualified_Cell_Lines"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(QC_Qualified_Cell_Lines_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers" & _
        """,{{""Object"", Int64.Type}, {""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", type text}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Trimmed Text"" = Table.TransformColumns(#""Changed Type"",{{""Char. Value"", Text.Trim}})," & Chr(10) & "    #""Grouped Rows"" = Table.Group(#""Trimmed Text"", {""Object""}, {{""Count"", each " & _
        "_, type table}})," & Chr(10) & "    #""Added Custom"" = Table.AddColumn(#""Grouped Rows"", ""Custom"", each Table.ToList(" & Chr(10) & "         Table.Transpose(" & Chr(10) & "              Table.FromList(" & Chr(10) & "                   Table.Column([Count], ""Char. Value"")" & Chr(10) & "                  )" & Chr(10) & "                  )," & Chr(10) & "    Combiner.CombineTextByDelimiter("", "")" & Chr(10) & "    ))," & Chr(10) & "    #""Expanded Custom"" = Table.ExpandListColumn(#""" & _
        "Added Custom"", ""Custom"")," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Expanded Custom"",{""Count""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Removed Columns"",{{""Custom"", ""QC_Qualified_Cell_Lines""}})" & Chr(10) & "in" & Chr(10) & "    #""Renamed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - QC_Qualified_Cell_Lines", _
        "Connection to the 'QC_Qualified_Cell_Lines' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QC_Qualified_Cell_Lines" _
        , "SELECT * FROM [QC_Qualified_Cell_Lines]", 2

    ActiveWorkbook.Queries.Add Name:="QC_TestedApplication", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    QC_TestedApplication_Sheet = Source{[Item=""QC_TestedApplication"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(QC_TestedApplication_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""Ob" & _
        "ject"", Int64.Type}, {""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", type text}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Trimmed Text"" = Table.TransformColumns(#""Changed Type"",{{""Char. Value"", Text.Trim}})," & Chr(10) & "    #""Grouped Rows"" = Table.Group(#""Trimmed Text"", {""Object""}, {{""Count"", each _, type t" & _
        "able}})," & Chr(10) & "    #""Added Custom"" = Table.AddColumn(#""Grouped Rows"", ""Custom"", each Table.ToList(" & Chr(10) & "         Table.Transpose(" & Chr(10) & "              Table.FromList(" & Chr(10) & "                   Table.Column([Count], ""Char. Value"")" & Chr(10) & "                  )" & Chr(10) & "                  )," & Chr(10) & "    Combiner.CombineTextByDelimiter("", "")" & Chr(10) & "    ))," & Chr(10) & "    #""Expanded Custom"" = Table.ExpandListColumn(#""Added Cus" & _
        "tom"", ""Custom"")," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Expanded Custom"",{{""Custom"", ""QC_TestedApplication""}})," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Renamed Columns"",{""Count""})" & Chr(10) & "in" & Chr(10) & "    #""Removed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - QC_TestedApplication", _
        "Connection to the 'QC_TestedApplication' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QC_TestedApplication" _
        , "SELECT * FROM [QC_TestedApplication]", 2

    ActiveWorkbook.Queries.Add Name:="QC_Datasheet_Low_Dilution", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    QC_Datasheet_Low_Dilution_Sheet = Source{[Item=""QC_Datasheet_Low_Dilution"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(QC_Datasheet_Low_Dilution_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted H" & _
        "eaders"",{{""Object"", Int64.Type}, {""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", Int64.Type}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Changed Type"",{""Counter"", ""Class Type"", ""Internal char."", ""Value from""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumn" & _
        "s(#""Removed Columns"",{{""Char. Value"", ""QC_Datasheet_Low_Dilution""}})" & Chr(10) & "in" & Chr(10) & "    #""Renamed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - QC_Datasheet_Low_Dilution", _
        "Connection to the 'QC_Datasheet_Low_Dilution' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QC_Datasheet_Low_Dilution" _
        , "SELECT * FROM [QC_Datasheet_Low_Dilution]", 2

    ActiveWorkbook.Queries.Add Name:="QC_Datasheet_High_Dilution", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    QC_Datasheet_High_Dilution_Sheet = Source{[Item=""QC_Datasheet_High_Dilution"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(QC_Datasheet_High_Dilution_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promote" & _
        "d Headers"",{{""Object"", Int64.Type}, {""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", Int64.Type}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Changed Type"",{""Counter"", ""Class Type"", ""Internal char."", ""Value from""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameCol" & _
        "umns(#""Removed Columns"",{{""Char. Value"", ""QC_Datasheet_High_Dilution""}})" & Chr(10) & "in" & Chr(10) & "    #""Renamed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - QC_Datasheet_High_Dilution", _
        "Connection to the 'QC_Datasheet_High_Dilution' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QC_Datasheet_High_Dilution" _
        , "SELECT * FROM [QC_Datasheet_High_Dilution]", 2

    ActiveWorkbook.Queries.Add Name:="QC_Molecular_Weight", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    QC_Molecular_Weight_Sheet = Source{[Item=""QC_Molecular_Weight"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(QC_Molecular_Weight_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""Objec" & _
        "t"", Int64.Type}, {""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", type text}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Trimmed Text"" = Table.TransformColumns(#""Changed Type"",{{""Char. Value"", Text.Trim}})," & Chr(10) & "    #""Grouped Rows"" = Table.Group(#""Trimmed Text"", {""Object""}, {{""Count"", each _, type tabl" & _
        "e}})," & Chr(10) & "    #""Added Custom"" = Table.AddColumn(#""Grouped Rows"", ""Custom"", each Table.ToList(" & Chr(10) & "         Table.Transpose(" & Chr(10) & "              Table.FromList(" & Chr(10) & "                   Table.Column([Count], ""Char. Value"")" & Chr(10) & "                  )" & Chr(10) & "                  )," & Chr(10) & "    Combiner.CombineTextByDelimiter("", "")" & Chr(10) & "    ))," & Chr(10) & "    #""Expanded Custom"" = Table.ExpandListColumn(#""Added Custom" & _
        """, ""Custom"")," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Expanded Custom"",{""Count""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Removed Columns"",{{""Custom"", ""QC_Molecular_Weight""}})" & Chr(10) & "in" & Chr(10) & "    #""Renamed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - QC_Molecular_Weight", _
        "Connection to the 'QC_Molecular_Weight' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QC_Molecular_Weight" _
        , "SELECT * FROM [QC_Molecular_Weight]", 2

    ActiveWorkbook.Queries.Add Name:="QC_Gel_Type", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    QC_Gel_Type_Sheet = Source{[Item=""QC_Gel_Type"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(QC_Gel_Type_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""Objec" & _
        "t"", Int64.Type}, {""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", type text}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Trimmed Text"" = Table.TransformColumns(#""Changed Type"",{{""Char. Value"", Text.Trim}})," & Chr(10) & "    #""Grouped Rows"" = Table.Group(#""Trimmed Text"", {""Object""}, {{""Count"", each _, type tabl" & _
        "e}})," & Chr(10) & "    #""Added Custom"" = Table.AddColumn(#""Grouped Rows"", ""Custom"", each Table.ToList(" & Chr(10) & "         Table.Transpose(" & Chr(10) & "              Table.FromList(" & Chr(10) & "                   Table.Column([Count], ""Char. Value"")" & Chr(10) & "                  )" & Chr(10) & "                  )," & Chr(10) & "    Combiner.CombineTextByDelimiter("", "")" & Chr(10) & "    ))," & Chr(10) & "    #""Expanded Custom"" = Table.ExpandListColumn(#""Added Custom" & _
        """, ""Custom"")," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Expanded Custom"",{""Count""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Removed Columns"",{{""Custom"", ""QC_Gel_Type""}})" & Chr(10) & "in" & Chr(10) & "    #""Renamed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - QC_Gel_Type", _
        "Connection to the 'QC_Gel_Type' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QC_Gel_Type" _
        , "SELECT * FROM [QC_Gel_Type]", 2

    ActiveWorkbook.Queries.Add Name:="Uniprot_Accession_Number", Formula:= _
        "let" & Chr(10) & "    Source = Excel.Workbook(File.Contents(""\\BETHYL-FS1\BethylShared\SAPData\ZCharValues.xlsm""), null, true)," & Chr(10) & "    Uniprot_Accession_Number_Sheet = Source{[Item=""Uniprot_Accession_Number"",Kind=""Sheet""]}[Data]," & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(Uniprot_Accession_Number_Sheet)," & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""Objec" & _
        "t"", Int64.Type}, {""Counter"", Int64.Type}, {""Class Type"", Int64.Type}, {""Internal char."", Int64.Type}, {""Char. Value"", type text}, {""Value from"", Int64.Type}})," & Chr(10) & "    #""Trimmed Text"" = Table.TransformColumns(#""Changed Type"",{{""Char. Value"", Text.Trim}})," & Chr(10) & "    #""Grouped Rows"" = Table.Group(#""Trimmed Text"", {""Object""}, {{""Count"", each _, type tabl" & _
        "e}})," & Chr(10) & "    #""Added Custom"" = Table.AddColumn(#""Grouped Rows"", ""Custom"", each Table.ToList(" & Chr(10) & "         Table.Transpose(" & Chr(10) & "              Table.FromList(" & Chr(10) & "                   Table.Column([Count], ""Char. Value"")" & Chr(10) & "                  )" & Chr(10) & "                  )," & Chr(10) & "    Combiner.CombineTextByDelimiter("", "")" & Chr(10) & "    ))," & Chr(10) & "    #""Expanded Custom"" = Table.ExpandListColumn(#""Added Custom" & _
        """, ""Custom"")," & Chr(10) & "    #""Removed Columns"" = Table.RemoveColumns(#""Expanded Custom"",{""Count""})," & Chr(10) & "    #""Renamed Columns"" = Table.RenameColumns(#""Removed Columns"",{{""Custom"", ""Uniprot_Accession_Number""}})" & Chr(10) & "in" & Chr(10) & "    #""Renamed Columns"""
    Workbooks("AllMatBatchCharValues.xlsm").Connections.Add2 _
        "Query - Uniprot_Accession_Number", _
        "Connection to the 'Uniprot_Accession_Number' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Uniprot_Accession_Number" _
        , "SELECT * FROM [Uniprot_Accession_Number]", 2
End Sub
Sub MakeCharValueTable()
    ActiveWorkbook.Queries("KSSK_INOB").Delete
    Sheets("AllMatBatchCharValues").Cells.Clear
    ActiveWorkbook.Queries.Add Name:="KSSK_INOB", Formula:= _
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
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=KSSK_INOB" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [KSSK_INOB]")
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
        .ListObject.DisplayName = "KSSK_INOB"
        .Refresh BackgroundQuery:=False
    End With
    'Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
End Sub
