Attribute VB_Name = "ZPL_Macros"
Option Explicit
Sub PrintLabel()
    Dim AbData As String
    Dim DilData As String, NumericalA300 As String
    Dim AbVol As Double, ReformulationShortText As String, ReformulationCount As Integer
    If IsNumeric(Range("AM" & ActiveCell.Row)) = True And Range("AM" & ActiveCell.Row) <> "" Then 'if the order has not been confirmed yet, this will execute
        If InStr(Range("AC" & ActiveCell.Row), "A30") > 0 Then
            NumericalA300 = CDbl(Trim(Replace(Replace(Replace(Replace(Range("AC" & ActiveCell.Row), "A30", ""), "A-T", ""), "A-M", ""), "-", "")))
        Else
            NumericalA300 = CDbl("0")
        End If
        If Right(Range("AC" & ActiveCell.Row), 1) = "T" Then 'if it's a T
            'set the Ab volume
            If Left(Range("AG" & ActiveCell.Row), 4) = "A500" Then
                If WorksheetFunction.VLookup(Range("AG" & ActiveCell.Row), Sheets("LabelData").Range("A:F"), 6, False) = "100 µl (10 blots)" Then
                    AbVol = 22
                    Else 'then it's the IHC version
                    AbVol = 11
                End If
            ElseIf Right(Range("AG" & ActiveCell.Row), 1) = "A" Then
                AbVol = 11
                NumericalA300 = CDbl(Trim(Replace(Replace(Replace(Replace(Range("AC" & ActiveCell.Row), "A30", ""), "A-T", ""), "A-M", ""), "-", "")))
            ElseIf Left(Range("AG" & ActiveCell.Row), 3) = "IHC" Then
                AbVol = 11
            ElseIf Left(Range("AG" & ActiveCell.Row), 3) = "A70" Then
                'mbb 07272016...multiply standard 22µl A700 trial volume by quotient of WBconc / ABconc...e.g. if the WB conc is 2µg/ml and the Ab stock is 1µg/ml (i.e. A700-002), then 44µl will be shipped
                Dim BOMQuantity As String
                BOMQuantity = WorksheetFunction.VLookup(Range("AC" & ActiveCell.Row), Sheets("BomChild").Range("A:E"), 5, False)
                AbVol = CInt(BOMQuantity) * 0.11
                'AbVol = 22 * (WorksheetFunction.VLookup(Range("AG" & ActiveCell.Row), Sheets("Zcharvalues").Range("A:F"), 4, False) / WorksheetFunction.VLookup(Range("AG" & ActiveCell.Row), Sheets("Zcharvalues").Range("A:F"), 6, False))
                'NumericalA300 = CDbl(Trim(Replace(Replace(Replace(Replace(Range("AC" & ActiveCell.Row), "A30", ""), "A-T", ""), "A-M", ""), "-", ""))) ' not necessary
            Else
                AbVol = 22
                NumericalA300 = CDbl(Trim(Replace(Replace(Replace(Replace(Range("AC" & ActiveCell.Row), "A30", ""), "A-T", ""), "A-M", ""), "-", "")))
            End If
            If InStr(Range("AC" & ActiveCell.Row).Value, "301-985") = 0 And NumericalA300 < 5576 And NumericalA300 <> 4842 And NumericalA300 <> 4923 Then
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                End If
            ElseIf InStr(Range("AC" & ActiveCell.Row).Value, "301-985") > 0 Then
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                End If
            ElseIf NumericalA300 > 5575 Then
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "-M" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "-M" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                End If
            ElseIf NumericalA300 <> 4842 Or NumericalA300 <> 4923 Then
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "-M" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "-M" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                End If
            End If
            Else 'it's an M
            If InStr(Range("AC" & ActiveCell.Row).Value, "301-985") = 0 And NumericalA300 < 5576 Then
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), "", Range("AI" & ActiveCell.Row))
                    End If
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                        Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", Range("AI" & ActiveCell.Row))
                        Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                    End If
                End If
            ElseIf InStr(Range("AC" & ActiveCell.Row).Value, "301-985") > 0 Then
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), "", Range("AI" & ActiveCell.Row))
                    End If
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                        Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", Range("AI" & ActiveCell.Row))
                        Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                    End If
                End If
            ElseIf NumericalA300 > 5575 Then
                If WorksheetFunction.VLookup(Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), "", Range("AI" & ActiveCell.Row))
                    End If
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                        Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", Range("AI" & ActiveCell.Row))
                        Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                    End If
                End If
            End If
        End If
    ElseIf InStr(Range("AM" & ActiveCell.Row), "Confirmed") > 0 And Range("AM" & ActiveCell.Row) <> "" Then 'if the order has been confirmed already, this will execute
        If Right(Range("AC" & ActiveCell.Row), 1) = "T" Then
            If Left(Range("AG" & ActiveCell.Row), 4) = "A500" Then
                If WorksheetFunction.VLookup(Range("AG" & ActiveCell.Row), Sheets("LabelData").Range("A:F"), 6, False) = "100 µl (10 blots)" Then
                    AbVol = 22
                    Else 'then it's the IHC version
                    AbVol = 11
                End If
            ElseIf Right(Range("AG" & ActiveCell.Row), 1) = "A" Then 'classic trial
                AbVol = 11
                NumericalA300 = CDbl(Trim(Replace(Replace(Replace(Replace(Range("AC" & ActiveCell.Row), "A30", ""), "A-T", ""), "A-M", ""), "-", "")))
            ElseIf Left(Range("AG" & ActiveCell.Row), 3) = "IHC" Then 'IHC trial
                AbVol = 11
                Else 'trial pulled from M
                Range("AP" & ActiveCell.Row) = Left(Range("AQ" & ActiveCell.Row), InStr(Range("AQ" & ActiveCell.Row), " "))
                AbVol = 22
                NumericalA300 = CDbl(Trim(Replace(Replace(Replace(Replace(Range("AC" & ActiveCell.Row), "A30", ""), "A-T", ""), "A-M", ""), "-", "")))
            End If

            If InStr(Range("AC" & ActiveCell.Row).Value, "301-985") = 0 And NumericalA300 < 5576 Then
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                End If
            ElseIf InStr(Range("AC" & ActiveCell.Row).Value, "301-985") > 0 Then
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                End If
            ElseIf NumericalA300 > 5575 Then
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "-M" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "-M" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                End If
            End If
            Else 'it's an M
            Range("AP" & ActiveCell.Row) = Left(Range("AQ" & ActiveCell.Row), InStr(Range("AQ" & ActiveCell.Row), " "))
            If InStr(Range("AC" & ActiveCell.Row).Value, "301-985") = 0 And NumericalA300 < 5576 Then
                If WorksheetFunction.VLookup(Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), "", Range("AI" & ActiveCell.Row))
                    End If
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                        Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), "", Range("AI" & ActiveCell.Row))
                        Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                    End If
                End If
            ElseIf InStr(Range("AC" & ActiveCell.Row).Value, "301-985") > 0 Then
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), "", Range("AI" & ActiveCell.Row))
                    End If
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                        Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), "", Range("AI" & ActiveCell.Row))
                        Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                    End If
                End If
            ElseIf NumericalA300 > 5575 Then
                If WorksheetFunction.VLookup(Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), "", Range("AI" & ActiveCell.Row))
                    End If
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                        Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), "", Range("AI" & ActiveCell.Row))
                        Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                    End If
                End If
            End If
        End If
    Else
        Stop 'what is the purpose of this section?
        If Range("AQ" & ActiveCell.Row) <> "" Then
            Range("AP" & ActiveCell.Row) = Left(Range("AQ" & ActiveCell.Row), InStr(Range("AQ" & ActiveCell.Row), " "))
            If Range("AP" & ActiveCell.Row) >= Range("AO" & ActiveCell.Row) Then
                If WorksheetFunction.VLookup(Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), "", Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), "", Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Range("AS" & ActiveCell.Row) = Range("AH" & ActiveCell.Row) & "0" & ReformulationCount
                End If
            Else
                AppActivate Application.Caption
                MsgBox ("You must have a production order and/or confirm the 10-Blot inventory before you can print a label.")
                Range("AM" & ActiveCell.Row).Select
            End If
        End If
    End If
    If Right(Range("AC" & ActiveCell.Row), 2) = "-M" And Range("AP" & ActiveCell.Row) = 0 Then
        'figure out how to handle Mstock
        Call PrintComboLabel_MStock
    ElseIf Right(Range("AC" & ActiveCell.Row), 2) = "-M" And Range("AP" & ActiveCell.Row) < 1.2 Then
        Call PrintComboLabel_MStock 'may need to make another routine to bump up the batch number and adjust the dilution factor
    End If
    If ReformulationShortText <> "" Then MsgBox Range("AC" & ActiveCell.Row) & " is a reformulation. Notify IT development to address vial label batch printing issue."
End Sub
Sub PrintLabel_previous()
    Dim AbData As String
    Dim DilData As String
    Dim AbVol As Double, ReformulationShortText As String, ReformulationCount As Integer
    If IsNumeric(Range("AM" & ActiveCell.Row)) = True And Range("AM" & ActiveCell.Row) <> "" Then 'if the order has not been confirmed yet, this will execute
        If Right(Range("AC" & ActiveCell.Row), 1) = "T" Then
            If Left(Range("AG" & ActiveCell.Row), 4) = "A500" Then
                If WorksheetFunction.VLookup(Range("AG" & ActiveCell.Row), Sheets("LabelData").Range("A:F"), 6, False) = "100 µl (10 blots)" Then
                    AbVol = 22
                    Else 'then it's the IHC version
                    AbVol = 11
                End If
            ElseIf Right(Range("AG" & ActiveCell.Row), 1) = "A" Then
                AbVol = 11
            ElseIf Left(Range("AG" & ActiveCell.Row), 3) = "IHC" Then
                AbVol = 11
            ElseIf Left(Range("AG" & ActiveCell.Row), 3) = "A70" Then
                'mbb 07272016...multiply standard 22µl A700 trial volume by quotient of WBconc / ABconc...e.g. if the WB conc is 2µg/ml and the Ab stock is 1µg/ml (i.e. A700-002), then 44µl will be shipped
                Dim BOMQuantity As String
                BOMQuantity = WorksheetFunction.VLookup(Range("AG" & ActiveCell.Row), Sheets("BomChild").Range("A:E"), 5, False)
                AbVol = CInt(BOMQuantity) * 0.11
                'AbVol = 22 * (WorksheetFunction.VLookup(Range("AG" & ActiveCell.Row), Sheets("Zcharvalues").Range("A:F"), 4, False) / WorksheetFunction.VLookup(Range("AG" & ActiveCell.Row), Sheets("Zcharvalues").Range("A:F"), 6, False))
            Else
                AbVol = 22
            End If
            If InStr(Range("AC" & ActiveCell.Row).Value, "301-985") = 0 Then
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                End If
            Else
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                End If
            End If
        Else
            If InStr(Range("AC" & ActiveCell.Row).Value, "301-985") = 0 Then
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), "", Range("AI" & ActiveCell.Row))
                    End If
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", Range("AI" & ActiveCell.Row))
                    End If
                End If
            Else
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), "", Range("AI" & ActiveCell.Row))
                    End If
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", Range("AI" & ActiveCell.Row))
                    End If
                End If
            End If
        End If
    ElseIf InStr(Range("AM" & ActiveCell.Row), "Confirmed") > 0 And Range("AM" & ActiveCell.Row) <> "" Then 'if the order has been confirmed already, this will execute
        If Right(Range("AC" & ActiveCell.Row), 1) = "T" Then
            If Left(Range("AG" & ActiveCell.Row), 4) = "A500" Then
                If WorksheetFunction.VLookup(Range("AG" & ActiveCell.Row), Sheets("LabelData").Range("A:F"), 6, False) = "100 µl (10 blots)" Then
                    AbVol = 22
                    Else 'then it's the IHC version
                    AbVol = 11
                End If
            ElseIf Right(Range("AG" & ActiveCell.Row), 1) = "A" Then 'classic trial
                AbVol = 11
            ElseIf Left(Range("AG" & ActiveCell.Row), 3) = "IHC" Then 'IHC trial
                AbVol = 11
                Else 'trial pulled from M
                Range("AP" & ActiveCell.Row) = Left(Range("AQ" & ActiveCell.Row), InStr(Range("AQ" & ActiveCell.Row), " "))
                AbVol = 22
            End If
            If InStr(Range("AC" & ActiveCell.Row).Value, "301-985") = 0 Then
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                End If
            Else

                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-T", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), AbVol, Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                End If
            End If
        Else
            Range("AP" & ActiveCell.Row) = Left(Range("AQ" & ActiveCell.Row), InStr(Range("AQ" & ActiveCell.Row), " "))


            If InStr(Range("AC" & ActiveCell.Row).Value, "301-985") = 0 Then
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), "", Range("AI" & ActiveCell.Row))
                    End If
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), "", Range("AI" & ActiveCell.Row))
                    End If
                End If
            Else
                If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), "", Range("AI" & ActiveCell.Row))
                    End If
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    'stop
                    If Range("AB" & ActiveCell.Row) = "" Then
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                    Else
                        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), Range("AJ" & ActiveCell.Row), Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, Range("AK" & ActiveCell.Row), "A400 Diluent", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), "", Range("AI" & ActiveCell.Row))
                    End If
                End If
            End If
        End If
    Else
        Stop
        If Range("AQ" & ActiveCell.Row) <> "" Then
            Range("AP" & ActiveCell.Row) = Left(Range("AQ" & ActiveCell.Row), InStr(Range("AQ" & ActiveCell.Row), " "))
            If Range("AP" & ActiveCell.Row) >= Range("AO" & ActiveCell.Row) Then
                If WorksheetFunction.VLookup(Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), "", Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row), "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                Else
                    ReformulationShortText = WorksheetFunction.VLookup(Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
                    ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
                    Stop
                    Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row), Range("AD" & ActiveCell.Row), "", Range("AG" & ActiveCell.Row), Range("AH" & ActiveCell.Row) & "0" & ReformulationCount, "", "", Range("AM" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row), Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
                End If
            Else
                AppActivate Application.Caption
                MsgBox ("You must have a production order and/or confirm the 10-Blot inventory before you can print a label.")
                Range("AM" & ActiveCell.Row).Select
            End If
        End If
    End If
    If Right(Range("AC" & ActiveCell.Row), 2) = "-M" And Range("AP" & ActiveCell.Row) = 0 Then
        'figure out how to handle Mstock
        Call PrintComboLabel_MStock
    ElseIf Right(Range("AC" & ActiveCell.Row), 2) = "-M" And Range("AP" & ActiveCell.Row) < 1.2 Then
        Call PrintComboLabel_MStock 'may need to make another routine to bump up the batch number and adjust the dilution factor
    End If
End Sub
Sub PrintSOLabel()
    Dim AbData As String
    Dim DilData As String
    Dim AbVol As Double
    If Range("AP" & ActiveCell.Row) >= Range("AO" & ActiveCell.Row) Then
        Call LandscapeTwoLabelPrint("\\C2186\Vialing1_GX430t", 1, Range("AC" & ActiveCell.Row) & "_" & GetMatBatch, Range("AD" & ActiveCell.Row), "", "", "", "", "", "SO:" & Range("AN" & ActiveCell.Row), Range("AC" & ActiveCell.Row) & "_" & GetMatBatch, Range("AO" & ActiveCell.Row), Range("AI" & ActiveCell.Row))
    Else
        AppActivate Application.Caption
        MsgBox ("You must have a production order and/or confirm the 10-Blot inventory before you can print a label.")
        Range("AM" & ActiveCell.Row).Select
    End If
End Sub
Sub LandscapeTwoLabelPrint(MyPrinter, NumLabels, MatNum, MatDesc, AbVol, CompName, CompBatch, DilVol, DilName, Datamatrix, Datamatrix2, Eaches, Sloc)
    Open MyPrinter For Output As #1
    Print #1, "^XA" & vbCrLf 'Start Label
    Print #1, "^LT25" & vbCrLf 'Label Top
    Print #1, "~SD15" & vbCrLf 'Media Darkness
    Print #1, "^PQ" & NumLabels & vbCrLf 'Num of Labels
    Print #1, "^PW760" & vbCrLf 'Label Print Width in Pixel
    Print #1, "^FO180,-5" & vbCrLf 'DataMatrix Location
    Print #1, "^BXN,3,200" & vbCrLf 'DataMatrix Density
    Print #1, "^FD" & Datamatrix & "^FS" & vbCrLf 'DataMatrix Data
    Print #1, "^FO620,-5" & vbCrLf 'DataMatrix Location
    Print #1, "^BXN,3,200" & vbCrLf 'DataMatrix Density
    Print #1, "^FD" & Datamatrix2 & "^FS" & vbCrLf 'DataMatrix Data
    If CompBatch <> "" Then
        Print #1, "^FO70,60^A@N,70,70,E:ARI000.FNT^FD" & MatNum & "_" & CompBatch & "^FS" & vbCrLf  'Line 1
    Else
        Print #1, "^FO70,60^A@N,70,70,E:ARI000.FNT^FD" & Datamatrix2 & "^FS" & vbCrLf  'Line 1
    End If
    Print #1, "^FB570,2,,," & vbCrLf 'Word Wrap
    Print #1, "^FO70,140^A@N,40,40,E:ARI000.FNT^FD" & MatDesc & "^FS" & vbCrLf  'Line 2

    If Sloc <> "" Then Print #1, "^FO340,15^A@N,40,40,E:ARI000.FNT^FD" & Sloc & "^FS" & vbCrLf 'Eaches

    If Eaches <> "" Then Print #1, "^FO620,220^A@N,40,40,E:ARI000.FNT^FD" & Eaches & "ea^FS" & vbCrLf 'Eaches
    'If AbVol <> "" Then Print #1, "^FO60,250^A@N,50,50,E:ARI000.FNT^FH^FD" & AbVol & " _E6l of " & CompName & "_5F" & CompBatch & "^FS" & vbCrLf  'Line 3
    If AbVol <> "" Then
        If AbVol = 11 Then
            Print #1, "^FO80,250^A@N,50,50,E:ARI000.FNT^FH^FD" & AbVol & " _E6l of " & CompName & "_5F" & CompBatch & "^FS" & vbCrLf  'Line 3
        Else
            'Print #1, "^FO55,220   ^GB505,45,85,,8^FS" 'print black box
            Print #1, "^FO55,220   ^GB535,45,85,,8^FS" 'print black box
            Print #1, "^LRY^FO80,240^A@N,50,50,E:ARI000.FNT^FH^FD" & AbVol & " _E6l of " & CompName & "_5F" & CompBatch & "^FS" & vbCrLf  'Line 3
        End If
    End If
    If DilVol <> "" Then Print #1, "^FO80,310^A@N,50,50,E:ARI000.FNT^FH^FD" & DilVol & " _E6l of " & DilName & "^FS" & vbCrLf 'Line 4
    Print #1, "^XZ" & vbCrLf 'End Label
    Close #1
End Sub
Sub PrintConfigLabel()
    Dim prtDevice As String
    Dim strQuote As String
    strQuote = Chr(34)
    prtDevice = "\\C2186\Vialing1_GX430t"
    ' open printer port
    Open prtDevice For Output As #1
    ' initialize printer
    Print #1, "^XA" & vbCrLf 'Start Label
    Print #1, "^WDZ:*.*" & vbCrLf 'Start Label
    Print #1, "^XZ" & vbCrLf 'End Label
    ' close printer port
    Close #1
End Sub
Sub ResetPrinterMemory()
    Dim prtDevice As String
    Dim strQuote As String
    strQuote = Chr(34)
    prtDevice = "\\C2186\Vialing1_GX430t"
    ' open printer port
    Open prtDevice For Output As #1
    ' initialize printer
    Print #1, "^XA" & vbCrLf 'Start Label
    Print #1, "^JB" & vbCrLf 'Start Label
    Print #1, "^XZ" & vbCrLf 'End Label
    ' close printer port
    Close #1
End Sub
Sub PrintAllVialing2()
    Dim MyPrinter, NumLabels, MatNum, OriginalMatDesc As String, MatDesc As Range, PrintMatDesc As String, AbVol, VialAmt, Count As Integer
    Dim CompName, CompBatch, DilVol, DilName, Datamatrix, Datamatrix2, Eaches, Sloc
    MyPrinter = "\\C2186\Vialing#2_S4M"
    Open MyPrinter For Output As #1
    Range("AC2").Select
    While Range("AC" & ActiveCell.Row) <> ""
        If Range("AB" & ActiveCell.Row) <> "Dilute 1st" Then
            NumLabels = Range("AO" & ActiveCell.Row) + 1
            'NumLabels = 1
            MatNum = Range("AC" & ActiveCell.Row)
            VialAmt = WorksheetFunction.VLookup(MatNum, Sheets("LabelData").Range("A:K"), 6, False)
            VialAmt = Replace(VialAmt, "µ", "_E6")
            PrintMatDesc = WorksheetFunction.VLookup(MatNum, Sheets("LabelData").Range("A:K"), 3, False)
            CompBatch = Range("AH" & ActiveCell.Row)
            Datamatrix = MatNum & "_" & CompBatch
            If Range("AS" & ActiveCell.Row).Value <> "" Then CompBatch = Range("AS" & ActiveCell.Row)
            Call VialLabelPrint(MyPrinter, NumLabels, MatNum, PrintMatDesc, CompBatch, Datamatrix, VialAmt)
        End If
        ActiveCell.Offset(1, 0).Select
        DoEvents
    Wend
    Close #1
End Sub
Sub PrintIndivVialing2()
    Dim MyPrinter, NumLabels, MatNum, OriginalMatDesc As String, MatDesc As Range, PrintMatDesc As String, AbVol, VialAmt, Count As Integer
    Dim CompName, CompBatch, DilVol, DilName, Datamatrix, Datamatrix2, Eaches, Sloc
    If ActiveCell.column <> 29 Or ActiveCell.Value = "" Then
        MsgBox "Please select a material number."
        Exit Sub
    End If
    MyPrinter = "\\C2186\Vialing#2_S4M"
    Open MyPrinter For Output As #1
    NumLabels = Range("AO" & ActiveCell.Row) + 1
    CompName = Range("AG" & ActiveCell.Row)
    'NumLabels = 1
    MatNum = Range("AC" & ActiveCell.Row)
    VialAmt = WorksheetFunction.VLookup(MatNum, Sheets("LabelData").Range("A:K"), 6, False)
    If InStr(VialAmt, "blot") > 0 And InStr(CompName, "-M") = 0 And InStr(MatNum, "-T") > 0 Then Range("AA" & ActiveCell.Row) = "Classic trial with 20 µl (2 blots) on label. Check label database."
    If InStr(VialAmt, "blot") = 0 And InStr(CompName, "-M") > 0 And InStr(MatNum, "-T") > 0 Then Range("AA" & ActiveCell.Row) = "Modified trial with 10 µl (X mg/ml) on label. Check label database."
    If InStr(VialAmt, "blot") > 0 And InStr(CompName, "-M") = 0 And InStr(MatNum, "-M") > 0 Then Range("AA" & ActiveCell.Row) = "Modified with 100 µl (X mg/ml) on label. Check label database."
    If InStr(VialAmt, "blot") = 0 And InStr(CompName, "-M") > 0 And InStr(MatNum, "-M") > 0 Then Range("AA" & ActiveCell.Row) = "Classic with 100 µl (10 blots) on label. Check label database."
    VialAmt = Replace(VialAmt, "µ", "_E6")
    PrintMatDesc = WorksheetFunction.VLookup(MatNum, Sheets("LabelData").Range("A:K"), 3, False)
    CompBatch = Range("AH" & ActiveCell.Row)
    Datamatrix = MatNum & "_" & CompBatch
    If Range("AS" & ActiveCell.Row).Value <> "" Then CompBatch = Range("AS" & ActiveCell.Row)
    Call VialLabelPrint(MyPrinter, NumLabels, MatNum, PrintMatDesc, CompBatch, Datamatrix, VialAmt)
    Close #1
    If WorksheetFunction.CountA(Range("AA:AA")) > 1 Then MsgBox "One or more labels printed with incorrect amounts. See notes at left."
End Sub
Sub VialLabelPrint_old(MyPrinter, NumLabels, MatNum, MatDesc, CompBatch, Datamatrix, VialAmt)
    Dim APLine As Integer, PrinterOpen As Boolean
    'Open MyPrinter For Output As #1
    Print #1, "^XA"                                     'Start Label   don't change
    Print #1, "^LS-50"                                   'Label shift
    Print #1, "^LT-10"                                   'Label Top                  might need to tweak for 1.5 x 0.75 label
    Print #1, "^MD7"                                   'Media Darkness             don't change yet
    Print #1, "^PQ" & NumLabels                         'Num of Labels              want two labels (one for sheet, one for vial)
    Print #1, "^PW450"                                  'Label Print Width in Pixel (1.5 inch = 450 dot width)
    '/////////////////// 1.5 x 0.75 = 450 x 225 dots (X =  0 to 450, Y = 0 to 225)
    Print #1, "^FO360,148"                              'DataMatrix Location
    Print #1, "^BXN,3,200"                            'DataMatrix Density
    Print #1, "^FD" & Datamatrix & "^FS"                        'DataMatrix Data (this is the yyyymmdd date)
    Print #1, "^FO90,70   ^GB42,170,21^FS" 'print black box
    If Len(MatNum) < 11 Then
        Print #1, "^LRY ^FO100,64  ^A@B,28,28,E:ARIALNB.FNT    ^FD" & MatNum & "    ^FS" 'print matnum bottom to top in black box (standard matnum length)
    Else
        Print #1, "^LRY ^FO100,64  ^A@B,26,26,E:ARIALNB.FNT    ^FD" & MatNum & "    ^FS" 'print matnum bottom to top in black box (long matnum length...may need to play with this)
    End If
    Print #1, "     ^FO100,20  ^A@B,28,28,E:ARIALN.FNT    ^FD" & CompBatch & "    ^FS" 'print batch above black box
    Print #1, "^FB320,2     ^FO140,40  ^A@N,24.5,24,E:ARIALN.FNT    ^FD" & MatDesc & "    ^FS" 'print matdesc to the right of batch...wrapped to second line if desc string is long
    APLine = 69
    If Len(MatDesc) > 25 Then APLine = 96  'if description is wrapped to second line (if length is > than width of FB), then push Affinity purified down to third line (65 to 90)
    Print #1, "     ^FO140, " & APLine & "  ^A@N,24.5,24,E:ARIALN.FNT       ^FD" & "Affinity Purified" & "^FS"
    Print #1, "^FO140,150^A@N,24.5,24,E:ARIALN.FNT^FH^FD" & VialAmt & "^FS"
    Print #1, "      ^FO140,185  ^A@N,24.5,24.5,E:ARIALN.FNT    ^FH^FD" & "Store at 2 - 8_F8 C" & "    ^FS"
    Print #1, "      ^FO140,217  ^A@N,17.7,15.3,E:ARIALNB.FNT    ^FD" & "For in vitro research use only" & "    ^FS"
    Print #1, "^FO345,200^IME:BL_BW_H30.GRF^FS" 'Bethyl Logo
    Print #1, "^XZ" & vbCrLf 'End Label
    'Close #1
End Sub
Sub VialLabelPrint_v2(MyPrinter, NumLabels, MatNum, MatDesc, CompBatch, Datamatrix, VialAmt)
    Dim APLine As Integer, PrinterOpen As Boolean
    'Open MyPrinter For Output As #1
    Print #1, "^XA"                                     'Start Label   don't change
    Print #1, "^LT0"                                   'Label Top                  might need to tweak for 1.5 x 0.75 label
    Print #1, "~SD15"                                   'Media Darkness             don't change yet
    Print #1, "^PQ" & NumLabels                         'Num of Labels              want two labels (one for sheet, one for vial)
    Print #1, "^PW450"                                  'Label Print Width in Pixel (1.5 inch = 450 dot width)
    '/////////////////// 1.5 x 0.75 = 450 x 225 dots (X =  0 to 450, Y = 0 to 225)
    Print #1, "^FO360,148"                              'DataMatrix Location
    Print #1, "^BXN,3,200"                            'DataMatrix Density
    Print #1, "^FD" & Datamatrix & "^FS"                        'DataMatrix Data (this is the yyyymmdd date)
    'Print #1, "^FO70,85   ^GB42,150,21^FS" 'print black box Original
    Print #1, "^FO70,75   ^GB42,180,21^FS" 'print black box NEW
    If Len(MatNum) <= 9 Then 'if material number is standard length
        Print #1, "^LRY ^FO80,65  ^A@B,30,30,E:ARIALNB.FNT    ^FD" & MatNum & "    ^FS" 'print matnum bottom to top in black box (standard matnum length)
    ElseIf Len(MatNum) > 9 And Len(MatNum) < 12 Then
        Print #1, "^LRY ^FO80,63  ^A@B,30,26,E:ARIALNB.FNT    ^FD" & MatNum & "    ^FS" 'print matnum bottom to top in black box (long matnum length...may need to play with this)
    ElseIf Len(MatNum) >= 12 Then
        Print #1, "^LRY ^FO80,60  ^A@B,30,26,E:ARIALNB.FNT    ^FD" & MatNum & "    ^FS" ' for the long ones
    End If
    Print #1, "     ^FO80,18  ^A@B,28,28,E:ARIALN.FNT    ^FD" & CompBatch & "    ^FS" 'print batch above black box
    Print #1, "^FB320,2     ^FO120,40  ^A@N,24.5,24,E:ARIALN.FNT    ^FD" & MatDesc & "    ^FS" 'print DescOne to the right of batch...wrapped to second line if DescOne string is long
    APLine = 69
    'Desc3Line = 96
    If Len(MatDesc) > 25 Then
        APLine = 96  'if matdesc is wrapped to second line (if length is > than width of FB), then push DescTwo down to third line
        'Desc3Line = 123 'if DescOne is wrapped to second line (if length is > than width of FB), then push DescThree down to fourth line
    End If
    Print #1, "     ^FO120, " & APLine & "  ^A@N,24.5,24,E:ARIALN.FNT       ^FD" & "Affinity Purified" & "^FS"
    'Print #1, "     ^FO120, " & Desc3Line & "  ^A@N,24.5,24,E:ARIALN.FNT       ^FD" & DescThree & "^FS"
    Print #1, "^FO120,150^A@N,24.5,24,E:ARIALN.FNT^FH^FD" & VialAmt & "^FS"
    Print #1, "      ^FO120,185  ^A@N,24.5,24.5,E:ARIALN.FNT    ^FH^FD" & "Store at 2 - 8_F8 C" & "    ^FS"
    Print #1, "      ^FO120,217  ^A@N,17.7,16,E:ARIALNB.FNT    ^FD" & "For in vitro research use only" & "    ^FS"
    Print #1, "^FO345,200^IME:BL_BW_H30.GRF^FS" 'Bethyl Logo
    Print #1, "^XZ" & vbCrLf 'End Label
    'Close #1
End Sub
Sub VialLabelPrint(MyPrinter, NumLabels, MatNum, MatDesc, CompBatch, Datamatrix, VialAmt)
    'toggle between this version and v2 when label is off after label roll is replaced
    Dim APLine As Integer, PrinterOpen As Boolean
    'CompBatch = "101"
    'Open MyPrinter For Output As #1
    Print #1, "^XA"                                     'Start Label   don't change
    Print #1, "^LT0"                                   'Label Top                  might need to tweak for 1.5 x 0.75 label
    Print #1, "~SD25"                                   'Media Darkness             don't change yet
    Print #1, "^PQ" & NumLabels                         'Num of Labels              want two labels (one for sheet, one for vial)
    Print #1, "^PW450"                                  'Label Print Width in Pixel (1.5 inch = 450 dot width)
    Print #1, "^PR6"
    '/////////////////// 1.5 x 0.75 = 450 x 225 dots (X =  0 to 450, Y = 0 to 225)
    Print #1, "^FO360,148"                              'DataMatrix Location
    Print #1, "^BXN,3,200"                            'DataMatrix Density
    Print #1, "^FD" & Datamatrix & "^FS"                        'DataMatrix Data (this is the yyyymmdd date)
    'Print #1, "^FO70,85   ^GB42,150,21^FS" 'print black box Original
    Print #1, "^FO90,75   ^GB42,170,21^FS" 'print black box NEW
    If Len(MatNum) <= 9 Then 'if material number is standard length
        Print #1, "^LRY ^FO100,65  ^A@B,30,30,E:ARIALNB.FNT    ^FD" & MatNum & "    ^FS" 'print matnum bottom to top in black box (standard matnum length)
    ElseIf Len(MatNum) > 9 And Len(MatNum) < 12 Then
        Print #1, "^LRY ^FO100,63  ^A@B,30,26,E:ARIALNB.FNT    ^FD" & MatNum & "    ^FS" 'print matnum bottom to top in black box (long matnum length...may need to play with this)
    ElseIf Len(MatNum) >= 12 Then
        Print #1, "^LRY ^FO100,60  ^A@B,30,26,E:ARIALNB.FNT    ^FD" & MatNum & "    ^FS" ' for the long ones
    End If
    Print #1, "     ^FO100,10  ^A@B,28,28,E:ARIALN.FNT    ^FD" & CompBatch & "    ^FS" 'print batch above black box
    Print #1, "^FB320,2     ^FO140,40  ^A@N,24.5,24,E:ARIALN.FNT    ^FD" & MatDesc & "    ^FS" 'print DescOne to the right of batch...wrapped to second line if DescOne string is long
    APLine = 69
    'Desc3Line = 96
    If Len(MatDesc) > 25 Then
        APLine = 96  'if matdesc is wrapped to second line (if length is > than width of FB), then push DescTwo down to third line
        'Desc3Line = 123 'if DescOne is wrapped to second line (if length is > than width of FB), then push DescThree down to fourth line
    End If
    Print #1, "     ^FO140, " & APLine & "  ^A@N,24.5,24,E:ARIALN.FNT       ^FD" & "Affinity Purified" & "^FS"
    'Print #1, "     ^FO120, " & Desc3Line & "  ^A@N,24.5,24,E:ARIALN.FNT       ^FD" & DescThree & "^FS"
    Print #1, "^FO140,150^A@N,24.5,24,E:ARIALN.FNT^FH^FD" & VialAmt & "^FS"
    Print #1, "      ^FO140,185  ^A@N,24.5,24.5,E:ARIALN.FNT    ^FH^FD" & "Store at 2 - 8_F8 C" & "    ^FS"
    Print #1, "      ^FO140,217  ^A@N,17.7,16,E:ARIALNB.FNT    ^FD" & "For in vitro research use only" & "    ^FS"
    Print #1, "^FO345,200^IME:BL_BW_H30.GRF^FS" 'Bethyl Logo
    Print #1, "^XZ" & vbCrLf 'End Label
    'Close #1
End Sub
Sub ResetPrinter()
    Dim MyPrinter
    MyPrinter = "\\C2186\Vialing#2_S4M"
    Open MyPrinter For Output As #1
    Print #1, "~JR"
    Close #1
End Sub
Sub ScanInMTops()
    Dim MyMBarcode As String, MyMatNum As String, MyBatch As String
    Range("AC" & Rows.Count).End(xlUp).Offset(1, 0).Select
    MyMBarcode = "something"
    While MyMBarcode <> ""
        MyMBarcode = InputBox("Scan in the '-M' material number.", "Scan barcode")
        If MyMBarcode <> "" Then
            MyMatNum = Replace(Left(MyMBarcode, InStr(MyMBarcode, "_")), "_", "")
            MyBatch = Replace(Replace(MyMBarcode, MyMatNum, ""), "_", "")
            ActiveCell.Value = MyMatNum
            ActiveCell.Offset(0, 5).Value = MyBatch
            ActiveCell.Offset(0, 12).Value = 1
            ActiveCell.Offset(0, 13).Value = 0
        End If
        ActiveCell.Offset(1, 0).Select
        DoEvents
    Wend
End Sub
Sub PrintAllComboLabels()
    Dim MyPrinter, NumLabels As String, MatNum As String, BatchNum As String, MatBatch As String, LineOne As String, LineTwo As String, LineThree As String, LineFour As String, CircleOne As String, CircleTwo As String, RowCount As Integer, i As Integer
    Dim FirstRow As Integer
    MyPrinter = "\\C2186\Vialing2_GX430t"
    Open MyPrinter For Output As #1
    Range("AC2").Select
    While Range("AC" & ActiveCell.Row) <> ""
        'If InStr(Range("AC" & ActiveCell.Row), "-M") > 0 And Range("AP" & ActiveCell.Row) = 0 Then
        If InStr(Range("AC" & ActiveCell.Row), "-M") > 0 Then
            MatNum = Range("AC" & ActiveCell.Row)
            BatchNum = Range("AH" & ActiveCell.Row)
            MatBatch = MatNum & "_" & BatchNum
            NumLabels = Range("AO" & ActiveCell.Row)
            NumLabels = 1
            MatBatch = MatNum & "_5f" & BatchNum
            LineOne = MatBatch
            CircleOne = MatBatch
            If NumLabels <> "" Then
                Call ComboLabelPrint(MyPrinter, NumLabels, LineOne, LineTwo, LineThree, LineFour, CircleOne, CircleTwo)
            End If
        End If
        ActiveCell.Offset(1, 0).Select
        DoEvents
        'Close #1
    Wend
    Close #1
End Sub
Sub ComboLabelPrint(MyPrinter, NumLabels, LineOne, LineTwo, LineThree, LineFour, CircleOne, CircleTwo)
    Print #1, "^XA"                                     'Start Label   don't change
    Print #1, "^LT25"                                   'Label Top                  might need to tweak for 1.5 x 0.75 label
    Print #1, "~SD0"                                   'Media Darkness             don't change yet
    Print #1, "^PQ" & NumLabels                         'Num of Labels              want two labels (one for sheet, one for vial)
    Print #1, "^PW440"                                  'Label Print Width in Pixel (1.5 inch = 450 dot width)
    'rectangle
    '    If InStr(LineOne, "-M") > 0 Then
    '        Print #1, "^FO8,20   ^GB240,80,45^FS" 'print black box for Ms
    '    End If
    Print #1, "^LRY^CI34^FO48,28^A@N,68,32,E:ARIALBD.TTF^FH^FD" & LineOne & "^FS"
    Print #1, "^XZ" 'End Label
End Sub
Sub PrintComboLabel_MStock()
    Dim MyPrinter, NumLabels As Integer, LineOne As String, LineTwo As String, LineThree As String, LineFour As String, CircleOne As String, CircleTwo As String, Datamatrix As String, ReformulationCount As Integer, ReformulationShortText As String
    MyPrinter = "\\C2186\Vialing2_GX430t"
    Open MyPrinter For Output As #1
    NumLabels = 1
    Datamatrix = Range("AC" & ActiveCell.Row) & "_" & Range("AH" & ActiveCell.Row)
    If InStr(Range("AC" & ActiveCell.Row).Value, "301-985") = 0 Then
        If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
            LineOne = Range("AC" & ActiveCell.Row) & " (lot " & Range("AH" & ActiveCell.Row) & ")"
        Else
            ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
            ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
            LineOne = Range("AC" & ActiveCell.Row) & " (lot " & Range("AH" & ActiveCell.Row) & "0" & ReformulationCount & ")"
            Range("AR" & ActiveCell.Row) = "RF"
            'stop
        End If
    Else
        If WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False) = "" Then
            LineOne = Range("AC" & ActiveCell.Row) & " (lot " & Range("AH" & ActiveCell.Row) & ")"
        Else
            ReformulationShortText = WorksheetFunction.VLookup(Replace(Range("AC" & ActiveCell.Row), "-M", "") & "100" & "_" & Range("AH" & ActiveCell.Row), Sheets("zcharvalues").Range("G:J"), 3, False)
            ReformulationCount = Len(ReformulationShortText) - Len(Replace(ReformulationShortText, "_", "")) + 1
            LineOne = Range("AC" & ActiveCell.Row) & " (lot " & Range("AH" & ActiveCell.Row) & "0" & ReformulationCount & ")"
            Range("AH" & ActiveCell.Row) = CInt(Range("AH" & ActiveCell.Row) & "0" & ReformulationCount)
            Range("AR" & ActiveCell.Row) = "RF"
            'stop
        End If
    End If
    LineTwo = Replace(Replace(Replace(WorksheetFunction.VLookup(Range("AC" & ActiveCell.Row), Sheets("LabelData").Range("A:K"), 3, False), "Rabbit", "RB"), "Goat", "GT"), "anti-", "X-")
    LineFour = Range("AJ" & ActiveCell.Row) & " / " & Range("AK" & ActiveCell.Row) 'stock + diluent
    CircleTwo = Replace(Replace(Range("AC" & ActiveCell.Row), "A30", ""), "A-M", "")
    CircleOne = Left(CircleTwo, 1)
    CircleTwo = Right(CircleTwo, 3)
    Call ComboLabelPrint_Mstock(MyPrinter, NumLabels, LineOne, LineTwo, LineThree, LineFour, CircleOne, CircleTwo, Datamatrix)
    Close #1
End Sub
Sub ComboLabelPrint_Mstock(MyPrinter, NumLabels, LineOne, LineTwo, LineThree, LineFour, CircleOne, CircleTwo, Datamatrix)
    Print #1, "^XA"
    Print #1, "^LS0"
    Print #1, "^LT0"
    Print #1, "~SD8"
    Print #1, "^PQ" & NumLabels
    Print #1, "^PW440"
    If Datamatrix <> "" Then
        Print #1, "^FO205,85"                              'DataMatrix Location
        Print #1, "^BXN,3,200"                            'DataMatrix Density
        Print #1, "^FD" & Datamatrix & "^FS"                   'DataMatrix Data (this is the yyyymmdd date)
    End If
    'rectangle
    Print #1, "^CI34        ^FO5,10  ^A@N,26,25,E:ARIALBD.TTF^FH^FD" & LineOne & "^FS" 'material and batch
    Print #1, "^CI34^FB260,2^FO5,55  ^A@N,25,24,E:ARIALBD.TTF^FH^FD" & LineTwo & "^FS"
    Print #1, "^CI34        ^FO5,112 ^A@N,27,26,E:ARIALBD.TTF^FH^FD" & LineFour & "^FS" 'stock + diluent
    'circle
    Print #1, "^CI34        ^FO367,30^A@N,45,45,E:ARIALBD.TTF^FH^FD" & CircleOne & "^FS" 'first part
    Print #1, "^CI34        ^FO337,70^A@N,55,55,E:ARIALBD.TTF^FH^FD" & CircleTwo & "^FS" 'second part
    Print #1, "^XZ" 'End Label
End Sub

