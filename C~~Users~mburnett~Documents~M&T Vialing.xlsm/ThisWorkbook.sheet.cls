Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Range("ZZ1") = vbNullString
    Range("AAA1") = vbNullString
    Sheets("Main").Range("AB2:AZ100000") = vbNullString
    With Sheets("Main").Range("AB2:AZ100000").Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Bold = False
        .Italic = False
    End With
    Sheets("Main").Range("AB2:AZ100000").Borders.LineStyle = xlNone
    Rows("2:100000").Interior.Pattern = xlNone
    ActiveWorkbook.Save
End Sub
Private Sub Workbook_Open()
    Call LogAll
    Call RetrieveData
    Sheets("Main").Range("DA1") = "Production"
End Sub