Private Sub Workbook_Open()

    Call UpdateSourceQueries
    Call MakeCharValueTable
    ActiveWorkbook.Save
    Application.Quit

    '    Dim objConnection As Object, bBackground As Boolean
    '    'refresh all source data tables first...
    '    For Each objConnection In ThisWorkbook.Connections
    '        If InStr(objConnection, "KSSK") = 0 Then
    '            bBackground = objConnection.OLEDBConnection.BackgroundQuery
    '            objConnection.OLEDBConnection.BackgroundQuery = False
    '            objConnection.Refresh
    '            objConnection.OLEDBConnection.BackgroundQuery = bBackground
    '        End If
    '    Next
    '    'then find KSSK_INOB and refresh that table (compiles data from all other queries in workbook)
    '    For Each objConnection In ThisWorkbook.Connections
    '        If InStr(objConnection, "KSSK") > 0 Then
    '            bBackground = objConnection.OLEDBConnection.BackgroundQuery
    '            objConnection.OLEDBConnection.BackgroundQuery = False
    '            objConnection.Refresh
    '            objConnection.OLEDBConnection.BackgroundQuery = bBackground
    '        End If
    '    Next
    '    'DoEvents

    ''    Application.Quit
End Sub
