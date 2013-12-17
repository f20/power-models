Sub ListWorksheets()
    On Error GoTo FAILED
    Dim ls As Worksheet
    Worksheets.Add
    Set ls = ActiveSheet
    ls.Cells(1, 1) = "Name"
    ls.Cells(1, 2) = "Rows"
    ls.Cells(1, 3) = "Columns"
    Dim sl As Object
    Dim ws As Worksheet
    Set sl = ActiveWorkbook.Worksheets
    slc = sl.Count
    For j = 1 To slc
        Set ws = sl(j)
        ls.Cells(j + 1, 1) = ws.Name
        If ws.Name <> ls.Name Then
        ls.Cells(j + 1, 2) = ws.UsedRange.Rows.Count
        ls.Cells(j + 1, 3) = ws.UsedRange.Columns.Count
        End If
    Next j
    Return
FAILED:
    MsgBox "Failed"
End Sub
