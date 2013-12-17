Private Sub Workbook_Open()
    Call MakeCommandBar
    Call ConvertVBASheets
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.CommandBars("Franck Spreadsheet Tools").Delete
End Sub
