Sub ConvertVBASheets()
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        Call ActivateVBACode(wb)
    Next wb
End Sub

Sub MakeVBAModules()
    Dim wb As Workbook
    On Error Resume Next
    Set wb = ActiveWorkbook
    If Not ws Is Nothing Then Call ActivateVBACode(wb)
End Sub

Sub ActivateVBACode(wb As Workbook)
    Dim vbaSheet As Worksheet
    On Error Resume Next
    Set vbaSheet = wb.sheets("VBACode")
    On Error GoTo 0
    If Not vbaSheet Is Nothing Then
        Dim vbaCode As String
        For Each c In vbaSheet.UsedRange
            vbaCode = vbaCode & c.Value & Chr$(10)
        Next c
        wb.VBProject.VBComponents.Add(vbext_ct_StdModule).CodeModule.AddFromString (vbaCode)
        da = Application.DisplayAlerts
        Application.DisplayAlerts = False
        vbaSheet.Delete
        Application.DisplayAlerts = da
        On Error Resume Next
        Application.Run "'" & Application.Substitute(wb.Name, "'", "''") & "'!Autorun"
        On Error GoTo 0
    End If
End Sub

Sub ImportVBA()
    ChDir ThisWorkbook.Path
    Dim c As VBComponent
    Dim cs As VBComponents
    Set cs = ThisWorkbook.VBProject.VBComponents
    For Each c In cs
        x = c.Name & ".bas"
        If True Then ' c.Type = vbext_ct_Document
            Dim st As String
            fnum = FreeFile()
            On Error Resume Next
            Open x For Input As fnum
            st = Input$(LOF(fnum), fnum)
            Close #fnum
            On Error GoTo 0
            If Len(st) > 3 Then
            With c.CodeModule
                .DeleteLines 1, .CountOfLines
                .InsertLines 1, st
            End With
            End If
        Else
            cs.Remove c
            cs.Import x
        End If
    Next c
    ThisWorkbook.Save
End Sub

Sub ExportVBA()
    ChDir ThisWorkbook.Path
    Dim c As VBComponent
    For Each c In ThisWorkbook.VBProject.VBComponents
        Dim x As String
        x = c.Name & ".bas"
        If True Then ' c.Type = vbext_ct_Document
            If c.CodeModule.CountOfLines > 0 Then
                fnum = FreeFile()
                On Error Resume Next
                Kill x
                On Error GoTo 0
                Open x For Output As fnum
                Dim st As String
                st = c.CodeModule.Lines(1, c.CodeModule.CountOfLines - 1)
                Print #fnum, st
                st = c.CodeModule.Lines(c.CodeModule.CountOfLines, 1)
                If st <> "" Then Print #fnum, st
                Close #fnum
            End If
        Else
            c.Export x
        End If
    Next c
End Sub
