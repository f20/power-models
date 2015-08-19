'
' Copyright 2013-2015 Franck Latremoliere, Reckon LLP and others.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'
' 1. Redistributions of source code must retain the above copyright notice,
' this list of conditions and the following disclaimer.
'
' 2. Redistributions in binary form must reproduce the above copyright notice,
' this list of conditions and the following disclaimer in the documentation
' and/or other materials provided with the distribution.
'
' THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
' EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
' THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'

Sub ConvertVBASheets()
    Dim wbook As Workbook
    For Each wbook In Application.Workbooks
        Call ActivateVBACode(wbook)
    Next wbook
End Sub

Sub MakeVBAModules()
    Dim wbook As Workbook
    On Error Resume Next
    Set wbook = ActiveWorkbook
    If Not ws Is Nothing Then Call ActivateVBACode(wbook)
End Sub

Sub ActivateVBACode(wbook As Workbook)
    Dim vbaSheet As Worksheet
    On Error Resume Next
    Set vbaSheet = wbook.sheets("VBACode")
    On Error GoTo 0
    If Not vbaSheet Is Nothing Then
        Dim vbaCode As String
        For Each component In vbaSheet.UsedRange
            vbaCode = vbaCode & component.Value & Chr$(10)
        Next component
        wbook.VBProject.VBComponents.Add(vbext_ct_StdModule).CodeModule.AddFromString (vbaCode)
        da = Application.DisplayAlerts
        Application.DisplayAlerts = False
        vbaSheet.Delete
        Application.DisplayAlerts = da
        On Error Resume Next
        Application.Run "'" & Application.Substitute(wbook.Name, "'", "''") & "'!Autorun"
        On Error GoTo 0
    End If
End Sub

Sub ImportVBAOne(wbook As Workbook)
    ChDir wbook.path
    Dim component As VBComponent
    Dim componentSet As VBComponents
    Set componentSet = wbook.VBProject.VBComponents
    For Each component In componentSet
        fileName = component.Name & ".bas"
        Dim vbaCode As String
        fnum = FreeFile()
        On Error Resume Next
        Open fileName For Input As fnum
        vbaCode = Input$(LOF(fnum), fnum)
        Close #fnum
        On Error GoTo 0
        If Len(vbaCode) > 3 Then
            With component.CodeModule
                .DeleteLines 1, .CountOfLines
                .InsertLines 1, vbaCode
            End With
        End If
    Next component
    wbook.Save
End Sub

Sub ExportVBAOne(wbook As Workbook)
    ChDir wbook.path
    Dim component As VBComponent
    For Each component In wbook.VBProject.VBComponents
        Dim fileName As String
        fileName = component.Name & ".bas"
        If component.CodeModule.CountOfLines > 0 Then
            fnum = FreeFile()
            On Error Resume Next
            Kill fileName
            On Error GoTo 0
            Open fileName For Output As fnum
            Dim vbaCode As String
            On Error Resume Next
            vbaCode = component.CodeModule.Lines(1, component.CodeModule.CountOfLines - 1)
            Print #fnum, vbaCode
            On Error GoTo 0
            vbaCode = component.CodeModule.Lines(component.CodeModule.CountOfLines, 1)
            If vbaCode <> "" Then Print #fnum, vbaCode
            Close #fnum
        End If
    Next component
End Sub

Sub ImportVBA()
    Call ImportVBAOne(ThisWorkbook)
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.path <> "" Then Call ImportVBAOne(wb)
    Next wb
End Sub

Sub ExportVBA()
    Call ExportVBAOne(ThisWorkbook)
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.path <> "" Then Call ExportVBAOne(wb)
    Next wb
End Sub
