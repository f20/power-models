'
' Copyright 2013 Franck Latrémolière, Reckon LLP and others.
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

Sub ShowFormulas()
    ShowFormulasUI.WholeBook.Value = True
    ShowFormulasUI.Done.Text = ""
    ShowFormulasUI.Show
End Sub
    
Sub ShowFormulasThisSheet()
    ShowFormulasUI.WholeBook.Value = False
    ShowFormulasUI.Done.Text = ""
    ShowFormulasUI.Show
End Sub
    
Sub sfWorkerSheet(ws)
    ws.Visible = True
    ws.Select
    If Not ShowFormulasUI.Replace.Value Then
        If Right(ws.Name, 1) = "=" Then GoTo NextSheet
        dispAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        On Error Resume Next
        sheets(ws.Name & "=").Delete
        On Error GoTo 0
        Application.DisplayAlerts = dispAlerts
        ws.Copy After:=ws
        ActiveSheet.Name = ws.Name & "="
        Set ws = ActiveSheet
    End If
    Dim prog As String
    prog = ShowFormulasUI.Done.Text
    ShowFormulasUI.Done.Text = prog & "Processing " & ws.Name & " (" & ws.UsedRange.Rows.Count & "x" & ws.UsedRange.Columns.Count & ")"
    DoEvents
    ws.Unprotect
    rowb = ws.UsedRange.row + ws.UsedRange.Rows.Count
    For r = rowb - 1 To ws.UsedRange.row Step -1
        f1 = ""
        For c = ws.UsedRange.Column To ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
            f = Cells(r, c).FormulaR1C1
            If Left(f, 1) = "=" Then
                f = Right(f, Len(f) - 1)
                If f = f1 Then
                    Cells(r, c) = ChrW$(&H2190)
                    If f = Cells(rowb, c).Value Then Cells(r + 1, c) = ChrW$(&H2191)
                    ' &H2196 diagonal arrow is not universally supported
                Else
                    f1 = f
                    g = Cells(r, c).Formula
                    On Error Resume Next
                    Cells(r, c) = "." & g
                    If Cells(r, c + 1).Formula <> "" Then Cells(r, c).WrapText = True
                    Cells(r, c).HorizontalAlignment = xlLeft
                    If f = Cells(rowb, c).Value Then Cells(r + 1, c) = ChrW$(&H2191)
                    On Error GoTo 0
                End If
            End If
            Cells(rowb, c) = f
        Next c
    Next r
        For c = ws.UsedRange.Column To ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
        Cells(rowb, c) = ""
    Next c
    prog = prog & ws.Name & " complete" & vbLf
    ShowFormulasUI.Done.Text = prog
    DoEvents
NextSheet:
End Sub

Sub sfWorker()

    On Error GoTo RestoreSettings
    screenUpdateState = Application.ScreenUpdating
    statusBarState = Application.DisplayStatusBar
    calcState = Application.Calculation
    eventsState = Application.EnableEvents
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    On Error GoTo 0
    ' So that we see potentially helpful error messages about the code below

    Dim sheetsToDo As Object
    If ShowFormulasUI.WholeBook.Value Then
        For Each ws In ActiveWorkbook.Worksheets
            sfWorkerSheet ws
        Next ws
    Else
        If Right(ActiveSheet.Name, 1) = "=" Then sheets(Left(ActiveSheet.Name, Len(ActiveSheet.Name) - 1)).Activate
        sfWorkerSheet ActiveSheet
    End If

RestoreSettings:
    Application.ScreenUpdating = screenUpdateState
    Application.DisplayStatusBar = statusBarState
    Application.Calculation = calcState
    Application.EnableEvents = eventsState

End Sub
