'
' Copyright 2009-2013 Franck Latremoliere, Reckon LLP and others.
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

Sub ShowInputDBUI()
    InputDbUI.Show
End Sub

Sub DoInputDB(modifyThisBook As Boolean, internalFormulas As Boolean, wsDatabase As Worksheet)
    On Error GoTo RestoreSettings
    screenUpdateState = Application.ScreenUpdating
    statusBarState = Application.DisplayStatusBar
    calcState = Application.Calculation
    eventsState = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo 0
    
    If modifyThisBook Then
        sheets.Add
        Set wsDatabase = ActiveSheet
        On Error Resume Next
        wsDatabase.Name = "InputDB"
        On Error GoTo 0
    End If

    Dim dbLine As Long
    dbLine = 1
    wsDatabase.Rows("1:100").Insert Shift:=xlDown
    wsDatabase.Rows("1:100").Interior.ColorIndex = 1
    wsDatabase.Rows("1:100").Font.ColorIndex = 4

    wsDatabase.Cells(dbLine, 1).Value = "Table"
    wsDatabase.Cells(dbLine, 2).Value = "Column"
    wsDatabase.Cells(dbLine, 3).Value = "Row"
    wsDatabase.Cells(dbLine, 4).Value = "Key"
    wsDatabase.Cells(dbLine, 5).Value = "Value from " & ActiveWorkbook.Name

    Dim wsInput As Worksheet
    Dim numRows, numCols As Integer
    Dim cl As Range
    Dim table As Integer
    Dim twidth As Integer
    For Each wsInput In Worksheets
        If wsInput.Name <> wsDatabase.Name And Not (wsInput.Name Like "Calc*") And wsInput.UsedRange.row = 1 And wsInput.UsedRange.Column = 1 Then
            If modifyThisBook Then wsInput.Unprotect
            numRows = wsInput.UsedRange.Rows.Count
            numCols = wsInput.UsedRange.Columns.Count
            table = 0
            For r = 1 To wsInput.UsedRange.Rows.Count
                Dim row As String
                On Error Resume Next
                row = wsInput.Cells(r, 1).Value
                On Error GoTo 0
                If row Like "###.*" Then
                    table = Left(row, 3)
                    twidth = 0
                End If
                If row Like "####.*" Then
                    table = Left(row, 4)
                    twidth = 0
                End If
                If table = 0 Then GoTo DontDoThisRow
                If row Like "> *" Then
                    row = Right(row, Len(row) - 2)
                End If
                If wsInput.Cells(r - 1, 1).Value = "" And wsInput.Cells(r + 1, 1).Value = "" Then row = ""
                For c = 1 To wsInput.UsedRange.Columns.Count
                    Set cl = wsInput.Cells(r, c)
                    If c > twidth Then
                        If cl.Value <> "" Then twidth = c Else GoTo DontDoThisCell
                    End If
                    If cl.Locked Then GoTo DontDoThisCell
                    
                    ' It does not seem to be possible to test for the existence of the Validation property
                    ' Credits to Jamie Ham for workaround idea
                    Dim validationType As Long
                    validationType = xlValidateCustom ' something which is not xlValidateTextLength
                    On Error Resume Next
                    validationType = cl.Validation.Type
                    If validationType = xlValidateTextLength Then GoTo DontDoThisCell
                    ' if there is a validation by text length then the cell is not to be used
                    
                    On Error GoTo 0
                    dbLine = dbLine + 1
                    If dbLine Mod 100 = 1 Then wsDatabase.Rows(dbLine & ":" & (dbLine + 99)).Insert Shift:=xlDown
                    wsDatabase.Cells(dbLine, 1).Value = table
                    wsDatabase.Cells(dbLine, 2).Value = c - 1
                    wsDatabase.Cells(dbLine, 3).Value = row
                    wsDatabase.Cells(dbLine, 4).Formula = "=CONCATENATE(A" & dbLine & ",""|"",B" & dbLine & ",""|"",C" & dbLine & ")"
                    If internalFormulas Then
                        wsDatabase.Cells(dbLine, 5).FormulaR1C1 = "=INDEX('" & wsInput.Name & "'!R" & r & "C1:R" & r & "C" & twidth & "," & c & ")"
                    Else
                        With wsDatabase.Cells(dbLine, 5)
                            .Value = cl.Value
                            .WrapText = False
                        End With
                    End If
                    If modifyThisBook Then
                        tx = False
                        If cl.NumberFormat = "@" Then
                            tx = True
                            cl.NumberFormat = "General"
                        End If
                        If row = "" Then
                            cl.FormulaR1C1 = "=VLOOKUP(""" & table & "|" & (c - 1) & "|"",'" & wsDatabase.Name & "'!C4:C5,2,FALSE)"
                        Else
                            cl.FormulaR1C1 = "=VLOOKUP(""" & table & "|" & (c - 1) & "|""&RC1,'" & wsDatabase.Name & "'!C4:C5,2,FALSE)"
                        End If
                        If tx Then cl.NumberFormat = "@"
                        cl.Locked = True
                    End If
DontDoThisCell:
                On Error GoTo 0
                Next c
DontDoThisRow:
            Next r
            If modifyThisBook Then
                wsInput.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
                wsInput.EnableSelection = xlNoRestrictions
            End If
        End If
    Next wsInput
    ' Columns("A:B").ColumnWidth = 7.5
    ' Columns("C:C").ColumnWidth = 60
    ' Columns("D:D").ColumnWidth = 10
    ' Columns("E:E").ColumnWidth = 20
    ' Range("A1").AutoFilter
RestoreSettings:
    Application.ScreenUpdating = screenUpdateState
    Application.DisplayStatusBar = statusBarState
    Application.Calculation = calcState
    Application.EnableEvents = eventsState
End Sub
