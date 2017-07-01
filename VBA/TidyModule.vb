'
' Copyright 2016-2017 Franck Latremoliere, Reckon LLP and others.
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

Sub TidySave()
    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = True
        ws.Select
        For Each sh In ws.Shapes
            If sh.Connector = msoFalse Then sh.IncrementTop 0
        Next sh
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
        Range("A1").Select
    Next ws
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
    Sheets(1).Select
    ActiveWindow.Width = 1280
    ActiveWindow.Height = 800
    ActiveWindow.Left = 4
    ActiveWindow.Top = 3
    On Error Resume Next
    For Each prop In ActiveWorkbook.BuiltinDocumentProperties
        prop.Value = ""
    Next
    On Error GoTo 0
    Let uName = Application.UserName
    Application.UserName = ChrW(&H2014)
    ' Use ChrW(&HD7) for a cross
    ' Use ChrW(&H2014) for an em dash
    ' Use ChrW(&H2702) for black scissors
    ' Use ChrW(&H263A) for smiling face
    ActiveWorkbook.Save
    Application.UserName = uName
    ActiveWorkbook.Close
End Sub
