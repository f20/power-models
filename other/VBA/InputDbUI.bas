'
' Copyright 2013 Franck Latremoliere, Reckon LLP and others.
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

Private Sub CommandButton1_Click()
    Call DoInputDB(True, False, Nothing)
End Sub

Private Sub CommandButton2_Click()
    Dim idbws As Worksheet
    On Error Resume Next
    Set idbws = ActiveWorkbook.sheets("InputDB")
    On Error GoTo 0
    If Not idbws Is Nothing Then
        Dim x As String
        x = Application.GetOpenFilename()
        If x Then
            Workbooks.Open x
            Call DoInputDB(False, False, idbws)
            idbws.Activate
            InputDbUI.Hide
        End If
    End If
End Sub
