

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
