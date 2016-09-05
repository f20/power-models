Function GetAPassword(sheet As Worksheet) As String
  ' Author unknown but submitted by brettdj of www.experts-exchange.com
  ' Modified by Franck Latrémolière
  Dim i As Integer, j As Integer, k As Integer
  Dim l As Integer, m As Integer, n As Integer
  Dim i1 As Integer, i2 As Integer, i3 As Integer
  Dim i4 As Integer, i5 As Integer, i6 As Integer
  On Error Resume Next
  For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
  For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
  For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
  For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
    sheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
        Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
        Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    If sheet.ProtectContents = False Then
        GetAPassword = Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & _
            Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
        Exit Function
    End If
  Next: Next: Next: Next: Next: Next
  Next: Next: Next: Next: Next: Next
End Function

Public Sub PasswordBreakerAllSheets()
    Set sl = ActiveWorkbook.Worksheets
    slc = sl.Count
    For sn2 = 1 To slc
        Dim sheet As Worksheet
        Set sheet = sl(sn2)
        If sheet.ProtectContents = True Then
            pass = GetAPassword(sheet)
            For sn = 1 To slc
                If sl(sn).ProtectContents = True Then
                    sl(sn).Unprotect pass
                End If
            Next sn
        End If
    Next sn2
End Sub

