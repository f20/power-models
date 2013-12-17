Private Sub Run_Click()
    Call sfWorker
    ShowFormulasUI.Hide
End Sub

Private Sub UserForm_Activate()
    If Not Me.WholeBook.Value Then
        Call sfWorker
        Me.Hide
    End If
End Sub
