Public Class FormDestinations
    Private Sub Button_OK_Click(sender As Object, e As EventArgs) Handles Button_OK.Click
        If Me.ListBoxDest.SelectedItem Is Nothing Then
            MsgBox("Select a Destination", MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
        Else
            DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End If
    End Sub

    Private Sub ListBoxDest_DoubleClick(sender As Object, e As EventArgs) Handles ListBoxDest.DoubleClick
        DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub
End Class