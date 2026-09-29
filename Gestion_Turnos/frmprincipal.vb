Public Class frmprincipal
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frmPacientes.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Application.Exit()
    End Sub
End Class
