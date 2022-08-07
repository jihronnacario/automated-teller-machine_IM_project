Public Class Form2
    Private Sub log_out()
        Dim ask_log_out As String

        ask_log_out = MsgBox("Sure ?", vbYesNo + vbQuestion, "Logout")
        If ask_log_out = vbYes Then
            Me.Dispose()
            Form1.Show()
            Form1.TextBox1.Focus()
        Else
            ask_log_out = vbCancel
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        log_out()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Dispose()
        Form3.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Dispose()
        Form4.get_account_number()
        Form4.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Dispose()
        Form6.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Dispose()
        Form5.Show()
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If usersType <> 1 Then
            Button3.Enabled = False
            Button4.Enabled = False
        End If
    End Sub
End Class