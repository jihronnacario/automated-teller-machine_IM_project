Imports System.Data.OleDb

Public Class Form1
    Dim cmd2 As OleDbCommand
    Dim reader As OleDbDataReader
    Dim con As OleDbConnection

    Private Sub clear_all()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub login()
        Dim pass, qry2, username, password, found, name As String

        found = ""
        name = ""
        Try

            Using con = New OleDbConnection(connection)
                Try
                    If TextBox1.Text <> "" And TextBox2.Text <> "" Then
                        pass = MD5hashing(TextBox2.Text)
                        con.Open()
                        qry2 = "SELECT * FROM tblUser;"
                        cmd2 = New OleDbCommand(qry2, con)
                        reader = cmd2.ExecuteReader()
                        While reader.Read()
                            username = reader(1)
                            name = reader(3)
                            password = reader(5)
                            usersType = reader(6)

                            If TextBox1.Text = username And pass = password Then
                                found = "password_username"
                                Exit While
                            Else
                                found = ""
                            End If
                        End While
                        reader.Close()
                        con.Close()

                        If found = "password_username" Then
                            MsgBox("Welcome " & name, vbOKOnly + vbInformation, "Login")
                            Me.Hide()
                            clear_all()
                            Form2.Show()
                        Else
                            MsgBox("Invalid password or username", vbOKOnly + vbCritical, "Failed")
                            clear_all()
                        End If
                    ElseIf TextBox1.Text = "" Then
                        MsgBox("Username cannot be empty.", vbOKOnly + vbCritical, "Failed")
                        TextBox1.Focus()
                    Else
                        MsgBox("Password cannot be empty.", vbOKOnly + vbCritical, "Failed")
                        TextBox2.Focus()
                    End If
                Catch ex As Exception
                    MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
                End Try
            End Using
        Catch ex As Exception
            MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
        End Try
    End Sub

    Private Sub quit_application()
        Dim ask_cancel As String

        ask_cancel = MsgBox("Are you sure?", vbYesNo + vbQuestion, "Confirm")
        If ask_cancel = vbYes Then
            Application.Exit()
        Else
            ask_cancel = vbCancel
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        quit_application()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        login()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Me.Hide()
        clear_all()
        Form7.Show()
    End Sub
End Class
