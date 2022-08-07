Imports System.Data.OleDb

Public Class Form8
    Private Sub cancel_btn()
        Dim ask_cancel As String

        ask_cancel = MsgBox("Sure?", vbYesNo + vbQuestion, "Cancel")
        If ask_cancel = vbYes Then
            Me.Dispose()
            Form1.Show()
        Else
            ask_cancel = vbCancel
        End If
    End Sub

    Private Sub clearAll()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub save_update()
        Dim cmd As New OleDbCommand
        Dim pass, verify As String

        If check_null_input() = 0 Then
            MsgBox("Please fill the for properly", vbOKOnly + vbInformation, "Error")
        ElseIf get_username() = 1 Then
            MsgBox("Invalid password.", vbOKOnly + vbInformation, "Invalid")
            TextBox2.Focus()
        ElseIf get_username() = 2 Then
            If TextBox3.Text <> TextBox4.Text Then
                MsgBox("New password don't match.", vbOKOnly + vbInformation, "Error")
                TextBox3.Focus()
            Else
                verify = MsgBox("All data input correct?", vbYesNo + vbQuestion, "Save Information")
                If verify = vbYes Then
                    Try
                        Using con = New OleDbConnection(connection)

                            Try
                                pass = MD5hashing(TextBox3.Text)
                                con.Open()
                                cmd.Connection = con
                                cmd.CommandText = "UPDATE tblUser SET user_password = ?" & _
                                                    "WHERE username = '" & TextBox1.Text & "';"
                                cmd.Parameters.Clear()
                                cmd.Parameters.AddWithValue("@p1", pass)
                                cmd.ExecuteNonQuery()
                                con.Close()

                                MsgBox("Record Updated", vbOKOnly + vbInformation, "")
                                clearAll()
                            Catch ex As Exception
                                MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
                            End Try
                        End Using
                    Catch ex As Exception
                        MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
                    End Try
                Else
                    verify = vbCancel
                End If
            End If
        Else
            MsgBox("Invalid account", vbOKOnly + vbInformation, "Invalid")
            TextBox1.Focus()
        End If
    End Sub

    Function check_null_input()
        Dim flag As Integer

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            flag = 0
        Else
            flag = 1
        End If
        Return flag
    End Function

    Function get_username()
        Dim cmd As OleDbCommand
        Dim reader As OleDbDataReader
        Dim qry, passw As String
        Dim found As Integer
        Try
            If check_null_input() = 0 Then
                MsgBox("Please fill the form properly.", vbOKOnly, "Empty Input")
            Else
                Using con = New OleDbConnection(connection)
                    Try
                        passw = MD5hashing(TextBox2.Text)

                        con.Open()
                        qry = "SELECT username, user_password FROM tblUser;"
                        cmd = New OleDbCommand(qry, con)
                        reader = cmd.ExecuteReader()
                        While reader.Read()
                            If reader(0) = TextBox1.Text And reader(1) = passw Then
                                found = 2
                                Exit While
                            ElseIf reader(0) = TextBox1.Text And reader(1) <> passw Then
                                found = 1
                                Exit While
                            Else
                                found = 0
                            End If
                        End While
                        reader.Close()
                        con.Close()
                    Catch ex As Exception
                        MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
                    End Try
                End Using
            End If
        Catch ex As Exception
            MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
        End Try
        Return found
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        cancel_btn()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        save_update()
    End Sub
End Class