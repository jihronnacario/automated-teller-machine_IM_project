Imports System.Data.OleDb

Public Class Form7
    Dim cmd, cmd2 As OleDbCommand
    Dim reader, reader2 As OleDbDataReader

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
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox1.Focus()
    End Sub

    Function check_account() As Object
        Dim qry, qry2, user As String
        Dim found, found2 As Integer

        Try
            If check_null_input() = 0 Then
                MsgBox("Please fill the form properly.", vbOKOnly, "Empty Input")
            Else
                Using con = New OleDbConnection(connection)
                    Try
                        con.Open()
                        qry = "SELECT usertype FROM tblUser;"
                        qry2 = "SELECT username FROM tblUser;"
                        cmd = New OleDbCommand(qry, con)
                        cmd2 = New OleDbCommand(qry2, con)
                        reader = cmd.ExecuteReader()
                        reader2 = cmd2.ExecuteReader()
                        While reader.Read()
                            If reader(0) = 1 Then
                                found = 1
                                Exit While
                            Else
                                found = 0
                            End If
                        End While
                        reader.Close()

                        While reader2.Read()
                            user = reader2(0)
                            If TextBox4.Text = user Then
                                found2 = 1
                                Exit While
                            Else
                                found2 = 0
                            End If
                        End While
                        reader2.Close()

                        con.Close()
                    Catch ex As Exception
                        MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
                    End Try
                End Using
            End If
        Catch ex As Exception
            MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
        End Try
        Return {found, found2}
    End Function

    Private Sub save_user()
        Dim verify, pass As String
        Dim type As Integer
        Dim obj As Object

        obj = check_account()

        If obj(0) = 0 Then
            type = 1 'ADMIN
        Else
            type = 2
        End If

        If check_null_input() = 1 Then
            If TextBox5.Text = TextBox6.Text Then
                verify = MsgBox("All data correct?", vbYesNo + vbQuestion, "Save Information")
                If verify = vbYes And obj(1) = 0 Then
                    Try
                        Using con = New OleDbConnection(connection)
                            Try
                                pass = MD5hashing(TextBox5.Text)
                                con.Open()
                                cmd.Connection = con
                                cmd.CommandText = "INSERT INTO tblUser(username, user_lastname," & _
                                        "user_firstname, user_middlename, user_password, usertype)" & _
                                        "VALUES(?, ?, ?, ?, ?, ?);"
                                cmd.Parameters.Clear()
                                cmd.Parameters.AddWithValue("@p1", TextBox4.Text)
                                cmd.Parameters.AddWithValue("@p2", TextBox1.Text)
                                cmd.Parameters.AddWithValue("@p3", TextBox2.Text)
                                cmd.Parameters.AddWithValue("@p4", TextBox3.Text)
                                cmd.Parameters.AddWithValue("@p5", pass)
                                cmd.Parameters.AddWithValue("@p6", type)
                                cmd.ExecuteNonQuery()
                                con.Close()

                                MsgBox("New Record Added", vbOKOnly + vbInformation, "")
                                clearAll()
                            Catch ex As Exception
                                MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
                            End Try
                        End Using
                    Catch ex As Exception
                        MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
                    End Try
                Else
                    MsgBox("Username already exist.", vbOKOnly + vbInformation, "Inavlid")
                    verify = vbCancel
                    TextBox4.Focus()
                End If
            Else
                MsgBox("Password dont't match.", vbOKOnly + vbInformation, "Invalid password")
                TextBox5.Clear()
                TextBox6.Clear()
                TextBox5.Focus()
            End If
        End If
    End Sub

    Function check_null_input()
        Dim flag As Integer

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or
            TextBox5.Text = "" Or TextBox6.Text = "" Then
            flag = 0
        Else
            flag = 1
        End If
        Return flag
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        cancel_btn()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Dispose()
        Form8.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        save_user()
    End Sub
End Class