Imports System.Data.OleDb

Public Class Form6
    Dim cmd, cmd2 As OleDbCommand
    Dim reader, reader2 As OleDbDataReader
    Dim con As OleDbConnection

    Private Sub cancel_btn()
        Dim ask_cancel As String

        ask_cancel = MsgBox("Sure?", vbYesNo + vbQuestion, "Cancel")
        If ask_cancel = vbYes Then
            Me.Dispose()
            Form2.Show()
            GroupBox1.Enabled = False
        Else
            ask_cancel = vbCancel
        End If
    End Sub

    Private Sub clearAll()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
    End Sub

    Private Sub clear_txtbx1()
        TextBox1.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub search_record()
        Dim qry, qry2 As String
        Dim txt, found As Integer

        Try
            If TextBox1.Text = "" Or IsNumeric(TextBox1.Text) = False Or TextBox1.Text.Length() < 8 Then
                MsgBox("Please provide with valid account number.", vbOKOnly, "Invalid Input")
                clear_txtbx1()
            Else
                Using con = New OleDbConnection(connection)

                    Try
                        account_number = Convert.ToInt32(TextBox1.Text)
                        con.Open()
                        qry = "SELECT lastname, firstname, middlename, birthdate, address, gender, job, company " & _
                                "FROM tblClient WHERE account_no = " & account_number & ";"
                        qry2 = "SELECT account_no FROM tblClient;"
                        cmd = New OleDbCommand(qry, con)
                        cmd2 = New OleDbCommand(qry2, con)
                        reader = cmd2.ExecuteReader()
                        While reader.Read()
                            txt = Convert.ToInt32(reader(0))
                            If account_number = txt Then
                                found = 1
                                Exit While
                            Else
                                found = 0
                            End If
                        End While
                        reader.Close()

                        If found = 0 Then
                            MsgBox("Account does not exist.", vbOKOnly + vbInformation, "No Record")
                            clear_txtbx1()
                        Else
                            GroupBox1.Enabled = True

                            reader2 = cmd.ExecuteReader()
                            While reader2.Read()
                                TextBox2.Text = reader2.Item(0).ToString()
                                TextBox3.Text = reader2.Item(1).ToString()
                                TextBox4.Text = reader2.Item(2).ToString()
                                DateTimePicker1.Value = reader2(3)
                                TextBox6.Text = reader2.Item(4).ToString()
                                TextBox7.Text = reader2.Item(5).ToString()
                                TextBox8.Text = reader2.Item(6).ToString()
                                TextBox9.Text = reader2.Item(7).ToString()
                            End While
                        End If
                        con.Close()
                    Catch ex As Exception
                        MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
                    End Try
                End Using
            End If
        Catch ex As Exception
            MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
        End Try
    End Sub

    Private Sub save_update()
        Dim date_string, verify As String
        Dim date_ As Date

        date_ = DateTimePicker1.Value()
        date_string = date_.ToString("yyyy-MM-dd")

        verify = MsgBox("All data input correct?", vbYesNo + vbQuestion, "Save Information")
        If verify = vbYes Then
            Try
                Using con = New OleDbConnection(connection)

                    Try
                        con.Open()
                        cmd.Connection = con
                        cmd.CommandText = "UPDATE tblClient SET lastname = ?, firstname = ?, " & _
                                "middlename = ?, birthdate = ?, address = ?, gender = ?, " & _
                                "job = ?, company = ? WHERE account_no = " & account_number & ";"

                        cmd.Parameters.Clear()
                        cmd.Parameters.AddWithValue("@p1", TextBox2.Text)
                        cmd.Parameters.AddWithValue("@p2", TextBox3.Text)
                        cmd.Parameters.AddWithValue("@p3", TextBox4.Text)
                        cmd.Parameters.AddWithValue("@p4", date_string)
                        cmd.Parameters.AddWithValue("@p5", TextBox6.Text)
                        cmd.Parameters.AddWithValue("@p6", TextBox7.Text)
                        cmd.Parameters.AddWithValue("@p7", TextBox8.Text)
                        cmd.Parameters.AddWithValue("@p8", TextBox9.Text)
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
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        cancel_btn()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "yyyy-MM-dd"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        search_record()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        save_update()
    End Sub
End Class