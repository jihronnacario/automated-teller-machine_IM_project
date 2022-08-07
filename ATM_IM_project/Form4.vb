Imports System.Data.OleDb

Public Class Form4
    Dim cmd, cmd2 As OleDbCommand
    Dim reader As OleDbDataReader
    Dim con As OleDbConnection


    Private Sub cancel_btn()
        Dim ask_cancel As String

        ask_cancel = MsgBox("Sure?", vbYesNo + vbQuestion, "Cancel")
        If ask_cancel = vbYes Then
            Me.Dispose()
            Form2.Show()
        Else
            ask_cancel = vbCancel
        End If
    End Sub

    Private Sub save_to_database()
        Dim date_string, verify, pin, status As String
        Dim date_ As Date

        date_ = DateTimePicker1.Value()
        date_string = date_.ToString("yyyy-MM-dd")
        status = ""

        If check_null_input() = 1 Then
            verify = MsgBox("All data input correct?", vbYesNo + vbQuestion, "Save Information")
            If verify = vbYes Then
                Try
                    Using con = New OleDbConnection(connection)
                        Try
                            pin = MD5hashing(TextBox2.Text)
                            con.Open()
                            cmd = New OleDbCommand
                            cmd.Connection = con
                            cmd.CommandText = "INSERT INTO tblClient(account_no, pin_code, lastname," & _
                                    "firstname, middlename, birthdate, address, gender, job, company)" & _
                                    "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?);"
                            cmd.Parameters.Clear()
                            cmd.Parameters.AddWithValue("@p1", Convert.ToInt32(TextBox1.Text))
                            cmd.Parameters.AddWithValue("@p2", pin)
                            cmd.Parameters.AddWithValue("@p3", TextBox3.Text)
                            cmd.Parameters.AddWithValue("@p4", TextBox4.Text)
                            cmd.Parameters.AddWithValue("@p5", TextBox5.Text)
                            cmd.Parameters.AddWithValue("@p6", date_string)
                            cmd.Parameters.AddWithValue("@p7", TextBox6.Text)
                            cmd.Parameters.AddWithValue("@p8", TextBox7.Text)
                            cmd.Parameters.AddWithValue("@p9", TextBox8.Text)
                            cmd.Parameters.AddWithValue("@p10", TextBox9.Text)
                            cmd.ExecuteNonQuery()
                            con.Close()

                            status = "ok"
                        Catch ex As Exception
                            MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
                        End Try
                    End Using
                    If status = "ok" Then
                        Try
                            Using con = New OleDbConnection(connection)
                                Try
                                    con.Open()
                                    cmd2 = New OleDbCommand
                                    cmd2.Connection = con
                                    cmd2.CommandText = "INSERT INTO tblTransaction(transID, balance, deposit_amount)" & _
                                            "VALUES(?, ?, ?);"
                                    cmd2.Parameters.Clear()
                                    cmd2.Parameters.AddWithValue("@p1", Convert.ToInt32(get_id(TextBox1.Text)))
                                    cmd2.Parameters.AddWithValue("@p2", CDbl(TextBox10.Text))
                                    cmd2.Parameters.AddWithValue("@p3", CDbl(TextBox10.Text))
                                    cmd2.ExecuteNonQuery()
                                    con.Close()

                                    MsgBox("New Record Added", vbOKOnly + vbInformation, "")
                                    clear_info()
                                    get_account_number()
                                Catch ex As Exception
                                    MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
                                End Try
                            End Using
                        Catch ex As Exception
                            MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
                        End Try
                    End If
                Catch ex As Exception
                MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
            End Try
            Else
                verify = vbCancel
            End If
        Else
            MsgBox("Fill the form properly", vbOKOnly + vbCritical, "Erro")
        End If
    End Sub

    Function check_null_input()
        Dim flag As Integer

        If TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or
            TextBox6.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Or TextBox9.Text = "" Or
            TextBox10.Text = "" Then
            flag = 0
        Else
            flag = 1
        End If
        Return flag
    End Function

    Private Sub clear_info()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()
        TextBox2.Focus()
    End Sub

    Function get_id(account As String)
        Dim qry As String
        Dim id As Integer

        Try
            Using con = New OleDbConnection(connection)
                Try
                    con.Open()
                    qry = "SELECT ID FROM tblClient WHERE account_no = " & account & ";"
                    cmd = New OleDbCommand(qry, con)
                    reader = cmd.ExecuteReader()
                    While reader.Read()
                        id = Convert.ToInt32(reader(0))
                    End While
                    reader.Close()
                    con.Close()
                Catch ex As Exception
                    MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error_in")
                End Try
            End Using
        Catch ex As Exception
            MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
        End Try

        Return id
    End Function

    Public Sub get_account_number()
        Dim qry As String
        Dim txt As Integer

        Try
            Using con = New OleDbConnection(connection)
                Try
                    con.Open()
                    qry = "SELECT account_no FROM tblClient;"
                    cmd = New OleDbCommand(qry, con)
                    reader = cmd.ExecuteReader()
                    While reader.Read()
                        txt = Convert.ToInt32(reader(0)) + 1
                    End While
                    If txt = 0 Then
                        TextBox1.Text = "10000001"
                    Else
                        TextBox1.Text = txt.ToString()
                    End If
                    reader.Close()
                    con.Close()
                Catch ex As Exception
                    MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error_in")
                End Try
            End Using
        Catch ex As Exception
            MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        cancel_btn()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "yyyy-MM-dd"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        save_to_database()
    End Sub
End Class