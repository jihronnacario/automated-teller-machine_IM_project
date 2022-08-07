Imports System.Data.OleDb

Public Class Form5
    Dim cmd, cmd2 As OleDbCommand
    Dim reader As OleDbDataReader
    Dim con As OleDbConnection

    Private Sub back_btn()
        Dim ask_back As String

        ask_back = MsgBox("Sure?", vbYesNo + vbQuestion, "Cancel")
        If ask_back = vbYes Then
            Me.Dispose()
            Form2.Show()
        Else
            ask_back = vbCancel
        End If
    End Sub

    Private Sub delete_record()
        Dim qry, ask, qry1, qry2 As String
        Dim txt, account_number, found As Integer

        Try
            If TextBox1.Text = "" Or IsNumeric(TextBox1.Text) = False Or TextBox1.Text.Length() < 8 Then
                MsgBox("Please provide with valid account number.", vbOKOnly, "Invalid Input")
                TextBox1.Clear()
                TextBox1.Focus()
            Else
                Using con = New OleDbConnection(connection)
                    Try
                        account_number = Convert.ToInt32(TextBox1.Text)
                        con.Open()
                        qry2 = "SELECT account_no FROM tblClient;"
                        qry = "DELETE FROM tblClient WHERE account_no = " & account_number & ";"
                        qry1 = "DELETE FROM tblTransaction WHERE transID = " & Form4.get_id(TextBox1.Text) & ";"
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
                        Else
                            ask = MsgBox("Are you sure to delete this record?", vbYesNo + vbQuestion, "Delete Record.")
                            If ask = vbYes Then
                                cmd = New OleDbCommand(qry, con)
                                cmd.ExecuteNonQuery()
                                cmd = New OleDbCommand(qry1, con)
                                cmd.ExecuteNonQuery()

                                MsgBox("Record Deleted.", vbOKOnly + vbInformation, "")
                                TextBox1.Clear()
                                TextBox1.Focus()
                            Else
                                ask = vbCancel
                            End If
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        back_btn()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        delete_record()
    End Sub
End Class