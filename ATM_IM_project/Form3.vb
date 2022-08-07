Imports System.Data.OleDb

Public Class Form3
    Dim cmd As OleDbCommand
    Dim da As OleDbDataAdapter
    Dim con As OleDbConnection

    Dim columnIndex As Integer

    Private Sub back_btn()
        Dim ask_back As String

        ask_back = MsgBox("Sure?", vbYesNo + vbQuestion, "Back")
        If ask_back = vbYes Then
            Me.Dispose()
            Form2.Show()
        Else
            ask_back = vbCancel
        End If
    End Sub

    Public Sub showTable()
        Dim dt As New DataTable
        Dim qry As String

        Try
            Using con = New OleDbConnection(connection)

                con.Open()
                qry = "SELECT tblClient.ID, tblClient.account_no, tblClient.pin_code, tblClient.lastname, tblClient.firstname, " & _
                    "tblClient.middlename, tblClient.birthdate, tblClient.address, tblClient.gender, tblClient.job, tblClient.company, " & _
                    "tblTransaction.balance, tblTransaction.deposit_amount, tblTransaction.withdraw_amount, tblTransaction.withdraw_date " & _
                    "FROM tblClient INNER JOIN tblTransaction " & _
                    "ON tblClient.ID = tblTransaction.transID;"
                da = New OleDbDataAdapter(qry, con)

                da.Fill(dt)
                DataGridView1.DataSource = dt.DefaultView
                con.Close()

            End Using
        Catch ex As Exception
            MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
        End Try
    End Sub

    Public Sub search_record()
        Dim dt As New DataTable

        Dim qry, key, query, lists() As String
        Dim flag, count As Integer

        query = ""
        lists = {"account_no", "lastname", "firstname", "middlename", "birthdate", "address", "gender", "job", "company"}
        key = TextBox1.Text()
        Try
            Using con = New OleDbConnection(connection)

                con.Open()
                If key <> "" Then
                    For count = 0 To lists.Length - 1
                        If IsNumeric(key) = True And count = 0 Then
                            qry = "SELECT tblClient.ID, tblClient.account_no, tblClient.pin_code, tblClient.lastname, tblClient.firstname, " & _
                                "tblClient.middlename, tblClient.birthdate, tblClient.address, tblClient.gender, tblClient.job, tblClient.company, " & _
                                "tblTransaction.balance, tblTransaction.deposit_amount, tblTransaction.withdraw_amount, tblTransaction.withdraw_date " & _
                                "FROM tblClient INNER JOIN tblTransaction " & _
                                "ON tblClient.ID = tblTransaction.transID " & _
                                "WHERE " & lists(count) & " = " & key & ";"
                            cmd = New OleDbCommand(qry, con)
                            If cmd.ExecuteScalar() = 0 Then
                                flag = 0
                            Else
                                flag = 1
                                query = qry
                                Exit For
                            End If
                        ElseIf count > 0 Then
                            qry = "SELECT tblClient.ID, tblClient.account_no, tblClient.pin_code, tblClient.lastname, tblClient.firstname, " & _
                                "tblClient.middlename, tblClient.birthdate, tblClient.address, tblClient.gender, tblClient.job, tblClient.company, " & _
                                "tblTransaction.balance, tblTransaction.deposit_amount, tblTransaction.withdraw_amount, tblTransaction.withdraw_date " & _
                                "FROM tblClient INNER JOIN tblTransaction " & _
                                "ON tblClient.ID = tblTransaction.transID " & _
                                "WHERE " & lists(count) & " LIKE '%" & key & "%';"

                            cmd = New OleDbCommand(qry, con)
                            If cmd.ExecuteScalar() = 0 Then
                                flag = 0
                            Else
                                flag = 1
                                query = qry
                                Exit For
                            End If
                        End If
                    Next

                    If flag = 0 Then
                        MsgBox("No record")
                    Else
                        da = New OleDbDataAdapter(query, con)
                        da.Fill(dt)
                        DataGridView1.DataSource = dt.DefaultView
                        TextBox1.Clear()
                    End If
                Else
                    MsgBox("Search cannot be empty", vbOKOnly + vbCritical, "Error")
                    TextBox1.Focus()
                End If
                con.Close()
            End Using
        Catch ex As Exception
            MsgBox(ex.ToString(), vbOKOnly + vbInformation, "Error")
        End Try
    End Sub

    Private Sub sort_asc()
        DataGridView1.Sort(DataGridView1.Columns(0),
        System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    Private Sub sort_desc()
        DataGridView1.Sort(DataGridView1.Columns(0),
        System.ComponentModel.ListSortDirection.Descending)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        back_btn()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        showTable()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        search_record()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)
        sort_desc()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        'columnIndex = DataGridView1.CurrentCell.ColumnIndex
        'MsgBox(columnIndex)
    End Sub
End Class