Imports System.Security.Cryptography

Module global_variable
    Public account_number As Integer 'CLASS VARIABLE

    Public usersType As Integer

    Public connection As String = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                       "Data Source=|DataDirectory|\db_atm.accdb;" & _
                       "Persist Security Info=True;Jet OLEDB:Database Password=admin"

    Public Function MD5hashing(text As String)
        Dim hashed_pass As String
        Using MD5hash As MD5 = MD5.Create()
            hashed_pass = System.Convert.ToBase64String(MD5hash.ComputeHash(System.Text.Encoding.ASCII.GetBytes(text)))
        End Using
        Return hashed_pass
    End Function

End Module
