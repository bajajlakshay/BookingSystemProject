Imports System.Data.OleDb
Public Class Form2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click


        If txtCID.Text = String.Empty Or txtFirstName.Text = String.Empty OrElse txtLastName.Text = String.Empty Then
            MsgBox("One or two fields cannot be empty")
        Else
            Dim stConString As String
            stConString = "Provider=Microsoft.ACE.OLEDB.12.0; Data source=F:\Project\BookingSystem\BS.accdb"
            Dim conAddCustomers As OleDbConnection
        conAddCustomers = New OleDbConnection
        conAddCustomers.ConnectionString = stConString
        conAddCustomers.Open()
        Dim stSQLInsert As String
        '  stSQLInsert = "INSERT INTO tblCustomers (CustomerID, FirstName,LastName) VALUES('DB001','Devaksha','Bhatt')"
        Dim cmdAddCustomer As OleDbCommand
            cmdAddCustomer = New OleDbCommand
            cmdAddCustomer.Connection = conAddCustomers
            stSQLInsert = "INSERT INTO tblCustomers(CustomerID,FirstName,LastName) VALUES('" & Me.txtCID.Text & "',' " & Me.txtFirstName.Text & " ' ,' " & Me.txtLastName.Text & " ')"
            cmdAddCustomer.CommandText = stSQLInsert
            cmdAddCustomer.ExecuteNonQuery()
        End If

    End Sub

End Class