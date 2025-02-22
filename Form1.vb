Imports System.Data.OleDb
Public Class Form1
    Dim availabeIcon As New System.Drawing.Bitmap(My.Resources.available)
    Dim provisionalIcon As New System.Drawing.Bitmap(My.Resources.Provisional)
    Dim bookedIcon As New System.Drawing.Bitmap(My.Resources.Booked)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim c As Control

        For Each c In Me.Controls
            If TypeOf (c) Is PictureBox Then
                CType(c, PictureBox).Image = availabeIcon
                AddHandler c.Click, AddressOf PictureBox10_Click

            End If
        Next
        Call UpdateBookings()
    End Sub
    Sub UpdateBookings()
        Dim stSQL As String
        stSQL = "Select BookingID, Customer, Seat FROM tblBookings"

        Dim stConString As String
        stConString = "Provider=Microsoft.ACE.OLEDB.12.0; Data source=/BookingSystemProject/BS.accdb"

        Dim conBookings As OleDbConnection
        conBookings = New OleDbConnection
        conBookings.ConnectionString = stConString
        conBookings.Open()

        Dim cmdSelectBookings As OleDbCommand
        cmdSelectBookings = New OleDbCommand
        cmdSelectBookings.CommandText = stSQL
        cmdSelectBookings.Connection = conBookings

        'Dim cmdSelectBookings As New OleDbCommand(stSQL, conBookings)
        Dim dsBookings As New DataSet
        Dim daBookings As New OleDbDataAdapter(cmdSelectBookings)
        daBookings.Fill(dsBookings, "Bookings")

        conBookings.Close()
        ' MsgBox(dsBookings.Tables("Bookings").Rows.Count)
        ' MsgBox(dsBookings.Tables("Bookings").Rows(0).Item(1))

        Dim stOut As String
        Dim t1 As DataTable = dsBookings.Tables("Bookings")
        Dim Row As DataRow
        For Each Row In t1.Rows
            stOut = stOut & Row(0) & " " & Row(1) & Row(2) & vbNewLine
            CType(Controls("PictureBox" & Row(2)), PictureBox).Image = bookedIcon
        Next

        '   MsgBox(stOut)

    End Sub
    Private Sub PictureBox10_Click(sender As Object, e As EventArgs) Handles PictureBox10.Click
        If CType(sender, PictureBox).Image Is availabeIcon Then
            CType(sender, PictureBox).Image = provisionalIcon
        ElseIf CType(sender, PictureBox).Image Is provisionalIcon Then
            CType(sender, PictureBox).Image = availabeIcon
        End If
    End Sub

    Private Sub btnContinue_Click(sender As Object, e As EventArgs) Handles btnContinue.Click
        'Dim conBookings As OleDbConnection
        Dim c As Control
        Dim bSelect
        For Each c In Me.Controls
            If TypeOf (c) Is PictureBox Then
                If CType(c, PictureBox).Image Is provisionalIcon Then
                    bSelect = True
                End If
                AddHandler c.Click, AddressOf PictureBox10_Click
            End If
        Next


        If bSelect = False Then
            MsgBox("Please select at least one seat to book")
            Exit Sub
        End If
        If textCustomer.Text = String.Empty Then
            MsgBox("Enter customer ID to book a seat")
        Else

            Dim stConString As String
            stConString = "Provider=Microsoft.ACE.OLEDB.12.0; Data source=/BookingSystemProject/BS.accdb"
            Dim conBookings As OleDbConnection
            conBookings = New OleDbConnection
            conBookings.ConnectionString = stConString
            conBookings.Open()
            Dim stSQLInsert As String
            stSQLInsert = "INSERT INTO tblBookings (Customer, Seat) VALUES('LB001', 88)"
            Dim cmdMakeBookings As OleDbCommand
            cmdMakeBookings = New OleDbCommand

            cmdMakeBookings.Connection = conBookings

            Dim iSeatNum As Integer
            For Each c In Me.Controls
                If TypeOf (c) Is PictureBox Then
                    If CType(c, PictureBox).Image Is provisionalIcon Then
                        iSeatNum = Mid(CType(c, PictureBox).Name, 11)
                        stSQLInsert = "INSERT INTO tblBookings (Customer, Seat) VALUES('" & Me.textCustomer.Text & "', " & iSeatNum & ")"
                        cmdMakeBookings.CommandText = stSQLInsert
                        cmdMakeBookings.ExecuteNonQuery()
                    End If
                End If
            Next
            conBookings.Close()
            Call UpdateBookings()
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnViewCustomers.Click
        Dim stSQL As String
        stSQL = "SELECT CustomerID, FirstName,LastName FROM tblCustomers"

        Dim stConString As String
        stConString = "Provider=Microsoft.ACE.OLEDB.12.0; Data source=/BookingSystemProject/BS.accdb"

        Dim conCustomers As OleDbConnection
        conCustomers = New OleDbConnection
        conCustomers.ConnectionString = stConString
        conCustomers.Open()

        Dim cmdSelectCustomers As OleDbCommand
        cmdSelectCustomers = New OleDbCommand
        cmdSelectCustomers.CommandText = stSQL
        cmdSelectCustomers.Connection = conCustomers

        'Dim cmdSelectBookings As New OleDbCommand(stSQL, conBookings)
        Dim dsCustomers As New DataSet
        Dim daCustomers As New OleDbDataAdapter(cmdSelectCustomers)
        daCustomers.Fill(dsCustomers, "Customers")

        conCustomers.Close()
        ' MsgBox(dsBookings.Tables("Bookings").Rows.Count)
        ' MsgBox(dsBookings.Tables("Bookings").Rows(0).Item(1))

        Dim stOut As String
        Dim t1 As DataTable = dsCustomers.Tables("Customers")
        Dim Row As DataRow
        For Each Row In t1.Rows
            stOut = stOut & Row(0) & " " & Row(1) & Row(2) & vbNewLine
            ' CType(Controls("PictureBox" & Row(2)), PictureBox).Image = bookedIcon
        Next

        MsgBox(stOut)
    End Sub

    Private Sub btnAddCustomer_Click(sender As Object, e As EventArgs) Handles btnAddCustomer.Click
        '    Dim stConString As String
        '  stConString = "Provider=Microsoft.ACE.OLEDB.12.0; Data source=C:\Users\laksh\source\repos\BookingSystem\BS.accdb"
        '   Dim conAddCustomers As OleDbConnection
        '   conAddCustomers = New OleDbConnection
        '   conAddCustomers.ConnectionString = stConString
        '  conAddCustomers.Open()
        '  Dim stSQLInsert As String
        '  stSQLInsert = "INSERT INTO tblCustomers (CustomerID, FirstName,LastName) VALUES('DB001','Devaksha','Bhatt')"
        '   Dim cmdAddCustomer As OleDbCommand
        '   cmdAddCustomer = New OleDbCommand

        '   cmdAddCustomer.Connection = conAddCustomers
        '   stSQLInsert = "INSERT INTO tblCustomers(CustomerID,FirstName,LastName) VALUES('" & Me.txtCID.Text & "',' " & Me.txtFirstName.Text & " ' ,' " & Me.txtLastName.Text & " ' )"
        '  cmdAddCustomer.CommandText = stSQLInsert
        '  cmdAddCustomer.ExecuteNonQuery()

        Form2.Show()



    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim c As Control
        For Each c In Me.Controls
            If TypeOf (c) Is PictureBox Then
                CType(c, PictureBox).Image = availabeIcon
                AddHandler c.Click, AddressOf PictureBox10_Click

            End If
        Next
        Dim stConString As String
        stConString = "Provider=Microsoft.ACE.OLEDB.12.0; Data source=/BookingSystemProject/BS.accdb"
        Dim conDeleteBookings As OleDbConnection
        conDeleteBookings = New OleDbConnection
        conDeleteBookings.ConnectionString = stConString
        conDeleteBookings.Open()
        Dim stSQLInsert As String
        Dim cmdDeleteBookings As OleDbCommand
        cmdDeleteBookings = New OleDbCommand

        cmdDeleteBookings.Connection = conDeleteBookings
        stSQLInsert = "DELETE * From tblBookings"
        cmdDeleteBookings.CommandText = stSQLInsert
        cmdDeleteBookings.ExecuteNonQuery()

        Me.Refresh()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        Me.Close()

    End Sub

    <Obsolete>
    Private Sub btnBookingsView_Click(sender As Object, e As EventArgs) Handles btnBookingsView.Click
        Dim stSQL As String
        stSQL = "Select BookingID, Customer, Seat FROM tblBookings"

        Dim stConString As String
        stConString = "Provider=Microsoft.ACE.OLEDB.12.0; Data source=/BookingSystemProject/BS.accdb"

        Dim conBookings As OleDbConnection
        conBookings = New OleDbConnection
        conBookings.ConnectionString = stConString
        conBookings.Open()

        Dim cmdSelectBookings As OleDbCommand
        cmdSelectBookings = New OleDbCommand
        cmdSelectBookings.CommandText = stSQL
        cmdSelectBookings.Connection = conBookings

        'Dim cmdSelectBookings As New OleDbCommand(stSQL, conBookings)
        Dim dsBookings As New DataSet
        Dim daBookings As New OleDbDataAdapter(cmdSelectBookings)
        daBookings.Fill(dsBookings, "Bookings")

        conBookings.Close()
        ' MsgBox(dsBookings.Tables("Bookings").Rows.Count)
        ' MsgBox(dsBookings.Tables("Bookings").Rows(0).Item(1))

        Dim stOut As String
        Dim t1 As DataTable = dsBookings.Tables("Bookings")
        Dim Row As DataRow
        For Each Row In t1.Rows
            stOut = stOut & Row(0) & " " & Row(1) & " " & Row(2) & vbNewLine
            CType(Controls("PictureBox" & Row(2)), PictureBox).Image = bookedIcon
        Next

        MsgBox(stOut)
    End Sub
End Class