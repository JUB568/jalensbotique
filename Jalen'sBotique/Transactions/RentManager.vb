Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions

Public Class RentManager
    Private form As Form1
    Private cmd As MySqlCommand

    Public Sub New(frm As Form1)
        form = frm
    End Sub

    Public Sub LoadRentTransactions()
        Dim query As String = "
            SELECT
                r.Rent_ID,
                c.Full_Name,
                cl.Clothes_Name,
                r.Date_Rented,
                r.Expected_Return_Date,
                r.Actual_Return_Date,
                r.Rental_Status
            FROM rent r
            INNER JOIN customer c ON r.Customer_ID = c.Customer_ID
            INNER JOIN clothes cl ON r.Clothes_ID = cl.Clothes_ID
            ORDER BY r.Rent_ID DESC"

        LoadToDGV(query, form.RentDGV)
        form.AutoSelectFirstRow(form.RentDGV)
    End Sub

    Public Sub LoadCustomersToRentCombo()
        Try
            ConnDB.OpenConn()  ' ✅ Use module method
            Dim query As String = "SELECT Customer_ID, Full_Name FROM customer ORDER BY Full_Name"
            cmd = New MySqlCommand(query, ConnDB.conn)  ' ✅ Use module conn
            Dim adapter As New MySqlDataAdapter(cmd)
            Dim dt As New DataTable
            adapter.Fill(dt)

            form.RentCustomerNameCMB.DataSource = dt
            form.RentCustomerNameCMB.DisplayMember = "Full_Name"
            form.RentCustomerNameCMB.ValueMember = "Customer_ID"
            form.RentCustomerNameCMB.SelectedIndex = -1
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            ConnDB.CloseConn()  ' ✅ Use module method
        End Try
    End Sub

    Public Sub LoadAvailableClothesToCombo()
        Try
            ConnDB.OpenConn()
            Dim query As String = "
                SELECT Clothes_ID, Clothes_Name
                FROM clothes
                WHERE Status = 'Available'
                ORDER BY Clothes_Name"

            cmd = New MySqlCommand(query, ConnDB.conn)
            Dim adapter As New MySqlDataAdapter(cmd)
            Dim dt As New DataTable
            adapter.Fill(dt)

            form.RentClothesCMB.DataSource = dt
            form.RentClothesCMB.DisplayMember = "Clothes_Name"
            form.RentClothesCMB.ValueMember = "Clothes_ID"
            form.RentClothesCMB.SelectedIndex = -1
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            ConnDB.CloseConn()
        End Try
    End Sub

    Public Sub RentItem()
        ' ============================================
        ' 🔒 COMPLETE VALIDATION - ALL FIELDS REQUIRED
        ' ============================================

        If form.RentCustomerNameCMB.SelectedIndex = -1 Then
            MessageBox.Show("❌ **Customer** is REQUIRED!", "Field Required",
                       MessageBoxButtons.OK, MessageBoxIcon.Warning)
            form.RentCustomerNameCMB.Focus()
            Return
        End If

        If form.RentClothesCMB.SelectedIndex = -1 Then
            MessageBox.Show("❌ **Clothes** is REQUIRED!", "Field Required",
                       MessageBoxButtons.OK, MessageBoxIcon.Warning)
            form.RentClothesCMB.Focus()
            Return
        End If

        Dim rentDate = form.RentDateRentedDTP.Value.Date
        If rentDate > Date.Today Then
            MessageBox.Show("❌ Rent date cannot be in the future!", "Invalid Date")
            form.RentDateRentedDTP.Focus()
            Return
        End If

        Dim expectedReturnDate = form.RentExpectedReturnDTP.Value.Date
        If expectedReturnDate <= rentDate Then
            MessageBox.Show("❌ Expected return date must be after rent date!", "Invalid Date")
            form.RentExpectedReturnDTP.Focus()
            Return
        End If

        Try
            ConnDB.OpenConn()

            ' Check if clothes already rented
            Dim checkQuery As String = "SELECT COUNT(*) FROM rent r JOIN clothes c ON r.Clothes_ID = c.Clothes_ID WHERE c.Clothes_ID = @clothesId AND r.Rental_Status != 'Returned'"
            cmd = New MySqlCommand(checkQuery, ConnDB.conn)
            cmd.Parameters.AddWithValue("@clothesId", form.RentClothesCMB.SelectedValue)
            Dim rentedCount = Convert.ToInt32(cmd.ExecuteScalar())

            If rentedCount > 0 Then
                MessageBox.Show("❌ Selected clothes are already rented!", "Item Unavailable")
                Return
            End If

            ' Insert rent record
            Dim insertQuery As String = "
                INSERT INTO rent
                (Customer_ID, Clothes_ID, Date_Rented, Expected_Return_Date, Rental_Status)
                VALUES (@customerId, @clothesId, @dateRented, @expectedReturn, 'Rented')"

            cmd = New MySqlCommand(insertQuery, ConnDB.conn)
            cmd.Parameters.AddWithValue("@customerId", form.RentCustomerNameCMB.SelectedValue)
            cmd.Parameters.AddWithValue("@clothesId", form.RentClothesCMB.SelectedValue)
            cmd.Parameters.AddWithValue("@dateRented", rentDate)
            cmd.Parameters.AddWithValue("@expectedReturn", expectedReturnDate)

            cmd.ExecuteNonQuery()

            ' Update clothes status
            Dim updateClothesQuery As String = "UPDATE clothes SET Status = 'Rented' WHERE Clothes_ID = @id"
            cmd = New MySqlCommand(updateClothesQuery, ConnDB.conn)
            cmd.Parameters.AddWithValue("@id", form.RentClothesCMB.SelectedValue)
            cmd.ExecuteNonQuery()

            MessageBox.Show($"✅ Item rented successfully!" & vbCrLf &
                       $"👤 {form.RentCustomerNameCMB.Text}" & vbCrLf &
                       $"👗 {form.RentClothesCMB.Text}", "Success",
                       MessageBoxButtons.OK, MessageBoxIcon.Information)

            ClearRentFields()
            RefreshAllData()

        Catch ex As Exception
            MessageBox.Show($"❌ Database Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ConnDB.CloseConn()
        End Try
    End Sub

    Public Sub ReturnItem()
        If String.IsNullOrWhiteSpace(form.ReturnItemRentIDTB.Text) Then
            MessageBox.Show("❌ **Please select rent transaction**", "Selection Required",
                       MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            ConnDB.OpenConn()
            Dim rentId As Integer = Convert.ToInt32(form.ReturnItemRentIDTB.Text)

            Dim getClothesQuery As String = "SELECT Clothes_ID FROM rent WHERE Rent_ID = @rentId"
            cmd = New MySqlCommand(getClothesQuery, ConnDB.conn)
            cmd.Parameters.AddWithValue("@rentId", rentId)
            Dim clothesId = Convert.ToInt32(cmd.ExecuteScalar())

            Dim returnQuery As String = "
                UPDATE rent SET Rental_Status = 'Returned', Actual_Return_Date = NOW() 
                WHERE Rent_ID = @rentId"
            cmd = New MySqlCommand(returnQuery, ConnDB.conn)
            cmd.Parameters.AddWithValue("@rentId", rentId)
            cmd.ExecuteNonQuery()

            Dim updateClothesQuery As String = "
                UPDATE clothes SET Status = 'Available' WHERE Clothes_ID = @clothesId"
            cmd = New MySqlCommand(updateClothesQuery, ConnDB.conn)
            cmd.Parameters.AddWithValue("@clothesId", clothesId)
            cmd.ExecuteNonQuery()

            MessageBox.Show("✅ Item returned successfully!", "Success",
                       MessageBoxButtons.OK, MessageBoxIcon.Information)
            RefreshAllData()

        Catch ex As Exception
            MessageBox.Show($"❌ Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ConnDB.CloseConn()
        End Try
    End Sub

    Public Sub MarkLostItem()
        If String.IsNullOrWhiteSpace(form.ReturnItemRentIDTB.Text) Then
            MessageBox.Show("❌ **Please select rent transaction**", "Selection Required",
                       MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            ConnDB.OpenConn()
            Dim rentId As Integer = Convert.ToInt32(form.ReturnItemRentIDTB.Text)

            Dim getClothesQuery As String = "SELECT Clothes_ID FROM rent WHERE Rent_ID = @rentId"
            cmd = New MySqlCommand(getClothesQuery, ConnDB.conn)
            cmd.Parameters.AddWithValue("@rentId", rentId)
            Dim clothesId = Convert.ToInt32(cmd.ExecuteScalar())

            Dim lostQuery As String = "UPDATE rent SET Rental_Status = 'Lost' WHERE Rent_ID = @rentId"
            cmd = New MySqlCommand(lostQuery, ConnDB.conn)
            cmd.Parameters.AddWithValue("@rentId", rentId)
            cmd.ExecuteNonQuery()

            Dim updateClothesQuery As String = "UPDATE clothes SET Status = 'Lost' WHERE Clothes_ID = @clothesId"
            cmd = New MySqlCommand(updateClothesQuery, ConnDB.conn)
            cmd.Parameters.AddWithValue("@clothesId", clothesId)
            cmd.ExecuteNonQuery()

            MessageBox.Show("✅ Item marked as lost!", "Success",
                       MessageBoxButtons.OK, MessageBoxIcon.Information)
            RefreshAllData()

        Catch ex As Exception
            MessageBox.Show($"❌ Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ConnDB.CloseConn()
        End Try
    End Sub

    Public Sub ExtendItem()
        If String.IsNullOrWhiteSpace(form.ReturnItemRentIDTB.Text) Then
            MessageBox.Show("❌ **Please select rent transaction**", "Selection Required",
                   MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            ConnDB.OpenConn()
            Dim query As String = "
            UPDATE rent SET Expected_Return_Date = DATE_ADD(Expected_Return_Date, INTERVAL 3 DAY)
            WHERE Rent_ID = @rentId"

            cmd = New MySqlCommand(query, ConnDB.conn)
            cmd.Parameters.AddWithValue("@rentId", Convert.ToInt32(form.ReturnItemRentIDTB.Text))
            cmd.ExecuteNonQuery()

            MessageBox.Show("✅ Return date extended by 3 days!", "Success",
                   MessageBoxButtons.OK, MessageBoxIcon.Information)

            RefreshAllData()

        Catch ex As Exception
            MessageBox.Show($"❌ Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ConnDB.CloseConn()
        End Try
    End Sub

    Public Sub SearchRent(searchText As String)
        Dim query As String = "
            SELECT
                r.Rent_ID,
                c.Full_Name,
                cl.Clothes_Name,
                r.Date_Rented,
                r.Expected_Return_Date,
                r.Actual_Return_Date,
                r.Rental_Status
            FROM rent r
            INNER JOIN customer c ON r.Customer_ID = c.Customer_ID
            INNER JOIN clothes cl ON r.Clothes_ID = cl.Clothes_ID
            WHERE
                c.Full_Name LIKE @search
                OR cl.Clothes_Name LIKE @search
                OR r.Rental_Status LIKE @search
            ORDER BY r.Rent_ID DESC"

        Try
            ConnDB.OpenConn()
            cmd = New MySqlCommand(query, ConnDB.conn)
            cmd.Parameters.AddWithValue("@search", "%" & searchText & "%")
            Dim adapter As New MySqlDataAdapter(cmd)
            Dim dt As New DataTable
            adapter.Fill(dt)
            form.RentDGV.DataSource = dt
        Finally
            ConnDB.CloseConn()
        End Try
    End Sub

    Public Sub LoadRentEditFields(row As DataGridViewRow)
        form.ReturnItemRentIDTB.Text = form.GetCellValue(row.Cells("Rent_ID"))
        form.ReturnItemCustomerNameTB.Text = form.GetCellValue(row.Cells("Full_Name"))
        form.ReturnItemClothesNameTB.Text = form.GetCellValue(row.Cells("Clothes_Name"))
        form.ReturnItemStatusTB.Text = form.GetCellValue(row.Cells("Rental_Status"))
    End Sub

    Public Sub ClearRentFields()
        form.RentCustomerNameCMB.SelectedIndex = -1
        form.RentClothesCMB.SelectedIndex = -1
        form.RentDateRentedDTP.Value = Date.Now
        form.RentExpectedReturnDTP.Value = Date.Now.AddDays(3)
    End Sub

    Private Sub RefreshAllData()
        LoadRentTransactions()
        LoadAvailableClothesToCombo()
        If form.clothesManager IsNot Nothing Then
            form.clothesManager.LoadClothes()
        End If

        form.ReturnItemRentIDTB.Clear()
        form.ReturnItemCustomerNameTB.Clear()
        form.ReturnItemClothesNameTB.Clear()
        form.ReturnItemStatusTB.Clear()

        If form.RentDGV IsNot Nothing Then
            form.RentDGV.ClearSelection()
            form.RentDGV.CurrentCell = Nothing
        End If
    End Sub

    Private Sub LoadToDGV(query As String, dgv As DataGridView)
        ConnDB.LoadToDGV(query, dgv)
    End Sub
End Class