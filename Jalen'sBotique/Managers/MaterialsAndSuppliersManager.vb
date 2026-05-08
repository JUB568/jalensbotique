Imports MySql.Data.MySqlClient

Public Class MaterialsAndSuppliersManager
    Private form As Form1

    Public Sub New(frm As Form1)
        form = frm
    End Sub

    ' =========================
    ' MATERIALS INVENTORY PANEL
    ' =========================
    Public Sub LoadMaterialsSuppliers()
        LoadMaterials()
        LoadSuppliers()
    End Sub

    Public Sub LoadMaterials()
        Dim query As String = "SELECT m.*, s.Supplier_Name 
                          FROM Materials m 
                          LEFT JOIN Supplier s ON m.Supplier_ID = s.Supplier_ID 
                          ORDER BY m.Material_ID DESC"
        LoadToDGV(query, form.MaterialsDGV)
        form.AutoSelectFirstRow(form.MaterialsDGV)
    End Sub

    Public Sub LoadSuppliers()
        Dim query As String = "SELECT * FROM Supplier ORDER BY Supplier_Name"
        LoadToDGV(query, form.SuppliersDGV)
    End Sub

    Public Function GetSuppliers() As DataTable

        Dim dt As New DataTable()

        Try

            OpenConn()

            Dim query As String =
            "SELECT Supplier_ID, Supplier_Name 
             FROM Supplier 
             ORDER BY Supplier_Name"

            Dim cmd As New MySqlCommand(query, conn)
            Dim adapter As New MySqlDataAdapter(cmd)

            adapter.Fill(dt)

            ' Ensure columns exist even if table empty
            If Not dt.Columns.Contains("Supplier_ID") Then
                dt.Columns.Add("Supplier_ID")
            End If

            If Not dt.Columns.Contains("Supplier_Name") Then
                dt.Columns.Add("Supplier_Name")
            End If

        Catch ex As Exception

            Debug.WriteLine("GetSuppliers Error: " & ex.Message)

            ' Create fallback columns
            If Not dt.Columns.Contains("Supplier_ID") Then
                dt.Columns.Add("Supplier_ID")
            End If

            If Not dt.Columns.Contains("Supplier_Name") Then
                dt.Columns.Add("Supplier_Name")
            End If

        Finally

            CloseConn()

        End Try

        Return dt

    End Function

    Public Sub SearchMaterials(searchText As String)
        searchText = searchText.Trim()

        If String.IsNullOrWhiteSpace(searchText) Then
            LoadSuppliers()
            Return
        End If

        ' FIXED: Changed "Suppliers" to "Supplier" (singular)
        Dim query As String = "SELECT m.*, s.Supplier_Name 
                          FROM Materials m 
                          LEFT JOIN Supplier s ON m.Supplier_ID = s.Supplier_ID 
                          WHERE m.Material_Name LIKE '%" & searchText & "%' 
                          OR m.Description LIKE '%" & searchText & "%'
                          OR s.Supplier_Name LIKE '%" & searchText & "%'
                          ORDER BY m.Material_ID DESC"

        Try
            OpenConn()

            Dim cmd As New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@search", "%" & searchText & "%")

            Dim da As New MySqlDataAdapter(cmd)
            Dim dt As New DataTable()
            da.Fill(dt)

            form.MaterialsDGV.DataSource = dt

        Catch ex As Exception
            MessageBox.Show("Search error: " & ex.Message)
        Finally
            CloseConn()
        End Try
    End Sub

    Public Sub SearchSuppliers(searchText As String)
        searchText = searchText.Trim()

        If String.IsNullOrWhiteSpace(searchText) Then
            LoadSuppliers()
            Return
        End If

        Dim query As String = "SELECT * FROM Supplier 
                          WHERE Supplier_Name LIKE @search 
                          OR Contact_Number LIKE @search 
                          OR Address LIKE @search
                          OR Email LIKE @search
                          ORDER BY Supplier_Name"

        Try
            OpenConn()

            Dim cmd As New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@search", "%" & searchText & "%")

            Dim da As New MySqlDataAdapter(cmd)
            Dim dt As New DataTable()
            da.Fill(dt)

            form.SuppliersDGV.DataSource = dt

        Catch ex As Exception
            MessageBox.Show("Search error: " & ex.Message)
        Finally
            CloseConn()
        End Try
    End Sub


    Public Sub AddSuppliers()
        ' ============================================
        ' 🔒 COMPLETE VALIDATION - ALL FIELDS REQUIRED
        ' ============================================

        ' 1️⃣ SUPPLIER NAME (REQUIRED + MIN LENGTH)
        Dim supplierName = form.SuppliersNameTB.Text.Trim()
        If String.IsNullOrWhiteSpace(supplierName) Then
            MessageBox.Show("❌ **Supplier Name** is REQUIRED!", "Field Required",
                       MessageBoxButtons.OK, MessageBoxIcon.Warning)
            form.SuppliersNameTB.Focus()
            form.SuppliersNameTB.BackColor = Color.LightPink
            Return
        End If
        If supplierName.Length < 2 OrElse supplierName.Length > 100 Then
            MessageBox.Show("❌ Supplier name must be 2-100 characters!", "Invalid Length")
            form.SuppliersNameTB.Focus()
            form.SuppliersNameTB.BackColor = Color.LightPink
            Return
        End If

        ' 2️⃣ CONTACT NUMBER (REQUIRED + VALID FORMAT)
        Dim contact = form.SuppliersContactTB.Text.Trim()
        If String.IsNullOrWhiteSpace(contact) Then
            MessageBox.Show("❌ **Contact Number** is REQUIRED!", "Field Required")
            form.SuppliersContactTB.Focus()
            form.SuppliersContactTB.BackColor = Color.LightPink
            Return
        End If
        If Not IsValidPhoneNumber(contact) Then
            MessageBox.Show("❌ Invalid phone number!\n📱 Use: 09xxxxxxxxx or +63xxxxxxxxxx", "Invalid Phone")
            form.SuppliersContactTB.Focus()
            form.SuppliersContactTB.BackColor = Color.LightPink
            Return
        End If

        ' 3️⃣ ADDRESS (REQUIRED + MIN LENGTH)
        Dim address = form.SuppliersAddressTB.Text.Trim()
        If String.IsNullOrWhiteSpace(address) Then
            MessageBox.Show("❌ **Address** is REQUIRED!", "Field Required")
            form.SuppliersAddressTB.Focus()
            form.SuppliersAddressTB.BackColor = Color.LightPink
            Return
        End If
        If address.Length < 5 OrElse address.Length > 255 Then
            MessageBox.Show("❌ Address must be 5-255 characters!", "Invalid Length")
            form.SuppliersAddressTB.Focus()
            form.SuppliersAddressTB.BackColor = Color.LightPink
            Return
        End If

        ' 4️⃣ EMAIL (OPTIONAL BUT VALID IF PROVIDED)
        Dim email = form.SuppliersEmailTB.Text.Trim()
        If Not String.IsNullOrWhiteSpace(email) AndAlso Not IsValidEmail(email) Then
            MessageBox.Show("❌ Invalid email format!\n📧 Example: supplier@example.com", "Invalid Email")
            form.SuppliersEmailTB.Focus()
            form.SuppliersEmailTB.BackColor = Color.LightPink
            Return
        End If

        Try
            OpenConn()

            ' 5️⃣ DUPLICATE CHECK
            Dim checkQuery As String = "SELECT COUNT(*) FROM Supplier WHERE Supplier_Name = @name OR Contact_Number = @contact"
            Dim checkCmd As New MySqlCommand(checkQuery, conn)
            checkCmd.Parameters.AddWithValue("@name", supplierName)
            checkCmd.Parameters.AddWithValue("@contact", contact)

            Dim existingCount = Convert.ToInt32(checkCmd.ExecuteScalar())
            If existingCount > 0 Then
                MessageBox.Show("❌ Supplier with this **name** or **contact** already exists!",
                           "Duplicate Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' ✅ ALL VALIDATIONS PASSED - INSERT
            Dim query As String = "INSERT INTO Supplier (Supplier_Name, Contact_Number, Address, Email) 
                              VALUES (@name, @contact, @address, @email)"

            cmd = New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@name", supplierName)
            cmd.Parameters.AddWithValue("@contact", contact)
            cmd.Parameters.AddWithValue("@address", address)
            cmd.Parameters.AddWithValue("@email", If(email = "", DBNull.Value, email))

            cmd.ExecuteNonQuery()

            ' 🎉 SUCCESS WITH DETAILS
            MessageBox.Show($"✅ Supplier added successfully!" & vbCrLf &
                       $"👤 {supplierName}" & vbCrLf &
                       $"📱 {contact}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ClearSupplierFields()  ' Reset colors too
            LoadSuppliers()
            form.LoadSuppliersToComboBox(form.AddMaterialSupplierCMB)
            form.LoadSuppliersToComboBox(form.UpdateMaterialSupplierCMB)

        Catch ex As Exception
            MessageBox.Show($"❌ Database Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            CloseConn()
        End Try
    End Sub

    Public Sub AddMaterial()
        If form.AddMaterialSupplierCMB.SelectedIndex = -1 Then
            MessageBox.Show("Please select a supplier")
            Return
        End If

        Try
            OpenConn()

            Dim query As String = "INSERT INTO Materials 
                              (Material_Name, Description, Quantity_in_Stock, Unit_of_Measure, Supplier_ID)
                              VALUES (@name, @desc, @qty, @unit, @supplierId)"

            cmd = New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@name", form.AddMaterialNameTB.Text)
            cmd.Parameters.AddWithValue("@desc", If(form.AddMaterialDescriptionTB.Text = "", DBNull.Value, form.AddMaterialDescriptionTB.Text))
            cmd.Parameters.AddWithValue("@qty", Convert.ToInt32(form.AddMaterialQuantityOnStockTB.Text))
            cmd.Parameters.AddWithValue("@unit", form.AddMaterialUnitOfMeasureTB.Text)
            cmd.Parameters.AddWithValue("@supplierId", form.AddMaterialSupplierCMB.SelectedValue)

            cmd.ExecuteNonQuery()
            MessageBox.Show("Material added successfully!")

            form.AddMaterialModalPanel.Visible = False
            ClearAddMaterialFields()
            LoadMaterials()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try
    End Sub

    Public Sub updateMaterials()
        If form.UpdateMaterialSupplierCMB.SelectedIndex = -1 Then
            MessageBox.Show("Please select a supplier")
            Return
        End If

        If form.UpdateMaterialIDTB Is Nothing OrElse form.UpdateMaterialIDTB.Text = "" Then
            MessageBox.Show("No material selected")
            Return
        End If

        Try
            OpenConn()

            Dim query As String = "UPDATE Materials SET 
                              Material_Name=@name,
                              Description=@desc,
                              Quantity_in_Stock=@qty,
                              Unit_of_Measure=@unit,
                              Supplier_ID=@supplierId
                              WHERE Material_ID=@id"

            cmd = New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@id", Convert.ToInt32(form.UpdateMaterialIDTB.Text))
            cmd.Parameters.AddWithValue("@name", form.UpdateMaterialNameTB.Text)
            cmd.Parameters.AddWithValue("@desc", If(form.UpdateMaterialDescriptionTB.Text = "", DBNull.Value, form.UpdateMaterialDescriptionTB.Text))
            cmd.Parameters.AddWithValue("@qty", Convert.ToInt32(form.UpdateMaterialQuantityOnStockTB.Text))
            cmd.Parameters.AddWithValue("@unit", form.UpdateMaterialUnitOfMeasureTB.Text)
            cmd.Parameters.AddWithValue("@supplierId", form.UpdateMaterialSupplierCMB.SelectedValue)

            cmd.ExecuteNonQuery()
            MessageBox.Show("Material updated successfully!")

            form.UpdateMaterialModalPanel.Visible = False
            LoadMaterials()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try
    End Sub

    Private Function IsValidPhoneNumber(phone As String) As Boolean

        Dim cleanPhone = System.Text.RegularExpressions.Regex.Replace(phone, "[^0-9+]", "")
        ' Check PH formats: 09xxxxxxxxx (11 digits) or +639xxxxxxxxx (12 chars)
        Return cleanPhone.Length = 11 AndAlso cleanPhone.StartsWith("09") OrElse
           cleanPhone.Length = 12 AndAlso cleanPhone.StartsWith("+63") AndAlso cleanPhone.Substring(3).StartsWith("9")
    End Function

    Private Function IsValidEmail(email As String) As Boolean
        Return System.Text.RegularExpressions.Regex.IsMatch(email,
        "^[^@\s]+@[^@\s]+\.[^@\s]+$")
    End Function

    Public Sub ClearAddMaterialFields()
        form.AddMaterialNameTB.Clear()
        form.AddMaterialDescriptionTB.Clear()
        form.AddMaterialQuantityOnStockTB.Clear()
        form.AddMaterialUnitOfMeasureTB.Clear()
        If form.AddMaterialSupplierCMB IsNot Nothing Then
            form.AddMaterialSupplierCMB.SelectedIndex = -1
        End If
    End Sub

    Public Sub ClearSupplierFields()
        form.SuppliersNameTB.Clear()
        form.SuppliersContactTB.Clear()
        form.SuppliersAddressTB.Clear()
        form.SuppliersEmailTB.Clear()
    End Sub

End Class
