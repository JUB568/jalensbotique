Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions

Public Class TailoringManager

    Private form As Form1
    Private cmd As MySqlCommand

    Private isCustomerOwnedMode As Boolean = False

    Public Sub New(frm As Form1)
        form = frm
    End Sub

    Public Sub LoadTransactions()
        Dim query As String = "
            SELECT 
                ts.Tailoring_Services_ID,
                ts.Type_of_Alteration,
                c.Full_Name,
                CASE WHEN ts.Is_Customer_Owned = 1 THEN '👕 CUSTOMER-OWNED' ELSE '👗 SHOP INVENTORY' END as Clothes_Source,
                CASE 
                    WHEN ts.Is_Customer_Owned = 1 THEN ts.Description_of_Work
                    ELSE CONCAT(cl.Clothes_Name, ' (', cl.Category, ' ', cl.Size, ')')
                END as Clothes_Details,
                ts.Service_Price,
                CONCAT(e.Full_Name, ' (', e.Position, ')') as Assigned_Employee,
                ts.Status,
                ts.Date_Requested
            FROM Tailoring_Services ts
            LEFT JOIN customer c ON ts.Customer_ID = c.Customer_ID
            LEFT JOIN Clothes cl ON ts.Clothes_ID = cl.Clothes_ID
            LEFT JOIN Employee e ON ts.Employee_ID = e.Employee_ID
            ORDER BY ts.Tailoring_Services_ID DESC"

        LoadToDGV(query, form.TailoringDGV)
        form.AutoSelectFirstRow(form.TailoringDGV)
    End Sub

    Public Sub LoadCustomersToCombo()
        Try
            ConnDB.OpenConn()
            Dim query As String = "SELECT Customer_ID, Full_Name FROM customer ORDER BY Full_Name"
            cmd = New MySqlCommand(query, ConnDB.conn)
            Dim adapter As New MySqlDataAdapter(cmd)
            Dim dt As New DataTable
            adapter.Fill(dt)

            form.TailoringCustomerCMB.DataSource = Nothing
            form.TailoringCustomerCMB.DataSource = dt
            form.TailoringCustomerCMB.DisplayMember = "Full_Name"
            form.TailoringCustomerCMB.ValueMember = "Customer_ID"
            form.TailoringCustomerCMB.SelectedIndex = -1
        Catch ex As Exception
            MessageBox.Show("Error loading customers: " & ex.Message)
        Finally
            ConnDB.CloseConn()
        End Try
    End Sub

    Public Sub LoadEmployeesToCombo()
        Try
            ConnDB.OpenConn()
            Dim query As String = "SELECT Employee_ID, CONCAT(Full_Name, ' - ', Position) as DisplayName FROM Employee WHERE Date_Hired IS NOT NULL ORDER BY Full_Name"
            cmd = New MySqlCommand(query, ConnDB.conn)
            Dim adapter As New MySqlDataAdapter(cmd)
            Dim dt As New DataTable
            adapter.Fill(dt)

            form.TailoringEmployeeCMB.DataSource = Nothing
            form.TailoringEmployeeCMB.DataSource = dt
            form.TailoringEmployeeCMB.DisplayMember = "DisplayName"
            form.TailoringEmployeeCMB.ValueMember = "Employee_ID"
            form.TailoringEmployeeCMB.SelectedIndex = -1
        Catch ex As Exception
            MessageBox.Show("Error loading employees: " & ex.Message)
        Finally
            ConnDB.CloseConn()
        End Try
    End Sub

    Public Sub LoadAvailableClothesToCombo()
        Try
            ConnDB.OpenConn()
            Dim query As String = "
                SELECT Clothes_ID, CONCAT(Clothes_Name, ' - ', Category, ' (', Size, ')') as DisplayName
                FROM Clothes WHERE Status IN ('Available', 'For Repair')
                ORDER BY Clothes_Name"

            cmd = New MySqlCommand(query, ConnDB.conn)
            Dim adapter As New MySqlDataAdapter(cmd)
            Dim dt As New DataTable
            adapter.Fill(dt)

            form.TailoringClothesCMB.DataSource = Nothing
            form.TailoringClothesCMB.DataSource = dt
            form.TailoringClothesCMB.DisplayMember = "DisplayName"
            form.TailoringClothesCMB.ValueMember = "Clothes_ID"
            form.TailoringClothesCMB.SelectedIndex = -1
        Catch ex As Exception
        Finally
            ConnDB.CloseConn()
        End Try
    End Sub

    Public Sub LoadCatalog()
        Dim catalogData = form.tailoringCatalogManager.Catalog.Select(Function(s) New With {
            .Catalog_ID = s.CatalogID,
            .Service_Name = s.ServiceName
        }).ToList()

        form.TailoringTypeCMB.DataSource = Nothing
        form.TailoringTypeCMB.DataSource = catalogData
        form.TailoringTypeCMB.DisplayMember = "Service_Name"
        form.TailoringTypeCMB.ValueMember = "Catalog_ID"
        form.TailoringTypeCMB.SelectedIndex = -1
    End Sub

    Public Sub CalculateSmartPrice()
        If form.TailoringTypeCMB.SelectedIndex < 0 Then
            form.TailoringPriceTB.Clear()
            Return
        End If

        Dim serviceName = form.TailoringTypeCMB.Text
        Dim service = form.tailoringCatalogManager.GetServiceByName(serviceName)
        If service Is Nothing Then
            form.TailoringPriceTB.Clear()
            Return
        End If

        form.TailoringPriceTB.Text = service.BasePrice.ToString("F2")
    End Sub

    Public Sub LoadRequiredMaterials()
        If form.TailoringTypeCMB.SelectedIndex < 0 Then
            form.TailoringMaterialsDGV.DataSource = Nothing
            Return
        End If

        Dim serviceName = form.TailoringTypeCMB.Text
        Dim service = form.tailoringCatalogManager.GetServiceByName(serviceName)
        If service Is Nothing Then Return

        Dim clothesDetails = GetSelectedClothesDetails()
        Dim requirements = form.tailoringCatalogManager.GetServiceRequirements(clothesDetails.Category, clothesDetails.FabricType, service.CatalogID)

        ' Try fallback if exact match not found
        If requirements.Count = 0 Then
            ShowNoMaterialsWarning()
            Return
        End If

        ' Show materials with stock status
        Dim materialsWithStock = New List(Of Object)
        For Each req In requirements
            Dim stockLevel As Integer = 0
            Try
                ConnDB.OpenConn()
                Dim stockQuery = "SELECT Quantity_in_Stock FROM Materials WHERE Material_ID = @id"
                cmd = New MySqlCommand(stockQuery, ConnDB.conn)
                cmd.Parameters.AddWithValue("@id", req.MaterialID)
                Dim result = cmd.ExecuteScalar()
                stockLevel = If(result IsNot Nothing, Convert.ToInt32(result), 0)
                ConnDB.CloseConn()
            Catch
                stockLevel = 0
            End Try

            Dim stockStatus As String
            If stockLevel >= req.DefaultQuantity Then
                stockStatus = $"✅ IN STOCK ({stockLevel})"
            ElseIf stockLevel > 0 Then
                stockStatus = $"⚠️ LOW ({stockLevel}/{req.DefaultQuantity})"
            Else
                stockStatus = "❌ OUT OF STOCK"
            End If

            materialsWithStock.Add(New With {
                .Material_ID = req.MaterialID,
                .Material_Name = req.MaterialName,
                .Default_Quantity = req.DefaultQuantity,
                .Unit_Measure = req.UnitMeasure,
                .Stock_Available = stockLevel,
                .Stock_Status = stockStatus,
                .Category_Match = clothesDetails.Category,
                .Fabric_Match = clothesDetails.FabricType
            })
        Next

        form.TailoringMaterialsDGV.Columns.Clear()
        form.TailoringMaterialsDGV.DataSource = materialsWithStock

        ' Auto-size columns
        form.TailoringMaterialsDGV.AutoResizeColumns()
        If form.TailoringMaterialsDGV.Columns.Contains("Material_ID") Then
            form.TailoringMaterialsDGV.Columns("Material_ID").Visible = False
        End If
    End Sub

    Public Sub CreateRequest()
        ' Validate service compatibility first
        Dim clothesDetails = GetSelectedClothesDetails()
        Dim serviceName = form.TailoringTypeCMB.Text
        Dim tailoringService = form.tailoringCatalogManager.GetServiceByName(serviceName)

        ' Incompatible services
        Dim incompatibleServices = {
            ("Pants", "Shorten Sleeves"), ("Pants", "Button Replacement"),
            ("Top", "Zipper Replacement"), ("Dress", "Button Replacement"),
            ("Suit", "Zipper Replacement"), ("Dress", "Zipper Replacement"),
            ("Top", "Hemming Pants")
        }
        For Each invalid In incompatibleServices
            If clothesDetails.Category = invalid.Item1 AndAlso serviceName.Contains(invalid.Item2) Then
                MessageBox.Show($"❌ '{serviceName}' not available for {clothesDetails.Category}!", "Service Incompatible")
                Return
            End If
        Next

        If form.TailoringCustomerCMB.SelectedIndex = -1 OrElse String.IsNullOrEmpty(form.TailoringTypeCMB.Text) Then
            MessageBox.Show("⚠️ Select Customer + Service Type")
            Return
        End If

        If form.TailoringEmployeeCMB.SelectedIndex = -1 Then
            MessageBox.Show("⚠️ Please assign an Employee")
            Return
        End If

        If String.IsNullOrEmpty(form.TailoringPriceTB.Text) Then
            MessageBox.Show("⚠️ Select Service Type first to auto-calculate price")
            Return
        End If

        Try
            ConnDB.OpenConn()

            Dim insertQuery = "
                INSERT INTO Tailoring_Services 
                (Type_of_Alteration, Description_of_Work, Service_Price, Date_Requested, Status, 
                 Customer_ID, Clothes_ID, Is_Customer_Owned, Employee_ID)  
                VALUES (@type, @desc, @price, NOW(), 'Pending', @customerId, @clothesId, @isOwned, @employeeId)"

            cmd = New MySqlCommand(insertQuery, ConnDB.conn)
            cmd.Parameters.AddWithValue("@type", form.TailoringTypeCMB.Text)
            cmd.Parameters.AddWithValue("@desc", If(isCustomerOwnedMode, form.TailoringCustomerClothesTB.Text, form.TailoringDescriptionTB.Text))
            cmd.Parameters.AddWithValue("@price", Convert.ToDecimal(form.TailoringPriceTB.Text))
            cmd.Parameters.AddWithValue("@customerId", form.TailoringCustomerCMB.SelectedValue)
            cmd.Parameters.AddWithValue("@clothesId", If(isCustomerOwnedMode, DBNull.Value, form.TailoringClothesCMB.SelectedValue))
            cmd.Parameters.AddWithValue("@isOwned", If(isCustomerOwnedMode, 1, 0))
            cmd.Parameters.AddWithValue("@employeeId", form.TailoringEmployeeCMB.SelectedValue)

            cmd.ExecuteNonQuery()

            ' Update clothes status if shop inventory
            If Not isCustomerOwnedMode AndAlso form.TailoringClothesCMB.SelectedValue IsNot Nothing Then
                Dim updateClothes = "UPDATE Clothes SET Status = 'In Tailoring' WHERE Clothes_ID = @id"
                cmd = New MySqlCommand(updateClothes, ConnDB.conn)
                cmd.Parameters.AddWithValue("@id", form.TailoringClothesCMB.SelectedValue)
                cmd.ExecuteNonQuery()
            End If

            ' Auto-deduct materials from stock
            If tailoringService IsNot Nothing Then
                Dim requirements = form.tailoringCatalogManager.GetServiceRequirements(clothesDetails.Category, clothesDetails.FabricType, tailoringService.CatalogID)
                Dim serviceId As Integer = Convert.ToInt32(cmd.LastInsertedId)

                For Each req In requirements
                    Try
                        Dim deductQuery = "UPDATE Materials SET Quantity_in_Stock = GREATEST(0, Quantity_in_Stock - @qty) WHERE Material_ID = @materialId"
                        cmd = New MySqlCommand(deductQuery, ConnDB.conn)
                        cmd.Parameters.AddWithValue("@qty", req.DefaultQuantity)
                        cmd.Parameters.AddWithValue("@materialId", req.MaterialID)
                        cmd.ExecuteNonQuery()

                        Dim recordQuery = "INSERT INTO Required_Materials (Tailoring_Services_ID, Material_ID, Quantity_Used, Unit_Measure_Used) VALUES (@serviceId, @materialId, @qty, @unit)"
                        cmd = New MySqlCommand(recordQuery, ConnDB.conn)
                        cmd.Parameters.AddWithValue("@serviceId", serviceId)
                        cmd.Parameters.AddWithValue("@materialId", req.MaterialID)
                        cmd.Parameters.AddWithValue("@qty", req.DefaultQuantity)
                        cmd.Parameters.AddWithValue("@unit", req.UnitMeasure)
                        cmd.ExecuteNonQuery()
                    Catch deductEx As Exception
                        Debug.WriteLine($"⚠️ Failed to deduct material {req.MaterialID}: {deductEx.Message}")
                    End Try
                Next
            End If

            MessageBox.Show($"✅ Tailoring request created!" & vbCrLf &
                           $"💰 ₱" & form.TailoringPriceTB.Text & vbCrLf &
                           $"👷 Assigned: " & form.TailoringEmployeeCMB.Text & vbCrLf &
                           $"📦 " & If(isCustomerOwnedMode, "👕 CUSTOMER-OWNED", "👗 SHOP INVENTORY"),
                           "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ClearFields()
            LoadTransactions()
            form.clothesManager.LoadClothes()  ' refresh clothes list

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            ConnDB.CloseConn()
        End Try
    End Sub

    Public Sub UpdateStatus()
        If String.IsNullOrWhiteSpace(form.TailoringServiceIDTB.Text) Then
            MessageBox.Show("❌ Select a tailoring transaction first!", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If form.TailoringStatusCMB.SelectedIndex = -1 OrElse form.TailoringStatusCMB.SelectedItem Is Nothing Then
            MessageBox.Show("❌ Please select a status from the dropdown!", "No Status Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim serviceId As Integer = Convert.ToInt32(form.TailoringServiceIDTB.Text.Trim())
        Dim newStatus As String = form.TailoringStatusCMB.SelectedItem.ToString().Trim()

        Dim transaction As MySqlTransaction = Nothing
        Try
            ConnDB.OpenConn()
            transaction = ConnDB.conn.BeginTransaction(IsolationLevel.ReadCommitted)

            ' 1. Update service status
            Dim updateServiceQuery As String = "UPDATE Tailoring_Services SET Status = @status, Date_Completed = NOW() WHERE Tailoring_Services_ID = @id"
            Using cmdService = New MySqlCommand(updateServiceQuery, ConnDB.conn, transaction)
                cmdService.Parameters.AddWithValue("@status", newStatus)
                cmdService.Parameters.AddWithValue("@id", serviceId)
                cmdService.ExecuteNonQuery()
            End Using

            ' 2. Get clothes details
            Dim clothesId As Integer? = Nothing
            Dim isCustomerOwned As Boolean = False
            Dim serviceTypeName As String = ""
            Dim clothesCategory As String = "Top"
            Dim clothesFabricType As String = "Cotton"

            Dim getDetailsQuery As String = "
                SELECT ts.Clothes_ID, ts.Is_Customer_Owned, ts.Type_of_Alteration, cl.Category, cl.Fabric_Type
                FROM Tailoring_Services ts
                LEFT JOIN Clothes cl ON ts.Clothes_ID = cl.Clothes_ID
                WHERE ts.Tailoring_Services_ID = @id"

            Using cmdDetails = New MySqlCommand(getDetailsQuery, ConnDB.conn, transaction)
                cmdDetails.Parameters.AddWithValue("@id", serviceId)
                Using reader = cmdDetails.ExecuteReader()
                    If reader.Read() Then
                        If Not IsDBNull(reader("Clothes_ID")) Then
                            clothesId = Convert.ToInt32(reader("Clothes_ID"))
                        End If
                        isCustomerOwned = Convert.ToBoolean(reader("Is_Customer_Owned"))
                        serviceTypeName = reader("Type_of_Alteration").ToString()
                        If Not isCustomerOwned AndAlso clothesId.HasValue Then
                            clothesCategory = If(IsDBNull(reader("Category")), "Top", reader("Category").ToString())
                            clothesFabricType = If(IsDBNull(reader("Fabric_Type")), "Cotton", reader("Fabric_Type").ToString())
                        End If
                    End If
                End Using
            End Using

            ' 3. Handle status-specific logic
            Dim clothesUpdated As Integer = 0
            Dim materialsRefunded As Integer = 0

            Select Case newStatus.ToUpperInvariant()
                Case "COMPLETED"
                    If clothesId.HasValue AndAlso Not isCustomerOwned Then
                        Dim updateClothesQuery = "UPDATE Clothes SET Status = 'Available', Clothes_Condition = 'Good' WHERE Clothes_ID = @clothesId"
                        Using cmdClothes = New MySqlCommand(updateClothesQuery, ConnDB.conn, transaction)
                            cmdClothes.Parameters.AddWithValue("@clothesId", clothesId.Value)
                            clothesUpdated = cmdClothes.ExecuteNonQuery()
                        End Using
                    End If

                Case "CANCELLED"
                    If clothesId.HasValue AndAlso Not isCustomerOwned Then
                        Dim updateClothesQuery = "UPDATE Clothes SET Status = 'Available' WHERE Clothes_ID = @clothesId"
                        Using cmdClothes = New MySqlCommand(updateClothesQuery, ConnDB.conn, transaction)
                            cmdClothes.Parameters.AddWithValue("@clothesId", clothesId.Value)
                            clothesUpdated = cmdClothes.ExecuteNonQuery()
                        End Using
                    End If

                    ' Refund materials
                    Dim service = form.tailoringCatalogManager.GetServiceByName(serviceTypeName)
                    If service IsNot Nothing Then
                        Dim refundRequirements = form.tailoringCatalogManager.GetServiceRequirements(clothesCategory, clothesFabricType, service.CatalogID)
                        For Each req In refundRequirements
                            Dim refundQuery = "UPDATE Materials SET Quantity_in_Stock = Quantity_in_Stock + @qty WHERE Material_ID = @materialId"
                            Using cmdRefund = New MySqlCommand(refundQuery, ConnDB.conn, transaction)
                                cmdRefund.Parameters.AddWithValue("@qty", req.DefaultQuantity)
                                cmdRefund.Parameters.AddWithValue("@materialId", req.MaterialID)
                                materialsRefunded += cmdRefund.ExecuteNonQuery()
                            End Using
                        Next
                    End If
            End Select

            transaction.Commit()

            Dim successMsg As New List(Of String) From {
                $"✅ Service #{serviceId} → '{newStatus}'",
                $"📅 {DateTime.Now:MMM dd, yyyy HH:mm}"
            }
            If clothesUpdated > 0 Then successMsg.Insert(0, "👗 Clothes returned to Available inventory!")
            If materialsRefunded > 0 Then successMsg.Insert(0, $"💰 {materialsRefunded} material units refunded!")

            MessageBox.Show(String.Join(vbCrLf, successMsg), "✅ Update Successful!", MessageBoxButtons.OK, MessageBoxIcon.Information)

            LoadTransactions()
            form.clothesManager.LoadClothes()
            LoadAvailableClothesToCombo()   ' refresh combo

        Catch ex As Exception
            transaction?.Rollback()
            MessageBox.Show($"❌ Update failed: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ConnDB.CloseConn()
        End Try
    End Sub

    Public Sub ClearFields()
        form.TailoringCustomerCMB.SelectedIndex = -1
        form.TailoringClothesCMB.SelectedIndex = -1
        form.TailoringEmployeeCMB.SelectedIndex = -1
        form.TailoringTypeCMB.SelectedIndex = -1
        form.TailoringCustomerClothesTB.Clear()
        form.TailoringDescriptionTB.Clear()
        form.TailoringPriceTB.Clear()
        form.TailoringMaterialsDGV.DataSource = Nothing
        form.TailoringServiceIDTB.Clear()
        ResetUI()
    End Sub

    Public Sub ResetUI()
        isCustomerOwnedMode = False
        form.CustomerOwnedToggleBTN.Text = "👕 CUSTOMER-OWNED"
        form.CustomerOwnedToggleBTN.BackColor = Color.DodgerBlue
        form.TailoringClothesCMB.Enabled = True
        form.TailoringCustomerClothesTB.Enabled = False
        form.TailoringClothesLabel.Text = "Shop Inventory Clothes:"
        form.TailoringDescriptionTB.Enabled = True
    End Sub

    Public Sub ToggleCustomerOwnedMode()
        isCustomerOwnedMode = Not isCustomerOwnedMode

        If isCustomerOwnedMode Then
            form.CustomerOwnedToggleBTN.Text = "👗 SHOP INVENTORY"
            form.CustomerOwnedToggleBTN.BackColor = Color.Orange
            form.TailoringClothesCMB.Enabled = False
            form.TailoringCustomerClothesTB.Enabled = True
            form.TailoringClothesLabel.Text = "Customer Clothes Description:"
            form.TailoringDescriptionTB.Enabled = False
        Else
            form.CustomerOwnedToggleBTN.Text = "👕 CUSTOMER-OWNED"
            form.CustomerOwnedToggleBTN.BackColor = Color.DodgerBlue
            form.TailoringClothesCMB.Enabled = True
            form.TailoringCustomerClothesTB.Enabled = False
            form.TailoringClothesLabel.Text = "Shop Inventory Clothes:"
            form.TailoringDescriptionTB.Enabled = True
        End If
    End Sub

    Public Sub LoadEditFields(row As DataGridViewRow)
        form.TailoringServiceIDTB.Text = form.GetCellValue(row.Cells("Tailoring_Services_ID"))
    End Sub

    ' ====================================================
    ' PRIVATE HELPERS
    ' ====================================================

    Private Function GetSelectedClothesDetails() As (Category As String, FabricType As String)
        If isCustomerOwnedMode Then
            Dim desc = form.TailoringCustomerClothesTB.Text.ToLower()
            Dim category = "Top"
            If desc.Contains("pants") Or desc.Contains("slacks") Or desc.Contains("jeans") Then
                category = "Pants"
            ElseIf desc.Contains("dress") Or desc.Contains("terno") Then
                category = "Dress"
            ElseIf desc.Contains("suit") Or desc.Contains("uniform") Then
                category = "Suit"
            End If
            Return (category, "Cotton")
        End If

        If form.TailoringClothesCMB.SelectedValue IsNot Nothing Then
            Try
                ConnDB.OpenConn()
                Dim query = "SELECT Category, Fabric_Type FROM Clothes WHERE Clothes_ID = @id"
                cmd = New MySqlCommand(query, ConnDB.conn)
                cmd.Parameters.AddWithValue("@id", form.TailoringClothesCMB.SelectedValue)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Dim category = reader("Category").ToString()
                        Dim fabric = If(IsDBNull(reader("Fabric_Type")), "Cotton", reader("Fabric_Type").ToString())
                        ConnDB.CloseConn()
                        Return (category, fabric)
                    End If
                End Using
                ConnDB.CloseConn()
            Catch
                ' fallback
            End Try
        End If
        Return ("Top", "Cotton")
    End Function

    Private Sub ShowNoMaterialsWarning()
        form.TailoringMaterialsDGV.DataSource = Nothing
        form.TailoringMaterialsDGV.Columns.Clear()
        form.TailoringMaterialsDGV.Columns.Add("MaterialID", "ID")
        form.TailoringMaterialsDGV.Columns.Add("MaterialName", "Material")
        form.TailoringMaterialsDGV.Columns.Add("Quantity", "Qty")
        form.TailoringMaterialsDGV.Columns.Add("Unit", "Unit")
        form.TailoringMaterialsDGV.Columns.Add("Stock", "Stock Status")
        form.TailoringMaterialsDGV.Columns.Add("Match", "Match")
        form.TailoringMaterialsDGV.Rows.Add("", "⚠️ No materials defined for this service + cloth combination", "", "", "CHECK SERVICE COMPATIBILITY", "")
    End Sub

    Private Sub LoadToDGV(query As String, dgv As DataGridView)
        ConnDB.LoadToDGV(query, dgv)
    End Sub

End Class