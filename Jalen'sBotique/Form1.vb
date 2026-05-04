Imports MySql.Data.MySqlClient


Public Class Form1

    Public Class TailoringService
        Public Property CatalogID As Integer
        Public Property ServiceName As String
        Public Property Description As String
        Public Property BasePrice As Decimal
        Public Property EstimatedTime As String
    End Class

    Public Class ServiceRequirement
        Public Property CatalogID As Integer
        Public Property MaterialName As String
        Public Property DefaultQuantity As Integer
        Public Property UnitMeasure As String

        Public Property MaterialID As Integer
    End Class

    Private ReadOnly TailoringCatalog As New List(Of TailoringService) From {
    New TailoringService With {.CatalogID = 1, .ServiceName = "Hemming Pants", .Description = "Shorten pants legs", .BasePrice = 150D, .EstimatedTime = "1 day"},
    New TailoringService With {.CatalogID = 2, .ServiceName = "Zipper Replacement", .Description = "Replace broken zipper", .BasePrice = 250D, .EstimatedTime = "2 days"},
    New TailoringService With {.CatalogID = 3, .ServiceName = "Button Replacement", .Description = "Replace missing buttons", .BasePrice = 80D, .EstimatedTime = "1 hour"},
    New TailoringService With {.CatalogID = 4, .ServiceName = "Taking In Waist", .Description = "Reduce waist size", .BasePrice = 200D, .EstimatedTime = "2 days"},
    New TailoringService With {.CatalogID = 5, .ServiceName = "Tapering Legs", .Description = "Slim fit leg taper", .BasePrice = 300D, .EstimatedTime = "3 days"},
    New TailoringService With {.CatalogID = 6, .ServiceName = "Shorten Sleeves", .Description = "Shorten shirt/jacket sleeves", .BasePrice = 180D, .EstimatedTime = "1 day"},
    New TailoringService With {.CatalogID = 7, .ServiceName = "Lengthen Hem", .Description = "Add fabric to hem", .BasePrice = 220D, .EstimatedTime = "2 days"}
}

    Private ServiceRequirements As New Dictionary(Of Integer, List(Of ServiceRequirement))

    Private Sub InitializeServiceRequirements()
        ServiceRequirements = New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
        {1, New List(Of ServiceRequirement) From {  ' Hemming Pants (ID 1)
            New ServiceRequirement With {.CatalogID = 1, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 2, .UnitMeasure = "rolls"},
            New ServiceRequirement With {.CatalogID = 1, .MaterialID = 5, .MaterialName = "Fabric Patch Cotton", .DefaultQuantity = 1, .UnitMeasure = "pieces"}
        }},
        {2, New List(Of ServiceRequirement) From {  ' Zipper Replacement (ID 3)
            New ServiceRequirement With {.CatalogID = 2, .MaterialID = 3, .MaterialName = "Zipper 7 Metal", .DefaultQuantity = 1, .UnitMeasure = "pieces"},
            New ServiceRequirement With {.CatalogID = 2, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 1, .UnitMeasure = "rolls"}
        }},
        {3, New List(Of ServiceRequirement) From {  ' Button Replacement (ID 4)
            New ServiceRequirement With {.CatalogID = 3, .MaterialID = 4, .MaterialName = "Buttons Plastic 15mm", .DefaultQuantity = 4, .UnitMeasure = "pieces"},
            New ServiceRequirement With {.CatalogID = 3, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 1, .UnitMeasure = "rolls"}
        }},
        {4, New List(Of ServiceRequirement) From {  ' Taking In Waist
            New ServiceRequirement With {.CatalogID = 4, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 3, .UnitMeasure = "rolls"}
        }},
        {5, New List(Of ServiceRequirement) From {  ' Tapering Legs
            New ServiceRequirement With {.CatalogID = 5, .MaterialID = 2, .MaterialName = "Polyester Thread Black", .DefaultQuantity = 4, .UnitMeasure = "rolls"},
            New ServiceRequirement With {.CatalogID = 5, .MaterialID = 6, .MaterialName = "Denim Patch Material", .DefaultQuantity = 2, .UnitMeasure = "pieces"}
        }},
        {6, New List(Of ServiceRequirement) From {  ' Shorten Sleeves
            New ServiceRequirement With {.CatalogID = 6, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 2, .UnitMeasure = "rolls"}
        }},
        {7, New List(Of ServiceRequirement) From {  ' Lengthen Hem
            New ServiceRequirement With {.CatalogID = 7, .MaterialID = 5, .MaterialName = "Fabric Patch Cotton", .DefaultQuantity = 3, .UnitMeasure = "pieces"},
            New ServiceRequirement With {.CatalogID = 7, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 2, .UnitMeasure = "rolls"}
        }}
    }
    End Sub



    Private ReadOnly FabricPricingRules As New Dictionary(Of String, Dictionary(Of String, Decimal)) From {
    {"Pantalon Hemming", New Dictionary(Of String, Decimal) From {{"Piña", 1.8D}, {"Cotton", 1D}, {"Denim", 1.2D}, {"Polo", 1.1D}}},
    {"Zipper Palit", New Dictionary(Of String, Decimal) From {{"Cotton", 1D}, {"Piña", 2D}, {"Denim", 1.3D}, {"Katsa", 1.5D}}},
    {"Button Palit", New Dictionary(Of String, Decimal) From {{"Piña", 1.5D}, {"Cotton", 1D}, {"Polo", 1.1D}}},
    {"Baywang Tanggal", New Dictionary(Of String, Decimal) From {{"Piña", 2D}, {"Cotton", 1D}, {"Denim", 1.4D}}},
    {"Pantalon Taper", New Dictionary(Of String, Decimal) From {{"Denim", 1.5D}, {"Cotton", 1D}, {"Piña", 2D}, {"Polo", 1.3D}}},
    {"Manggas Shorten", New Dictionary(Of String, Decimal) From {{"Piña", 1.6D}, {"Cotton", 1D}, {"Polo", 1.2D}}},
    {"Basta Lengthen", New Dictionary(Of String, Decimal) From {{"Piña", 2D}, {"Cotton", 1D}, {"Katsa", 1.4D}}},
    {"Barong Waist In", New Dictionary(Of String, Decimal) From {{"Piña", 2.2D}, {"Polyester", 1D}}},
    {"Terno Hemming", New Dictionary(Of String, Decimal) From {{"Silk", 2.5D}, {"Piña", 2D}, {"Cotton", 1.5D}}}
}

    ' Tailoring mode tracking
    Private isCustomerOwnedMode As Boolean = False

    ' Dragging variables
    Private isDragging As Boolean = False
    Private dragOffset As Point

    Public userRole As String = "admin" ' 👉 set default for testing

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConnDB.TestConnection()
        UpdateConnectionStatus()

        isArchiveMode = False
        UpdateArchiveModeUI()

        LoadClothes()
        LoadMaterials()
        LoadSuppliers()

        InitializeServiceRequirements()
        ClotheAddModalPanel.Visible = False
        ClotheEditModalPanel.Visible = False
        AddMaterialModalPanel.Visible = False
        UpdateMaterialModalPanel.Visible = False
        AddCustomerModalPanel.Visible = False
        UpdateCustomerModalPanel.Visible = False
        ' =========================
        ' ❌ DISABLED LOGIN SYSTEM
        ' =========================

        'MainPanel.Visible = False
        'SidebarPanel.Visible = False

        'overlayPanel.Parent = Me
        'overlayPanel.Dock = DockStyle.Fill
        'overlayPanel.Visible = True
        'overlayPanel.BringToFront()

        'LogInPanel.Visible = True
        'LogInPanel.BringToFront()
        'CenterLoginPanel()

        ' =========================
        ' ✅ DEBUG MODE (SHOW EVERYTHING)
        ' =========================
        MainPanel.Visible = True
        SidebarPanel.Visible = True

        ' Form settings
        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = FormBorderStyle.Sizable

        ' Load UI based on role
        SetupUI()

    End Sub

    ' =========================
    ' ❌ LOGIN DISABLED
    ' =========================
    'Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
    'End Sub

    ' =========================
    ' ROLE-BASED UI
    ' =========================
    Private Sub SetupUI()

        ' Default visibility
        UsersBTN.Visible = True
        ReportsBTN.Visible = True
        MaterialBTN.Visible = True

        If userRole = "admin" Then

            ShowPanel(AdminDashboardPanel)

        ElseIf userRole = "employee" Then

            UsersBTN.Visible = False
            ReportsBTN.Visible = False

            ShowPanel(EmployeeDashboardPanel)

        End If

    End Sub

    ' =========================
    ' PANEL SWITCHING
    ' =========================
    Private Sub ShowPanel(panel As Panel)

        For Each ctrl As Control In MainPanel.Controls
            If TypeOf ctrl Is Panel Then
                ctrl.Visible = False
            End If
        Next

        panel.Visible = True
        panel.BringToFront()

    End Sub

    ' =========================
    ' BUTTON EVENTS
    ' =========================
    Private Sub DashboardBTN_Click(sender As Object, e As EventArgs) Handles DashboardBTN.Click

        If userRole = "admin" Then
            ShowPanel(AdminDashboardPanel)
        Else
            ShowPanel(EmployeeDashboardPanel)
        End If

    End Sub

    Private Sub ClothesBTN_Click(sender As Object, e As EventArgs) Handles ClothesBTN.Click
        ShowPanel(ClothesPanel)
    End Sub

    Private Sub CustomerBTN_Click(sender As Object, e As EventArgs) Handles CustomerBTN.Click
        ShowPanel(CustomerPanel)
    End Sub

    Private Sub RentBTN_Click(sender As Object, e As EventArgs) Handles RentBTN.Click
        ShowPanel(RentPanel)
    End Sub

    Private Sub TailoringBTN_Click(sender As Object, e As EventArgs) Handles TailoringBTN.Click
        ShowPanel(TailoringPanel)
    End Sub

    Private Sub MaterialBTN_Click(sender As Object, e As EventArgs) Handles MaterialBTN.Click
        ShowPanel(MaterialsPanel)
    End Sub

    Private Sub ReportsBTN_Click(sender As Object, e As EventArgs) Handles ReportsBTN.Click

        If userRole <> "admin" Then
            MessageBox.Show("Access Denied!")
            Exit Sub
        End If

        ShowPanel(ReportsPanel)

    End Sub

    Private Sub UsersBTN_Click(sender As Object, e As EventArgs)

        If userRole <> "admin" Then
            MessageBox.Show("Access Denied!")
            Exit Sub
        End If

        ShowPanel(UserPanel)

    End Sub


    ' =========================
    ' CLOTHES INVENTORY PANEL
    ' =========================

    Private Sub LoadClothes()
        LoadActiveClothes()
    End Sub

    Private Sub LoadActiveClothes()
        Dim query As String = "
        SELECT Clothes_ID,
            Clothes_Name,
            Category,
            Size,
            Fabric_Type,
            Color,
            Clothes_Condition,
            Status,
            Date_Added FROM Clothes 
        WHERE Status != 'Archived' 
        ORDER BY Clothes_ID DESC"
        LoadToDGV(query, ClothesDGV)
        AutoSelectFirstRow()
    End Sub

    Private Sub LoadArchiveClothes()
        Dim query As String = "
        SELECT 
            Clothes_ID,
            Clothes_Name,
            Category,
            Size,
            Fabric_Type,
            Color,
            Clothes_Condition,
            Status,
            Archived_Date
        FROM Clothes 
        WHERE Status = 'Archived'
        ORDER BY Archived_Date DESC, Clothes_ID DESC"
        LoadToDGV(query, ArchiveClothesDGV)
        AutoSelectFirstRow()
    End Sub

    Private Sub ClotheSearchTB_TextChanged(sender As Object, e As EventArgs) Handles ClotheSearchTB.TextChanged
        Dim searchText As String = ClotheSearchTB.Text.Trim()

        If isArchiveMode Then
            Dim query As String = "
            SELECT Clothes_ID, Clothes_Name, Category, Size, Fabric_Type, 
                   Color, Clothes_Condition, Status, Archived_Date
            FROM Clothes 
            WHERE Status = 'Archived'
              AND (Clothes_Name LIKE '%" & searchText & "%'
                OR Category LIKE '%" & searchText & "%'
                OR Color LIKE '%" & searchText & "%')
            ORDER BY Archived_Date DESC"
            LoadToDGV(query, ArchiveClothesDGV)
        Else
            Dim query As String = "
            SELECT * FROM Clothes 
            WHERE Status != 'Archived'
              AND (Clothes_Name LIKE '%" & searchText & "%'
                OR Category LIKE '%" & searchText & "%'
                OR Color LIKE '%" & searchText & "%')
            ORDER BY Clothes_ID DESC"
            LoadToDGV(query, ClothesDGV)
        End If
        AutoSelectFirstRow()
    End Sub

    Private Sub ClothesEditBTN_Click(sender As Object, e As EventArgs) Handles ClothesEditBTN.Click

        ' check if may selected item
        If ClothesEditModalClothesIDTB.Text = "" Then
            MessageBox.Show("Please select item first")
            Return
        End If

        ' show modal
        ClotheEditModalPanel.Visible = True
        ClotheEditModalPanel.BringToFront()

    End Sub

    Private Sub ClothesAddBTN_Click(sender As Object, e As EventArgs) Handles ClothesAddBTN.Click
        ClotheAddModalPanel.Visible = True
        ClotheAddModalPanel.BringToFront()
    End Sub
    Private Sub ClothesAddModalCancelBTN_Click(sender As Object, e As EventArgs) Handles ClothesAddModalCancelBTN.Click
        ClotheAddModalPanel.Visible = False
    End Sub

    Private Sub ClothesAddModalConfirmBTN_Click(sender As Object, e As EventArgs) Handles ClothesAddModalConfirmBTN.Click

        Try
            OpenConn()

            Dim query As String = "INSERT INTO Clothes 
        (Clothes_Name, Category, Size, Fabric_Type, Color, Clothes_Condition, Status)
        VALUES (@name, @cat, @size, @fabric, @color, 'Good', 'Available')"

            cmd = New MySqlCommand(query, conn)

            cmd.Parameters.AddWithValue("@name", ClothesAddModalClothesNameTB.Text)
            cmd.Parameters.AddWithValue("@cat", ClothesAddModalCategoryTB.Text)
            cmd.Parameters.AddWithValue("@size", ClothesAddModalSizeTB.Text)
            cmd.Parameters.AddWithValue("@fabric", ClothesAddModalFabricTypeTB.Text)
            cmd.Parameters.AddWithValue("@color", ClothesAddModalColorTB.Text)

            cmd.ExecuteNonQuery()

            MessageBox.Show("Clothes added successfully!")

            ClotheAddModalPanel.Visible = False
            ClearAddFields()
            LoadActiveClothes()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try

    End Sub

    Private Sub ClothesDGVs_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles ClothesDGV.CellClick, ArchiveClothesDGV.CellClick
        Dim dgv As DataGridView = DirectCast(sender, DataGridView)

        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = dgv.Rows(e.RowIndex)
            LoadEditFieldsFromRow(row)
        End If
    End Sub

    Private Sub AutoSelectFirstRow()
        Dim currentDGV As DataGridView = If(isArchiveMode, ArchiveClothesDGV, ClothesDGV)

        If currentDGV IsNot Nothing AndAlso currentDGV.Rows.Count > 0 Then
            ' Select first row
            currentDGV.ClearSelection()
            currentDGV.CurrentCell = currentDGV.Rows(0).Cells(0)
            currentDGV.Rows(0).Selected = True

            ' Load fields automatically
            LoadEditFieldsFromRow(currentDGV.Rows(0))
        Else
            ClearEditFields()
        End If
    End Sub

    Private Sub LoadEditFieldsFromRow(row As DataGridViewRow)
        Try
            With row
                ClothesEditModalClothesIDTB.Text = GetCellValue(.Cells("Clothes_ID"))
                ClothesEditModalClothesNameTB.Text = GetCellValue(.Cells("Clothes_Name"))
                ClothesEditModalCategoryTB.Text = GetCellValue(.Cells("Category"))
                ClothesEditModalSizeTB.Text = GetCellValue(.Cells("Size"))
                ClothesEditModalFabricTypeTB.Text = GetCellValue(.Cells("Fabric_Type"))
                ClothesEditModalColorTB.Text = GetCellValue(.Cells("Color"))
                ClothesEditModalConditionTB.Text = GetCellValue(.Cells("Clothes_Condition"))
            End With

        Catch ex As Exception
            Debug.WriteLine("LoadEditFieldsFromRow Error: " & ex.Message)
        End Try
    End Sub

    ' Helper function
    Private Function GetCellValue(cell As DataGridViewCell) As String
        If cell Is Nothing OrElse cell.Value Is Nothing Then
            Return ""
        End If
        Return cell.Value.ToString()
    End Function



    Private Sub ClothesEditModalConfirmBTN_Click(sender As Object, e As EventArgs) Handles ClothesEditModalConfirmBTN.Click

        Try
            OpenConn()

            Dim query As String = "UPDATE Clothes SET 
            Clothes_Name=@name,
            Category=@cat,
            Size=@size,
            Fabric_Type=@fabric,
            Color=@color,
            Clothes_Condition=@condition
            WHERE Clothes_ID=@id"

            cmd = New MySqlCommand(query, conn)

            cmd.Parameters.AddWithValue("@id", ClothesEditModalClothesIDTB.Text)
            cmd.Parameters.AddWithValue("@name", ClothesEditModalClothesNameTB.Text)
            cmd.Parameters.AddWithValue("@cat", ClothesEditModalCategoryTB.Text)
            cmd.Parameters.AddWithValue("@size", ClothesEditModalSizeTB.Text)
            cmd.Parameters.AddWithValue("@fabric", ClothesEditModalFabricTypeTB.Text)
            cmd.Parameters.AddWithValue("@color", ClothesEditModalColorTB.Text)
            cmd.Parameters.AddWithValue("@condition", ClothesEditModalConditionTB.Text)

            cmd.ExecuteNonQuery()

            MessageBox.Show("Updated successfully!")

            ClotheEditModalPanel.Visible = False
            LoadClothes()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try

    End Sub

    ' Archive (Soft Delete) - Changes Status to 'Archived'
    Private Sub ClothesDeleteBTN_Click(sender As Object, e As EventArgs) Handles ClothesDeleteBTN.Click
        If ConnDB.IsOfflineMode Then
            MessageBox.Show("Cannot modify while offline!", "Offline Mode")
            Return
        End If

        If ClothesEditModalClothesIDTB.Text = "" Then
            MessageBox.Show("Select item first!")
            Return
        End If

        Dim actionText As String = If(isArchiveMode, "RESTORE to Active", "ARCHIVE to Storage")
        Dim newStatus As String = If(isArchiveMode, "Available", "Archived")

        If MessageBox.Show($"📦 {actionText}?" & vbCrLf & $"Status will be: {newStatus}",
                          If(isArchiveMode, "Restore Item", "Archive Item"),
                          MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            Try
                OpenConn()

                Dim query As String = "
                    UPDATE Clothes SET 
                        Status = @newStatus,
                        Archived_Date = " & If(newStatus = "Archived", "NOW()", "NULL") & "
                    WHERE Clothes_ID = @id"

                cmd = New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", Convert.ToInt32(ClothesEditModalClothesIDTB.Text))
                cmd.Parameters.AddWithValue("@newStatus", newStatus)

                cmd.ExecuteNonQuery()

                MessageBox.Show($"✅ Item {actionText} successfully!" & vbCrLf &
                              $"📅 " & DateTime.Now.ToString("MMM dd, yyyy HH:mm"))

                ClearEditFields()
                LoadCurrentDGV()

            Catch ex As Exception
                MessageBox.Show("Operation failed: " & ex.Message)
            Finally
                CloseConn()
            End Try
        End If
    End Sub

    Private Sub RestoreFromArchiveBTN_Click(sender As Object, e As EventArgs) Handles RestoreFromArchiveBTN.Click
        If ConnDB.IsOfflineMode Then
            MessageBox.Show("Cannot restore while offline!", "Offline Mode", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If ClothesEditModalClothesIDTB.Text = "" Then
            MessageBox.Show("Please select an archived item first!", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Confirm restore
        Dim result = MessageBox.Show(
        "👗 Restore this item to ACTIVE INVENTORY?" & vbCrLf & vbCrLf &
        "✓ Status: Available" & vbCrLf &
        "✓ Clears archive date" & vbCrLf &
        "✓ Item returns to main inventory",
        "Restore from Storage",
        MessageBoxButtons.YesNoCancel,
        MessageBoxIcon.Question
    )

        If result <> DialogResult.Yes Then Return

        Try
            OpenConn()

            ' Restore: Set Status back to 'Available' and clear archive fields
            Dim restoreQuery As String = "
            UPDATE Clothes SET 
                Status = 'Available',
                Archived_Date = NULL
            WHERE Clothes_ID = @id AND Status = 'Archived'"

            cmd = New MySqlCommand(restoreQuery, conn)
            cmd.Parameters.AddWithValue("@id", Convert.ToInt32(ClothesEditModalClothesIDTB.Text))

            Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

            If rowsAffected > 0 Then
                MessageBox.Show(
                "✅ Item restored to ACTIVE INVENTORY!" & vbCrLf & vbCrLf &
                "📅 Restored: " & DateTime.Now.ToString("MMM dd, yyyy hh:mm tt") & vbCrLf &
                "👗 Now visible in Active Inventory",
                "Restore Successful",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )

                ' Clear fields and refresh
                ClearEditFields()

                ' Switch to active mode and refresh
                isArchiveMode = False
                ToggleDGVVisibility()
                LoadActiveClothes()
                UpdateToggleUI()

            Else
                MessageBox.Show("Item not found in archive or already restored!", "Restore Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            MessageBox.Show("Restore failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            CloseConn()
        End Try
    End Sub


    Private Sub ClearEditFields()
        ClothesEditModalClothesIDTB.Clear()
        ClothesEditModalClothesNameTB.Clear()
        ClothesEditModalCategoryTB.Clear()
        ClothesEditModalSizeTB.Clear()
        ClothesEditModalFabricTypeTB.Clear()
        ClothesEditModalColorTB.Clear()
        ClothesEditModalConditionTB.Clear()
    End Sub



    Private Sub ClothesMarkAsRepairBTN_Click(sender As Object, e As EventArgs) Handles ClothesMarkAsRepairBTN.Click

        If ClothesEditModalClothesIDTB.Text = "" Then
            MessageBox.Show("Select item first")
            Return
        End If

        Try
            OpenConn()

            Dim query As String = "UPDATE Clothes SET Status='For Repair' WHERE Clothes_ID=@id"
            cmd = New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@id", ClothesEditModalClothesIDTB.Text)

            cmd.ExecuteNonQuery()

            MessageBox.Show("Marked as For Repair")
            LoadClothes()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try

    End Sub

    Private Sub ClearAddFields()
        ClothesAddModalClothesNameTB.Clear()
        ClothesAddModalCategoryTB.Clear()
        ClothesAddModalSizeTB.Clear()
        ClothesAddModalFabricTypeTB.Clear()
        ClothesAddModalColorTB.Clear()
    End Sub


    ' =========================
    ' ARCHIVE TOGGLE SYSTEM
    ' =========================

    ' Archive mode tracking
    Private isArchiveMode As Boolean = False

    Private Sub ToggleArchiveBTN_Click(sender As Object, e As EventArgs) Handles ToggleArchiveBTN.Click
        If ConnDB.IsOfflineMode Then
            MessageBox.Show("Cannot switch views while offline!", "Offline Mode")
            Return
        End If

        isArchiveMode = Not isArchiveMode
        ToggleDGVVisibility()
        LoadCurrentDGV()
        UpdateToggleUI()
    End Sub

    Private Sub ToggleDGVVisibility()
        If isArchiveMode Then
            ArchiveClothesDGV.Visible = True
            ArchiveClothesDGV.BringToFront()
            ClothesDGV.Visible = False
            RestoreFromArchiveBTN.Visible = True
            ClothesDeleteBTN.Visible = False
        Else
            ClothesDGV.Visible = True
            ClothesDGV.BringToFront()
            ArchiveClothesDGV.Visible = False
            RestoreFromArchiveBTN.Visible = False
            ClothesDeleteBTN.Visible = True
        End If
    End Sub

    Private Sub UpdateToggleUI()
        If isArchiveMode Then
            ToggleArchiveBTN.Text = "👗 SHOW ACTIVE INVENTORY"
            ToggleArchiveBTN.BackColor = Color.OrangeRed
            ToggleArchiveBTN.ForeColor = Color.White
            ArchiveModeLabel.Text = "📦 ARCHIVE/STORAGE MODE"
            ArchiveModeLabel.ForeColor = Color.OrangeRed

        Else
            ToggleArchiveBTN.Text = "📦 SHOW ARCHIVE/STORAGE"
            ToggleArchiveBTN.BackColor = Color.DodgerBlue
            ToggleArchiveBTN.ForeColor = Color.White
            ArchiveModeLabel.Text = "👗 ACTIVE INVENTORY MODE"
            ArchiveModeLabel.ForeColor = Color.Green

        End If
    End Sub
    Private Sub LoadCurrentDGV()
        If isArchiveMode Then
            LoadArchiveClothes()
        Else
            LoadClothes()
        End If
    End Sub

    Private Sub UpdateArchiveModeUI()
        If isArchiveMode Then
            ToggleArchiveBTN.Text = "📦 ACTIVE CLOTHES"
            ToggleArchiveBTN.BackColor = Color.Orange
            ToggleArchiveBTN.ForeColor = Color.White
            ArchiveModeLabel.Text = "📦 ARCHIVE MODE - Storage"
            ArchiveModeLabel.ForeColor = Color.Orange

        Else
            ToggleArchiveBTN.Text = "📦 ARCHIVE MODE"
            ToggleArchiveBTN.BackColor = Color.DodgerBlue
            ToggleArchiveBTN.ForeColor = Color.White
            ArchiveModeLabel.Text = "👗 ACTIVE INVENTORY"
            ArchiveModeLabel.ForeColor = Color.Green

        End If
    End Sub

    Private Sub LoadCurrentView()
        If ConnDB.IsOfflineMode Then
            MessageBox.Show("Cannot load data while offline!", "Offline Mode")
            Return
        End If

        If isArchiveMode Then
            LoadArchiveClothes()
        Else
            LoadClothes()
        End If
    End Sub



    ' =========================
    ' MATERIALS INVENTORY PANEL
    ' =========================

    Private Sub LoadMaterials()
        Dim query As String = "SELECT m.*, s.Supplier_Name 
                          FROM Materials m 
                          LEFT JOIN Supplier s ON m.Supplier_ID = s.Supplier_ID 
                          ORDER BY m.Material_ID DESC"
        LoadToDGV(query, MaterialsDGV)
    End Sub

    Private Sub LoadSuppliers()
        Dim query As String = "SELECT * FROM Supplier ORDER BY Supplier_Name"
        LoadToDGV(query, SuppliersDGV)
    End Sub

    Private Sub LoadSuppliersToComboBox(comboBox As ComboBox)
        Try
            OpenConn()
            Dim query As String = "SELECT Supplier_ID, Supplier_Name FROM Supplier ORDER BY Supplier_Name"
            cmd = New MySqlCommand(query, conn)
            Dim adapter As New MySqlDataAdapter(cmd)
            Dim dt As New DataTable
            adapter.Fill(dt)

            comboBox.DataSource = dt
            comboBox.DisplayMember = "Supplier_Name"
            comboBox.ValueMember = "Supplier_ID"
            comboBox.SelectedIndex = -1

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try
    End Sub

    ' =========================
    ' SUPPLIER PANEL EVENTS
    ' =========================

    Private Sub AddSuppliersBTN_Click(sender As Object, e As EventArgs) Handles AddSuppliersBTN.Click
        Try
            OpenConn()

            Dim query As String = "INSERT INTO Supplier (Supplier_Name, Contact_Number, Address, Email) 
                              VALUES (@name, @contact, @address, @email)"

            cmd = New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@name", SuppliersNameTB.Text)
            cmd.Parameters.AddWithValue("@contact", SuppliersContactTB.Text)
            cmd.Parameters.AddWithValue("@address", SuppliersAddressTB.Text)
            cmd.Parameters.AddWithValue("@email", If(SuppliersEmailTB.Text = "", DBNull.Value, SuppliersEmailTB.Text))

            cmd.ExecuteNonQuery()
            MessageBox.Show("Supplier added successfully!")

            ClearSupplierFields()
            LoadSuppliers()
            LoadSuppliersToComboBox(AddMaterialSupplierCMB)
            LoadSuppliersToComboBox(UpdateMaterialSupplierCMB)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try
    End Sub

    Private Sub SuppliersDGV_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles SuppliersDGV.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = SuppliersDGV.Rows(e.RowIndex)
            SuppliersNameTB.Text = row.Cells("Supplier_Name").Value.ToString()
            SuppliersContactTB.Text = row.Cells("Contact_Number").Value.ToString()
            SuppliersAddressTB.Text = row.Cells("Address").Value.ToString()
            If Not IsDBNull(row.Cells("Email").Value) Then
                SuppliersEmailTB.Text = row.Cells("Email").Value.ToString()
            Else
                SuppliersEmailTB.Text = ""
            End If
        End If
    End Sub

    Private Sub ClearSupplierFields()
        SuppliersNameTB.Clear()
        SuppliersContactTB.Clear()
        SuppliersAddressTB.Clear()
        SuppliersEmailTB.Clear()
    End Sub

    ' =========================
    ' MATERIALS PANEL EVENTS
    ' =========================

    Private Sub AddMaterialBTN_Click(sender As Object, e As EventArgs) Handles AddMaterialBTN.Click
        AddMaterialModalPanel.Visible = True
        AddMaterialModalPanel.BringToFront()
        ClearAddMaterialFields()
        LoadSuppliersToComboBox(AddMaterialSupplierCMB)
    End Sub

    Private Sub UpdateMaterialBTN_Click(sender As Object, e As EventArgs) Handles UpdateMaterialBTN.Click
        If UpdateMaterialIDTB IsNot Nothing AndAlso UpdateMaterialIDTB.Text = "" Then
            MessageBox.Show("Please select a material first")
            Return
        End If
        UpdateMaterialModalPanel.Visible = True
        UpdateMaterialModalPanel.BringToFront()
        LoadSuppliersToComboBox(UpdateMaterialSupplierCMB)
    End Sub

    Private Sub AddMaterialModalCancelBTN_Click(sender As Object, e As EventArgs) Handles AddMaterialModalCancelBTN.Click
        AddMaterialModalPanel.Visible = False
        ClearAddMaterialFields()
    End Sub

    Private Sub UpdateMaterialModalCancelBTN_Click(sender As Object, e As EventArgs) Handles UpdateMaterialModalCancelBTN.Click
        UpdateMaterialModalPanel.Visible = False
    End Sub

    Private Sub AddMaterialModalConfirmBTN_Click(sender As Object, e As EventArgs) Handles AddMaterialModalConfirmBTN.Click
        If AddMaterialSupplierCMB.SelectedIndex = -1 Then
            MessageBox.Show("Please select a supplier")
            Return
        End If

        Try
            OpenConn()

            Dim query As String = "INSERT INTO Materials 
                              (Material_Name, Description, Quantity_in_Stock, Unit_of_Measure, Supplier_ID)
                              VALUES (@name, @desc, @qty, @unit, @supplierId)"

            cmd = New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@name", AddMaterialNameTB.Text)
            cmd.Parameters.AddWithValue("@desc", If(AddMaterialDescriptionTB.Text = "", DBNull.Value, AddMaterialDescriptionTB.Text))
            cmd.Parameters.AddWithValue("@qty", Convert.ToInt32(AddMaterialQuantityOnStockTB.Text))
            cmd.Parameters.AddWithValue("@unit", AddMaterialUnitOfMeasureTB.Text)
            cmd.Parameters.AddWithValue("@supplierId", AddMaterialSupplierCMB.SelectedValue)

            cmd.ExecuteNonQuery()
            MessageBox.Show("Material added successfully!")

            AddMaterialModalPanel.Visible = False
            ClearAddMaterialFields()
            LoadMaterials()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try
    End Sub

    Private Sub MaterialsDGV_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles MaterialsDGV.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = MaterialsDGV.Rows(e.RowIndex)

            ' Assuming you have UpdateMaterialIDTB text box for storing ID
            If UpdateMaterialIDTB IsNot Nothing Then
                UpdateMaterialIDTB.Text = row.Cells("Material_ID").Value.ToString()
            End If

            UpdateMaterialNameTB.Text = row.Cells("Material_Name").Value.ToString()
            If Not IsDBNull(row.Cells("Description").Value) Then
                UpdateMaterialDescriptionTB.Text = row.Cells("Description").Value.ToString()
            Else
                UpdateMaterialDescriptionTB.Text = ""
            End If
            UpdateMaterialQuantityOnStockTB.Text = row.Cells("Quantity_in_Stock").Value.ToString()
            UpdateMaterialUnitOfMeasureTB.Text = row.Cells("Unit_of_Measure").Value.ToString()
            UpdateMaterialSupplierCMB.SelectedValue = row.Cells("Supplier_ID").Value
        End If
    End Sub

    Private Sub UpdateMaterialModalConfirmBTN_Click(sender As Object, e As EventArgs) Handles UpdateMaterialModalConfirmBTN.Click
        If UpdateMaterialSupplierCMB.SelectedIndex = -1 Then
            MessageBox.Show("Please select a supplier")
            Return
        End If

        If UpdateMaterialIDTB Is Nothing OrElse UpdateMaterialIDTB.Text = "" Then
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
            cmd.Parameters.AddWithValue("@id", Convert.ToInt32(UpdateMaterialIDTB.Text))
            cmd.Parameters.AddWithValue("@name", UpdateMaterialNameTB.Text)
            cmd.Parameters.AddWithValue("@desc", If(UpdateMaterialDescriptionTB.Text = "", DBNull.Value, UpdateMaterialDescriptionTB.Text))
            cmd.Parameters.AddWithValue("@qty", Convert.ToInt32(UpdateMaterialQuantityOnStockTB.Text))
            cmd.Parameters.AddWithValue("@unit", UpdateMaterialUnitOfMeasureTB.Text)
            cmd.Parameters.AddWithValue("@supplierId", UpdateMaterialSupplierCMB.SelectedValue)

            cmd.ExecuteNonQuery()
            MessageBox.Show("Material updated successfully!")

            UpdateMaterialModalPanel.Visible = False
            LoadMaterials()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try
    End Sub

    Private Sub ClearAddMaterialFields()
        AddMaterialNameTB.Clear()
        AddMaterialDescriptionTB.Clear()
        AddMaterialQuantityOnStockTB.Clear()
        AddMaterialUnitOfMeasureTB.Clear()
        If AddMaterialSupplierCMB IsNot Nothing Then
            AddMaterialSupplierCMB.SelectedIndex = -1
        End If
    End Sub

    Private Sub ClearUpdateMaterialFields()
        If UpdateMaterialIDTB IsNot Nothing Then UpdateMaterialIDTB.Clear()
        UpdateMaterialNameTB.Clear()
        UpdateMaterialDescriptionTB.Clear()
        UpdateMaterialQuantityOnStockTB.Text = "0"
        UpdateMaterialUnitOfMeasureTB.Clear()
        If UpdateMaterialSupplierCMB IsNot Nothing Then
            UpdateMaterialSupplierCMB.SelectedIndex = -1
        End If
    End Sub

    ' =========================
    ' SEARCH FUNCTIONALITY (Optional - Add TextBox for search)
    ' =========================
    ' Uncomment and add search textbox event if you have one
    'Private Sub MaterialSearchTB_TextChanged(sender As Object, e As EventArgs) Handles MaterialSearchTB.TextChanged
    '    Dim searchText As String = MaterialSearchTB.Text
    '    Dim query As String = "SELECT m.*, s.Supplier_Name 
    '                          FROM Materials m 
    '                          LEFT JOIN Suppliers s ON m.Supplier_ID = s.Supplier_ID 
    '                          WHERE m.Material_Name LIKE '%" & searchText & "%' 
    '                          OR m.Description LIKE '%" & searchText & "%'
    '                          OR s.Supplier_Name LIKE '%" & searchText & "%'
    '                          ORDER BY m.Material_ID DESC"
    '    LoadToDGV(query, MaterialsDGV)
    'End Sub

    ' =========================
    ' LOAD DATA WHEN MATERIALS PANEL IS SHOWN
    ' =========================
    Private Sub MaterialsPanel_VisibleChanged(sender As Object, e As EventArgs) Handles MaterialsPanel.VisibleChanged
        If MaterialsPanel.Visible Then
            LoadMaterials()
            LoadSuppliers()
            LoadSuppliersToComboBox(AddMaterialSupplierCMB)
            LoadSuppliersToComboBox(UpdateMaterialSupplierCMB)
        End If
    End Sub



    ' =========================
    ' CUSTOMER MANAGEMENT PANEL
    ' =========================

    Private Sub LoadCustomers()
        Dim query As String = "SELECT * FROM customer ORDER BY Customer_ID DESC"
        LoadToDGV(query, CustomerDGV)
    End Sub

    ' =========================
    ' CUSTOMER BUTTON EVENTS
    ' =========================

    Private Sub AddCustomerBTN_Click(sender As Object, e As EventArgs) Handles AddCustomerBTN.Click
        AddCustomerModalPanel.Visible = True
        AddCustomerModalPanel.BringToFront()
        ClearAddCustomerFields()
    End Sub

    Private Sub UpdateCustomerBTN_Click(sender As Object, e As EventArgs) Handles UpdateCustomerBTN.Click
        If UpdateCustomerModalIDTB.Text = "" Then
            MessageBox.Show("Please select a customer first")
            Return
        End If
        UpdateCustomerModalPanel.Visible = True
        UpdateCustomerModalPanel.BringToFront()
    End Sub

    Private Sub DeleteCustomerBTN_Click(sender As Object, e As EventArgs) Handles DeleteCustomerBTN.Click
        If UpdateCustomerModalIDTB.Text = "" Then
            MessageBox.Show("Please select a customer first")
            Return
        End If

        If MessageBox.Show("Are you sure you want to delete this customer?", "Confirm Delete",
                           MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then

            Try
                OpenConn()

                Dim query As String = "DELETE FROM customer WHERE Customer_ID = @id"
                cmd = New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", Convert.ToInt32(UpdateCustomerModalIDTB.Text))

                cmd.ExecuteNonQuery()
                MessageBox.Show("Customer deleted successfully!")

                UpdateCustomerModalIDTB.Clear()
                LoadCustomers()

            Catch ex As Exception
                MessageBox.Show("Error deleting customer: " & ex.Message)
            Finally
                CloseConn()
            End Try
        End If
    End Sub

    ' =========================
    ' ADD CUSTOMER MODAL EVENTS
    ' =========================

    Private Sub AddCustomerModalCancelBTN_Click(sender As Object, e As EventArgs) Handles AddCustomerModalCancelBTN.Click
        AddCustomerModalPanel.Visible = False
        ClearAddCustomerFields()
    End Sub

    Private Sub AddCustomerModalConfirmBTN_Click(sender As Object, e As EventArgs) Handles AddCustomerModalConfirmBTN.Click
        ' Validation
        If AddCustomerModalNameTB.Text.Trim() = "" Then
            MessageBox.Show("Please enter customer name")
            Return
        End If

        If AddCustomerModalContactNoTB.Text.Trim() = "" Then
            MessageBox.Show("Please enter contact number")
            Return
        End If

        If AddCustomerModalAddressTB.Text.Trim() = "" Then
            MessageBox.Show("Please enter address")
            Return
        End If

        Try
            OpenConn()

            Dim query As String = "INSERT INTO customer (Full_Name, Contact_Number, Address) 
                              VALUES (@name, @contact, @address)"

            cmd = New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@name", AddCustomerModalNameTB.Text.Trim())
            cmd.Parameters.AddWithValue("@contact", AddCustomerModalContactNoTB.Text.Trim())
            cmd.Parameters.AddWithValue("@address", AddCustomerModalAddressTB.Text.Trim())

            cmd.ExecuteNonQuery()
            MessageBox.Show("Customer added successfully!")

            AddCustomerModalPanel.Visible = False
            ClearAddCustomerFields()
            LoadCustomers()

        Catch ex As Exception
            MessageBox.Show("Error adding customer: " & ex.Message)
        Finally
            CloseConn()
        End Try
    End Sub

    ' =========================
    ' UPDATE CUSTOMER MODAL EVENTS
    ' =========================

    Private Sub UpdateCustomerModalCancelBTN_Click(sender As Object, e As EventArgs) Handles UpdateCustomerModalCancelBTN.Click
        UpdateCustomerModalPanel.Visible = False
    End Sub

    Private Sub UpdateCustomerModalConfirmBTN_Click(sender As Object, e As EventArgs) Handles UpdateCustomerModalConfirmBTN.Click
        ' Validation
        If UpdateCustomerModalIDTB.Text = "" Then
            MessageBox.Show("No customer selected")
            Return
        End If

        If UpdateCustomerModalNameTB.Text.Trim() = "" Then
            MessageBox.Show("Please enter customer name")
            Return
        End If

        If UpdateCustomerModalContactNoTB.Text.Trim() = "" Then
            MessageBox.Show("Please enter contact number")
            Return
        End If

        If UpdateCustomerModalAddressTB.Text.Trim() = "" Then
            MessageBox.Show("Please enter address")
            Return
        End If

        Try
            OpenConn()

            Dim query As String = "UPDATE customer SET 
                              Full_Name = @name,
                              Contact_Number = @contact,
                              Address = @address
                              WHERE Customer_ID = @id"

            cmd = New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@id", Convert.ToInt32(UpdateCustomerModalIDTB.Text))
            cmd.Parameters.AddWithValue("@name", UpdateCustomerModalNameTB.Text.Trim())
            cmd.Parameters.AddWithValue("@contact", UpdateCustomerModalContactNoTB.Text.Trim())
            cmd.Parameters.AddWithValue("@address", UpdateCustomerModalAddressTB.Text.Trim())

            cmd.ExecuteNonQuery()
            MessageBox.Show("Customer updated successfully!")

            UpdateCustomerModalPanel.Visible = False
            LoadCustomers()

        Catch ex As Exception
            MessageBox.Show("Error updating customer: " & ex.Message)
        Finally
            CloseConn()
        End Try
    End Sub

    ' =========================
    ' DGV CELL CLICK - POPULATE UPDATE FIELDS
    ' =========================

    Private Sub CustomerDGV_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles CustomerDGV.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = CustomerDGV.Rows(e.RowIndex)

            UpdateCustomerModalIDTB.Text = row.Cells("Customer_ID").Value.ToString()
            UpdateCustomerModalNameTB.Text = row.Cells("Full_Name").Value.ToString()
            UpdateCustomerModalContactNoTB.Text = row.Cells("Contact_Number").Value.ToString()
            UpdateCustomerModalAddressTB.Text = row.Cells("Address").Value.ToString()
        End If
    End Sub

    ' =========================
    ' SEARCH FUNCTIONALITY
    ' =========================

    Private Sub CustomerSearchTB_TextChanged(sender As Object, e As EventArgs) Handles CustomerSearchTB.TextChanged
        Dim searchText As String = CustomerSearchTB.Text.Trim()
        Dim query As String = "SELECT * FROM customer 
                          WHERE Full_Name LIKE '%" & searchText & "%'
                          OR Contact_Number LIKE '%" & searchText & "%'
                          OR Address LIKE '%" & searchText & "%'
                          ORDER BY Customer_ID DESC"
        LoadToDGV(query, CustomerDGV)
    End Sub

    ' =========================
    ' CLEAR FIELDS
    ' =========================

    Private Sub ClearAddCustomerFields()
        AddCustomerModalNameTB.Clear()
        AddCustomerModalContactNoTB.Clear()
        AddCustomerModalAddressTB.Clear()
    End Sub

    ' =========================
    ' LOAD DATA WHEN CUSTOMER PANEL IS SHOWN
    ' =========================

    Private Sub CustomerPanel_VisibleChanged(sender As Object, e As EventArgs) Handles CustomerPanel.VisibleChanged
        If CustomerPanel.Visible Then
            LoadCustomers()
        End If
    End Sub



    ' =========================
    ' RENT TRANSACTION PANEL
    ' =========================

    Private Sub LoadRentTransactions()

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

        LoadToDGV(query, RentDGV)

    End Sub

    Private Sub LoadCustomersToRentCombo()

        Try
            OpenConn()

            Dim query As String = "SELECT Customer_ID, Full_Name FROM customer ORDER BY Full_Name"

            cmd = New MySqlCommand(query, conn)

            Dim adapter As New MySqlDataAdapter(cmd)
            Dim dt As New DataTable

            adapter.Fill(dt)

            RentCustomerNameCMB.DataSource = dt
            RentCustomerNameCMB.DisplayMember = "Full_Name"
            RentCustomerNameCMB.ValueMember = "Customer_ID"
            RentCustomerNameCMB.SelectedIndex = -1

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try

    End Sub

    Private Sub LoadAvailableClothesToCombo()

        Try
            OpenConn()

            Dim query As String = "
        SELECT Clothes_ID, Clothes_Name
        FROM clothes
        WHERE Status = 'Available'
        ORDER BY Clothes_Name"

            cmd = New MySqlCommand(query, conn)

            Dim adapter As New MySqlDataAdapter(cmd)
            Dim dt As New DataTable

            adapter.Fill(dt)

            RentClothesCMB.DataSource = dt
            RentClothesCMB.DisplayMember = "Clothes_Name"
            RentClothesCMB.ValueMember = "Clothes_ID"
            RentClothesCMB.SelectedIndex = -1

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try

    End Sub

    Private Sub RentPanel_VisibleChanged(sender As Object, e As EventArgs) Handles RentPanel.VisibleChanged

        If RentPanel.Visible Then
            LoadRentTransactions()
            LoadCustomersToRentCombo()
            LoadAvailableClothesToCombo()
        End If

    End Sub

    Private Sub RentItemBTN_Click(sender As Object, e As EventArgs) Handles RentItemBTN.Click

        If RentCustomerNameCMB.SelectedIndex = -1 Then
            MessageBox.Show("Please select customer")
            Return
        End If

        If RentClothesCMB.SelectedIndex = -1 Then
            MessageBox.Show("Please select clothes")
            Return
        End If

        Try
            OpenConn()

            Dim insertQuery As String = "
        INSERT INTO rent
        (
            Customer_ID,
            Clothes_ID,
            Date_Rented,
            Expected_Return_Date,
            Rental_Status
        )
        VALUES
        (
            @customerId,
            @clothesId,
            @dateRented,
            @expectedReturn,
            'Rented'
        )"

            cmd = New MySqlCommand(insertQuery, conn)

            cmd.Parameters.AddWithValue("@customerId", RentCustomerNameCMB.SelectedValue)
            cmd.Parameters.AddWithValue("@clothesId", RentClothesCMB.SelectedValue)
            cmd.Parameters.AddWithValue("@dateRented", RentDateRentedDTP.Value.Date)
            cmd.Parameters.AddWithValue("@expectedReturn", RentExpectedReturnDTP.Value.Date)

            cmd.ExecuteNonQuery()

            Dim updateClothesQuery As String = "
        UPDATE clothes
        SET Status = 'Rented'
        WHERE Clothes_ID = @id"

            cmd = New MySqlCommand(updateClothesQuery, conn)
            cmd.Parameters.AddWithValue("@id", RentClothesCMB.SelectedValue)

            cmd.ExecuteNonQuery()

            MessageBox.Show("Item rented successfully!")

            ClearRentFields()
            LoadRentTransactions()
            LoadAvailableClothesToCombo()
            LoadClothes()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try

    End Sub

    Private Sub ClearRentFields()

        RentCustomerNameCMB.SelectedIndex = -1
        RentClothesCMB.SelectedIndex = -1

        RentDateRentedDTP.Value = Date.Now
        RentExpectedReturnDTP.Value = Date.Now.AddDays(3)

    End Sub

    Private Sub ClearRentItemBTN_Click(sender As Object, e As EventArgs) Handles ClearRentItemBTN.Click
        ClearRentFields()
    End Sub

    Private Sub RentDGV_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles RentDGV.CellClick

        If e.RowIndex >= 0 Then

            Dim row As DataGridViewRow = RentDGV.Rows(e.RowIndex)

            ReturnItemRentIDTB.Text = row.Cells("Rent_ID").Value.ToString()
            ReturnItemCustomerNameTB.Text = row.Cells("Full_Name").Value.ToString()
            ReturnItemClothesNameTB.Text = row.Cells("Clothes_Name").Value.ToString()
            ReturnItemStatusTB.Text = row.Cells("Rental_Status").Value.ToString()

        End If

    End Sub

    Private Sub ReturnItemBTN_Click(sender As Object, e As EventArgs) Handles ReturnItemBTN.Click

        If ReturnItemRentIDTB.Text = "" Then
            MessageBox.Show("Please select rent transaction")
            Return
        End If

        Try
            OpenConn()

            Dim clothesId As Integer = 0

            Dim getClothesQuery As String = "
        SELECT Clothes_ID
        FROM rent
        WHERE Rent_ID = @rentId"

            cmd = New MySqlCommand(getClothesQuery, conn)
            cmd.Parameters.AddWithValue("@rentId", Convert.ToInt32(ReturnItemRentIDTB.Text))

            clothesId = Convert.ToInt32(cmd.ExecuteScalar())

            Dim returnQuery As String = "
        UPDATE rent
        SET
            Rental_Status = 'Returned',
            Actual_Return_Date = NOW()
        WHERE Rent_ID = @rentId"

            cmd = New MySqlCommand(returnQuery, conn)
            cmd.Parameters.AddWithValue("@rentId", Convert.ToInt32(ReturnItemRentIDTB.Text))

            cmd.ExecuteNonQuery()

            Dim updateClothesQuery As String = "
        UPDATE clothes
        SET Status = 'Available'
        WHERE Clothes_ID = @clothesId"

            cmd = New MySqlCommand(updateClothesQuery, conn)
            cmd.Parameters.AddWithValue("@clothesId", clothesId)

            cmd.ExecuteNonQuery()

            MessageBox.Show("Item returned successfully!")

            LoadRentTransactions()
            LoadAvailableClothesToCombo()
            LoadClothes()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try

    End Sub

    Private Sub MarkLostItemBTN_Click(sender As Object, e As EventArgs) Handles MarkLostItemBTN.Click

        If ReturnItemRentIDTB.Text = "" Then
            MessageBox.Show("Please select rent transaction")
            Return
        End If

        Try
            OpenConn()

            Dim clothesId As Integer = 0

            Dim getClothesQuery As String = "
        SELECT Clothes_ID
        FROM rent
        WHERE Rent_ID = @rentId"

            cmd = New MySqlCommand(getClothesQuery, conn)
            cmd.Parameters.AddWithValue("@rentId", Convert.ToInt32(ReturnItemRentIDTB.Text))

            clothesId = Convert.ToInt32(cmd.ExecuteScalar())

            Dim lostQuery As String = "
        UPDATE rent
        SET Rental_Status = 'Lost'
        WHERE Rent_ID = @rentId"

            cmd = New MySqlCommand(lostQuery, conn)
            cmd.Parameters.AddWithValue("@rentId", Convert.ToInt32(ReturnItemRentIDTB.Text))

            cmd.ExecuteNonQuery()

            Dim updateClothesQuery As String = "
        UPDATE clothes
        SET Status = 'Lost'
        WHERE Clothes_ID = @clothesId"

            cmd = New MySqlCommand(updateClothesQuery, conn)
            cmd.Parameters.AddWithValue("@clothesId", clothesId)

            cmd.ExecuteNonQuery()

            MessageBox.Show("Item marked as lost!")

            LoadRentTransactions()
            LoadClothes()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try

    End Sub

    Private Sub ExtendItemBTN_Click(sender As Object, e As EventArgs) Handles ExtendItemBTN.Click

        If ReturnItemRentIDTB.Text = "" Then
            MessageBox.Show("Please select rent transaction")
            Return
        End If

        Try
            OpenConn()

            Dim query As String = "
        UPDATE rent
        SET Expected_Return_Date = DATE_ADD(Expected_Return_Date, INTERVAL 3 DAY)
        WHERE Rent_ID = @rentId"

            cmd = New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@rentId", Convert.ToInt32(ReturnItemRentIDTB.Text))

            cmd.ExecuteNonQuery()

            MessageBox.Show("Return date extended by 3 days!")

            LoadRentTransactions()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try

    End Sub

    Private Sub RentSearchTB_TextChanged(sender As Object, e As EventArgs) Handles RentSearchTB.TextChanged

        Dim searchText As String = RentSearchTB.Text.Trim()

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
        c.Full_Name LIKE '%" & searchText & "%'
        OR cl.Clothes_Name LIKE '%" & searchText & "%'
        OR r.Rental_Status LIKE '%" & searchText & "%'
    ORDER BY r.Rent_ID DESC"

        LoadToDGV(query, RentDGV)

    End Sub

    ' =========================
    ' TAILORING TRANSACTION PANEL
    ' =========================

    Private Sub LoadTailoringTransactions()
        Try
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

            LoadToDGV(query, TailoringDGV)
        Catch ex As Exception
            MessageBox.Show("Error loading tailoring transactions: " & ex.Message)
        End Try
    End Sub

    Private Sub CustomerOwnedToggleBTN_Click(sender As Object, e As EventArgs) Handles CustomerOwnedToggleBTN.Click
        isCustomerOwnedMode = Not isCustomerOwnedMode

        If isCustomerOwnedMode Then
            CustomerOwnedToggleBTN.Text = "👗 SHOP INVENTORY"
            CustomerOwnedToggleBTN.BackColor = Color.Orange
            TailoringClothesCMB.Enabled = False
            TailoringCustomerClothesTB.Enabled = True
            TailoringClothesLabel.Text = "Customer Clothes Description:"
            TailoringDescriptionTB.Enabled = False
        Else
            CustomerOwnedToggleBTN.Text = "👕 CUSTOMER-OWNED"
            CustomerOwnedToggleBTN.BackColor = Color.DodgerBlue
            TailoringClothesCMB.Enabled = True
            TailoringCustomerClothesTB.Enabled = False
            TailoringClothesLabel.Text = "Shop Inventory Clothes:"
            TailoringDescriptionTB.Enabled = True
        End If
    End Sub



    Private Sub TailoringPanel_VisibleChanged(sender As Object, e As EventArgs) Handles TailoringPanel.VisibleChanged
        If TailoringPanel.Visible Then
            LoadTailoringTransactions()
            LoadCustomersToTailoringCombo()
            LoadTailoringCatalog()
            LoadAvailableClothesToTailoringCombo()
            LoadEmployeesToTailoringCombo()
            ResetTailoringUI()
        End If
    End Sub

    Private Sub LoadCustomersToTailoringCombo()
        Try
            OpenConn()
            Dim query As String = "SELECT Customer_ID, Full_Name FROM customer ORDER BY Full_Name"
            cmd = New MySqlCommand(query, conn)
            Dim adapter As New MySqlDataAdapter(cmd)
            Dim dt As New DataTable
            adapter.Fill(dt)

            TailoringCustomerCMB.DataSource = Nothing
            TailoringCustomerCMB.DataSource = dt
            TailoringCustomerCMB.DisplayMember = "Full_Name"
            TailoringCustomerCMB.ValueMember = "Customer_ID"
            TailoringCustomerCMB.SelectedIndex = -1
        Catch ex As Exception
        Finally
            CloseConn()
        End Try
    End Sub

    Private Sub LoadEmployeesToTailoringCombo()
        Try
            OpenConn()
            Dim query As String = "SELECT Employee_ID, CONCAT(Full_Name, ' - ', Position) as DisplayName FROM Employee WHERE Date_Hired IS NOT NULL ORDER BY Full_Name"
            cmd = New MySqlCommand(query, conn)
            Dim adapter As New MySqlDataAdapter(cmd)
            Dim dt As New DataTable
            adapter.Fill(dt)

            TailoringEmployeeCMB.DataSource = Nothing
            TailoringEmployeeCMB.DataSource = dt
            TailoringEmployeeCMB.DisplayMember = "DisplayName"
            TailoringEmployeeCMB.ValueMember = "Employee_ID"
            TailoringEmployeeCMB.SelectedIndex = -1
        Catch ex As Exception
            MessageBox.Show("Error loading employees: " & ex.Message)
        Finally
            CloseConn()
        End Try
    End Sub


    Private Sub LoadAvailableClothesToTailoringCombo()
        Try
            OpenConn()
            Dim query As String = "
            SELECT Clothes_ID, CONCAT(Clothes_Name, ' - ', Category, ' (', Size, ')') as DisplayName
            FROM Clothes WHERE Status IN ('Available', 'For Repair')
            ORDER BY Clothes_Name"

            cmd = New MySqlCommand(query, conn)
            Dim adapter As New MySqlDataAdapter(cmd)
            Dim dt As New DataTable
            adapter.Fill(dt)

            TailoringClothesCMB.DataSource = Nothing
            TailoringClothesCMB.DataSource = dt
            TailoringClothesCMB.DisplayMember = "DisplayName"
            TailoringClothesCMB.ValueMember = "Clothes_ID"
            TailoringClothesCMB.SelectedIndex = -1
        Catch ex As Exception
        Finally
            CloseConn()
        End Try
    End Sub

    Private Sub LoadTailoringCatalog()
        ' ✅ INSTANT LOAD FROM MEMORY (NO DB!)
        Dim catalogData = TailoringCatalog.Select(Function(s) New With {
        .Catalog_ID = s.CatalogID,
        .Service_Name = s.ServiceName
    }).ToList()

        TailoringTypeCMB.DataSource = Nothing
        TailoringTypeCMB.DataSource = catalogData
        TailoringTypeCMB.DisplayMember = "Service_Name"
        TailoringTypeCMB.ValueMember = "Catalog_ID"
        TailoringTypeCMB.SelectedIndex = -1
    End Sub

    ' Smart Features
    Private Sub TailoringTypeCMB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TailoringTypeCMB.SelectedIndexChanged
        LoadRequiredMaterials()
        If TailoringClothesCMB.SelectedIndex >= 0 Then CalculateSmartPrice()
    End Sub

    Private Sub TailoringClothesCMB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TailoringClothesCMB.SelectedIndexChanged
        CalculateSmartPrice()
    End Sub

    Private Sub CalculateSmartPrice()
        Dim serviceName = TailoringTypeCMB.Text
        Dim fabricType = GetSelectedClothesFabric()

        If String.IsNullOrEmpty(serviceName) Then Return

        Dim service = TailoringCatalog.FirstOrDefault(Function(s) s.ServiceName = serviceName)
        If service Is Nothing Then Return

        Dim multiplier = 1D
        If FabricPricingRules.ContainsKey(serviceName) AndAlso FabricPricingRules(serviceName).ContainsKey(fabricType) Then
            multiplier = FabricPricingRules(serviceName)(fabricType)
        End If

        Dim finalPrice = Math.Max(service.BasePrice * multiplier, 50D) ' Minimum ₱50
        TailoringPriceTB.Text = finalPrice.ToString("F2") ' Just the number
        TailoringPriceLabel.Text = $"Base: ₱{service.BasePrice:F2} × {multiplier:F1}x ({fabricType})"
    End Sub

    Private Function GetSelectedClothesFabric() As String
        If TailoringClothesCMB.SelectedIndex >= 0 AndAlso TailoringClothesCMB.SelectedValue IsNot Nothing Then
            Try
                OpenConn()
                Dim query = "SELECT Fabric_Type FROM Clothes WHERE Clothes_ID = @id"
                cmd = New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", TailoringClothesCMB.SelectedValue)

                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing Then
                    Return result.ToString()
                End If
                Return "Cotton"
            Catch ex As Exception
                Return "Cotton"
            Finally
                CloseConn()
            End Try
        End If
        Return "Cotton"
    End Function

    Private Sub LoadRequiredMaterials()
        Dim serviceName = TailoringTypeCMB.Text
        If String.IsNullOrEmpty(serviceName) Then Return

        Dim service = TailoringCatalog.FirstOrDefault(Function(s) s.ServiceName = serviceName)
        If service Is Nothing Then Return

        Dim requirements = ServiceRequirements.GetValueOrDefault(service.CatalogID, New List(Of ServiceRequirement))

        ' ✅ ENHANCED: Show actual stock availability
        Dim materialsWithStock = New List(Of Object)

        For Each req In requirements
            Dim stockLevel As Integer = 0
            Try
                OpenConn()
                Dim stockQuery = "SELECT Quantity_in_Stock FROM Materials WHERE Material_ID = @id"
                cmd = New MySqlCommand(stockQuery, conn)
                cmd.Parameters.AddWithValue("@id", req.MaterialID)
                stockLevel = Convert.ToInt32(cmd.ExecuteScalar())
                CloseConn()
            Catch
                stockLevel = 0
            End Try

            Dim stockStatus As String
            If stockLevel >= req.DefaultQuantity Then
                stockStatus = $"✅ IN STOCK ({stockLevel})"
            ElseIf stockLevel > 0 Then
                stockStatus = $"⚠️ LOW STOCK ({stockLevel}/{req.DefaultQuantity})"
            Else
                stockStatus = "❌ OUT OF STOCK"
            End If

            materialsWithStock.Add(New With {
            .Material_ID = req.MaterialID,
            .Material_Name = req.MaterialName,
            .Default_Quantity = req.DefaultQuantity,
            .Unit_Measure = req.UnitMeasure,
            .Stock_Available = stockLevel,
            .Stock_Status = stockStatus
        })
        Next

        TailoringMaterialsDGV.DataSource = materialsWithStock
    End Sub

    ' Create Request
    Private Sub CreateTailoringRequestBTN_Click(sender As Object, e As EventArgs) Handles CreateTailoringRequestBTN.Click
        If TailoringCustomerCMB.SelectedIndex = -1 OrElse String.IsNullOrEmpty(TailoringTypeCMB.Text) Then
            MessageBox.Show("⚠️ Select Customer + Service Type")
            Return
        End If

        If TailoringEmployeeCMB.SelectedIndex = -1 Then  ' ✅ VALIDATION
            MessageBox.Show("⚠️ Please assign an Employee")
            Return
        End If

        If String.IsNullOrEmpty(TailoringPriceTB.Text) Then
            MessageBox.Show("⚠️ Select Service Type first to auto-calculate price")
            Return
        End If

        Try
            OpenConn()

            Dim insertQuery = "
        INSERT INTO Tailoring_Services 
        (Type_of_Alteration, Description_of_Work, Service_Price, Date_Requested, Status, 
         Customer_ID, Clothes_ID, Is_Customer_Owned, Employee_ID)  
        VALUES (@type, @desc, @price, NOW(), 'Pending', @customerId, @clothesId, @isOwned, @employeeId)"

            cmd = New MySqlCommand(insertQuery, conn)
            cmd.Parameters.AddWithValue("@type", TailoringTypeCMB.Text)
            cmd.Parameters.AddWithValue("@desc", If(isCustomerOwnedMode, TailoringCustomerClothesTB.Text, TailoringDescriptionTB.Text))
            cmd.Parameters.AddWithValue("@price", Convert.ToDecimal(TailoringPriceTB.Text))
            cmd.Parameters.AddWithValue("@customerId", TailoringCustomerCMB.SelectedValue)
            cmd.Parameters.AddWithValue("@clothesId", If(isCustomerOwnedMode, DBNull.Value, TailoringClothesCMB.SelectedValue))
            cmd.Parameters.AddWithValue("@isOwned", If(isCustomerOwnedMode, 1, 0))
            cmd.Parameters.AddWithValue("@employeeId", TailoringEmployeeCMB.SelectedValue)  ' ✅ ADD EMPLOYEE

            cmd.ExecuteNonQuery()

            ' Update clothes status if shop inventory
            If Not isCustomerOwnedMode AndAlso TailoringClothesCMB.SelectedValue IsNot Nothing Then
                Dim updateClothes = "UPDATE Clothes SET Status = 'In Tailoring' WHERE Clothes_ID = @id"
                cmd = New MySqlCommand(updateClothes, conn)
                cmd.Parameters.AddWithValue("@id", TailoringClothesCMB.SelectedValue)
                cmd.ExecuteNonQuery()
            End If

            ' 🔥 NEW: AUTO-DEDUCT MATERIALS FROM STOCK
            Dim service = TailoringCatalog.FirstOrDefault(Function(s) s.ServiceName = TailoringTypeCMB.Text)
            If service IsNot Nothing Then
                Dim requirements = ServiceRequirements.GetValueOrDefault(service.CatalogID, New List(Of ServiceRequirement))
                Dim serviceId As Integer = Convert.ToInt32(cmd.LastInsertedId)

                For Each req In requirements
                    Try
                        ' 1. Deduct from Materials stock
                        Dim deductQuery = "UPDATE Materials SET Quantity_in_Stock = GREATEST(0, Quantity_in_Stock - @qty) WHERE Material_ID = @materialId"
                        cmd = New MySqlCommand(deductQuery, conn)
                        cmd.Parameters.AddWithValue("@qty", req.DefaultQuantity)
                        cmd.Parameters.AddWithValue("@materialId", req.MaterialID)
                        cmd.ExecuteNonQuery()

                        ' 2. Record in Required_Materials table (for audit trail)
                        Dim recordQuery = "INSERT INTO Required_Materials (Tailoring_Services_ID, Material_ID, Quantity_Used, Unit_Measure_Used) VALUES (@serviceId, @materialId, @qty, @unit)"
                        cmd = New MySqlCommand(recordQuery, conn)
                        cmd.Parameters.AddWithValue("@serviceId", serviceId)
                        cmd.Parameters.AddWithValue("@materialId", req.MaterialID)
                        cmd.Parameters.AddWithValue("@qty", req.DefaultQuantity)
                        cmd.Parameters.AddWithValue("@unit", req.UnitMeasure)
                        cmd.ExecuteNonQuery()

                    Catch deductEx As Exception
                        ' Log but don't stop the whole transaction
                        Debug.WriteLine($"⚠️ Failed to deduct material {req.MaterialID}: {deductEx.Message}")
                    End Try
                Next
            End If

            MessageBox.Show("✅ Tailoring request created!" & vbCrLf &
                   $"💰 ₱" & TailoringPriceTB.Text & vbCrLf &
                   $"👷 Assigned: " & TailoringEmployeeCMB.Text & vbCrLf &
                   $"📦 " & If(isCustomerOwnedMode, "👕 CUSTOMER-OWNED", "👗 SHOP INVENTORY"),
                   "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ClearTailoringFields()
            LoadTailoringTransactions()
            LoadClothes()

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            CloseConn()
        End Try
    End Sub

    Private Sub TailoringDGV_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles TailoringDGV.CellClick
        If e.RowIndex >= 0 Then
            Dim row = TailoringDGV.Rows(e.RowIndex)
            TailoringServiceIDTB.Text = row.Cells("Tailoring_Services_ID").Value.ToString()
        End If
    End Sub

    Private Sub UpdateTailoringStatusBTN_Click(sender As Object, e As EventArgs) Handles UpdateTailoringStatusBTN.Click
        If TailoringServiceIDTB.Text = "" Then
            MessageBox.Show("Select a transaction first")
            Return
        End If

        Try
            OpenConn()
            Dim updateQuery = "UPDATE Tailoring_Services SET Status = @status, Date_Completed = NOW() WHERE Tailoring_Services_ID = @id"
            cmd = New MySqlCommand(updateQuery, conn)
            cmd.Parameters.AddWithValue("@status", TailoringStatusCMB.Text)
            cmd.Parameters.AddWithValue("@id", TailoringServiceIDTB.Text)
            cmd.ExecuteNonQuery()

            If TailoringStatusCMB.Text = "Completed" Then
                Dim updateClothes = "UPDATE Clothes SET Status = 'Available' WHERE Clothes_ID = (SELECT Clothes_ID FROM Tailoring_Services WHERE Tailoring_Services_ID = @id) AND Clothes_ID IS NOT NULL"
                cmd = New MySqlCommand(updateClothes, conn)
                cmd.Parameters.AddWithValue("@id", TailoringServiceIDTB.Text)
                cmd.ExecuteNonQuery()
            End If

            MessageBox.Show("Status updated!")
            LoadTailoringTransactions()
            LoadClothes()

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            CloseConn()
        End Try
    End Sub

    Private Sub ClearTailoringFields()
        TailoringCustomerCMB.SelectedIndex = -1
        TailoringClothesCMB.SelectedIndex = -1
        TailoringEmployeeCMB.SelectedIndex = -1
        TailoringTypeCMB.SelectedIndex = -1
        TailoringCustomerClothesTB.Clear()
        TailoringDescriptionTB.Clear()
        TailoringPriceTB.Clear()
        TailoringPriceLabel.Text = ""
        TailoringMaterialsDGV.DataSource = Nothing
        TailoringServiceIDTB.Clear()
    End Sub

    Private Sub ClearTailoringBTN_Click(sender As Object, e As EventArgs) Handles ClearTailoringBTN.Click
        ClearTailoringFields()
        ResetTailoringUI()
    End Sub

    Private Sub ResetTailoringUI()
        isCustomerOwnedMode = False
        CustomerOwnedToggleBTN.Text = "👕 CUSTOMER-OWNED"
        CustomerOwnedToggleBTN.BackColor = Color.DodgerBlue
        TailoringClothesCMB.Enabled = True
        TailoringCustomerClothesTB.Enabled = False
        TailoringClothesLabel.Text = "Shop Inventory Clothes:"
        TailoringDescriptionTB.Enabled = True
    End Sub


    ' =========================
    ' OFFLINE MODE HANDLER
    ' =========================
    Private Sub UpdateConnectionStatus()
        If ConnDB.IsOfflineMode Then
            ConnectionStatusLabel.ForeColor = Color.Red
            ConnectionStatusLabel.Text = "⚠️ OFFLINE MODE - No database connection"

            ClearAllDataGridViews()
            SetOfflineUI()
        Else
            ConnectionStatusLabel.ForeColor = Color.Green
            ConnectionStatusLabel.Text = "🟢 ONLINE - Connected to database"

            ' Re-enable UI
            SetOnlineUI()
        End If
    End Sub

    Private Sub ClearAllDataGridViews()
        ClothesDGV.DataSource = Nothing
        MaterialsDGV.DataSource = Nothing
        SuppliersDGV.DataSource = Nothing
        CustomerDGV.DataSource = Nothing
        ' Add other DGV controls here
    End Sub

    Private Sub SetOfflineUI()
        ' Disable buttons that require DB
        ClothesAddBTN.Enabled = False
        ClothesEditBTN.Enabled = False
        ClothesDeleteBTN.Enabled = False
        ClothesMarkAsRepairBTN.Enabled = False

        AddMaterialBTN.Enabled = False
        UpdateMaterialBTN.Enabled = False
        AddSuppliersBTN.Enabled = False

        AddCustomerBTN.Enabled = False
        UpdateCustomerBTN.Enabled = False
        DeleteCustomerBTN.Enabled = False

        RentItemBTN.Enabled = False
        ReturnItemBTN.Enabled = False
        MarkLostItemBTN.Enabled = False
        ExtendItemBTN.Enabled = False

        ' You can add more controls here
    End Sub

    Private Sub SetOnlineUI()
        ' Re-enable buttons
        ClothesAddBTN.Enabled = True
        ClothesEditBTN.Enabled = True
        ClothesDeleteBTN.Enabled = True
        ClothesMarkAsRepairBTN.Enabled = True

        AddMaterialBTN.Enabled = True
        UpdateMaterialBTN.Enabled = True
        AddSuppliersBTN.Enabled = True

        AddCustomerBTN.Enabled = True
        UpdateCustomerBTN.Enabled = True
        DeleteCustomerBTN.Enabled = True

        RentItemBTN.Enabled = True
        ReturnItemBTN.Enabled = True
        MarkLostItemBTN.Enabled = True
        ExtendItemBTN.Enabled = True
    End Sub

    ' =========================
    ' RECONNECTION HANDLER
    ' =========================
    Private Sub CheckConnectionStatus()
        Dim wasOffline = ConnDB.IsOfflineMode

        If ConnDB.TestConnection() Then
            If wasOffline Then
                MessageBox.Show("✅ Reconnected to database!", "Connection Restored",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            UpdateConnectionStatus()
        Else
            UpdateConnectionStatus()
        End If
    End Sub


End Class