Imports MySql.Data.MySqlClient


Public Class Form1

    Public materialsManager As MaterialsAndSuppliersManager
    Public clothesManager As ClothesInventoryManager
    Public tailoringCatalogManager As TailoringCatalogManager
    Public rentManager As RentManager
    Public tailoringManager As TailoringManager
    Public reportsManager As ReportsManager
    Public dashboardManager As DashboardManager
    Public Sub New()
        InitializeComponent()
        materialsManager = New MaterialsAndSuppliersManager(Me)
        clothesManager = New ClothesInventoryManager(Me)
        tailoringCatalogManager = New TailoringCatalogManager()
        rentManager = New RentManager(Me)
        tailoringManager = New TailoringManager(Me)
        reportsManager = New ReportsManager(Me)
    End Sub

    Private isDragging As Boolean = False
    Private dragOffset As Point

    Private hasRefreshedAfterReconnect As Boolean = False

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ConnDB.LoadConfig()
        ConnDB.TestConnection()
        UpdateConnectionStatus()
        dashboardManager = New DashboardManager(Me)
        dashboardManager.LoadDashboard()
        InitializeManagers()
        InitializeUI()
        LoadInitialData()

    End Sub

    Private Sub InitializeManagers()
        clothesManager.UpdateArchiveModeUI()
    End Sub

    Private Sub InitializeUI()

        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = FormBorderStyle.Sizable

        MainPanel.Visible = True
        SidebarPanel.Visible = True

        HideAllModals()

    End Sub

    Private Sub LoadInitialData()
        clothesManager.LoadClothes()
        materialsManager.LoadMaterials()
        materialsManager.LoadSuppliers()
    End Sub

    Private Sub HideAllModals()

        Dim modals As Panel() = {
        ClotheAddModalPanel,
        ClotheEditModalPanel,
        AddMaterialModalPanel,
        UpdateMaterialModalPanel,
        AddCustomerModalPanel,
        UpdateCustomerModalPanel
    }

        For Each pnl In modals
            pnl.Visible = False
        Next

    End Sub

    Private Sub ShowPanel(panel As Panel)

        For Each ctrl As Control In MainPanel.Controls
            If TypeOf ctrl Is Panel Then
                ctrl.Visible = False
            End If
        Next

        panel.Visible = True
        panel.BringToFront()

    End Sub


    Private Sub LoadConfigToUI()

        ServerTB.Text = ConnDB.serverName
        DatabaseTB.Text = ConnDB.databaseName
        UsernameTB.Text = ConnDB.dbUsername
        PasswordTB.Text = ConnDB.dbPassword

    End Sub


    ' =========================
    ' BUTTON EVENTS
    ' =========================
    Private Sub DashboardBTN_Click(sender As Object, e As EventArgs) Handles DashboardBTN.Click
        ShowPanel(AdminDashboardPanel)
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

    Private Sub SettingsBTN_Click(sender As Object, e As EventArgs) Handles SettingsBTN.Click
        ShowPanel(SettingsPanel)
        LoadConfigToUI()
    End Sub

    Private Sub ReportsBTN_Click(sender As Object, e As EventArgs) Handles ReportsBTN.Click
        ShowPanel(ReportsPanel)
    End Sub

    ' =========================
    ' CLOTHES INVENTORY PANEL
    ' =========================
    Private Sub ClotheSearchTB_TextChanged(
        sender As Object,
        e As EventArgs
    ) Handles ClotheSearchTB.TextChanged
        clothesManager.SearchClothes(ClotheSearchTB.Text)
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
        clothesManager.AddClothes()
    End Sub

    Private Sub ClothesDGVs_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles ClothesDGV.CellClick, ArchiveClothesDGV.CellClick
        Dim dgv As DataGridView = DirectCast(sender, DataGridView)
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = dgv.Rows(e.RowIndex)
            clothesManager.LoadEditFieldsFromRow(row)
        End If
    End Sub

    Public Sub AutoSelectFirstRow(Optional dgv As DataGridView = Nothing)
        ' Auto-detect current DGV if none specified
        If dgv Is Nothing Then
            dgv = GetCurrentActiveDGV()
        End If

        If dgv IsNot Nothing AndAlso dgv.Rows.Count > 0 Then
            ' Clear any existing selection
            dgv.ClearSelection()
            dgv.CurrentCell = dgv.Rows(0).Cells(0)
            dgv.Rows(0).Selected = True

            ' Trigger the cell click event automatically
            dgv_CellClick(dgv, New DataGridViewCellEventArgs(0, 0))
        Else
            ' Clear fields if no data
            ClearAllEditFields()
        End If
    End Sub

    ' Helper to detect which DGV is currently active
    Public Function GetCurrentActiveDGV() As DataGridView
        If clothesManager.IsArchiveMode Then Return ArchiveClothesDGV
        If MaterialsPanel.Visible Then Return MaterialsDGV
        If CustomerPanel.Visible Then Return CustomerDGV
        If RentPanel.Visible Then Return RentDGV
        If TailoringPanel.Visible Then Return TailoringDGV
        Return ClothesDGV ' Default
    End Function

    Private Sub dgv_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles _
    ClothesDGV.CellClick, ArchiveClothesDGV.CellClick, MaterialsDGV.CellClick,
     CustomerDGV.CellClick, RentDGV.CellClick, TailoringDGV.CellClick

        Dim dgv As DataGridView = DirectCast(sender, DataGridView)
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = dgv.Rows(e.RowIndex)
            LoadFieldsFromRow(dgv, row)
        End If
    End Sub

    Private Sub LoadFieldsFromRow(dgv As DataGridView, row As DataGridViewRow)
        Try
            Select Case dgv.Name
                Case "ClothesDGV", "ArchiveClothesDGV"
                    clothesManager.LoadEditFieldsFromRow(row)

                Case "MaterialsDGV"
                    LoadMaterialsEditFields(row)

                Case "SuppliersDGV"
                    LoadSuppliersEditFields(row)

                Case "CustomerDGV"
                    LoadCustomerEditFields(row)

                Case "RentDGV"
                    rentManager.LoadRentEditFields(row)
                Case "TailoringDGV"
                    tailoringManager.LoadEditFields(row)
            End Select
        Catch ex As Exception
            Debug.WriteLine($"LoadFieldsFromRow Error ({dgv.Name}): {ex.Message}")
        End Try
    End Sub


    ' Materials Edit Fields
    Private Sub LoadMaterialsEditFields(row As DataGridViewRow)
        If UpdateMaterialIDTB IsNot Nothing Then
            UpdateMaterialIDTB.Text = GetCellValue(row.Cells("Material_ID"))
        End If
        UpdateMaterialNameTB.Text = GetCellValue(row.Cells("Material_Name"))
        UpdateMaterialDescriptionTB.Text = GetCellValue(row.Cells("Description"))
        UpdateMaterialQuantityOnStockTB.Text = GetCellValue(row.Cells("Quantity_in_Stock"))
        UpdateMaterialUnitOfMeasureTB.Text = GetCellValue(row.Cells("Unit_of_Measure"))
        ' Supplier combo will be set in MaterialsDGV_CellClick
    End Sub

    ' Suppliers Edit Fields
    Private Sub LoadSuppliersEditFields(row As DataGridViewRow)
        SuppliersNameTB.Text = GetCellValue(row.Cells("Supplier_Name"))
        SuppliersContactTB.Text = GetCellValue(row.Cells("Contact_Number"))
        SuppliersAddressTB.Text = GetCellValue(row.Cells("Address"))
        SuppliersEmailTB.Text = GetCellValue(row.Cells("Email"))
    End Sub

    ' Customer Edit Fields
    Private Sub LoadCustomerEditFields(row As DataGridViewRow)
        UpdateCustomerModalIDTB.Text = GetCellValue(row.Cells("Customer_ID"))
        UpdateCustomerModalNameTB.Text = GetCellValue(row.Cells("Full_Name"))
        UpdateCustomerModalContactNoTB.Text = GetCellValue(row.Cells("Contact_Number"))
        UpdateCustomerModalAddressTB.Text = GetCellValue(row.Cells("Address"))
    End Sub

    ' Tailoring Edit Fields
    Private Sub LoadTailoringEditFields(row As DataGridViewRow)
        TailoringServiceIDTB.Text = GetCellValue(row.Cells("Tailoring_Services_ID"))
    End Sub

    ' Clear ALL edit fields
    Private Sub ClearAllEditFields()
        clothesManager.ClearEditFields()
        ClearUpdateMaterialFields()
        materialsManager.ClearSupplierFields()
        UpdateCustomerModalIDTB.Clear()
        ReturnItemRentIDTB.Clear()
        TailoringServiceIDTB.Clear()
    End Sub

    ' Helper function
    Public Function GetCellValue(cell As DataGridViewCell) As String
        If cell Is Nothing OrElse cell.Value Is Nothing Then
            Return ""
        End If
        Return cell.Value.ToString()
    End Function

    Private Sub ClothesEditModalConfirmBTN_Click(sender As Object, e As EventArgs) Handles ClothesEditModalConfirmBTN.Click
        clothesManager.UpdateClothes()
    End Sub

    ' Archive (Soft Delete) - Changes Status to 'Archived'
    Private Sub ClothesDeleteBTN_Click(sender As Object, e As EventArgs) Handles ClothesDeleteBTN.Click
        clothesManager.SafeClothesDelete()
    End Sub

    Private Sub RestoreFromArchiveBTN_Click(sender As Object, e As EventArgs) Handles RestoreFromArchiveBTN.Click
        clothesManager.RestoreDeletedClothes()
    End Sub


    Private Sub ClothesMarkAsRepairBTN_Click(sender As Object, e As EventArgs) Handles ClothesMarkAsRepairBTN.Click
        clothesManager.MarkAsRepairClothes()
    End Sub

    ' =========================
    ' ARCHIVE TOGGLE SYSTEM
    ' =========================
    Public Sub ToggleArchiveBTN_Click(sender As Object, e As EventArgs) Handles ToggleArchiveBTN.Click
        clothesManager.ToggleArchiveMode()
    End Sub


    ' =========================
    ' MATERIALS INVENTORY PANEL
    ' =========================

    Public Sub LoadSuppliersToComboBox(comboBox As ComboBox)

        Try

            Dim dt As DataTable = materialsManager.GetSuppliers()

            ' check if no data returned
            If dt Is Nothing Then
                comboBox.DataSource = Nothing
                Return
            End If

            ' check if required columns exist
            If Not dt.Columns.Contains("Supplier_Name") OrElse
           Not dt.Columns.Contains("Supplier_ID") Then

                comboBox.DataSource = Nothing
                Return
            End If

            comboBox.DataSource = dt
            comboBox.DisplayMember = "Supplier_Name"
            comboBox.ValueMember = "Supplier_ID"
            comboBox.SelectedIndex = -1

        Catch ex As Exception

            comboBox.DataSource = Nothing

            Debug.WriteLine("LoadSuppliersToComboBox Error: " & ex.Message)

        End Try

    End Sub

    ' =========================
    ' SUPPLIER PANEL EVENTS
    ' =========================
    Private Sub AddSuppliersBTN_Click(sender As Object, e As EventArgs) Handles AddSuppliersBTN.Click
        materialsManager.AddSuppliers()
    End Sub

    ' =========================
    ' MATERIALS PANEL EVENTS
    ' =========================

    Private Sub AddMaterialBTN_Click(sender As Object, e As EventArgs) Handles AddMaterialBTN.Click
        AddMaterialModalPanel.Visible = True
        AddMaterialModalPanel.BringToFront()
        materialsManager.ClearAddMaterialFields()
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
        materialsManager.ClearAddMaterialFields()
    End Sub

    Private Sub UpdateMaterialModalCancelBTN_Click(sender As Object, e As EventArgs) Handles UpdateMaterialModalCancelBTN.Click
        UpdateMaterialModalPanel.Visible = False
    End Sub

    Private Sub AddMaterialModalConfirmBTN_Click(sender As Object, e As EventArgs) Handles AddMaterialModalConfirmBTN.Click
        materialsManager.AddMaterial()
    End Sub
    Private Sub UpdateMaterialModalConfirmBTN_Click(sender As Object, e As EventArgs) Handles UpdateMaterialModalConfirmBTN.Click
        materialsManager.updateMaterials()
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
    ' SEARCH FUNCTIONALITY 
    ' =========================
    Private Sub MaterialSearchTB_TextChanged(sender As Object, e As EventArgs) Handles MaterialSearchTB.TextChanged
        materialsManager.SearchMaterials(MaterialSearchTB.Text)
    End Sub

    Private Sub SupplierSearchTB_TextChanged(sender As Object, e As EventArgs) Handles SupplierSearchTB.TextChanged
        materialsManager.SearchSuppliers(SupplierSearchTB.Text)
    End Sub

    ' =========================
    ' LOAD DATA WHEN MATERIALS PANEL IS SHOWN
    ' =========================
    Private Sub MaterialsPanel_VisibleChanged(sender As Object, e As EventArgs) Handles MaterialsPanel.VisibleChanged
        If MaterialsPanel.Visible Then
            CheckAndUpdateConnectionStatus()
            If materialsManager Is Nothing Then
                materialsManager = New MaterialsAndSuppliersManager(Me)
            End If
            materialsManager.LoadMaterialsSuppliers()
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
        AutoSelectFirstRow(CustomerDGV)
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
            CheckAndUpdateConnectionStatus()
            LoadCustomers()
        End If
    End Sub



    ' =========================
    ' RENT TRANSACTION PANEL
    ' =========================

    Private Sub RentPanel_VisibleChanged(sender As Object, e As EventArgs) Handles RentPanel.VisibleChanged
        If RentPanel.Visible Then
            CheckAndUpdateConnectionStatus()
            rentManager.LoadRentTransactions()
            rentManager.LoadCustomersToRentCombo()
            rentManager.LoadAvailableClothesToCombo()
        End If
    End Sub

    Private Sub RentItemBTN_Click(sender As Object, e As EventArgs) Handles RentItemBTN.Click
        rentManager.RentItem()
    End Sub


    Private Sub ClearRentItemBTN_Click(sender As Object, e As EventArgs) Handles ClearRentItemBTN.Click
        rentManager.ClearRentFields()
    End Sub


    Private Sub ReturnItemBTN_Click(sender As Object, e As EventArgs) Handles ReturnItemBTN.Click
        rentManager.ReturnItem()
    End Sub


    Private Sub MarkLostItemBTN_Click(sender As Object, e As EventArgs) Handles MarkLostItemBTN.Click
        rentManager.MarkLostItem()
    End Sub

    Private Sub ExtendItemBTN_Click(sender As Object, e As EventArgs) Handles ExtendItemBTN.Click
        rentManager.ExtendItem()
    End Sub

    Private Sub RentSearchTB_TextChanged(sender As Object, e As EventArgs) Handles RentSearchTB.TextChanged
        rentManager.SearchRent(RentSearchTB.Text.Trim())
    End Sub

    ' =========================
    ' TAILORING TRANSACTION PANEL
    ' =========================

    Private Sub TailoringPanel_VisibleChanged(sender As Object, e As EventArgs) Handles TailoringPanel.VisibleChanged
        If TailoringPanel.Visible Then
            CheckAndUpdateConnectionStatus()
            tailoringManager.LoadTransactions()
            tailoringManager.LoadCustomersToCombo()
            tailoringManager.LoadEmployeesToCombo()
            tailoringManager.LoadAvailableClothesToCombo()
            tailoringManager.LoadCatalog()
            tailoringManager.ResetUI()
        End If
    End Sub

    Private Sub CustomerOwnedToggleBTN_Click(sender As Object, e As EventArgs) Handles CustomerOwnedToggleBTN.Click
        tailoringManager.ToggleCustomerOwnedMode()
    End Sub

    Private Sub TailoringTypeCMB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TailoringTypeCMB.SelectedIndexChanged
        tailoringManager.LoadRequiredMaterials()
        tailoringManager.CalculateSmartPrice()
    End Sub

    Private Sub TailoringClothesCMB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TailoringClothesCMB.SelectedIndexChanged
        If Not String.IsNullOrEmpty(TailoringTypeCMB.Text) Then
            tailoringManager.CalculateSmartPrice()
        End If
    End Sub

    Private Sub CreateTailoringRequestBTN_Click(sender As Object, e As EventArgs) Handles CreateTailoringRequestBTN.Click
        tailoringManager.CreateRequest()
    End Sub

    Private Sub UpdateTailoringStatusBTN_Click(sender As Object, e As EventArgs) Handles UpdateTailoringStatusBTN.Click
        tailoringManager.UpdateStatus()
    End Sub

    Private Sub ClearTailoringBTN_Click(sender As Object, e As EventArgs) Handles ClearTailoringBTN.Click
        tailoringManager.ClearFields()
    End Sub

    ' =========================
    ' REPORTS PANEL
    ' =========================

    Private Sub ReportsPanel_VisibleChanged(sender As Object, e As EventArgs) Handles ReportsPanel.VisibleChanged
        If ReportsPanel.Visible Then
            CheckAndUpdateConnectionStatus()
            reportsManager.InitializeReportsPanel()
        End If
    End Sub

    Private Sub GenerateReportBTN_Click(sender As Object, e As EventArgs) Handles GenerateReportBTN.Click
        reportsManager.GenerateSelectedReport()
    End Sub

    Private Sub ExportPDFBTN_Click(sender As Object, e As EventArgs) Handles ExportPDFBTN.Click
        reportsManager.ExportToExcel()
    End Sub

    'Private Sub ExportExcelBTN_Click(sender As Object, e As EventArgs) Handles ExportExcelBTN.Click
    '    reportsManager.ExportToExcel()
    'End Sub

    ' =========================
    ' OFFLINE MODE HANDLER
    ' =========================
    Private Sub UpdateConnectionStatus()
        CheckAndUpdateConnectionStatus()
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

    Private Sub TestConnectionBTN_Click(sender As Object, e As EventArgs) Handles TestConnectionBTN.Click
        ConnDB.serverName = ServerTB.Text.Trim()
        ConnDB.databaseName = DatabaseTB.Text.Trim()
        ConnDB.dbUsername = UsernameTB.Text.Trim()
        ConnDB.dbPassword = PasswordTB.Text

        If ConnDB.TestConnection() Then

            ConfigStatusLabel.Text = "✅ Connection Successful"
            ConfigStatusLabel.ForeColor = Color.Green

        Else

            ConfigStatusLabel.Text = "❌ " & ConnDB.LastConnectionError
            ConfigStatusLabel.ForeColor = Color.Red

        End If
    End Sub

    Private Sub SaveConfigBTN_Click(sender As Object, e As EventArgs) Handles SaveConfigBTN.Click
        ConnDB.serverName = ServerTB.Text.Trim()
        ConnDB.databaseName = DatabaseTB.Text.Trim()
        ConnDB.dbUsername = UsernameTB.Text.Trim()
        ConnDB.dbPassword = PasswordTB.Text

        If ConnDB.TestConnection() Then

            ConnDB.SaveConfig()

            ConfigStatusLabel.Text = "✅ Configuration Saved"
            ConfigStatusLabel.ForeColor = Color.Green

        Else

            ConfigStatusLabel.Text = "❌ Invalid Connection"
            ConfigStatusLabel.ForeColor = Color.Red

        End If
    End Sub

    Private Sub BackupNowBTN_Click(sender As Object, e As EventArgs) Handles BackupNowBTN.Click
        Dim sfd As New SaveFileDialog()

        sfd.Filter = "SQL File|*.sql"
        sfd.FileName =
            $"{databaseName}_Backup_{Date.Now:yyyyMMdd_HHmmss}.sql"

        If sfd.ShowDialog() = DialogResult.OK Then

            If ConnDB.BackupDatabase(sfd.FileName) Then

                MessageBox.Show(
                    "Database backup successful!",
                    "Backup",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                )

            Else

                MessageBox.Show(
                    "Backup failed!",
                    "Backup",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                )

            End If

        End If
    End Sub

    Private Sub RestoreNowBTN_Click(sender As Object, e As EventArgs) Handles RestoreNowBTN.Click
        If MessageBox.Show(
        "This will overwrite current database data. Continue?",
        "WARNING",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Warning
    ) = DialogResult.No Then
            Return
        End If

        Dim ofd As New OpenFileDialog()

        ofd.Filter = "SQL File|*.sql"

        If ofd.ShowDialog() = DialogResult.OK Then

            If ConnDB.RestoreDatabase(ofd.FileName) Then

                MessageBox.Show(
                "Database restored successfully!",
                "Restore",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )

            Else

                MessageBox.Show(
                "Restore failed!",
                "Restore",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )

            End If

        End If

    End Sub

    Public Sub RefreshAllData()

        Try

            If dashboardManager IsNot Nothing Then
                dashboardManager.LoadDashboard()
            End If

            If clothesManager IsNot Nothing Then
                clothesManager.LoadClothes()
            End If

            If materialsManager IsNot Nothing Then
                materialsManager.LoadMaterials()
                materialsManager.LoadSuppliers()
            End If

            If rentManager IsNot Nothing Then
                rentManager.LoadRentTransactions()
                rentManager.LoadCustomersToRentCombo()
                rentManager.LoadAvailableClothesToCombo()
            End If

            If tailoringManager IsNot Nothing Then
                tailoringManager.LoadTransactions()
                tailoringManager.LoadCustomersToCombo()
                tailoringManager.LoadEmployeesToCombo()
                tailoringManager.LoadAvailableClothesToCombo()
                tailoringManager.LoadCatalog()
            End If

        Catch ex As Exception
            Debug.WriteLine("RefreshAllData Error: " & ex.Message)
        End Try

    End Sub

    Public Function EnsureOnline() As Boolean

        If ConnDB.IsOfflineMode Then

            If ConnDB.TestConnection() Then

                UpdateConnectionStatus()
                RefreshAllData()

                MessageBox.Show(
                    "✅ Connection restored!",
                    "Online",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                )

                Return True

            Else

                UpdateConnectionStatus()

                MessageBox.Show(
                    "❌ Database still offline.",
                    "Offline Mode",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )

                Return False

            End If

        End If

        Return True

    End Function


    ' Add this method to check and update connection status
    Public Sub CheckAndUpdateConnectionStatus()
        ' Test connection to get current status
        Dim wasOffline = ConnDB.IsOfflineMode
        Dim isNowOnline = ConnDB.TestConnection()

        ' Update the UI based on current status
        If ConnDB.IsOfflineMode Then
            ConnectionStatusLabel.ForeColor = Color.Red
            ConnectionStatusLabel.Text = "⚠️ OFFLINE MODE - No database connection"

            If Not wasOffline Then
                ' Just went offline
                ClearAllDataGridViews()
                SetOfflineUI()

                ' Show offline in dashboard too
                If dashboardManager IsNot Nothing AndAlso AdminDashboardPanel.Visible Then
                    dashboardManager.LoadDashboard()
                End If
            End If

            hasRefreshedAfterReconnect = False
        Else
            ConnectionStatusLabel.ForeColor = Color.Green
            ConnectionStatusLabel.Text = "🟢 ONLINE - Connected to database"

            If wasOffline Then
                ' Just came online
                SetOnlineUI()

                ' Auto-refresh data when coming online
                If Not hasRefreshedAfterReconnect Then
                    RefreshAllData()
                    hasRefreshedAfterReconnect = True
                End If
            End If
        End If
    End Sub


    Private Sub RefreshCurrentPanelData()
        If ConnDB.IsOfflineMode Then
            ' Try to reconnect
            If ConnDB.TestConnection() Then
                UpdateConnectionStatus()
            Else
                Return
            End If
        End If

        ' Refresh data for the currently visible panel
        If AdminDashboardPanel.Visible Then
            If dashboardManager IsNot Nothing Then
                dashboardManager.LoadDashboard()
            End If
        ElseIf ClothesPanel.Visible Then
            clothesManager.LoadCurrentDGV()
        ElseIf CustomerPanel.Visible Then
            LoadCustomers()
        ElseIf RentPanel.Visible Then
            rentManager.LoadRentTransactions()
            rentManager.LoadCustomersToRentCombo()
            rentManager.LoadAvailableClothesToCombo()
        ElseIf TailoringPanel.Visible Then
            tailoringManager.LoadTransactions()
            tailoringManager.LoadCustomersToCombo()
            tailoringManager.LoadEmployeesToCombo()
            tailoringManager.LoadAvailableClothesToCombo()
            tailoringManager.LoadCatalog()
        ElseIf MaterialsPanel.Visible Then
            materialsManager.LoadMaterials()
            materialsManager.LoadSuppliers()
        ElseIf ReportsPanel.Visible Then
            reportsManager.InitializeReportsPanel()
        End If
    End Sub

    Private Sub AdminDashboardPanel_VisibleChanged(sender As Object, e As EventArgs) Handles AdminDashboardPanel.VisibleChanged
        If AdminDashboardPanel.Visible Then
            If ConnDB.IsOfflineMode Then
                ConnDB.TestConnection()
                UpdateConnectionStatus()
            End If

            If dashboardManager IsNot Nothing Then
                dashboardManager.LoadDashboard()
            End If
        End If
    End Sub

    Private Sub ClothesPanel_VisibleChanged(sender As Object, e As EventArgs) Handles ClothesPanel.VisibleChanged
        If ClothesPanel.Visible Then
            CheckAndUpdateConnectionStatus()  ' Add this line
            clothesManager.LoadCurrentDGV()
        End If
    End Sub
End Class