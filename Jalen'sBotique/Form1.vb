Imports MySql.Data.MySqlClient


Public Class Form1

    Private clothesManager As ClothesInventoryManager

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

    Private ServiceRequirements As New Dictionary(Of String, Dictionary(Of String, Dictionary(Of Integer, List(Of ServiceRequirement))))


    Private Sub InitializeServiceRequirements()
        ServiceRequirements = New Dictionary(Of String, Dictionary(Of String, Dictionary(Of Integer, List(Of ServiceRequirement)))) From {
        {"Pants", New Dictionary(Of String, Dictionary(Of Integer, List(Of ServiceRequirement))) From {
            {"Cotton", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                {1, New List(Of ServiceRequirement) From {  ' Hemming Pants
                    New ServiceRequirement With {.CatalogID = 1, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 2, .UnitMeasure = "rolls"},
                    New ServiceRequirement With {.CatalogID = 1, .MaterialID = 7, .MaterialName = "Cotton Fabric Patch", .DefaultQuantity = 1, .UnitMeasure = "pieces"}
                }},
                {2, New List(Of ServiceRequirement) From {  ' Zipper Replacement (Pants only)
                    New ServiceRequirement With {.CatalogID = 2, .MaterialID = 5, .MaterialName = "Zipper 7 inch Metal", .DefaultQuantity = 1, .UnitMeasure = "pieces"},
                    New ServiceRequirement With {.CatalogID = 2, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 1, .UnitMeasure = "rolls"}
                }}
            }},
            {"Denim", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                {1, New List(Of ServiceRequirement) From {  ' Hemming Pants
                    New ServiceRequirement With {.CatalogID = 1, .MaterialID = 3, .MaterialName = "Heavy Duty Thread", .DefaultQuantity = 3, .UnitMeasure = "rolls"},
                    New ServiceRequirement With {.CatalogID = 1, .MaterialID = 8, .MaterialName = "Denim Patch Material", .DefaultQuantity = 1, .UnitMeasure = "pieces"}
                }},
                {2, New List(Of ServiceRequirement) From {  ' Zipper Replacement
                    New ServiceRequirement With {.CatalogID = 2, .MaterialID = 5, .MaterialName = "Zipper 7 inch Metal", .DefaultQuantity = 1, .UnitMeasure = "pieces"},
                    New ServiceRequirement With {.CatalogID = 2, .MaterialID = 3, .MaterialName = "Heavy Duty Thread", .DefaultQuantity = 2, .UnitMeasure = "rolls"}
                }}
            }},
            {"Gabardine", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                {1, New List(Of ServiceRequirement) From {
                    New ServiceRequirement With {.CatalogID = 1, .MaterialID = 3, .MaterialName = "Heavy Duty Thread", .DefaultQuantity = 3, .UnitMeasure = "rolls"}
                }},
                {2, New List(Of ServiceRequirement) From {
                    New ServiceRequirement With {.CatalogID = 2, .MaterialID = 5, .MaterialName = "Zipper 7 inch Metal", .DefaultQuantity = 1, .UnitMeasure = "pieces"},
                    New ServiceRequirement With {.CatalogID = 2, .MaterialID = 3, .MaterialName = "Heavy Duty Thread", .DefaultQuantity = 2, .UnitMeasure = "rolls"}
                }}
            }}
        }},
        {"Top", New Dictionary(Of String, Dictionary(Of Integer, List(Of ServiceRequirement))) From {
            {"Cotton", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                {3, New List(Of ServiceRequirement) From {  ' Button Replacement (Tops only)
                    New ServiceRequirement With {.CatalogID = 3, .MaterialID = 6, .MaterialName = "Buttons Plastic 15mm", .DefaultQuantity = 4, .UnitMeasure = "pieces"},
                    New ServiceRequirement With {.CatalogID = 3, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 1, .UnitMeasure = "rolls"}
                }},
                {6, New List(Of ServiceRequirement) From {  ' Shorten Sleeves (Tops only)
                    New ServiceRequirement With {.CatalogID = 6, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 2, .UnitMeasure = "rolls"}
                }}
            }},
            {"Silk", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                {3, New List(Of ServiceRequirement) From {
                    New ServiceRequirement With {.CatalogID = 3, .MaterialID = 6, .MaterialName = "Buttons Plastic 15mm", .DefaultQuantity = 4, .UnitMeasure = "pieces"},
                    New ServiceRequirement With {.CatalogID = 3, .MaterialID = 4, .MaterialName = "Silk Thread Fine", .DefaultQuantity = 1, .UnitMeasure = "rolls"}
                }},
                {6, New List(Of ServiceRequirement) From {
                    New ServiceRequirement With {.CatalogID = 6, .MaterialID = 4, .MaterialName = "Silk Thread Fine", .DefaultQuantity = 2, .UnitMeasure = "rolls"},
                    New ServiceRequirement With {.CatalogID = 6, .MaterialID = 10, .MaterialName = "Silk Fabric Patch", .DefaultQuantity = 1, .UnitMeasure = "pieces"}
                }}
            }},
            {"Polyester", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                {3, New List(Of ServiceRequirement) From {
                    New ServiceRequirement With {.CatalogID = 3, .MaterialID = 6, .MaterialName = "Buttons Plastic 15mm", .DefaultQuantity = 4, .UnitMeasure = "pieces"},
                    New ServiceRequirement With {.CatalogID = 3, .MaterialID = 2, .MaterialName = "Polyester Thread Black", .DefaultQuantity = 1, .UnitMeasure = "rolls"}
                }},
                {6, New List(Of ServiceRequirement) From {
                    New ServiceRequirement With {.CatalogID = 6, .MaterialID = 2, .MaterialName = "Polyester Thread Black", .DefaultQuantity = 2, .UnitMeasure = "rolls"}
                }}
            }}
        }},
        {"Dress", New Dictionary(Of String, Dictionary(Of Integer, List(Of ServiceRequirement))) From {
            {"Silk", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                {4, New List(Of ServiceRequirement) From {  ' Taking In Waist (Dress)
                    New ServiceRequirement With {.CatalogID = 4, .MaterialID = 4, .MaterialName = "Silk Thread Fine", .DefaultQuantity = 3, .UnitMeasure = "rolls"}
                }},
                {7, New List(Of ServiceRequirement) From {  ' Lengthen Hem (Dress)
                    New ServiceRequirement With {.CatalogID = 7, .MaterialID = 15, .MaterialName = "Silk Fabric Roll", .DefaultQuantity = 1, .UnitMeasure = "yards"},
                    New ServiceRequirement With {.CatalogID = 7, .MaterialID = 4, .MaterialName = "Silk Thread Fine", .DefaultQuantity = 2, .UnitMeasure = "rolls"}
                }}
            }},
            {"Linen", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                {4, New List(Of ServiceRequirement) From {
                    New ServiceRequirement With {.CatalogID = 4, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 3, .UnitMeasure = "rolls"}
                }},
                {7, New List(Of ServiceRequirement) From {
                    New ServiceRequirement With {.CatalogID = 7, .MaterialID = 17, .MaterialName = "Linen Fabric Roll", .DefaultQuantity = 2, .UnitMeasure = "yards"},
                    New ServiceRequirement With {.CatalogID = 7, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 2, .UnitMeasure = "rolls"}
                }}
            }}
        }},
        {"Suit", New Dictionary(Of String, Dictionary(Of Integer, List(Of ServiceRequirement))) From {
            {"Twill", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                {4, New List(Of ServiceRequirement) From {
                    New ServiceRequirement With {.CatalogID = 4, .MaterialID = 3, .MaterialName = "Heavy Duty Thread", .DefaultQuantity = 3, .UnitMeasure = "rolls"},
                    New ServiceRequirement With {.CatalogID = 4, .MaterialID = 11, .MaterialName = "Interfacing Cloth", .DefaultQuantity = 1, .UnitMeasure = "meters"}
                }},
                {7, New List(Of ServiceRequirement) From {
                    New ServiceRequirement With {.CatalogID = 7, .MaterialID = 16, .MaterialName = "Twill Fabric Roll", .DefaultQuantity = 2, .UnitMeasure = "yards"},
                    New ServiceRequirement With {.CatalogID = 7, .MaterialID = 3, .MaterialName = "Heavy Duty Thread", .DefaultQuantity = 2, .UnitMeasure = "rolls"}
                }}
            }}
        }}
    }
    End Sub


    ' Tailoring mode tracking
    Private isCustomerOwnedMode As Boolean = False

    ' Dragging variables
    Private isDragging As Boolean = False
    Private dragOffset As Point

    Public userRole As String = "admin" ' 👉 set default for testing asd

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        clothesManager = New ClothesInventoryManager(Me)

        ConnDB.TestConnection()
        UpdateConnectionStatus()

        clothesManager.UpdateArchiveModeUI()

        clothesManager.LoadClothes()
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


    End Sub

    ' =========================
    ' ❌ LOGIN DISABLED
    ' =========================
    'Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
    'End Sub

    ' =========================
    ' ROLE-BASED UI
    ' =========================


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
    SuppliersDGV.CellClick, CustomerDGV.CellClick, RentDGV.CellClick, TailoringDGV.CellClick

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
                    LoadRentEditFields(row)

                Case "TailoringDGV"
                    LoadTailoringEditFields(row)
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

    ' Rent Edit Fields
    Private Sub LoadRentEditFields(row As DataGridViewRow)
        ReturnItemRentIDTB.Text = GetCellValue(row.Cells("Rent_ID"))
        ReturnItemCustomerNameTB.Text = GetCellValue(row.Cells("Full_Name"))
        ReturnItemClothesNameTB.Text = GetCellValue(row.Cells("Clothes_Name"))
        ReturnItemStatusTB.Text = GetCellValue(row.Cells("Rental_Status"))
    End Sub

    ' Tailoring Edit Fields
    Private Sub LoadTailoringEditFields(row As DataGridViewRow)
        TailoringServiceIDTB.Text = GetCellValue(row.Cells("Tailoring_Services_ID"))
    End Sub

    ' Clear ALL edit fields
    Private Sub ClearAllEditFields()
        clothesManager.ClearEditFields()
        ClearUpdateMaterialFields()
        ClearSupplierFields()
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

    Private Sub LoadMaterials()
        Dim query As String = "SELECT m.*, s.Supplier_Name 
                          FROM Materials m 
                          LEFT JOIN Supplier s ON m.Supplier_ID = s.Supplier_ID 
                          ORDER BY m.Material_ID DESC"
        LoadToDGV(query, MaterialsDGV)
        AutoSelectFirstRow(MaterialsDGV)
    End Sub

    Private Sub LoadSuppliers()
        Dim query As String = "SELECT * FROM Supplier ORDER BY Supplier_Name"
        LoadToDGV(query, SuppliersDGV)
        AutoSelectFirstRow(SuppliersDGV)
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
    ' DGV CELL CLICK - POPULATE UPDATE FIELDS
    ' =========================



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
        AutoSelectFirstRow(RentDGV)
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
            clothesManager.LoadClothes()

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
            clothesManager.LoadClothes()

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
            clothesManager.LoadClothes()

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
            AutoSelectFirstRow(TailoringDGV)
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
        CalculateSmartPrice()
    End Sub

    Private Sub TailoringClothesCMB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TailoringClothesCMB.SelectedIndexChanged
        If Not String.IsNullOrEmpty(TailoringTypeCMB.Text) Then
            CalculateSmartPrice()
        End If
    End Sub

    Private Sub CalculateSmartPrice()
        If TailoringTypeCMB.SelectedIndex < 0 Then
            TailoringPriceTB.Clear()
            Return
        End If

        Dim serviceName = TailoringTypeCMB.Text
        Dim service = TailoringCatalog.FirstOrDefault(Function(s) s.ServiceName = serviceName)
        If service Is Nothing Then
            TailoringPriceTB.Clear()
            Return
        End If

        TailoringPriceTB.Text = service.BasePrice.ToString("F2")
    End Sub


    Private Sub LoadRequiredMaterials()
        If TailoringTypeCMB.SelectedIndex < 0 Then
            TailoringMaterialsDGV.DataSource = Nothing
            Return
        End If

        Dim serviceName = TailoringTypeCMB.Text
        Dim service = TailoringCatalog.FirstOrDefault(Function(s) s.ServiceName = serviceName)
        If service Is Nothing Then Return

        Dim clothesDetails = GetSelectedClothesDetails()
        Dim requirements As List(Of ServiceRequirement) = New List(Of ServiceRequirement)

        ' 🔍 Priority 1: Category + Fabric specific
        If ServiceRequirements.ContainsKey(clothesDetails.Category) AndAlso
       ServiceRequirements(clothesDetails.Category).ContainsKey(clothesDetails.FabricType) AndAlso
       ServiceRequirements(clothesDetails.Category)(clothesDetails.FabricType).ContainsKey(service.CatalogID) Then

            requirements = ServiceRequirements(clothesDetails.Category)(clothesDetails.FabricType)(service.CatalogID)

            ' 🔍 Priority 2: Category + Any Fabric
        ElseIf ServiceRequirements.ContainsKey(clothesDetails.Category) Then
            Dim categoryDict = ServiceRequirements(clothesDetails.Category)
            Dim fabricDict = categoryDict.Values.FirstOrDefault(Function(d) d.ContainsKey(service.CatalogID))
            If fabricDict IsNot Nothing Then
                requirements = fabricDict(service.CatalogID)
            End If

            ' 🔍 Priority 3: Any Category + Fabric
        Else
            For Each categoryDict In ServiceRequirements.Values
                Dim fabricDict = categoryDict.Values.FirstOrDefault(Function(d) d.ContainsKey(service.CatalogID))
                If fabricDict IsNot Nothing Then
                    requirements = fabricDict(service.CatalogID)
                    Exit For
                End If
            Next
        End If

        If requirements.Count = 0 Then
            TailoringMaterialsDGV.DataSource = Nothing
            TailoringMaterialsDGV.Columns.Clear()

            ' Add columns first, then rows
            TailoringMaterialsDGV.Columns.Add("MaterialID", "ID")
            TailoringMaterialsDGV.Columns.Add("MaterialName", "Material")
            TailoringMaterialsDGV.Columns.Add("Quantity", "Qty")
            TailoringMaterialsDGV.Columns.Add("Unit", "Unit")
            TailoringMaterialsDGV.Columns.Add("Stock", "Stock Status")
            TailoringMaterialsDGV.Columns.Add("Match", "Match")

            ' Now add the warning row
            TailoringMaterialsDGV.Rows.Add("", "⚠️ No materials defined for this service + cloth combination", "", "", "CHECK SERVICE COMPATIBILITY", "")
            Return
        End If

        ' Show materials with stock status (same as before)
        Dim materialsWithStock = New List(Of Object)
        For Each req In requirements
            Dim stockLevel As Integer = 0
            Try
                OpenConn()
                Dim stockQuery = "SELECT Quantity_in_Stock FROM Materials WHERE Material_ID = @id"
                cmd = New MySqlCommand(stockQuery, conn)
                cmd.Parameters.AddWithValue("@id", req.MaterialID)
                Dim result = cmd.ExecuteScalar()
                stockLevel = If(result IsNot Nothing, Convert.ToInt32(result), 0)
                CloseConn()
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

        TailoringMaterialsDGV.Columns.Clear()
        TailoringMaterialsDGV.DataSource = materialsWithStock

        ' Auto-size columns
        TailoringMaterialsDGV.AutoResizeColumns()
        TailoringMaterialsDGV.Columns("Material_ID").Visible = False  ' Hide ID column
    End Sub

    Private Function GetSelectedClothesDetails() As (Category As String, FabricType As String)
        If isCustomerOwnedMode Then
            ' Parse customer description for category hints
            Dim desc = TailoringCustomerClothesTB.Text.ToLower()
            Dim category = "Top" ' Default
            If desc.Contains("pants") Or desc.Contains("slacks") Or desc.Contains("jeans") Then
                category = "Pants"
            ElseIf desc.Contains("dress") Or desc.Contains("terno") Then
                category = "Dress"
            ElseIf desc.Contains("suit") Or desc.Contains("uniform") Then
                category = "Suit"
            End If
            Return (category, "Cotton") ' Default fabric for customer-owned
        End If

        If TailoringClothesCMB.SelectedValue IsNot Nothing Then
            Try
                OpenConn()
                Dim query = "SELECT Category, Fabric_Type FROM Clothes WHERE Clothes_ID = @id"
                cmd = New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", TailoringClothesCMB.SelectedValue)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Dim category = reader("Category").ToString()
                        Dim fabric = If(IsDBNull(reader("Fabric_Type")), "Cotton", reader("Fabric_Type").ToString())
                        CloseConn()
                        Return (category, fabric)
                    End If
                End Using
                CloseConn()
            Catch
                ' Fallback
            End Try
        End If
        Return ("Top", "Cotton") ' Ultimate fallback
    End Function

    ' Create Request
    Private Sub CreateTailoringRequestBTN_Click(sender As Object, e As EventArgs) Handles CreateTailoringRequestBTN.Click
        ' 🆕 VALIDATE SERVICE COMPATIBILITY
        Dim clothesDetails = GetSelectedClothesDetails()
        Dim serviceName = TailoringTypeCMB.Text
        Dim tailoringService = TailoringCatalog.FirstOrDefault(Function(s) s.ServiceName = TailoringTypeCMB.Text)

        ' Block incompatible services
        Dim incompatibleServices = {
            ("Pants", "Shorten Sleeves"), ("Pants", "Button Replacement"),
            ("Top", "Zipper Replacement"), ("Dress", "Button Replacement"),
            ("Suit", "Zipper Replacement")
}

        For Each invalid In incompatibleServices
            If clothesDetails.Category = invalid.Item1 AndAlso serviceName.Contains(invalid.Item2) Then
                MessageBox.Show($"❌ '{serviceName}' not available for {clothesDetails.Category}!", "Service Incompatible")
                Return
            End If
        Next

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
            cmd.Parameters.AddWithValue("@employeeId", TailoringEmployeeCMB.SelectedValue)

            cmd.ExecuteNonQuery()

            ' Update clothes status if shop inventory
            If Not isCustomerOwnedMode AndAlso TailoringClothesCMB.SelectedValue IsNot Nothing Then
                Dim updateClothes = "UPDATE Clothes SET Status = 'In Tailoring' WHERE Clothes_ID = @id"
                cmd = New MySqlCommand(updateClothes, conn)
                cmd.Parameters.AddWithValue("@id", TailoringClothesCMB.SelectedValue)
                cmd.ExecuteNonQuery()
            End If

            ' 🔥 NEW: AUTO-DEDUCT MATERIALS FROM STOCK
            If tailoringService IsNot Nothing Then  ' ✅ Fixed: Use consistent variable name
                Dim requirements = GetCurrentServiceRequirements(tailoringService.CatalogID)  ' ✅ Fixed: Use tailoringService
                Dim serviceId As Integer = Convert.ToInt32(cmd.LastInsertedId)

                For Each req In requirements
                    Try
                        Dim deductQuery = "UPDATE Materials SET Quantity_in_Stock = GREATEST(0, Quantity_in_Stock - @qty) WHERE Material_ID = @materialId"
                        cmd = New MySqlCommand(deductQuery, conn)
                        cmd.Parameters.AddWithValue("@qty", req.DefaultQuantity)
                        cmd.Parameters.AddWithValue("@materialId", req.MaterialID)
                        cmd.ExecuteNonQuery()

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
            clothesManager.LoadClothes()

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            CloseConn()
        End Try
    End Sub

    Private Function GetCurrentServiceRequirements(catalogId As Integer) As List(Of ServiceRequirement)
        Dim clothesDetails = GetSelectedClothesDetails()

        ' Priority 1: Exact Category + Fabric match
        If ServiceRequirements.ContainsKey(clothesDetails.Category) AndAlso
       ServiceRequirements(clothesDetails.Category).ContainsKey(clothesDetails.FabricType) AndAlso
       ServiceRequirements(clothesDetails.Category)(clothesDetails.FabricType).ContainsKey(catalogId) Then
            Return ServiceRequirements(clothesDetails.Category)(clothesDetails.FabricType)(catalogId)
        End If

        ' Priority 2: Category match, any fabric
        If ServiceRequirements.ContainsKey(clothesDetails.Category) Then
            Dim categoryDict = ServiceRequirements(clothesDetails.Category)
            For Each fabricDict In categoryDict.Values
                If fabricDict.ContainsKey(catalogId) Then
                    Return fabricDict(catalogId)
                End If
            Next
        End If

        ' Priority 3: Any category + fabric match
        For Each categoryDict In ServiceRequirements.Values
            If categoryDict.ContainsKey(clothesDetails.FabricType) AndAlso
           categoryDict(clothesDetails.FabricType).ContainsKey(catalogId) Then
                Return categoryDict(clothesDetails.FabricType)(catalogId)
            End If
        Next

        ' Priority 4: Any category, any fabric
        For Each categoryDict In ServiceRequirements.Values
            For Each fabricDict In categoryDict.Values
                If fabricDict.ContainsKey(catalogId) Then
                    Return fabricDict(catalogId)
                End If
            Next
        Next

        Return New List(Of ServiceRequirement)()
    End Function
    Private Sub UpdateTailoringStatusBTN_Click(sender As Object, e As EventArgs) Handles UpdateTailoringStatusBTN.Click
        If TailoringServiceIDTB.Text = "" OrElse String.IsNullOrEmpty(TailoringServiceIDTB.Text) Then
            MessageBox.Show("❌ Select a tailoring transaction first!", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If TailoringStatusCMB.SelectedIndex = -1 OrElse TailoringStatusCMB.SelectedItem Is Nothing Then
            MessageBox.Show("❌ Please select a status from the dropdown!", "No Status Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim serviceId As Integer = Convert.ToInt32(TailoringServiceIDTB.Text.Trim())
        Dim newStatus As String = TailoringStatusCMB.SelectedItem.ToString().Trim()
        Dim fabricType As String = "Cotton"

        ' 🔒 TRANSACTION for safety
        Dim transaction As MySqlTransaction = Nothing
        Try
            OpenConn()
            transaction = conn.BeginTransaction(IsolationLevel.ReadCommitted)

            ' 1️⃣ UPDATE Tailoring Service
            Dim updateServiceQuery As String = "UPDATE Tailoring_Services SET Status = @status, Date_Completed = NOW() WHERE Tailoring_Services_ID = @id"
            Using cmdService = New MySqlCommand(updateServiceQuery, conn, transaction)
                cmdService.Parameters.AddWithValue("@status", newStatus)
                cmdService.Parameters.AddWithValue("@id", serviceId)
                Dim serviceRows = cmdService.ExecuteNonQuery()
                Debug.WriteLine($"📝 Service {serviceId} → '{newStatus}' ({serviceRows} rows)")
            End Using

            ' 2️⃣ 🆕 Get Clothes_ID, Is_Customer_Owned & Service Type
            Dim clothesId As Integer? = Nothing
            Dim isCustomerOwned As Boolean = False
            Dim serviceTypeName As String = ""

            Dim getDetailsQuery As String = "
            SELECT Clothes_ID, Is_Customer_Owned, Type_of_Alteration 
            FROM Tailoring_Services 
            WHERE Tailoring_Services_ID = @id"

            Using cmdDetails = New MySqlCommand(getDetailsQuery, conn, transaction)
                cmdDetails.Parameters.AddWithValue("@id", serviceId)
                Using reader = cmdDetails.ExecuteReader()
                    If reader.Read() Then
                        If Not IsDBNull(reader("Clothes_ID")) Then
                            clothesId = Convert.ToInt32(reader("Clothes_ID"))
                        End If
                        isCustomerOwned = Convert.ToBoolean(reader("Is_Customer_Owned"))
                        serviceTypeName = reader("Type_of_Alteration").ToString()
                    End If
                End Using
            End Using

            ' 3️⃣ HANDLE EACH STATUS
            Dim clothesUpdated As Integer = 0
            Dim materialsRefunded As Integer = 0

            Select Case newStatus.ToUpperInvariant()
                Case "COMPLETED"
                    ' 🎉 Return shop clothes to Available
                    If clothesId.HasValue AndAlso Not isCustomerOwned Then
                        Dim updateClothesQuery As String = "
                        UPDATE Clothes 
                        SET Status = 'Available', 
                            Clothes_Condition = 'Good'
                        WHERE Clothes_ID = @clothesId"
                        Using cmdClothes = New MySqlCommand(updateClothesQuery, conn, transaction)
                            cmdClothes.Parameters.AddWithValue("@clothesId", clothesId.Value)
                            clothesUpdated = cmdClothes.ExecuteNonQuery()
                        End Using
                    End If

                Case "CANCELLED"
                    ' Return shop clothes to Available
                    If clothesId.HasValue AndAlso Not isCustomerOwned Then
                        Dim updateClothesQuery As String = "UPDATE Clothes SET Status = 'Available' WHERE Clothes_ID = @clothesId"
                        Using cmdClothes = New MySqlCommand(updateClothesQuery, conn, transaction)
                            cmdClothes.Parameters.AddWithValue("@clothesId", clothesId.Value)
                            clothesUpdated = cmdClothes.ExecuteNonQuery()
                        End Using
                    End If

                    ' 🔥 FIXED: Fabric-aware material refund
                    Dim service = TailoringCatalog.FirstOrDefault(Function(s) s.ServiceName.Equals(serviceTypeName, StringComparison.OrdinalIgnoreCase))
                    If service IsNot Nothing Then
                        ' Determine fabric type
                        Dim refundFabricType As String = "Cotton"
                        If clothesId.HasValue AndAlso Not isCustomerOwned Then
                            Using cmdFabric = New MySqlCommand("SELECT Fabric_Type FROM Clothes WHERE Clothes_ID = @id", conn, transaction)
                                cmdFabric.Parameters.AddWithValue("@id", clothesId.Value)
                                Dim fabricResult = cmdFabric.ExecuteScalar()
                                If fabricResult IsNot Nothing Then refundFabricType = fabricResult.ToString()
                            End Using
                        End If

                        ' Get fabric-specific requirements or fallback
                        Dim refundRequirements As New List(Of ServiceRequirement)
                        Dim categoryDict As Dictionary(Of String, Dictionary(Of Integer, List(Of ServiceRequirement))) = Nothing
                        Dim fabricDict As Dictionary(Of Integer, List(Of ServiceRequirement)) = Nothing

                        If ServiceRequirements.ContainsKey(refundFabricType) Then
                            categoryDict = ServiceRequirements(refundFabricType)
                            If categoryDict IsNot Nothing AndAlso categoryDict.ContainsKey(service.CatalogID.ToString()) Then
                                fabricDict = categoryDict(service.CatalogID.ToString())
                                If fabricDict IsNot Nothing AndAlso fabricDict.ContainsKey(service.CatalogID) Then
                                    refundRequirements = fabricDict(service.CatalogID)
                                End If
                            End If
                        End If

                        If refundRequirements.Count = 0 Then
                            For Each catDict In ServiceRequirements.Values
                                For Each fabDict In catDict.Values
                                    If fabDict.ContainsKey(service.CatalogID) Then
                                        refundRequirements = fabDict(service.CatalogID)
                                        Exit For
                                    End If
                                Next
                                If refundRequirements.Count > 0 Then Exit For
                            Next
                        End If

                        ' Refund each material
                        For Each req In refundRequirements
                            Dim refundQuery As String = "UPDATE Materials SET Quantity_in_Stock = Quantity_in_Stock + @qty WHERE Material_ID = @materialId"
                            Using cmdRefund = New MySqlCommand(refundQuery, conn, transaction)
                                cmdRefund.Parameters.AddWithValue("@qty", req.DefaultQuantity)
                                cmdRefund.Parameters.AddWithValue("@materialId", req.MaterialID)
                                materialsRefunded += cmdRefund.ExecuteNonQuery()
                            End Using
                            Debug.WriteLine($"💰 Refunded {req.DefaultQuantity} {req.MaterialName} ({refundFabricType})")
                        Next
                    End If

                Case Else
                    ' Pending/In Progress/Ready for Pickup: No special handling
                    Debug.WriteLine($"ℹ️ Status '{newStatus}' - no special clothes/material handling")

            End Select

            ' 4️⃣ COMMIT
            transaction.Commit()

            ' 🎉 DETAILED SUCCESS REPORT
            Dim successMsg As New List(Of String) From {
            $"✅ Service #{serviceId} → '{newStatus}'",
            $"📅 {DateTime.Now:MMM dd, yyyy HH:mm}"
        }

            If clothesUpdated > 0 Then
                successMsg.Insert(0, "👗 Clothes returned to Available inventory!")
            End If

            If materialsRefunded > 0 Then
                successMsg.Insert(0, $"💰 {materialsRefunded} material units refunded!")
            ElseIf newStatus.ToUpperInvariant() = "COMPLETED" AndAlso Not isCustomerOwned AndAlso clothesId.HasValue Then
                successMsg.Insert(0, "🎉 Service completed - clothes ready!")
            ElseIf newStatus.ToUpperInvariant() = "COMPLETED" Then
                successMsg.Insert(0, "✅ Service completed (customer-owned)")
            End If

            MessageBox.Show(String.Join(vbCrLf, successMsg), "✅ Update Successful!",
                       MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' 🔄 FULL REFRESH
            LoadTailoringTransactions()
            clothesManager.LoadClothes()
            If MaterialsPanel.Visible Then LoadMaterials()  ' Refresh stock display
            If TailoringPanel.Visible Then LoadAvailableClothesToTailoringCombo()

        Catch ex As Exception
            transaction?.Rollback()
            MessageBox.Show($"❌ Update failed: {ex.Message}", "Database Error",
                       MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine($"TAILORING ERROR: {ex.ToString()}")
        Finally
            If conn?.State = ConnectionState.Open Then conn.Close()
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