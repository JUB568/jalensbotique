Imports MySql.Data.MySqlClient


Public Class ClothesInventoryManager

    Private form As Form1

    Private _isArchiveMode As Boolean = False

    Public ReadOnly Property IsArchiveMode As Boolean
        Get
            Return _isArchiveMode
        End Get
    End Property
    Public Sub New(frm As Form1)
        form = frm
    End Sub

    ' =========================
    ' CLOTHES INVENTORY PANEL
    ' =========================

    ' LOAD CLOTHES
    Public Sub LoadClothes()
        LoadActiveClothes()
    End Sub

    Public Sub LoadActiveClothes()
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
        LoadToDGV(query, form.ClothesDGV)
        form.AutoSelectFirstRow(form.ClothesDGV)
    End Sub

    Public Sub LoadArchiveClothes()
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
        LoadToDGV(query, form.ArchiveClothesDGV)
        form.AutoSelectFirstRow(form.ArchiveClothesDGV)
    End Sub


    'SEARCH FUNCTION
    Public Sub SearchClothes(searchText As String)

        searchText = searchText.Trim()

        If _isArchiveMode Then

            Dim query As String = "
            SELECT Clothes_ID,
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
              AND (
                    Clothes_Name LIKE @search
                    OR Category LIKE @search
                    OR Color LIKE @search
                  )"

            Try
                OpenConn()

                Dim cmd As New MySqlCommand(query, conn)

                cmd.Parameters.AddWithValue("@search", "%" & searchText & "%")

                Dim da As New MySqlDataAdapter(cmd)

                Dim dt As New DataTable()

                da.Fill(dt)

                form.ArchiveClothesDGV.DataSource = dt

            Catch ex As Exception
                MessageBox.Show(ex.Message)

            Finally
                CloseConn()
            End Try

        Else

            Dim query As String = "
            SELECT *
            FROM Clothes
            WHERE Status != 'Archived'
              AND (
                    Clothes_Name LIKE @search
                    OR Category LIKE @search
                    OR Color LIKE @search
                  )"

            Try
                OpenConn()

                Dim cmd As New MySqlCommand(query, conn)

                cmd.Parameters.AddWithValue("@search", "%" & searchText & "%")

                Dim da As New MySqlDataAdapter(cmd)

                Dim dt As New DataTable()

                da.Fill(dt)

                form.ClothesDGV.DataSource = dt

            Catch ex As Exception
                MessageBox.Show(ex.Message)

            Finally
                CloseConn()
            End Try

        End If

    End Sub

    'ADD FUNCTION
    Public Sub AddClothes()

        Try
            OpenConn()

            Dim query As String = "INSERT INTO Clothes 
        (Clothes_Name, Category, Size, Fabric_Type, Color, Clothes_Condition, Status)
        VALUES (@name, @cat, @size, @fabric, @color, 'Good', 'Available')"

            cmd = New MySqlCommand(query, conn)

            cmd.Parameters.AddWithValue("@name", form.ClothesAddModalClothesNameTB.Text)
            cmd.Parameters.AddWithValue("@cat", form.ClothesAddModalCategoryTB.Text)
            cmd.Parameters.AddWithValue("@size", form.ClothesAddModalSizeTB.Text)
            cmd.Parameters.AddWithValue("@fabric", form.ClothesAddModalFabricTypeTB.Text)
            cmd.Parameters.AddWithValue("@color", form.ClothesAddModalColorTB.Text)

            cmd.ExecuteNonQuery()

            MessageBox.Show("Clothes added successfully!")

            form.ClotheAddModalPanel.Visible = False
            ClearAddFields()
            LoadActiveClothes()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try

    End Sub

    'EDIT FUNCTION
    Public Sub UpdateClothes()
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

            cmd.Parameters.AddWithValue("@id", form.ClothesEditModalClothesIDTB.Text)
            cmd.Parameters.AddWithValue("@name", form.ClothesEditModalClothesNameTB.Text)
            cmd.Parameters.AddWithValue("@cat", form.ClothesEditModalCategoryTB.Text)
            cmd.Parameters.AddWithValue("@size", form.ClothesEditModalSizeTB.Text)
            cmd.Parameters.AddWithValue("@fabric", form.ClothesEditModalFabricTypeTB.Text)
            cmd.Parameters.AddWithValue("@color", form.ClothesEditModalColorTB.Text)
            cmd.Parameters.AddWithValue("@condition", form.ClothesEditModalConditionTB.Text)

            cmd.ExecuteNonQuery()

            MessageBox.Show("Updated successfully!")

            form.ClotheEditModalPanel.Visible = False
            LoadClothes()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try
    End Sub

    'FOR SAFE DELETE(ARCHIVED)
    Public Sub SafeClothesDelete()
        If ConnDB.IsOfflineMode Then
            MessageBox.Show("Cannot modify while offline!", "Offline Mode")
            Return
        End If

        If form.ClothesEditModalClothesIDTB.Text = "" Then
            MessageBox.Show("Select item first!")
            Return
        End If

        Dim actionText As String = If(_isArchiveMode, "RESTORE to Active", "ARCHIVE to Storage")
        Dim newStatus As String = If(_isArchiveMode, "Available", "Archived")

        If MessageBox.Show($"📦 {actionText}?" & vbCrLf & $"Status will be: {newStatus}",
                          If(_isArchiveMode, "Restore Item", "Archive Item"),
                          MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            Try
                OpenConn()

                Dim query As String = "
                    UPDATE Clothes SET 
                        Status = @newStatus,
                        Archived_Date = " & If(newStatus = "Archived", "NOW()", "NULL") & "
                    WHERE Clothes_ID = @id"

                cmd = New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", Convert.ToInt32(form.ClothesEditModalClothesIDTB.Text))
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

    'RESTORE FROM ARCHIVE
    Public Sub RestoreDeletedClothes()
        If ConnDB.IsOfflineMode Then
            MessageBox.Show("Cannot restore while offline!", "Offline Mode", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If form.ClothesEditModalClothesIDTB.Text = "" Then
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
        MessageBoxButtons.YesNo,
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
            cmd.Parameters.AddWithValue("@id", Convert.ToInt32(form.ClothesEditModalClothesIDTB.Text))

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
                If _isArchiveMode Then
                    ToggleArchiveMode()
                End If

            Else
                MessageBox.Show("Item not found in archive or already restored!", "Restore Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            MessageBox.Show("Restore failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            CloseConn()
        End Try
    End Sub

    'MARK AS REPAIR
    Public Sub MarkAsRepairClothes()

        If form.ClothesEditModalClothesIDTB.Text = "" Then
            MessageBox.Show("Select item first")
            Return
        End If

        Try
            OpenConn()

            Dim query As String = "UPDATE Clothes SET Status='For Repair' WHERE Clothes_ID=@id"
            cmd = New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@id", form.ClothesEditModalClothesIDTB.Text)

            cmd.ExecuteNonQuery()

            MessageBox.Show("Marked as For Repair")
            LoadClothes()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConn()
        End Try
    End Sub

    'LOAD SELECTED ROW FIELDS
    Public Sub LoadEditFieldsFromRow(row As DataGridViewRow)
        Try
            With row
                form.ClothesEditModalClothesIDTB.Text = form.GetCellValue(.Cells("Clothes_ID"))
                form.ClothesEditModalClothesNameTB.Text = form.GetCellValue(.Cells("Clothes_Name"))
                form.ClothesEditModalCategoryTB.Text = form.GetCellValue(.Cells("Category"))
                form.ClothesEditModalSizeTB.Text = form.GetCellValue(.Cells("Size"))
                form.ClothesEditModalFabricTypeTB.Text = form.GetCellValue(.Cells("Fabric_Type"))
                form.ClothesEditModalColorTB.Text = form.GetCellValue(.Cells("Color"))
                form.ClothesEditModalConditionTB.Text = form.GetCellValue(.Cells("Clothes_Condition"))
            End With

        Catch ex As Exception
            Debug.WriteLine("LoadEditFieldsFromRow Error: " & ex.Message)
        End Try
    End Sub

    'HANDLES CLEARING FIELDS
    Public Sub ClearAddFields()
        form.ClothesAddModalClothesNameTB.Clear()
        form.ClothesAddModalCategoryTB.Clear()
        form.ClothesAddModalSizeTB.Clear()
        form.ClothesAddModalFabricTypeTB.Clear()
        form.ClothesAddModalColorTB.Clear()
    End Sub

    Public Sub ClearEditFields()
        form.ClothesEditModalClothesIDTB.Clear()
        form.ClothesEditModalClothesNameTB.Clear()
        form.ClothesEditModalCategoryTB.Clear()
        form.ClothesEditModalSizeTB.Clear()
        form.ClothesEditModalFabricTypeTB.Clear()
        form.ClothesEditModalColorTB.Clear()
        form.ClothesEditModalConditionTB.Clear()
    End Sub

    'TOGGLE FUNCTION
    Public Sub ToggleArchiveMode()
        If ConnDB.IsOfflineMode Then
            MessageBox.Show("Cannot switch views while offline!", "Offline Mode")
            Return
        End If

        _isArchiveMode = Not _isArchiveMode
        ToggleDGVVisibility()
        LoadCurrentDGV()
        UpdateToggleUI()
    End Sub

    Public Sub LoadCurrentDGV()
        If _isArchiveMode Then
            LoadArchiveClothes()
        Else
            LoadClothes()
        End If
    End Sub

    Public Sub ToggleDGVVisibility()
        If _isArchiveMode Then
            form.ArchiveClothesDGV.Visible = True
            form.ArchiveClothesDGV.BringToFront()
            form.ClothesDGV.Visible = False
            form.RestoreFromArchiveBTN.Visible = True
            form.ClothesDeleteBTN.Visible = False
        Else
            form.ClothesDGV.Visible = True
            form.ClothesDGV.BringToFront()
            form.ArchiveClothesDGV.Visible = False
            form.RestoreFromArchiveBTN.Visible = False
            form.ClothesDeleteBTN.Visible = True
        End If
    End Sub
    Public Sub UpdateToggleUI()
        If _isArchiveMode Then
            form.ToggleArchiveBTN.Text = "👗 SHOW ACTIVE INVENTORY"
            form.ToggleArchiveBTN.BackColor = Color.OrangeRed
            form.ToggleArchiveBTN.ForeColor = Color.White
            form.ArchiveModeLabel.Text = "📦 ARCHIVE/STORAGE MODE"
            form.ArchiveModeLabel.ForeColor = Color.OrangeRed

        Else
            form.ToggleArchiveBTN.Text = "📦 SHOW ARCHIVE/STORAGE"
            form.ToggleArchiveBTN.BackColor = Color.DodgerBlue
            form.ToggleArchiveBTN.ForeColor = Color.White
            form.ArchiveModeLabel.Text = "👗 ACTIVE INVENTORY MODE"
            form.ArchiveModeLabel.ForeColor = Color.Green

        End If
    End Sub

    Public Sub LoadCurrentView()
        If ConnDB.IsOfflineMode Then
            MessageBox.Show("Cannot load data while offline!", "Offline Mode")
            Return
        End If

        If _isArchiveMode Then
            LoadArchiveClothes()
        Else
            LoadClothes()
        End If
    End Sub


    Public Sub UpdateArchiveModeUI()
        If _isArchiveMode Then
            form.ToggleArchiveBTN.Text = "📦 ACTIVE CLOTHES"
            form.ToggleArchiveBTN.BackColor = Color.Orange
            form.ToggleArchiveBTN.ForeColor = Color.White
            form.ArchiveModeLabel.Text = "📦 ARCHIVE MODE - Storage"
            form.ArchiveModeLabel.ForeColor = Color.Orange

        Else
            form.ToggleArchiveBTN.Text = "📦 ARCHIVE MODE"
            form.ToggleArchiveBTN.BackColor = Color.DodgerBlue
            form.ToggleArchiveBTN.ForeColor = Color.White
            form.ArchiveModeLabel.Text = "👗 ACTIVE INVENTORY"
            form.ArchiveModeLabel.ForeColor = Color.Green

        End If
    End Sub
End Class