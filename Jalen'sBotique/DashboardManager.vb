Imports MySql.Data.MySqlClient

Public Class DashboardManager

    Private form As Form1

    Public Sub New(frm As Form1)
        form = frm
    End Sub

    ' =========================
    ' MAIN DASHBOARD LOAD
    ' =========================

    Public Sub LoadDashboard()
        ' Try to reconnect if offline
        If ConnDB.IsOfflineMode Then
            ConnDB.TestConnection()
        End If

        If ConnDB.IsOfflineMode Then
            ShowOfflineIndicators()
            Return
        End If

        Try
            LoadTotalClothes()
            LoadAvailableClothes()
            LoadRentedClothes()
            LoadRepairsClothes()
            LoadQuickReports()
            LoadLowStockAlert()
            LoadRecentActivity()

        Catch ex As Exception
            Debug.WriteLine($"Error loading dashboard: {ex.Message}")
            ShowOfflineIndicators()
        End Try
    End Sub

    Private Sub ShowOfflineIndicators()
        form.TotalClothesLabel.Text = "0"
        form.AvailableLabel.Text = "0"
        form.RentedLabel.Text = "0"
        form.RepairsLabel.Text = "0"

        If form.QuickReportsListBox IsNot Nothing Then
            form.QuickReportsListBox.Items.Clear()
            form.QuickReportsListBox.Items.Add("⚠️ OFFLINE MODE - Connect to database")
        End If

        If form.LowStockListBox IsNot Nothing Then
            form.LowStockListBox.Items.Clear()
            form.LowStockListBox.Items.Add("⚠️ Cannot load - Database offline")
        End If

        If form.RecentActivityListBox IsNot Nothing Then
            form.RecentActivityListBox.Items.Clear()
            form.RecentActivityListBox.Items.Add("⚠️ Cannot load - Database offline")
        End If
    End Sub

    ' =========================
    ' CARD COUNTS
    ' =========================

    Public Sub LoadTotalClothes()
        Dim total As Integer = 0

        Try
            If ConnDB.IsOfflineMode Then
                form.TotalClothesLabel.Text = "0"
                Return
            End If

            OpenConn()
            If ConnDB.IsOfflineMode Then
                form.TotalClothesLabel.Text = "0"
                Return
            End If

            Dim query As String = "SELECT COUNT(*) FROM clothes"
            Using cmd As New MySqlCommand(query, conn)
                total = Convert.ToInt32(cmd.ExecuteScalar())
            End Using
            form.TotalClothesLabel.Text = total.ToString()

        Catch ex As Exception
            form.TotalClothesLabel.Text = "0"
            Debug.WriteLine($"LoadTotalClothes Error: {ex.Message}")
        Finally
            CloseConn()
        End Try
    End Sub

    Public Sub LoadAvailableClothes()
        Dim available As Integer = 0

        Try
            If ConnDB.IsOfflineMode Then
                form.AvailableLabel.Text = "0"
                Return
            End If

            OpenConn()
            If ConnDB.IsOfflineMode Then
                form.AvailableLabel.Text = "0"
                Return
            End If

            Dim query As String = "SELECT COUNT(*) FROM clothes WHERE Status = 'Available'"
            Using cmd As New MySqlCommand(query, conn)
                available = Convert.ToInt32(cmd.ExecuteScalar())
            End Using
            form.AvailableLabel.Text = available.ToString()

        Catch ex As Exception
            form.AvailableLabel.Text = "0"
            Debug.WriteLine($"LoadAvailableClothes Error: {ex.Message}")
        Finally
            CloseConn()
        End Try
    End Sub

    Public Sub LoadRentedClothes()
        Dim rented As Integer = 0

        Try
            If ConnDB.IsOfflineMode Then
                form.RentedLabel.Text = "0"
                Return
            End If

            OpenConn()
            If ConnDB.IsOfflineMode Then
                form.RentedLabel.Text = "0"
                Return
            End If

            Dim query As String = "SELECT COUNT(*) FROM clothes WHERE Status = 'Rented'"
            Using cmd As New MySqlCommand(query, conn)
                rented = Convert.ToInt32(cmd.ExecuteScalar())
            End Using
            form.RentedLabel.Text = rented.ToString()

        Catch ex As Exception
            form.RentedLabel.Text = "0"
            Debug.WriteLine($"LoadRentedClothes Error: {ex.Message}")
        Finally
            CloseConn()
        End Try
    End Sub

    Public Sub LoadRepairsClothes()
        Dim repairs As Integer = 0

        Try
            If ConnDB.IsOfflineMode Then
                form.RepairsLabel.Text = "0"
                Return
            End If

            OpenConn()
            If ConnDB.IsOfflineMode Then
                form.RepairsLabel.Text = "0"
                Return
            End If

            Dim query As String = "SELECT COUNT(*) FROM clothes WHERE Status = 'For Repair'"
            Using cmd As New MySqlCommand(query, conn)
                repairs = Convert.ToInt32(cmd.ExecuteScalar())
            End Using
            form.RepairsLabel.Text = repairs.ToString()

        Catch ex As Exception
            form.RepairsLabel.Text = "0"
            Debug.WriteLine($"LoadRepairsClothes Error: {ex.Message}")
        Finally
            CloseConn()
        End Try
    End Sub

    ' =========================
    ' QUICK REPORTS
    ' =========================

    Public Sub LoadQuickReports()
        If form.QuickReportsListBox Is Nothing Then Return

        form.QuickReportsListBox.Items.Clear()

        Try
            If ConnDB.IsOfflineMode Then
                form.QuickReportsListBox.Items.Add("⚠️ OFFLINE MODE - Connect to database")
                Return
            End If

            OpenConn()
            If ConnDB.IsOfflineMode Then
                form.QuickReportsListBox.Items.Add("⚠️ Cannot load reports - Database offline")
                Return
            End If

            ' 1. Most rented clothes (top 5)
            Dim topRentedQuery = "
                SELECT cl.Clothes_Name, COUNT(r.Rent_ID) as Rental_Count
                FROM clothes cl
                LEFT JOIN rent r ON cl.Clothes_ID = r.Clothes_ID
                GROUP BY cl.Clothes_ID
                ORDER BY Rental_Count DESC
                LIMIT 3"

            cmd = New MySqlCommand(topRentedQuery, conn)
            cmdRead = cmd.ExecuteReader()
            form.QuickReportsListBox.Items.Add("📊 TOP RENTED CLOTHES:")
            Dim count As Integer = 0
            While cmdRead.Read()
                count += 1
                form.QuickReportsListBox.Items.Add($"   {count}. {cmdRead("Clothes_Name")} ({cmdRead("Rental_Count")} rentals)")
            End While
            cmdRead.Close()

            If count = 0 Then
                form.QuickReportsListBox.Items.Add("   No rental data yet")
            End If

            ' 2. Recent tailoring revenue
            Dim revenueQuery = "
                SELECT COALESCE(SUM(Service_Price), 0) as Total_Revenue
                FROM tailoring_services 
                WHERE Status = 'Completed' 
                AND Date_Requested >= DATE_SUB(NOW(), INTERVAL 30 DAY)"

            cmd = New MySqlCommand(revenueQuery, conn)
            Dim revenue = Convert.ToDecimal(cmd.ExecuteScalar())
            form.QuickReportsListBox.Items.Add($"💰 Last 30 days revenue: ₱{revenue:N2}")

            ' 3. Pending tailoring services
            Dim pendingQuery = "
                SELECT COUNT(*) as Pending_Count
                FROM tailoring_services 
                WHERE Status = 'Pending'"

            cmd = New MySqlCommand(pendingQuery, conn)
            Dim pending = Convert.ToInt32(cmd.ExecuteScalar())
            form.QuickReportsListBox.Items.Add($"✂️ Pending tailoring services: {pending}")

            ' 4. Active rentals count
            Dim activeRentalsQuery = "
                SELECT COUNT(*) as Active_Count
                FROM rent 
                WHERE Rental_Status = 'Rented'"

            cmd = New MySqlCommand(activeRentalsQuery, conn)
            Dim activeRentals = Convert.ToInt32(cmd.ExecuteScalar())
            form.QuickReportsListBox.Items.Add($"📦 Active rentals: {activeRentals}")

            CloseConn()

        Catch ex As Exception
            form.QuickReportsListBox.Items.Clear()
            form.QuickReportsListBox.Items.Add("⚠️ Could not load quick reports")
            Debug.WriteLine($"LoadQuickReports Error: {ex.Message}")
            CloseConn()
        End Try
    End Sub

    ' =========================
    ' LOW STOCK ALERT
    ' =========================

    Public Sub LoadLowStockAlert()
        If form.LowStockListBox Is Nothing Then Return

        form.LowStockListBox.Items.Clear()

        Try
            If ConnDB.IsOfflineMode Then
                form.LowStockListBox.Items.Add("⚠️ Cannot load - Database offline")
                Return
            End If

            OpenConn()
            If ConnDB.IsOfflineMode Then
                form.LowStockListBox.Items.Add("⚠️ Cannot load stock alerts - Database offline")
                Return
            End If

            ' Get materials with low stock (<= 10)
            Dim query = "
                SELECT Material_Name, Quantity_in_Stock, Unit_of_Measure
                FROM materials
                WHERE Quantity_in_Stock <= 10
                ORDER BY Quantity_in_Stock ASC
                LIMIT 5"

            cmd = New MySqlCommand(query, conn)
            cmdRead = cmd.ExecuteReader()

            Dim hasLowStock As Boolean = False
            While cmdRead.Read()
                hasLowStock = True
                Dim name = cmdRead("Material_Name").ToString()
                Dim qty = Convert.ToInt32(cmdRead("Quantity_in_Stock"))
                Dim unit = cmdRead("Unit_of_Measure").ToString()

                If qty <= 0 Then
                    form.LowStockListBox.Items.Add($"❌ {name}: OUT OF STOCK")
                ElseIf qty <= 5 Then
                    form.LowStockListBox.Items.Add($"🔴 {name}: Only {qty} {unit} left (CRITICAL)")
                Else
                    form.LowStockListBox.Items.Add($"⚠️ {name}: {qty} {unit} left")
                End If
            End While
            cmdRead.Close()

            If Not hasLowStock Then
                form.LowStockListBox.Items.Add("✅ All materials have sufficient stock")
            End If

            ' Also check clothes that need repair
            Dim repairQuery = "SELECT COUNT(*) FROM clothes WHERE Status = 'For Repair'"
            cmd = New MySqlCommand(repairQuery, conn)
            Dim repairCount = Convert.ToInt32(cmd.ExecuteScalar())

            If repairCount > 0 Then
                form.LowStockListBox.Items.Add($"🛠️ {repairCount} clothes need repair")
            End If

            CloseConn()

        Catch ex As Exception
            form.LowStockListBox.Items.Clear()
            form.LowStockListBox.Items.Add("⚠️ Could not load stock alerts")
            Debug.WriteLine($"LoadLowStockAlert Error: {ex.Message}")
            CloseConn()
        End Try
    End Sub

    ' =========================
    ' RECENT ACTIVITY - FIXED COLUMN NAMES
    ' =========================

    Public Sub LoadRecentActivity()
        If form.RecentActivityListBox Is Nothing Then Return

        form.RecentActivityListBox.Items.Clear()

        Try
            If ConnDB.IsOfflineMode Then
                form.RecentActivityListBox.Items.Add("⚠️ Cannot load - Database offline")
                Return
            End If

            OpenConn()
            If ConnDB.IsOfflineMode Then
                form.RecentActivityListBox.Items.Add("⚠️ Cannot load activity - Database offline")
                Return
            End If

            ' 1. Recent rentals
            Dim rentalsQuery = "
            SELECT 'RENTAL' as Activity_Type, 
                   CONCAT(c.Full_Name, ' rented ', cl.Clothes_Name) as Description,
                   r.Date_Rented as Activity_Date
            FROM Rent r
            INNER JOIN customer c ON r.Customer_ID = c.Customer_ID
            INNER JOIN Clothes cl ON r.Clothes_ID = cl.Clothes_ID
            ORDER BY r.Date_Rented DESC
            LIMIT 3"

            cmd = New MySqlCommand(rentalsQuery, conn)
            cmdRead = cmd.ExecuteReader()
            While cmdRead.Read()
                Dim activityDate As DateTime
                If DateTime.TryParse(cmdRead("Activity_Date").ToString(), activityDate) Then
                    Dim dateStr = activityDate.ToString("MMM dd, hh:mm tt")
                    form.RecentActivityListBox.Items.Add($"📱 {dateStr} - {cmdRead("Description")}")
                Else
                    form.RecentActivityListBox.Items.Add($"📱 - {cmdRead("Description")}")
                End If
            End While
            cmdRead.Close()

            ' 2. Recent tailoring services
            Dim tailoringQuery = "
            SELECT 'TAILORING' as Activity_Type,
                   CONCAT(IFNULL(c.Full_Name, 'Customer'), ' requested: ', ts.Type_of_Alteration) as Description,
                   ts.Date_Requested as Activity_Date
            FROM Tailoring_Services ts
            LEFT JOIN customer c ON ts.Customer_ID = c.Customer_ID
            WHERE ts.Date_Requested IS NOT NULL
            ORDER BY ts.Date_Requested DESC
            LIMIT 3"

            cmd = New MySqlCommand(tailoringQuery, conn)
            cmdRead = cmd.ExecuteReader()
            While cmdRead.Read()
                Dim activityDate As DateTime
                If DateTime.TryParse(cmdRead("Activity_Date").ToString(), activityDate) Then
                    Dim dateStr = activityDate.ToString("MMM dd, hh:mm tt")
                    form.RecentActivityListBox.Items.Add($"✂️ {dateStr} - {cmdRead("Description")}")
                Else
                    form.RecentActivityListBox.Items.Add($"✂️ - {cmdRead("Description")}")
                End If
            End While
            cmdRead.Close()

            ' 3. New customers (using Date_Added column if exists)
            Dim customersQuery = "
            SELECT 'NEW_CUSTOMER' as Activity_Type,
                   CONCAT('New customer registered: ', Full_Name) as Description,
                   Date_Added as Activity_Date
            FROM customer
            WHERE Date_Added IS NOT NULL
            ORDER BY Date_Added DESC
            LIMIT 2"

            cmd = New MySqlCommand(customersQuery, conn)
            cmdRead = cmd.ExecuteReader()
            While cmdRead.Read()
                Dim activityDate As DateTime
                If DateTime.TryParse(cmdRead("Activity_Date").ToString(), activityDate) Then
                    Dim dateStr = activityDate.ToString("MMM dd, hh:mm tt")
                    form.RecentActivityListBox.Items.Add($"👤 {dateStr} - {cmdRead("Description")}")
                Else
                    form.RecentActivityListBox.Items.Add($"👤 - {cmdRead("Description")}")
                End If
            End While
            cmdRead.Close()

            ' 4. Also show completed tailoring services as activity
            Dim completedQuery = "
            SELECT 'COMPLETED' as Activity_Type,
                   CONCAT('Tailoring completed: ', ts.Type_of_Alteration, ' for ', IFNULL(c.Full_Name, 'customer')) as Description,
                   ts.Date_Completed as Activity_Date
            FROM Tailoring_Services ts
            LEFT JOIN customer c ON ts.Customer_ID = c.Customer_ID
            WHERE ts.Status = 'Completed' AND ts.Date_Completed IS NOT NULL
            ORDER BY ts.Date_Completed DESC
            LIMIT 2"

            cmd = New MySqlCommand(completedQuery, conn)
            cmdRead = cmd.ExecuteReader()
            While cmdRead.Read()
                Dim activityDate As DateTime
                If DateTime.TryParse(cmdRead("Activity_Date").ToString(), activityDate) Then
                    Dim dateStr = activityDate.ToString("MMM dd, hh:mm tt")
                    form.RecentActivityListBox.Items.Add($"✅ {dateStr} - {cmdRead("Description")}")
                Else
                    form.RecentActivityListBox.Items.Add($"✅ - {cmdRead("Description")}")
                End If
            End While
            cmdRead.Close()

            If form.RecentActivityListBox.Items.Count = 0 Then
                form.RecentActivityListBox.Items.Add("No recent activity to display")
            End If

            CloseConn()

        Catch ex As Exception
            form.RecentActivityListBox.Items.Clear()
            form.RecentActivityListBox.Items.Add("⚠️ Could not load recent activity")
            Debug.WriteLine($"LoadRecentActivity Error: {ex.ToString()}")
            CloseConn()
        End Try
    End Sub

    ' =========================
    ' REFRESH DASHBOARD
    ' =========================

    Public Sub RefreshDashboard()
        LoadDashboard()
    End Sub

End Class