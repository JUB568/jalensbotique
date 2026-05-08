Imports System.Data
Imports System.IO
Imports System.Text
Imports MySql.Data.MySqlClient

Public Class ReportsManager
    Private form As Form1

    Public Sub New(frm As Form1)
        form = frm
    End Sub

    ' =========================
    ' REPORTS PANEL INITIALIZATION
    ' =========================

    Public Sub InitializeReportsPanel()
        ' Populate report types
        form.ReportTypeCMB.Items.Clear()
        form.ReportTypeCMB.Items.AddRange({
            "📊 All Reports Summary",
            "👗 Clothes Inventory Report",
            "👥 Customer Report",
            "📦 Materials & Stock Report",
            "📱 Rental Transactions Report",
            "✂️ Tailoring Services Report",
            "🏪 Supplier Report",
            "📋 Archived Clothes Report"
        })
        form.ReportTypeCMB.SelectedIndex = 0

        ' Set default dates
        form.ReportDateFromDTP.Value = DateTime.Now.AddDays(-30)
        form.ReportDateToDTP.Value = DateTime.Now

        form.ReportsDGV.AutoGenerateColumns = False
    End Sub

    ' =========================
    ' MAIN REPORT GENERATION
    ' =========================

    Public Sub GenerateSelectedReport()
        If ConnDB.IsOfflineMode Then
            MessageBox.Show("❌ Cannot generate reports offline!", "Offline Mode")
            Return
        End If

        Dim reportType = form.ReportTypeCMB.SelectedItem.ToString()
        Dim fromDate = form.ReportDateFromDTP.Value.Date
        Dim toDate = form.ReportDateToDTP.Value.Date

        Try
            form.ReportsDGV.DataSource = Nothing
            form.ReportsDGV.Columns.Clear()
            form.ReportsDGV.AutoGenerateColumns = True

            Select Case True
                Case reportType.Contains("Summary")
                    LoadSummaryReport()

                Case reportType.Contains("Archived")
                    LoadArchivedClothesReport()

                Case reportType.Contains("Clothes")
                    LoadClothesReport(fromDate, toDate)

                Case reportType.Contains("Customer")
                    LoadCustomerReport()

                Case reportType.Contains("Materials")
                    LoadMaterialsReport()

                Case reportType.Contains("Rental")
                    LoadRentalReport(fromDate, toDate)

                Case reportType.Contains("Tailoring")
                    LoadTailoringReport(fromDate, toDate)

                Case reportType.Contains("Supplier")
                    LoadSupplierReport()
            End Select

            FormatReportsDGV()

        Catch ex As Exception
            MessageBox.Show($"❌ Report Error: {ex.Message}", "Error")
            Debug.WriteLine($"Report Error: {ex.ToString()}")
        End Try
    End Sub

    ' =========================
    ' DATABASE HELPER
    ' =========================

    Private Function GetReportDataTable(query As String, ParamArray parameters() As Object) As DataTable
        Dim dt As New DataTable()
        Try
            OpenConn()
            Dim cmd As New MySqlCommand(query, conn)

            If parameters IsNot Nothing AndAlso parameters.Length > 0 Then
                For i = 0 To parameters.Length - 1 Step 2
                    If i + 1 < parameters.Length Then
                        cmd.Parameters.AddWithValue(parameters(i).ToString(), parameters(i + 1))
                    End If
                Next
            End If

            Dim adapter As New MySqlDataAdapter(cmd)
            adapter.Fill(dt)

            Debug.WriteLine($"✅ Loaded {dt.Rows.Count} rows for query: {query.Substring(0, Math.Min(100, query.Length))}...")

        Catch ex As Exception
            Debug.WriteLine($"❌ GetReportDataTable Error: {ex.Message}")
            Throw ex
        Finally
            CloseConn()
        End Try
        Return dt
    End Function

    ' =========================
    ' INDIVIDUAL REPORTS
    ' =========================

    ' 📊 1. SUMMARY REPORT
    Private Sub LoadSummaryReport()
        form.ReportTitleLabel.Text = "📊 BUSINESS SUMMARY REPORT"

        Dim summaryData As New DataTable()
        summaryData.Columns.Add("Metric", GetType(String))
        summaryData.Columns.Add("Count", GetType(Integer))
        summaryData.Columns.Add("Amount", GetType(Decimal))

        Try
            OpenConn()

            ' Total Clothes (Active only)
            cmd = New MySqlCommand("SELECT COUNT(*) FROM Clothes WHERE Status != 'Archived'", conn)
            Dim totalClothes = Convert.ToInt32(cmd.ExecuteScalar())

            ' Active Rentals
            cmd = New MySqlCommand("SELECT COUNT(*) FROM Rent WHERE Rental_Status = 'Rented'", conn)
            Dim activeRentals = Convert.ToInt32(cmd.ExecuteScalar())

            ' Completed Tailoring Services
            cmd = New MySqlCommand("SELECT COUNT(*) FROM Tailoring_Services WHERE Status = 'Completed'", conn)
            Dim completedServices = Convert.ToInt32(cmd.ExecuteScalar())

            ' Total Revenue
            cmd = New MySqlCommand("SELECT COALESCE(SUM(Service_Price), 0) FROM Tailoring_Services WHERE Status = 'Completed'", conn)
            Dim totalRevenue = Convert.ToDecimal(cmd.ExecuteScalar())

            ' Low Stock Materials
            cmd = New MySqlCommand("SELECT COUNT(*) FROM Materials WHERE Quantity_in_Stock <= 10", conn)
            Dim lowStock = Convert.ToInt32(cmd.ExecuteScalar())

            ' Total Customers
            cmd = New MySqlCommand("SELECT COUNT(*) FROM customer", conn)
            Dim totalCustomers = Convert.ToInt32(cmd.ExecuteScalar())

            CloseConn()

            ' Add rows
            summaryData.Rows.Add("👗 Total Clothes", totalClothes, 0)
            summaryData.Rows.Add("📦 Active Rentals", activeRentals, 0)
            summaryData.Rows.Add("✂️ Completed Services", completedServices, 0)
            summaryData.Rows.Add("💰 Total Revenue", 0, totalRevenue)
            summaryData.Rows.Add("⚠️ Low Stock Items", lowStock, 0)
            summaryData.Rows.Add("👥 Total Customers", totalCustomers, 0)

        Catch ex As Exception
            CloseConn()
            Throw ex
        End Try

        form.ReportsDGV.DataSource = summaryData
    End Sub

    ' 👗 2. CLOTHES REPORT
    Private Sub LoadClothesReport(fromDate As Date, toDate As Date)
        form.ReportTitleLabel.Text = "👗 CLOTHES INVENTORY REPORT"

        Dim query = "
        SELECT 
            Clothes_ID,
            Clothes_Name,
            Category,
            Size,
            Status,
            Clothes_Condition,
            Date_Added,
            CASE 
                WHEN Status = 'Rented' THEN '🔴 RENTED'
                WHEN Status = 'For Repair' THEN '🛠️ REPAIR'
                WHEN Status = 'Lost' THEN '💥 LOST'
                ELSE CONCAT('✅ ', Status)
            END as Status_Display
        FROM Clothes
        WHERE Date_Added BETWEEN @from AND @to
        ORDER BY Date_Added DESC"

        Dim dt = GetReportDataTable(query, "@from", fromDate, "@to", toDate)
        form.ReportsDGV.DataSource = dt
    End Sub

    ' 👥 3. CUSTOMER REPORT
    Private Sub LoadCustomerReport()
        form.ReportTitleLabel.Text = "👥 CUSTOMER REPORT"

        Dim query = "
        SELECT 
            Customer_ID,
            Full_Name,
            Contact_Number,
            Address,
            (SELECT COUNT(*) FROM Rent r WHERE r.Customer_ID = c.Customer_ID) as Total_Rentals,
            (SELECT COUNT(*) FROM Tailoring_Services ts WHERE ts.Customer_ID = c.Customer_ID) as Total_Services
        FROM customer c
        ORDER BY Customer_ID DESC"

        form.ReportsDGV.DataSource = GetReportDataTable(query)
    End Sub

    ' 📦 4. MATERIALS REPORT
    Private Sub LoadMaterialsReport()
        form.ReportTitleLabel.Text = "📦 MATERIALS & STOCK REPORT"

        Dim query = "
        SELECT 
            m.Material_ID,
            m.Material_Name,
            m.Quantity_in_Stock,
            m.Unit_of_Measure,
            s.Supplier_Name,
            CASE 
                WHEN m.Quantity_in_Stock <= 10 THEN '⚠️ LOW STOCK'
                WHEN m.Quantity_in_Stock = 0 THEN '❌ OUT OF STOCK'
                ELSE '✅ IN STOCK'
            END as Stock_Status
        FROM Materials m
        LEFT JOIN Supplier s ON m.Supplier_ID = s.Supplier_ID
        ORDER BY m.Quantity_in_Stock ASC"

        form.ReportsDGV.DataSource = GetReportDataTable(query)
    End Sub

    ' 📱 5. RENTAL REPORT
    Private Sub LoadRentalReport(fromDate As Date, toDate As Date)
        form.ReportTitleLabel.Text = "📱 RENTAL TRANSACTIONS REPORT"

        Dim query = "
        SELECT 
            r.Rent_ID,
            c.Full_Name,
            cl.Clothes_Name,
            r.Date_Rented,
            r.Expected_Return_Date,
            r.Rental_Status,
            CASE 
                WHEN r.Rental_Status = 'Rented' THEN '🔴 ACTIVE'
                WHEN r.Rental_Status = 'Returned' THEN '✅ RETURNED'
                WHEN r.Rental_Status = 'Lost' THEN '💥 LOST'
                ELSE r.Rental_Status
            END as Status_Display
        FROM Rent r
        INNER JOIN customer c ON r.Customer_ID = c.Customer_ID
        INNER JOIN Clothes cl ON r.Clothes_ID = cl.Clothes_ID
        WHERE r.Date_Rented BETWEEN @from AND @to
        ORDER BY r.Date_Rented DESC"

        Dim dt = GetReportDataTable(query, "@from", fromDate, "@to", toDate)
        form.ReportsDGV.DataSource = dt
    End Sub

    ' ✂️ 6. TAILORING REPORT
    Private Sub LoadTailoringReport(fromDate As Date, toDate As Date)
        form.ReportTitleLabel.Text = "✂️ TAILORING SERVICES REPORT"

        Dim query = "
        SELECT 
            Tailoring_Services_ID,
            Type_of_Alteration,
            c.Full_Name,
            Service_Price,
            Status,
            Date_Requested,
            CASE 
                WHEN Is_Customer_Owned = 1 THEN '👕 CUSTOMER OWNED'
                ELSE '👗 SHOP INVENTORY'
            END as Source_Type,
            CASE 
                WHEN Status = 'Completed' THEN '✅ COMPLETED'
                WHEN Status = 'Pending' THEN '⏳ PENDING'
                WHEN Status = 'Cancelled' THEN '❌ CANCELLED'
                ELSE Status
            END as Status_Display
        FROM Tailoring_Services ts
        LEFT JOIN customer c ON ts.Customer_ID = c.Customer_ID
        WHERE Date_Requested BETWEEN @from AND @to
        ORDER BY Date_Requested DESC"

        Dim dt = GetReportDataTable(query, "@from", fromDate, "@to", toDate)
        form.ReportsDGV.DataSource = dt
    End Sub

    ' 🏪 7. SUPPLIER REPORT
    Private Sub LoadSupplierReport()
        form.ReportTitleLabel.Text = "🏪 SUPPLIER REPORT"

        Dim query = "
        SELECT 
            Supplier_ID,
            Supplier_Name,
            Contact_Number,
            Address,
            Email,
            (SELECT COUNT(*) FROM Materials m WHERE m.Supplier_ID = s.Supplier_ID) as Total_Materials
        FROM Supplier s
        ORDER BY Supplier_Name"

        form.ReportsDGV.DataSource = GetReportDataTable(query)
    End Sub

    ' 📋 8. ARCHIVED CLOTHES REPORT
    Private Sub LoadArchivedClothesReport()
        form.ReportTitleLabel.Text = "📋 ARCHIVED CLOTHES REPORT"

        Dim query = "
        SELECT 
            Clothes_ID,
            Clothes_Name,
            Category,
            Archived_Date,
            DATEDIFF(NOW(), Archived_Date) as Days_Archived,
            CASE 
                WHEN DATEDIFF(NOW(), Archived_Date) > 90 THEN '🗑️ LONG TERM (>90 days)'
                WHEN DATEDIFF(NOW(), Archived_Date) > 30 THEN '📦 MEDIUM TERM (30-90 days)'
                ELSE '📦 RECENT (<30 days)'
            END as Archive_Status
        FROM Clothes
        WHERE Status = 'Archived'
        ORDER BY Archived_Date DESC"

        form.ReportsDGV.DataSource = GetReportDataTable(query)
    End Sub

    ' =========================
    ' FORMATTING HELPERS
    ' =========================

    Private Sub FormatReportsDGV()
        Try
            For Each col As DataGridViewColumn In form.ReportsDGV.Columns
                Select Case col.Name.ToUpper()
                    Case "SERVICE_PRICE", "AMOUNT"
                        col.DefaultCellStyle.Format = "C2"
                    Case "DATE_ADDED", "DATE_RENTED", "DATE_REQUESTED", "ARCHIVED_DATE"
                        col.DefaultCellStyle.Format = "MMM dd, yyyy"
                    Case "QUANTITY_IN_STOCK"
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                End Select
            Next
            form.ReportsDGV.AutoResizeColumns()
        Catch ex As Exception
            ' Ignore formatting errors
        End Try
    End Sub

    ' =========================
    ' EXPORT TO EXCEL (CSV format - opens in Excel)
    ' =========================

    Public Sub ExportToExcel()
        If form.ReportsDGV.Rows.Count = 0 Then
            MessageBox.Show("Please generate a report first!", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Excel CSV File (*.csv)|*.csv|Excel XLSX File (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        saveFileDialog.FileName = GenerateFileName()
        saveFileDialog.DefaultExt = "csv"
        saveFileDialog.AddExtension = True

        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Try
                If saveFileDialog.FilterIndex = 1 OrElse saveFileDialog.FileName.ToLower().EndsWith(".csv") Then
                    ' Export as CSV
                    ExportToCSV(saveFileDialog.FileName)
                Else
                    ' Export as XLSX
                    ExportToXLSX(saveFileDialog.FileName)
                End If
            Catch ex As Exception
                MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Function GenerateFileName() As String
        Dim reportType = form.ReportTypeCMB.SelectedItem.ToString()
        ' Remove emojis and special characters for filename
        reportType = System.Text.RegularExpressions.Regex.Replace(reportType, "[^\w\s]", "")
        reportType = reportType.Replace(" ", "_")
        Return $"{reportType}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
    End Function

    Private Sub ExportToCSV(filePath As String)
        Using writer As New StreamWriter(filePath, False, Encoding.UTF8)
            ' Write headers
            Dim headers As New List(Of String)
            For Each col As DataGridViewColumn In form.ReportsDGV.Columns
                headers.Add(EscapeCSVValue(col.HeaderText))
            Next
            writer.WriteLine(String.Join(",", headers))

            ' Write data rows
            For Each row As DataGridViewRow In form.ReportsDGV.Rows
                If Not row.IsNewRow Then
                    Dim rowValues As New List(Of String)
                    For Each cell As DataGridViewCell In row.Cells
                        rowValues.Add(EscapeCSVValue(GetCellDisplayValue(cell)))
                    Next
                    writer.WriteLine(String.Join(",", rowValues))
                End If
            Next
        End Using

        MessageBox.Show($"✅ Report exported successfully to:{vbCrLf}{filePath}{vbCrLf}{vbCrLf}Total rows: {form.ReportsDGV.Rows.Count}",
                       "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ExportToXLSX(filePath As String)
        ' For XLSX export, we'll create a simple HTML table that Excel can open
        ' Or you can add EPPlus NuGet package for proper XLSX generation

        Dim htmlContent As New System.Text.StringBuilder()
        htmlContent.AppendLine("<html>")
        htmlContent.AppendLine("<head>")
        htmlContent.AppendLine("<meta charset='UTF-8'>")
        htmlContent.AppendLine($"<title>{form.ReportTitleLabel.Text}</title>")
        htmlContent.AppendLine("<style>")
        htmlContent.AppendLine("th { background-color: #4CAF50; color: white; padding: 8px; }")
        htmlContent.AppendLine("td { padding: 6px; border: 1px solid #ddd; }")
        htmlContent.AppendLine("table { border-collapse: collapse; width: 100%; }")
        htmlContent.AppendLine("tr:nth-child(even) { background-color: #f2f2f2; }")
        htmlContent.AppendLine("</style>")
        htmlContent.AppendLine("</head>")
        htmlContent.AppendLine("<body>")
        htmlContent.AppendLine($"<h1>{form.ReportTitleLabel.Text}</h1>")
        htmlContent.AppendLine($"<p><strong>Generated:</strong> {DateTime.Now}</p>")
        htmlContent.AppendLine($"<p><strong>Date Range:</strong> {form.ReportDateFromDTP.Value:MMM dd, yyyy} - {form.ReportDateToDTP.Value:MMM dd, yyyy}</p>")
        htmlContent.AppendLine("<table border='1'>")

        ' Write headers
        htmlContent.AppendLine("<tr>")
        For Each col As DataGridViewColumn In form.ReportsDGV.Columns
            htmlContent.AppendLine($"<th>{EscapeHTMLValue(col.HeaderText)}</th>")
        Next
        htmlContent.AppendLine("</tr>")

        ' Write data rows
        For Each row As DataGridViewRow In form.ReportsDGV.Rows
            If Not row.IsNewRow Then
                htmlContent.AppendLine("<tr>")
                For Each cell As DataGridViewCell In row.Cells
                    htmlContent.AppendLine($"<td>{EscapeHTMLValue(GetCellDisplayValue(cell))}</td>")
                Next
                htmlContent.AppendLine("</tr>")
            End If
        Next

        htmlContent.AppendLine("</table>")
        htmlContent.AppendLine($"<p><strong>Total Records:</strong> {form.ReportsDGV.Rows.Count}</p>")
        htmlContent.AppendLine("</body>")
        htmlContent.AppendLine("</html>")

        File.WriteAllText(filePath, htmlContent.ToString(), Encoding.UTF8)

        MessageBox.Show($"✅ Report exported successfully to:{vbCrLf}{filePath}{vbCrLf}{vbCrLf}Total rows: {form.ReportsDGV.Rows.Count}{vbCrLf}{vbCrLf}Note: File will open in Excel as HTML table.",
                       "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Function GetCellDisplayValue(cell As DataGridViewCell) As String
        If cell.Value Is Nothing Then
            Return ""
        End If

        ' Handle numeric formatting
        If TypeOf cell.Value Is Decimal OrElse TypeOf cell.Value Is Double OrElse TypeOf cell.Value Is Single Then
            Return Convert.ToDecimal(cell.Value).ToString("N2")
        ElseIf TypeOf cell.Value Is DateTime Then
            Return Convert.ToDateTime(cell.Value).ToString("yyyy-MM-dd HH:mm:ss")
        Else
            Return cell.Value.ToString()
        End If
    End Function

    Private Function EscapeCSVValue(value As String) As String
        If String.IsNullOrEmpty(value) Then
            Return ""
        End If

        ' Remove any existing quotes and replace with double quotes for CSV
        value = value.Replace("""", """""")

        ' If value contains comma, newline, or double quote, wrap in quotes
        If value.Contains(",") OrElse value.Contains(vbCrLf) OrElse value.Contains("""") Then
            Return $"""{value}"""
        End If

        Return value
    End Function

    Private Function EscapeHTMLValue(value As String) As String
        If String.IsNullOrEmpty(value) Then
            Return ""
        End If

        ' Escape HTML special characters
        value = value.Replace("&", "&amp;")
        value = value.Replace("<", "&lt;")
        value = value.Replace(">", "&gt;")
        value = value.Replace("""", "&quot;")
        value = value.Replace("'", "&#39;")

        Return value
    End Function
End Class