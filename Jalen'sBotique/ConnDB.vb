Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports MySql.Data.MySqlClient

Module ConnDB

    Public conn As MySqlConnection
    Public cmd As MySqlCommand
    Public cmdRead As MySqlDataReader

    Public IsOfflineMode As Boolean = False
    Public LastConnectionError As String = ""

    Public serverName As String = "localhost"
    Public databaseName As String = "jalensbotique"
    Public dbUsername As String = "root"
    Public dbPassword As String = ""

    Private configPath As String = Application.StartupPath & "\dbconfig.txt"

    Public Sub SaveConfig()

        Try

            Dim lines As String() = {
            serverName,
            databaseName,
            dbUsername,
            Encrypt(dbPassword)
        }

            File.WriteAllLines(configPath, lines)

        Catch ex As Exception
            MessageBox.Show("Error saving config: " & ex.Message)
        End Try

    End Sub

    Public Sub LoadConfig()

        Try

            If File.Exists(configPath) Then

                Dim lines() As String = File.ReadAllLines(configPath)

                If lines.Length >= 4 Then

                    serverName = lines(0)
                    databaseName = lines(1)
                    dbUsername = lines(2)
                    dbPassword = Decrypt(lines(3))

                End If

            End If

        Catch ex As Exception
            MessageBox.Show("Error loading config: " & ex.Message)
        End Try

    End Sub

    Private Function GetConnectionString() As String
        Return $"server={serverName};uid={dbUsername};password={dbPassword};database={databaseName};AllowUserVariables=True;"
    End Function

    Public Function TestConnection() As Boolean
        Try
            Using testConn As New MySqlConnection(GetConnectionString())
                testConn.Open()
                IsOfflineMode = False
                LastConnectionError = ""
                Return True
            End Using
        Catch ex As Exception
            LastConnectionError = ex.Message
            IsOfflineMode = True
            Return False
        End Try
    End Function

    Public Sub OpenConn()

        Try
            ' If we think we're online but connection is broken, test again
            If Not IsOfflineMode Then
                If conn Is Nothing OrElse conn.State <> ConnectionState.Open Then
                    If Not TestConnection() Then
                        IsOfflineMode = True
                        Return
                    End If
                End If
            End If

            If IsOfflineMode Then
                ' Only try to reconnect if we have config
                If Not String.IsNullOrEmpty(serverName) Then
                    If TestConnection() Then
                        ' We're online now, continue
                    Else
                        Return
                    End If
                Else
                    Return
                End If
            End If

            If conn Is Nothing Then
                conn = New MySqlConnection(GetConnectionString())
            ElseIf conn.ConnectionString <> GetConnectionString() Then
                conn.Close()
                conn.Dispose()
                conn = New MySqlConnection(GetConnectionString())
            End If

            If conn.State = ConnectionState.Closed OrElse
           conn.State = ConnectionState.Broken Then

                conn.Open()
                IsOfflineMode = False
                LastConnectionError = ""

            End If

        Catch ex As Exception
            IsOfflineMode = True
            LastConnectionError = ex.Message
        End Try

    End Sub

    Public Sub CloseConn()
        Try
            If conn IsNot Nothing AndAlso conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        Catch
        End Try
    End Sub

    Public Sub ReadQuery(ByVal sql As String)
        Try
            If IsOfflineMode Then
                If Not TestConnection() Then Return
            End If
            OpenConn()
            If IsOfflineMode Then Return
            cmd = New MySqlCommand(sql, conn)
            cmdRead = cmd.ExecuteReader()
        Catch ex As Exception
            IsOfflineMode = True
            LastConnectionError = ex.Message
        End Try
    End Sub

    Public Function LoadToDGV(ByVal query As String, ByVal dgv As DataGridView) As Integer
        Try
            If IsOfflineMode Then
                ' Try to reconnect once
                If Not TestConnection() Then
                    dgv.DataSource = Nothing
                    dgv.Rows.Clear()
                    Return 0
                End If
            End If

            OpenConn()

            If IsOfflineMode Then
                dgv.DataSource = Nothing
                dgv.Rows.Clear()
                Return 0
            End If

            Dim adapter As New MySqlDataAdapter(query, conn)
            Dim dt As New DataTable
            adapter.Fill(dt)
            dgv.DataSource = dt
            Return dt.Rows.Count

        Catch ex As Exception
            IsOfflineMode = True
            LastConnectionError = ex.Message
            dgv.DataSource = Nothing
            dgv.Rows.Clear()
            Return 0
        Finally
            CloseConn()
        End Try
    End Function

    Public Function BackupDatabase(savePath As String) As Boolean

        Try

            Dim mysqldumpPath As String =
            "D:\xampp\mysql\bin\mysqldump.exe"

            Dim arguments As String =
            $"-h {serverName} -u {dbUsername} "

            If dbPassword <> "" Then
                arguments &= $"-p{dbPassword} "
            End If

            arguments &= $"{databaseName} --result-file=""{savePath}"""

            Dim process As New Process()

            process.StartInfo.FileName = mysqldumpPath
            process.StartInfo.Arguments = arguments
            process.StartInfo.UseShellExecute = False
            process.StartInfo.CreateNoWindow = True

            process.Start()

            process.WaitForExit()

            Return process.ExitCode = 0

        Catch ex As Exception

            MessageBox.Show(
            "Backup Error: " & ex.Message
        )

            Return False

        End Try

    End Function

    Public Function RestoreDatabase(sqlFilePath As String) As Boolean

        Try

            CloseConn()

            Dim mysqlPath As String =
            "D:\xampp\mysql\bin\mysql.exe"

            Dim arguments As String =
            $"-h {serverName} -u {dbUsername} "

            If dbPassword <> "" Then
                arguments &= $"-p{dbPassword} "
            End If

            arguments &= $"{databaseName}"

            Dim process As New Process()

            process.StartInfo.FileName = mysqlPath
            process.StartInfo.Arguments = arguments
            process.StartInfo.UseShellExecute = False
            process.StartInfo.RedirectStandardInput = True
            process.StartInfo.CreateNoWindow = True

            process.Start()

            Dim sql As String = File.ReadAllText(sqlFilePath)

            process.StandardInput.WriteLine(sql)
            process.StandardInput.Close()

            process.WaitForExit()

            Return process.ExitCode = 0

        Catch ex As Exception

            MessageBox.Show(
            "Restore Error: " & ex.Message
        )

            Return False

        End Try

    End Function


    ' =========================
    ' ENCRYPT/DECRYPT (Always available)
    ' =========================
    Public Function Encrypt(ByVal clearText As String) As String
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Dim clearBytes As Byte() = Encoding.UTF8.GetBytes(clearText)

        Using aes As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, Encoding.UTF8.GetBytes("salt123"))
            aes.Key = pdb.GetBytes(32)
            aes.IV = pdb.GetBytes(16)

            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, aes.CreateEncryptor(), CryptoStreamMode.Write)
                    cs.Write(clearBytes, 0, clearBytes.Length)
                End Using
                Return Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
    End Function

    Public Function Decrypt(ByVal cipherText As String) As String
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)

        Using aes As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, Encoding.UTF8.GetBytes("salt123"))
            aes.Key = pdb.GetBytes(32)
            aes.IV = pdb.GetBytes(16)

            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, aes.CreateDecryptor(), CryptoStreamMode.Write)
                    cs.Write(cipherBytes, 0, cipherBytes.Length)
                End Using
                Return Encoding.UTF8.GetString(ms.ToArray())
            End Using
        End Using
    End Function

    Public Sub EnsureDatabaseExists()

        Try

            Dim connStr As String =
                $"server={serverName};uid={dbUsername};password={dbPassword};"

            Using tempConn As New MySqlConnection(connStr)

                tempConn.Open()

                Dim query As String =
                    $"CREATE DATABASE IF NOT EXISTS `{databaseName}`"

                Using cmd As New MySqlCommand(query, tempConn)
                    cmd.ExecuteNonQuery()
                End Using

            End Using

        Catch ex As Exception

            MessageBox.Show(
                "Create DB Error: " & ex.Message
            )

        End Try

    End Sub

End Module