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

    Private serverName As String = "localhost"
    Private databaseName As String = "jalensbotique"
    Private dbUsername As String = "root"
    Private dbPassword As String = ""

    Private Function GetConnectionString() As String
        Return $"server={serverName};uid={dbUsername};password={dbPassword};database={databaseName};AllowUserVariables=True;"
    End Function

    Public Function TestConnection() As Boolean
        Try
            Using testConn As New MySqlConnection(GetConnectionString())
                testConn.Open()
                IsOfflineMode = False
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
            If IsOfflineMode Then Return

            If conn Is Nothing Then
                conn = New MySqlConnection(GetConnectionString())
            End If

            If conn.State = ConnectionState.Closed OrElse conn.State = ConnectionState.Broken Then
                conn.Open()
                IsOfflineMode = False
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
            If IsOfflineMode Then Return
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
                dgv.DataSource = Nothing
                dgv.Rows.Clear()
                Return 0
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

End Module