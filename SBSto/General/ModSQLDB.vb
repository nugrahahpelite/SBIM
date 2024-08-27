Imports System.Data.SqlClient
Imports GlobalUtil

Module ModSQLDB
    'Public Const cbuDBMaster = "SB_DATAWH.dbo."

    Public Function fbuConnectSQL(vriSQLConnMst As SqlClient.SqlConnection) As Boolean
        Try
            pbMsgError = ""
            Dim vnDBServer As String
            Dim vnDBName As String
            Dim vnDBUser As String
            Dim vnDBPassword As String

            vnDBServer = EncryptDecrypt.Decrypt(ConfigurationManager.AppSettings("DBServer"), "MyEncryptPassword")
            vnDBName = EncryptDecrypt.Decrypt(ConfigurationManager.AppSettings("DBName"), "MyEncryptPassword")
            vnDBUser = EncryptDecrypt.Decrypt(ConfigurationManager.AppSettings("DBUserName"), "MyEncryptPassword")
            vnDBPassword = EncryptDecrypt.Decrypt(ConfigurationManager.AppSettings("DBUserPassword"), "MyEncryptPassword")

            Dim vnSQLConnString As String
            vnSQLConnString = String.Format("Data Source={0};Initial Catalog={1};Integrated Security=False;User ID={2};Password={3};Connect Timeout=300;", vnDBServer, vnDBName, vnDBUser, vnDBPassword)

            If vnDBName = "Agus_Testing" Then
                vbuDevelopmentStatus = "DB Testing"
                'vnSQLConnString = "Server=mssql-cluster;Database=" & vnDBName & ";Trusted_Connection=True;"
            ElseIf vnDBName = "WMS_Staging" Then
                vbuDevelopmentStatus = "DB STAGING"
            End If

            fbuConnectSQL = False
            vriSQLConnMst.ConnectionString = vnSQLConnString
            vriSQLConnMst.Open()
            fbuConnectSQL = True
        Catch ex As Exception
            fbuConnectSQL = False
            pbMsgError = ex.Message
        End Try
    End Function

    Public Function fbuGetDBMaster() As String
        Dim vnDBMaster As String
        vnDBMaster = ConfigurationManager.AppSettings("DBMaster")
        If vnDBMaster <> "" Then
            vnDBMaster = EncryptDecrypt.Decrypt(vnDBMaster, "MyEncryptPassword")
        End If
        Return vnDBMaster
    End Function

    Public Function fbuGetDBDcm() As String
        Dim vnDBDcm As String
        vnDBDcm = ConfigurationManager.AppSettings("DBDcm")
        If vnDBDcm <> "" Then
            vnDBDcm = EncryptDecrypt.Decrypt(vnDBDcm, "MyEncryptPassword")
        End If
        Return vnDBDcm
    End Function

    Public Function fbuConnectSQLDWH(vriSQLConn As SqlClient.SqlConnection) As Boolean
        Try
            pbMsgError = ""
            Dim vnDBServer As String
            Dim vnDBName As String
            Dim vnDBUser As String
            Dim vnDBPassword As String

            vnDBServer = EncryptDecrypt.Decrypt(ConfigurationManager.AppSettings("DBServerDWH"), "MyEncryptPassword")
            vnDBName = EncryptDecrypt.Decrypt(ConfigurationManager.AppSettings("DBNameDWH"), "MyEncryptPassword")
            vnDBUser = EncryptDecrypt.Decrypt(ConfigurationManager.AppSettings("DBUserNameDWH"), "MyEncryptPassword")
            vnDBPassword = EncryptDecrypt.Decrypt(ConfigurationManager.AppSettings("DBUserPasswordDWH"), "MyEncryptPassword")

            Dim vnSQLConnString As String
            vnSQLConnString = String.Format("Data Source={0};Initial Catalog={1};Integrated Security=False;User ID={2};Password={3};Connect Timeout=300;", vnDBServer, vnDBName, vnDBUser, vnDBPassword)

            fbuConnectSQLDWH = False
            vriSQLConn.ConnectionString = vnSQLConnString
            vriSQLConn.Open()
            fbuConnectSQLDWH = True
        Catch ex As Exception
            fbuConnectSQLDWH = False
            pbMsgError = ex.Message
        End Try
    End Function

    Public Function fbuConnectSQLHris(vriSQLConn As SqlClient.SqlConnection) As Boolean
        Try
            pbMsgError = ""
            Dim vnDBServer As String = EncryptDecrypt.Decrypt(ConfigurationManager.AppSettings("DBServerH"), "MyEncryptPassword")
            Dim vnDBName As String = EncryptDecrypt.Decrypt(ConfigurationManager.AppSettings("DBNameH"), "MyEncryptPassword")
            Dim vnDBUser As String = EncryptDecrypt.Decrypt(ConfigurationManager.AppSettings("DBUserNameH"), "MyEncryptPassword")
            Dim vnDBPassword As String = EncryptDecrypt.Decrypt(ConfigurationManager.AppSettings("DBPasswordH"), "MyEncryptPassword")

            Dim vnSQLConnString As String
            vnSQLConnString = String.Format("Data Source={0};Initial Catalog={1};Integrated Security=False;User ID={2};Password={3};Connect Timeout=300;", vnDBServer, vnDBName, vnDBUser, vnDBPassword)

            fbuConnectSQLHris = False
            vriSQLConn.ConnectionString = vnSQLConnString
            vriSQLConn.Open()
            fbuConnectSQLHris = True
        Catch ex As Exception
            fbuConnectSQLHris = False
            pbMsgError = ex.Message
        End Try
    End Function

    Public Sub pbuCloseSQLConn(vriSQLConn As SqlClient.SqlConnection)
        vriSQLConn.Close()
        vriSQLConn.Dispose()
        vriSQLConn = Nothing
    End Sub

    Public Sub pbuFillDtbSQL(vriDtb As DataTable, vriQuery As String, vriSQLConn As SqlClient.SqlConnection)
        Dim vnSda As New SqlDataAdapter
        vnSda.SelectCommand = New SqlCommand(vriQuery, vriSQLConn)
        vnSda.SelectCommand.CommandTimeout = "150"
        vnSda.Fill(vriDtb)
    End Sub

    Public Sub pbuFillDtbSQLTrans(vriDtb As DataTable, vriQuery As String, vriSQLConn As SqlClient.SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnSda As New SqlDataAdapter
        vnSda.SelectCommand = New SqlCommand(vriQuery, vriSQLConn, vriSQLTrans)
        vnSda.Fill(vriDtb)
    End Sub

    Public Sub pbuExecuteSQLTrans(vriQuery As String, vriStatus As Integer, vriSQLConn As SqlClient.SqlConnection, vriSQLTrans As SqlTransaction)
        Dim vnSQLAdp As New SqlDataAdapter()
        If vriStatus = cbuActionNew Then
            vnSQLAdp.InsertCommand = New SqlCommand(vriQuery, vriSQLConn, vriSQLTrans)
            vnSQLAdp.InsertCommand.CommandTimeout = 1000
            vnSQLAdp.InsertCommand.ExecuteNonQuery()
        ElseIf vriStatus = cbuActionEdit Then
            vnSQLAdp.UpdateCommand = New SqlCommand(vriQuery, vriSQLConn, vriSQLTrans)
            vnSQLAdp.UpdateCommand.CommandTimeout = 1000
            vnSQLAdp.UpdateCommand.ExecuteNonQuery()
        ElseIf vriStatus = cbuActionDel Then
            vnSQLAdp.DeleteCommand = New SqlCommand(vriQuery, vriSQLConn, vriSQLTrans)
            vnSQLAdp.DeleteCommand.CommandTimeout = 1000
            vnSQLAdp.DeleteCommand.ExecuteNonQuery()
        End If
    End Sub

    Public Sub pbuExecuteSQL(vriQuery As String, vriStatus As Integer, vriSQLConn As SqlClient.SqlConnection)
        Dim vnSQLAdp As New SqlDataAdapter()
        If vriStatus = cbuActionNew Then
            vnSQLAdp.InsertCommand = New SqlCommand(vriQuery, vriSQLConn)
            vnSQLAdp.InsertCommand.CommandTimeout = 300
            vnSQLAdp.InsertCommand.ExecuteNonQuery()
        ElseIf vriStatus = cbuActionEdit Then
            vnSQLAdp.UpdateCommand = New SqlCommand(vriQuery, vriSQLConn)
            vnSQLAdp.UpdateCommand.CommandTimeout = 300
            vnSQLAdp.UpdateCommand.ExecuteNonQuery()
        ElseIf vriStatus = cbuActionDel Then
            vnSQLAdp.DeleteCommand = New SqlCommand(vriQuery, vriSQLConn)
            vnSQLAdp.DeleteCommand.CommandTimeout = 300
            vnSQLAdp.DeleteCommand.ExecuteNonQuery()
        End If
    End Sub

    Public Function fbuExecuteSQLTransScalar(vriQuery As String, vriSQLConn As SqlClient.SqlConnection, vriSQLTrans As SqlTransaction) As Integer
        Dim vnSQLCmd As New SqlCommand(vriQuery, vriSQLConn)
        Dim vnSQLAdp As New SqlDataAdapter(vnSQLCmd)
        vnSQLAdp.InsertCommand = New SqlCommand(vriQuery, vriSQLConn, vriSQLTrans)
        fbuExecuteSQLTransScalar = CInt(vnSQLAdp.InsertCommand.ExecuteScalar())
    End Function

    Public Function fbuExecuteSQLTransNonQuery(vriQuery As String, vriSQLConn As SqlClient.SqlConnection, vriSQLTrans As SqlTransaction) As Integer
        Dim vnSQLCmd As New SqlCommand(vriQuery, vriSQLConn)
        Dim vnSQLAdp As New SqlDataAdapter(vnSQLCmd)
        vnSQLAdp.InsertCommand = New SqlCommand(vriQuery, vriSQLConn, vriSQLTrans)
        fbuExecuteSQLTransNonQuery = CInt(vnSQLAdp.InsertCommand.ExecuteNonQuery())
    End Function

    Public Function fbuGetDataNumSQL(vriQuery As String, vriSQLConn As SqlClient.SqlConnection) As Single
        Dim vnSQLCmd As New SqlCommand(vriQuery, vriSQLConn)
        Dim vnSQLAdp As New SqlDataAdapter(vnSQLCmd)
        Dim vnDtb As New DataTable
        vnSQLAdp.Fill(vnDtb)

        If vnDtb.Rows.Count = 0 Then
            fbuGetDataNumSQL = 0
        Else
            If IsDBNull(vnDtb.Rows(0).Item(0)) Then
                fbuGetDataNumSQL = 0
            Else
                fbuGetDataNumSQL = vnDtb.Rows(0).Item(0)
            End If
        End If

        vnSQLAdp.Dispose()
        vnSQLCmd.Dispose()
    End Function

    Public Function fbuGetDataNumSQLArr(vriQuery As String, vriSQLConn As SqlClient.SqlConnection) As Single()
        Dim vnReturn() As Single
        Dim vnSQLCmd As New SqlCommand(vriQuery, vriSQLConn)
        Dim vnSQLAdp As New SqlDataAdapter(vnSQLCmd)
        Dim vnDtb As New DataTable
        vnSQLAdp.Fill(vnDtb)

        ReDim vnReturn(vnDtb.Columns.Count - 1)

        If vnDtb.Rows.Count = 0 Then
            For vn = 0 To vnDtb.Columns.Count - 1
                vnReturn(vn) = 0
            Next
        Else
            For vn = 0 To vnDtb.Columns.Count - 1
                If IsDBNull(vnDtb.Rows(0).Item(vn)) Then
                    vnReturn(vn) = 0
                Else
                    vnReturn(vn) = vnDtb.Rows(0).Item(vn)
                End If
            Next
        End If

        vnSQLAdp.Dispose()
        vnSQLCmd.Dispose()
        Return vnReturn
    End Function

    Public Function fbuGetDataNumSQLTrans(vriQuery As String, vriSQLConn As SqlClient.SqlConnection, vriSQLTrans As SqlTransaction) As Single
        Dim vnSQLCmd As New SqlCommand(vriQuery, vriSQLConn, vriSQLTrans)
        Dim vnSQLAdp As New SqlDataAdapter(vnSQLCmd)
        Dim vnDtb As New DataTable
        vnSQLAdp.Fill(vnDtb)

        If vnDtb.Rows.Count = 0 Then
            fbuGetDataNumSQLTrans = 0
        Else
            If IsDBNull(vnDtb.Rows(0).Item(0)) Then
                fbuGetDataNumSQLTrans = 0
            Else
                fbuGetDataNumSQLTrans = vnDtb.Rows(0).Item(0)
            End If
        End If

        vnSQLAdp.Dispose()
        vnSQLCmd.Dispose()
    End Function

    Public Function fbuGetDataStrSQL(vriQuery As String, vriSQLConn As SqlClient.SqlConnection) As String
        Dim vnSQLCmd As New SqlCommand(vriQuery, vriSQLConn)
        Dim vnSQLAdp As New SqlDataAdapter(vnSQLCmd)
        Dim vnDtb As New DataTable
        vnSQLAdp.Fill(vnDtb)

        If vnDtb.Rows.Count = 0 Then
            fbuGetDataStrSQL = ""
        Else
            If IsDBNull(vnDtb.Rows(0).Item(0)) Then
                fbuGetDataStrSQL = ""
            Else
                fbuGetDataStrSQL = vnDtb.Rows(0).Item(0)
            End If
        End If

        vnSQLAdp.Dispose()
        vnSQLCmd.Dispose()
    End Function

    Public Function fbuGetDataStrSQLArr(vriQuery As String, vriSQLConn As SqlClient.SqlConnection) As String()
        Dim vnReturn() As String
        Dim vnSQLCmd As New SqlCommand(vriQuery, vriSQLConn)
        Dim vnSQLAdp As New SqlDataAdapter(vnSQLCmd)
        Dim vnDtb As New DataTable
        vnSQLAdp.Fill(vnDtb)

        ReDim vnReturn(vnDtb.Columns.Count - 1)

        If vnDtb.Rows.Count = 0 Then
            For vn = 0 To vnDtb.Columns.Count - 1
                vnReturn(vn) = ""
            Next
        Else
            For vn = 0 To vnDtb.Columns.Count - 1
                If IsDBNull(vnDtb.Rows(0).Item(vn)) Then
                    vnReturn(vn) = ""
                Else
                    vnReturn(vn) = vnDtb.Rows(0).Item(vn)
                End If
            Next
        End If

        vnSQLAdp.Dispose()
        vnSQLCmd.Dispose()
        Return vnReturn
    End Function
    Public Function fbuGetDataStrSQLTrans(vriQuery As String, vriSQLConn As SqlClient.SqlConnection, vriSQLTrans As SqlTransaction) As String
        Dim vnSQLCmd As New SqlCommand(vriQuery, vriSQLConn, vriSQLTrans)
        Dim vnSQLAdp As New SqlDataAdapter(vnSQLCmd)
        Dim vnDtb As New DataTable
        vnSQLAdp.Fill(vnDtb)

        If vnDtb.Rows.Count = 0 Then
            fbuGetDataStrSQLTrans = ""
        Else
            If IsDBNull(vnDtb.Rows(0).Item(0)) Then
                fbuGetDataStrSQLTrans = ""
            Else
                fbuGetDataStrSQLTrans = vnDtb.Rows(0).Item(0)
            End If
        End If

        vnSQLAdp.Dispose()
        vnSQLCmd.Dispose()
    End Function

    Public Function fbuGetDateNowSQL(vriSQLConn As SqlClient.SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select convert(varchar(11),getdate(),106)+' '+convert(varchar(8),getdate(),108)"
        Dim vnSQLCmd As New SqlCommand(vnQuery, vriSQLConn)
        Dim vnSQLAdp As New SqlDataAdapter(vnSQLCmd)
        Dim vnDtb As New DataTable
        vnSQLAdp.Fill(vnDtb)

        fbuGetDateNowSQL = vnDtb.Rows(0).Item(0)

        vnSQLAdp.Dispose()
        vnSQLCmd.Dispose()
    End Function

    Public Function fbuGetDateNowSQLTrans(vriSQLConn As SqlClient.SqlConnection, vriSQLTrans As SqlTransaction) As String
        Dim vnQuery As String
        vnQuery = "Select convert(varchar(11),getdate(),106)+' '+convert(varchar(8),getdate(),108)"
        Dim vnSQLCmd As New SqlCommand(vnQuery, vriSQLConn, vriSQLTrans)
        Dim vnSQLAdp As New SqlDataAdapter(vnSQLCmd)
        Dim vnDtb As New DataTable
        vnSQLAdp.Fill(vnDtb)

        fbuGetDateNowSQLTrans = vnDtb.Rows(0).Item(0)

        vnSQLAdp.Dispose()
        vnSQLCmd.Dispose()
    End Function

    Public Function fbuGetDateTodaySQL(vriSQLConn As SqlClient.SqlConnection) As String
        Dim vnQuery As String
        vnQuery = "Select convert(varchar(11),getdate(),106)"
        Dim vnSQLCmd As New SqlCommand(vnQuery, vriSQLConn)
        Dim vnSQLAdp As New SqlDataAdapter(vnSQLCmd)
        Dim vnDtb As New DataTable
        vnSQLAdp.Fill(vnDtb)

        fbuGetDateTodaySQL = vnDtb.Rows(0).Item(0)

        vnSQLAdp.Dispose()
        vnSQLCmd.Dispose()
    End Function
End Module
