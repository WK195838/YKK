Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient

Namespace Utility

    Public Class Common

        ''' <summary>
        ''' 檢查資料是否為DBNull型態並以字元替代
        ''' </summary>
        ''' <param name="obj">要檢查的物件</param>
        ''' <param name="replaceString">如果是DBNull則要替代的值</param>
        ''' <returns>回傳原有值或替代值</returns>
        ''' <remarks>通常用在資料繫結時的檢查動作</remarks>
        Public Function ReplaceDBnull(ByVal obj As Object, ByVal replaceString As String) As String
            If Not IsDBNull(obj) Then
                Return obj
            Else
                Return replaceString
            End If
        End Function

        ''' <summary>
        ''' 取得web.config中的ConnectionString值
        ''' </summary>
        ''' <param name="key">key</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetConnectionString(ByVal key As String) As String
            Dim ConnectionString As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings(key).ToString
            Return ConnectionString
        End Function


        Public Function GetAppSetting(ByVal key As String) As String
            Dim appString As String = System.Web.Configuration.WebConfigurationManager.AppSettings(key).ToString
            Return appString
        End Function
    End Class

    Public Class JScript

        ''' <summary>
        ''' 彈出訊息視窗
        ''' </summary>
        ''' <param name="p">Me</param>
        ''' <param name="message">訊息</param>
        ''' <remarks></remarks>
        Public Sub PopMsg(ByRef p As Page, ByVal message As String)
            p.ClientScript.RegisterClientScriptBlock(GetType(String), "PopMsg", GetJScript(message), True)
        End Sub

        ''' <summary>
        ''' 彈出訊息視窗並導向
        ''' </summary>
        ''' <param name="p">Me</param>
        ''' <param name="message">訊息</param>
        ''' <param name="Url">導向連結</param>
        ''' <remarks></remarks>
        Public Sub PopMsg(ByRef p As Page, ByVal message As String, ByVal Url As String)
            p.ClientScript.RegisterClientScriptBlock(GetType(String), "PopMsg", GetJScript(message) & "location.href('" & Url & "');", True)
        End Sub

        ''' <summary>
        ''' 傳回JScript
        ''' </summary>
        ''' <param name="message">訊息</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetJScript(ByVal message As String) As String
            Dim msg As String
            msg = "alert('" & message & "');"
            Return msg
        End Function

        ''' <summary>
        ''' 註冊JavaScript
        ''' </summary>
        ''' <param name="p"></param>
        ''' <param name="key"></param>
        ''' <param name="js"></param>
        ''' <remarks></remarks>
        Public Sub RegJavaScript(ByVal p As Page, ByVal key As String, ByVal js As String)
            p.ClientScript.RegisterClientScriptBlock(GetType(String), key, js, True)
        End Sub

        '彈出視窗傳回值
        Public Function ReturnValue(ByVal controlId As String(), ByVal attribute As String(), ByVal value As String()) As String
            Dim script As String = ""

            For i As Integer = 0 To controlId.Length - 1
                script &= "window.opener.document.all." & controlId(i) & "." & attribute(i) & "='" & value(i) & "';"
            Next
            script &= "window.close();"
            Return script
        End Function


    End Class

    Public Class DataBase

        Private _defaultConnection As SqlConnection

        Dim UtilityCommon As New Utility.Common


        ''' <summary>
        ''' 取資枓的方式
        ''' </summary>
        ''' <remarks></remarks>
        Enum SelectType
            Rows '列
            Columns '欄
        End Enum

        Friend Function GetConnection(ByVal connection As SqlConnection) As SqlConnection
            If connection Is Nothing Then
                connection = New SqlConnection(DefaultConnection.ConnectionString)
            End If
            Return connection
        End Function

        ''' <summary>
        ''' 預設的連線字串,若要設定網站預設可在Application_Start時設定或將連線字串設為預設的DBConnection
        ''' </summary>
        ''' <value>SqlConnection</value>
        ''' <returns>SqlConnection</returns>
        ''' <remarks></remarks>
        Public Property DefaultConnection() As SqlClient.SqlConnection
            Get
                If _defaultConnection Is Nothing Then
                    Throw New Exception("預設連線或指定連線不存在")

                End If
                Return _defaultConnection
            End Get
            Set(ByVal value As SqlClient.SqlConnection)
                _defaultConnection = value
            End Set
        End Property


        '傳回單一結果
        Friend Function ExcuteScalar(ByVal cmd As SqlCommand, ByVal connection As SqlConnection) As String

            cmd.Connection = connection
            Dim re As Object
            Try
                connection.Open()
                re = cmd.ExecuteScalar()
            Catch ex As Exception
                Throw ex
            Finally
                If connection.State = ConnectionState.Open Then
                    connection.Close()
                    connection.Dispose()
                End If
                cmd.Dispose()
            End Try


            Return UtilityCommon.ReplaceDBnull(re, "")

        End Function

        ''' <summary>
        ''' 傳回DataTable
        ''' </summary>
        ''' <param name="sqlQuery">sqlQuery</param>
        ''' <param name="connection">SqlConnection ,預設為DefaultConnection</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDataTable(ByVal sqlQuery As String, Optional ByVal connection As SqlConnection = Nothing) As Data.DataTable
            connection = GetConnection(connection)
            Dim da As New SqlDataAdapter(sqlQuery, connection)
            Dim ds As New Data.DataSet
            da.Fill(ds)
            Return ds.Tables(0)
        End Function

        ''' <summary>
        ''' 傳回DataTable
        ''' </summary>
        ''' <param name="cmd">SqlCommand</param>
        ''' <param name="connection">SqlConnection ,預設為DefaultConnection</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDataTable(ByVal cmd As SqlCommand, Optional ByVal connection As SqlConnection = Nothing) As Data.DataTable

            connection = GetConnection(connection)

            Dim _dt As New System.Data.DataTable()
            Dim _reader As SqlDataReader = Nothing
            Try

                connection.Open()
                cmd.Connection = connection
                _reader = cmd.ExecuteReader

                Dim _table As DataTable = _reader.GetSchemaTable
                Dim _dc As DataColumn
                Dim _row As DataRow
                Dim _al As New ArrayList

                For i As Integer = 0 To _table.Rows.Count - 1
                    _dc = New DataColumn
                    If Not _dt.Columns.Contains(_table.Rows(i)("ColumnName")) Then
                        _dc.ColumnName = _table.Rows(i)("ColumnName").ToString()
                        _dc.Unique = Convert.ToBoolean(_table.Rows(i)("IsUnique"))
                        _dc.AllowDBNull = Convert.ToBoolean(_table.Rows(i)("AllowDBNull"))
                        _dc.ReadOnly = Convert.ToBoolean(_table.Rows(i)("IsReadOnly"))
                        _al.Add(_dc.ColumnName)
                        _dt.Columns.Add(_dc)
                    End If
                Next
                While _reader.Read
                    _row = _dt.NewRow()
                    For i As Integer = 0 To _al.Count - 1
                        _row.Item(_al(i)) = _reader.Item(_al(i))
                    Next
                    _dt.Rows.Add(_row)
                End While


            Catch ex As Exception
                Throw ex
            Finally
                If connection.State = ConnectionState.Open Then
                    connection.Close()
                    connection.Dispose()
                End If
                If Not _reader Is Nothing Then
                    If Not _reader.IsClosed Then
                        _reader.Close()
                    End If
                End If

            End Try
            Return _dt

        End Function

        ''' <summary>
        ''' 判斷是否有重覆值
        ''' </summary>
        ''' <param name="tableName">資料表名稱</param>
        ''' <param name="conditional">條件,ex:(xx_id = no)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function CheckDuplicateColumn(ByVal tableName As String, ByVal conditional As String) As Boolean

            If String.IsNullOrEmpty(tableName) Then
                Throw New Exception("無指定資料表")
            End If
            If String.IsNullOrEmpty(conditional) Then
                Throw New Exception("無指定條件")
            End If

            Try
                If SelectQuery("Count(*)", tableName, conditional) <> "0" Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Throw ex
            End Try

        End Function

        ''' <summary>
        ''' 查詢資料表,傳回某值(String)
        ''' </summary>
        ''' <param name="columnName">欄位名稱</param>
        ''' <param name="tableName">資料表名稱</param>
        ''' <param name="conditional">條件,ex:(xx_id = no)</param>
        ''' <param name="connection">SqlConnection ,預設為DefaultConnection</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SelectQuery(ByVal columnName As String, ByVal tableName As String, ByVal conditional As String, Optional ByVal connection As SqlConnection = Nothing) As String

            If String.IsNullOrEmpty(columnName) Then
                Throw New Exception("無指定欄位")
            End If

            If String.IsNullOrEmpty(tableName) Then
                Throw New Exception("無指定資料表")
            End If

            If String.IsNullOrEmpty(conditional) Then
                Throw New Exception("無指定條件")
            End If



            Dim sqlQuery As String = "Select " & columnName & " From " & tableName & " Where " & conditional
            Dim cmd As New SqlCommand(sqlQuery)
            connection = GetConnection(connection)
            Return ExcuteScalar(cmd, connection)

        End Function

        ''' <summary>
        ''' 查詢資料表,傳回某值(String)
        ''' </summary>
        ''' <param name="sqlQuery">Sql Query</param>
        ''' <param name="connection">SqlConnection ,預設為DefaultConnection</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SelectQuery(ByVal sqlQuery As String, Optional ByVal connection As SqlConnection = Nothing) As String

            If String.IsNullOrEmpty(sqlQuery) Then
                Throw New Exception("無查詢指令")
            End If

            Dim cmd As New SqlCommand(sqlQuery)
            connection = GetConnection(connection)
            Return ExcuteScalar(cmd, connection)

        End Function

        ''' <summary>
        ''' 查詢資料表,傳回某值(String)
        ''' </summary>
        ''' <param name="sqlCommand">Sql Command</param>
        ''' <param name="connection">SqlConnection ,預設為DefaultConnection</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SelectQuery(ByVal sqlCommand As SqlCommand, Optional ByVal connection As SqlConnection = Nothing) As String

            If String.IsNullOrEmpty(sqlCommand.CommandText) Then
                Throw New Exception("無查詢指令")
            End If
            connection = GetConnection(connection)
            Return ExcuteScalar(sqlCommand, connection)

        End Function

        ''' <summary>
        ''' 查詢資料表,傳回一列或一行的值
        ''' </summary>
        ''' <param name="columnName">欄位名稱()</param>
        ''' <param name="tableName">資料表名稱</param>
        ''' <param name="conditional">v</param>
        ''' <param name="type">取值方式,預設為Rows</param>
        ''' <param name="divisionString">分割的字元,預設為","</param>
        ''' <param name="connection">SqlConnection ,預設為DefaultConnection</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SelectQuery(ByVal columnName() As String, ByVal tableName As String, ByVal conditional As String, ByVal type As SelectType, Optional ByVal divisionString As String = ",", Optional ByVal connection As SqlConnection = Nothing) As String

            If columnName.Length = 0 Then
                Throw New Exception("無指定欄位")
            End If
            If String.IsNullOrEmpty(tableName) Then
                Throw New Exception("無指定資料表")
            End If
            If String.IsNullOrEmpty(conditional) Then
                Throw New Exception("無指定條件")
            End If
            Dim columnNameTemp As String = ""
            For Each s As String In columnName
                If String.IsNullOrEmpty(s) Then
                    columnNameTemp &= s & ","
                End If
            Next
            If String.IsNullOrEmpty(columnNameTemp) Then
                Throw New Exception("無指定欄位")
            Else
                columnNameTemp = Mid(columnNameTemp, 1, columnNameTemp.Length - 1)
            End If

            Dim sqlQuery As String = "Select " & columnNameTemp & " From " & tableName & " Where " & conditional

            Dim cmd As New SqlCommand(sqlQuery)
            connection = GetConnection(connection)
            Dim dt As DataTable = GetDataTable(cmd, connection)
            If dt.Rows.Count <> 0 Then
                Dim value As String = ""
                For i As Integer = 0 To IIf(type = SelectType.Rows, dt.Rows.Count, dt.Columns.Count) - 1
                    For j As Integer = 0 To IIf(type = SelectType.Rows, dt.Columns.Count, dt.Rows.Count) - 1
                        value += UtilityCommon.ReplaceDBnull(dt.Rows(IIf(type = SelectType.Rows, i, j)).Item(IIf(type = SelectType.Rows, j, i)), "") & divisionString
                    Next
                Next
                If value.Length <> 0 Then
                    value = Mid(value, 1, value.Length - 1)
                End If
                Return value
            Else
                Return String.Empty
            End If

        End Function

        ''' <summary>
        ''' 查詢資料表,傳回一列或一行的值
        ''' </summary>
        ''' <param name="sqlQuery">sqlQuery</param>
        ''' <param name="type">取值方式,預設為Rows</param>
        ''' <param name="divisionString">分割的字元,預設為","</param>
        ''' <param name="connection">SqlConnection ,預設為DefaultConnection</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SelectQuery(ByVal sqlQuery As String, ByVal type As SelectType, Optional ByVal divisionString As String = ",", Optional ByVal connection As SqlConnection = Nothing) As String

            If String.IsNullOrEmpty(sqlQuery) Then
                Throw New Exception("無查詢指令")
            End If
            connection = GetConnection(connection)
            Dim cmd As New SqlCommand(sqlQuery)
            Dim dt As DataTable = GetDataTable(cmd, connection)
            If dt.Rows.Count <> 0 Then
                Dim value As String = ""
                For i As Integer = 0 To IIf(type = SelectType.Rows, dt.Rows.Count, dt.Columns.Count) - 1
                    For j As Integer = 0 To IIf(type = SelectType.Rows, dt.Columns.Count, dt.Rows.Count) - 1
                        value += UtilityCommon.ReplaceDBnull(dt.Rows(IIf(type = SelectType.Rows, i, j)).Item(IIf(type = SelectType.Rows, j, i)), "") & divisionString
                    Next
                Next
                If value.Length <> 0 Then
                    value = Mid(value, 1, value.Length - 1)
                End If
                Return value
            Else
                Return String.Empty
            End If

        End Function

        ''' <summary>
        ''' 查詢資料表,傳回一列或一行的值
        ''' </summary>
        ''' <param name="sqlCommand">sqlCommand</param>
        ''' <param name="type">取值方式,預設為Rows</param>
        ''' <param name="divisionString">分割的字元,預設為","</param>
        ''' <param name="connection">SqlConnection ,預設為DefaultConnection</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SelectQuery(ByVal sqlCommand As SqlCommand, ByVal type As SelectType, Optional ByVal divisionString As String = ",", Optional ByVal connection As SqlConnection = Nothing) As String

            If String.IsNullOrEmpty(sqlCommand.CommandText) Then
                Throw New Exception("無查詢指令")
            End If
            connection = GetConnection(connection)
            Dim dt As DataTable = GetDataTable(sqlCommand, connection)
            If dt.Rows.Count <> 0 Then
                Dim value As String = ""
                For i As Integer = 0 To IIf(type = SelectType.Rows, dt.Rows.Count, dt.Columns.Count) - 1
                    For j As Integer = 0 To IIf(type = SelectType.Rows, dt.Columns.Count, dt.Rows.Count) - 1
                        value += UtilityCommon.ReplaceDBnull(dt.Rows(IIf(type = SelectType.Rows, i, j)).Item(IIf(type = SelectType.Rows, j, i)), "") & divisionString
                    Next
                Next
                If value.Length <> 0 Then
                    value = Mid(value, 1, value.Length - 1)
                End If
                Return value
            Else
                Return String.Empty
            End If

        End Function

        ''' <summary>
        ''' 執行SqlQuery
        ''' </summary>
        ''' <param name="sqlQuery">sqlQuery</param>
        '''  <param name="startTransAction">交易處理,預設為否</param>
        ''' <param name="connection">SqlConnection ,預設為DefaultConnection</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ExecuteNonQuery(ByVal sqlQuery As String, Optional ByVal startTransAction As Boolean = False, Optional ByVal connection As SqlConnection = Nothing) As Integer

            If String.IsNullOrEmpty(sqlQuery) Then
                Throw New Exception("無查詢指令")
            End If
            connection = GetConnection(connection)


            Dim transaction As SqlTransaction
            Dim cmd As New SqlCommand(sqlQuery, connection)
            Try
                connection.Open()


                If startTransAction Then
                    transaction = connection.BeginTransaction()
                    cmd.Transaction = transaction
                End If
                Dim result As Integer = cmd.ExecuteNonQuery()
                If startTransAction Then
                    transaction.Commit()
                End If
                Return result

            Catch ex As Exception

                If startTransAction Then
                    Try
                        transaction.Rollback()
                    Catch e As SqlException
                        If Not transaction.Connection Is Nothing Then
                            Throw e
                        End If
                    End Try
                End If

                Throw ex
            Finally
                If connection.State = ConnectionState.Open Then
                    connection.Close()
                    connection.Dispose()
                    cmd.Dispose()
                End If
            End Try
        End Function

        ''' <summary>
        ''' 執行SQL Command
        ''' </summary>
        ''' <param name="cmd">Sql Command</param>
        ''' <param name="startTransAction">啟動交易處理,預設為否</param>
        ''' <param name="connection">SqlConnection ,預設為DefaultConnection</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ExecuteNonQuery(ByVal cmd As SqlCommand, Optional ByVal startTransAction As Boolean = False, Optional ByVal connection As SqlConnection = Nothing) As Integer

            If cmd Is Nothing Then
                Throw New Exception("無效的SqlCommand")
            End If


            connection = GetConnection(connection)

            Dim transaction As SqlTransaction

            Try
                connection.Open()
                cmd.Connection = connection
                If startTransAction Then
                    transaction = connection.BeginTransaction()
                    cmd.Transaction = transaction
                End If
                Dim result As Integer = cmd.ExecuteNonQuery()
                If startTransAction Then
                    transaction.Commit()
                End If
                Return result

            Catch ex As Exception

                If startTransAction Then
                    Try
                        transaction.Rollback()
                    Catch e As SqlException
                        If Not transaction.Connection Is Nothing Then
                            Throw e
                        End If
                    End Try
                End If

                Throw ex
            Finally
                If connection.State = ConnectionState.Open Then
                    connection.Close()
                    connection.Dispose()
                    cmd.Dispose()
                End If
            End Try
        End Function

        ''' <summary>
        ''' 新增後回傳新的識別ID
        ''' </summary>
        ''' <param name="cmd">SqlCommand</param>
        ''' <param name="connection">SqlConnection ,預設為DefaultConnection</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function InsertAndReturnID(ByVal cmd As SqlCommand, Optional ByVal connection As SqlConnection = Nothing) As String

            If cmd Is Nothing Then
                Throw New Exception("無效的SqlCommand")
            End If


            connection = GetConnection(connection)

            cmd.CommandText = cmd.CommandText & ";select @NewID = @@IDENTITY"
            cmd.Parameters.AddWithValue("@NewID", 0)
            cmd.Parameters("@NewID").Direction = Data.ParameterDirection.Output

            Try
                connection.Open()
                cmd.Connection = connection
                cmd.ExecuteNonQuery()
                Return cmd.Parameters("@NewID").Value.ToString

            Catch ex As Exception
                Throw ex
            Finally
                If connection.State = ConnectionState.Open Then
                    connection.Close()
                    connection.Dispose()
                    cmd.Dispose()
                End If
            End Try
        End Function

        ''' <summary>
        ''' 新增後回傳新的識別ID
        ''' </summary>
        ''' <param name="sqlQuery">sqlQuery</param>
        ''' <param name="connection">SqlConnection ,預設為DefaultConnection</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function InsertAndReturnID(ByVal sqlQuery As String, Optional ByVal connection As SqlConnection = Nothing) As String

            If sqlQuery Is Nothing Then
                Throw New Exception("無查詢指令")
            End If


            sqlQuery &= ";select @NewID = @@IDENTITY"
            connection = GetConnection(connection)
            Dim cmd As New SqlCommand(sqlQuery, connection)

            cmd.Parameters.AddWithValue("@NewID", 0)
            cmd.Parameters("@NewID").Direction = Data.ParameterDirection.Output

            Try
                connection.Open()
                cmd.ExecuteNonQuery()
                Return cmd.Parameters("@NewID").Value.ToString

            Catch ex As Exception
                Throw ex
            Finally
                If connection.State = ConnectionState.Open Then
                    connection.Close()
                    connection.Dispose()
                    cmd.Dispose()
                End If
            End Try
        End Function

        ''' <summary>
        ''' 取得資料庫中的IMAGE欄位值
        ''' </summary>
        ''' <param name="sqlQuery"></param>
        ''' <param name="connection"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDBBinrary(ByVal sqlQuery As String, Optional ByVal connection As SqlConnection = Nothing) As Object
            connection = GetConnection(connection)
            Dim cmd As New SqlCommand(sqlQuery, connection)

            Dim re As Object
            Try
                connection.Open()
                re = cmd.ExecuteScalar()
            Catch ex As Exception
                Throw ex
            Finally
                If connection.State = ConnectionState.Open Then
                    connection.Close()
                    connection.Dispose()
                End If
                cmd.Dispose()
            End Try
            Return re
        End Function

    End Class

End Namespace

