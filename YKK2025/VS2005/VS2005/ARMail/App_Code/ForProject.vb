Imports Microsoft.VisualBasic
Imports System.Data

Public Class ForProject

    Public Shared strSavedMessage = "已儲存"
    Public Shared strSaveAlertMessage = "是否儲存?"
    Public Shared strDeletedMessage = "已刪除"
    Public Shared strDeleteAlertMessage = "是否刪除?"
    Public Shared strRecordExist = "此筆記錄已存在"
    Public Shared strRecordNotExist = "此筆記錄不存在"
    Public Shared strCannotIns = "無新增權限"
    Public Shared strCannotMod = "無修改權限"
    Public Shared strCannotDel = "無刪除權限"
    Public Shared strChoiceMsg = "請先選擇主項目"

    '取得連線字串
    Public Shared Function GetConnectionString() As String
        Dim strConnectionString As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("OAOPEN_Con").ToString
        Return strConnectionString
    End Function

    '檢查是否為管理者
    Public Shared Function CheckAdmin(ByVal userId As String) As Boolean
        Dim uDataBase As New Utility.DataBase
        uDataBase.DefaultConnection = New System.Data.SqlClient.SqlConnection(GetConnectionString())

        Dim strSql As String = "Select Count(*) from M_Referp where [cat] = '998' and [DKey] = '" & userId & "'"
        Dim strResult As String = uDataBase.SelectQuery(strSql)
        If strResult = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    '檢查是否為使用管理者
    Public Shared Function CheckUserAdmin(ByVal userId As String) As Boolean
        Dim uDataBase As New Utility.DataBase
        uDataBase.DefaultConnection = New System.Data.SqlClient.SqlConnection(GetConnectionString())

        Dim strSql As String = "Select Count(*) from M_Referp where [cat] = '906' and [DKey] = '" & userId & "'"
        Dim strResult As String = uDataBase.SelectQuery(strSql)
        If strResult = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    '插入維護履歷資料
    Public Shared Sub InsertMaintHistory(ByVal procUser As String, ByVal procProgram As String, ByVal func As String, ByVal before As String, ByVal after As String)
        Dim uDataBase As New Utility.DataBase
        uDataBase.DefaultConnection = New System.Data.SqlClient.SqlConnection(GetConnectionString())

        Dim strSql As String = "Insert into T_MaintHistory (ProcTime,ProcUser,ProcUserName,ProcProgram,[function],before,after) values (@ProcTime,@ProcUser,@ProcUserName,@ProcProgram,@function,@before,@after)"
        Dim CmdSql As New SqlClient.SqlCommand(strSql)
        CmdSql.Parameters.AddWithValue("@ProcTime", Now)
        CmdSql.Parameters.AddWithValue("@ProcUser", procUser)
        CmdSql.Parameters.AddWithValue("@ProcUserName", GetUserName(procUser))
        CmdSql.Parameters.AddWithValue("@ProcProgram", procProgram)
        CmdSql.Parameters.AddWithValue("@function", func)
        CmdSql.Parameters.AddWithValue("@before", before)
        CmdSql.Parameters.AddWithValue("@after", after)
        uDataBase.ExecuteNonQuery(CmdSql)
    End Sub

    '取得UserName from WORKFLOW DB
    Public Shared Function GetUserName(ByVal userId As String) As String
        Dim strConnectionString As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("WORKFLOW_Con").ToString
        Dim uDataBase As New Utility.DataBase
        uDataBase.DefaultConnection = New System.Data.SqlClient.SqlConnection(strConnectionString)

        Dim strSql As String = "Select UserName from  m_users where lower(UserId)='" & userId.ToLower & "'"
        Return uDataBase.SelectQuery(strSql)

    End Function
End Class