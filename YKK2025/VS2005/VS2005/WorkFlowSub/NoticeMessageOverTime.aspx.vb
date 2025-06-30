Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports Dundas.Charting.WebControl

Partial Class NoticeMessageOverTime
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wFormNo As String           '表單號碼
    Dim wUserID As String           '使請者ID
    Dim NowDateTime As String       '現在日期時間
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then   '不是PostBack
            CheckReadMessageHistory()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '現在日時
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wUserID = Request.QueryString("pUserID")    '使請者ID
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     CheckReadMessageHistory
    '**
    '*****************************************************************
    Sub CheckReadMessageHistory()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        OleDbConnection1.Open()
        If wFormNo <> "" And wUserID <> "" Then
            SQL = "Select * From W_ReadMessageHistory "
            SQL = SQL & " Where FormNo = '" & wFormNo & "' "
            SQL = SQL & "   And ReadUser = '" & wUserID & "' "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "MessageHistory")
            If DBDataSet1.Tables("MessageHistory").Rows.Count <= 0 Then
                SQL = "Insert into W_ReadMessageHistory "
                SQL = SQL + "( "
                SQL = SQL + "FormNo, ReadFlag, ReadUser, FirstReadTime "
                SQL = SQL + ")  "
                SQL = SQL + "VALUES( "
                SQL = SQL + " '" + wFormNo + "', "
                SQL = SQL + " '" + "0" + "', "
                SQL = SQL + " '" + wUserID + "', "
                SQL = SQL + " '" + NowDateTime + "' "
                SQL = SQL + " ) "
                OleDBCommand1.Connection = OleDbConnection1
                OleDBCommand1.CommandText = SQL
                OleDBCommand1.ExecuteNonQuery()
            End If
        End If
        OleDbConnection1.Close()
    End Sub

    Protected Sub BClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BClose.Click
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand
        OleDbConnection1.Open()

        SQL = "Select * From W_ReadMessageHistory "
        SQL = SQL & " Where FormNo = '" & wFormNo & "' "
        SQL = SQL & "   And ReadUser = '" & wUserID & "' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "MessageHistory")
        If DBDataSet1.Tables("MessageHistory").Rows.Count > 0 Then

            SQL = "Update W_ReadMessageHistory Set "
            If DReadFlag.Checked = True Then
                SQL = SQL + " ReadFlag = '" & "1" & "', "
                SQL = SQL + " LastReadTime = '" & NowDateTime & "' "
            Else
                SQL = SQL + " ReadFlag = '" & "0" & "', "
                SQL = SQL + " LastReadTime = null "
            End If
            SQL = SQL & " Where FormNo = '" & wFormNo & "' "
            SQL = SQL & "   And ReadUser = '" & wUserID & "' "

            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()
        End If
        OleDbConnection1.Close()
        '
        Response.Write("<script> parent.window.opener=null;parent.window.close(); </script>")
    End Sub
End Class
