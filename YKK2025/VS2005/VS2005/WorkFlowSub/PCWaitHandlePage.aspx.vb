Imports System.Data
Imports System.Data.OleDb

Partial Class PCWaitHandlePage
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "PCWaitHandlePage.aspx"
        If Not Me.IsPostBack Then
            WaitHandle_DataList()
        End If
    End Sub

    Sub WaitHandle_DataList()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        Dim Http As String = System.Configuration.ConfigurationManager.AppSettings("Http")  'Http Address
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        DLow.Text = "(0)"
        DHigh.Text = "(0)"

        OleDbConnection1.Open()
        '一般件
        SQL = "SELECT "
        SQL = SQL + "Count(*) As Low "
        SQL = SQL + "FROM V_WaitHandle_01 "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And (Sts = '0'  Or  Sts = '4') "
        SQL = SQL + "  And Delay = '0' "
        SQL = SQL + "  And DecideID = '" + Request.QueryString("pUserID") + "'"
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "WaitHandle")
        If DBDataSet1.Tables("WaitHandle").Rows(0).Item("Low") > 0 Then
            DLow.Text = "(" + DBDataSet1.Tables("WaitHandle").Rows(0).Item("Low").ToString + ")"
        End If
        '急件
        DBDataSet1.Clear()
        SQL = "SELECT "
        SQL = SQL + "Count(*) As High "
        SQL = SQL + "FROM V_WaitHandle_01 "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And (Sts = '0'  Or  Sts = '4') "
        SQL = SQL + "  And Delay = '1' "
        SQL = SQL + "  And DecideID = '" + Request.QueryString("pUserID") + "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "WaitHandle")
        If DBDataSet1.Tables("WaitHandle").Rows(0).Item("High") > 0 Then
            DHigh.Text = "(" + DBDataSet1.Tables("WaitHandle").Rows(0).Item("High").ToString + ")"
        End If

        OleDbConnection1.Close()

    End Sub
End Class
