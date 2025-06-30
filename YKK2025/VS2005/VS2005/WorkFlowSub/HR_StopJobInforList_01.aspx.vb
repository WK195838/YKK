Imports System.Data
Imports System.Data.OleDb

Partial Class HR_StopJobInforList_01
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wDepoID As String
    Dim wEmpID As String

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()
        If Not Me.IsPostBack Then
            StopJobData()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        wDepoID = Request.QueryString("pDepoID")
        wEmpID = Request.QueryString("pEmpID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選請假資料
    '**
    '*****************************************************************
    Sub StopJobData()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        '--------------------------------------------------------------------------------------------------------------------
        '系統資料
        SQL = "SELECT "
        SQL = SQL + "Convert(VARCHAR(10), StopStart, 111) As StopStart, "
        SQL = SQL + "Convert(VARCHAR(10), StopEnd, 111) As StopEnd, "
        SQL = SQL + "Days "
        SQL = SQL + "FROM HR_StopJobTime "
        SQL = SQL + "Where DepoCode = '" & wDepoID & "' "
        SQL = SQL + "  And EmpID    = '" & wEmpID & "' "
        SQL = SQL + "Order by StopStart "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Vacation")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()
        '--------------------------------------------------------------------------------------------------------------------
        OleDbConnection1.Close()
    End Sub

End Class
