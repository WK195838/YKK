Imports System.Data
Imports System.Data.OleDb

Partial Class INQ_ITMasterStatus
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

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()
        If Not Me.IsPostBack Then
            ShowData()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        DMessqge1.Style("left") = -500 & "px"

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選加班資料
    '**
    '*****************************************************************
    Sub ShowData()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        'Count
        SQL = "SELECT "
        SQL = SQL + "a.formno , b.formname, count(*) as RecCount "
        SQL = SQL + "FROM T_WaitHandle a, m_form b "
        SQL = SQL + "where a.formno=b.formno "
        SQL = SQL + "and a.formno in ('003102','003103','003104','008001') "
        SQL = SQL + "and a.workid='it004' "
        SQL = SQL + "and a.active=1 "
        SQL = SQL + "GROUP BY a.formno, b.formname "
        SQL = SQL + "ORDER BY  a.formno, b.formname "

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "MASTER")

        If DBDataSet1.Tables("MASTER").Rows.Count > 0 Then
            GridView1.DataSource = DBDataSet1
            GridView1.DataBind()
            '
            '003102
            DBDataSet1.Clear()
            SQL = "SELECT top 1 "
            SQL = SQL + "a.formno , b.formname, a.ReceiptTime "
            SQL = SQL + "FROM T_WaitHandle a, m_form b "
            SQL = SQL + "where a.formno=b.formno "
            SQL = SQL + "and a.formno in ('003102') "
            SQL = SQL + "and a.workid='it004' "
            SQL = SQL + "and a.active=1 "
            SQL = SQL + "order by a.ReceiptTime "

            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "MASTER")
            GridView2.DataSource = DBDataSet1
            GridView2.DataBind()
            '
            '003103
            DBDataSet1.Clear()

            SQL = "SELECT top 1 "
            SQL = SQL + "a.formno , b.formname, a.ReceiptTime "
            SQL = SQL + "FROM T_WaitHandle a, m_form b "
            SQL = SQL + "where a.formno=b.formno "
            SQL = SQL + "and a.formno in ('003103') "
            SQL = SQL + "and a.workid='it004' "
            SQL = SQL + "and a.active=1 "
            SQL = SQL + "order by a.ReceiptTime "

            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet1, "MASTER")
            GridView3.DataSource = DBDataSet1
            GridView3.DataBind()
            '
            '003104
            DBDataSet1.Clear()

            SQL = "SELECT top 1 "
            SQL = SQL + "a.formno , b.formname, a.ReceiptTime "
            SQL = SQL + "FROM T_WaitHandle a, m_form b "
            SQL = SQL + "where a.formno=b.formno "
            SQL = SQL + "and a.formno in ('003104') "
            SQL = SQL + "and a.workid='it004' "
            SQL = SQL + "and a.active=1 "
            SQL = SQL + "order by a.ReceiptTime "

            Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter4.Fill(DBDataSet1, "MASTER")
            GridView4.DataSource = DBDataSet1
            GridView4.DataBind()
            '
            '008001
            DBDataSet1.Clear()

            SQL = "SELECT top 1 "
            SQL = SQL + "a.formno , b.formname, a.ReceiptTime "
            SQL = SQL + "FROM T_WaitHandle a, m_form b "
            SQL = SQL + "where a.formno=b.formno "
            SQL = SQL + "and a.formno in ('008001') "
            SQL = SQL + "and a.workid='it004' "
            SQL = SQL + "and a.active=1 "
            SQL = SQL + "order by a.ReceiptTime "

            Dim DBAdapter5 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter5.Fill(DBDataSet1, "MASTER")
            GridView5.DataSource = DBDataSet1
            GridView5.DataBind()
            '
            OleDbConnection1.Close()
        Else
            DMessqge1.Style("left") = 20 & "px"

        End If
        ' --------------------------------------------
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     合計處理-加班
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     選取
    '**
    '*****************************************************************
    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        '
    End Sub
End Class
