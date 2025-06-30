Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class EDXList01
    Inherits System.Web.UI.Page


    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim uJavaScript As New Utility.JScript
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String
    Dim pSize As String
    Dim pPuller As String
    Dim pColor As String


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "EDXList01.aspx"

        If Not Me.IsPostBack Then
            SetParameter()          '設定共用參數
            DataList()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        pSize = Request.QueryString("pSize")
        pPuller = Request.QueryString("pPuller")
        pColor = Request.QueryString("pColor")
        '
        DSPEC.Text = pPuller
        GridView1.Visible = False
        '
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '
        SQL = "SELECT "
        SQL = SQL & "[Seqno],[Cat],[Size],[Family],[Body],[Puller],[Color],[Finish],[Supplier] "
        SQL = SQL & "FROM M_EDX "
        SQL = SQL & "where ( "
        SQL = SQL & "   Puller = '" & pPuller & "' "
        SQL = SQL & "OR Puller = '" & pPuller & "-B' "
        SQL = SQL & "OR Puller = '" & pPuller & "SK' "
        SQL = SQL & ") "
        SQL = SQL & "and Color  = '" & pColor & "' "
        SQL = SQL & "Order By [Seqno],[Cat],[Size],[Family],[Body] "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "EDX")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()
        GridView1.Visible = True
        OleDbConnection1.Close()
        '
    End Sub


    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '
        If DSPEC.Text <> "" Then
            SQL = "SELECT "
            SQL = SQL & "[Seqno],[Cat],[Size],[Family],[Body],[Puller],[Color],[Finish],[Supplier] "
            SQL = SQL & "FROM M_EDX "
            SQL = SQL & "where [Cat]+[Size]+[Family]+[Body]+[Puller]+[Color]+[Finish]+[Supplier] like '%" & DSPEC.Text & "%' "
            SQL = SQL & "Order By [Seqno],[Cat],[Size],[Family],[Body] "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "EDX")
            GridView1.DataSource = DBDataSet1
            GridView1.DataBind()
            GridView1.Visible = True
            OleDbConnection1.Close()
        End If
    End Sub
End Class
