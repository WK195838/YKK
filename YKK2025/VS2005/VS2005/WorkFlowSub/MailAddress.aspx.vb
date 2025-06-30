Imports System.Data
Imports System.Data.OleDb

Partial Class MailAddress
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not Me.IsPostBack Then
            CheckAuthority()
            SetSearchItem()
            DataList()
        End If
    End Sub

    Sub CheckAuthority()
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub SetSearchItem()
        Dim i As Integer
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        '公司
        SQL = "SELECT DepoName FROM M_Users "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "Group by DepoName  "
        SQL = SQL + "Order by DepoName  "
        OleDbConnection1.Open()

        DDepoName.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("DepoName")
            ListItem1.Value = DBTable1.Rows(i).Item("DepoName")
            DDepoName.Items.Add(ListItem1)
        Next

        '部門
        DBDataSet1.Clear()
        SQL = "SELECT HRWDivName FROM M_Users "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  and DepoName = '" & DDepoName.SelectedValue & "' "
        SQL = SQL + "Group by HRWDivName "
        SQL = SQL + "Order by HRWDivName "

        DDivision.Items.Clear()
        DDivision.Items.Add("ALL")
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("HRWDivName")
            ListItem1.Value = DBTable1.Rows(i).Item("HRWDivName")
            DDivision.Items.Add(ListItem1)
        Next

        '姓
        DFirstName.Text = ""

        'DBDataSet1.Clear()
        'SQL = "SELECT substring(UserName,1,1) as FirstName FROM M_Users "
        'SQL = SQL + "Where Active = '1' "
        'SQL = SQL + "  And DepoName = '" & DDepoName.SelectedValue & "' "
        'If DDivision.SelectedValue <> "ALL" Then
        '    SQL = SQL + " And   HRWDivName = '" + DDivision.SelectedValue + "'"
        'End If
        'SQL = SQL + "Group by substring(UserName,1,1) "
        'SQL = SQL + "Order by substring(UserName,1,1) "

        'DFirstName.Items.Clear()
        'DFirstName.Items.Add("ALL")
        'Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter3.Fill(DBDataSet1, "FirstName")
        'DBTable1 = DBDataSet1.Tables("FirstName")
        'For i = 0 To DBTable1.Rows.Count - 1
        '    Dim ListItem1 As New ListItem
        '    ListItem1.Text = DBTable1.Rows(i).Item("FirstName")
        '    ListItem1.Value = DBTable1.Rows(i).Item("FirstName")
        '    DFirstName.Items.Add(ListItem1)
        'Next
    End Sub

    Sub DataList()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        SQL = "SELECT "
        SQL = SQL + "DepoName as DepoNameDesc, HRWDivName as DivisionDesc, UserName as NameDesc, "
        SQL = SQL + "Case Mail When 'inq@ykk.com.tw' then '' When 'ssd@ykk.com.tw' then '' Else Mail End as I_MailAddress, "
        SQL = SQL + "'Mailto:'+Mail as MailURL "
        SQL = SQL + "FROM M_Users "
        SQL = SQL + "Where DepoName = '" + DDepoName.SelectedValue + "' "
        SQL = SQL + "  and Active = '1' "
        SQL = SQL + "  and UserName <> '資訊管理' "
        SQL = SQL + "  and UserName <> '工讀生' "

        '部門
        If DDivision.SelectedValue <> "ALL" Then
            SQL = SQL + " And   HRWDivName = '" + DDivision.SelectedValue + "'"
        End If
        '姓
        If DFirstName.Text <> "" Then
            SQL = SQL + "  And UserName Like '" + DFirstName.Text + "%' "
        End If
        'If DFirstName.SelectedValue <> "ALL" Then
        '    SQL = SQL + "  And UserName Like '" + DFirstName.SelectedValue + "%' "
        'End If
        SQL = SQL + " Order by HRWDivName, UserName "

        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Mail")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()
        OleDbConnection1.Close()
    End Sub

    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim wAllowPaging As Boolean = GridView1.AllowPaging   '程式別不同

        GridView1.AllowPaging = False   '程式別不同
        DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=MailAddress_List.xls")     '程式別不同
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        GridView1.AllowPaging = wAllowPaging        '程式別不同
    End Sub

    Protected Sub DDepoName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDepoName.SelectedIndexChanged
        Dim i As Integer
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        '部門
        DBDataSet1.Clear()
        SQL = "SELECT HRWDivName FROM M_Users "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  and DepoName = '" & DDepoName.SelectedValue & "' "
        SQL = SQL + "Group by HRWDivName "
        SQL = SQL + "Order by HRWDivName "

        DDivision.Items.Clear()
        DDivision.Items.Add("ALL")
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("HRWDivName")
            ListItem1.Value = DBTable1.Rows(i).Item("HRWDivName")
            DDivision.Items.Add(ListItem1)
        Next

        '姓
        'DFirstName.Text = ""

        'DBDataSet1.Clear()
        'SQL = "SELECT substring(UserName,1,1) as FirstName FROM M_Users "
        'SQL = SQL + "Where Active = '1' "
        'SQL = SQL + "  And DepoName = '" & DDepoName.SelectedValue & "' "
        'If DDivision.SelectedValue <> "ALL" Then
        '    SQL = SQL + " And   HRWDivName = '" + DDivision.SelectedValue + "'"
        'End If
        'SQL = SQL + "Group by substring(UserName,1,1) "
        'SQL = SQL + "Order by substring(UserName,1,1) "

        'DFirstName.Items.Clear()
        'DFirstName.Items.Add("ALL")
        'Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter3.Fill(DBDataSet1, "FirstName")
        'DBTable1 = DBDataSet1.Tables("FirstName")
        'For i = 0 To DBTable1.Rows.Count - 1
        '    Dim ListItem1 As New ListItem
        '    ListItem1.Text = DBTable1.Rows(i).Item("FirstName")
        '    ListItem1.Value = DBTable1.Rows(i).Item("FirstName")
        '    DFirstName.Items.Add(ListItem1)
        'Next

        DataList()

    End Sub

    Protected Sub DDivision_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDivision.SelectedIndexChanged
        Dim i As Integer
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        '姓
        'DFirstName.Text = ""

        'DBDataSet1.Clear()
        'SQL = "SELECT substring(UserName,1,1) as FirstName FROM M_Users "
        'SQL = SQL + "Where Active = '1' "
        'SQL = SQL + "  And DepoName = '" & DDepoName.SelectedValue & "' "
        'If DDivision.SelectedValue <> "ALL" Then
        '    SQL = SQL + " And   HRWDivName = '" + DDivision.SelectedValue + "'"
        'End If
        'SQL = SQL + "Group by substring(UserName,1,1) "
        'SQL = SQL + "Order by substring(UserName,1,1) "

        'DFirstName.Items.Clear()
        'DFirstName.Items.Add("ALL")
        'Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter3.Fill(DBDataSet1, "FirstName")
        'DBTable1 = DBDataSet1.Tables("FirstName")
        'For i = 0 To DBTable1.Rows.Count - 1
        '    Dim ListItem1 As New ListItem
        '    ListItem1.Text = DBTable1.Rows(i).Item("FirstName")
        '    ListItem1.Value = DBTable1.Rows(i).Item("FirstName")
        '    DFirstName.Items.Add(ListItem1)
        'Next

        DataList()

    End Sub


    Protected Sub BGo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGo.Click
        DataList()
    End Sub
End Class
