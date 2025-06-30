Imports System.Data
Imports System.Data.OleDb

Public Class AgentList
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DUserName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAllForm As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAgentName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "AgentList.aspx"

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
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        '表單
        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And FormNo <= '900000' "
        SQL = SQL + "  And (InqAuthority = '0' "
        SQL = SQL + "       Or (InqAuthority = '1' "
        SQL = SQL + "           And (   InqUser     like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                Or InqUserName like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "               ) "
        SQL = SQL + "          ) "
        SQL = SQL + "      ) "
        SQL = SQL + "Order by FormNo "

        OleDbConnection1.Open()
        DFormName.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Form")
        DBTable1 = DBDataSet1.Tables("M_Form")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("FormName")
            ListItem1.Value = DBTable1.Rows(i).Item("FormNo")
            DFormName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()

        '原擔當
        SQL = "SELECT UserName, UserID FROM M_Agent "
        SQL = SQL + " Where Active = '1' "
        SQL = SQL + "Group by UserName, UserID "
        SQL = SQL + "Order by UserName, UserID "

        OleDbConnection1.Open()
        DUserName.Items.Clear()
        DUserName.Items.Add("ALL")
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Agent")
        DBTable1 = DBDataSet1.Tables("M_Agent")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("UserName")
            ListItem1.Value = DBTable1.Rows(i).Item("UserID")
            DUserName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()

        '代理人
        SQL = "SELECT AgentName, AgentID FROM M_Agent "
        SQL = SQL + " Where Active = '1' "
        SQL = SQL + "Group by AgentName, AgentID "
        SQL = SQL + "Order by AgentName, AgentID "

        OleDbConnection1.Open()
        DAgentName.Items.Clear()
        DAgentName.Items.Add("ALL")
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "M_Agent1")
        DBTable1 = DBDataSet1.Tables("M_Agent1")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("AgentName")
            ListItem1.Value = DBTable1.Rows(i).Item("AgentID")
            DAgentName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()
    End Sub

    Sub DataList()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        SQL = "SELECT "
        SQL = SQL + "FormName, FormNo, Description, UserName, AgentName, "
        SQL = SQL + "Case AllForm When 0 Then '全委託單' Else '單一委託單' End As AllFormDesc, "
        SQL = SQL + "Convert(VarChar, StartDate, 20) + '~' + Convert(VarChar, EndDate, 20) As RangeDesc "
        SQL = SQL + "FROM M_Agent "
        SQL = SQL + "Where Active = '1' "
        '代理類型
        If DAllForm.SelectedValue <> "ALL" Then
            SQL = SQL + " And   AllForm = '" + DAllForm.SelectedValue + "'"
        End If
        '表單
        If DAllForm.SelectedValue <> "0" And DAllForm.SelectedValue <> "ALL" Then
            SQL = SQL + " And   FormNo = '" + DFormName.SelectedValue + "'"
        End If
        '原擔當
        If DUserName.SelectedValue <> "ALL" Then
            SQL = SQL + " And   UserID = '" + DUserName.SelectedValue + "'"
        End If
        '代理人
        If DAgentName.SelectedValue <> "ALL" Then
            SQL = SQL + " And   AgentID = '" + DAgentName.SelectedValue + "'"
        End If

        'Sort
        SQL = SQL + " Order by StartDate, EndDate "

        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Agent")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()
        OleDbConnection1.Close()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        DataList()
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataList()
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataGrid1.DataBind()
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
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '程式別不同

        DataGrid1.AllowPaging = False   '程式別不同
        DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=Agent_List.xls")     '程式別不同
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        DataGrid1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        DataGrid1.AllowPaging = wAllowPaging        '程式別不同
    End Sub

End Class
