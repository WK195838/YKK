Imports System.Data
Imports System.Data.OleDb

Public Class Progress
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DDateTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents LDevelop As System.Web.UI.WebControls.HyperLink
    Protected WithEvents Hyperlink1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DRefresh As System.Web.UI.WebControls.TextBox
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
        Response.Cookies("PGM").Value = "Progress.aspx"

        If Not Me.IsPostBack Then
            CheckAuthority()
            DataList()
        End If
    End Sub

    Sub CheckAuthority()
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        SQL = "Select * From W_Progress "
        SQL = SQL & " Order by Unique_ID Desc "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "wProgress")
        If DBDataSet1.Tables("wProgress").Rows.Count > 0 Then
            DDateTime.Text = DBDataSet1.Tables("wProgress").Rows(0).Item("CreateTime") + "時點"
            DRefresh.Text = "更新頻率：１次／１日"

        End If
        OleDbConnection1.Close()

        SQL = "SELECT "
        SQL = SQL + " *, "
        SQL = SQL + "'OPProgress.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pSts=Delay' "
        SQL = SQL + " As DelayURL, "
        SQL = SQL + "'OPProgress.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pSts=ReadDelay' "
        SQL = SQL + " As ReadDelayURL, "
        SQL = SQL + "'OPProgress.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pSts=Normal' "
        SQL = SQL + " As NormalURL, "

        SQL = SQL + "'OPInqCommission.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pSts=OK' "
        SQL = SQL + " As OKURL, "
        SQL = SQL + "'OPInqCommission.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pSts=NG' "
        SQL = SQL + " As NGURL, "
        SQL = SQL + "'OPInqCommission.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pSts=Cancel' "
        SQL = SQL + " As CancelURL, "

        SQL = SQL + "'OPInqCommission.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pSts=OKY' "
        SQL = SQL + " As OKYURL, "
        SQL = SQL + "'OPInqCommission.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pSts=NGY' "
        SQL = SQL + " As NGYURL, "
        SQL = SQL + "'OPInqCommission.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pSts=CancelY' "
        SQL = SQL + " As CancelYURL, "

        SQL = SQL + "'OPInqCommission.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pSts=OKBY' "
        SQL = SQL + " As OKBYURL, "
        SQL = SQL + "'OPInqCommission.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pSts=NGBY' "
        SQL = SQL + " As NGBYURL, "
        SQL = SQL + "'OPInqCommission.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pSts=CancelBY' "
        SQL = SQL + " As CancelBYURL "

        SQL = SQL + "FROM W_Progress "
        SQL = SQL + " Order by Unique_ID "

        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "Progress")
        DataGrid1.DataSource = DBDataSet2
        DataGrid1.DataBind()
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
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '程式別不同

        DataGrid1.AllowPaging = False   '程式別不同
        DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=OP_Management_List.xls")     '程式別不同
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
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
