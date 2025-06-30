Imports System.Data
Imports System.Data.OleDb

Public Class MapManagement
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents Datagrid2 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton
    Protected WithEvents BExcel1 As System.Web.UI.WebControls.ImageButton

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

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wMapNo As String = ""

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "MapManagement.aspx"

        If Not Me.IsPostBack Then
            CheckAuthority()
            SetKey()
            DataList()
        End If
    End Sub

    Sub CheckAuthority()
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub SetKey()
        Dim i As Integer
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        '部門
        DBDataSet1.Clear()
        SQL = "Select DivName From M_Users Group by DivName Order by DivName "
        DDivision.Items.Clear()
        DDivision.Items.Add("ALL")
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("DivName")
            ListItem1.Value = DBTable1.Rows(i).Item("DivName")
            DDivision.Items.Add(ListItem1)
        Next

        'Buyer
        DBuyer.Text = "Buyer"
        '圖面
        DMapNo.Text = "圖號"
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim OleDbConnection1 As New OleDbConnection
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("MapFilePath")
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        SQL = "SELECT "   '取得DB資料
        SQL = SQL & "No, MapNo, FormSno, CompletedTime, Buyer, "
        SQL = SQL & "Case Sts When 0 Then '開發中' When 1 Then '開發OK' When 2 Then '開發NG' ELSE '取消' END AS StsDesc, "
        SQL = SQL & "'" & Path & "' + MapFile As URL "
        SQL = SQL & "FROM F_MapSheet "
        SQL = SQL + "Where Sts <> '0' And Sts <> '3' "
        '部門
        If DDivision.SelectedValue <> "ALL" Then
            SQL = SQL + " And   Division = '" + DDivision.SelectedValue + "'"
        End If
        'Buyer
        If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
            SQL = SQL + " And   Buyer Like '%" + DBuyer.Text + "%'"
        End If
        '圖面
        If DMapNo.Text <> "圖號" And DMapNo.Text <> "" Then
            SQL = SQL + "And MapNo Like '%" & DMapNo.Text & "%' "
        End If
        SQL = SQL + "Order by CompletedTime Desc "

        Dim DBDataSet1 As New DataSet
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_MapSheet")
        Datagrid2.DataSource = DBDataSet1
        Datagrid2.DataBind()

        OleDbConnection1.Close()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        Datagrid2.CurrentPageIndex = 0
        DataList()

        DataGrid1.DataSource = Nothing
        DataGrid1.DataBind()
    End Sub

    Private Sub Datagrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid2.ItemCommand
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("MapModFilePath")

        If e.CommandSource.CommandName = "ShowInformation" Then  '點選選取檢查
            Dim Key As String = Datagrid2.DataKeys(e.Item.ItemIndex)  '所選取的Map No
            wMapNo = Datagrid2.DataKeys(e.Item.ItemIndex)
            Dim SQL As String

            DataGrid1.CurrentPageIndex = 0

            'DB連結設定
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
            SQL = "SELECT "   '取得DB資料
            SQL = SQL & "Case Sts When 0 then '開發中' When 1 then '開發完成-OK' When 2 then '開發完成-NG' Else '開發取消' End As StsDesc, "
            SQL = SQL & "No, MapNo, FormSno, CompletedTime, ModReasonCode, ModContent, Buyer, "
            SQL = SQL & "'" & Path & "' + MapFile As URL "
            SQL = SQL & "FROM F_MapModSheet "
            'SQL = SQL + "Where Sts = '1' "
            'SQL = SQL + "And OriMapNo Like '%" & Key & "%' "
            SQL = SQL + "Where OriMapNo Like '%" & Key & "%' "
            SQL = SQL + "Order by CompletedTime Desc "

            Dim DBDataSet2 As New DataSet
            OleDbConnection1.Open()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "F_MapModSheet")
            DataGrid1.DataSource = DBDataSet2
            DataGrid1.DataBind()

            If DBDataSet2.Tables("F_MapModSheet").Rows.Count > 0 Then
                BExcel1.Visible = True
            Else
                BExcel1.Visible = False
            End If

            'DB連結關閉
            OleDbConnection1.Close()
        End If
    End Sub

    Private Sub Datagrid2_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles Datagrid2.PageIndexChanged
        Datagrid2.CurrentPageIndex = e.NewPageIndex   'DataGrid跳上下頁
        DataList()

        DataGrid1.DataSource = Nothing
        DataGrid1.DataBind()
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As System.Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs)
        Dim SQL As String
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("MapModFilePath")
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        DataGrid1.CurrentPageIndex = e.NewPageIndex   'DataGrid跳上下頁

        SQL = "SELECT "   '取得DB資料
        SQL = SQL & "No, MapNo, FormSno, CompletedTime, ModReason, ModReasonDesc, Buyer, "
        SQL = SQL & "'" & Path & "' + MapFile As URL "
        SQL = SQL & "FROM F_MapModSheet "
        SQL = SQL + "Where Sts = '1' "
        SQL = SQL + "And MapNo Like '%" & wMapNo & "%' "
        SQL = SQL + "Order by CompletedTime Desc "

        Dim DBDataSet2 As New DataSet
        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "F_MapModSheet")
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
        Dim wAllowPaging As Boolean = Datagrid2.AllowPaging   '程式別不同

        'Datagrid2.AllowPaging = False   '程式別不同
        'DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=Map_Management_List.xls")     '程式別不同
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        Datagrid2.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        Datagrid2.AllowPaging = wAllowPaging        '程式別不同
    End Sub

    Private Sub BExcel1_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click, BExcel1.Click
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '程式別不同

        'DataGrid1.AllowPaging = False   '程式別不同
        'DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=Map_Management_Mod_List.xls")     '程式別不同
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
