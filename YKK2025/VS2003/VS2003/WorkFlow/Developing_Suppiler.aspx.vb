Imports System.Data
Imports System.Data.OleDb

Public Class Developing_Suppiler
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DDateTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DRefresh As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DApplyName As System.Web.UI.WebControls.TextBox
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSts As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DProgress As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDays As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWDivision As System.Web.UI.WebControls.DropDownList

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
        Response.Cookies("PGM").Value = "Developing_Suppiler.aspx"

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
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        '表單
        SQL = "SELECT FormNo, FormName FROM W_Developing_Suppiler "
        SQL = SQL + "Group by FormNo, FormName "
        SQL = SQL + "Order by FormNo, FormName "

        OleDbConnection1.Open()
        DFormName.Items.Clear()

        Dim ListItem1 As New ListItem
        ListItem1.Text = "ALL"
        ListItem1.Value = "ALL"
        ListItem1.Selected = True
        DFormName.Items.Add(ListItem1)

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "W_Developing_Suppiler")
        DBTable1 = DBDataSet1.Tables("W_Developing_Suppiler")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem2 As New ListItem
            ListItem2.Text = DBTable1.Rows(i).Item("FormName")
            ListItem2.Value = DBTable1.Rows(i).Item("FormNo")
            DFormName.Items.Add(ListItem2)
        Next

        '部門
        DBDataSet1.Clear()
        SQL = "SELECT WDivision FROM W_Developing_Suppiler "
        SQL = SQL + " Where WDivision <> '' "
        SQL = SQL + "Group by WDivision "
        SQL = SQL + "Order by WDivision "

        DWDivision.Items.Clear()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "W_Developing")
        DBTable1 = DBDataSet1.Tables("W_Developing")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem3 As New ListItem
            ListItem3.Text = DBTable1.Rows(i).Item("WDivision")
            ListItem3.Value = DBTable1.Rows(i).Item("WDivision")
            DWDivision.Items.Add(ListItem3)
        Next

        OleDbConnection1.Close()

        DApplyName.Text = "依賴者"
        DBuyer.Text = "Buyer"
        DSliderCode.Text = "Slider Code"
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        SQL = "Select * From W_Developing "
        SQL = SQL & " Order by Unique_ID Desc "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "wDevelop")
        If DBDataSet1.Tables("wDevelop").Rows.Count > 0 Then
            DDateTime.Text = DBDataSet1.Tables("wDevelop").Rows(0).Item("CreateTime") + "時點"
            DRefresh.Text = "更新頻率：１次／１日"
        Else
            DDateTime.Text = ""
            DRefresh.Text = ""
        End If
        OleDbConnection1.Close()

        SQL = "SELECT "
        SQL = SQL + "FormNo, FormSno, FormName, SeqNo, Progress, ApplyID, ApplyName, ApplyTime, Step, "
        SQL = SQL + "StepName, DecideID, DecideName, BStartTime, BEndTime, ViewURL, Suppiler, Buyer, SliderCode, "
        SQL = SQL + "Case No When '' Then '未編號' Else No End As No, "
        SQL = SQL + "Case Sts When '0' Then '開發中' When '1' Then '開發完成(OK)' When '2' Then '開發完成(NG)' Else '開發取消' End As Sts, "

        SQL = SQL + "'....' as WorkFlow, "
        SQL = SQL + "'BefOPList.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pFormSno=' + str(FormSno,Len(FormSno)) + "
        SQL = SQL + "'&pStep='    + str(Step,Len(Step)) + "
        SQL = SQL + "'&pSeqNo='   + str(SeqNo,Len(SeqNo)) + "
        SQL = SQL + "'&pApplyID=' + ApplyID "
        SQL = SQL + " As OPURL "

        SQL = SQL + "FROM W_Developing_Suppiler "
        '------------------------------------------------ 
        SQL = SQL + "Where Sts <> '9' "
        '表單
        If DFormName.SelectedValue <> "ALL" Then
            SQL = SQL + " And FormNo = '" + DFormName.SelectedValue + "'"
        End If

        '開發狀態
        If DSts.SelectedValue <> "ALL" Then
            If DSts.SelectedValue = "1" Then
                SQL = SQL + " And Sts =  '0'  "
            Else
                SQL = SQL + " And Sts <> '0'  "
            End If
        End If

        '部門
        If DWDivision.SelectedValue <> "ALL" Then
            SQL = SQL + " And WDivision Like '%" + DWDivision.SelectedValue + "%'"
        End If

        '進度狀態
        If DProgress.SelectedValue <> "ALL" Then
            SQL = SQL + " And Progress Like '%" + DProgress.SelectedValue + "%'"

            If DProgress.SelectedValue = "延遲" Then
                If DDays.SelectedValue = "3" Then SQL = SQL + " And BWorkTime   > '0'    And BWorkTime <= '1440' "
                If DDays.SelectedValue = "6" Then SQL = SQL + " And BWorkTime   > '1440' And BWorkTime <= '2880' "
                If DDays.SelectedValue = "10" Then SQL = SQL + " And BWorkTime  > '2880' And BWorkTime <= '4800' "
                If DDays.SelectedValue = "999" Then SQL = SQL + " And BWorkTime > '4800' "
            End If
        End If

        '依賴者
        If DApplyName.Text <> "" And DApplyName.Text <> "依賴者" Then
            SQL = SQL + " And ApplyName Like '%" + DApplyName.Text + "%'"
        End If
        'Buyer
        If DBuyer.Text <> "" And DBuyer.Text <> "Buyer" Then
            SQL = SQL + " And Buyer Like '%" + DBuyer.Text + "%'"
        End If
        'Slider Code
        If DSliderCode.Text <> "" And DSliderCode.Text <> "Slider Code" Then
            SQL = SQL + " And SliderCode Like '%" + DSliderCode.Text + "%'"
        End If

        SQL = SQL + " Order by ApplyTime Desc "

        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "Progress")
        DataGrid1.DataSource = DBDataSet2
        DataGrid1.DataBind()
        OleDbConnection1.Close()
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged

        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataList()

    End Sub

    Private Sub DFormName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        DataGrid1.CurrentPageIndex = 0
        DataList()

    End Sub

    Private Sub DProgress_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        DataGrid1.CurrentPageIndex = 0
        DataList()

    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        DataList()

    End Sub

    Private Sub DProgress_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles DProgress.SelectedIndexChanged
        Dim i As Integer
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        '部門
        DBDataSet1.Clear()
        SQL = "SELECT WDivision FROM W_Developing_Suppiler "
        SQL = SQL + " Where WDivision <> '' "
        SQL = SQL + "   And Progress Like '%" + DProgress.SelectedValue + "%'"
        SQL = SQL + "Group by WDivision "
        SQL = SQL + "Order by WDivision "

        DWDivision.Items.Clear()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "W_Developing")
        DBTable1 = DBDataSet1.Tables("W_Developing")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem3 As New ListItem
            ListItem3.Text = DBTable1.Rows(i).Item("WDivision")
            ListItem3.Value = DBTable1.Rows(i).Item("WDivision")
            DWDivision.Items.Add(ListItem3)
        Next
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

        Response.AppendHeader("Content-Disposition", "attachment;filename=Suppiler_Developing_List.xls")     '程式別不同
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
