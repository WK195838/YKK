Imports System.Data
Imports System.Data.OleDb

Public Class OPProgressAdv
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DActive As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DStepName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSts As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSDate As System.Web.UI.WebControls.Button
    Protected WithEvents DEDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BEDate As System.Web.UI.WebControls.Button
    Protected WithEvents DSortKey As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSort As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Go As System.Web.UI.WebControls.Button

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
        Response.Cookies("PGM").Value = "OPProgressAdv.aspx"

        If Not Me.IsPostBack Then
            CheckAuthority()
            SetSearchItem()
            DataList()
        End If
    End Sub

    Sub CheckAuthority()
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
        '日期選擇
        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"
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
        SQL = SQL + "  And FormNo <= '100000' "
        SQL = SQL + "  And (Authority = '0' "
        SQL = SQL + "       Or (Authority = '1' "
        SQL = SQL + "           And (Users1 like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                Or Users2 like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
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

        '工程
        SQL = "SELECT StepName, Step FROM V_Flow_01 "
        SQL = SQL + " Where Active = '1' "
        SQL = SQL + " And   FormNo = '" + DFormName.SelectedValue + "'"
        SQL = SQL + "Order by Step, StepName "

        OleDbConnection1.Open()
        DStepName.Items.Clear()
        DStepName.Items.Add("ALL")
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "V_Flow")
        DBTable1 = DBDataSet1.Tables("V_Flow")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("StepName")
            ListItem1.Value = DBTable1.Rows(i).Item("Step")
            DStepName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()

        '日期
        DSDate.Text = CStr(DateAdd("d", -30, DateTime.Now.Today))
        DEDate.Text = CStr(DateTime.Now.Today)
        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"
    End Sub

    Sub DataList()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        SQL = "SELECT "
        SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
        SQL = SQL + "StsDesc, FormName, FlowTypeDesc, ApplyName, DecideName, StepNameDesc, "
        SQL = SQL + "'流程資訊' As WorkFlow, "
        SQL = SQL + "'申請時間：[' + Convert(VarChar, ApplyTime, 20) + '], ' + "
        SQL = SQL + "'收件時間：[' + Convert(VarChar, ReceiptTime, 20) + '], ' + "
        SQL = SQL + "'閱讀期限：[' + Convert(VarChar, ReadTimeLimit, 20) + '], ' + "
        SQL = SQL + "'首次閱讀：[' + FirstReadTimeDesc + '], ' + "
        SQL = SQL + "'最後閱讀：[' + LastReadTimeDesc  + '], ' + "
        SQL = SQL + "'預定開始：[' + Convert(VarChar, BStartTime, 20) + '], ' + "
        SQL = SQL + "'預定完成：[' + Convert(VarChar, BEndTime, 20) + '], ' + "
        SQL = SQL + "'實際開始：[' + Convert(VarChar, AStartTime, 20) + '], ' + "
        SQL = SQL + "'實際完成：[' + AEndTimeDesc + '] '  "
        SQL = SQL + " As Description, ViewURL, "

        SQL = SQL + "'BefOPList.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pFormSno=' + str(FormSno,Len(FormSno)) + "
        SQL = SQL + "'&pStep='    + str(Step,Len(Step)) + "
        SQL = SQL + "'&pSeqNo='   + str(SeqNo,Len(SeqNo)) + "
        SQL = SQL + "'&pApplyID=' + ApplyID "
        SQL = SQL + " As OPURL "

        SQL = SQL + "FROM V_WaitHandle_01 "
        SQL = SQL + "Where Active <> '9' "
        '表單
        If DFormName.SelectedValue <> "ALL" Then
            SQL = SQL + " And   FormNo = '" + DFormName.SelectedValue + "'"
        End If
        '工程
        If DStepName.SelectedValue <> "ALL" Then
            SQL = SQL + " And   Step = '" + DStepName.SelectedValue + "'"
        End If
        '狀態
        SQL = SQL + " And   Active = '" + DActive.SelectedValue + "'"
        '延遲
        SQL = SQL + " And   Delay = '" + DSts.SelectedValue + "'"
        '日期
        If DSDate.Text = "" Then DSDate.Text = CStr(DateAdd("d", -30, DateTime.Now.Today))
        If DEDate.Text = "" Then DEDate.Text = CStr(DateTime.Now.Today)
        SQL = SQL + " And   ApplyTime >= '" + DSDate.Text + " 00:00:00" + "'"
        SQL = SQL + " And   ApplyTime <= '" + DEDate.Text + " 23:59:59" + "'"

        'Sort
        SQL = SQL + " Order by " + DSortKey.SelectedValue + " " + DSort.SelectedValue

        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "WaitHandle")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()
        OleDbConnection1.Close()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        DataList()
    End Sub

    Private Sub DataGrid1_PageIndexChanged1(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataList()
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataGrid1.DataBind()
    End Sub

    Private Sub DFormName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged

        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        '工程
        SQL = "SELECT StepName, Step FROM V_Flow_01 "
        SQL = SQL + " Where Active = '1' "
        SQL = SQL + " And   FormNo = '" + DFormName.SelectedValue + "'"
        SQL = SQL + "Order by Step, StepName "

        OleDbConnection1.Open()
        DStepName.Items.Clear()
        DStepName.Items.Add("ALL")
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "V_Flow")
        DBTable1 = DBDataSet1.Tables("V_Flow")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("StepName")
            ListItem1.Value = DBTable1.Rows(i).Item("Step")
            DStepName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()

    End Sub

End Class
