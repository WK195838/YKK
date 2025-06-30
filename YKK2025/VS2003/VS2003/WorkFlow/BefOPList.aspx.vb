Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class BefOPList
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents BRead As System.Web.UI.HtmlControls.HtmlInputButton
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

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As String          '表單流水號
    Dim wStep As Integer            '工程代碼
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者
    Dim wKeepData As String         '封存

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()
        If Not Me.IsPostBack Then
            ShowReadButton()
            SetOPDataList()
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
        wStep = Request.QueryString("pStep")        '工程代碼
        wSeqNo = Request.QueryString("pSeqNo")      '序號
        wApplyID = Request.QueryString("pApplyID")  '申請者
        wKeepData = Request.QueryString("pKeepData")  '封存
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定閱讀按鈕
    '**
    '*****************************************************************
    Sub ShowReadButton()
        If wApplyID = "" Then
            BRead.Visible = True
        Else
            BRead.Visible = False
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定工程資料
    '**
    '*****************************************************************
    Sub SetOPDataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        SQL = "SELECT "
        SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
        SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
        SQL = SQL + "'預定開始：[' + BStartTimeDesc + '], ' + "
        SQL = SQL + "'預定完成：[' + BEndTimeDesc + '], ' + "
        SQL = SQL + "'實際開始：[' + AStartTimeDesc + '], ' + "
        SQL = SQL + "'實際完成：[' + AEndTimeDesc + '] ' As Description, "
        SQL = SQL + "URL, "
        'POP-ADD-Start
        If wFormNo = "000003" Then
            SQL = SQL + "Case Step When 80 Then 'POP' Else '' End As POP, "
            SQL = SQL + "Case Step When 80 Then 'http://10.245.1.10/WorkFlowSub/POPList.aspx?pSys=SPD&pNo=' + No + '&pFormNo=000003&pFormSno=' + LTrim(str(FormSno)) + '&pStep=' + LTrim(str(Step)) Else '' End As POPURL "

        Else
            SQL = SQL + "'' As POP, "
            SQL = SQL + "'' As POPURL "
        End If
        'POP-ADD-End
        If wKeepData = "1" Then
            SQL = SQL + "FROM V_WaitHandle_OLD_01 "
        Else
            SQL = SQL + "FROM V_WaitHandle_01 "
        End If
        SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
        SQL = SQL + "  And FormSno = '" & wFormSno & "' "
        'SQL = SQL + "Order by Unique_ID Desc "
        SQL = SQL + "Order by CreateTime Desc, Step Desc, SeqNo Desc "

        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "WaitHandle")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()
        OleDbConnection1.Close()
    End Sub


    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged

        DataGrid1.CurrentPageIndex = e.NewPageIndex   'DataGrid跳上下頁
        SetOPDataList()

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     閱讀完成
    '**
    '*****************************************************************
    Private Sub BRead_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BRead.ServerClick
        Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.close();</script>", "Form1.DOPReady", "已閱讀"))

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
        SetOPDataList()                 '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=OP_Detail_List.xls")     '程式別不同
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
