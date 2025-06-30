Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class DevelopLeadTime_EA_List
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
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
    Dim pFormNo As String
    Dim pFormSno As Integer
    Dim pStart As Integer
    Dim pEnd As Integer
    Dim pStartTime As String
    Dim pEndTime As String
    Dim NowDateTime As String       '現在日期時間
    Dim pTable As String

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "DevelopLeadTime_EAList.aspx"

        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
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
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '現在日時

        pFormNo = Request.QueryString("pFormNo")    '表單號碼
        pFormSno = Request.QueryString("pFormSno")  '表單號碼
        pStart = Request.QueryString("pStart")      'Start Step
        pEnd = Request.QueryString("pEnd")          'End Step
        pStartTime = Request.QueryString("pStartTime")   'Start Time
        pEndTime = Request.QueryString("pEndTime")   'End Time
        pTable = Request.QueryString("pTable")
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDBCommand1 As New OleDbCommand
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()


        SQL = "SELECT Step, ReceiptTime, Convert(VarChar, AStartTime, 20) + '~' + Convert(VarChar, AEndTime, 20) as SingleTimeRange, AWorkTime  "
        SQL = SQL + "FROM T_WaitHandle "
        SQL = SQL + "Where FormNo = '" & pFormNo & "' "
        SQL = SQL + "  and FormSno = '" & CStr(pFormSno) & "' "
        SQL = SQL + "  and Step >= '" & CStr(pStart) & "' "
        SQL = SQL + "  and Step <= '" & CStr(pEnd) & "' "
        SQL = SQL + "  and AEndTime >= '" & pStartTime & "' "
        SQL = SQL + "  and AEndTime <= '" & pEndTime & "' "
        SQL = SQL + "  and FlowType <> '0' "
        SQL = SQL + "Order by Step, AEndTime "

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Header")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()

        OleDbConnection1.Close()

    End Sub

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '程式別不同

        DataGrid1.AllowPaging = False   '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=DevelopLeadTime_DetailList.xls")     '程式別不同
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
