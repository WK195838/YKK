Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class OPPerformanceWorkID
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DAStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndTime As System.Web.UI.WebControls.TextBox

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
    Dim pStep As Integer
    Dim pSeqNo As Integer
    Dim pWorkID As String
    Dim NowDateTime As String       '現在日期時間

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "OPPerformanceWorkID.aspx"

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
        pFormSno = Request.QueryString("pFormSno")    '表單號碼
        pStep = Request.QueryString("pStep")
        pSeqNo = Request.QueryString("pSeqNo")
        pWorkID = Request.QueryString("pWorkID")
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim wEndTime As String = ""
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDBCommand1 As New OleDbCommand
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()

        SQL = "SELECT AStartTime, AEndTime FROM T_WaitHandle "
        SQL = SQL + "Where FormNo  = '" + pFormNo + "'"
        SQL = SQL + "  And FormSno = '" + CStr(pFormSno) + "'"
        SQL = SQL + "  And Step    = '" + CStr(pStep) + "' "
        SQL = SQL + "  And SeqNo   = '" + CStr(pSeqNo) + "' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "WaitHandle")
        If DBDataSet1.Tables("WaitHandle").Rows.Count > 0 Then
            wEndTime = CStr(DBDataSet1.Tables("WaitHandle").Rows(0).Item("AEndTime").Date) + " " + _
                       CStr(DBDataSet1.Tables("WaitHandle").Rows(0).Item("AEndTime").Hour) + ":" + _
                       CStr(DBDataSet1.Tables("WaitHandle").Rows(0).Item("AEndTime").Minute) + ":" + _
                       CStr(DBDataSet1.Tables("WaitHandle").Rows(0).Item("AEndTime").Second)

            DAStartTime.Text = CStr(DBDataSet1.Tables("WaitHandle").Rows(0).Item("AStartTime").Date) + " " + _
                               CStr(DBDataSet1.Tables("WaitHandle").Rows(0).Item("AStartTime").Hour) + ":" + _
                               CStr(DBDataSet1.Tables("WaitHandle").Rows(0).Item("AStartTime").Minute) + ":" + _
                               CStr(DBDataSet1.Tables("WaitHandle").Rows(0).Item("AStartTime").Second)
            DAEndTime.Text = wEndTime
        End If

        SQL = "SELECT Top 20 No, FormName, StepName+'('+str(Step)+')' as StepName, UserName As DecideName, DivName, "
        SQL = SQL + "Convert(VarChar, AStartTime, 20) + '~' + Convert(VarChar, AEndTime, 20) As TimeRange, AWorkTime "
        SQL = SQL + "FROM V_WaitHandle_01 "

        SQL = SQL + "Where Active = '0' "
        SQL = SQL + "  and FlowType <> '0' "
        SQL = SQL + "  and WorkID = '" & pWorkID & "' "
        SQL = SQL + "  and AEndTime <= '" & wEndTime & "' "
        SQL = SQL + "Order by AEndTime Desc "

        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "Header")
        DataGrid1.DataSource = DBDataSet2
        DataGrid1.DataBind()

        OleDbConnection1.Close()

    End Sub

End Class
