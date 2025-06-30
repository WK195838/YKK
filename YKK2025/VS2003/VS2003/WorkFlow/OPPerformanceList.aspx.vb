Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class OPPerformanceList
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid

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
    Dim NowDateTime As String       '現在日期時間
    Dim pTable As String

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "OPPerformanceList.aspx"

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

        SQL = "SELECT No, FormNo, FormSno, Step, SStep, EStep, Division, Buyer, Finished, Suppiler, Level, BTime, ATime, TimeRange, "
        SQL = SQL + "Case Sts When 0 Then '開發中' When 1 Then 'OK' When 2 Then 'NG' Else '取消' End As Sts, "
        SQL = SQL + "BTime-ATime As Balance, "

        SQL = SQL + "'OPPerformanceWorkID.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pFormSno=' + str(FormSno,Len(FormSno)) + "
        SQL = SQL + "'&pStep=' + str(Step,Len(Step)) + "
        SQL = SQL + "'&pSeqNo=' + str(SeqNo,Len(SeqNo)) + "
        SQL = SQL + "'&pWorkID=' + WorkID "
        SQL = SQL + " As URL "

        SQL = SQL + "FROM " + pTable + " "
        SQL = SQL + "Where RecordType = '0' "
        SQL = SQL + "  and FormNo = '" & pFormNo & "' "
        SQL = SQL + "  and FormSno = '" & CStr(pFormSno) & "' "
        SQL = SQL + "Order by Step "

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Header")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()

        OleDbConnection1.Close()

    End Sub


End Class
