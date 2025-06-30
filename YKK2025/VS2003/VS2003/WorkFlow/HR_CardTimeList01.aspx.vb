Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class HR_CardTimeList01
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
    Dim wMonth As String        '工作月
    Dim wEmpID As String        'Emp-ID
    Dim wDepoID As String           '公司別

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()
        If Not Me.IsPostBack Then
            CardData()
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        wMonth = Request.QueryString("pMonth")    '工作月
        wEmpID = Request.QueryString("pEmpID")    'Emp-ID
        wDepoID = Request.QueryString("pDepoID")        '公司別
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選刷卡資料
    '**
    '*****************************************************************
    Sub CardData()
        Dim wStartDate As String = wMonth + "/1"
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()

        SQL = "SELECT "
        SQL = SQL + "Convert(VARCHAR(10), CDate, 111) as CDateDesc, "
        SQL = SQL + "EmpID, TimeA, TimeB "
        SQL = SQL + "FROM HR_WorkTime "
        SQL = SQL + "Where EmpID  = '" & wEmpID & "' "
        SQL = SQL + "  And DepoID = '" & wDepoID & "' "
        SQL = SQL + "  And CDate >= '" & wStartDate & "' "
        SQL = SQL + "Order by CDate Desc "

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "WorkTime")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()

        OleDbConnection1.Close()
    End Sub

End Class
