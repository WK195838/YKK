Imports System.Data
Imports System.Data.OleDb

Public Class MapList
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
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wCode As String             'Slider Group Code 
    Dim wFun As String              'Function

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
        wCode = Request.QueryString("pCode")        'Slider Group Code
        wFun = Request.QueryString("pFun")          'Function
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        If wFun = "MAP" Then
            SQL = "SELECT "
            SQL = SQL + "Case Sts When 0 then '未結' When 1 then '已結OK' When 2 then '已結NG' else '抽單' end as StsDesc, "
            SQL = SQL + "No, Convert(Varchar, Date,111) as DateDesc, ModReasonCode as ModReason, MapNo, "
            SQL = SQL + "OriMapNo, BefMapNo, "
            SQL = SQL + "'MapModSheet_02.aspx?' + "
            SQL = SQL + "'pFormNo='   + FormNo + "
            SQL = SQL + "'&pFormSno=' + str(FormSno,Len(FormSno)) "
            SQL = SQL + " As URL "
            SQL = SQL + "FROM F_MapModSheet "
            SQL = SQL + "Where OriFormNo = '" & wFormNo & "'"
            SQL = SQL + "  and OriFormSno = '" & CStr(wFormSno) & "'"
        End If
        SQL = SQL + "Order by No Desc "

        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "Sheet")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()
        OleDbConnection1.Close()
    End Sub

    Private Sub DataGrid1_PageIndexChanged1(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataList()
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataGrid1.DataBind()
    End Sub

End Class
