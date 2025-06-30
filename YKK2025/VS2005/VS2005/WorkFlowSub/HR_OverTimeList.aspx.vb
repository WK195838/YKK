Imports System.Data
Imports System.Data.OleDb

Partial Class HR_OverTimeList
    Inherits System.Web.UI.Page

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
    Dim wMonth As String           '工作年月
    Dim wEmpID As String           'Emp-ID
    Dim wDepoID As String           '公司別
    Dim SumCommon As Double = 0    '平日加班總時數
    Dim SumVacation As Double = 0  '假日加班總時數
    Dim SumCVacation As Double = 0 '國假加班總時數

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()
        If Not Me.IsPostBack Then
            OverTimeData()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        wMonth = Request.QueryString("pMonth")   '工作年月
        wEmpID = Request.QueryString("pEmpID")   'Emp-ID
        wDepoID = Request.QueryString("pDepoID")        '公司別
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選加班資料
    '**
    '*****************************************************************
    Sub OverTimeData()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        '加班
        SQL = "SELECT "
        SQL = SQL + "Convert(VARCHAR(10), OverTimeDate, 111) As OverTimeDateDesc, "
        SQL = SQL + "'http://10.245.1.10/WorkFlow/HR_OverTimeSheet_02.aspx?' + "
        SQL = SQL + "'&pFormNo=' + FormNo + "
        SQL = SQL + "'&pFormSno=' + str(FormSno) "
        SQL = SQL + " As OverTimeURL, "
        SQL = SQL + "No, "
        SQL = SQL + "Case Sts When 0 Then '未核' Else '已核' End As StsDesc, "
        SQL = SQL + "FAH1*60 + FAH2*60 + FAM1 + FAM2 as Common, "
        SQL = SQL + "FBH*60 + FBM as Vacation, "
        SQL = SQL + "FCH*60 + FCM as CVacation "
        SQL = SQL + "FROM F_OverTimeSheet "
        SQL = SQL + "Where (Sts = '1' or Sts = '0') "
        SQL = SQL + "  And EmpID  = '" & wEmpID & "' "
        SQL = SQL + "  And DepoCode = '" & wDepoID & "' "
        SQL = SQL + "  And SalaryYM = '" & wMonth & "' "
        SQL = SQL + "Order by OverTimeDate "

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "OverTime")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     合計處理-加班
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            e.Row.Cells(3).Text = FormatNumber(CInt(e.Row.Cells(3).Text.ToString) / 60, 1)
            e.Row.Cells(4).Text = FormatNumber(CInt(e.Row.Cells(4).Text.ToString) / 60, 1)
            e.Row.Cells(5).Text = FormatNumber(CInt(e.Row.Cells(5).Text.ToString) / 60, 1)
            SumCommon = SumCommon + CDbl(e.Row.Cells(3).Text.ToString)
            SumVacation = SumVacation + CDbl(e.Row.Cells(4).Text.ToString)
            SumCVacation = SumCVacation + CDbl(e.Row.Cells(5).Text.ToString)
        End If
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(2).Text = "合計"
            e.Row.Cells(3).Text = FormatNumber(SumCommon, 1)
            e.Row.Cells(4).Text = FormatNumber(SumVacation, 1)
            e.Row.Cells(5).Text = FormatNumber(SumCVacation, 1)
        End If
    End Sub

End Class
