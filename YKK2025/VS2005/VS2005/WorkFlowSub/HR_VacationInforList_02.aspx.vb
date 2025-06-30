Imports System.Data
Imports System.Data.OleDb

Partial Class HR_VacationInforList_02
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
    Dim wDepoID As String          '公司別
    Dim wEmpID As String           'EmpID
    Dim wYear As String            '年度
    Dim SumDays As Double = 0      '日數合計 

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()
        If Not Me.IsPostBack Then
            ShowData()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        wDepoID = Request.QueryString("pDepo")      '公司別
        wEmpID = Request.QueryString("pEmpID")        'EmpID
        wYear = Request.QueryString("pYear")          '年度
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選資料
    '**
    '*****************************************************************
    Sub ShowData()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        Dim wStartDate As String = wYear + "/01/01"     '起始日
        Dim wEndDate As String = wYear + "/12/31"       '截止日
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        Dim wHttp As String = System.Configuration.ConfigurationManager.AppSettings("Http")

        OleDbConnection1.Open()
        DBDataSet1.Clear()
        SQL = "Select No, VacationCode+':'+Vacation As Vacation, ADays, "
        SQL = SQL & "Convert(VARCHAR(10), AStartDate, 111) + ' ' + str(AStartH) + ':00' + '~' + "
        SQL = SQL & "Convert(VARCHAR(10), AEndDate, 111)   + ' ' + str(AEndH) + ':00' as VacationTime, "
        SQL = SQL + "'" + wHttp + "' + '/WorkFlow/HR_TimeOffSheet_02.aspx?' + "
        SQL = SQL + "'&pFormNo=' + FormNo + "
        SQL = SQL + "'&pFormSno=' + str(FormSno) "
        SQL = SQL + " As URL "
        SQL = SQL & "From F_TimeOffSheet "
        SQL = SQL & " Where DepoCode = '" & wDepoID & "'"
        SQL = SQL & "   And EmpID    = '" & wEmpID & "'"
        SQL = SQL & "   And Sts      = '1' "
        SQL = SQL & "   And ( "
        SQL = SQL & "           VacationCode >= 'I' "
        SQL = SQL & "        OR VacationCode >= 'Y' "
        SQL = SQL & "        OR VacationCode >= 'S' "
        SQL = SQL & "        OR VacationCode >= 'Z' "
        SQL = SQL & "        OR VacationCode >= 'M' "
        SQL = SQL & "        OR VacationCode >= 'X' "
        SQL = SQL & "        OR VacationCode >= 'P' "
        SQL = SQL & "        OR VacationCode >= 'Q' "
        SQL = SQL & "       ) "
        SQL = SQL & "   And AStartDate >= '" & wStartDate & "'"
        SQL = SQL & "   And AEndDate   <= '" & wEndDate & "'"
        SQL = SQL & " Order by VacationCode, AStartDate "
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "Data")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()

        OleDbConnection1.Close()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.Header Then
            SumDays = 0      '日數合計 
        End If
        If e.Row.RowType = DataControlRowType.DataRow Then
            SumDays = SumDays + CDbl(e.Row.Cells(3).Text.ToString)
        End If
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(2).Text = "合計"
            e.Row.Cells(3).Text = FormatNumber(SumDays, 1)
        End If
    End Sub


End Class
