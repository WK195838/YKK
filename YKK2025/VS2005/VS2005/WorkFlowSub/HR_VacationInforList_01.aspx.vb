Imports System.Data
Imports System.Data.OleDb

Partial Class HR_VacationInforList_01
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
    Dim wBaseDate As String        '基準日
    Dim wDepoID As String          '公司別
    Dim wEmpID As String           'EmpID
    Dim wInDate As String          '入社日
    Dim wOldInDate As String       '留職停薪前入社日

    Dim SumDays As Double = 0      '日數合計 
    Dim SumDaysLY As Double = 0    'LastYear日數合計 

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
        wBaseDate = Request.QueryString("pBaseDate")  '基準日
        wDepoID = Request.QueryString("pDepoID")      '公司別
        wEmpID = Request.QueryString("pEmpID")        'EmpID
        wInDate = Request.QueryString("pInDate")      '入社日
        wOldInDate = Request.QueryString("pOldInDate")      '留職停薪前入社日
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選資料
    '**
    '*****************************************************************
    Sub ShowData()
        Dim SQL As String
        Dim wSYYMM As String = ""
        Dim wEYYMM As String = ""
        Dim wStartDate As String = ""
        Dim wEndDate As String = ""
        Dim wStartDate1 As String = ""
        Dim wEndDate1 As String = ""
        Dim DBDataSet1, DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        Dim wHttp As String = System.Configuration.ConfigurationManager.AppSettings("Http")
        '
        '計算起迄日期
        '非年休
        If wBaseDate >= CDate(Now.Date.Year.ToString + "/4/1") Then
            wSYYMM = Now.Date.Year.ToString + "/04"
            wEYYMM = CStr(CInt(Now.Date.Year.ToString) + 1) + "/03"
        Else
            wSYYMM = CStr(CInt(Now.Date.Year.ToString) - 1) + "/04"
            wEYYMM = Now.Date.Year.ToString + "/03"
        End If
        '年休
        If wBaseDate >= CDate(Now.Date.Year.ToString + Mid(wInDate, 5, 6)) Then
            wStartDate = Now.Date.Year.ToString + Mid(wInDate, 5, 6)
            wEndDate = CStr(CInt(Now.Date.Year.ToString) + 1) + Mid(wInDate, 5, 6)
        Else
            wStartDate = CStr(CInt(Now.Date.Year.ToString) - 1) + Mid(wInDate, 5, 6)
            wEndDate = Now.Date.Year.ToString + Mid(wInDate, 5, 6)
        End If
        '
        OleDbConnection1.Open()

        If wOldInDate <> "" Then
            '
            If wBaseDate >= CDate(Now.Date.Year.ToString + Mid(wOldInDate, 5, 6)) Then
                wStartDate1 = Now.Date.Year.ToString + Mid(wOldInDate, 5, 6)
                wEndDate1 = CStr(CInt(Now.Date.Year.ToString) + 1) + Mid(wOldInDate, 5, 6)
            Else
                wStartDate1 = CStr(CInt(Now.Date.Year.ToString) - 1) + Mid(wOldInDate, 5, 6)
                wEndDate1 = Now.Date.Year.ToString + Mid(wOldInDate, 5, 6)
            End If
            '
            SQL = "Select Days From HR_StopJobTime "
            SQL = SQL & " Where DepoCode = '" & wDepoID & "'"
            SQL = SQL & "   And EmpID    = '" & wEmpID & "'"
            SQL = SQL & "   And StopStart >= '" & wStartDate1 & "'"
            SQL = SQL & "   And StopEnd   <= '" & wEndDate1 & "'"
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet2, "StopJob")
            If DBDataSet2.Tables("StopJob").Rows.Count > 0 Then
                wStartDate = CDate(wStartDate).Year.ToString + Mid(wOldInDate, 5, 6)
            Else

            End If
        End If
        '
        '上線前-今年
        DBDataSet1.Clear()
        SQL = "Select VacationDate, '上線前' As No, Code+':'+CodeName As Vacation, CodeValue As ADays, "
        SQL = SQL & "Convert(VARCHAR(10), VacationDate, 111) As VacationTime "
        SQL = SQL & "From HR_BeforeVacationList "
        SQL = SQL & " Where VacationDate >= '" & wStartDate & "'"
        SQL = SQL & "   And VacationDate <  '" & wEndDate & "'"
        SQL = SQL & "   And DepoCode = '" & wDepoID & "'"
        SQL = SQL & "   And EmpID    = '" & wEmpID & "'"
        SQL = SQL & " Order by Code "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Data")
        GridView2.DataSource = DBDataSet1
        GridView2.DataBind()
        '
        '系統-今年
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
        SQL = SQL & "   And VacationCode >= 'A' And VacationCode <= 'Z' "
        SQL = SQL & "   And (    "
        SQL = SQL & "          (  "
        SQL = SQL & "                 VacationCode <> 'A' "
        SQL = SQL & "             And SalaryYM >= '" & wSYYMM & "'"
        SQL = SQL & "             And SalaryYM <= '" & wEYYMM & "'"
        SQL = SQL & "          )  "
        SQL = SQL & "          Or  "
        SQL = SQL & "          (  "
        SQL = SQL & "                 VacationCode = 'A' "
        SQL = SQL & "             And AStartDate >= '" & wStartDate & "'"
        SQL = SQL & "             And AEndDate   < '" & wEndDate & "'"
        SQL = SQL & "          )  "
        SQL = SQL & "       )  "
        SQL = SQL & " Order by AStartDate, AStartH, AEndDate, AEndH "
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

    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        If e.Row.RowType = DataControlRowType.Header Then
            SumDaysLY = 0      '日數合計 
        End If
        If e.Row.RowType = DataControlRowType.DataRow Then
            SumDaysLY = SumDaysLY + CDbl(e.Row.Cells(3).Text.ToString)
        End If
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(2).Text = "合計"
            e.Row.Cells(3).Text = FormatNumber(SumDaysLY, 1)
        End If

    End Sub
End Class
