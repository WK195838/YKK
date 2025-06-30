Imports System.Data
Imports System.Data.OleDb

Partial Class HR_OverTimeReport03
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
    Dim NowDateTime As String       '現在日期時間
    Dim SumFood As Integer = 0      '合計伙食 
    Dim SumTraffic As Integer = 0   '合計交通 
    Dim SumTotal As Double = 0      '合計核定時 
    Dim SumCVTotal As Double = 0    '合計換休時 

    Dim SumTwoIn As Double = 0      '合計2內時 
    Dim SumTwoOut As Double = 0     '合計2外時 
    Dim SumVacation As Double = 0   '合計假日時 
    Dim SumCVacation As Double = 0  '合計國定假日時 
    Dim wSalaryYM As String = ""    '年月
    Dim wEmpID As String = ""        '指定人
    Dim wUserID As String = ""      'UserID

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Response.Cookies("PGM").Value = "HR_OverTimeReport03.aspx"
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
        NowDateTime = CStr(DateTime.Now.Date) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)       '現在日時
        wSalaryYM = Request.QueryString("pSalaryYM")  '指定年月
        wEmpID = Request.QueryString("pEmpID")        '指定人
        wUserID = Request.QueryString("pUserID")      'UserID
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選資料
    '**
    '*****************************************************************
    Sub DataList()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        SQL = "SELECT "
        SQL = SQL + "HRWDivName + '(' + HRWDivid + ')' as DivisionDesc, "
        SQL = SQL + "Name  + '(' + EmpID + ')' as NameDesc, "
        SQL = SQL + "Convert(VARCHAR(10), OverTimeDate, 111) as YMDDesc, "

        '日期URL
        SQL = SQL + "'http://10.245.1.10/WorkFlow/HR_OverTimeSheet_02.aspx?' + "
        SQL = SQL + "'&pFormNo=' + FormNo + "
        SQL = SQL + "'&pFormSno=' + str(FormSno) "
        SQL = SQL + " As YMDURL, "

        SQL = SQL + "FoodAmt as FoodDesc, "
        SQL = SQL + "TrafficAmt as TrafficDesc, "
        '核定分數
        SQL = SQL + "FAH1*60 + FAH2*60 + FBH*60 + FCH*60 + FAM1 + FAM2 + FBM + FCM as Total, "
        '換休分數
        SQL = SQL + "CVTime, "

        '2內分數
        SQL = SQL + "FAH1*60 + FAM1 As TwoIn, "
        '2外分數
        SQL = SQL + "FAH2*60 + FAM2 As TwoOut, "
        '假日分數
        SQL = SQL + "FBH*60 + FBM As Vacation, "
        '國定假日分數
        SQL = SQL + "FCH*60 + FCM As CVacation "
        SQL = SQL + "FROM V_OverTimeSheet_02 "
        '篩選條件Joo
        SQL = SQL + "Where Sts = '1' "
        SQL = SQL + "  And SalaryYM = '" + wSalaryYM + "' "
        SQL = SQL + "  And EmpID = '" + wEmpID + "' "
        SQL = SQL + "Order by OverTimeDate "

        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "OverTime")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()

        GridView2.DataSource = DBDataSet1
        GridView2.DataBind()

        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     合計處理
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.Header Then
            SumFood = 0      '合計伙食 
            SumTraffic = 0   '合計交通 
            SumTotal = 0      '合計核定時 
            SumCVTotal = 0   '合計換休時 

            SumTwoIn = 0      '合計2內時 
            SumTwoOut = 0     '合計2外時 
            SumVacation = 0   '合計假日時 
            SumCVacation = 0  '合計國定假日時 
        End If
        If e.Row.RowType = DataControlRowType.DataRow Then
            '伙食
            SumFood = SumFood + CInt(e.Row.Cells(3).Text.ToString)
            '交通
            SumTraffic = SumTraffic + CInt(e.Row.Cells(4).Text.ToString)
            '核定時數
            e.Row.Cells(5).Text = FormatNumber(CInt(e.Row.Cells(5).Text.ToString) / 60, 1)
            SumTotal = SumTotal + CDbl(e.Row.Cells(5).Text.ToString)
            '換休時數
            e.Row.Cells(6).Text = FormatNumber(CInt(e.Row.Cells(6).Text.ToString) / 60, 1)
            SumCVTotal = SumCVTotal + CDbl(e.Row.Cells(6).Text.ToString)
            '2內時數
            e.Row.Cells(7).Text = FormatNumber(CInt(e.Row.Cells(7).Text.ToString) / 60, 1)
            SumTwoIn = SumTwoIn + CDbl(e.Row.Cells(7).Text.ToString)
            '2外時數
            e.Row.Cells(8).Text = FormatNumber(CInt(e.Row.Cells(8).Text.ToString) / 60, 1)
            SumTwoOut = SumTwoOut + CDbl(e.Row.Cells(8).Text.ToString)
            '假日時數
            e.Row.Cells(9).Text = FormatNumber(CInt(e.Row.Cells(9).Text.ToString) / 60, 1)
            SumVacation = SumVacation + CDbl(e.Row.Cells(9).Text.ToString)
            '國定假日時數
            e.Row.Cells(10).Text = FormatNumber(CInt(e.Row.Cells(10).Text.ToString) / 60, 1)
            SumCVacation = SumCVacation + CDbl(e.Row.Cells(10).Text.ToString)
        End If
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(2).Text = "合計"
            e.Row.Cells(3).Text = FormatNumber(SumFood, 0)
            e.Row.Cells(4).Text = FormatNumber(SumTraffic, 0)
            e.Row.Cells(5).Text = FormatNumber(SumTotal, 1)
            e.Row.Cells(6).Text = FormatNumber(SumCVTotal, 1)
            e.Row.Cells(7).Text = FormatNumber(SumTwoIn, 1)
            e.Row.Cells(8).Text = FormatNumber(SumTwoOut, 1)
            e.Row.Cells(9).Text = FormatNumber(SumVacation, 1)
            e.Row.Cells(10).Text = FormatNumber(SumCVacation, 1)
        End If
    End Sub
    '*****************************************************************
    '**
    '**     合計處理(Excel)
    '**
    '*****************************************************************
    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        If e.Row.RowType = DataControlRowType.Header Then
            SumFood = 0      '合計伙食 
            SumTraffic = 0   '合計交通 
            SumTotal = 0      '合計核定時 
            SumCVTotal = 0   '合計換休時 

            SumTwoIn = 0      '合計2內時 
            SumTwoOut = 0     '合計2外時 
            SumVacation = 0   '合計假日時 
            SumCVacation = 0  '合計國定假日時 
        End If
        If e.Row.RowType = DataControlRowType.DataRow Then
            '伙食
            SumFood = SumFood + CInt(e.Row.Cells(3).Text.ToString)
            '交通
            SumTraffic = SumTraffic + CInt(e.Row.Cells(4).Text.ToString)
            '核定時數
            e.Row.Cells(5).Text = FormatNumber(CInt(e.Row.Cells(5).Text.ToString) / 60, 1)
            SumTotal = SumTotal + CDbl(e.Row.Cells(5).Text.ToString)
            '換休時數
            e.Row.Cells(6).Text = FormatNumber(CInt(e.Row.Cells(6).Text.ToString) / 60, 1)
            SumTotal = SumTotal + CDbl(e.Row.Cells(6).Text.ToString)
            '2內時數
            e.Row.Cells(7).Text = FormatNumber(CInt(e.Row.Cells(7).Text.ToString) / 60, 1)
            SumTwoIn = SumTwoIn + CDbl(e.Row.Cells(7).Text.ToString)
            '2外時數
            e.Row.Cells(8).Text = FormatNumber(CInt(e.Row.Cells(8).Text.ToString) / 60, 1)
            SumTwoOut = SumTwoOut + CDbl(e.Row.Cells(8).Text.ToString)
            '假日時數
            e.Row.Cells(9).Text = FormatNumber(CInt(e.Row.Cells(9).Text.ToString) / 60, 1)
            SumVacation = SumVacation + CDbl(e.Row.Cells(9).Text.ToString)
            '國定假日時數
            e.Row.Cells(10).Text = FormatNumber(CInt(e.Row.Cells(10).Text.ToString) / 60, 1)
            SumCVacation = SumCVacation + CDbl(e.Row.Cells(10).Text.ToString)
        End If
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(2).Text = "合計"
            e.Row.Cells(3).Text = FormatNumber(SumFood, 0)
            e.Row.Cells(4).Text = FormatNumber(SumTraffic, 0)
            e.Row.Cells(5).Text = FormatNumber(SumTotal, 1)
            e.Row.Cells(6).Text = FormatNumber(SumCVTotal, 1)
            e.Row.Cells(7).Text = FormatNumber(SumTwoIn, 1)
            e.Row.Cells(8).Text = FormatNumber(SumTwoOut, 1)
            e.Row.Cells(9).Text = FormatNumber(SumVacation, 1)
            e.Row.Cells(10).Text = FormatNumber(SumCVacation, 1)
        End If
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
        Response.AppendHeader("Content-Disposition", "attachment;filename=HR_OverTimeReport_Day.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=BIG5>")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)

        GridView2.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()
    End Sub

End Class
