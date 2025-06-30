Imports System.Data
Imports System.Data.OleDb

Partial Class HR_OverTimeAndTimeOffList
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
    Dim SumTotal As Double = 0      '合計核定時 
    Dim wYear As String           '工作年
    Dim wEmpID As String = ""        '指定人
    Dim wUserID As String = ""      'UserID
    Dim wDepoID As String           '公司別

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Response.Cookies("PGM").Value = "HR_OverTimeAndTimeOffList.aspx"
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
        wYear = Left(Request.QueryString("pMonth"), 4)  '工作年
        wEmpID = Request.QueryString("pEmpID")        '指定人
        wUserID = Request.QueryString("pUserID")      'UserID
        wDepoID = Request.QueryString("pDepoID")        '公司別
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選資料
    '**
    '*****************************************************************
    Sub DataList()
        Dim i As Integer = 0
        Dim wStartDate As String = ""

        If CInt(Now.Date.Month.ToString) < 4 Then
            wStartDate = CStr(CInt(Now.Date.Year.ToString) - 1) + "/4/1"
        Else
            wStartDate = Now.Date.Year.ToString + "/4/1"
        End If

        Dim SQL As String
        Dim DBDataSet1, DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()

        '加班資料
        SQL = "SELECT "
        SQL = SQL + "Name  + '(' + EmpID + ')' as NameDesc, "
        SQL = SQL + "Convert(VARCHAR(10), OverTimeDate, 111) as OverTimeDateDesc, "
        SQL = SQL + "CVacation as CVacationDesc, "
        SQL = SQL + "FAH1*60 + FAH2*60 + FBH*60 + FCH*60 + FAM1 + FAM2 + FBM + FCM as Total "
        SQL = SQL + "FROM V_OverTimeSheet_02 "
        SQL = SQL + "Where Sts = '1' "
        SQL = SQL + "  And OverTimeDate >= '" & wStartDate & "' "
        SQL = SQL + "  And DepoCode = '" & wDepoID & "' "
        SQL = SQL + "  And EmpID = '" + wEmpID + "' "
        SQL = SQL + "  And CVacation Like '2.%' "
        SQL = SQL + "Order by OverTimeDate "

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "OverTime")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()

        '請假資料
        SQL = "SELECT "
        SQL = SQL + "Name  + '(' + EmpID + ')' as NameDesc, "
        SQL = SQL + "Vacation As VacationDesc, "
        SQL = SQL + "Convert(VARCHAR(10), AStartDate, 111) + ' ' + str(AStartH) + ':00 ~ ' + "
        SQL = SQL + "Convert(VARCHAR(10), AEndDate, 111)   + ' ' + str(AEndH) + ':00'  As VacationDateDesc, "
        SQL = SQL + "ADays As Total "
        SQL = SQL + "FROM F_TimeOffSheet "
        SQL = SQL + "Where Sts = '1' "
        SQL = SQL + "  And AStartDate >= '" & wStartDate & "' "
        SQL = SQL + "  And DepoCode = '" & wDepoID & "' "
        SQL = SQL + "  And EmpID  =  '" & wEmpID & "' "
        SQL = SQL + "  And VacationCode  = '9' "
        SQL = SQL + "Order by AStartDate "

        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "TimeOff")
        GridView2.DataSource = DBDataSet2
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
            SumTotal = 0      '合計核定時 
        End If
        If e.Row.RowType = DataControlRowType.DataRow Then
            '核定時數
            e.Row.Cells(3).Text = FormatNumber(CInt(e.Row.Cells(3).Text.ToString) / 60, 1)
            SumTotal = SumTotal + CDbl(e.Row.Cells(3).Text.ToString)
        End If
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(2).Text = "合計"
            e.Row.Cells(3).Text = FormatNumber(SumTotal, 1)
        End If
    End Sub
    '*****************************************************************
    '**
    '**     合計處理(Excel)
    '**
    '*****************************************************************
    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        If e.Row.RowType = DataControlRowType.Header Then
            SumTotal = 0      '合計核定時 
        End If
        If e.Row.RowType = DataControlRowType.DataRow Then
            '核定時數
            e.Row.Cells(3).Text = FormatNumber(CInt(e.Row.Cells(3).Text.ToString), 1)
            SumTotal = SumTotal + CDbl(e.Row.Cells(3).Text.ToString)
        End If
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(2).Text = "合計"
            e.Row.Cells(3).Text = FormatNumber(SumTotal, 1)
        End If
    End Sub

End Class
