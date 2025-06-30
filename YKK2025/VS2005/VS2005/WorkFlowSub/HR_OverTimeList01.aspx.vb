Imports System.Data
Imports System.Data.OleDb

Partial Class HR_OverTimeList01
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
    Dim wYear As String            '工作年
    Dim wEmpID As String           'Emp-ID
    Dim wDepoID As String           '公司別
    Dim SumOverTime As Double = 0  '加班總時數
    Dim SumVacation As Double = 0  '請假總時數

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
        wYear = Left(Request.QueryString("pMonth"), 4)  '工作年
        wEmpID = Request.QueryString("pEmpID")          'Emp-ID
        wDepoID = Request.QueryString("pDepoID")        '公司別
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選加班資料
    '**
    '*****************************************************************
    Sub OverTimeData()
        Dim wStartDate As String = ""
        Dim wEndDate As String = ""
        If CInt(Now.Date.Month.ToString) < 4 Then
            wStartDate = CStr(CInt(Now.Date.Year.ToString) - 1) + "/4/1"
            wEndDate = Now.Date.Year.ToString + "/3/31"
        Else
            wStartDate = Now.Date.Year.ToString + "/4/1"
            wEndDate = CStr(CInt(Now.Date.Year.ToString) + 1) + "/3/31"
        End If

        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        '加班
        SQL = "SELECT "
        SQL = SQL + "Convert(VARCHAR(10), OverTimeDate, 111) As OverTimeDateDesc, "
        SQL = SQL + "'http://10.245.1.6/WorkFlow/HR_OverTimeSheet_02.aspx?' + "
        SQL = SQL + "'&pFormNo=' + FormNo + "
        SQL = SQL + "'&pFormSno=' + str(FormSno) "
        SQL = SQL + " As OverTimeURL, "
        SQL = SQL + "No, "
        SQL = SQL + "FAH1*60 + FAH2*60 + FBH*60 + FCH*60 + FAM1 + FAM2 + FBM + FCM as Total "  '核定分數
        SQL = SQL + "FROM F_OverTimeSheet "

        SQL = SQL + "Where EmpID  = '" & wEmpID & "' "
        SQL = SQL + "  And DepoCode = '" & wDepoID & "' "
        SQL = SQL + "  And OverTimeDate >= '" & wStartDate & "' "
        SQL = SQL + "  And OverTimeDate <= '" & wEndDate & "' "
        SQL = SQL + "  And CVacation Like '2%' "
        SQL = SQL + "  And Sts       = '1' "
        SQL = SQL + "Order by OverTimeDate Desc "

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
            SumOverTime = SumOverTime + CInt(e.Row.Cells(3).Text.ToString)
        End If
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(2).Text = "合計"
            e.Row.Cells(3).Text = FormatNumber(SumOverTime, 1)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     關閉
    '**
    '*****************************************************************
    Protected Sub BClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.Click
        Dim i, j As Integer
        Dim wNo(5) As String
        Dim Cmd As String
        Dim wHour(5), wTotal As Double
        '初值
        For j = 1 To 5
            wNo(j) = ""
            wHour(j) = 0
        Next
        wTotal = 0
        '篩選資料
        j = 1
        For i = 0 To GridView1.Rows.Count - 1
            Dim cb As CheckBox = CType(GridView1.Rows(i).FindControl("CheckBox1"), CheckBox)
            If cb.Checked Then
                wNo(j) = GridView1.Rows(i).Cells(2).Text
                wHour(j) = CDbl(GridView1.Rows(i).Cells(3).Text)
                wTotal = wTotal + wHour(j)
                j = j + 1
            End If
        Next
        'Return Main Form
        Cmd = String.Format("<script>window.opener.document.{0}.value = '{1:d}'; " + _
                                    "window.opener.document.{2}.value = '{3:d}'; " + _
                                    "window.opener.document.{4}.value = '{5:d}'; " + _
                                    "window.opener.document.{6}.value = '{7:d}'; " + _
                                    "window.opener.document.{8}.value = '{9:d}'; " + _
                                    "window.opener.document.{10}.value = '{11:d}'; " + _
                                    "window.opener.document.{12}.value = '{13:d}'; " + _
                                    "window.opener.document.{14}.value = '{15:d}'; " + _
                                    "window.opener.document.{16}.value = '{17:d}'; " + _
                                    "window.opener.document.{18}.value = '{19:d}'; " + _
                                    "window.opener.document.{20}.value = '{21:d}'; " + _
                                    "window.close(); " + _
                            "</script>", "Form1.DOTNo1", wNo(1), "Form1.DOTHours1", FormatNumber(wHour(1), 1), "Form1.DOTNo2", wNo(2), "Form1.DOTHours2", FormatNumber(wHour(2), 1), "Form1.DOTNo3", wNo(3), "Form1.DOTHours3", FormatNumber(wHour(3), 1), "Form1.DOTNo4", wNo(4), "Form1.DOTHours4", FormatNumber(wHour(4), 1), "Form1.DOTNo5", wNo(5), "Form1.DOTHours5", FormatNumber(wHour(5), 1), "Form1.DOTHours", FormatNumber(wTotal, 1))
        Response.Write(Cmd)
    End Sub

End Class
