Imports System.Data
Imports System.Data.OleDb

Partial Class OverTimeList01
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
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選加班資料
    '**
    '*****************************************************************
    Sub OverTimeData()
        Dim wStartDate As String = wYear + "/1/1"
        Dim wEndDate As String = wYear + "/12/31"
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
        SQL = SQL + "FAH1*60 + FAH2*60 + FBH*60 + FCH*60 + FAM1 + FAM2 + FBM + FCM as Total "  '核定分數
        SQL = SQL + "FROM F_OverTimeSheet "

        SQL = SQL + "Where EmpID  = '" & wEmpID & "' "
        SQL = SQL + "  And OverTimeDate >= '" & wStartDate & "' "
        SQL = SQL + "  And OverTimeDate <= '" & wEndDate & "' "
        SQL = SQL + "  And CVacation Like '2%' "
        SQL = SQL + "  And Sts       = '1' "
        SQL = SQL + "Order by OverTimeDate Desc "

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "OverTime")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()
        '請假
        SQL = "SELECT "
        SQL = SQL + "Convert(VARCHAR(10), AStartDate, 111) + ' ' + str(AStartH) + ' ~ ' + "
        SQL = SQL + "Convert(VARCHAR(10), AEndDate, 111)   + ' ' + str(AEndH)   As AVacationDate, "
        SQL = SQL + "'http://10.245.1.10/WorkFlow/HR_TimeOffSheet_02.aspx?' + "
        SQL = SQL + "'&pFormNo=' + FormNo + "
        SQL = SQL + "'&pFormSno=' + str(FormSno) "
        SQL = SQL + " As VacationURL, "
        SQL = SQL + "No, "
        SQL = SQL + "ADays As Total "  '核定分數
        SQL = SQL + "FROM F_TimeOffSheet "

        SQL = SQL + "Where EmpID  = '" & wEmpID & "' "
        SQL = SQL + "  And AStartDate >= '" & wStartDate & "' "
        SQL = SQL + "  And AStartDate <= '" & wEndDate & "' "
        SQL = SQL + "  And VacationCode = 'R' "
        SQL = SQL + "  And Sts       = '1' "
        SQL = SQL + "Order by Date Desc "

        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "Vacation")
        GridView2.DataSource = DBDataSet1
        GridView2.DataBind()

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
    '**     合計處理-請假
    '**
    '*****************************************************************
    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            e.Row.Cells(2).Text = FormatNumber(CInt(e.Row.Cells(3).Text.ToString) / 60, 1)
            SumOverTime = SumOverTime + CInt(e.Row.Cells(3).Text.ToString)
        End If
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(1).Text = "合計"
            e.Row.Cells(2).Text = FormatNumber(SumOverTime, 1)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     關閉
    '**
    '*****************************************************************
    Protected Sub BClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.Click
        Dim i As Integer

        For i = 0 To GridView1.Rows.Count - 1
            Dim cb As CheckBox = CType(GridView1.Rows(i).FindControl("CheckBox1"), CheckBox)
            If cb.Checked Then
                MsgBox(GridView1.Rows(i).Cells(2).Text)
            End If
        Next
    End Sub
End Class
