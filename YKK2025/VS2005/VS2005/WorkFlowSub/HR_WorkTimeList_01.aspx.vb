Imports System.Data
Imports System.Data.OleDb

Partial Class HR_WorkTimeList_01
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
    Dim wFormNo As String          '表單
    Dim wDate As String            '出勤日
    Dim wEmpID As String           'EmpID


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()
        If Not Me.IsPostBack Then
            DWorkDate.Text = "出勤日：" + wDate
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
        wFormNo = Request.QueryString("pFormNo")  '表單
        wDate = Request.QueryString("pDate")      '出勤日
        wEmpID = Request.QueryString("pID")          'EmpID
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選加班資料
    '**
    '*****************************************************************
    Sub ShowData()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()

        DBDataSet1.Clear()

        SQL = ""
        '加班
        If wFormNo = "001001" Then
            GridView1.Columns.Item(0).HeaderText = "加班單No."
            GridView1.Columns.Item(1).HeaderText = ""
            GridView1.Columns.Item(2).HeaderText = "時數"

            SQL = "SELECT "
            SQL = SQL + "No, '' as Field1, "
            SQL = SQL + "'http://10.245.1.10/WorkFlow/HR_OverTimeSheet_02.aspx?' + "
            SQL = SQL + "'&pFormNo=' + FormNo + "
            SQL = SQL + "'&pFormSno=' + str(FormSno) "
            SQL = SQL + " As URL, "
            SQL = SQL + "FAH1*60 + FAH2*60 + FBH*60 + FCH*60 + FAM1 + FAM2 + FBM + FCM as Time "
            SQL = SQL + "FROM F_OverTimeSheet "
            SQL = SQL + "Where EmpID  = '" & wEmpID & "' "
            SQL = SQL + "  And Sts    = '1' "
            SQL = SQL + "  And OverTimeDate = '" & wDate & "' "
            SQL = SQL + "Order by OverTimeDate, No "
        End If

        '請假
        If wFormNo = "001002" Then
            GridView1.Columns.Item(0).HeaderText = "請假單No."
            GridView1.Columns.Item(1).HeaderText = "假別"
            GridView1.Columns.Item(2).HeaderText = "日數"

            SQL = "SELECT "
            SQL = SQL + "No, Vacation As Field1, "
            SQL = SQL + "'http://10.245.1.10/WorkFlow/HR_TimeOffSheet_02.aspx?' + "
            SQL = SQL + "'&pFormNo=' + FormNo + "
            SQL = SQL + "'&pFormSno=' + str(FormSno) "
            SQL = SQL + " As URL, "
            SQL = SQL + " ADays*8 as Time "
            SQL = SQL + "FROM F_TimeOffSheet "
            SQL = SQL + "Where EmpID  = '" & wEmpID & "' "
            SQL = SQL + "  And Sts    = '1' "
            SQL = SQL + "  And AStartDate <= '" & wDate & "' "
            SQL = SQL + "  And AEndDate   >= '" & wDate & "' "
            SQL = SQL + "Order by AStartDate, No "
        End If

        '外出
        If wFormNo = "001003" Then
            GridView1.Columns.Item(0).HeaderText = "外出單No."
            GridView1.Columns.Item(1).HeaderText = "目的地"
            GridView1.Columns.Item(2).HeaderText = "日數"

            SQL = "SELECT "
            SQL = SQL + "No, Place as Field1, "
            SQL = SQL + "'http://10.245.1.10/WorkFlow/HR_AWaySheet_02.aspx?' + "
            SQL = SQL + "'&pFormNo=' + FormNo + "
            SQL = SQL + "'&pFormSno=' + str(FormSno) "
            SQL = SQL + " As URL, "
            SQL = SQL + " ADay*8+AHour as Time "
            SQL = SQL + "FROM F_AWaySheet "
            SQL = SQL + "Where EmpID  = '" & wEmpID & "' "
            SQL = SQL + "  And Sts    = '1' "
            SQL = SQL + "  And AStartDate <= '" & wDate & "' "
            SQL = SQL + "  And AEndDate   >= '" & wDate & "' "
            SQL = SQL + "Order by AStartDate, No "
        End If

        '補工
        If wFormNo = "001004" Then
            GridView1.Columns.Item(0).HeaderText = "補工單No."
            GridView1.Columns.Item(1).HeaderText = ""
            GridView1.Columns.Item(2).HeaderText = "時數"

            SQL = "SELECT "
            SQL = SQL + "No, '' as Field1, "
            SQL = SQL + "'http://10.245.1.10/WorkFlow/HR_AddWorkSheet_02.aspx?' + "
            SQL = SQL + "'&pFormNo=' + FormNo + "
            SQL = SQL + "'&pFormSno=' + str(FormSno) "
            SQL = SQL + " As URL, "
            SQL = SQL + " AH*60+AM as Time "
            SQL = SQL + "FROM F_AddWorkSheet "
            SQL = SQL + "Where EmpID  = '" & wEmpID & "' "
            SQL = SQL + "  And Sts    = '1' "
            SQL = SQL + "  And WorkDate = '" & wDate & "' "
            SQL = SQL + "Order by WorkDate, No "
        End If

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Data")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()

        OleDbConnection1.Close()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then

            If wFormNo = "001001" Then
                e.Row.Cells(2).Text = FormatNumber(CInt(e.Row.Cells(2).Text.ToString) / 60, 1)
            End If
            If wFormNo = "001002" Then
                e.Row.Cells(2).Text = FormatNumber(CInt(e.Row.Cells(2).Text.ToString) / 8, 1)
            End If
            If wFormNo = "001003" Then
                e.Row.Cells(2).Text = FormatNumber(CInt(e.Row.Cells(2).Text.ToString) / 8, 1)
            End If
            If wFormNo = "001004" Then
                e.Row.Cells(2).Text = FormatNumber(CInt(e.Row.Cells(2).Text.ToString) / 60, 1)
            End If

        End If
    End Sub

End Class
