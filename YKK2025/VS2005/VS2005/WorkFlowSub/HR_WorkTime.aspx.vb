Imports System.Data
Imports System.Data.OleDb

Partial Class HR_WorkTime
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim oCommon As New Common.CommonService

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim SumOverTime As Double = 0  '合計加班 
    Dim SumTimeOff As Double = 0   '合計請假 
    Dim SumAway As Double = 0      '合計外出 
    Dim SumAddWork As Double = 0   '合計補工 
    Dim wLevel As String = ""
    Dim NowDateTime As String       '現在日期時間
    'Add-Start by Joy 2009/11/20(2010行事曆對應)
    Dim wDepo As String = "TP1"      '群組行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
    'Add-End

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        Response.Cookies("PGM").Value = "HR_WorkTime.aspx"

        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            SetSearchItem()
            'DataList()
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
                      CStr(DateTime.Now.Second)     '現在日時

        'Add-Start by Joy 2009/11/20(2010行事曆對應)
        wDepo = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))
        'Add-End

        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub SetSearchItem()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        '**** 設定 部門/姓名
        OleDbConnection1.Open()

        '取得篩選權限
        SQL = "Select * From M_Referp  "
        SQL = SQL + "Where Cat='1999'  "
        SQL = SQL + "  and DKey='" & "AUTHORITY-" & Request.QueryString("pUserID") & "' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Referp")
        If DBDataSet1.Tables("M_Referp").Rows.Count > 0 Then
            If DBDataSet1.Tables("M_Referp").Rows(0).Item("Data") = "ALL" Then
                wLevel = "ALL"
            Else
                wLevel = "DIVISION"
            End If
        Else
            wLevel = "PERSON"
        End If

        DBDataSet1.Clear()
        DDivision.Items.Clear()
        DName.Items.Clear()
        If wLevel = "PERSON" Then
            '取得個人資訊
            SQL = "Select isnull(Name,'') as UserName, isnull(HRWDivName,'') as HRWDivName From V_WorkTime_01 "
            SQL = SQL + "Where UserID='" & Request.QueryString("pUserID") & "' "
            SQL = SQL + "  And Not HRWDivName is Null "
            SQL = SQL + "  And Not Name is Null "
            SQL = SQL + "Order by Name "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "M_Users")
            DBTable1 = DBDataSet1.Tables("M_Users")
            If DBTable1.Rows.Count > 0 Then
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(0).Item("HRWDivName")
                ListItem1.Value = DBTable1.Rows(0).Item("HRWDivName")
                DDivision.Items.Add(ListItem1)
                '姓名
                Dim ListItem2 As New ListItem
                ListItem2.Text = DBTable1.Rows(0).Item("UserName")
                ListItem2.Value = DBTable1.Rows(0).Item("UserName")
                DName.Items.Add(ListItem2)
            End If
        Else
            If wLevel = "DIVISION" Then
                '取得所指定部門
                SQL = "Select * From M_Referp  "
                SQL = SQL + "Where Cat='1999'  "
                SQL = SQL + "  and DKey='" & "DIVISION-" & Request.QueryString("pUserID") & "' "
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    DDivision.Items.Add(ListItem1)
                Next
                '取得個人資訊
                DBDataSet1.Clear()
                SQL = "Select isnull(Name,'') as UserName From V_WorkTime_01 "
                SQL = SQL + "Where HRWDivName = '" & DDivision.SelectedValue & "' "
                SQL = SQL + "  And Not HRWDivName is Null "
                SQL = SQL + "  And Not Name is Null "
                SQL = SQL + "Group by Name "
                SQL = SQL + "Order by Name "
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    '姓名
                    Dim ListItem2 As New ListItem
                    ListItem2.Text = DBTable1.Rows(i).Item("UserName")
                    ListItem2.Value = DBTable1.Rows(i).Item("UserName")
                    DName.Items.Add(ListItem2)
                Next
            Else
                '取得全部部門
                If wLevel = "ALL" Then
                    SQL = "Select isnull(HRWDivName,'') as HRWDivName From V_WorkTime_01 "
                    SQL = SQL + "Where Not HRWDivName is Null "
                    SQL = SQL + "  And Not Name is Null "
                    SQL = SQL + "Group by HRWDivName "
                    SQL = SQL + "Order by HRWDivName "
                    Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter2.Fill(DBDataSet1, "M_Users")
                    DBTable1 = DBDataSet1.Tables("M_Users")
                    For i = 0 To DBTable1.Rows.Count - 1
                        Dim ListItem1 As New ListItem
                        ListItem1.Text = DBTable1.Rows(i).Item("HRWDivName")
                        ListItem1.Value = DBTable1.Rows(i).Item("HRWDivName")
                        DDivision.Items.Add(ListItem1)
                    Next
                    '姓名
                    DBDataSet1.Clear()
                    SQL = "Select isnull(Name,'') as UserName From V_WorkTime_01  "
                    SQL = SQL + "Where HRWDivName = '" & DDivision.SelectedValue & "' "
                    SQL = SQL + "  And Not HRWDivName is Null "
                    SQL = SQL + "  And Not Name is Null "
                    SQL = SQL + "Group by Name "
                    SQL = SQL + "Order by Name "
                    Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter3.Fill(DBDataSet1, "M_Users")
                    DBTable1 = DBDataSet1.Tables("M_Users")
                    For i = 0 To DBTable1.Rows.Count - 1
                        '姓名
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("UserName")
                        ListItem2.Value = DBTable1.Rows(i).Item("UserName")
                        DName.Items.Add(ListItem2)
                    Next
                End If
            End If
        End If
        OleDbConnection1.Close()

        '**** 設定 加班年月
        Dim wYY As String = CStr(DateTime.Now.Year)
        Dim wMM As String = CStr(DateTime.Now.Month)
        '加班年
        DYear.Items.Clear()
        For i = 2008 To 2020
            Dim ListItem1 As New ListItem
            ListItem1.Text = CStr(i)
            ListItem1.Value = CStr(i)
            If ListItem1.Value = wYY Then ListItem1.Selected = True
            DYear.Items.Add(ListItem1)
        Next
        '加班月
        If CInt(wMM) < 10 Then wMM = "0" & wMM
        DMonth.Items.Clear()
        For i = 1 To 12
            Dim ListItem1 As New ListItem
            If i < 10 Then
                ListItem1.Text = "0" & CStr(i)
                ListItem1.Value = "0" & CStr(i)
            Else
                ListItem1.Text = CStr(i)
                ListItem1.Value = CStr(i)
            End If
            If ListItem1.Value = wMM Then ListItem1.Selected = True
            DMonth.Items.Add(ListItem1)
        Next

    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        SQL = "SELECT "
        SQL = SQL + "Convert(VARCHAR(10), CDate, 111) as DateDesc, WeekDesc, "
        SQL = SQL + "RTrim(Name)+'('+RTrim(EmpID)+')' as NameDesc, "
        SQL = SQL + "HRWDivName, JobName, TimeA, TimeB, Remark, "
        SQL = SQL + "OTDesc, OTURL, VADesc, VAURL, AWayDesc, AWayURL, AWorkDesc, AWorkURL, "
        SQL = SQL + "OT_Hours As SOverTime, ABS1 As STimeoff, ABS2 As SAway, ABS3 As SAddwork, "
        SQL = SQL + "ABS_FormSno1, ABS_FormSno2 "
        SQL = SQL + "FROM V_WorkTime_01 "
        '年月
        SQL = SQL + "Where SelectYM = '" + DYear.SelectedValue + "/" + DMonth.SelectedValue + "' "
        SQL = SQL + "  And pluralism = '0' "        '非兼職
        '部門
        If DDivision.SelectedValue <> "ALL" Then
            SQL = SQL + " And HRWDivName = '" + DDivision.SelectedValue + "'"
        End If
        '申請者
        If DName.SelectedValue <> "ALL" Then
            SQL = SQL + " And Name = '" + DName.SelectedValue + "'"
        End If
        'Sort
        SQL = SQL + " Order by HRWDivName, EmpID, CDate "

        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "WaitHandle")
        GridView1.DataSource = DBDataSet2
        GridView1.DataBind()

        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView--合併儲存格/設定背景顏色
    '**
    '*****************************************************************
    Protected Sub GridView1_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PreRender
        Dim mySingleRow As GridViewRow
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        '設定背景顏色
        For Each mySingleRow In GridView1.Rows
            DBDataSet1.Clear()
            SQL = "SELECT * FROM M_Vacation "

            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
            'SQL = SQL + "Where Depo = 'TP' "
            '
            SQL = SQL + "Where Depo = '" + wDepo + "' "
            'Modify-End

            SQL = SQL + "  And YMD  = '" + GridView1.Rows(CInt(mySingleRow.RowIndex)).Cells(2).Text.Trim() + "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "Vacation")
            If DBDataSet1.Tables("Vacation").Rows.Count > 0 Then
                If DBDataSet1.Tables("Vacation").Rows(0).Item("VacationType") = 1 Then
                    '國定假日
                    GridView1.Rows(CInt(mySingleRow.RowIndex)).BackColor = Drawing.Color.Cyan
                Else
                    '假日
                    GridView1.Rows(CInt(mySingleRow.RowIndex)).BackColor = Drawing.Color.LightPink
                End If
            End If
        Next
        OleDbConnection1.Close()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataList()
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
        Dim wAllowPaging As Boolean = GridView1.AllowPaging   '程式別不同

        GridView1.AllowPaging = False   '程式別不同
        DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=HR_WorkTimeList.xls")     '程式別不同

        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        GridView1.AllowPaging = wAllowPaging        '程式別不同
    End Sub

    Protected Sub DDivision_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDivision.SelectedIndexChanged
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        DName.Items.Clear()
        'DName.Items.Add("ALL")
        OleDbConnection1.Open()

        SQL = "Select isnull(name,'') as UserName From V_WorkTime_01 "
        SQL = SQL + "Where HRWDivName = '" & DDivision.SelectedValue & "' "
        SQL = SQL + "  And Not HRWDivName is Null "
        SQL = SQL + "  And Not Name is Null "
        SQL = SQL + "Group by Name "
        SQL = SQL + "Order by Name "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            '姓名
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("UserName")
            ListItem1.Value = DBTable1.Rows(i).Item("UserName")
            DName.Items.Add(ListItem1)
        Next

        OleDbConnection1.Close()

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     合計處理(Excel)
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        If e.Row.RowType = DataControlRowType.Header Then
            SumOverTime = 0
            SumTimeOff = 0
            SumAway = 0
            SumAddWork = 0
            e.Row.Cells(11).Visible = False
            e.Row.Cells(12).Visible = False
            e.Row.Cells(13).Visible = False
            e.Row.Cells(14).Visible = False
            e.Row.Cells(15).Visible = False
            e.Row.Cells(16).Visible = False
        End If
        If e.Row.RowType = DataControlRowType.DataRow Then
            If e.Row.Cells(2).Text < "2010/07/01" Then
                If e.Row.Cells(4).Text > "08:20" Then
                    e.Row.Cells(4).ForeColor = Drawing.Color.Red
                Else
                    e.Row.Cells(4).ForeColor = Drawing.Color.Black
                End If
            Else
                If e.Row.Cells(4).Text > "08:30" Then
                    e.Row.Cells(4).ForeColor = Drawing.Color.Red
                Else
                    e.Row.Cells(4).ForeColor = Drawing.Color.Black
                End If
            End If
            '加班
            SumOverTime = SumOverTime + CDbl(e.Row.Cells(11).Text.ToString)
            '請假
            If CDbl(e.Row.Cells(12).Text.ToString) > 0 Then
                DBDataSet1.Clear()
                OleDbConnection1.Open()
                SQL = "SELECT * FROM F_TimeOffSheet "
                SQL = SQL + "Where Sts = '1' "
                SQL = SQL + "  And Name =  '" + DName.SelectedValue + "' "
                SQL = SQL + "  And AStartDate = '" + e.Row.Cells(2).Text.ToString + "'"
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "Vacation")
                If DBDataSet1.Tables("Vacation").Rows.Count > 0 Then
                    SumTimeOff = SumTimeOff + CDbl(e.Row.Cells(12).Text.ToString)
                Else
                    e.Row.Cells(7).Text = "''"
                End If
                OleDbConnection1.Close()
            End If
            '外出  
            If CDbl(e.Row.Cells(13).Text.ToString) > 0 Then
                DBDataSet1.Clear()
                OleDbConnection1.Open()
                SQL = "SELECT * FROM F_AWaySheet "
                SQL = SQL + "Where Sts = '1' "
                SQL = SQL + "  And Name =  '" + DName.SelectedValue + "' "
                SQL = SQL + "  And AStartDate = '" + e.Row.Cells(2).Text.ToString + "'"
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "AWay")
                If DBDataSet1.Tables("AWay").Rows.Count > 0 Then
                    SumAway = SumAway + CDbl(e.Row.Cells(13).Text.ToString)
                Else
                    e.Row.Cells(8).Text = "''"
                End If
                OleDbConnection1.Close()
            End If
            '補工
            SumAddWork = SumAddWork + CDbl(e.Row.Cells(14).Text.ToString)

            e.Row.Cells(11).Visible = False
            e.Row.Cells(12).Visible = False
            e.Row.Cells(13).Visible = False
            e.Row.Cells(14).Visible = False
            e.Row.Cells(15).Visible = False
            e.Row.Cells(16).Visible = False
        End If
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(5).Text = "合計"
            e.Row.Cells(6).Text = FormatNumber(SumOverTime, 1) & " 時"
            e.Row.Cells(7).Text = FormatNumber(SumTimeOff, 1) & " 日"
            e.Row.Cells(8).Text = FormatNumber(SumAway, 1) & " 日"
            e.Row.Cells(9).Text = FormatNumber(SumAddWork, 1) & " 時"
            e.Row.Cells(11).Visible = False
            e.Row.Cells(12).Visible = False
            e.Row.Cells(13).Visible = False
            e.Row.Cells(14).Visible = False
            e.Row.Cells(15).Visible = False
            e.Row.Cells(16).Visible = False
        End If
    End Sub
End Class
