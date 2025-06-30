Imports System.Data
Imports System.Data.OleDb

Partial Class HR_OverTimeReport01
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

    Dim wLevel As String = ""       '篩選Level
    Dim wEmpID As String = ""       '指定人
    Dim wUserID As String = ""      'UserID

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Response.Cookies("PGM").Value = "HR_OverTimeReport01.aspx"
        SetParameter()          '設定共用參數

        If Not Me.IsPostBack Then
            SetSearchItem(True)
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

        wUserID = Request.QueryString("pUserID")      'UserID
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定篩選條件
    '**
    '*****************************************************************
    Sub SetSearchItem(ByVal First As Boolean)
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()

        '取得篩選權限
        SQL = "Select * From M_Referp  "
        SQL = SQL + "Where Cat='1999'  "
        SQL = SQL + "  and DKey='" & "AUTHORITY-" & wUserID & "' "
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

        If First Then
            DBDataSet1.Clear()
            DDivision.Items.Clear()
            If wLevel = "PERSON" Then
                '取得個人資訊
                SQL = "Select * From M_Users  "
                SQL = SQL + "Where UserID='" & wUserID & "' "
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                If DBTable1.Rows.Count > 0 Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("HRWDivName")
                    ListItem1.Value = DBTable1.Rows(i).Item("HRWDivName")
                    DDivision.Items.Add(ListItem1)
                    wEmpID = DBTable1.Rows(i).Item("EmpID")
                End If
            Else
                If wLevel = "DIVISION" Then
                    '取得所指定部門
                    SQL = "Select * From M_Referp  "
                    SQL = SQL + "Where Cat='1999'  "
                    SQL = SQL + "  and DKey='" & "DIVISION-" & wUserID & "' "
                    Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter2.Fill(DBDataSet1, "M_Referp")
                    DBTable1 = DBDataSet1.Tables("M_Referp")
                    For i = 0 To DBTable1.Rows.Count - 1
                        Dim ListItem1 As New ListItem
                        ListItem1.Text = DBTable1.Rows(i).Item("Data")
                        ListItem1.Value = DBTable1.Rows(i).Item("Data")
                        DDivision.Items.Add(ListItem1)
                    Next
                Else
                    '取得全部部門
                    If wLevel = "ALL" Then
                        SQL = "Select HRWDivName From M_Users Group by HRWDivName Order by HRWDivName "
                        DDivision.Items.Add("ALL")
                        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                        DBAdapter2.Fill(DBDataSet1, "M_Users")
                        DBTable1 = DBDataSet1.Tables("M_Users")
                        For i = 0 To DBTable1.Rows.Count - 1
                            Dim ListItem1 As New ListItem
                            ListItem1.Text = DBTable1.Rows(i).Item("HRWDivName")
                            ListItem1.Value = DBTable1.Rows(i).Item("HRWDivName")
                            DDivision.Items.Add(ListItem1)
                        Next
                    End If
                End If
            End If

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
        End If
        OleDbConnection1.Close()
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

        If wLevel = "PERSON" Then
            SQL = SQL + "Name  + '(' + EmpID + ')' as NameDesc, "
        Else
            SQL = SQL + "'' as NameDesc, "
        End If

        SQL = SQL + "SalaryYM as YMDesc, "
        SQL = SQL + "Sum(FoodAmt) as FoodDesc, "
        SQL = SQL + "Sum(TrafficAmt) as TrafficDesc, "
        '核定分數
        SQL = SQL + "(Sum(FAH1)*60 + Sum(FAH2)*60 + Sum(FBH)*60 + Sum(FCH)*60 + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM)) as Total, "
        '換休分數
        SQL = SQL + "Sum(CVTime) As CVTime, "
        '2內分數
        SQL = SQL + "(Sum(FAH1)*60 + Sum(FAM1)) As TwoIn, "
        '2外分數
        SQL = SQL + "(Sum(FAH2)*60 + Sum(FAM2)) As TwoOut, "
        '假日分數
        SQL = SQL + "(Sum(FBH)*60 + Sum(FBM)) As Vacation, "
        '國定假日分數
        SQL = SQL + "(Sum(FCH)*60 + Sum(FCM)) As CVacation, "
        '部門URL
        If wLevel = "ALL" Or wLevel = "DIVISION" Then
            SQL = SQL + "'HR_OverTimeReport02.aspx?' + "
            SQL = SQL + "'&pYear=' + SubString(SalaryYM,1,4) + "
            SQL = SQL + "'&pMonth=' + SubString(SalaryYM,6,2) + "
            SQL = SQL + "'&pLevel=" + wLevel + "' + "
            SQL = SQL + "'&pDivision=' + HRWDivID + "
            SQL = SQL + "'&pUserID=" + wUserID + "' "
            SQL = SQL + " As DivisionURL, "
            '姓名URL
            SQL = SQL + "'' As NameURL "
        Else
            SQL = SQL + "'' As DivisionURL, "
            '姓名URL
            SQL = SQL + "'HR_OverTimeReport03.aspx?' + "
            SQL = SQL + "'&pSalaryYM=' + SalaryYM + "
            SQL = SQL + "'&pEmpID=' + EmpID + "
            SQL = SQL + "'&pUserID=" + wUserID + "' "
            SQL = SQL + " As NameURL "
        End If
        SQL = SQL + "FROM V_OverTimeSheet_02 "
        '篩選條件Joo
        SQL = SQL + "Where Sts = '1' "
        SQL = SQL + "  And SalaryYM = '" + DYear.SelectedValue + "/" + DMonth.SelectedValue + "' "

        If DDivision.SelectedValue <> "ALL" Then
            SQL = SQL + " And HRWDivName = '" + DDivision.SelectedValue + "' "
        End If

        If wLevel = "ALL" Or wLevel = "DIVISION" Then
            SQL = SQL + "Group by SalaryYM, HRWDivName, HRWDivid "
            SQL = SQL + "Order by SalaryYM, HRWDivName, HRWDivid "
        End If

        If wLevel = "PERSON" Then
            SQL = SQL + "And EmpID = '" + wEmpID + "' "
            SQL = SQL + "Group by SalaryYM, HRWDivName, HRWDivid, Name, Empid "
            SQL = SQL + "Order by SalaryYM, HRWDivName, HRWDivid, Name, Empid "
        End If

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
            SumTotal = 0     '合計核定時 
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
    '---------------------------------------------------------------------------------------------------
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
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Click Go Button
    '**
    '*****************************************************************
    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()

        '取得篩選權限
        SQL = "Select * From M_Referp  "
        SQL = SQL + "Where Cat='1999'  "
        SQL = SQL + "  and DKey='" & "AUTHORITY-" & wUserID & "' "
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

        If wLevel = "PERSON" Then
            '取得個人資訊
            SQL = "Select * From M_Users  "
            SQL = SQL + "Where UserID='" & wUserID & "' "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "M_Users")
            DBTable1 = DBDataSet1.Tables("M_Users")
            If DBTable1.Rows.Count > 0 Then
                wEmpID = DBTable1.Rows(i).Item("EmpID")
            End If
        End If
        OleDbConnection1.Close()

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
        Response.AppendHeader("Content-Disposition", "attachment;filename=HR_OverTimeReport_Division.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")

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
