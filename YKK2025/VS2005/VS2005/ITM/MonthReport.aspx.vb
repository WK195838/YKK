Imports System.Data
Imports System.Data.OleDb

Partial Class MonthReport
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_Class   'YKK SPD系共通涵式
    Dim fpObj As New ForProject             ' 操作db的物件
    'Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDate As Date       '現在日期時間
    Dim wUserID As String     'User ID
    Dim wNowYear As String     '當年
    Dim wNowMonth As Integer   '當月
    Dim wStartYear, wEndYear As String      '起迄年
    Dim wStartMonth, wEndMonth As Integer   '起迄月
    Dim wKeyField As String    '指定Key欄位
    Dim wKeyData As String     '指定Key值
    Dim wSearch As String      '指定Search值

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            SetFunHyperLink()      '設定機能連結
            ShowSchedule()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '設定Cookies及傳遞參數
        NowDate = DateTime.Now.Date                  '現在日期
        Response.Cookies("PGM").Value = "MonthReport.aspx"
        wUserID = Request.QueryString("pUserID")     'User ID
        wSearch = Request.QueryString("pSearch")     'Search
        wSearch = Replace(wSearch, "%20", "")
        If DSearch.Text = "" Then
            If wSearch <> "" Then
                DSearch.Text = wSearch
            End If
        End If

        '設定Key欄位
        If Request.QueryString("pKeyField") = "" Then
            wKeyField = "HeadDivision"
            wKeyData = "IT"
        Else
            wKeyField = Request.QueryString("pKeyField")
            wKeyData = Request.QueryString("pKeyData")
        End If
        '現在年及週別
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As New DataTable
        Dim SQL As String

        Dim OleDbConnection1 As New OleDbConnection
        Dim OleDBCommand1 As New OleDbCommand
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn1")  'SQL連結設定

        If Request.QueryString("pReportMonth") <> "" And Request.QueryString("pReportYear") <> "" Then
            wNowYear = Request.QueryString("pReportYear")
            wNowMonth = CInt(Request.QueryString("pReportMonth"))
        Else
            SQL = "Select ReportYear, ReportMonth From V_WeekReport_01 Where ReportDate = '" + NowDate + "' "
            OleDbConnection1.Open()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "WeekReport")
            DBTable1 = DBDataSet1.Tables("WeekReport")
            If DBTable1.Rows.Count > 0 Then
                wNowYear = DBTable1.Rows(0).Item("ReportYear")
                wNowMonth = DBTable1.Rows(0).Item("ReportMonth")
            Else
                wNowYear = NowDate.Year.ToString
                wNowMonth = 1
            End If
            OleDbConnection1.Close()
        End If
        DReportYear.Text = wNowYear
        '計算起迄年,週別
        DBDataSet1.Clear()
        wStartYear = wNowYear       '起始年
        wStartMonth = wNowMonth     '起始月
        wEndYear = wStartYear       '截止年
        wEndMonth = wStartMonth     '截止月
        '
        DUserID.Style("left") = -500 & "px"
        DUserID.Text = UCase(Request.QueryString("pUserID"))
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定機能連結
    '**
    '*****************************************************************
    Sub SetFunHyperLink()
        '隱藏機能連結
        DFunLink1.Visible = False
        DFunLink2.Visible = False
        DFunMark1.Visible = False
        '設定/開啟機能連結
        'IT
        DFunLink1.Text = "IT"
        DFunLink1.NavigateUrl = "MonthReport.aspx?pUserID=" & wUserID & _
                                                         "&pReportYear=" & wNowYear & _
                                                         "&pReportMonth=" & CStr(wNowMonth) & _
                                                         "&pKeyField=" & "HeadDivision" & _
                                                         "&pKeyData=" & Server.UrlEncode("IT") & _
                                                         "&pSearch=" & Server.UrlEncode(DSearch.Text)
        DFunLink1.Visible = True
        '部門別/業務員別/客戶別
        If wKeyField <> "HeadDivision" Then
            DFunLink2.Text = wKeyData
            DFunLink2.NavigateUrl = "MonthReport.aspx?pUserID=" & wUserID & _
                                                             "&pReportYear=" & wNowYear & _
                                                             "&pReportMonth=" & CStr(wNowMonth) & _
                                                             "&pKeyField=" & wKeyField & _
                                                             "&pKeyData=" & Server.UrlEncode(wKeyData) & _
                                                             "&pSearch=" & Server.UrlEncode(DSearch.Text)
            DFunMark1.Visible = True
            DFunLink2.Visible = True
        Else
            DFunMark1.Visible = False
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     顯示行程
    '**
    '*****************************************************************
    Sub ShowSchedule()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn1")  'SQL連結設定

        SQL = "SELECT ReportMonth+'/'+ReportDay As ReportDate, "
        SQL = SQL + "Case Week When 1 Then '日' When 2 Then '一' When 3 Then '二' When 4 Then '三' When 5 Then '四' When 6 Then '五' Else '六' End As WeekDesc, "
        SQL = SQL + "Division, Section, Name, wType, wContent,  wRemark, Vacation, "
        SQL = SQL + "'" + wUserID + "' As pUserID, "
        SQL = SQL + "'" + wNowYear + "' As pNowYear, "
        SQL = SQL + "'" + CStr(wNowMonth) + "' As pNowMonth, "
        SQL = SQL + "wID, "
        SQL = SQL + "Case HeadDivision When '' Then '' Else '....' End As Maint "
        SQL = SQL + "From V_WeekReport_01 "
        SQL = SQL + "Where ReportYear >= '" + wStartYear + "' "
        SQL = SQL + "  And ReportYear <= '" + wEndYear + "' "
        SQL = SQL + "  And ReportMonth >= '" + CStr(wStartMonth) + "' "
        SQL = SQL + "  And ReportMonth <= '" + CStr(wEndMonth) + "' "
        SQL = SQL + "  And ( "
        SQL = SQL + "        " + wKeyField + " = '" + wKeyData + "' "
        SQL = SQL + "    or  HeadDivision = '" + "" + "' "
        SQL = SQL + "      ) "
        If wSearch <> "" Then
            SQL = SQL + "  And ( replace(wContent,' ','') Like '%" + Replace(wSearch, " ", "") + "%' OR replace(wRemark,' ','') Like '%" & Replace(wSearch, " ", "") & "%' ) "
        End If
        SQL = SQL + "Order by ReportYear, Weeks, Week, Division, Section, Name, wType "

        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "WeekReport")
        DBTable1 = DBDataSet1.Tables("WeekReport")

        With GridView1
            .DataSource = DBTable1
            .DataBind()
        End With
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView--合併儲存格/設定背景顏色
    '**
    '*****************************************************************
    Protected Sub GridView1_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PreRender
        Dim i As Integer = 1
        Dim wRowSpan As Integer = 1
        Dim MergeIdx As Integer = 0
        Dim mySingleRow As GridViewRow
        Dim mergeColumns(6) As Integer
        mergeColumns(0) = 0
        mergeColumns(1) = 1
        mergeColumns(2) = 9
        mergeColumns(3) = 10
        mergeColumns(4) = 11
        mergeColumns(5) = 12

        '合併儲存格
        For MergeIdx = 0 To 5
            i = 1
            For Each mySingleRow In GridView1.Rows
                If CInt(mySingleRow.RowIndex) > 0 Then  '資料筆數>0

                    'MsgBox(CStr(mergeColumns(MergeIdx)))
                    'MsgBox(mySingleRow.Cells(mergeColumns(MergeIdx)).Text.Trim())
                    'MsgBox(CStr(mySingleRow.RowIndex))
                    'MsgBox(GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(MergeIdx)).Text.Trim())

                    If mySingleRow.Cells(mergeColumns(MergeIdx)).Text.Trim() = GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(MergeIdx)).Text.Trim() Then

                        If mySingleRow.Cells(mergeColumns(MergeIdx)).Text.Trim() <> "&nbsp;" Then  '空白是否
                            '
                            '日期/星期
                            If mergeColumns(MergeIdx) < 2 Then   '顯示欄位是否
                                '各欄位合併處理
                                GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(MergeIdx)).RowSpan += 1
                                mySingleRow.Cells(mergeColumns(MergeIdx)).Visible = False
                                i = i + 1
                            End If
                            '
                            'GROUP
                            If mergeColumns(MergeIdx) = 9 Then
                                If mySingleRow.Cells(9).Text.Trim() = GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(9).Text.Trim() Then
                                    If GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(2).RowSpan = 0 Then
                                        GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(2).RowSpan = 1
                                    End If
                                    GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(2).RowSpan += 1
                                    mySingleRow.Cells(2).Visible = False
                                    i = i + 1
                                End If
                            End If
                            '
                            'SECTION
                            If mergeColumns(MergeIdx) = 10 Then
                                If mySingleRow.Cells(10).Text.Trim() = GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(10).Text.Trim() Then
                                    If GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(3).RowSpan = 0 Then
                                        GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(3).RowSpan = 1
                                    End If
                                    GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(3).RowSpan += 1
                                    mySingleRow.Cells(3).Visible = False
                                    i = i + 1
                                End If
                            End If
                            '
                            'MEMBER
                            If mergeColumns(MergeIdx) = 11 Then
                                If mySingleRow.Cells(11).Text.Trim() = GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(11).Text.Trim() Then
                                    If GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(4).RowSpan = 0 Then
                                        GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(4).RowSpan = 1
                                    End If
                                    GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(4).RowSpan += 1
                                    mySingleRow.Cells(4).Visible = False
                                    i = i + 1
                                End If
                            End If
                            '
                            '作業
                            If mergeColumns(MergeIdx) = 12 Then
                                If mySingleRow.Cells(12).Text.Trim() = GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(12).Text.Trim() Then
                                    If GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(5).RowSpan = 0 Then
                                        GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(5).RowSpan = 1
                                    End If
                                    GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(5).RowSpan += 1
                                    mySingleRow.Cells(5).Visible = False
                                    i = i + 1
                                End If
                            End If


                        End If   '空白是否
                    Else
                        GridView1.Rows(CInt(mySingleRow.RowIndex)).Cells(mergeColumns(MergeIdx)).RowSpan = 1
                        i = 1
                    End If

                Else  '資料筆數>0
                    mySingleRow.Cells(mergeColumns(MergeIdx)).RowSpan = 1
                End If
            Next
        Next

        '設定背景顏色
        For Each mySingleRow In GridView1.Rows
            '休假日
            If GridView1.Rows(CInt(mySingleRow.RowIndex)).Cells(8).Text.Trim() = "1" Then
                GridView1.Rows(CInt(mySingleRow.RowIndex)).Cells(0).BackColor = Drawing.Color.LightPink
            End If
            '當日
            If GridView1.Rows(CInt(mySingleRow.RowIndex)).Cells(0).Text.Trim() = CStr(DateTime.Now.Month) + "/" + CStr(DateTime.Now.Day) Then
                GridView1.Rows(CInt(mySingleRow.RowIndex)).Cells(0).BackColor = Drawing.Color.Cyan
            End If
        Next
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView--欄位換行處理
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            Dim str As String
            '
            '內容
            str = Server.HtmlDecode(DataBinder.Eval(e.Row.DataItem, "wContent"))
            str = Replace(str, Chr(10), "<br />")
            e.Row.Cells(6).Text = str
            '
            '說明
            str = Server.HtmlDecode(DataBinder.Eval(e.Row.DataItem, "wRemark"))
            str = Replace(str, Chr(10), "<br />")
            e.Row.Cells(7).Text = str
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView--欄位隱藏處理
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        e.Row.Cells(8).Visible = False   '休假
        e.Row.Cells(9).Visible = False   'GROUP
        e.Row.Cells(10).Visible = False   'SECTION
        e.Row.Cells(11).Visible = False  '成員
        e.Row.Cells(12).Visible = False  '作業
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     上一月
    '**
    '*****************************************************************
    Protected Sub BUp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BUp.Click
        Dim URL As String

        If wNowMonth = 1 Then
            Dim DBDataSet1 As New DataSet
            Dim DBTable1 As New DataTable
            Dim SQL As String

            Dim OleDbConnection1 As New OleDbConnection
            Dim OleDBCommand1 As New OleDbCommand
            OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn1")  'SQL連結設定

            SQL = "Select ReportYear, ReportMonth From V_WeekReport_01 "
            SQL = SQL + "Where ReportYear < '" + wNowYear + "' "
            SQL = SQL + "Order by ReportDate Desc "
            OleDbConnection1.Open()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "WeekReport")
            DBTable1 = DBDataSet1.Tables("WeekReport")
            If DBTable1.Rows.Count > 0 Then
                wNowYear = DBTable1.Rows(0).Item("ReportYear")
                wNowMonth = DBTable1.Rows(0).Item("ReportMonth")
            End If
            OleDbConnection1.Close()
        Else
            wNowMonth = wNowMonth - 1
        End If

        URL = "MonthReport.aspx?pUserID=" & wUserID & _
                                       "&pReportYear=" & wNowYear & _
                                       "&pReportMonth=" & CStr(wNowMonth) & _
                                       "&pKeyField=" & wKeyField & _
                                       "&pKeyData=" & Server.UrlEncode(wKeyData) & _
                                       "&pSearch=" & DSearch.Text
        Response.Redirect(URL)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     下一月
    '**
    '*****************************************************************
    Protected Sub BDown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BDown.Click
        Dim URL As String

        If wNowMonth = 12 Then
            wNowYear = CStr(CInt(wNowYear) + 1)
            wNowMonth = 1

            'Dim DBDataSet1 As New DataSet
            'Dim DBTable1 As New DataTable
            'Dim SQL As String

            'Dim OleDbConnection1 As New OleDbConnection
            'Dim OleDBCommand1 As New OleDbCommand
            'OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn1")  'SQL連結設定

            'SQL = "Select ReportYear, ReportMonth From V_WeekReport_01 "
            'SQL = SQL + "Where ReportYear = '" + wNowYear + "' "
            'SQL = SQL + "  And ReportMonth > '" + CStr(wNowMonth) + "' "
            'SQL = SQL + "Order by ReportDate Desc "
            'OleDbConnection1.Open()
            'Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            'DBAdapter1.Fill(DBDataSet1, "WeekReport")
            'DBTable1 = DBDataSet1.Tables("WeekReport")
            'If DBTable1.Rows.Count > 0 Then
            '    wNowMonth = wNowMonth + 1
            'Else
            '    wNowYear = CStr(CInt(wNowYear) + 1)
            '    wNowMonth = 1
            'End If
            'OleDbConnection1.Close()
        Else
            wNowMonth = wNowMonth + 1
        End If

        URL = "MonthReport.aspx?pUserID=" & wUserID & _
                                       "&pReportYear=" & wNowYear & _
                                       "&pReportMonth=" & CStr(wNowMonth) & _
                                       "&pKeyField=" & wKeyField & _
                                       "&pKeyData=" & Server.UrlEncode(wKeyData) & _
                                       "&pSearch=" & DSearch.Text
        Response.Redirect(URL)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     當月
    '**
    '*****************************************************************
    Protected Sub BThis_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BThis.Click
        Dim URL As String

        URL = "MonthReport.aspx?pUserID=" & wUserID & _
                                       "&pReportYear=" & wNowYear & _
                                       "&pReportMonth=" & CStr(wNowMonth) & _
                                       "&pKeyField=" & wKeyField & _
                                       "&pKeyData=" & Server.UrlEncode(wKeyData) & _
                                       "&pSearch=" & DSearch.Text
        Response.Redirect(URL)
    End Sub

    Protected Sub BWeek_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BWeek.Click
        Dim Cmd As String
        Dim URL As String

        URL = "WeekReport.aspx?pUserID=" & wUserID & _
                                       "&pReportYear=" & wNowYear & _
                                       "&pWeeks=" & "" & _
                                       "&pKeyField=" & wKeyField & _
                                       "&pKeyData=" & Server.UrlEncode(wKeyData)
        Cmd = "<script>" + _
                "window.location.href='" & URL & "';" + _
              "</script>"
        Response.Write(Cmd)
        '    
        ShowSchedule()
    End Sub
End Class

