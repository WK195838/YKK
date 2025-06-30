Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class IRWNoticePage
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    '外部Object
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String
    Dim YYMM(6) As String

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "IRWNoticePage.aspx"

        If Not Me.IsPostBack Then
            SetParameter()          '設定共用參數
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
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        GridView2.Visible = False
        GridView3.Visible = False
        GridView4.Visible = False
        GridView5.Visible = False
        GridView6.Visible = False
        GridView7.Visible = False
        GridView8.Visible = False
        GridView9.Visible = False
        GridView10.Visible = False
        GridView11.Visible = False
        '
        Dim Sql As String
        Sql = "SELECT * FROM W_IRWSeqNo Order By Unique_id "
        Dim dt As DataTable = uDataBase.GetDataTable(Sql)
        For i As Integer = 0 To dt.Rows.Count - 1
            YYMM(i + 1) = dt.Rows(i).Item("YYMM")
        Next
        '
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim wTop As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBDataSet3 As New DataSet
        Dim DBDataSet4 As New DataSet
        Dim DBDataSet5 As New DataSet
        Dim DBDataSet6 As New DataSet
        Dim DBDataSet7 As New DataSet
        Dim DBDataSet8 As New DataSet
        Dim DBDataSet9 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '
        '[全體]未受注顧客服務率
        SQL = "SELECT "
        SQL = SQL & "'15 %' AS DataType, MM1_PER, MM2_PER, MM3_PER, MM4_PER, MM5_PER, MM6_PER "
        SQL = SQL & "FROM V_IRWNOActivityPeecent "
        SQL = SQL & "Where DataType = 'ALL' "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "ALL")
        GridView10.DataSource = DBDataSet1
        GridView10.DataBind()
        GridView10.Visible = True
        OleDbConnection1.Close()
        '
        '[部門]未受注顧客服務率
        SQL = "SELECT "
        SQL = SQL & "ROW_NUMBER() OVER (ORDER BY MM3_PER desc) as Rank, "
        'SQL = SQL & "DataType, MM1_PER, MM2_PER, MM3_PER, MM4_PER, MM5_PER, MM6_PER "
        '
        SQL = SQL & "DataType, "
        SQL = SQL & "case when MM1_PER<0 then 0 else MM1_PER end as MM1_PER, "
        SQL = SQL & "case when MM2_PER<0 then 0 else MM2_PER end as MM2_PER, "
        SQL = SQL & "case when MM3_PER<0 then 0 else MM3_PER end as MM3_PER, "
        SQL = SQL & "case when MM4_PER<0 then 0 else MM4_PER end as MM4_PER, "
        SQL = SQL & "case when MM5_PER<0 then 0 else MM5_PER end as MM5_PER, "
        SQL = SQL & "case when MM6_PER<0 then 0 else MM6_PER end as MM6_PER  "
        '
        SQL = SQL & "FROM V_IRWNOActivityPeecent "
        SQL = SQL & "Where DataType <> 'ALL' "
        SQL = SQL & "Order By MM3_PER DESC "
        '
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "DEP")
        GridView11.DataSource = DBDataSet2
        GridView11.DataBind()
        GridView11.Visible = True
        OleDbConnection1.Close()
        '
        '--------------------------------
        '無效工時(RANKING)
        wTop = 390
        GridView9.Style("top") = wTop & "px"
        '
        SQL = "SELECT TOP 10  "
        SQL = SQL & "ROW_NUMBER() OVER (ORDER BY  SUM( CAST( REPLACE( REPLACE(NoworkHour1, ' H','') ,' ','' ) AS decimal(18, 2) ) ) desc) as Rank, "
        SQL = SQL & "DEPNAME + '-' + APPNAME as DepName, "
        '
        SQL = SQL & "SUM( CAST( REPLACE( REPLACE(NoworkHour1, ' H','') ,' ','' ) AS decimal(18, 2) ) ) as Hour, "
        SQL = SQL & "(select SUM( CAST( REPLACE( REPLACE(NoworkHour1, ' H','') ,' ','' ) AS decimal(18, 2) ) ) from F_NoOrderReortSheet) as HourTotal, "
        SQL = SQL & "SUM( CAST( REPLACE( REPLACE(NoworkHour1, ' H','') ,' ','' ) AS decimal(18, 2) ) ) "
        SQL = SQL & "/"
        SQL = SQL & "(select SUM( CAST( REPLACE( REPLACE(NoworkHour1, ' H','') ,' ','' ) AS decimal(18, 2) ) ) from F_NoOrderReortSheet) "
        SQL = SQL & "*"
        SQL = SQL & "100"
        SQL = SQL & "as HourPer, "
        '
        SQL = SQL & "'http://10.245.1.6/IRW/IRWNoticePageRanking.aspx?pEmpID=' + AppID as V2URL, 'Link' as V2 "
        '
        SQL = SQL & "FROM F_NoOrderReortSheet "
        SQL = SQL & "GROUP BY DEPNAME, APPNAME, AppID "
        SQL = SQL & "ORDER BY SUM( CAST( REPLACE( REPLACE(NoworkHour1, ' H','') ,' ','' ) AS decimal(18, 2) ) ) desc "
        '
        Dim DBAdapter9 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter9.Fill(DBDataSet9, "SUM")
        '
        If DBDataSet9.Tables("SUM").Rows.Count > 0 Then
            GridView9.DataSource = DBDataSet9
            GridView9.DataBind()
            GridView9.Visible = True
            '
            'ADJUST
            GridView8.Style("top") = (DBDataSet9.Tables("SUM").Rows.Count + 4) * 23 + wTop & "px"
            wTop = (DBDataSet9.Tables("SUM").Rows.Count + 4) * 40 + wTop
        Else
            GridView8.Style("top") = 1 * 20 + wTop & "px"
            wTop = (DBDataSet9.Tables("SUM").Rows.Count + 4) * 40 + wTop
        End If
        '
        OleDbConnection1.Close()        '
        '--------------------------------
        '監查報告
        '
        SQL = "SELECT "
        SQL = SQL & "[NO], case when sts=0 then '報告中' when sts=1 then '報告完成' else '報告取消(離)' end as Status, "
        SQL = SQL & "[DepName] + '-' + [AppName] as AppName, "
        SQL = SQL & "SUBSTRING(convert(varchar, DATEADD(MONTH, -4,  [LeadDate]), 112),1,4) + '/' + SUBSTRING(convert(varchar, DATEADD(MONTH, -4,  [LeadDate]), 112),5,2) as YYMM, "
        SQL = SQL & "[ApplyQty] as ALLCount, "
        SQL = SQL & "[NoOrderQty] as NG, "
        SQL = SQL & "[Percen] as NGPer, "
        SQL = SQL & "[NoworkHour1] as NoWorkHour, "
        '
        SQL = SQL & "'http://10.245.1.6/IRW/IRWNoOrderReportSheet_02.aspx?pUserid=&pFormNo=' + [FormNo] + '&pFormSno=' + ltrim(rtrim(str(FormSno))) as NoURL "
        '
        SQL = SQL & "FROM F_NoOrderReortSheet "
        SQL = SQL & "where sts in (0) "
        SQL = SQL & "   or (sts in (1,3) and SUBSTRING(convert(varchar, [LeadDate], 112),1,6) = SUBSTRING(convert(varchar, GETDATE(), 112),1,6) ) "
        SQL = SQL & "Order By sts, LeadDate, [DepName] + '-' + [AppName] "
        '
        Dim DBAdapter8 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter8.Fill(DBDataSet8, "WFS")
        '
        If DBDataSet8.Tables("WFS").Rows.Count > 0 Then
            GridView8.DataSource = DBDataSet8
            GridView8.DataBind()
            GridView8.Visible = True
            '
            'ADJUST
            GridView4.Style("top") = (DBDataSet8.Tables("WFS").Rows.Count + 4) * 10 + wTop & "px"
            wTop = (DBDataSet8.Tables("WFS").Rows.Count + 4) * 10 + wTop
        Else
            GridView4.Style("top") = 1 * 40 + wTop & "px"
            wTop = (DBDataSet8.Tables("WFS").Rows.Count + 4) * 40 + wTop
        End If
        '
        OleDbConnection1.Close()
        '
        '報告提出者
        SQL = "SELECT "
        SQL = SQL & "DepName,A_YES,A_NO,A_TOTAL,A_PERCENT,S_PERCENT,S_NO,H_NO,S_NO+H_NO AS L_NO,B_NO "
        SQL = SQL & "FROM W_IRWNOActivityV2 "
        SQL = SQL & "where Warning='*' "
        SQL = SQL & "and YM = SUBSTRING(convert(varchar, DATEADD(MONTH, -4,  GETDATE()), 112),1,6) "
        SQL = SQL & "Order By B_NO DESC "
        '
        Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter4.Fill(DBDataSet4, "NO")
        GridView4.DataSource = DBDataSet4
        GridView4.DataBind()
        GridView4.Visible = True
        '
        'ADJUST
        If DBDataSet4.Tables("NO").Rows.Count > 0 Then
            'ADJUST
            GridView5.Style("top") = (DBDataSet4.Tables("NO").Rows.Count + 4) * 40 + wTop - 380 & "px"
            wTop = (DBDataSet4.Tables("NO").Rows.Count + 4) * 40 + wTop
        End If
        '
        OleDbConnection1.Close()
        '
        '次月需注意部門
        SQL = "SELECT "
        SQL = SQL & "[DepName],[EmpName],[A_YES],[A_NO],[A_TOTAL],[A_PERCENT],[Remark], "
        '
        SQL = SQL & "'IRWNoticePage2310.aspx?pDepCode=' + DepCode + '&pYM=' + [YM] as V2URL, 'Link' as V2 "
        '
        SQL = SQL & "FROM W_IRWNOActivityV2Person  "
        SQL = SQL & "where Warning='*' "
        SQL = SQL & "and YM = SUBSTRING(convert(varchar, DATEADD(MONTH, -4,  GETDATE()), 112),1,6) "
        SQL = SQL & "Order By W1 Desc, [DepName], [EmpName] "
        '
        Dim DBAdapter5 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter5.Fill(DBDataSet5, "CHECK")
        GridView5.DataSource = DBDataSet5
        GridView5.DataBind()
        GridView5.Visible = True
        '
        If DBDataSet5.Tables("CHECK").Rows.Count > 0 Then
            GridView6.Style("top") = (DBDataSet5.Tables("CHECK").Rows.Count + 4) * 40 + wTop & "px"
            wTop = (DBDataSet5.Tables("CHECK").Rows.Count + 4) * 40 + wTop
        End If
        '
        OleDbConnection1.Close()
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-4行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "TOP"
            H4row.Cells.Add(H4tc0)

            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "NAME"
            H4row.Cells.Add(H4tc1)

            Dim H4tc1A As TableCell = New TableCell
            H4tc1A.Text = "COUNT"
            H4row.Cells.Add(H4tc1A)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------

            ' 表頭-3行
            '
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "年累計  (2022/8~現在)"
            H3tc1.ColumnSpan = 3
            H3tc1.BackColor = Color.Red
            H3row.Cells.Add(H3tc1)

            gv.Controls(0).Controls.AddAt(0, H3row)

        End If
    End Sub
    '

    Protected Sub GridView2_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-4行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "TOP"
            H4row.Cells.Add(H4tc0)

            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "NAME"
            H4row.Cells.Add(H4tc1)

            Dim H4tc1A As TableCell = New TableCell
            H4tc1A.Text = "COUNT"
            H4row.Cells.Add(H4tc1A)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            ' 表頭-3行
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "月累計 " & "(" & Mid(NowDate, 1, 4) & "/" & Mid(NowDate, 5, 2) & ")"
            H3tc1.ColumnSpan = 3
            H3tc1.BackColor = Color.Red
            H3row.Cells.Add(H3tc1)
            gv.Controls(0).Controls.AddAt(0, H3row)
        End If
    End Sub

    Protected Sub GridView3_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView3.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim SQL, xDataTime As String
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-4行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "TOP"
            H4row.Cells.Add(H4tc0)

            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "NAME"
            H4row.Cells.Add(H4tc1)

            Dim H4tc1A As TableCell = New TableCell
            H4tc1A.Text = "COUNT"
            H4row.Cells.Add(H4tc1A)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            ' 表頭-3行

            xDataTime = ""
            SQL = "SELECT TOP 1 Convert(varchar, CreateTime, 111) + ' ' + SUBSTRING(Convert(varchar, CreateTime, 108),1,2) As DataTime "
            SQL = SQL & "FROM W_IRWNOTICE "
            SQL = SQL & "Order By COUNT DESC "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "NO")
            If DBDataSet1.Tables("NO").Rows.Count > 0 Then
                xDataTime = DBDataSet1.Tables("NO").Rows(0).Item("DataTime")
            End If
            '
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "日累計　　" & Mid(xDataTime, 6) & "點更新"
            H3tc1.ColumnSpan = 3
            H3tc1.BackColor = Color.Red
            H3row.Cells.Add(H3tc1)
            gv.Controls(0).Controls.AddAt(0, H3row)
        End If
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer

        If (e.Row.RowType = DataControlRowType.DataRow) Then
            For i = 0 To 2
                Select Case i
                    Case 2
                        e.Row.Cells(i).Text = Format(CDbl(e.Row.Cells(i).Text), "###,###,###")
                    Case Else
                End Select
            Next
        End If

    End Sub
    '--------------------------------
    Protected Sub GridView4_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView4.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim SQL, xDataTime, xYM As String
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
            '
            Dim gv As GridView = sender
            Dim H2row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-N行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "部門" & "<BR>" & "　"
            H4row.Cells.Add(H4tc0)

            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "YES" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc1)

            Dim H4tc1A As TableCell = New TableCell
            H4tc1A.Text = "NG" & "<BR>" & "件數"
            H4tc1A.BackColor = Color.Blue
            H4row.Cells.Add(H4tc1A)

            Dim H4tc1B As TableCell = New TableCell
            H4tc1B.Text = "TOTAL" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc1B)

            Dim H4tc1C As TableCell = New TableCell
            H4tc1C.Text = "NG" & "<BR>" & "比例"
            H4row.Cells.Add(H4tc1C)

            Dim H4tc1D As TableCell = New TableCell
            H4tc1D.Text = "NG" & "<BR>" & "比例"
            H4row.Cells.Add(H4tc1D)

            Dim H4tc1E As TableCell = New TableCell
            H4tc1E.Text = "NG" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc1E)

            Dim H4tc1F As TableCell = New TableCell
            H4tc1F.Text = "支援" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc1F)

            Dim H4tc1G As TableCell = New TableCell
            H4tc1G.Text = "調整後" & "<BR>" & "件數"
            H4tc1G.BackColor = Color.Blue
            H4row.Cells.Add(H4tc1G)

            Dim H4tc1H As TableCell = New TableCell
            H4tc1H.Text = "超過" & "<BR>" & "件數"
            H4tc1H.BackColor = Color.Blue
            H4row.Cells.Add(H4tc1H)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            ' 表頭-N-1行
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "實績"
            H3tc1.ColumnSpan = 5
            H3tc1.BackColor = Color.Green
            H3row.Cells.Add(H3tc1)

            Dim H3tc2 As TableCell = New TableCell
            H3tc2.Text = "目標(a)"
            H3tc2.ColumnSpan = 2
            H3tc2.BackColor = Color.Green
            H3row.Cells.Add(H3tc2)

            Dim H3tc3 As TableCell = New TableCell
            H3tc3.Text = "他部門(b)"
            H3tc3.ColumnSpan = 1
            H3tc3.BackColor = Color.Green
            H3row.Cells.Add(H3tc3)

            Dim H3tc4 As TableCell = New TableCell
            H3tc4.Text = "(a+b)"
            H3tc4.ColumnSpan = 1
            H3tc4.BackColor = Color.Green
            H3row.Cells.Add(H3tc4)

            Dim H3tc5 As TableCell = New TableCell
            H3tc5.Text = "結果"
            H3tc5.ColumnSpan = 1
            H3tc5.BackColor = Color.Green
            H3row.Cells.Add(H3tc5)

            gv.Controls(0).Controls.AddAt(0, H3row)

            '-----------------------------------------
            ' 表頭-N-2行
            '
            xDataTime = ""
            xYM = ""
            SQL = "SELECT TOP 1 Convert(varchar, CreateTime, 111) + ' ' + SUBSTRING(Convert(varchar, CreateTime, 108),1,2) As DataTime, YM "
            SQL = SQL & "FROM W_IRWNOActivityV2 "
            SQL = SQL & "where Warning='*' "
            SQL = SQL & "and YM = SUBSTRING(convert(varchar, DATEADD(MONTH, -4,  GETDATE()), 112),1,6) "
            SQL = SQL & "Order By B_NO DESC "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "NO")
            If DBDataSet1.Tables("NO").Rows.Count > 0 Then
                xDataTime = DBDataSet1.Tables("NO").Rows(0).Item("DataTime")
                xYM = Mid(DBDataSet1.Tables("NO").Rows(0).Item("YM"), 1, 4) & "/" & Mid(DBDataSet1.Tables("NO").Rows(0).Item("YM"), 5, 2)
            End If
            '
            Dim H2tc1 As TableCell = New TableCell
            H2tc1.Text = "違反部門"
            H2tc1.ColumnSpan = 3
            H2tc1.BackColor = Color.Red
            H2tc1.HorizontalAlign = HorizontalAlign.Left
            H2row.Cells.Add(H2tc1)

            Dim H2tc11 As TableCell = New TableCell
            H2tc11.Text = "對象年月=" & xYM & "申請"
            H2tc11.ColumnSpan = 3
            H2tc11.BackColor = Color.White
            H2tc11.ForeColor = Color.Black
            H2tc11.HorizontalAlign = HorizontalAlign.Left
            H2row.Cells.Add(H2tc11)

            Dim H2tc2 As TableCell = New TableCell
            H2tc2.Text = Mid(xDataTime, 6) & "點更新"
            H2tc2.ColumnSpan = 4
            H2tc2.HorizontalAlign = HorizontalAlign.Right
            H2tc2.BackColor = Color.White
            H2tc2.ForeColor = Color.Black
            H2row.Cells.Add(H2tc2)

            gv.Controls(0).Controls.AddAt(0, H2row)
        End If
    End Sub

    Protected Sub GridView4_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView4.RowDataBound
        Dim i As Integer
        '
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            For i = 0 To 9
                Select Case i
                    Case 0
                        e.Row.Cells(i).BackColor = Color.Khaki
                    Case 2, 8, 9
                        e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                        e.Row.Cells(i).BackColor = Color.Pink
                    Case 4
                        If CDbl(e.Row.Cells(i).Text) < 0 Then
                            e.Row.Cells(i).Text = "0" & "%"
                        Else
                            e.Row.Cells(i).Text = e.Row.Cells(i).Text & "%"
                        End If
                    Case 5
                        e.Row.Cells(i).Text = "15" & "%"
                    Case Else
                End Select
            Next
        End If
    End Sub

    Protected Sub GridView5_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView5.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            '
            Dim SQL, xDataTime, xYM As String
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

            Dim gv As GridView = sender
            Dim H2row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H5row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-N行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "部門" & "<BR>" & "　"
            H4row.Cells.Add(H4tc0)

            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "NAME" & "<BR>" & "　"
            H4row.Cells.Add(H4tc1)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = "YES" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc2)

            Dim H4tc1A As TableCell = New TableCell
            H4tc1A.Text = "NG" & "<BR>" & "件數"
            H4tc1A.BackColor = Color.Blue
            H4row.Cells.Add(H4tc1A)

            Dim H4tc1B As TableCell = New TableCell
            H4tc1B.Text = "TOTAL" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc1B)

            Dim H4tc1C As TableCell = New TableCell
            H4tc1C.Text = "NG" & "<BR>" & "比例"
            H4row.Cells.Add(H4tc1C)

            Dim H4tc1D As TableCell = New TableCell
            H4tc1D.Text = "REMARK" & "<BR>" & "　"
            H4row.Cells.Add(H4tc1D)

            Dim H4tc1E As TableCell = New TableCell
            H4tc1E.Text = ""
            H4row.Cells.Add(H4tc1E)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            ' 表頭-N-1行
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "實績"
            H3tc1.ColumnSpan = 6
            H3tc1.BackColor = Color.Green
            H3row.Cells.Add(H3tc1)

            Dim H3tc2 As TableCell = New TableCell
            H3tc2.Text = "結果"
            H3tc2.ColumnSpan = 2
            H3tc2.BackColor = Color.Green
            H3row.Cells.Add(H3tc2)

            gv.Controls(0).Controls.AddAt(0, H3row)
            '-----------------------------------------
            ' 表頭-N-2行
            xDataTime = ""
            xYM = ""
            SQL = "SELECT TOP 1 Convert(varchar, CreateTime, 111) + ' ' + SUBSTRING(Convert(varchar, CreateTime, 108),1,2) As DataTime, YM "
            SQL = SQL & "FROM W_IRWNOActivityV2  "
            SQL = SQL & "where Warning='*' "
            SQL = SQL & "and YM = SUBSTRING(convert(varchar, DATEADD(MONTH, -4,  GETDATE()), 112),1,6) "
            SQL = SQL & "Order By B_NO DESC "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "NO")
            If DBDataSet1.Tables("NO").Rows.Count > 0 Then
                xDataTime = DBDataSet1.Tables("NO").Rows(0).Item("DataTime")
                xYM = Mid(DBDataSet1.Tables("NO").Rows(0).Item("YM"), 1, 4) & "/" & Mid(DBDataSet1.Tables("NO").Rows(0).Item("YM"), 5, 2)
            End If
            '
            Dim H2tc1 As TableCell = New TableCell
            H2tc1.Text = "報告提出者"
            H2tc1.ColumnSpan = 3
            H2tc1.BackColor = Color.Red
            H2tc1.HorizontalAlign = HorizontalAlign.Left
            H2row.Cells.Add(H2tc1)

            Dim H2tc11 As TableCell = New TableCell
            H2tc11.Text = "對象年月=" & xYM & "申請"
            H2tc11.ColumnSpan = 3
            H2tc11.BackColor = Color.White
            H2tc11.ForeColor = Color.Black
            H2tc11.HorizontalAlign = HorizontalAlign.Left
            H2row.Cells.Add(H2tc11)

            Dim H2tc2 As TableCell = New TableCell
            H2tc2.Text = Mid(xDataTime, 6) & "點更新"
            H2tc2.ColumnSpan = 2
            H2tc2.HorizontalAlign = HorizontalAlign.Right
            H2tc2.BackColor = Color.White
            H2tc2.ForeColor = Color.Black
            H2row.Cells.Add(H2tc2)

            gv.Controls(0).Controls.AddAt(0, H2row)

        End If
    End Sub

    Protected Sub GridView5_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView5.RowDataBound
        Dim i As Integer
        '
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            For i = 0 To 6
                Select Case i
                    Case 3
                        e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                        e.Row.Cells(i).BackColor = Color.Pink
                    Case 6
                        e.Row.Cells(i).ForeColor = Color.Red

                    Case Else
                End Select
            Next
        End If

    End Sub


    Protected Sub GridView6_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView6.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim SQL, xDataTime, xYM As String
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
            '
            Dim gv As GridView = sender
            Dim H2row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-N行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "部門" & "<BR>" & "　"
            H4row.Cells.Add(H4tc0)

            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "YES" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc1)

            Dim H4tc1A As TableCell = New TableCell
            H4tc1A.Text = "NG" & "<BR>" & "件數"
            H4tc1A.BackColor = Color.Blue
            H4row.Cells.Add(H4tc1A)

            Dim H4tc1B As TableCell = New TableCell
            H4tc1B.Text = "TOTAL" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc1B)

            Dim H4tc1C As TableCell = New TableCell
            H4tc1C.Text = "NG" & "<BR>" & "比例"
            H4row.Cells.Add(H4tc1C)

            Dim H4tc1D As TableCell = New TableCell
            H4tc1D.Text = "NG" & "<BR>" & "比例"
            H4row.Cells.Add(H4tc1D)

            Dim H4tc1E As TableCell = New TableCell
            H4tc1E.Text = "NG" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc1E)

            Dim H4tc1F As TableCell = New TableCell
            H4tc1F.Text = "支援" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc1F)

            Dim H4tc1G As TableCell = New TableCell
            H4tc1G.Text = "調整後" & "<BR>" & "件數"
            H4tc1G.BackColor = Color.Blue
            H4row.Cells.Add(H4tc1G)

            Dim H4tc1H As TableCell = New TableCell
            H4tc1H.Text = "超過" & "<BR>" & "件數"
            H4tc1H.BackColor = Color.Blue
            H4row.Cells.Add(H4tc1H)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            ' 表頭-N-1行
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "實績"
            H3tc1.ColumnSpan = 5
            H3tc1.BackColor = Color.Green
            H3row.Cells.Add(H3tc1)

            Dim H3tc2 As TableCell = New TableCell
            H3tc2.Text = "目標(a)"
            H3tc2.ColumnSpan = 2
            H3tc2.BackColor = Color.Green
            H3row.Cells.Add(H3tc2)

            Dim H3tc3 As TableCell = New TableCell
            H3tc3.Text = "他部門(b)"
            H3tc3.ColumnSpan = 1
            H3tc3.BackColor = Color.Green
            H3row.Cells.Add(H3tc3)

            Dim H3tc4 As TableCell = New TableCell
            H3tc4.Text = "(a+b)"
            H3tc4.ColumnSpan = 1
            H3tc4.BackColor = Color.Green
            H3row.Cells.Add(H3tc4)

            Dim H3tc5 As TableCell = New TableCell
            H3tc5.Text = "結果"
            H3tc5.ColumnSpan = 1
            H3tc5.BackColor = Color.Green
            H3row.Cells.Add(H3tc5)

            gv.Controls(0).Controls.AddAt(0, H3row)

            '-----------------------------------------
            ' 表頭-N-2行
            '
            xDataTime = ""
            xYM = ""
            SQL = "SELECT TOP 1 Convert(varchar, CreateTime, 111) + ' ' + SUBSTRING(Convert(varchar, CreateTime, 108),1,2) As DataTime, YM "
            SQL = SQL & "FROM W_IRWNOActivityV2  "
            SQL = SQL & "where Warning='*' "
            SQL = SQL & "and YM = SUBSTRING(convert(varchar, DATEADD(MONTH, -3,  GETDATE()), 112),1,6) "
            SQL = SQL & "Order By B_NO DESC "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "NO")
            If DBDataSet1.Tables("NO").Rows.Count > 0 Then
                xDataTime = DBDataSet1.Tables("NO").Rows(0).Item("DataTime")
                xYM = Mid(DBDataSet1.Tables("NO").Rows(0).Item("YM"), 1, 4) & "/" & Mid(DBDataSet1.Tables("NO").Rows(0).Item("YM"), 5, 2)
            End If
            '
            Dim H2tc1 As TableCell = New TableCell
            H2tc1.Text = "[注意] 次月需注意部門"
            H2tc1.ColumnSpan = 3
            H2tc1.BackColor = Color.Navy
            H2tc1.HorizontalAlign = HorizontalAlign.Left
            H2row.Cells.Add(H2tc1)

            Dim H2tc11 As TableCell = New TableCell
            H2tc11.Text = "對象年月=" & xYM & "申請"
            H2tc11.ColumnSpan = 3
            H2tc11.BackColor = Color.White
            H2tc11.ForeColor = Color.Black
            H2tc11.HorizontalAlign = HorizontalAlign.Left
            H2row.Cells.Add(H2tc11)

            Dim H2tc2 As TableCell = New TableCell
            H2tc2.Text = Mid(xDataTime, 6) & "點更新"
            H2tc2.ColumnSpan = 4
            H2tc2.HorizontalAlign = HorizontalAlign.Right
            H2tc2.BackColor = Color.White
            H2tc2.ForeColor = Color.Black
            H2row.Cells.Add(H2tc2)

            gv.Controls(0).Controls.AddAt(0, H2row)
        End If
    End Sub

    Protected Sub GridView6_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView6.RowDataBound
        Dim i As Integer
        '
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            For i = 0 To 9
                Select Case i
                    Case 0
                        e.Row.Cells(i).BackColor = Color.Khaki
                    Case 2, 8, 9
                        e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                        e.Row.Cells(i).BackColor = Color.Pink
                    Case 4
                        If CDbl(e.Row.Cells(i).Text) < 0 Then
                            e.Row.Cells(i).Text = "0" & "%"
                        Else
                            e.Row.Cells(i).Text = e.Row.Cells(i).Text & "%"
                        End If
                    Case 5
                        e.Row.Cells(i).Text = "15" & "%"
                    Case Else
                End Select
            Next
        End If
    End Sub

    Protected Sub GridView7_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView7.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            '
            Dim SQL, xDataTime, xYM As String
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

            Dim gv As GridView = sender
            Dim H2row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H5row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-N行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "部門" & "<BR>" & "　"
            H4row.Cells.Add(H4tc0)

            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "NAME" & "<BR>" & "　"
            H4row.Cells.Add(H4tc1)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = "YES" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc2)

            Dim H4tc1A As TableCell = New TableCell
            H4tc1A.Text = "NG" & "<BR>" & "件數"
            H4tc1A.BackColor = Color.Blue
            H4row.Cells.Add(H4tc1A)

            Dim H4tc1B As TableCell = New TableCell
            H4tc1B.Text = "TOTAL" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc1B)

            Dim H4tc1C As TableCell = New TableCell
            H4tc1C.Text = "NG" & "<BR>" & "比例"
            H4row.Cells.Add(H4tc1C)

            Dim H4tc1D As TableCell = New TableCell
            H4tc1D.Text = "REMARK" & "<BR>" & "　"
            H4row.Cells.Add(H4tc1D)

            Dim H4tc1E As TableCell = New TableCell
            H4tc1E.Text = ""
            H4row.Cells.Add(H4tc1E)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            ' 表頭-N-1行
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "實績"
            H3tc1.ColumnSpan = 6
            H3tc1.BackColor = Color.Green
            H3row.Cells.Add(H3tc1)

            Dim H3tc2 As TableCell = New TableCell
            H3tc2.Text = "結果"
            H3tc2.ColumnSpan = 2
            H3tc2.BackColor = Color.Green
            H3row.Cells.Add(H3tc2)

            gv.Controls(0).Controls.AddAt(0, H3row)
            '-----------------------------------------
            ' 表頭-N-2行
            xDataTime = ""
            xYM = ""
            SQL = "SELECT TOP 1 Convert(varchar, CreateTime, 111) + ' ' + SUBSTRING(Convert(varchar, CreateTime, 108),1,2) As DataTime, YM "
            SQL = SQL & "FROM W_IRWNOActivityV2  "
            SQL = SQL & "where Warning='*' "
            SQL = SQL & "and YM = SUBSTRING(convert(varchar, DATEADD(MONTH, -3,  GETDATE()), 112),1,6) "
            SQL = SQL & "Order By B_NO DESC "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "NO")
            If DBDataSet1.Tables("NO").Rows.Count > 0 Then
                xDataTime = DBDataSet1.Tables("NO").Rows(0).Item("DataTime")
                xYM = Mid(DBDataSet1.Tables("NO").Rows(0).Item("YM"), 1, 4) & "/" & Mid(DBDataSet1.Tables("NO").Rows(0).Item("YM"), 5, 2)
            End If
            '
            Dim H2tc1 As TableCell = New TableCell
            H2tc1.Text = "[注意] 次月需注意個人"
            H2tc1.ColumnSpan = 3
            H2tc1.BackColor = Color.Navy
            H2tc1.HorizontalAlign = HorizontalAlign.Left
            H2row.Cells.Add(H2tc1)

            Dim H2tc11 As TableCell = New TableCell
            H2tc11.Text = "對象年月=" & xYM & "申請"
            H2tc11.ColumnSpan = 3
            H2tc11.BackColor = Color.White
            H2tc11.ForeColor = Color.Black
            H2tc11.HorizontalAlign = HorizontalAlign.Left
            H2row.Cells.Add(H2tc11)

            Dim H2tc2 As TableCell = New TableCell
            H2tc2.Text = Mid(xDataTime, 6) & "點更新"
            H2tc2.ColumnSpan = 2
            H2tc2.HorizontalAlign = HorizontalAlign.Right
            H2tc2.BackColor = Color.White
            H2tc2.ForeColor = Color.Black
            H2row.Cells.Add(H2tc2)

            gv.Controls(0).Controls.AddAt(0, H2row)

        End If
    End Sub

    Protected Sub GridView7_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView7.RowDataBound
        Dim i As Integer
        '
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            For i = 0 To 6
                Select Case i
                    Case 3
                        e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                        e.Row.Cells(i).BackColor = Color.Pink
                    Case 6
                        e.Row.Cells(i).ForeColor = Color.Red

                    Case Else
                End Select
            Next
        End If
    End Sub


    Protected Sub GridView8_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView8.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            '
            Dim SQL, xDataTime, xYM As String
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

            Dim gv As GridView = sender
            Dim H2row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H5row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-N行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "No." & "<BR>" & ""
            H4row.Cells.Add(H4tc0)

            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "Status" & "<BR>" & ""
            H4row.Cells.Add(H4tc1)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = "Name" & "<BR>" & ""
            H4row.Cells.Add(H4tc2)

            Dim H4tc1A As TableCell = New TableCell
            H4tc1A.Text = "對象年月" & "<BR>" & ""
            H4row.Cells.Add(H4tc1A)

            Dim H4tc1B As TableCell = New TableCell
            H4tc1B.Text = "申請" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc1B)

            Dim H4tc1C As TableCell = New TableCell
            H4tc1C.Text = "未受注" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc1C)

            Dim H4tc1D As TableCell = New TableCell
            H4tc1D.Text = "未受注" & "<BR>" & "%"
            H4row.Cells.Add(H4tc1D)

            Dim H4tc1E As TableCell = New TableCell
            H4tc1E.Text = "無效" & "<BR>" & "工時"
            H4row.Cells.Add(H4tc1E)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            ' 表頭-N-2行
            xDataTime = ""
            xYM = ""
            SQL = "SELECT "
            SQL = SQL & "TOP 1 Convert(varchar,  [LeadDate], 111) + ' ' + SUBSTRING(Convert(varchar,  [LeadDate], 108),1,2) As DataTime, "
            SQL = SQL & "SUBSTRING(convert(varchar, DATEADD(MONTH, -4,  [LeadDate]), 112),1,6) as YM "
            SQL = SQL & "FROM F_NoOrderReortSheet "
            SQL = SQL & "where SUBSTRING(convert(varchar, [LeadDate], 112),1,6) = SUBSTRING(convert(varchar, GETDATE(), 112),1,6) "
            'SQL = SQL & "where SUBSTRING(convert(varchar, [LeadDate], 112),1,6) = '202401' "
            SQL = SQL & "Order By [DepName] + '-' + [AppName] "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "NO")
            If DBDataSet1.Tables("NO").Rows.Count > 0 Then
                xDataTime = DBDataSet1.Tables("NO").Rows(0).Item("DataTime")
                xYM = Mid(DBDataSet1.Tables("NO").Rows(0).Item("YM"), 1, 4) & "/" & Mid(DBDataSet1.Tables("NO").Rows(0).Item("YM"), 5, 2)
            End If
            '
            Dim H2tc1 As TableCell = New TableCell
            H2tc1.Text = "監查報告"
            H2tc1.ColumnSpan = 3
            H2tc1.BackColor = Color.Red
            H2tc1.HorizontalAlign = HorizontalAlign.Left
            H2row.Cells.Add(H2tc1)

            Dim H2tc11 As TableCell = New TableCell
            H2tc11.Text = "對象年月=2023/10 ~ " & xYM & "申請"
            H2tc11.ColumnSpan = 3
            H2tc11.BackColor = Color.White
            H2tc11.ForeColor = Color.Black
            H2tc11.HorizontalAlign = HorizontalAlign.Left
            H2row.Cells.Add(H2tc11)

            Dim H2tc2 As TableCell = New TableCell
            H2tc2.Text = "截止日:" & Mid(xDataTime, 6, 5)
            H2tc2.ColumnSpan = 2
            H2tc2.HorizontalAlign = HorizontalAlign.Right
            H2tc2.BackColor = Color.White
            H2tc2.ForeColor = Color.Black
            H2row.Cells.Add(H2tc2)

            gv.Controls(0).Controls.AddAt(0, H2row)

        End If
    End Sub

    Protected Sub GridView8_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView8.RowDataBound
        Dim i As Integer
        '
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            For i = 0 To 7
                Select Case i
                    Case 2
                        e.Row.Cells(i).BackColor = Color.Khaki
                    Case 5, 6, 7
                        e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                        e.Row.Cells(i).BackColor = Color.Pink
                    Case Else
                End Select
            Next
        End If

    End Sub

    Protected Sub GridView9_RowCreated1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView9.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            '
            Dim SQL, xDataTime, xYM As String
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

            Dim gv As GridView = sender
            Dim H2row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H5row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-N行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "TOP." & "<BR>" & ""
            H4row.Cells.Add(H4tc0)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = "Name" & "<BR>" & ""
            H4row.Cells.Add(H4tc2)

            Dim H4tc1A As TableCell = New TableCell
            H4tc1A.Text = "無效工時" & "<BR>" & "個人(a)"
            H4row.Cells.Add(H4tc1A)

            Dim H4tc1b As TableCell = New TableCell
            H4tc1b.Text = "無效工時" & "<BR>" & "總(b)"
            H4row.Cells.Add(H4tc1b)

            Dim H4tc1c As TableCell = New TableCell
            H4tc1c.Text = "比例" & "<BR>" & "a / b"
            H4row.Cells.Add(H4tc1c)

            Dim H4tc1d As TableCell = New TableCell
            H4tc1d.Text = "" & "<BR>" & ""
            H4row.Cells.Add(H4tc1d)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            '

            ' 表頭-3行
            xDataTime = ""
            SQL = "SELECT TOP 1 Convert(varchar, CreateTime, 111) + ' ' + SUBSTRING(Convert(varchar, CreateTime, 108),1,2) As DataTime "
            SQL = SQL & "FROM W_IRWNOTICE "
            SQL = SQL & "Order By COUNT DESC "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "NO")
            If DBDataSet1.Tables("NO").Rows.Count > 0 Then
                xDataTime = DBDataSet1.Tables("NO").Rows(0).Item("DataTime")
            End If

            'SQL = "SELECT TOP 1 Convert(varchar, CreateTime, 111) + ' ' + SUBSTRING(Convert(varchar, CreateTime, 108),1,2) As DataTime, YM "
            'SQL = SQL & "FROM W_IRWNOActivityV2  "
            'SQL = SQL & "where Warning='*' "
            'SQL = SQL & "and YM = SUBSTRING(convert(varchar, DATEADD(MONTH, -4,  GETDATE()), 112),1,6) "
            'SQL = SQL & "Order By B_NO DESC "
            'Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            'DBAdapter1.Fill(DBDataSet1, "NO")
            'If DBDataSet1.Tables("NO").Rows.Count > 0 Then
            '    xDataTime = DBDataSet1.Tables("NO").Rows(0).Item("DataTime")
            '    xYM = Mid(DBDataSet1.Tables("NO").Rows(0).Item("YM"), 1, 4) & "/" & Mid(DBDataSet1.Tables("NO").Rows(0).Item("YM"), 5, 2)
            'End If
            '
            Dim H2tc1 As TableCell = New TableCell
            H2tc1.Text = "無效工時(RANKING)"
            H2tc1.ColumnSpan = 2
            H2tc1.BackColor = Color.Red
            H2tc1.HorizontalAlign = HorizontalAlign.Left
            H2row.Cells.Add(H2tc1)

            Dim H2tc2 As TableCell = New TableCell
            H2tc2.Text = "對象期間=2023/10~2024/09 (累計) "
            H2tc2.ColumnSpan = 2
            H2tc2.BackColor = Color.White
            H2tc2.ForeColor = Color.Black
            H2tc2.HorizontalAlign = HorizontalAlign.Left
            H2row.Cells.Add(H2tc2)

            Dim H2tc3 As TableCell = New TableCell
            H2tc3.Text = Mid(xDataTime, 6) & "點更新"
            H2tc3.ColumnSpan = 2
            H2tc3.HorizontalAlign = HorizontalAlign.Right
            H2tc3.BackColor = Color.White
            H2tc3.ForeColor = Color.Black
            H2row.Cells.Add(H2tc3)

            gv.Controls(0).Controls.AddAt(0, H2row)

        End If
    End Sub

    Protected Sub GridView9_RowDataBound1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView9.RowDataBound
        Dim i, idx As Integer
        '
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            For i = 0 To 5
                Select Case i
                    Case 0
                        e.Row.Cells(i).Font.Bold = True
                        e.Row.Cells(i).Font.Size = 12
                        idx = CInt(e.Row.Cells(i).Text)
                        Select Case idx
                            Case Is < 4
                                e.Row.Cells(i).Text = CStr(e.Row.RowIndex + 1)
                                e.Row.Cells(i).ForeColor = Color.Red
                            Case Is < 7
                                e.Row.Cells(i).Text = CStr(e.Row.RowIndex + 1)
                                e.Row.Cells(i).ForeColor = Color.Blue
                            Case Else
                                e.Row.Cells(i).Text = CStr(e.Row.RowIndex + 1)
                        End Select
                    Case 1
                        e.Row.Cells(i).BackColor = Color.Khaki
                    Case 2
                        e.Row.Cells(i).Text = e.Row.Cells(i).Text & " H"
                        e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                        e.Row.Cells(i).BackColor = Color.Pink
                    Case 3
                        e.Row.Cells(i).Text = e.Row.Cells(i).Text & " H"
                    Case 4
                        e.Row.Cells(i).Text = CStr(CInt(CDbl(e.Row.Cells(i).Text) * 100)) / 100 & " %"
                        e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                        e.Row.Cells(i).BackColor = Color.Pink
                    Case Else
                End Select
            Next
        End If


    End Sub

    Protected Sub GridView10_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView10.RowCreated
        '
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim SQL, xDataTime As String
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-4行

            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "目標"
            H4row.Cells.Add(H4tc0)

            Dim H4tc1B As TableCell = New TableCell
            H4tc1B.Text = YYMM(1)
            H4row.Cells.Add(H4tc1B)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = YYMM(2)
            H4row.Cells.Add(H4tc2)

            Dim H4tc2A As TableCell = New TableCell
            H4tc2A.Text = YYMM(3)
            H4tc2A.BackColor = Color.Blue
            H4row.Cells.Add(H4tc2A)

            Dim H4tc2B As TableCell = New TableCell
            H4tc2B.Text = YYMM(4)
            H4row.Cells.Add(H4tc2B)

            Dim H4tc3 As TableCell = New TableCell
            H4tc3.Text = YYMM(5)
            H4row.Cells.Add(H4tc3)

            Dim H4tc3A As TableCell = New TableCell
            H4tc3A.Text = YYMM(6)
            H4row.Cells.Add(H4tc3A)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            ' 表頭-3行
            xDataTime = ""
            SQL = "SELECT TOP 1 Convert(varchar, CreateTime, 111) + ' ' + SUBSTRING(Convert(varchar, CreateTime, 108),1,2) As DataTime "
            SQL = SQL & "FROM W_IRWNOTICE "
            SQL = SQL & "Order By COUNT DESC "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "NO")
            If DBDataSet1.Tables("NO").Rows.Count > 0 Then
                xDataTime = DBDataSet1.Tables("NO").Rows(0).Item("DataTime")
            End If

            Dim H2tc1 As TableCell = New TableCell
            H2tc1.Text = "[全體]未受注顧客服務率"
            H2tc1.ColumnSpan = 2
            H2tc1.BackColor = Color.Red
            H2tc1.HorizontalAlign = HorizontalAlign.Left
            H3row.Cells.Add(H2tc1)

            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = Mid(xDataTime, 6) & "點更新"
            H3tc1.ColumnSpan = 5
            H3tc1.BackColor = Color.White
            H3tc1.ForeColor = Color.Black
            H3tc1.HorizontalAlign = HorizontalAlign.Right
            H3row.Cells.Add(H3tc1)

            gv.Controls(0).Controls.AddAt(0, H3row)
        End If
    End Sub

    Protected Sub GridView10_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView10.RowDataBound
        Dim i As Integer
        '
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            For i = 0 To 6
                Select Case i
                    Case 0
                        e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                        e.Row.Cells(i).BackColor = Color.Pink
                    Case 3
                        e.Row.Cells(i).Text = e.Row.Cells(i).Text & " %"
                        e.Row.Cells(i).ForeColor = Color.Red
                    Case Else
                        e.Row.Cells(i).Text = e.Row.Cells(i).Text & " %"
                End Select
            Next
        End If

    End Sub

    Protected Sub GridView11_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView11.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim SQL, xDataTime As String
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-4行
            Dim H4tcZ As TableCell = New TableCell
            H4tcZ.Text = "TOP." & "<BR>" & ""
            H4row.Cells.Add(H4tcZ)

            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "部門"
            H4row.Cells.Add(H4tc0)

            Dim H4tc1B As TableCell = New TableCell
            H4tc1B.Text = YYMM(1)
            H4row.Cells.Add(H4tc1B)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = YYMM(2)
            H4row.Cells.Add(H4tc2)

            Dim H4tc2A As TableCell = New TableCell
            H4tc2A.Text = YYMM(3)
            H4tc2A.BackColor = Color.Blue
            H4row.Cells.Add(H4tc2A)

            Dim H4tc2B As TableCell = New TableCell
            H4tc2B.Text = YYMM(4)
            H4row.Cells.Add(H4tc2B)

            Dim H4tc3 As TableCell = New TableCell
            H4tc3.Text = YYMM(5)
            H4row.Cells.Add(H4tc3)

            Dim H4tc3A As TableCell = New TableCell
            H4tc3A.Text = YYMM(6)
            H4row.Cells.Add(H4tc3A)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            ' 表頭-3行
            xDataTime = ""
            SQL = "SELECT TOP 1 Convert(varchar, CreateTime, 111) + ' ' + SUBSTRING(Convert(varchar, CreateTime, 108),1,2) As DataTime "
            SQL = SQL & "FROM W_IRWNOTICE "
            SQL = SQL & "Order By COUNT DESC "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "NO")
            If DBDataSet1.Tables("NO").Rows.Count > 0 Then
                xDataTime = DBDataSet1.Tables("NO").Rows(0).Item("DataTime")
            End If

            Dim H2tc1 As TableCell = New TableCell
            H2tc1.Text = "[部門]未受注顧客服務率"
            H2tc1.ColumnSpan = 3
            H2tc1.BackColor = Color.Red
            H2tc1.HorizontalAlign = HorizontalAlign.Left
            H3row.Cells.Add(H2tc1)

            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = Mid(xDataTime, 6) & "點更新"
            H3tc1.ColumnSpan = 5
            H3tc1.BackColor = Color.White
            H3tc1.ForeColor = Color.Black
            H3tc1.HorizontalAlign = HorizontalAlign.Right
            H3row.Cells.Add(H3tc1)

            gv.Controls(0).Controls.AddAt(0, H3row)
        End If
    End Sub

    Protected Sub GridView11_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView11.RowDataBound
        Dim i, idx As Integer

        If (e.Row.RowType = DataControlRowType.DataRow) Then
            For i = 0 To 7
                Select Case i
                    Case 0
                        e.Row.Cells(i).Font.Bold = True
                        e.Row.Cells(i).Font.Size = 12
                        idx = CInt(e.Row.Cells(i).Text)
                        Select Case idx
                            Case Is < 4
                                e.Row.Cells(i).Text = CStr(e.Row.RowIndex + 1)
                                e.Row.Cells(i).ForeColor = Color.Red
                            Case Is < 7
                                e.Row.Cells(i).Text = CStr(e.Row.RowIndex + 1)
                                e.Row.Cells(i).ForeColor = Color.Blue
                            Case Else
                                e.Row.Cells(i).Text = CStr(e.Row.RowIndex + 1)
                        End Select
                    Case 1
                    Case 4
                        e.Row.Cells(i).Text = e.Row.Cells(i).Text & " %"
                        e.Row.Cells(i).ForeColor = Color.Red
                    Case Else
                        e.Row.Cells(i).Text = e.Row.Cells(i).Text & " %"
                End Select
            Next
        End If

    End Sub
End Class
