Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class IRWNoticePage2310
    Inherits System.Web.UI.Page


    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim uJavaScript As New Utility.JScript
    '---------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String
    Dim pDepCode As String
    Dim pYM As String

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "IRWNoticePageV2.aspx"

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
        pDepCode = Request.QueryString("pDepCode")
        pYM = Request.QueryString("pYM")
        GridView4.Visible = False
        GridView9.Visible = False
        '
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet4 As New DataSet
        Dim DBDataSet9 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '
        SQL = "SELECT "
        SQL = SQL & "Warning, DepName,A_YES,A_NO,A_TOTAL,A_PERCENT,S_PERCENT,S_NO,H_NO,S_NO+H_NO AS L_NO,B_NO, H2P_NO "
        SQL = SQL & "FROM W_IRWNOActivityV2  "
        SQL = SQL & "WHERE YM = '" & pYM & "' "
        SQL = SQL & "Order By Warning DESC, H_NO DESC, B_NO DESC "
        '
        Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter4.Fill(DBDataSet4, "D")
        GridView4.DataSource = DBDataSet4
        GridView4.DataBind()
        GridView4.Visible = True
        OleDbConnection1.Close()

        '
        SQL = "SELECT "
        SQL = SQL & "[DepName],[EmpName],[A_YES],[A_NO],[A_TOTAL],[A_PERCENT],[G_No],[L_No],[Remark] "
        '
        SQL = SQL & "FROM W_IRWNOActivityV2Person  "
        SQL = SQL & "where DepCode = '" & pDepCode & "' "
        SQL = SQL & "AND YM = '" & pYM & "' "
        SQL = SQL & "ORDER BY UNIQUE_ID "
        '
        Dim DBAdapter9 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter9.Fill(DBDataSet9, "CHECK")
        GridView9.DataSource = DBDataSet9
        GridView9.DataBind()
        GridView9.Visible = True
        OleDbConnection1.Close()
        '
    End Sub

    Protected Sub GridView4_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView4.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            '
            Dim SQL, xDataTime, xYM, xCheckYM As String
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
            Dim H4tcZ As TableCell = New TableCell
            H4tcZ.Text = "" & "<BR>" & "　"
            H4row.Cells.Add(H4tcZ)

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
            H3tc1.ColumnSpan = 6
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
            SQL = SQL & "AND YM = '" & pYM & "' "
            SQL = SQL & "Order By B_NO DESC "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "NO")
            If DBDataSet1.Tables("NO").Rows.Count > 0 Then
                xDataTime = DBDataSet1.Tables("NO").Rows(0).Item("DataTime")
                xYM = Mid(DBDataSet1.Tables("NO").Rows(0).Item("YM"), 1, 4) & "/" & Mid(DBDataSet1.Tables("NO").Rows(0).Item("YM"), 5, 2)
            End If
            '
            DBDataSet1.Clear()
            xCheckYM = ""
            SQL = "SELECT COUNT(*) as RCount "
            SQL = SQL & "FROM W_IRWNOActivityV2  "
            SQL = SQL & "where YM > '" & pYM & "' "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "NO")
            If DBDataSet1.Tables("NO").Rows.Count > 0 Then
                If DBDataSet1.Tables("NO").Rows(0).Item("RCount") > 0 Then
                    xCheckYM = "Y"
                End If
            End If
            '-----------------------------------------
            Dim H2tc1 As TableCell = New TableCell
            If xCheckYM = "Y" Then
                H2tc1.Text = "違反部門"
                H2tc1.BackColor = Color.Red
            Else
                H2tc1.Text = "[注意] 次月需注意部門"
                H2tc1.BackColor = Color.Navy
            End If
            H2tc1.ColumnSpan = 3
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
            H2tc2.ColumnSpan = 5
            H2tc2.HorizontalAlign = HorizontalAlign.Right
            H2tc2.BackColor = Color.White
            H2tc2.ForeColor = Color.Black
            H2row.Cells.Add(H2tc2)

            gv.Controls(0).Controls.AddAt(0, H2row)
        End If
    End Sub

    Protected Sub GridView9_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView9.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            '
            Dim SQL, xDataTime, xYM, xCheckYM As String
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
            '
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
            H4row.Cells.Add(H4tc1A)

            Dim H4tc1B As TableCell = New TableCell
            H4tc1B.Text = "TOTAL" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc1B)

            Dim H4tc1C As TableCell = New TableCell
            H4tc1C.Text = "NG" & "<BR>" & "比例"
            H4row.Cells.Add(H4tc1C)

            Dim H4tc1CA As TableCell = New TableCell
            H4tc1CA.Text = "NG" & "<BR>" & "COUNT"
            H4row.Cells.Add(H4tc1CA)

            Dim H4tc1CB As TableCell = New TableCell
            H4tc1CB.Text = "限制" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc1CB)

            Dim H4tc1D As TableCell = New TableCell
            H4tc1D.Text = "REMARK" & "<BR>" & "　"
            H4row.Cells.Add(H4tc1D)

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
            H3tc2.ColumnSpan = 3
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
            SQL = SQL & "AND YM = '" & pYM & "' "
            SQL = SQL & "Order By B_NO DESC "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "NO")
            If DBDataSet1.Tables("NO").Rows.Count > 0 Then
                xDataTime = DBDataSet1.Tables("NO").Rows(0).Item("DataTime")
                xYM = Mid(DBDataSet1.Tables("NO").Rows(0).Item("YM"), 1, 4) & "/" & Mid(DBDataSet1.Tables("NO").Rows(0).Item("YM"), 5, 2)
            End If
            '
            DBDataSet1.Clear()
            xCheckYM = ""
            SQL = "SELECT COUNT(*) as RCount "
            SQL = SQL & "FROM W_IRWNOActivityV2  "
            SQL = SQL & "where YM > '" & pYM & "' "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "NO")
            If DBDataSet1.Tables("NO").Rows.Count > 0 Then
                If DBDataSet1.Tables("NO").Rows(0).Item("RCount") > 0 Then
                    xCheckYM = "Y"
                End If
            End If
            '
            Dim H2tc1 As TableCell = New TableCell
            If xCheckYM = "Y" Then
                H2tc1.Text = "報告提出者"
                H2tc1.BackColor = Color.Red
            Else
                H2tc1.Text = "[注意] 次月需注意個人"
                H2tc1.BackColor = Color.Navy
            End If
            H2tc1.ColumnSpan = 3
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
            H2tc2.ColumnSpan = 3
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
            For i = 0 To 10
                Select Case i
                    Case 0, 1
                        e.Row.Cells(i).BackColor = Color.Khaki
                    Case 3, 9, 10
                        If e.Row.Cells(0).Text = "*" Then
                            e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                            e.Row.Cells(i).BackColor = Color.Pink
                        End If
                    Case 5
                        If CDbl(e.Row.Cells(i).Text) < 0 Then
                            e.Row.Cells(i).Text = "0" & "%"
                        Else
                            e.Row.Cells(i).Text = e.Row.Cells(i).Text & "%"
                        End If
                    Case 6
                        e.Row.Cells(i).Text = "15" & "%"
                    Case Else
                End Select
            Next
        End If
    End Sub


    Protected Sub GridView9_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView9.RowDataBound
        Dim i As Integer
        '
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            For i = 0 To 8
                If CDbl(e.Row.Cells(6).Text) > CDbl(e.Row.Cells(7).Text) Then
                    e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                    e.Row.Cells(i).BackColor = Color.Pink
                End If
                '
                Select Case i
                    Case 6
                        If CDbl(e.Row.Cells(6).Text) > CDbl(e.Row.Cells(7).Text) Then
                            e.Row.Cells(i).Attributes.Add("style", "border:1px solid blue ")
                            e.Row.Cells(i).BackColor = Color.Blue
                            e.Row.Cells(i).ForeColor = Color.White
                        End If
                    Case Else
                End Select
            Next
        End If

    End Sub

End Class
