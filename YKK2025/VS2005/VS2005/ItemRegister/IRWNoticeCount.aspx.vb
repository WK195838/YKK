Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class IRWNoticeCount
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim uJavaScript As New Utility.JScript
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "IRWNoticeCount.aspx"

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
        '
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '
        SQL = "SELECT "
        SQL = SQL & "YYMM, DayCount, WorkDay, Records as Records, LYRecords, Banlance, Remark "
        SQL = SQL & "FROM W_IRWCount "
        SQL = SQL & "Order By YYMM DESC "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "YY")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()
        GridView1.Visible = True
        OleDbConnection1.Close()
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
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
            ' 表頭-8行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "YYMM"
            H4row.Cells.Add(H4tc0)

            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "日均" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc1)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = "工作" & "<BR>" & "日數"
            H4row.Cells.Add(H4tc2)

            Dim H4tc3 As TableCell = New TableCell
            H4tc3.Text = "實績/推移" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc3)

            Dim H4tc4 As TableCell = New TableCell
            H4tc4.Text = "去年同月" & "<BR>" & "件數"
            H4row.Cells.Add(H4tc4)

            Dim H4tc5 As TableCell = New TableCell
            H4tc5.Text = "" & "<BR>" & "(%)"
            H4row.Cells.Add(H4tc5)

            Dim H4tc6 As TableCell = New TableCell
            H4tc6.Text = "備註"
            H4row.Cells.Add(H4tc6)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            '' 表頭-行
            '
            xDataTime = ""
            SQL = "SELECT TOP 1 Convert(varchar, CreateTime, 111) + ' ' + SUBSTRING(Convert(varchar, CreateTime, 108),1,2) As DataTime "
            SQL = SQL & "FROM  W_IRWCount "
            SQL = SQL & "Order By Unique_ID "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "ZIP")
            If DBDataSet1.Tables("ZIP").Rows.Count > 0 Then
                xDataTime = DBDataSet1.Tables("ZIP").Rows(0).Item("DataTime")
            End If
            '
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "ITEM申請實績&推移"
            H3tc1.ColumnSpan = 6
            H3tc1.BackColor = Color.Blue
            H3row.Cells.Add(H3tc1)

            Dim H3tc2 As TableCell = New TableCell
            H3tc2.Text = Mid(xDataTime, 6) & "點更新"
            H3tc2.ColumnSpan = 1
            H3tc2.BackColor = Color.Blue
            H3row.Cells.Add(H3tc2)
            '
            gv.Controls(0).Controls.AddAt(0, H3row)
        End If

    End Sub

    Protected Sub GridView2_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim SQL, xDataTime As String
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H5row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-8行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = ""
            H4row.Cells.Add(H4tc0)

            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "NAME"
            H4row.Cells.Add(H4tc1)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = "08"
            H4row.Cells.Add(H4tc2)

            Dim H4tc3 As TableCell = New TableCell
            H4tc3.Text = "09"
            H4row.Cells.Add(H4tc3)

            Dim H4tc4 As TableCell = New TableCell
            H4tc4.Text = "10"
            H4row.Cells.Add(H4tc4)

            Dim H4tc5 As TableCell = New TableCell
            H4tc5.Text = "11"
            H4row.Cells.Add(H4tc5)

            Dim H4tc6 As TableCell = New TableCell
            H4tc6.Text = "12"
            H4row.Cells.Add(H4tc6)

            Dim H4tc7 As TableCell = New TableCell
            H4tc7.Text = "01"
            H4row.Cells.Add(H4tc7)

            Dim H4tc8 As TableCell = New TableCell
            H4tc8.Text = "02"
            H4row.Cells.Add(H4tc8)

            Dim H4tc9 As TableCell = New TableCell
            H4tc9.Text = "03"
            H4row.Cells.Add(H4tc9)

            Dim H4tc10 As TableCell = New TableCell
            H4tc10.Text = "04"
            H4row.Cells.Add(H4tc10)

            Dim H4tc11 As TableCell = New TableCell
            H4tc11.Text = "05"
            H4row.Cells.Add(H4tc11)

            Dim H4tc12 As TableCell = New TableCell
            H4tc12.Text = "06"
            H4row.Cells.Add(H4tc12)

            Dim H4tc13 As TableCell = New TableCell
            H4tc13.Text = "07"
            H4row.Cells.Add(H4tc13)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            ' 表頭-行
            '
            Dim H5tc1 As TableCell = New TableCell
            H5tc1.Text = "[Y]: >40件　[Y/R]:需報告　[Y/N]:監查時點Y後降至<40件"
            H5tc1.ColumnSpan = 14
            H5tc1.BackColor = Color.White
            H5tc1.ForeColor = Color.Red
            H5tc1.Font.Bold = False
            H5tc1.Style.Add("text-align", "left")
            H5row.Cells.Add(H5tc1)
            '
            gv.Controls(0).Controls.AddAt(0, H5row)

            '-----------------------------------------
            ' 表頭-行
            '
            xDataTime = ""
            SQL = "SELECT TOP 1 Convert(varchar, CreateTime, 111) + ' ' + SUBSTRING(Convert(varchar, CreateTime, 108),1,2) As DataTime "
            SQL = SQL & "FROM  W_IRWTopNoActive_WARN "
            SQL = SQL & "Order By Unique_ID "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "CHECK")
            If DBDataSet1.Tables("CHECK").Rows.Count > 0 Then
                xDataTime = DBDataSet1.Tables("CHECK").Rows(0).Item("DataTime")
            End If
            '
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "監查警告履歷 (2022/8~2023/7)"
            H3tc1.ColumnSpan = 8
            H3tc1.BackColor = Color.Blue
            H3row.Cells.Add(H3tc1)

            Dim H3tc2 As TableCell = New TableCell
            H3tc2.Text = Mid(xDataTime, 6) & "點更新"
            H3tc2.ColumnSpan = 6
            H3tc2.BackColor = Color.Blue
            H3row.Cells.Add(H3tc2)
            '
            gv.Controls(0).Controls.AddAt(0, H3row)
        End If
    End Sub
    '
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        Dim str As String
        '
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            For i = 0 To 6
                Select Case i
                    Case 3, 4
                        e.Row.Cells(i).Text = Format(CDbl(e.Row.Cells(i).Text), "###,###,###")
                    Case 5
                        e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                        e.Row.Cells(i).BackColor = Color.Pink
                        str = Format(100 - CDbl(e.Row.Cells(i).Text) * 100, "###,###,###")
                        If str = "" Then str = "-0"
                        If CDbl(str) > 0 Then
                            e.Row.Cells(i).Text = "↓ " & Replace(str, "-", "") & "%"
                        Else
                            If CDbl(str) < 0 Then
                                e.Row.Cells(i).Text = "↑ " & Replace(str, "-", "") & "%"
                            Else
                                e.Row.Cells(i).Text = "　 " & Replace(str, "-", "") & "%"
                            End If
                        End If
                    Case Else
                End Select
            Next
        End If

    End Sub

    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        Dim i As Integer
        Dim xCount As Integer
        '
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            xCount = 0
            For i = 0 To 13
                If e.Row.Cells(i).Text = "Y" Or e.Row.Cells(i).Text = "Y/R" Then
                    xCount = xCount + 1
                Else
                    xCount = 0
                End If
                '
                If e.Row.Cells(i).Text = "Y" Or e.Row.Cells(i).Text = "Y/R" Then
                    If xCount > 1 Then
                        e.Row.Cells(i - 1).Attributes.Add("style", "border:1px solid red ")
                        e.Row.Cells(i - 1).BackColor = Color.Pink
                        e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                        e.Row.Cells(i).BackColor = Color.Pink
                    End If
                End If
                '
                Select Case i
                    Case Is > 1
                        'e.Row.Cells(i).Font.Size = 7
                    Case Else
                End Select
            Next
        End If

    End Sub
End Class
