Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class IRWNoticePageRanking
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    '外部Object
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim uJavaScript As New Utility.JScript
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String
    Dim UserID As String            'UserID
    Dim pEmpID As String

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "IRWNoticePageRanking.aspx"

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
        UserID = Request.QueryString("pUserID")             'UserID
        pEmpID = Request.QueryString("pEmpID")
        '
    End Sub

    Sub DataList()
        Dim SQL, xTop As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet9 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '-------------------------------
        '
        xTop = ""
        SQL = "SELECT TOP 10  "
        SQL = SQL & "ROW_NUMBER() OVER (ORDER BY  SUM( CAST( REPLACE( REPLACE(NoworkHour1, ' H','') ,' ','' ) AS decimal(18, 2) ) ) desc) as Rank, "
        SQL = SQL & "DEPNAME + '-' + APPNAME as DepName, AppID "
        SQL = SQL & "FROM F_NoOrderReortSheet "
        SQL = SQL & "GROUP BY DEPNAME, APPNAME, AppID "
        SQL = SQL & "ORDER BY SUM( CAST( REPLACE( REPLACE(NoworkHour1, ' H','') ,' ','' ) AS decimal(18, 2) ) ) desc "
        Dim dt_Order As DataTable = uDataBase.GetDataTable(SQL)
        For i As Integer = 0 To dt_Order.Rows.Count - 1
            If dt_Order.Rows(i)("AppID").ToString = pEmpID Then
                xTop = dt_Order.Rows(i)("Rank").ToString
                Exit For
            End If
        Next
        '
        SQL = "SELECT "
        SQL = SQL & " '" & xTop & "' as Rank, DEPNAME + '-' + APPNAME as DepName, "
        '
        SQL = SQL & "SUM( CAST( REPLACE( REPLACE(NoworkHour1, ' H','') ,' ','' ) AS decimal(18, 2) ) ) as Hour, "
        SQL = SQL & "(select SUM( CAST( REPLACE( REPLACE(NoworkHour1, ' H','') ,' ','' ) AS decimal(18, 2) ) ) from F_NoOrderReortSheet) as HourTotal, "
        SQL = SQL & "SUM( CAST( REPLACE( REPLACE(NoworkHour1, ' H','') ,' ','' ) AS decimal(18, 2) ) ) "
        SQL = SQL & "/"
        SQL = SQL & "(select SUM( CAST( REPLACE( REPLACE(NoworkHour1, ' H','') ,' ','' ) AS decimal(18, 2) ) ) from F_NoOrderReortSheet) "
        SQL = SQL & "*"
        SQL = SQL & "100"
        SQL = SQL & "as HourPer "
        '
        SQL = SQL & "FROM F_NoOrderReortSheet "
        SQL = SQL & "Where AppID = '" & pEmpID & "' "
        SQL = SQL & "GROUP BY DEPNAME, APPNAME, AppID "
        SQL = SQL & "ORDER BY SUM( CAST( REPLACE( REPLACE(NoworkHour1, ' H','') ,' ','' ) AS decimal(18, 2) ) ) desc "
        '
        Dim DBAdapter9 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter9.Fill(DBDataSet9, "SUM")
        '
        GridView9.DataSource = DBDataSet9
        GridView9.DataBind()
        GridView9.Visible = True
        '
        OleDbConnection1.Close()
        '-------------------------------
        SQL = "SELECT "
        SQL = SQL & "[FormNo],[FormSno], Description, "
        SQL = SQL & "[No], 'http://10.245.1.6/IRW/IRWNoOrderReportSheet_02.aspx?pUserid=&pFormNo=001171&pFormSno=' + ltrim(rtrim(formsno)) as NOUrl, "
        '
        SQL = SQL & "[LeadDate],[AppID],[DepName] + '-' + [AppName] as DepName, "
        SQL = SQL & "[ApplyQty],[OrderQty],[NoOrderQty],[Percen], "
        SQL = SQL & "[NoworkHour1],[NoworkHour2] "
        '
        SQL = SQL & "FROM F_NoOrderReortSheet "
        SQL = SQL & "Where AppID = '" & pEmpID & "' "
        SQL = SQL & "Order By LeadDate DESC "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "NoOrder")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()
        GridView1.Visible = True
        OleDbConnection1.Close()
        '
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            Dim H4tc11 As TableCell = New TableCell
            H4tc11.Text = "年月" & "<BR>" & ""
            H4row.Cells.Add(H4tc11)

            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "No" & "<BR>" & ""
            H4row.Cells.Add(H4tc0)

            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "NAME" & "<BR>" & ""
            H4row.Cells.Add(H4tc1)

            Dim H4tc1A As TableCell = New TableCell
            H4tc1A.Text = "ITEM" & "<BR>" & "申請件數(a)"
            H4tc1A.BackColor = Color.Blue
            H4row.Cells.Add(H4tc1A)

            Dim H4tc1B As TableCell = New TableCell
            H4tc1B.Text = "受注數" & "<BR>" & "(b)"
            H4tc1B.BackColor = Color.Blue
            H4row.Cells.Add(H4tc1B)

            Dim H4tc1C As TableCell = New TableCell
            H4tc1C.Text = "未受注數" & "<BR>" & "(c)"
            H4tc1C.BackColor = Color.Blue
            H4row.Cells.Add(H4tc1C)

            Dim H4tc1D As TableCell = New TableCell
            H4tc1D.Text = "比例" & "<BR>" & "(c) / (a)"
            H4tc1D.BackColor = Color.Green
            H4row.Cells.Add(H4tc1D)

            Dim H4tc1E As TableCell = New TableCell
            H4tc1E.Text = "無效工時" & "<BR>" & ""
            H4tc1E.BackColor = Color.Red
            H4row.Cells.Add(H4tc1E)

            Dim H4tc1F As TableCell = New TableCell
            H4tc1F.Text = "估算公式" & "<BR>" & ""
            H4row.Cells.Add(H4tc1F)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
        End If
    End Sub
    '
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer

        If (e.Row.RowType = DataControlRowType.DataRow) Then
            For i = 0 To 8
                'Select Case i
                '    Case 2
                '        e.Row.Cells(i).Text = Format(CDbl(e.Row.Cells(i).Text), "###,###,###")
                '    Case Else
                'End Select
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

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            '
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
            H2tc1.Text = "無效工時(RANKING)"
            H2tc1.ColumnSpan = 2
            H2tc1.BackColor = Color.Red
            H2tc1.HorizontalAlign = HorizontalAlign.Left
            H2row.Cells.Add(H2tc1)

            Dim H2tc2 As TableCell = New TableCell
            H2tc2.Text = "對象期間=2023/10~2023/12 (累計) "
            H2tc2.ColumnSpan = 2
            H2tc2.BackColor = Color.White
            H2tc2.ForeColor = Color.Black
            H2tc2.HorizontalAlign = HorizontalAlign.Left
            H2row.Cells.Add(H2tc2)

            Dim H2tc3 As TableCell = New TableCell
            H2tc3.Text = Mid(xDataTime, 6) & "點更新"
            H2tc3.ColumnSpan = 1
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
            For i = 0 To 4
                Select Case i
                    Case 0
                        e.Row.Cells(i).Font.Bold = True
                        e.Row.Cells(i).Font.Size = 12
                        idx = CInt(e.Row.Cells(i).Text)
                        Select Case idx
                            Case Is < 4
                                e.Row.Cells(i).ForeColor = Color.Red
                            Case Is < 7
                                e.Row.Cells(i).ForeColor = Color.Blue
                            Case Else
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
    '
End Class
