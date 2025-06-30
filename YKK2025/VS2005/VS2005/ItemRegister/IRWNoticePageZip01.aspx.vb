Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class IRWNoticePageZip01
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
    Dim pYYMM As String
    Dim pDataType As String
    Dim pCat As String


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "IRWNoticePageZIP01.aspx"

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
        pYYMM = Request.QueryString("pYYMM")
        pDataType = Request.QueryString("pDataType")
        pCat = Request.QueryString("pCat")
        '
        GridView1.Visible = False
        '
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '
        SQL = "SELECT "
        SQL = SQL & "YYMM, "
        SQL = SQL & "case when datatype='A' then '1~10' when datatype='B' then '11~20' else '21~ ' end AS DATATYPEDESC, "
        SQL = SQL & "cat, no, '....' as OP, furl as nourl, opurl as opurl, REJECT, CASE WHEN REJECT='Y' THEN RHH ELSE 0 END AS RHH, CASE WHEN REJECT='Y' THEN RHHSales ELSE 0 END AS RHHSales,  "
        SQL = SQL & "convert(varchar, sendtime, 120) AS sendtime, "
        SQL = SQL & "convert(varchar, CODETIME, 120) AS codetime, "
        SQL = SQL & "convert(varchar, starttime, 120) AS starttime, "
        SQL = SQL & "convert(varchar, endTime, 120) AS endtime, "
        '
        SQL = SQL & "hh, "
        SQL = SQL & "case when YYMM >= '2023/10' then '4' else '8' end AS stdhh, "
        SQL = SQL & "case when YYMM >= '2023/10' then hh-4 else hh-8 end AS Balance "
        'SQL = SQL & "hh, '4' as stdhh, hh-4 as Balance "
        '
        SQL = SQL & "FROM W_IRWFlowTime_01 "
        SQL = SQL & "where YYMM = '" & pYYMM & "' "
        If pDataType <> "ALL" Then
            SQL = SQL & "and DataType = '" & pDataType & "' "
        End If

        SQL = SQL & "and Cat = '" & pCat & "' "
        SQL = SQL & "Order By YYMM , DataType DESC, cat, no Desc "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "YY")
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
            Dim H9row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()

            '-----------------------------------------
            ' 表頭-8行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "YYMM"
            H4row.Cells.Add(H4tc0)

            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "日"
            H4row.Cells.Add(H4tc1)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = "類別"
            H4row.Cells.Add(H4tc2)

            Dim H4tc3 As TableCell = New TableCell
            H4tc3.Text = "NO"
            H4row.Cells.Add(H4tc3)

            Dim H4tc4 As TableCell = New TableCell
            H4tc4.Text = "承認" & "<br>" & "履歷"
            H4row.Cells.Add(H4tc4)

            Dim H4tcA2 As TableCell = New TableCell
            H4tcA2.Text = ""
            H4row.Cells.Add(H4tcA2)

            Dim H4tcA3 As TableCell = New TableCell
            H4tcA3.Text = "[營+生]" & "<br>" & "(H)"
            H4row.Cells.Add(H4tcA3)

            'Dim H4tcA4 As TableCell = New TableCell
            'H4tcA4.Text = "[營]" & "<br>" & "(H)"
            'H4row.Cells.Add(H4tcA4)

            Dim H4tc5 As TableCell = New TableCell
            H4tc5.Text = "營業完成"
            H4row.Cells.Add(H4tc5)

            Dim H4tcA1 As TableCell = New TableCell
            H4tcA1.Text = "CODE完成"
            H4row.Cells.Add(H4tcA1)

            Dim H4tc6 As TableCell = New TableCell
            H4tc6.Text = "計算-開始"
            H4row.Cells.Add(H4tc6)

            Dim H4tc7 As TableCell = New TableCell
            H4tc7.Text = "計算-完成"
            H4row.Cells.Add(H4tc7)

            Dim H4tc8 As TableCell = New TableCell
            H4tc8.Text = "實績" & "<br>" & "(H)"
            H4row.Cells.Add(H4tc8)

            Dim H4tc9 As TableCell = New TableCell
            H4tc9.Text = "目標" & "<br>" & "(H)"
            H4row.Cells.Add(H4tc9)

            Dim H4tcA As TableCell = New TableCell
            H4tcA.Text = "差" & "<br>" & "(H)"
            H4row.Cells.Add(H4tcA)

            gv.Controls(0).Controls.AddAt(0, H4row)
            ''-----------------------------------------
            ' '' 表頭-行
            '-----------------------------------------
            ' 表頭-8行
            Dim H9tc0 As TableCell = New TableCell
            H9tc0.Text = ""
            H9row.Cells.Add(H9tc0)

            Dim H9tc1 As TableCell = New TableCell
            H9tc1.Text = ""
            H9row.Cells.Add(H9tc1)

            Dim H9tc2 As TableCell = New TableCell
            H9tc2.Text = ""
            H9row.Cells.Add(H9tc2)

            Dim H9tc3 As TableCell = New TableCell
            H9tc3.Text = "委託書"
            H9tc3.ColumnSpan = 2
            H9row.Cells.Add(H9tc3)

            Dim H9tc4 As TableCell = New TableCell
            H9tc4.Text = "退件工時"
            H9tc4.ColumnSpan = 2
            H9row.Cells.Add(H9tc4)

            Dim H9tc5 As TableCell = New TableCell
            H9tc5.Text = "處理時間"
            H9tc5.ColumnSpan = 2
            H9row.Cells.Add(H9tc5)

            Dim H9tc6 As TableCell = New TableCell
            H9tc6.Text = "統計時間"
            H9tc6.ColumnSpan = 2
            H9row.Cells.Add(H9tc6)

            Dim H9tc7 As TableCell = New TableCell
            H9tc7.Text = "目標"
            H9tc7.ColumnSpan = 3
            H9row.Cells.Add(H9tc7)

            gv.Controls(0).Controls.AddAt(0, H9row)

            '-----------------------------------------
            '' 表頭-行
        End If
    End Sub
    '
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer

        If (e.Row.RowType = DataControlRowType.DataRow) Then

            For i = 0 To 14
                Select Case i
                    Case 6
                        If CDbl(e.Row.Cells(i).Text) > 0 Then
                            If CDbl(e.Row.Cells(14).Text) > 0 Then
                                If CDbl(e.Row.Cells(14).Text) <= CDbl(e.Row.Cells(i).Text) Then
                                    e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                                    e.Row.Cells(i).BackColor = Color.Pink
                                End If
                            End If
                        Else
                            e.Row.Cells(i).Text = "0"
                        End If
                    Case 7
                        If CDbl(e.Row.Cells(i).Text) < 0 Then
                            If CDbl(e.Row.Cells(6).Text) > 0.1 Then
                                e.Row.Cells(i).Text = "0.1"
                            Else
                                e.Row.Cells(i).Text = e.Row.Cells(6).Text
                            End If
                        Else
                            If CDbl(e.Row.Cells(i).Text) <= 0 Then e.Row.Cells(i).Text = "0"
                        End If
                        e.Row.Cells(i).Visible = False
                    Case 12
                        If CDbl(e.Row.Cells(i).Text) > 4 Then
                            e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                            e.Row.Cells(i).BackColor = Color.Pink
                        End If
                    Case 14
                        If CDbl(e.Row.Cells(i).Text) > 0 Then
                            e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                            e.Row.Cells(i).BackColor = Color.Pink
                        End If
                    Case Else
                End Select
            Next
        End If

    End Sub

End Class
