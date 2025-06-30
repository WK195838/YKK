Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class IRWNoticePageZip
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
        Response.Cookies("PGM").Value = "IRWNoticePageZIP.aspx"

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
        LSearch.Visible = True
        '
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim xDays As Double
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '
        SQL = "SELECT "
        SQL = SQL & "case when YYMM='2022/07' then '2022/07' else YYMM end AS YYMM, "

        SQL = SQL & "case when datatype='A' then '1 ~10' when datatype='B' then '11~20' when datatype='C' then '21~ E' when datatype='ALL' then 'ALL' else '-' end AS DATATYPEDESC, "
        SQL = SQL & "sales, zip, sld, total, target, banlance, "
        '
        SQL = SQL & "'IRWNoticePageZIP01.aspx?pYYMM=' + YYMM + '&pDataType=' + datatype + '&pCat=1.SALES' as salesurl, "
        SQL = SQL & "'IRWNoticePageZIP01.aspx?pYYMM=' + YYMM + '&pDataType=' + datatype + '&pCat=2.ZIP' as zipurl, "
        SQL = SQL & "'IRWNoticePageZIP01.aspx?pYYMM=' + YYMM + '&pDataType=' + datatype + '&pCat=3.SLD' as SLDurl "
        SQL = SQL & "FROM W_IRWFlowTime "
        'JOY-TEST
        'SQL = SQL & "WHERE YYMM < '2023/10' "
        '
        SQL = SQL & "Order By YYMM Desc, DataType Desc "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "YY")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()
        GridView1.Visible = True
        OleDbConnection1.Close()
        '
        SQL = "SELECT top 1  "
        SQL = SQL & "Total, Convert(varchar, CreateTime, 111) + ' ' + SUBSTRING(Convert(varchar, CreateTime, 108),1,2) As DataTime "
        SQL = SQL & "FROM W_IRWFlowTime "
        SQL = SQL & "Order By YYMM DESC, DATATYPE DESC "
        '
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "Leadtime")
        If DBDataSet2.Tables("Leadtime").Rows.Count > 0 Then
            xDays = CInt((DBDataSet2.Tables("Leadtime").Rows(0).Item("Total") / 8 + 0.005) * 100) / 100
            DNowDays.Text = CStr(xDays + 30)
            DIRWDays.Text = CStr(DBDataSet2.Tables("Leadtime").Rows(0).Item("Total")) & " H ≒ " & CStr(xDays) & " 日"
            '
            Label2.Text = Mid(DBDataSet2.Tables("Leadtime").Rows(0).Item("DataTime"), 6) & "點更新"
        End If
        OleDbConnection1.Close()
        '
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
            H4tc1.Text = "日"
            H4row.Cells.Add(H4tc1)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = "SALES"
            H4row.Cells.Add(H4tc2)

            Dim H4tc3 As TableCell = New TableCell
            H4tc3.Text = "ZIP"
            H4row.Cells.Add(H4tc3)

            Dim H4tc4 As TableCell = New TableCell
            H4tc4.Text = "SLD"
            H4row.Cells.Add(H4tc4)

            Dim H4tc5 As TableCell = New TableCell
            H4tc5.Text = "實績(H)"
            H4row.Cells.Add(H4tc5)

            Dim H4tc6 As TableCell = New TableCell
            H4tc6.Text = "目標(H)"
            H4row.Cells.Add(H4tc6)

            Dim H4tc7 As TableCell = New TableCell
            H4tc7.Text = "差(H)"
            H4row.Cells.Add(H4tc7)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            '' 表頭-行
            '
            xDataTime = ""
            SQL = "SELECT TOP 1 Convert(varchar, CreateTime, 111) + ' ' + SUBSTRING(Convert(varchar, CreateTime, 108),1,2) As DataTime "
            SQL = SQL & "FROM W_IRWFlowTime "
            SQL = SQL & "Order By Unique_ID "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "ZIP")
            If DBDataSet1.Tables("ZIP").Rows.Count > 0 Then
                xDataTime = DBDataSet1.Tables("ZIP").Rows(0).Item("DataTime")
            End If
            '
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "ITEM申請實績"
            H3tc1.ColumnSpan = 6
            H3tc1.BackColor = Color.Blue
            H3row.Cells.Add(H3tc1)

            Dim H3tc2 As TableCell = New TableCell
            H3tc2.Text = Mid(xDataTime, 6) & "點更新"
            H3tc2.ColumnSpan = 2
            H3tc2.BackColor = Color.Blue
            H3row.Cells.Add(H3tc2)
            '
            gv.Controls(0).Controls.AddAt(0, H3row)
        End If

    End Sub
    '
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        If (e.Row.RowType = DataControlRowType.DataRow) Then

            For i = 0 To 7
                Select Case i
                    Case 0
                        If InStr(e.Row.Cells(0).Text, "2022/07") > 0 Then
                            e.Row.Cells(i).Text = "Before."
                        End If
                    Case 2, 3, 4
                        If InStr(e.Row.Cells(0).Text, "Before.") > 0 Then
                            e.Row.Cells(i).Text = "-"
                        End If
                    Case 5
                        If CDbl(e.Row.Cells(i).Text) > 16 Then
                            e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                            e.Row.Cells(i).BackColor = Color.Pink
                        End If
                    Case 7
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
