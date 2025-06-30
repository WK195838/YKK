Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class InqManufOutList
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DYear As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DMonth As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSOrB As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSts As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

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
    Dim pFormNo As String
    Dim NowYear As String
    Dim NowDateTime As String       '現在日期時間

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "InqManufOutList.aspx"

        SetParameter()
        If Not Me.IsPostBack Then
            SetYear()
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '現在日時
        NowYear = CStr(DateTime.Now.Year)

        ''pFormNo = Request.QueryString("pFormNo")    '表單號碼
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")

    End Sub

    '*****************************************************************
    '**
    '**     篩選外注委託書委託年份
    '**
    '*****************************************************************
    Sub SetYear()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        SQL = "SELECT CONVERT(varchar(4), CreateTime, 120) AS DYear FROM F_ManufOutSheet "
        SQL = SQL + "GROUP BY CONVERT(varchar(4), CreateTime, 120)"

        OleDbConnection1.Open()
        DYear.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ManufOutSheet")
        DBTable1 = DBDataSet1.Tables("F_ManufOutSheet")
        '將篩選出的年份加入DropDownList
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("DYear")
            ListItem1.Value = DBTable1.Rows(i).Item("DYear")
            If ListItem1.Value = NowYear Then ListItem1.Selected = True
            DYear.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        CreateWorkTable()
    End Sub

    Sub CreateWorkTable()
        Dim i, j As Integer
        Dim MonthlyCount, CumulativeCount As Integer             '每月件數,累計件數
        Dim SQL, SQL1, SQL2 As String
        Dim wTableName As String = "F_ManufOutSheet"
        Dim wMDate As String = DYear.SelectedValue + "-" + DMonth.SelectedValue + "-1"  '篩選日期範圍
        Dim wCDate As String = DYear.SelectedValue + "-1-1"                             '累計日期範圍
        Dim wMTotal, wCTotal As Integer             'wMTotal:月件數總計,wCTotal:年件數總計(累計)
        Dim MonthlyRate, CumulativeRate As String   '月外注率,年外注率
        Dim wTColumn As String = "CreateTime"        '當月委託或當月開發完成的篩選欄位

        Dim WorkTable As String = "Temp_InqManufOutList_" & Request.Cookies("UserID").Value

        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBTable1 As DataTable
        Dim DBDataRow As DataRow

        Dim OleDBCommand1 As New OleDbCommand
        Dim OleDbConnection1 As New OleDbConnection

        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        If DSts.SelectedValue = 1 Then  '判斷是否選擇開發完成
            SQL2 = "and Sts = 1 "
            wTColumn = "CompletedTime"
        End If
        '篩選月份資料
        SQL1 = "SELECT Case " + DSOrB.SelectedValue + " When '' Then '無名稱' Else " + DSOrB.SelectedValue + " End AS SOrB, COUNT(*) AS MonthlyCount FROM F_ManufOutSheet "
        SQL1 = SQL1 + "WHERE (" + wTColumn + " > DATEADD(d, 0, '" + wMDate + "')) AND (" + wTColumn + " < DATEADD(m, 1, '" + wMDate + "')) "
        SQL1 = SQL1 + SQL2
        SQL1 = SQL1 + "GROUP BY " + DSOrB.SelectedValue
        Dim DBAdapter1 As New OleDbDataAdapter(SQL1, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Monthly")
        '月件數總計
        SQL1 = "SELECT COUNT(*) AS MTotal FROM F_ManufOutSheet "
        SQL1 = SQL1 + "WHERE (" + wTColumn + " > DATEADD(d, 0, '" + wMDate + "')) AND (" + wTColumn + " < DATEADD(m, 1, '" + wMDate + "')) "
        SQL1 = SQL1 + SQL2
        Dim DBAdapter2 As New OleDbDataAdapter(SQL1, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "Total")

        wMTotal = DBDataSet1.Tables("Total").Rows(0).Item("MTotal")
        If wMTotal = 0 Then
            DataGrid1.Visible = False
            Exit Sub
        Else
            DataGrid1.Visible = True
        End If
        '篩選年份資料
        SQL1 = "SELECT Case " + DSOrB.SelectedValue + " When '' Then '無名稱' Else " + DSOrB.SelectedValue + " End AS SOrB, COUNT(*) AS CumulativeCount FROM F_ManufOutSheet "
        SQL1 = SQL1 + "WHERE (" + wTColumn + " > DATEADD(d, 0, '" + wCDate + "')) AND (" + wTColumn + " < DATEADD(m, 1, '" + wMDate + "')) "
        SQL1 = SQL1 + SQL2
        SQL1 = SQL1 + "GROUP BY " + DSOrB.SelectedValue
        Dim DBAdapter3 As New OleDbDataAdapter(SQL1, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "Cumulative")
        '年件數總計
        SQL1 = "SELECT COUNT(*) AS CTotal FROM F_ManufOutSheet "
        SQL1 = SQL1 + "WHERE (" + wTColumn + " > DATEADD(d, 0, '" + wCDate + "')) AND (" + wTColumn + " < DATEADD(m, 1, '" + wMDate + "')) "
        SQL1 = SQL1 + SQL2
        Dim DBAdapter4 As New OleDbDataAdapter(SQL1, OleDbConnection1)
        DBAdapter4.Fill(DBDataSet1, "CumulativeTotal")

        wCTotal = DBDataSet1.Tables("CumulativeTotal").Rows(0).Item("CTotal")

        'Call Stored Procedure Create WorkTable
        SQL = "Exec sp_Temp_InqManufOutList '" & WorkTable & "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQL
        OleDBCommand1.ExecuteNonQuery()

        For i = 0 To DBDataSet1.Tables("Cumulative").Rows.Count - 1
            MonthlyCount = 0
            MonthlyRate = "0"
            CumulativeCount = CInt(DBDataSet1.Tables("Cumulative").Rows(i).Item("CumulativeCount"))
            '比月資料是否包含在年資料中,若有的話,則將月件數取出
            For j = 0 To DBDataSet1.Tables("Monthly").Rows.Count - 1
                If DBDataSet1.Tables("Cumulative").Rows(i).Item("SOrB") = DBDataSet1.Tables("Monthly").Rows(j).Item("SOrB") Then
                    MonthlyCount = CInt(DBDataSet1.Tables("Monthly").Rows(j).Item("MonthlyCount"))
                    Exit For
                End If
            Next j
            '計算月外注率
            If MonthlyCount <> 0 Then MonthlyRate = Format(MonthlyCount / wMTotal * 100, "##0.00")
            '計算年外注率
            CumulativeRate = Format(CumulativeCount / wCTotal * 100, "##0.00")

            SQL = "Insert into " & WorkTable & " "
            SQL = SQL + "( "
            SQL = SQL + "SOrB, MonthlyCount, MonthlyRate, CumulativeCount, CumulativeRate, CreateUser, CreateTime "
            SQL = SQL + ")  "
            SQL = SQL + "VALUES( "
            SQL = SQL + " '" + DBDataSet1.Tables("Cumulative").Rows(i).Item("SOrB") + "', "
            SQL = SQL + " '" + CStr(MonthlyCount) + "', "
            SQL = SQL + " '" + CStr(MonthlyRate) + "', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Cumulative").Rows(i).Item("CumulativeCount")) + "', "
            SQL = SQL + " '" + CStr(CumulativeRate) + "', "
            SQL = SQL + " '" + Request.Cookies("UserID").Value + "', "     '作成者
            SQL = SQL + " '" + NowDateTime + "' "       '作成時間
            SQL = SQL + " ) "

            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()
        Next i

        SQL = "SELECT SOrB, MonthlyCount, MonthlyRate, CumulativeCount, CumulativeRate "
        SQL = SQL + "FROM " + WorkTable + " "
        SQL = SQL + "Order by SOrB"

        Dim DBAdapter5 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter5.Fill(DBDataSet2, "Header")
        DBTable1 = DBDataSet2.Tables("Header")

        DBDataRow = DBTable1.NewRow()
        DBDataRow.Item(0) = DYear.SelectedValue + "年" + DMonth.SelectedValue + "月"
        DBTable1.Rows.InsertAt(DBDataRow, 0)    '插入年月在第一列

        DBDataRow = DBTable1.NewRow()
        DBDataRow.Item(0) = "總計"
        DBDataRow.Item(1) = wMTotal
        DBDataRow.Item(2) = "100"
        DBDataRow.Item(3) = wCTotal
        DBDataRow.Item(4) = "100"
        DBTable1.Rows.Add(DBDataRow)            '新增總計資料在最後一列

        DataGrid1.DataSource = DBTable1
        DataGrid1.DataBind()

        DataGrid1.Items(0).Cells(0).ColumnSpan = 5  '將第一列顯示為跨欄
        For i = 1 To 4
            DataGrid1.Items(0).Cells(i).Visible = False '隱藏後面4欄
        Next i

        OleDbConnection1.Close()
    End Sub

    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '程式別不同

        ''pFormNo = DFormName.SelectedValue
        DataGrid1.AllowPaging = False   '程式別不同
        CreateWorkTable()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=Commission_List.xls")     '程式別不同
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        DataGrid1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        DataGrid1.AllowPaging = wAllowPaging        '程式別不同
    End Sub
End Class
