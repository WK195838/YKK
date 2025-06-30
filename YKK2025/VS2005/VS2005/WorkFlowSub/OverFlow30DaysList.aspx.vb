Imports System.Data
Imports System.Data.OleDb

Partial Class OverFlow30DaysList
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
    Dim pFormNo As String
    Dim wLevel As String = ""
    Dim wDivision As String = ""
    Dim wName As String = ""
    Dim NowDateTime As String       '現在日期時間

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "SPD_Delay30List.aspx"

        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            SetSearchItem()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(Now.Date) + " " + _
                      CStr(Now.Hour) + ":" + _
                      CStr(Now.Minute) + ":" + _
                      CStr(Now.Second)     '現在日時
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub SetSearchItem()
        pFormNo = DFormNo.SelectedValue
        Search_Item_Attribute()
    End Sub

    Sub DataList()
        Dim i As Integer = 0
        Dim wTableName As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        SQL = "SELECT FormNo, FormName, TableName1 FROM M_Form "
        SQL = SQL + " Where Active  =  '1' "
        SQL = SQL + "   And FormNo = '" + DFormNo.SelectedValue + "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Flow")
        If DBDataSet1.Tables("Flow").Rows.Count > 0 Then
            wTableName = "F_" + DBDataSet1.Tables("Flow").Rows(0).Item("TableName1")
        End If
        OleDbConnection1.Close()

        If wTableName <> "" Then
            GridView1.Columns.Item(0).HeaderText = "委託書"
            GridView1.Columns.Item(1).HeaderText = "委託(a)"
            GridView1.Columns.Item(2).HeaderText = "現在(b)"
            GridView1.Columns.Item(3).HeaderText = "休假(c)"
            GridView1.Columns.Item(4).HeaderText = "(b-a)"
            GridView1.Columns.Item(5).HeaderText = "(b-a-c)"
            GridView1.Columns.Item(6).HeaderText = "ＮＯ"
            GridView1.Columns.Item(7).HeaderText = "狀態"
            GridView1.Columns.Item(8).HeaderText = "Buyer"

            SQL = "SELECT "
            SQL = SQL + wTableName + ".CreateTime as Field2, "
            SQL = SQL + "GetDate()  as Field3, "

            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
            'SQL = SQL + "(select count(*) from m_vacation where depo='cl' and ymd>= " + wTableName + ".CreateTime and ymd<=Getdate()) as Field4, "
            '
            SQL = SQL + "(select count(*) from m_vacation where depo='CL1' and ymd>= " + wTableName + ".CreateTime and ymd<=Getdate()) as Field4, "
            'Modify-End

            SQL = SQL + "DateDiff(day, " + wTableName + ".CreateTime, getdate()) as Field5, "

            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
            'SQL = SQL + "DateDiff(day, " + wTableName + ".CreateTime, getdate())-(select count(*) from m_vacation where depo='cl' and ymd>= " + wTableName + ".CreateTime and ymd<=Getdate()) As Field6, "
            '
            SQL = SQL + "DateDiff(day, " + wTableName + ".CreateTime, getdate())-(select count(*) from m_vacation where depo='CL1' and ymd>= " + wTableName + ".CreateTime and ymd<=Getdate()) As Field6, "
            'Modify-End

            SQL = SQL + wTableName + ".No as Field7, "
            SQL = SQL + "'開發中' as Field8, "
            SQL = SQL + wTableName + ".Buyer as Field9, "

            If pFormNo = "000001" Then
                GridView1.Columns.Item(9).HeaderText = "Size/ChainType/Class"
                SQL = SQL + "'圖面' as Field1, "
                SQL = SQL + wTableName + ".Spec as Field10 "
            End If
            If pFormNo = "000002" Then
                GridView1.Columns.Item(9).HeaderText = ""
                SQL = SQL + "'修圖' as Field1, "
                SQL = SQL + "'' as Field10 "
            End If
            If pFormNo = "000003" Then
                GridView1.Columns.Item(9).HeaderText = "Slider"
                SQL = SQL + "'內製' as Field1, "
                SQL = SQL + "SliderCode as Field10 "
            End If
            If pFormNo = "000007" Then
                GridView1.Columns.Item(9).HeaderText = "Slider"
                SQL = SQL + "'外注' as Field1, "
                SQL = SQL + "SliderCode as Field10 "
            End If
            If pFormNo = "000014" Then
                GridView1.Columns.Item(9).HeaderText = "Size/ChainType/Class"
                SQL = SQL + "'表面處理' as Field1, "
                SQL = SQL + "Spec as Field10 "
            End If
            SQL = SQL + "From " + wTableName + " "
            SQL = SQL + "Where Sts = '0' "
            If DVType.SelectedValue = "1" Then
                '不含假日
                'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                'SQL = SQL + "  And DateDiff(day, " + wTableName + ".CreateTime, getdate())-(select count(*) from m_vacation where depo='cl' and ymd>= " + wTableName + ".CreateTime and ymd<=Getdate()) > '" + DDays.SelectedValue + "' "
                '
                SQL = SQL + "  And DateDiff(day, " + wTableName + ".CreateTime, getdate())-(select count(*) from m_vacation where depo='CL1' and ymd>= " + wTableName + ".CreateTime and ymd<=Getdate()) > '" + DDays.SelectedValue + "' "
                'Modify-End
                SQL = SQL + "order by Field6 Desc "
            Else
                '含假日
                SQL = SQL + "  And DateDiff(day, " + wTableName + ".CreateTime, getdate()) > '" + DDays.SelectedValue + "' "
                SQL = SQL + "order by Field5 Desc "
            End If

            OleDbConnection1.Open()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "DelayList")
            GridView1.DataSource = DBDataSet2
            GridView1.DataBind()
            OleDbConnection1.Close()
        End If

    End Sub

    Sub Search_Item_Attribute()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        pFormNo = DFormNo.SelectedValue
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

        pFormNo = DFormNo.SelectedValue
        GridView1.AllowPaging = False   '程式別不同
        DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=Overflow30DaysList.xls")     '程式別不同
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

End Class
