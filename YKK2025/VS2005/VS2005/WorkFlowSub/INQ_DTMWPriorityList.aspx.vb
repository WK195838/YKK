Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class INQ_DTMWPriorityList
    Inherits System.Web.UI.Page


    '*****************************************************************
    '**
    '**     全域變數  
    '**     CALL WFS890R1    SQLRPG      ORDER REAL S8G00 JOY
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間

    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            SetParameter()                              '設定參數
            GetData()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '-----------------------------------------------------------------
        '-- 系統參數
        '-----------------------------------------------------------------
        Server.ScriptTimeout = 900                                  '設定逾時時間
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        GridView1.Visible = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetData)
    '**     取得資料
    '**
    '*****************************************************************
    Sub GetData()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '
        SQL = "SELECT "
        SQL = SQL & "[NAME], pNO, Priority, YKKColorCode, '' AS OrderSts, "
        SQL = SQL & "FormName, convert(varchar, DATE, 111) AS DATE, "
        SQL = SQL & "customercode + '(' + CUSTOMER + ')' AS CUSTOMERD, "
        SQL = SQL & "BUYERcode + '(' + BUYER + ')' AS BUYERD, "
        SQL = SQL & "Sts, WFUrl, ViewUrl, "
        SQL = SQL & "'....' as WFUrlDesc, '....' as ViewUrlDesc "
        '
        SQL = SQL & "FROM V_DTMWPriorityList "
        SQL = SQL & "Order By [NAME], PNO "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "ITEM")
        If DBDataSet1.Tables("ITEM").Rows.Count > 0 Then
            GridView1.DataSource = DBDataSet1
            GridView1.DataBind()
            GridView1.Visible = True
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
        End If
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
            ' 表頭-8行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "Name"
            H4row.Cells.Add(H4tc0)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = "No."
            H4row.Cells.Add(H4tc2)

            Dim H4tc3 As TableCell = New TableCell
            H4tc3.Text = "優先度"
            H4row.Cells.Add(H4tc3)

            Dim H4tc4 As TableCell = New TableCell
            H4tc4.Text = "Color"
            H4tc4.BackColor = Color.Green
            H4row.Cells.Add(H4tc4)

            Dim H4tc5 As TableCell = New TableCell
            H4tc5.Text = "Order Status"
            H4tc5.BackColor = Color.Blue
            H4row.Cells.Add(H4tc5)

            Dim H4tc5a As TableCell = New TableCell
            H4tc5a.Text = "DTM Status"
            H4tc5a.BackColor = Color.Orange

            H4row.Cells.Add(H4tc5a)

            Dim H4tc5b As TableCell = New TableCell
            H4tc5b.Text = "Link(工程)"
            H4tc5b.BackColor = Color.Tomato
            H4row.Cells.Add(H4tc5b)

            Dim H4tc5c As TableCell = New TableCell
            H4tc5c.Text = "Link(申請)"
            H4tc5c.BackColor = Color.Tomato
            H4row.Cells.Add(H4tc5c)

            Dim H4tc6 As TableCell = New TableCell
            H4tc6.Text = "申請日"
            H4row.Cells.Add(H4tc6)

            Dim H4tc7 As TableCell = New TableCell
            H4tc7.Text = "客戶"
            H4row.Cells.Add(H4tc7)

            Dim H4tc8 As TableCell = New TableCell
            H4tc8.Text = "Buyer"
            H4row.Cells.Add(H4tc8)

            gv.Controls(0).Controls.AddAt(0, H4row)
        End If
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim SQL As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
        '
        cn.ConnectionString = ConnectString

        If (e.Row.RowType = DataControlRowType.DataRow) Then
            '
            If e.Row.Cells(3).Text <> "&nbsp;" Then
                SQL = "SELECT "
                SQL = SQL & "A.ORDN5E AS ORNO, 'CONFIRM=' || B.OCNU5C || ')' AS CFDATE, B.OCNU5C as CDate, "
                SQL = SQL & "'[' || A.ITMC5E || ']' || C.IT1IA0 AS ITEM1, "
                SQL = SQL & "C.IT2IA0 AS ITEM2 "
                '
                SQL = SQL & "FROM WAVEDLIB.S5E00 A, WAVEDLIB.S5C00 B, WAVEDLIB.FA000 C "
                SQL = SQL + "WHERE A.ORDN5E = B.ORDN5C AND A.ITMC5E = C.ITMCA0 "
                SQL = SQL + "AND B.CTYC5C <> '' "
                SQL = SQL + "AND C.ST1CA0 || ST4CA0 IN ('11', '12') "
                SQL = SQL + "AND ( "
                SQL = SQL + "      C.SF1CA0='GREEN-F' OR C.SF2CA0='GREEN-F' OR C.SF3CA0='GREEN-F' OR C.SF4CA0='GREEN-F' OR C.SF5CA0='GREEN-F' OR C.SF6CA0='GREEN-F' "
                SQL = SQL + ") "
                SQL = SQL + "AND A.CLRC5E = '" & e.Row.Cells(3).Text & "' "
                SQL = SQL + "ORDER BY A.RADU5E DESC "
                SQL = SQL + "FETCH FIRST 1 ROW ONLY "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, cn)
                DBAdapter1.Fill(ds, "ORDER")
                If ds.Tables("ORDER").Rows.Count > 0 Then

                    If ds.Tables("ORDER").Rows(0).Item("CDATE") > 0 Then
                        e.Row.Cells(4).Text = ds.Tables("ORDER").Rows(0).Item("ORNO") & _
                                             "<BR>" & ds.Tables("ORDER").Rows(0).Item("CFDATE") & _
                                             "<BR>" & ds.Tables("ORDER").Rows(0).Item("ITEM1") & _
                                             "<BR>" & ds.Tables("ORDER").Rows(0).Item("ITEM2")

                        e.Row.Cells(4).ForeColor = Color.Blue
                    Else
                        e.Row.Cells(4).Text = ds.Tables("ORDER").Rows(0).Item("ORNO") & _
                                             "<BR>" & ds.Tables("ORDER").Rows(0).Item("CFDATE") & _
                                             "<BR>" & ds.Tables("ORDER").Rows(0).Item("ITEM1") & _
                                             "<BR>" & ds.Tables("ORDER").Rows(0).Item("ITEM2")

                        e.Row.Cells(4).ForeColor = Color.Blue
                        '
                        e.Row.Cells(1).Attributes.Add("style", "border:1px solid red ")
                        e.Row.Cells(1).BackColor = Color.Pink
                        e.Row.Cells(2).Attributes.Add("style", "border:1px solid red ")
                        e.Row.Cells(2).BackColor = Color.Pink
                        e.Row.Cells(3).Attributes.Add("style", "border:1px solid red ")
                        e.Row.Cells(3).BackColor = Color.Pink
                        e.Row.Cells(4).Attributes.Add("style", "border:1px solid red ")
                        e.Row.Cells(4).BackColor = Color.Pink
                        e.Row.Cells(5).Attributes.Add("style", "border:1px solid red ")
                        e.Row.Cells(5).BackColor = Color.Pink
                    End If
                Else
                    e.Row.Cells(1).Attributes.Add("style", "border:1px solid red ")
                    e.Row.Cells(1).BackColor = Color.Pink
                    e.Row.Cells(2).Attributes.Add("style", "border:1px solid red ")
                    e.Row.Cells(2).BackColor = Color.Pink
                    e.Row.Cells(3).Attributes.Add("style", "border:1px solid red ")
                    e.Row.Cells(3).BackColor = Color.Pink
                    e.Row.Cells(4).Attributes.Add("style", "border:1px solid red ")
                    e.Row.Cells(4).BackColor = Color.Pink
                    e.Row.Cells(5).Attributes.Add("style", "border:1px solid red ")
                    e.Row.Cells(5).BackColor = Color.Pink
                End If

            End If
        End If
        '
        cn.Close()
    End Sub

End Class
