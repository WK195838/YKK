Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class INQ_REQStatus
    Inherits System.Web.UI.Page


    '*****************************************************************
    '**
    '**     全域變數
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
        Dim Sql As String = ""
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        cn.ConnectionString = ConnectString
        Sql = "select "
        'Sql &= "YYMM, "
        'Sql &= "CASE SMPF5C WHEN '1' THEN 'SAMPLE' ELSE 'BULK' END AS BULK, "
        Sql &= "YYMM, CTYC5C AS BULK, "
        'Sql &= "case when CTYC5C='E' OR CTYC5C='K' then 'Y' else '' end as bulk, "
        Sql &= "sum(case when rdlu5m < rdlu5c then 1 else 0 end) as Postpone, "     '延後
        Sql &= "sum(case when rdlu5m > rdlu5c then 1 else 0 end) as Advance, "      '提前      
        Sql &= "sum(case when rdlu5m = rdlu5c then 1 else 0 end) as Same, "
        Sql &= "count(*) as Total, "
        Sql &= "'' as PURL, "
        Sql &= "'' as AURL, "
        Sql &= "'' as SURL, "
        Sql &= "'...' as PSelect, '...' as ASelect, '...' as SSelect, "
        Sql &= "'%' as PPer, '%' as APer, '%' as SPer "
        '
        Sql &= "FROM SONGOLIB.DELSTATUS@ "
        Sql &= "GROUP BY YYMM, CTYC5C "
        Sql &= "ORDER BY YYMM DESC, CTYC5C "
        'Sql &= "WHERE SMPF5C = '' "
        'Sql &= "GROUP BY YYMM, case when CTYC5C='E' OR CTYC5C='K' then 'Y' else '' end "
        'Sql &= "ORDER BY YYMM DESC, case when CTYC5C='E' OR CTYC5C='K' then 'Y' else '' end "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(Sql, cn)
        DBAdapter1.Fill(ds, "S5M00")
        If ds.Tables("S5M00").Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = ds
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If

        cn.Close()

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim gv As GridView = DirectCast(sender, GridView)
            Dim gvRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear()

            Dim tc1_1 As New TableCell()
            tc1_1.Text = "Month"
            tc1_1.Wrap = False
            tc1_1.Width = 50
            tc1_1.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc1_1)

            Dim tc1_11 As New TableCell()
            tc1_11.Text = ""
            tc1_11.Wrap = False
            tc1_11.Width = 50
            tc1_11.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc1_11)

            Dim tc1_2 As New TableCell()
            tc1_2.Text = "後延"
            tc1_2.ColumnSpan = 3
            tc1_2.Wrap = False
            tc1_2.Width = 150
            tc1_2.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc1_2)

            Dim tc1_3 As New TableCell()
            tc1_3.Text = "前倒"
            tc1_3.ColumnSpan = 3
            tc1_3.Wrap = False
            tc1_3.Width = 150
            tc1_3.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc1_3)

            Dim tc1_4 As New TableCell()
            tc1_4.Text = "同樣"
            tc1_4.ColumnSpan = 2
            tc1_4.Wrap = False
            tc1_4.Width = 150
            tc1_4.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc1_4)

            Dim tc1_5 As New TableCell()
            tc1_5.Text = "計"
            tc1_5.Wrap = False
            tc1_5.Width = 150
            tc1_5.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc1_5)

            gv.Controls(0).Controls.AddAt(0, gvRow)
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            For i = 0 To 10
                Select Case e.Row.Cells(i).Text
                    Case "%"
                        e.Row.Cells(i).Text = Math.Round(CDbl(e.Row.Cells(i - 1).Text) / CDbl(e.Row.Cells(10).Text) * 100, 1) & " %"
                    Case "..."
                        Select Case i
                            Case 4
                                Dim h1 As New HyperLink
                                h1.Target = "_blank"
                                h1.Text = e.Row.Cells(i).Text
                                h1.NavigateUrl = "http://10.245.1.6/WorkFlowSub/INQ_REQStatus01.aspx?pMonth=" & e.Row.Cells(0).Text & "&pCTYC=" & e.Row.Cells(1).Text & "&pAction=P"
                                e.Row.Cells(i).Text = ""
                                e.Row.Cells(i).Controls.Add(h1)
                            Case 7
                                Dim h1 As New HyperLink
                                h1.Target = "_blank"
                                h1.Text = e.Row.Cells(i).Text
                                h1.NavigateUrl = "http://10.245.1.6/WorkFlowSub/INQ_REQStatus01.aspx?pMonth=" & e.Row.Cells(0).Text & "&pCTYC=" & e.Row.Cells(1).Text & "&pAction=A"
                                e.Row.Cells(i).Text = ""
                                e.Row.Cells(i).Controls.Add(h1)
                            Case 10
                                'Dim h1 As New HyperLink
                                'h1.Target = "_blank"
                                'h1.Text = e.Row.Cells(i).Text
                                'h1.NavigateUrl = "http://localhost:49391/WorkFlowSub/INQ_REQStatus01.aspx?pMonth=" & e.Row.Cells(0).Text & "&Action=S"
                                'e.Row.Cells(i).Text = ""
                                'e.Row.Cells(i).Controls.Add(h1)
                            Case Else
                        End Select
                    Case Else
                End Select
            Next
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub

End Class
