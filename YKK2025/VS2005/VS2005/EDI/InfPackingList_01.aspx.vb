Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfPackingList_01
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim wUserID As String
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim uEDIMapping As New EDI2011.MappingService
    Dim uEDICommon As New EDI2011.CommonService
    Dim uWFSCommon As New WFS.CommonService
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        SetPopupFunction()                          '設定彈出視窗事件
        If Not IsPostBack Then                      'PostBack
            ShowData()
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
        Response.Cookies("PGM").Value = "InfPackingList.aspx"       '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyy/MM/dd")   '現在日期時間
        DBuyer.Text = Request.QueryString("pBuyer")
        DPO.Text = Request.QueryString("pPO")
        DOrderNo.Text = Request.QueryString("pOrderNo")
        wUserID = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        BRefresh.Attributes("onclick") = "var ok=window.confirm('" + "是否要更新 [Packing List] 資料？" + "');if(!ok){return false;}"
        DBuyer.ReadOnly = True
        DPO.ReadOnly = True
        DOrderNo.ReadOnly = True
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(更新)
    '**     Refresh Button
    '**
    '*****************************************************************
    Protected Sub BRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BRefresh.Click
        '
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        uEDICommon.UpdatePackingList(LogID, DBuyer.Text, DOrderNo.Text, wUserID)

        ShowData()

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim sql As String
        Dim xCaseNo, i As Integer
        Dim new_row As DataRow
        '
        '最後更新日
        sql = "SELECT CreateTime From B_PackingInstruction "
        sql &= "Where CustomerBuyer = '" & DBuyer.Text & "' "
        sql &= "  And PO = '" & DPO.Text & "' "
        sql &= "  And OrderNo = '" & DOrderNo.Text & "' "
        sql &= "Order by CreateTime Desc "
        Dim dt_OrderProgress As DataTable = uDataBase.GetDataTable(sql)
        If dt_OrderProgress.Rows.Count > 0 Then
            DLastUpdate.Text = "最後更新 : " + CDate(dt_OrderProgress.Rows(0).Item("CreateTime")).ToString("yyyy/MM/dd HH:mm:ss")
        End If
        '
        '篩選資料
        sql = "SELECT "
        sql &= "PO, Seqno, "
        sql &= "Case Delivery When '1' Then '' Else '★' End As Delivery, "
        sql &= "OrderNo, OrderSubNo, Item, "
        sql &= "rtrim(ltrim(ItemName1)) + ' ' + rtrim(ltrim(ItemName2)) + ' ' + rtrim(ltrim(ItemName3)) as ItemName, "
        sql &= "CaseNoStart as CaseNo, "
        sql &= "Length, Color, PackQty, Count, ItemNet, OuterNet, Gross, "
        sql &= "OItem, "
        sql &= "rtrim(ltrim(OItemName1)) + ' ' + rtrim(ltrim(OItemName2)) + ' ' + rtrim(ltrim(OItemName3)) as OItemName, "
        sql &= "'0' as EndFlag "
        '
        sql &= "From B_PackingInstruction "
        sql &= "Where CustomerBuyer = '" & DBuyer.Text & "' "
        sql &= "  And PO = '" & DPO.Text & "' "
        sql &= "  And OrderNo = '" & DOrderNo.Text & "' "
        sql &= "Order by CaseNoStart, CaseNoEnd, OrderNo, OrderSubNo "
        '
        Dim dtB2BData As DataTable = uDataBase.GetDataTable(sql)
        If dtB2BData.Rows.Count > 0 Then
            '
            xCaseNo = dtB2BData.Rows(0).Item("CaseNo")
            For i = 0 To dtB2BData.Rows.Count - 1
                If xCaseNo <> dtB2BData.Rows(i).Item("CaseNo") Then

                    '-- 手動新增一行 DataRow
                    new_row = dtB2BData.NewRow()
                    new_row("PO") = dtB2BData.Rows(i - 1)(0)
                    new_row("SeqNo") = 999999
                    new_row("Delivery") = ""
                    new_row("OrderNo") = dtB2BData.Rows(i - 1)(3)
                    new_row("OrderSubNo") = 999999
                    new_row("Item") = dtB2BData.Rows(i - 1)(15)
                    new_row("ItemName") = dtB2BData.Rows(i - 1)(16)
                    new_row("CaseNo") = dtB2BData.Rows(i - 1)(7)
                    new_row("Length") = 999999
                    new_row("Color") = ""
                    new_row("PackQty") = 999999
                    new_row("Count") = 999999
                    new_row("ItemNet") = dtB2BData.Rows(i - 1)(12)
                    new_row("OuterNet") = dtB2BData.Rows(i - 1)(13)
                    new_row("Gross") = dtB2BData.Rows(i - 1)(14)
                    new_row("OItem") = dtB2BData.Rows(i - 1)(15)
                    new_row("OItemName") = dtB2BData.Rows(i - 1)(16)
                    new_row("EndFlag") = "1"
                    dtB2BData.Rows.Add(new_row)

                    xCaseNo = dtB2BData.Rows(i).Item("CaseNo")
                End If
            Next
            '
            '-- 手動新增一行 DataRow
            new_row = dtB2BData.NewRow()
            new_row("PO") = dtB2BData.Rows(i - 1)(0)
            new_row("SeqNo") = 999999
            new_row("Delivery") = ""
            new_row("OrderNo") = dtB2BData.Rows(i - 1)(3)
            new_row("OrderSubNo") = 999999
            new_row("Item") = dtB2BData.Rows(i - 1)(15)
            new_row("ItemName") = dtB2BData.Rows(i - 1)(16)
            new_row("CaseNo") = dtB2BData.Rows(i - 1)(7)
            new_row("Length") = 999999
            new_row("Color") = ""
            new_row("PackQty") = 999999
            new_row("Count") = 999999
            new_row("ItemNet") = dtB2BData.Rows(i - 1)(12)
            new_row("OuterNet") = dtB2BData.Rows(i - 1)(13)
            new_row("Gross") = dtB2BData.Rows(i - 1)(14)
            new_row("OItem") = dtB2BData.Rows(i - 1)(15)
            new_row("OItemName") = dtB2BData.Rows(i - 1)(16)
            new_row("EndFlag") = "1"
            dtB2BData.Rows.Add(new_row)
            '
            dtB2BData.DefaultView.Sort = "CaseNo ASC"
            '
            GridView1.DataSource = dtB2BData
            GridView1.DataBind()
        End If

    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        '
        Response.AppendHeader("Content-Disposition", "attachment;filename=PackingList_01.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        '
        ShowData()
        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()
    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Protected Sub GridView1_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PreRender
        Dim i As Integer = 1
        Dim MergeIdx As Integer = 0
        Dim mySingleRow As GridViewRow
        Dim mergeColumns(9) As Integer
        mergeColumns(0) = 0
        mergeColumns(1) = 3
        mergeColumns(2) = 5
        mergeColumns(3) = 6
        mergeColumns(4) = 7
        mergeColumns(5) = 9
        mergeColumns(6) = 12
        mergeColumns(7) = 13
        mergeColumns(8) = 14

        '合併儲存格
        For MergeIdx = 0 To 9
            i = 1
            For Each mySingleRow In GridView1.Rows
                If CInt(mySingleRow.RowIndex) > 0 Then  '資料筆數>0

                    '一般合併
                    If mySingleRow.Cells(mergeColumns(MergeIdx)).Text.Trim() = GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(MergeIdx)).Text.Trim() And _
                       (mySingleRow.Cells(7).Text.Trim() = GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(7).Text.Trim() Or mySingleRow.Cells(7).Text.Trim() = "") Then

                        If mySingleRow.Cells(mergeColumns(MergeIdx)).Text.Trim() <> "&nbsp;" Then   '空白是否

                            '各欄位合併處理
                            GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(MergeIdx)).RowSpan += 1
                            mySingleRow.Cells(mergeColumns(MergeIdx)).Visible = False
                            i = i + 1
                        End If   '空白是否
                    Else
                        GridView1.Rows(CInt(mySingleRow.RowIndex)).Cells(mergeColumns(MergeIdx)).RowSpan = 1
                        i = 1
                    End If

                Else  '資料筆數>0
                    mySingleRow.Cells(mergeColumns(MergeIdx)).RowSpan = 1
                End If
            Next
            '
        Next

        '設定背景顏色
        Dim xColor As Color = Drawing.Color.LightBlue
        For Each mySingleRow In GridView1.Rows
            If GridView1.Rows(CInt(mySingleRow.RowIndex)).Cells(15).Text.Trim() = "1" Then
                For i = 0 To 14
                    GridView1.Rows(CInt(mySingleRow.RowIndex)).Cells(i).BackColor = xColor
                Next

                If xColor = Drawing.Color.LightBlue Then
                    xColor = Drawing.Color.YellowGreen
                Else
                    xColor = Drawing.Color.LightBlue
                End If
            Else
                For i = 0 To 14
                    GridView1.Rows(CInt(mySingleRow.RowIndex)).Cells(i).BackColor = xColor
                Next
            End If
        Next
    End Sub
    '
    '*****************************************************************
    '**
    '**     GridView資料重整
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
            e.Row.Cells(15).Visible = False
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            If e.Row.Cells(15).Text = "1" Then
                e.Row.Cells(1).Text = ""
                e.Row.Cells(4).Text = "外箱"
                e.Row.Cells(7).Text = ""
                e.Row.Cells(8).Text = ""
                e.Row.Cells(9).Text = ""
                e.Row.Cells(10).Text = ""
                e.Row.Cells(11).Text = ""
            End If
            e.Row.Cells(15).Visible = False
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
        '
        For i As Integer = 0 To 14
            e.Row.Cells(i).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        Next
    End Sub

End Class
