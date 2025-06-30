Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfOrderProgress
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim wUserID As String           'UserID
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
            '
            Dim sql As String
            sql = "Select Top 1 RegisterTime "
            Sql &= "From B_OrderProgress "
            sql &= "Where CustomerBuyer = '" & DBuyer.Text & "' "
            sql &= "Order by RegisterTime Desc "
            Dim dt_RegisterTime As DataTable = uDataBase.GetDataTable(sql)
            If dt_RegisterTime.Rows.Count > 0 Then
                DInputDate.Text = CDate(dt_RegisterTime.Rows(0).Item("RegisterTime")).ToString("yyyyMMdd")
            Else
                DInputDate.Text = Now.ToString("yyyyMMdd")
            End If
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
        Response.Cookies("PGM").Value = "InfOrderProgress.aspx"     '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyy/MM/dd")   '現在日期時間
        'DBuyer.Text = Request.QueryString("pBuyer")
        wUserID = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        BRefresh.Attributes("onclick") = "var ok=window.confirm('" + "是否要更新 [Order Progress] 資料？" + "');if(!ok){return false;}"
        BSelectPO.Attributes("onclick") = "window.open('OpenPOPicker.aspx?pBuyer=" & DBuyer.Text & "&pFun=PROGRESS" & "','POPicker','status=0,toolbar=0,width=300,height=500,resizable=yes,scrollbars=yes');"    '選取PO資料
        BRefresh.Style("left") = -500 & "px"
        LCustomerOrder.Style("left") = -500 & "px"
        LCustomerOrder.NavigateUrl = "InfCustomerOrder.aspx?pUserID=" + wUserID + "&pBuyer=" + DBuyer.Text + "&pPO=" + DPO.Text + "&pDate=" + DInputDate.Text
        'DBuyer.ReadOnly = True
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Go)
    '**     Go Button
    '**
    '*****************************************************************
    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        '
        ShowData()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(更新)
    '**     Refresh Button
    '**
    '*****************************************************************
    Protected Sub BRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BRefresh.Click
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")
        Dim sql As String
        '
        sql = "SELECT * From B_OrderProgress "
        sql &= "Where CustomerBuyer = '" & DBuyer.Text & "' "
        sql &= "  And PO = '" & DPO.Text & "' "
        sql &= "Order by PO, SeqNo "
        Dim dt_OrderProgress As DataTable = uDataBase.GetDataTable(sql)
        If dt_OrderProgress.Rows.Count > 0 Then
            uEDICommon.UpdateOrderProgress(LogID, DBuyer.Text, wUserID, DPO.Text)
            '
            ShowData()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            BRefresh.Style("left") = -500 & "px"
            LCustomerOrder.Style("left") = -500 & "px"
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim sql As String
        Dim xStartTime As String = Mid(DInputDate.Text, 1, 4) + "/" + Mid(DInputDate.Text, 5, 2) + "/" + Mid(DInputDate.Text, 7, 2) + " 00:00:00"
        Dim xEndTime As String = Mid(DInputDate.Text, 1, 4) + "/" + Mid(DInputDate.Text, 5, 2) + "/" + Mid(DInputDate.Text, 7, 2) + " 23:59:59"
        '
        '最後更新日
        sql = "SELECT CreateTime From B_OrderProgress "
        sql &= "Where CustomerBuyer = '" & DBuyer.Text & "' "
        If DPO.Text <> "" Then
            sql &= "  And PO = '" & DPO.Text & "' "
        Else
            If DInputDate.Text <> "" Then
                sql &= "  And RegisterTime >= '" & xStartTime & "' "
                sql &= "  And RegisterTime <= '" & xEndTime & "' "
            Else
                sql &= " And PO = '" & DPO.Text & "' "
            End If
        End If
        sql &= "Order by CreateTime Desc "
        Dim dt_OrderProgress As DataTable = uDataBase.GetDataTable(sql)
        If dt_OrderProgress.Rows.Count > 0 Then
            DLastUpdate.Text = "最後更新 : " + CDate(dt_OrderProgress.Rows(0).Item("CreateTime")).ToString("yyyy/MM/dd HH:mm:ss")
        End If
        '
        '篩選資料
        sql = "SELECT "
        sql &= "PO, Seqno, OrderNo, OrderSubNo, Item, Length, Unit, Color, PriceType + PriceVersion As PriceVersion, "
        sql &= "Status, SalesPrice, SalesAmount, "
        sql &= "OrderQty, PackQty, DeliveryQty, InvoiceQty, "

        sql &= "Case OrderDate "
        sql &= "When '1900/1/1' Then '' "
        sql &= "Else Convert(VARCHAR(10), OrderDate, 111) "
        sql &= "End As OrderDate, "

        sql &= "Case ReqDate "
        sql &= "When '1900/1/1' Then '' "
        sql &= "Else Convert(VARCHAR(10), ReqDate, 111) "
        sql &= "End As ReqDate, "

        sql &= "Case PlanDate "
        sql &= "When '1900/1/1' Then '' "
        sql &= "Else Convert(VARCHAR(10), PlanDate, 111) "
        sql &= "End As PlanDate, "

        sql &= "Case CompletedDate "
        sql &= "When '1900/1/1' Then '' "
        sql &= "Else Convert(VARCHAR(10), CompletedDate, 111) "
        sql &= "End As CompletedDate, "

        sql &= "ItemName1 + ' ' + ItemName2 + ' ' + ItemName3 as ItemName, ItemName1 "

        sql &= "From B_OrderProgress "
        sql &= "Where CustomerBuyer = '" & DBuyer.Text & "' "

        If DPO.Text <> "" Then
            sql &= "  And PO = '" & DPO.Text & "' "
        Else
            If DInputDate.Text <> "" Then
                sql &= "  And RegisterTime >= '" & xStartTime & "' "
                sql &= "  And RegisterTime <= '" & xEndTime & "' "
            Else
                sql &= " And PO = '" & DPO.Text & "' "
            End If
        End If

        sql &= "Order by PO, SeqNo "
        '
        Dim dtB2BData As DataTable = uDataBase.GetDataTable(sql)
        If dtB2BData.Rows.Count > 0 Then
            GridView1.DataSource = dtB2BData
            GridView1.DataBind()
            '
            BRefresh.Style("left") = 471 & "px"
            LCustomerOrder.Style("left") = 10 & "px"
            LCustomerOrder.NavigateUrl = "InfCustomerOrder.aspx?pUserID=" + wUserID + "&pBuyer=" + DBuyer.Text + "&pPO=" + DPO.Text + "&pDate=" + DInputDate.Text
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            BRefresh.Style("left") = -500 & "px"
            LCustomerOrder.Style("left") = -500 & "px"
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
        Response.AppendHeader("Content-Disposition", "attachment;filename=OrderProgress.xls")     '程式別不同
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
    '
    '*****************************************************************
    '**
    '**     GridView資料重整
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        GetMyMultiRow(e)
        '
        For i As Integer = 0 To 14
            e.Row.Cells(i).Attributes.Add("class", "text")
        Next
    End Sub
    '
    Protected Sub GetMyMultiRow(ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        '
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim C_Blank As TableCell = New TableCell
            Dim C_Progress As TableCell = New TableCell
            Dim C_PlanDate As TableCell = New TableCell
            Dim C_PackQty As TableCell = New TableCell
            Dim C_DeliveryQty As TableCell = New TableCell
            Dim C_InvoiceQty As TableCell = New TableCell
            Dim C_ItemName As TableCell = New TableCell
            '
            C_Blank.BackColor = Color.Silver
            C_Blank.Text = ""
            row.Cells.Add(C_Blank)

            C_Progress.Font.Bold = True
            C_Progress.BackColor = Color.Silver
            C_Progress.HorizontalAlign = HorizontalAlign.Center
            C_Progress.ColumnSpan = 6
            C_Progress.Text = "Progress"
            row.Cells.Add(C_Progress)

            C_PlanDate.Font.Bold = True
            C_PlanDate.BackColor = Color.Silver
            C_PlanDate.HorizontalAlign = HorizontalAlign.Center
            C_PlanDate.Text = "Completed Date"
            row.Cells.Add(C_PlanDate)

            C_PackQty.Font.Bold = True
            C_PackQty.BackColor = Color.Silver
            C_PackQty.HorizontalAlign = HorizontalAlign.Center
            C_PackQty.Text = "Pack Qty"
            row.Cells.Add(C_PackQty)

            C_DeliveryQty.Font.Bold = True
            C_DeliveryQty.BackColor = Color.Silver
            C_DeliveryQty.HorizontalAlign = HorizontalAlign.Center
            C_DeliveryQty.Text = "Delivery Qty"
            row.Cells.Add(C_DeliveryQty)

            C_InvoiceQty.Font.Bold = True
            C_InvoiceQty.BackColor = Color.Silver
            C_InvoiceQty.HorizontalAlign = HorizontalAlign.Center
            C_InvoiceQty.Text = "Invoice Qty"
            row.Cells.Add(C_InvoiceQty)

            C_ItemName.Font.Bold = True
            C_ItemName.BackColor = Color.Silver
            C_ItemName.HorizontalAlign = HorizontalAlign.Center
            C_ItemName.ColumnSpan = 4
            C_ItemName.Text = "Item Name"
            row.Cells.Add(C_ItemName)

            e.Row.Parent.Controls.Add(row)
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            Dim row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim C_Blank As TableCell = New TableCell
            Dim C_Progress As TableCell = New TableCell
            Dim C_PlanDate As TableCell = New TableCell
            Dim C_PackQty As TableCell = New TableCell
            Dim C_DeliveryQty As TableCell = New TableCell
            Dim C_InvoiceQty As TableCell = New TableCell
            Dim C_ItemName As TableCell = New TableCell
            '
            'OK          	Color.Khaki
            '> PlanData 	Color.LightPink
            '> RequestData	Color.GreenYellow
            Dim xColor As System.Drawing.Color = Color.White
            Dim xDelay As String = ""
            '
            If DataBinder.Eval(e.Row.DataItem, "CompletedDate") = "" Then
                If NowDate > DataBinder.Eval(e.Row.DataItem, "ReqDate") And NowDate > DataBinder.Eval(e.Row.DataItem, "PlanDate") Then
                    xColor = Color.LightPink
                    xDelay = "★"
                Else
                    If NowDate > DataBinder.Eval(e.Row.DataItem, "ReqDate") Then
                        xColor = Color.LightPink
                        xDelay = "●"
                    Else
                        If NowDate > DataBinder.Eval(e.Row.DataItem, "PlanDate") Then
                            xColor = Color.LightPink
                            xDelay = "★"
                        End If
                    End If
                End If
            Else
                If DataBinder.Eval(e.Row.DataItem, "CompletedDate") > DataBinder.Eval(e.Row.DataItem, "ReqDate") And _
                   DataBinder.Eval(e.Row.DataItem, "CompletedDate") > DataBinder.Eval(e.Row.DataItem, "PlanDate") Then
                    xColor = Color.LightPink
                    xDelay = "★"
                Else
                    If DataBinder.Eval(e.Row.DataItem, "CompletedDate") > DataBinder.Eval(e.Row.DataItem, "ReqDate") Then
                        xColor = Color.LightPink
                        xDelay = "●"
                    Else
                        If DataBinder.Eval(e.Row.DataItem, "CompletedDate") > DataBinder.Eval(e.Row.DataItem, "PlanDate") Then
                            xColor = Color.LightPink
                            xDelay = "★"
                        End If
                    End If
                End If
            End If
            '
            e.Row.Cells(0).BackColor = xColor
            e.Row.Cells(1).BackColor = xColor
            e.Row.Cells(2).BackColor = xColor
            e.Row.Cells(3).BackColor = xColor
            e.Row.Cells(4).BackColor = xColor
            e.Row.Cells(5).BackColor = xColor
            e.Row.Cells(6).BackColor = xColor
            e.Row.Cells(7).BackColor = xColor
            e.Row.Cells(8).BackColor = xColor
            e.Row.Cells(9).BackColor = xColor
            '
            e.Row.Cells(10).BackColor = xColor
            e.Row.Cells(11).BackColor = xColor
            e.Row.Cells(12).BackColor = xColor
            e.Row.Cells(13).BackColor = xColor
            e.Row.Cells(14).BackColor = xColor
            '
            C_Blank.BackColor = xColor
            C_Blank.ForeColor = Color.Red
            C_Blank.Text = xDelay
            row.Cells.Add(C_Blank)

            C_Progress.BackColor = xColor
            C_Progress.ColumnSpan = 6
            C_Progress.Text = DataBinder.Eval(e.Row.DataItem, "Status")
            row.Cells.Add(C_Progress)

            C_PlanDate.BackColor = xColor
            C_PlanDate.Text = DataBinder.Eval(e.Row.DataItem, "CompletedDate")
            row.Cells.Add(C_PlanDate)

            C_PackQty.BackColor = xColor
            If DataBinder.Eval(e.Row.DataItem, "PackQty") = 0 Then
                C_PackQty.Text = ""
            Else
                C_PackQty.Text = DataBinder.Eval(e.Row.DataItem, "PackQty")
            End If
            row.Cells.Add(C_PackQty)

            C_DeliveryQty.BackColor = xColor
            If DataBinder.Eval(e.Row.DataItem, "DeliveryQty") = 0 Then
                C_DeliveryQty.Text = ""
            Else
                C_DeliveryQty.Text = DataBinder.Eval(e.Row.DataItem, "DeliveryQty")
            End If
            row.Cells.Add(C_DeliveryQty)

            C_InvoiceQty.BackColor = xColor
            If DataBinder.Eval(e.Row.DataItem, "InvoiceQty") = 0 Then
                C_InvoiceQty.Text = ""
            Else
                C_InvoiceQty.Text = DataBinder.Eval(e.Row.DataItem, "InvoiceQty")
            End If
            row.Cells.Add(C_InvoiceQty)

            C_ItemName.BackColor = xColor
            C_ItemName.ColumnSpan = 4
            C_ItemName.Text = DataBinder.Eval(e.Row.DataItem, "ItemName1")
            row.Cells.Add(C_ItemName)

            e.Row.Parent.Controls.Add(row)
        End If
    End Sub

End Class
