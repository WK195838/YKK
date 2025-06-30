Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfPackingList
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

            Dim sql As String
            sql = "Select Top 1 RegisterTime "
            sql &= "From B_PackingInstruction "
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
        Response.Cookies("PGM").Value = "InfPackingList.aspx"       '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyy/MM/dd")   '現在日期時間
        DBuyer.Text = Request.QueryString("pBuyer")
        wUserID = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        BSelectPO.Attributes("onclick") = "window.open('OpenPOPicker.aspx?pBuyer=" & DBuyer.Text & "&pFun=PACKING" & "','POPicker','status=0,toolbar=0,width=300,height=500,resizable=yes,scrollbars=yes');"    '選取PO資料
        DBuyer.ReadOnly = True
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
        sql = "SELECT CreateTime From B_PackingInstruction "
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
        sql &= "PO, Seqno, "
        sql &= "Case Delivery When '1' Then '' Else '★' End As Delivery, "
        sql &= "OrderNo, OrderSubNo, Item, "
        sql &= "rtrim(ltrim(ItemName1)) + ' ' + rtrim(ltrim(ItemName2)) + ' ' + rtrim(ltrim(ItemName3)) as ItemName, "
        sql &= "rtrim(ltrim(str(CaseNoStart))) + ' ~ ' + rtrim(ltrim(str(CaseNoEnd))) as CaseNo, "
        sql &= "Length, Color, PackQty, Count, ItemNet, OuterNet, Gross, "
        '
        sql = sql + "'InfPackingList_01.aspx?' + "
        sql = sql + "'pBuyer='   + CustomerBuyer + "
        sql = sql + "'&pPO=' + PO + "
        sql = sql + "'&pOrderNo=' + OrderNo + "
        sql = sql + "'&pUserID=' + '" & wUserID & "' "
        sql = sql + " As OrderURL "
        '
        sql &= "From B_PackingInstruction "
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
        sql &= "Order by PO, SeqNo, CaseNoStart, CaseNoEnd "
        '
        Dim dtB2BData As DataTable = uDataBase.GetDataTable(sql)
        If dtB2BData.Rows.Count > 0 Then
            GridView1.DataSource = dtB2BData
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
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
        Response.AppendHeader("Content-Disposition", "attachment;filename=PackingList.xls")     '程式別不同
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
        Dim mergeColumns(6) As Integer
        mergeColumns(0) = 0
        mergeColumns(1) = 3
        mergeColumns(2) = 5
        mergeColumns(3) = 6
        mergeColumns(4) = 7
        mergeColumns(5) = 9

        '合併儲存格
        For MergeIdx = 0 To 6
            i = 1
            For Each mySingleRow In GridView1.Rows
                If CInt(mySingleRow.RowIndex) > 0 Then  '資料筆數>0

                    If mySingleRow.Cells(mergeColumns(MergeIdx)).Text.Trim() = GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(MergeIdx)).Text.Trim() And _
                       CType(mySingleRow.Cells(3).Controls(0), HyperLink).Text = CType(GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(3).Controls(0), HyperLink).Text Then

                        If mySingleRow.Cells(mergeColumns(MergeIdx)).Text.Trim() <> "&nbsp;" Then   '空白是否
                            '各欄位合併處理
                            GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(MergeIdx)).RowSpan += 1
                            mySingleRow.Cells(mergeColumns(MergeIdx)).Visible = False
                            i = i + 1
                        End If
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
        Dim xColor As Color = Drawing.Color.Plum
        For Each mySingleRow In GridView1.Rows

            If CInt(mySingleRow.RowIndex) > 0 Then

                If CType(GridView1.Rows(CInt(mySingleRow.RowIndex)).Cells(3).Controls(0), HyperLink).Text <> CType(GridView1.Rows(CInt(mySingleRow.RowIndex) - 1).Cells(3).Controls(0), HyperLink).Text Then

                    If xColor = Drawing.Color.Plum Then
                        xColor = Drawing.Color.Khaki
                    Else
                        xColor = Drawing.Color.Plum
                    End If
                End If
            End If

            For i = 0 To 11
                GridView1.Rows(CInt(mySingleRow.RowIndex)).Cells(i).BackColor = xColor
            Next

        Next
    End Sub
    '
    '*****************************************************************
    '**
    '**     GridView資料重整
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        '
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
            e.Row.Cells(12).Visible = False
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            e.Row.Cells(12).Visible = False
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

        For i As Integer = 0 To 11
            e.Row.Cells(i).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        Next
    End Sub
    '
End Class
