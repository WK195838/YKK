Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfF_OrderProgress
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim Buyer As String             'Buyer
    Dim UserID As String            'UserID
    Dim OrderNo As String           'OrderProgressGrid
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
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
            SetDefaultValue()                       '設定初值
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
        Response.Cookies("PGM").Value = "InfF_OrderProgress.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        Buyer = Request.QueryString("pBuyer")
        If Buyer <> "000013T" And Buyer <> "TW0371T" Then
            Buyer = Mid(Request.QueryString("pBuyer"), 1, 6)
        End If
        UserID = Request.QueryString("pUserID")
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        PGGridView.Visible = False
        LSGridView.Visible = False
        FCTGridView.Visible = False
        DBuyer.ReadOnly = True
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        DBuyer.Text = Buyer
        DOrderNo.Text = ""

        Label2.Visible = False
        DFCTNo.Visible = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Inq)
    '**     查詢
    '**
    '*****************************************************************
    Protected Sub BInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BInq.Click
        If Not InputError() Then
            ' FCT Process
            If DFCTNo.Text <> "" Then
                GetForcastData(DFCTNo.Text)
            End If
            ' Buyer LS Process
            If DBLSNo.Text <> "" Then
                GetBuyerLocalStockData(DBLSNo.Text)
            End If
            ' Order Process
            If (DOrderNo.Text <> "") Or (DFCTNo.Text = "" And DBLSNo.Text = "" And DOrderNo.Text = "") Then

                If DOrderNo.Text <> "" Then
                    OrderNo = DOrderNo.Text
                Else
                    OrderNo = "ALL"
                End If
                GetOrderProgressData(OrderNo)
            End If
        Else
            uJavaScript.PopMsg(Me, "未輸入資料或輸入資料異常,請確認!!")
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetOrderProgressData)
    '**     查詢 Order Progress
    '**
    '*****************************************************************
    Protected Sub GetOrderProgressData(ByVal pOrderNo As String)
        Dim sql As String
        ' 篩選資料
        sql = "SELECT *, "
        sql &= "ItemName1 + ' ' + ItemName2 As ItemName, Cast(OrderQty As int) As Qty, "
        sql &= "'InfF_TraceLSOrder.aspx?' + "
        sql &= "'pBuyer=' + '" + DBuyer.Text + "' + "
        sql &= "'&pFCTNo=' + '" + "" + "' + "
        sql &= "'&pBLSNo=' + '" + "" + "' + "
        sql &= "'&pOrderNo=' + OrderNo + '-' + ltrim(rtrim(str(OrderSubNo))) + "
        sql &= "'&pAction=' + '" + "ORDER" + "' "
        sql &= " As URL "
        '
        ' ADIDAS=000001
        If DBuyer.Text = "000001" Then
            sql &= "From I_ADIDAS_LSOrder "
        End If
        ' NIKE=000013
        If DBuyer.Text = "000013" Then
            sql &= "From I_NIKE_LSOrder "
        End If
        ' REEBOK=000016
        If DBuyer.Text = "000016" Then
            sql &= "From I_REEBOK_LSOrder "
        End If
        ' TNF=000021
        If DBuyer.Text = "000021" Then
            sql &= "From I_TNF_LSOrder "
        End If
        ' UNDERARMOUR=TW0371
        If DBuyer.Text = "TW0371" Then
            sql &= "From I_UNDERARMOUR_LSOrder "
        End If
        ' T&P NIKE=000013T
        If DBuyer.Text = "000013T" Then
            sql &= "From I_TPNIKE_LSOrder "
        End If
        ' UABAGS=TW0371T
        If DBuyer.Text = "TW0371T" Then
            sql &= "From I_UABAGS_LSOrder "
        End If
        ' LULULEMON=TW1741
        If DBuyer.Text = "TW1741" Then
            sql &= "From I_LULULEMON_LSOrder "
        End If
        If pOrderNo = "ALL" Then
            ' 查詢異常
            sql &= "Where CompletedFlag = '" & "" & "' "
            sql &= "  And BLSNo <> '" & "" & "' "
            sql &= "  And PlanDate >= " & NowDate & " "
            sql &= "  And SubString(LTrim(RTrim(Str(Reqdate))),1,6) < SubString(LTrim(RTrim(Str(PlanDate))),1,6) "
            sql &= "Order by ReqDate, PlanDate "
        Else
            sql &= "Where OrderNo = '" & pOrderNo & "' "
            sql &= "Order by OrderSubNo "
        End If
        '
        Dim dt_LSOrder As DataTable = uDataBase.GetDataTable(sql)
        If dt_LSOrder.Rows.Count > 0 Then
            PGGridView.Style("Top") = 40 & "px"
            PGGridView.Visible = True
            PGGridView.DataSource = dt_LSOrder
            PGGridView.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetBuyerLocalStockData)
    '**     查詢 Buyer LocalStock Plan
    '**
    '*****************************************************************
    Protected Sub GetBuyerLocalStockData(ByVal pLSNo As String)
        Dim sql As String
        ' 篩選資料
        sql = "SELECT *, "
        sql &= "GR_04 + ' ' + GR_05 As ItemName, '|' As Blank, "
        sql &= "'InfF_TraceLSOrder.aspx?' + "
        sql &= "'pBuyer=' + '" + DBuyer.Text + "' + "
        sql &= "'&pFCTNo=' + '" + "" + "' + "
        sql &= "'&pBLSNo=' + BULSNo + '-' + ltrim(rtrim(str(BULSSubNo))) + "
        sql &= "'&pOrderNo=' + '" + "" + "' + "
        sql &= "'&pAction=' + '" + "LS" + "' "
        sql &= " As URL "
        sql &= "From LF_BuyerLocalStockPlan "
        sql &= "Where BULSNo = '" & DBLSNo.Text & "' "
        sql &= "Order by BULSSubNo "
        ' 
        Dim dt_BuyerLocalStockPlan As DataTable = uDataBase.GetDataTable(sql)
        If dt_BuyerLocalStockPlan.Rows.Count > 0 Then
            LSGridView.Style("Top") = 40 & "px"
            LSGridView.Visible = True
            LSGridView.DataSource = dt_BuyerLocalStockPlan
            LSGridView.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetForcastData)
    '**     查詢 Forcast Plan
    '**
    '*****************************************************************
    Protected Sub GetForcastData(ByVal pFCTNo As String)
        Dim sql As String
        ' 篩選資料
        sql = "SELECT *, "
        sql &= "Y_ItemCode, Y_ItemName1 + ' ' + Y_ItemName1 as Y_ItemName, Y_Color, '|' As Blank, "
        sql &= "'InfF_TraceLSOrder.aspx?' + "
        sql &= "'pBuyer=' + '" + DBuyer.Text + "' + "
        sql &= "'&pFCTNo=' + FCTNo + '-' + ltrim(rtrim(str(FCTSubNo))) + "
        sql &= "'&pBLSNo=' + '" + "" + "' + "
        sql &= "'&pOrderNo=' + '" + "" + "' + "
        sql &= "'&pAction=' + '" + "FCT" + "' "
        sql &= " As URL "
        sql &= "From LF_ForcastPlan "
        sql &= "Where FCTNo = '" & DFCTNo.Text & "' "
        sql &= "Order by FCTSubNo "
        ' 
        Dim dt_ForcastPlan As DataTable = uDataBase.GetDataTable(sql)
        If dt_ForcastPlan.Rows.Count > 0 Then
            FCTGridView.Style("Top") = 40 & "px"
            FCTGridView.Visible = True
            FCTGridView.DataSource = dt_ForcastPlan
            FCTGridView.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(OrderProgress GridView Page Change)
    '**     PGGridView_PageIndexChanging
    '**
    '*****************************************************************
    Protected Sub PGGridView_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles PGGridView.PageIndexChanging
        '
        If (DOrderNo.Text <> "") Or (DFCTNo.Text = "" And DBLSNo.Text = "" And DOrderNo.Text = "") Then
            If DOrderNo.Text <> "" Then
                OrderNo = DOrderNo.Text
            Else
                OrderNo = "ALL"
            End If
            GetOrderProgressData(OrderNo)
        End If
        '
        PGGridView.PageIndex = e.NewPageIndex
        GetOrderProgressData(OrderNo)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(InputError)
    '**     InputError
    '**
    '*****************************************************************
    Public Function InputError() As Boolean
        Dim RtnCode As Boolean = False
        '
        Try
            If (DFCTNo.Text <> "" And DBLSNo.Text <> "") Or (DFCTNo.Text <> "" And DOrderNo.Text <> "") Or _
               (DBLSNo.Text <> "" And DOrderNo.Text <> "") Then
                RtnCode = True
            End If
        Catch ex As Exception
            RtnCode = True
        End Try
        '
        Return RtnCode
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯BuyerLocalStockData GridView
    '**
    '*****************************************************************
    Protected Sub LSGridView_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles LSGridView.RowDataBound
        Dim row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
        Dim C_Blank As TableCell = New TableCell
        Dim C_Line As TableCell = New TableCell
        Dim C_SP01 As TableCell = New TableCell
        Dim C_OP01 As TableCell = New TableCell
        Dim C_FS01 As TableCell = New TableCell
        Dim C_PS01 As TableCell = New TableCell
        Dim C_IS01 As TableCell = New TableCell
        Dim C1_Line As TableCell = New TableCell
        Dim C1_SP01 As TableCell = New TableCell
        Dim C1_OP01 As TableCell = New TableCell
        Dim C1_FS01 As TableCell = New TableCell
        Dim C1_PS01 As TableCell = New TableCell
        Dim C1_IS01 As TableCell = New TableCell
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
            '--N3
            C_Blank.BackColor = System.Drawing.ColorTranslator.FromHtml("#990000")
            C_Blank.ForeColor = Color.White
            C_Blank.HorizontalAlign = HorizontalAlign.Center
            C_Blank.Font.Bold = True
            C_Blank.ColumnSpan = 13
            C_Blank.Text = ""
            row.Cells.Add(C_Blank)

            C_Line.BackColor = System.Drawing.ColorTranslator.FromHtml("#990000")
            C_Line.ForeColor = Color.White
            C_Line.HorizontalAlign = HorizontalAlign.Center
            C_Line.Font.Bold = True
            C_Line.ColumnSpan = 1
            C_Line.Text = "|"
            row.Cells.Add(C_Line)

            C_SP01.BackColor = System.Drawing.ColorTranslator.FromHtml("#990000")
            C_SP01.ForeColor = Color.White
            C_SP01.HorizontalAlign = HorizontalAlign.Center
            C_SP01.Font.Bold = True
            C_SP01.ColumnSpan = 1
            C_SP01.Text = "N3_SCHE"
            row.Cells.Add(C_SP01)

            C_OP01.BackColor = System.Drawing.ColorTranslator.FromHtml("#990000")
            C_OP01.ForeColor = Color.White
            C_OP01.HorizontalAlign = HorizontalAlign.Center
            C_OP01.Font.Bold = True
            C_OP01.ColumnSpan = 1
            C_OP01.Text = "N3_ON"
            row.Cells.Add(C_OP01)

            C_FS01.BackColor = System.Drawing.ColorTranslator.FromHtml("#990000")
            C_FS01.ForeColor = Color.White
            C_FS01.HorizontalAlign = HorizontalAlign.Center
            C_FS01.Font.Bold = True
            C_FS01.ColumnSpan = 1
            C_FS01.Text = "N3_F"
            row.Cells.Add(C_FS01)

            C_PS01.BackColor = System.Drawing.ColorTranslator.FromHtml("#990000")
            C_PS01.ForeColor = Color.White
            C_PS01.HorizontalAlign = HorizontalAlign.Center
            C_PS01.Font.Bold = True
            C_PS01.ColumnSpan = 1
            C_PS01.Text = "N3_P"
            row.Cells.Add(C_PS01)

            C_IS01.BackColor = System.Drawing.ColorTranslator.FromHtml("#990000")
            C_IS01.ForeColor = Color.White
            C_IS01.HorizontalAlign = HorizontalAlign.Center
            C_IS01.Font.Bold = True
            C_IS01.ColumnSpan = 1
            C_IS01.Text = "N3_I"
            row.Cells.Add(C_IS01)
            '-----------------------------------------------------------------------------
            '
            '--N4 
            C1_Line.BackColor = System.Drawing.ColorTranslator.FromHtml("#990000")
            C1_Line.ForeColor = Color.White
            C1_Line.HorizontalAlign = HorizontalAlign.Center
            C1_Line.Font.Bold = True
            C1_Line.ColumnSpan = 1
            C1_Line.Text = "|"
            row.Cells.Add(C1_Line)

            C1_SP01.BackColor = System.Drawing.ColorTranslator.FromHtml("#990000")
            C1_SP01.ForeColor = Color.White
            C1_SP01.HorizontalAlign = HorizontalAlign.Center
            C1_SP01.Font.Bold = True
            C1_SP01.ColumnSpan = 1
            C1_SP01.Text = "N4_SCHE"
            row.Cells.Add(C1_SP01)

            C1_OP01.BackColor = System.Drawing.ColorTranslator.FromHtml("#990000")
            C1_OP01.ForeColor = Color.White
            C1_OP01.HorizontalAlign = HorizontalAlign.Center
            C1_OP01.Font.Bold = True
            C1_OP01.ColumnSpan = 1
            C1_OP01.Text = "N4_ON"
            row.Cells.Add(C1_OP01)

            C1_FS01.BackColor = System.Drawing.ColorTranslator.FromHtml("#990000")
            C1_FS01.ForeColor = Color.White
            C1_FS01.HorizontalAlign = HorizontalAlign.Center
            C1_FS01.Font.Bold = True
            C1_FS01.ColumnSpan = 1
            C1_FS01.Text = "N4_F"
            row.Cells.Add(C1_FS01)

            C1_PS01.BackColor = System.Drawing.ColorTranslator.FromHtml("#990000")
            C1_PS01.ForeColor = Color.White
            C1_PS01.HorizontalAlign = HorizontalAlign.Center
            C1_PS01.Font.Bold = True
            C1_PS01.ColumnSpan = 1
            C1_PS01.Text = "N4_P"
            row.Cells.Add(C1_PS01)

            C1_IS01.BackColor = System.Drawing.ColorTranslator.FromHtml("#990000")
            C1_IS01.ForeColor = Color.White
            C1_IS01.HorizontalAlign = HorizontalAlign.Center
            C1_IS01.Font.Bold = True
            C1_IS01.ColumnSpan = 1
            C1_IS01.Text = "N4_I"
            row.Cells.Add(C1_IS01)
            ' 加行
            e.Row.Parent.Controls.Add(row)
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            '
            '--N3 
            C_Blank.ColumnSpan = 13
            C_Blank.Text = ""
            row.Cells.Add(C_Blank)

            C_Line.ColumnSpan = 1
            C_Line.Text = "|"
            row.Cells.Add(C_Line)

            C_SP01.ColumnSpan = 1
            C_SP01.HorizontalAlign = HorizontalAlign.Right
            C_SP01.Text = DataBinder.Eval(e.Row.DataItem, "SP_03")
            row.Cells.Add(C_SP01)

            C_OP01.ColumnSpan = 1
            C_OP01.HorizontalAlign = HorizontalAlign.Right
            C_OP01.Text = DataBinder.Eval(e.Row.DataItem, "OP_03")
            row.Cells.Add(C_OP01)

            C_FS01.ColumnSpan = 1
            C_FS01.HorizontalAlign = HorizontalAlign.Right
            C_FS01.Text = DataBinder.Eval(e.Row.DataItem, "FS_03")
            row.Cells.Add(C_FS01)

            C_PS01.ColumnSpan = 1
            C_PS01.HorizontalAlign = HorizontalAlign.Right
            C_PS01.Text = DataBinder.Eval(e.Row.DataItem, "PS_03")
            row.Cells.Add(C_PS01)

            C_IS01.ColumnSpan = 1
            C_IS01.HorizontalAlign = HorizontalAlign.Right
            C_IS01.Text = DataBinder.Eval(e.Row.DataItem, "IS_03")
            row.Cells.Add(C_IS01)
            '----------------------------------------------
            '
            '--N4
            C1_Line.ColumnSpan = 1
            C1_Line.Text = "|"
            row.Cells.Add(C1_Line)

            C1_SP01.ColumnSpan = 1
            C1_SP01.HorizontalAlign = HorizontalAlign.Right
            C1_SP01.Text = DataBinder.Eval(e.Row.DataItem, "SP_04")
            row.Cells.Add(C1_SP01)

            C1_OP01.ColumnSpan = 1
            C1_OP01.HorizontalAlign = HorizontalAlign.Right
            C1_OP01.Text = DataBinder.Eval(e.Row.DataItem, "OP_04")
            row.Cells.Add(C1_OP01)

            C1_FS01.ColumnSpan = 1
            C1_FS01.HorizontalAlign = HorizontalAlign.Right
            C1_FS01.Text = DataBinder.Eval(e.Row.DataItem, "FS_04")
            row.Cells.Add(C1_FS01)

            C1_PS01.ColumnSpan = 1
            C1_PS01.HorizontalAlign = HorizontalAlign.Right
            C1_PS01.Text = DataBinder.Eval(e.Row.DataItem, "PS_04")
            row.Cells.Add(C1_PS01)

            C1_IS01.ColumnSpan = 1
            C1_IS01.HorizontalAlign = HorizontalAlign.Right
            C1_IS01.Text = DataBinder.Eval(e.Row.DataItem, "IS_04")
            row.Cells.Add(C1_IS01)
            ' 加行
            e.Row.Parent.Controls.Add(row)
        End If
    End Sub
    ''
    ''*****************************************************************
    ''**
    ''**     轉Excel
    ''**
    ''*****************************************************************
    'Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
    '    '
    '    Response.AppendHeader("Content-Disposition", "attachment;filename=OrderProgress.xls")     '程式別不同
    '    Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

    '    Response.ContentType = "application/vnd.ms-excel"
    '    Response.Charset = ""
    '    Me.EnableViewState = False

    '    Dim tw As New System.IO.StringWriter
    '    Dim hw As New System.Web.UI.HtmlTextWriter(tw)
    '    '
    '    'ShowData()
    '    'GridView1.RenderControl(hw)
    '    Response.Write(tw.ToString())
    '    Response.End()
    'End Sub
    ''
    ''*****************************************************************
    ''**
    ''**     轉Excel共用程式
    ''**
    ''*****************************************************************
    'Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
    '    'Confirms that an HtmlForm control is rendered for the specified ASP.NET
    '    ' server control at run time. */
    'End Sub



End Class
