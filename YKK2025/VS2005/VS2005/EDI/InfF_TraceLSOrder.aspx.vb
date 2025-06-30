Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfF_TraceLSOrder
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim FCTNo As String             'FCT-NO
    Dim LSNo As String              'LS-NO
    Dim BLSNo As String             'BULS-NO
    Dim OrderNo As String           'Order-NO
    Dim gBuyer As String            'Buyer
    Dim Action As String            'Action

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
            GetAllKeyField()                        '取得FCTNo.+BULSNo.+OrderNo.資訊
            ShowTraceInf()                          '顯示資訊
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
        Response.Cookies("PGM").Value = "InfF_TraceLSOrder.aspx"    '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyy/MM/dd")                '現在日期時間
        FCTNo = Request.QueryString("pFCTNo")               'FCT-NO
        LSNo = ""                                           'LS-NO
        BLSNo = Request.QueryString("pBLSNo")               'BULS-NO
        OrderNo = Request.QueryString("pOrderNo")           'ORDER-NO
        gBuyer = Request.QueryString("pBuyer")              'BUYER
        Action = Request.QueryString("pAction")             'Action
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        FCTGridView.Visible = False
        LSGridView.Visible = False
        PGGridView.Visible = False
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
    '**(GetAllKeyField)
    '**     取得FCTNo.+BULSNo.+OrderNo.資訊
    '**
    '*****************************************************************
    Sub GetAllKeyField()
        If Action = "FCT" Then
            If FCTNo <> "" Then
                LSNo = GetKeyNo("FCT2LS", FCTNo)
            End If

            If LSNo <> "" Then
                BLSNo = GetKeyNo("LS2BULS", LSNo)
            End If

            If BLSNo <> "" Then
                OrderNo = GetKeyNo("BULS2ORDER", BLSNo)
            End If
        Else
            If Action = "LS" Then
                OrderNo = GetKeyNo("BULS2ORDER", BLSNo)
                LSNo = GetKeyNo("BULS2LS", BLSNo)
                FCTNo = GetKeyNo("LS2FCT", LSNo)
            Else
                BLSNo = GetKeyNo("ORDER2BULS", OrderNo)
                LSNo = GetKeyNo("BULS2LS", BLSNo)
                FCTNo = GetKeyNo("LS2FCT", LSNo)
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetKeyNo)
    '**     取得FCTNo.+BULSNo.+OrderNo.
    '**
    '*****************************************************************
    Public Function GetKeyNo(ByVal pKeyType, ByVal pKey) As String
        Dim RtnString As String = ""
        '
        Try
            Dim sql As String
            ' ** FCT <--> LS
            If pKeyType = "FCT2LS" Then
                sql = "SELECT LSNo  + '-' + ltrim(rtrim(str(LSSubNo))) As LSNo "
                sql &= "From LF_ForcastPlan "
                sql &= "Where FCTNo = '" & Mid(pKey, 1, 11) & "' "
                sql &= "  And FCTSubNo = " & Mid(pKey, 13) & " "
                sql &= "  And Version = 99 "
                sql &= "Order by LSNo, LSSubNo "
                '
                Dim dt_FCTPlan As DataTable = uDataBase.GetDataTable(sql)
                If dt_FCTPlan.Rows.Count > 0 Then
                    If dt_FCTPlan.Rows(0).Item("LSNo").ToString <> "-0" Then
                        RtnString = dt_FCTPlan.Rows(0).Item("LSNo").ToString
                    End If
                End If
            End If
            If pKeyType = "LS2FCT" Then
                sql = "SELECT FCTNo  + '-' + ltrim(rtrim(str(FCTSubNo))) As FCTNo "
                sql &= "From LF_ForcastPlan "
                sql &= "Where LSNo = '" & Mid(pKey, 1, 10) & "' "
                sql &= "  And LSSubNo = " & Mid(pKey, 12) & " "
                sql &= "  And Version = 99 "
                sql &= "Order by FCTNo, FCTSubNo "
                '
                Dim dt_FCTPlan As DataTable = uDataBase.GetDataTable(sql)
                If dt_FCTPlan.Rows.Count > 0 Then
                    If dt_FCTPlan.Rows(0).Item("FCTNo").ToString <> "-0" Then
                        RtnString = dt_FCTPlan.Rows(0).Item("FCTNo").ToString
                    End If
                End If
            End If
            ' ** LS <--> BULS
            If pKeyType = "LS2BULS" Then
                sql = "SELECT BULSNo  + '-' + ltrim(rtrim(str(BULSSubNo))) As BULSNo "
                sql &= "From LF_LocalStockPlan "
                sql &= "Where LSNo = '" & Mid(pKey, 1, 10) & "' "
                sql &= "  And LSSubNo = " & Mid(pKey, 12) & " "
                sql &= "  And Version = 99 "
                sql &= "Order by BULSNo, BULSSubNo "
                '
                Dim dt_LSPlan As DataTable = uDataBase.GetDataTable(sql)
                If dt_LSPlan.Rows.Count > 0 Then
                    If dt_LSPlan.Rows(0).Item("BULSNo").ToString <> "-0" Then
                        RtnString = dt_LSPlan.Rows(0).Item("BULSNo").ToString
                    End If
                End If
            End If
            If pKeyType = "BULS2LS" Then
                sql = "SELECT LSNo  + '-' + ltrim(rtrim(str(LSSubNo))) As LSNo "
                sql &= "From LF_LocalStockPlan "
                sql &= "Where BULSNo = '" & Mid(pKey, 1, 10) & "' "
                sql &= "  And BULSSubNo = " & Mid(pKey, 12) & " "
                sql &= "  And Version = 99 "
                sql &= "Order by LSNo, LSSubNo "
                '
                Dim dt_LSPlan As DataTable = uDataBase.GetDataTable(sql)
                If dt_LSPlan.Rows.Count > 0 Then
                    If dt_LSPlan.Rows(0).Item("LSNo").ToString <> "-0" Then
                        RtnString = dt_LSPlan.Rows(0).Item("LSNo").ToString
                    End If
                End If
            End If
            ' ** BULS <--> ORDER
            If pKeyType = "BULS2ORDER" Then
                sql = "SELECT OrderNo  + '-' + ltrim(rtrim(str(OrderSubNo))) As OrderNo "
                ' ADIDAS=000001
                If gBuyer = "000001" Then
                    sql &= "From I_ADIDAS_LSOrder "
                End If
                ' NIKE=000013
                If gBuyer = "000013" Then
                    sql &= "From I_NIKE_LSOrder "
                End If
                ' REEBOK=000016
                If gBuyer = "000016" Then
                    sql &= "From I_REEBOK_LSOrder "
                End If
                ' REEBOK=000021
                If gBuyer = "000021" Then
                    sql &= "From I_TNF_LSOrder "
                End If
                ' COLUMBIA=000003
                If gBuyer = "000003" Then
                    sql &= "From I_COLUMBIA_LSOrder "
                End If
                ' UA=TW0371
                If gBuyer = "TW0371" Then
                    sql &= "From I_UNDERARMOUR_LSOrder "
                End If
                ' T&P NIKE=000013T
                If gBuyer = "000013T" Then
                    sql &= "From I_TPNIKE_LSOrder "
                End If
                ' LULULEMON=TW1741
                If gBuyer = "TW1741" Then
                    sql &= "From I_LULULEMON_LSOrder "
                End If
                '
                sql &= "Where BLSNo = '" & Mid(pKey, 1, 10) & "' "
                sql &= "  And BLSSubNo = " & Mid(pKey, 12) & " "
                sql &= "Order by OrderNo, OrderSubNo "
                '
                Dim dt_LSOrder As DataTable = uDataBase.GetDataTable(sql)
                If dt_LSOrder.Rows.Count > 0 Then
                    If dt_LSOrder.Rows(0).Item("OrderNo").ToString <> "-0" Then
                        RtnString = dt_LSOrder.Rows(0).Item("OrderNo").ToString
                    End If
                End If
            End If
            If pKeyType = "ORDER2BULS" Then
                sql = "SELECT BLSNo  + '-' + ltrim(rtrim(str(BLSSubNo))) As BULSNo "
                ' ADIDAS=000001
                If gBuyer = "000001" Then
                    sql &= "From I_ADIDAS_LSOrder "
                End If
                ' NIKE=000013
                If gBuyer = "000013" Then
                    sql &= "From I_NIKE_LSOrder "
                End If
                ' REEBOK=000016
                If gBuyer = "000016" Then
                    sql &= "From I_REEBOK_LSOrder "
                End If
                ' REEBOK=000021
                If gBuyer = "000021" Then
                    sql &= "From I_TNF_LSOrder "
                End If
                ' COLUMBIA=000003
                If gBuyer = "000003" Then
                    sql &= "From I_COLUMBIA_LSOrder "
                End If
                ' UA=TW0371
                If gBuyer = "TW0371" Then
                    sql &= "From I_UNDERARMOUR_LSOrder "
                End If
                ' T&P NIKE=000013T
                If gBuyer = "000013T" Then
                    sql &= "From I_TPNIKE_LSOrder "
                End If
                ' LULULEMON=TW1741
                If gBuyer = "TW1741" Then
                    sql &= "From I_LULULEMON_LSOrder "
                End If
                '
                sql &= "Where OrderNo = '" & Mid(pKey, 1, 10) & "' "
                sql &= "  And OrderSubNo = " & Mid(pKey, 12) & " "
                sql &= "Order by BLSNo, BLSSubNo "
                '
                Dim dt_LSOrder As DataTable = uDataBase.GetDataTable(sql)
                If dt_LSOrder.Rows.Count > 0 Then
                    If dt_LSOrder.Rows(0).Item("BULSNo").ToString <> "-0" Then
                        RtnString = dt_LSOrder.Rows(0).Item("BULSNo").ToString
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
        '
        Return RtnString
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowTraceInf)
    '**     顯示資訊
    '**
    '*****************************************************************
    Sub ShowTraceInf()
        Dim sql As String
        ' ** Show FCT Inf.
        If LSNo <> "" Then
            sql = "SELECT *, "
            sql &= "Y_ItemCode, Y_ItemName1 + ' ' + Y_ItemName1 as Y_ItemName, Y_Color, '|' As Blank "
            sql &= "From LF_ForcastPlan "
            sql &= "Where LSNo = '" & Mid(LSNo, 1, 10) & "' "
            'sql &= "  And LSSubNo = " & Mid(LSNo, 12) & " "
            'sql &= "Where FCTNo = '" & Mid(FCTNo, 1, 11) & "' "
            'sql &= "  And FCTSubNo = " & Mid(FCTNo, 13) & " "
            sql &= "  And Version = 99 "
            sql &= "Order by FCTSubNo "
            ' 
            Dim dt_ForcastPlan As DataTable = uDataBase.GetDataTable(sql)
            If dt_ForcastPlan.Rows.Count > 0 Then
                FCTGridView.Visible = True
                FCTGridView.DataSource = dt_ForcastPlan
                FCTGridView.DataBind()
            End If
        End If
        ' ** Show Buyer LocalStock Inf.
        If BLSNo <> "" Then
            sql = "SELECT *, GR_04 + ' ' + GR_05 As ItemName, '|' As Blank "
            sql &= "From LF_BuyerLocalStockPlan "
            sql &= "Where BULSNo = '" & Mid(BLSNo, 1, 10) & "' "
            'sql &= "  And BULSSubNo = " & Mid(BLSNo, 12) & " "
            sql &= "  And Version = 99 "
            sql &= "Order by BULSSubNo "
            ' 
            Dim dt_BuyerLocalStockPlan As DataTable = uDataBase.GetDataTable(sql)
            If dt_BuyerLocalStockPlan.Rows.Count > 0 Then
                LSGridView.Visible = True
                LSGridView.DataSource = dt_BuyerLocalStockPlan
                LSGridView.DataBind()
            End If
        End If
        ' ** Wave's Order Inf.
        If BLSNo <> "" Then
            sql = "SELECT *, ItemName1 + ' ' + ItemName2 As ItemName, Cast(OrderQty As int) As Qty "
            ' ADIDAS=000001
            If gBuyer = "000001" Then
                sql &= "From I_ADIDAS_LSOrder "
            End If
            ' NIKE=000013
            If gBuyer = "000013" Then
                sql &= "From I_NIKE_LSOrder "
            End If
            ' REEBOK=000016
            If gBuyer = "000016" Then
                sql &= "From I_REEBOK_LSOrder "
            End If
            ' REEBOK=000021
            If gBuyer = "000021" Then
                sql &= "From I_TNF_LSOrder "
            End If
            ' COLUMBIA=000003
            If gBuyer = "000003" Then
                sql &= "From I_COLUMBIA_LSOrder "
            End If
            ' UA=TW0371
            If gBuyer = "TW0371" Then
                sql &= "From I_UNDERARMOUR_LSOrder "
            End If
            ' T&P NIKE=000013T
            If gBuyer = "000013T" Then
                sql &= "From I_TPNIKE_LSOrder "
            End If
            ' LULULEMON=TW1741
            If gBuyer = "TW1741" Then
                sql &= "From I_LULULEMON_LSOrder "
            End If
            '
            sql &= "Where BLSNo = '" & Mid(BLSNo, 1, 10) & "' "
            sql &= "Order by BLSNo, BLSSubNo, BLSMM, BLSMMSeqNo "
            'sql &= "Where OrderNo = '" & Mid(OrderNo, 1, 10) & "' "
            'sql &= "  And OrderSubNo = " & Mid(OrderNo, 12) & " "
            'sql &= "Order by OrderSubNo "
            '
            Dim dt_LSOrder As DataTable = uDataBase.GetDataTable(sql)
            If dt_LSOrder.Rows.Count > 0 Then
                PGGridView.Visible = True
                PGGridView.DataSource = dt_LSOrder
                PGGridView.DataBind()
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯BuyerLocalStockData GridView
    '**
    '*****************************************************************
    Protected Sub LSGridView_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles LSGridView.RowDataBound
        Dim row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
        Dim row1 As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
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
