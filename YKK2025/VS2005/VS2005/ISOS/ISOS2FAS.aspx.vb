Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class ISOS2FAS
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '外部Object
    Dim fpObj As New ForProject
    Dim oDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim oCommon As New Utility.Common
    Dim oJavaScript As New Utility.JScript
    Dim oWaves As New Waves.CommonService
    Dim oFASCommon As New EDI2011.FCommonService

    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    '
    Dim pBuyer As String            'Buyer
    Dim UserID As String            'UserID
    Dim pItem As String             'Item
    Dim pItemName1 As String        'ItemName1
    Dim pItemName2 As String        'ItemName2
    Dim pItemName3 As String        'ItemName3
    Dim EDLConnect As String = oCommon.GetAppSetting("EDLDB")

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                          '設定參數
        If Not IsPostBack Then                  'PostBack
            SetDefaultValue()                   '設定初值
            'ShowData()
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
        Server.ScriptTimeout = 900                          '設定逾時時間
        Response.Cookies("PGM").Value = "ISOS2FAS.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '
        pBuyer = Request.QueryString("pBuyer")
        UserID = Request.QueryString("pUserID")
        pItem = Request.QueryString("pItem")
        pItemName1 = Request.QueryString("pItemName1")        'ItemName1
        pItemName2 = Request.QueryString("pItemName2")        'ItemName2
        pItemName3 = Request.QueryString("pItemName3")        'ItemName3
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        DBUYER.Text = "FALL-" & pBuyer
        DItem.Text = pItem
        DColor.Text = "TN858"
        DKeepCode.Text = ""
        DPuller.Text = ""
        '
        DItemName1.Text = pItemName1
        DItemName2.Text = pItemName2
        DItemName3.Text = pItemName3
        '
        DISOS2FASFile.Style("left") = -500 & "px"
        DISOS2FASFile.Text = oCommon.GetAppSetting("DataPrepareFile") + "ISOS2FAS_" & UserID & ".xlsm"
        '動作按鈕設定
        BInput.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('ISOS2FASExcel')}"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Inq)
    '**     查詢
    '**
    '*****************************************************************
    Protected Sub BInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BInq.Click
        '單筆
        ShowData()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Inq)
    '**     查詢
    '**
    '*****************************************************************
    Protected Sub BManyInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BManyInq.Click
        '多筆
        ShowManyData()
    End Sub
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim xKeepCode(20) As String
        Dim i, xCode As Integer
        Dim sql As String
        '
        '轉大寫
        DBUYER.Text = UCase(Trim(DBUYER.Text))
        DItem.Text = UCase(Trim(DItem.Text))
        DColor.Text = UCase(DColor.Text)
        DKeepCode.Text = UCase(Trim(DKeepCode.Text))
        DPuller.Text = UCase(Trim(DPuller.Text))
        '
        DItemName1.Text = UCase(Trim(DItemName1.Text))
        DItemName2.Text = UCase(Trim(DItemName2.Text))
        DItemName3.Text = UCase(Trim(DItemName3.Text))
        '
        '清除資料
        sql = "DELETE FROM ForcastPlan_ISOS " & _
            " Where ModifyUser = '" & UserID & "' "
        oDataBase.ExecuteNonQuery(sql)
        sql = "DELETE FROM LocalStockPlan_ISOS " & _
            " Where ModifyUser = '" & UserID & "' "
        oDataBase.ExecuteNonQuery(sql)
        '
        'GET KEEPCODE LIST
        For i = 1 To 20
            xKeepCode(i) = ""
        Next
        If DKeepCode.Text <> "" Then
            xKeepCode(1) = DKeepCode.Text
        Else
            sql = "Select ShortKeepCode From M_NativeVendor "
            sql = sql & " Where Buyer = '" & DBUYER.Text & "' "
            sql = sql & "   And Active = 1 "
            sql = sql & "   And ShortKeepCode <> '' "
            sql = sql & "Group by ShortKeepCode "
            sql = sql & "Order by ShortKeepCode "
            Dim dt_Referp As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_Referp.Rows.Count - 1
                xKeepCode(i + 1) = dt_Referp.Rows(i).Item("ShortKeepCode")
            Next
        End If
        '
        'INSERT FORCAST DATA
        For i = 1 To 20
            If xKeepCode(i) = "" Then
                Exit For
            Else
                sql = "INSERT INTO ForcastPlan_ISOS " & _
                      " SELECT "
                'Buyer, ID, BuyerGroup, FCTNo, FCTSubNo, 
                sql &= "'" & DBUYER.Text & "', " & _
                      "1" & ", " & _
                      "'" & pBuyer & "B" & "', " & _
                      "'" & "ZZF00000001" & "', " & _
                      "1" & ", "
                'Version, C_Code, C_Color, C_SpecialRequest, C_Season, 
                sql &= "99" & ", " & _
                      "'" & DItem.Text & "', " & _
                      "'" & DPuller.Text & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', "
                'C_ShortenLT, C_Style, C_A1, C_B1, C_C1, 
                sql &= "'" & xKeepCode(i) & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', "
                'C_D1, C_E1, C_F1, C_G1, C_H1, 
                sql &= "'" & "" & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', "
                'C_I1, C_J1, C_K1, C_L1, C_M1, C_N1, C_O1
                sql &= "'" & "" & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', "
                'Y_LEVEL, Y_ItemCode, Y_ItemName1, Y_ItemName2, Y_ItemName3,
                sql &= "0" & ", " & _
                      "'" & DItem.Text & "', " & _
                      "'" & DItemName1.Text & "', " & _
                      "'" & DItemName2.Text & "', " & _
                      "'" & DItemName3.Text & " ', "
                'Y_Color, Y_A1, Y_B1, Y_C1, Y_D1, Y_E1,
                sql &= "'" & DColor.Text & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', "
                'Y_F1, Y_G1, Y_H1, Y_I1, Y_J1,
                sql &= "'" & "" & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', " & _
                      "'" & "" & "', "
                'Total, N_F, N1_F, N2_F, N3_F,
                sql &= "1000" & ", " & _
                      "0" & ", " & _
                      "1000" & ", " & _
                      "0" & ", " & _
                      "0" & ", "
                'N4_F, N5_F, N6_F, N7_F, N8_F, N9_F, N10_F, N11_F, N12_F,
                sql &= "0" & ", " & _
                      "0" & ", " & _
                      "0" & ", " & _
                      "0" & ", " & _
                      "0" & ", " & _
                      "0" & ", " & _
                      "0" & ", " & _
                      "0" & ", " & _
                      "0" & ", "
                'LSNo, LSSubNo, CreateUser, CreateTime, ModifyUser, ModifyTime
                sql &= "'" & "" & "', " & _
                      "0" & ", " & _
                      "'" & "ISOS" & "', " & _
                      "'" & NowDateTime & "', " & _
                      "'" & UserID & "', " & _
                      "'" & NowDateTime & "' "
                '
                oDataBase.ExecuteNonQuery(sql)
            End If
        Next
        '
        '展開
        xCode = oFASCommon.ISOS2FAS(Now.ToString("yyyyMMddHHmmss"), DBUYER.Text, UserID, pBuyer & "B", "91111XXXXX", 0)
        If xCode = 0 Then
            sql = "SELECT "
            sql &= "BUYER, C_CODE, ITEM_NAME1 + ' ' + ITEM_NAME2 + ' ' + ITEM_NAME3 as ITEMNAME, Y_COLOR, C_COLOR, C_ShortenLT, N1_F "
            sql &= "From V_ISOS2FAS_FCT "
            sql &= "Where Buyer = '" & DBUYER.Text & "' "
            sql &= "  And Y_LEVEL = '" & "0" & "' "
            sql &= "  And ModifyUser = '" & UserID & "' "
            sql &= "Order by BUYER, C_CODE, Y_COLOR, C_ShortenLT "
            '
            Dim dt_FCTPlanList As DataTable = oDataBase.GetDataTable(sql)
            If dt_FCTPlanList.Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = dt_FCTPlanList
                GridView1.DataBind()
                '
                sql = "SELECT "
                sql &= "Buyer, GR_02, GR_03, GR_04, GR_05, GR_06, GR_07, GR_08, "
                sql &= "MinimumStock, FS_01, N_ScheProd, N_OnProd, N_FreeInv, N_KeepInv, N_Total, "
                sql &= "'" & NowDateTime & "' as DataTime "
                sql &= "From LocalStockPlan_ISOS "
                sql &= "Where Buyer = '" & DBUYER.Text & "' "
                sql &= "  And GR_01 IN ('" & "LS" & "', '" & "LIST" & "') "
                sql &= "  And ModifyUser = '" & UserID & "' "
                sql &= "Order by Buyer, GR_01, GR_02, GR_03, GR_04, GR_05, GR_06, GR_07, GR_08 "
                '
                Dim dt_LSPlanList As DataTable = oDataBase.GetDataTable(sql)
                If dt_LSPlanList.Rows.Count > 0 Then
                    GridView2.Visible = True
                    GridView2.DataSource = dt_LSPlanList
                    GridView2.DataBind()
                End If
            Else
                oJavaScript.PopMsg(Me, "搜尋不到資料,請確認! (No Data)")
            End If
            '
        Else
            oJavaScript.PopMsg(Me, "搜尋不到資料,請確認! (ISOS2FAS)")
        End If
        '
    End Sub
    '*****************************************************************
    '**(ShowManyData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowManyData()
        Dim xCode As Integer
        Dim sql As String
        '
        '展開
        xCode = oFASCommon.ISOS2FAS(Now.ToString("yyyyMMddHHmmss"), DBUYER.Text, UserID, pBuyer & "B", "91111XXXXX", 0)
        xCode = 0
        If xCode = 0 Then
            sql = "SELECT "
            sql &= "BUYER, C_CODE, ITEM_NAME1 + ' ' + ITEM_NAME2 + ' ' + ITEM_NAME3 as ITEMNAME, Y_COLOR, C_COLOR, C_ShortenLT, N1_F "
            sql &= "From V_ISOS2FAS_FCT "
            sql &= "Where Buyer = '" & DBUYER.Text & "' "
            sql &= "  And Y_LEVEL = '" & "0" & "' "
            sql &= "  And ModifyUser = '" & UserID & "' "
            sql &= "Order by BUYER, C_CODE, Y_COLOR, C_ShortenLT "
            '
            Dim dt_FCTPlanList As DataTable = oDataBase.GetDataTable(sql)
            If dt_FCTPlanList.Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = dt_FCTPlanList
                GridView1.DataBind()
                '
                sql = "SELECT "
                sql &= "Buyer, GR_02, GR_03, GR_04, GR_05, GR_06, GR_07, GR_08, "
                sql &= "MinimumStock, FS_01, N_ScheProd, N_OnProd, N_FreeInv, N_KeepInv, N_Total, "
                sql &= "'" & NowDateTime & "' as DataTime "
                sql &= "From LocalStockPlan_ISOS "
                sql &= "Where Buyer = '" & DBUYER.Text & "' "
                sql &= "  And GR_01 IN ('" & "LS" & "', '" & "LIST" & "') "
                sql &= "  And ModifyUser = '" & UserID & "' "
                sql &= "Order by Buyer, GR_01, GR_02, GR_03, GR_04, GR_05, GR_06, GR_07, GR_08 "
                '
                Dim dt_LSPlanList As DataTable = oDataBase.GetDataTable(sql)
                If dt_LSPlanList.Rows.Count > 0 Then
                    GridView2.Visible = True
                    GridView2.DataSource = dt_LSPlanList
                    GridView2.DataBind()
                End If
            Else
                oJavaScript.PopMsg(Me, "搜尋不到資料,請確認! (No Data)")
            End If
            '
        Else
            oJavaScript.PopMsg(Me, "搜尋不到資料,請確認! (ISOS2FAS)")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        '
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
        End If
        '---------------------------------------
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
        End If
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel
    '**
    '*****************************************************************
    Protected Sub BExcel1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel1.Click
        Dim sql As String
        '
        Response.Clear()
        Response.Buffer = True

        Response.AppendHeader("Content-Disposition", "attachment;filename=ItemStockList1.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        Dim style As String = "<style> .text { mso-number-format:\@; } </style> " '文字樣式字串
        '
        sql = "SELECT "
        sql &= "BUYER, C_CODE, ITEM_NAME1 + ' ' + ITEM_NAME2 + ' ' + ITEM_NAME3 as ITEMNAME, Y_COLOR, C_COLOR, C_ShortenLT, N1_F "
        sql &= "From V_ISOS2FAS_FCT "
        sql &= "Where Buyer = '" & DBUYER.Text & "' "
        sql &= "  And Y_LEVEL = '" & "0" & "' "
        sql &= "  And ModifyUser = '" & UserID & "' "
        sql &= "Order by BUYER, C_CODE, Y_COLOR, C_ShortenLT "
        '
        Dim dt_FCTPlanList As DataTable = oDataBase.GetDataTable(sql)
        If dt_FCTPlanList.Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = dt_FCTPlanList
            GridView1.DataBind()
        End If
        '
        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()
    End Sub

    Protected Sub BExcel2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel2.Click
        Dim sql As String
        '
        Response.Clear()
        Response.Buffer = True

        Response.AppendHeader("Content-Disposition", "attachment;filename=ItemStockList2.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        Dim style As String = "<style> .text { mso-number-format:\@; } </style> " '文字樣式字串
        '
        Sql = "SELECT "
        Sql &= "Buyer, GR_02, GR_03, GR_04, GR_05, GR_06, GR_07, GR_08, "
        Sql &= "MinimumStock, FS_01, N_ScheProd, N_OnProd, N_FreeInv, N_KeepInv, N_Total, "
        Sql &= "'" & NowDateTime & "' as DataTime "
        Sql &= "From LocalStockPlan_ISOS "
        Sql &= "Where Buyer = '" & DBUYER.Text & "' "
        sql &= "  And GR_01 IN ('" & "LS" & "', '" & "LIST" & "') "
        sql &= "  And ModifyUser = '" & UserID & "' "
        Sql &= "Order by Buyer, GR_01, GR_02, GR_03, GR_04, GR_05, GR_06, GR_07, GR_08 "
        '
        Dim dt_LSPlanList As DataTable = oDataBase.GetDataTable(Sql)
        If dt_LSPlanList.Rows.Count > 0 Then
            GridView2.Visible = True
            GridView2.DataSource = dt_LSPlanList
            GridView2.DataBind()
        End If
        '
        GridView2.RenderControl(hw)
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
    '---------------------------------------------------------------------------------------------------
End Class
