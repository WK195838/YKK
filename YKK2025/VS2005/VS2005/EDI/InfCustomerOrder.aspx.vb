Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfCustomerOrder
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
        Response.Cookies("PGM").Value = "InfCustomerOrder.aspx"     '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyy/MM/dd")   '現在日期時間
        wUserID = Request.QueryString("pUserID")
        DBuyer.Text = Request.QueryString("pBuyer")
        DPO.Text = Request.QueryString("pPO")
        DInputDate.Text = Request.QueryString("pDate")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        'DBuyer.ReadOnly = True
        'DPO.ReadOnly = True
        'DInputDate.ReadOnly = True
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim sql, xGroupCode As String
        Dim xStartTime As String = Mid(DInputDate.Text, 1, 4) + "/" + Mid(DInputDate.Text, 5, 2) + "/" + Mid(DInputDate.Text, 7, 2) + " 00:00:00"
        Dim xEndTime As String = Mid(DInputDate.Text, 1, 4) + "/" + Mid(DInputDate.Text, 5, 2) + "/" + Mid(DInputDate.Text, 7, 2) + " 23:59:59"
        '
        ' 取得 Buyer GroupCode
        sql = "SELECT BuyerGroup FROM M_ControlRecord "
        sql &= "Where Buyer = '" & DBuyer.Text & "' "
        Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
        If dt_ControlRecord.Rows.Count > 0 Then
            xGroupCode = dt_ControlRecord.Rows(0).Item("BuyerGroup")
        End If
        '
        '篩選資料
        If fpObj.GetFunctionCode(xGroupCode, 2) = "P" Then
            '
            sql = "SELECT "
            sql &= "BM1 As PO, BN1 As Seqno, "
            sql &= "A1, B1, C1, D1, E1, F1, G1, H1, I1, J1, K1, L1, M1, "
            sql &= "N1, O1, P1, Q1, R1, S1, T1, U1, V1, W1, X1, Y1, Z1  "
            sql &= "From B_CustomerRequest "
            sql &= "Where Buyer = '" & DBuyer.Text & "' "

            If DPO.Text <> "" Then
                sql &= "  And BM1 = '" & DPO.Text & "' "
            Else
                If DInputDate.Text <> "" Then
                    sql &= "  And CreateTime >= '" & xStartTime & "' "
                    sql &= "  And CreateTime <= '" & xEndTime & "' "
                Else
                    sql &= " And BM1 = '" & DPO.Text & "' "
                End If
            End If
            sql &= "Order by BM1, BN1 "
        Else
            '
            sql = "SELECT "
            sql &= "E1 As PO, D1 As Seqno, "
            sql &= "A1, B1, C1, D1, E1, F1, G1, H1, I1, J1, K1, L1, M1, "
            sql &= "N1, O1, P1, Q1, R1, S1, T1, U1, V1, W1, X1, Y1, Z1  "
            sql &= "From B_CustomerRequest "
            sql &= "Where Buyer = '" & DBuyer.Text & "' "

            If DPO.Text <> "" Then
                sql &= "  And E1 = '" & DPO.Text & "' "
            Else
                If DInputDate.Text <> "" Then
                    sql &= "  And CreateTime >= '" & xStartTime & "' "
                    sql &= "  And CreateTime <= '" & xEndTime & "' "
                Else
                    sql &= " And E1 = '" & DPO.Text & "' "
                End If
            End If
            sql &= "Order by E1, D1 "
        End If
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
        Response.AppendHeader("Content-Disposition", "attachment;filename=CustomerOrder.xls")     '程式別不同
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

End Class
