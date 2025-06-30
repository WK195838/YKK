Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing


Partial Class FASBYCheckSumQty
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim xBuyer As String             'Buyer
    Dim UserID As String            'UserID

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
        Response.Cookies("PGM").Value = "FASBYCheckSumQty.aspx"     '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        xBuyer = Request.QueryString("pBuyer")
        UserID = Request.QueryString("pUserID")             'UserID
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
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowData)
    '**     查詢
    '**
    '*****************************************************************
    Sub ShowData()
        Dim sql As String
        '-- ----------------------------------------------------------------
        '-- W_InputSheetBY 
        '-- 原始 BUYER FCT
        '-- ----------------------------------------------------------------
        sql = "SELECT "
        sql &= "'' as TYPE, "
        sql &= "'' as N_F, "
        sql &= "'' AS N1_F, "
        sql &= "'' AS N2_F, "
        sql &= "'' AS N3_F, "
        sql &= "'' AS N4_F, "
        sql &= "'' AS N5_F, "
        sql &= "'' AS N5_F, "
        sql &= "'' AS N6_F,"
        sql &= "'' AS N7_F, "
        sql &= "'' AS N8_F, "
        sql &= "'' AS N9_F, "
        sql &= "'' AS N10_F, "
        sql &= "'' AS N11_F, "
        sql &= "'' AS N12_F, "
        sql &= "'' AS TTL_F "
        'sql &= "FROM W_InputSheetBY "
        '
        Dim dt_BYFCT As DataTable = uDataBase.GetDataTable(sql)
        If dt_BYFCT.Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = dt_BYFCT
            GridView1.DataBind()
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯表頭
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim gv As GridView = sender
            Dim H1row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-1行
            Dim H1tc1 As TableCell = New TableCell
            H1tc1.Text = ""
            H1row.Cells.Add(H1tc1)

            Dim H1tc2 As TableCell = New TableCell
            H1tc2.Text = "N"
            H1row.Cells.Add(H1tc2)

            Dim H1tc3 As TableCell = New TableCell
            H1tc3.Text = "N+1"
            H1row.Cells.Add(H1tc3)

            Dim H1tc4 As TableCell = New TableCell
            H1tc4.Text = "N+2"
            H1row.Cells.Add(H1tc4)

            Dim H1tc5 As TableCell = New TableCell
            H1tc5.Text = "N+3"
            H1row.Cells.Add(H1tc5)

            Dim H1tc6 As TableCell = New TableCell
            H1tc6.Text = "N+4"
            H1row.Cells.Add(H1tc6)

            Dim H1tc7 As TableCell = New TableCell
            H1tc7.Text = "N+5"
            H1row.Cells.Add(H1tc7)

            Dim H1tc8 As TableCell = New TableCell
            H1tc8.Text = "N+6"
            H1row.Cells.Add(H1tc8)

            Dim H1tc9 As TableCell = New TableCell
            H1tc9.Text = "N+7"
            H1row.Cells.Add(H1tc9)

            Dim H1tcA As TableCell = New TableCell
            H1tcA.Text = "N+8"
            H1row.Cells.Add(H1tcA)

            Dim H1tcB As TableCell = New TableCell
            H1tcB.Text = "N+9"
            H1row.Cells.Add(H1tcB)

            Dim H1tcC As TableCell = New TableCell
            H1tcC.Text = "N+10"
            H1row.Cells.Add(H1tcC)

            Dim H1tcD As TableCell = New TableCell
            H1tcD.Text = "N+11"
            H1row.Cells.Add(H1tcD)

            Dim H1tcE As TableCell = New TableCell
            H1tcE.Text = "N+12"
            H1row.Cells.Add(H1tcE)

            Dim H1tcF As TableCell = New TableCell
            H1tcF.Text = "TOTAL"
            H1row.Cells.Add(H1tcF)

            gv.Controls(0).Controls.AddAt(0, H1row)
            '-----------------------------------------
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            Dim sql As String

            Dim H1row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim H2row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim H5row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim H6row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            ' 清除
            e.Row.Cells.Clear()
            '-- ----------------------------------------------------------------
            '-- W_InputSheetBY 
            '-- 原始 BUYER FCT
            '-- ----------------------------------------------------------------
            Dim H1hd1 As TableCell = New TableCell
            H1hd1.Text = "FCT"
            H1row.Cells.Add(H1hd1)

            sql = "SELECT "
            sql &= "sum(CAST(s1 AS float)) as N, "
            sql &= "sum(CAST(t1 AS float)) AS N1, "
            sql &= "sum(CAST(u1 AS float)) AS N2, "
            sql &= "sum(CAST(v1 AS float)) AS N3, "
            sql &= "sum(CAST(w1 AS float)) AS N4, "
            sql &= "sum(CAST(x1 AS float)) AS N5, "
            sql &= "sum(CAST(y1 AS float)) AS N6,"
            sql &= "sum(CAST(z1 AS float)) AS N7, "
            sql &= "sum(CAST(aa1 AS float)) AS N8, "
            sql &= "sum(CAST(ab1 AS float)) AS N9, "
            sql &= "sum(CAST(ac1 AS float)) AS N10, "
            sql &= "sum(CAST(ad1 AS float)) AS N11, "
            sql &= "sum(CAST(ae1 AS float)) AS N12, "
            sql &= "sum(CAST(af1 AS float)) AS Total "
            sql &= "FROM W_InputSheetBY "
            sql &= " Where Buyer = '" & xBuyer & "' "
            Dim dt_FCTQty As DataTable = uDataBase.GetDataTable(sql)
            If dt_FCTQty.Rows.Count > 0 Then

                Dim H1tc1 As TableCell = New TableCell
                H1tc1.Text = Format(dt_FCTQty.Rows(0).Item("N"), "###,###,##0")
                H1row.Cells.Add(H1tc1)

                Dim H1tc2 As TableCell = New TableCell
                H1tc2.Text = Format(dt_FCTQty.Rows(0).Item("N1"), "###,###,##0")
                H1row.Cells.Add(H1tc2)

                Dim H1tc3 As TableCell = New TableCell
                H1tc3.Text = Format(dt_FCTQty.Rows(0).Item("N2"), "###,###,##0")
                H1row.Cells.Add(H1tc3)

                Dim H1tc4 As TableCell = New TableCell
                H1tc4.Text = Format(dt_FCTQty.Rows(0).Item("N3"), "###,###,##0")
                H1row.Cells.Add(H1tc4)

                Dim H1tc5 As TableCell = New TableCell
                H1tc5.Text = Format(dt_FCTQty.Rows(0).Item("N4"), "###,###,##0")
                H1row.Cells.Add(H1tc5)

                Dim H1tc6 As TableCell = New TableCell
                H1tc6.Text = Format(dt_FCTQty.Rows(0).Item("N5"), "###,###,##0")
                H1row.Cells.Add(H1tc6)

                Dim H1tc7 As TableCell = New TableCell
                H1tc7.Text = Format(dt_FCTQty.Rows(0).Item("N6"), "###,###,##0")
                H1row.Cells.Add(H1tc7)

                Dim H1tc8 As TableCell = New TableCell
                H1tc8.Text = Format(dt_FCTQty.Rows(0).Item("N7"), "###,###,##0")
                H1row.Cells.Add(H1tc8)

                Dim H1tc9 As TableCell = New TableCell
                H1tc9.Text = Format(dt_FCTQty.Rows(0).Item("N8"), "###,###,##0")
                H1row.Cells.Add(H1tc9)

                Dim H1tcA As TableCell = New TableCell
                H1tcA.Text = Format(dt_FCTQty.Rows(0).Item("N9"), "###,###,##0")
                H1row.Cells.Add(H1tcA)

                Dim H1tcB As TableCell = New TableCell
                H1tcB.Text = Format(dt_FCTQty.Rows(0).Item("N10"), "###,###,##0")
                H1row.Cells.Add(H1tcB)

                Dim H1tcC As TableCell = New TableCell
                H1tcC.Text = Format(dt_FCTQty.Rows(0).Item("N11"), "###,###,##0")
                H1row.Cells.Add(H1tcC)

                Dim H1tcD As TableCell = New TableCell
                H1tcD.Text = Format(dt_FCTQty.Rows(0).Item("N12"), "###,###,##0")
                H1row.Cells.Add(H1tcD)

                Dim H1tcE As TableCell = New TableCell
                H1tcE.Text = Format(dt_FCTQty.Rows(0).Item("Total"), "###,###,##0")
                H1row.Cells.Add(H1tcE)
            End If
            e.Row.Parent.Controls.Add(H1row)
            '-- ----------------------------------------------------------------
            '-- E_InputSheetBY
            '-- SUM QTY (原始BUYER FCT)
            '-- ----------------------------------------------------------------
            Dim H2hd1 As TableCell = New TableCell
            H2hd1.Text = "DATA CHECK"
            H2row.Cells.Add(H2hd1)

            sql = "SELECT "
            sql &= "sum(CAST(s1 AS float)) as N, "
            sql &= "sum(CAST(t1 AS float)) AS N1, "
            sql &= "sum(CAST(u1 AS float)) AS N2, "
            sql &= "sum(CAST(v1 AS float)) AS N3, "
            sql &= "sum(CAST(w1 AS float)) AS N4, "
            sql &= "sum(CAST(x1 AS float)) AS N5, "
            sql &= "sum(CAST(y1 AS float)) AS N6,"
            sql &= "sum(CAST(z1 AS float)) AS N7, "
            sql &= "sum(CAST(aa1 AS float)) AS N8, "
            sql &= "sum(CAST(ab1 AS float)) AS N9, "
            sql &= "sum(CAST(ac1 AS float)) AS N10, "
            sql &= "sum(CAST(ad1 AS float)) AS N11, "
            sql &= "sum(CAST(ae1 AS float)) AS N12, "
            sql &= "sum(CAST(af1 AS float)) AS Total "
            sql &= "FROM E_InputSheetBY "
            sql &= " Where Buyer = '" & xBuyer & "' "
            Dim dt_FCTQty1 As DataTable = uDataBase.GetDataTable(sql)
            If dt_FCTQty1.Rows.Count > 0 Then

                Dim H2tc1 As TableCell = New TableCell
                H2tc1.Text = Format(dt_FCTQty1.Rows(0).Item("N"), "###,###,##0")
                H2row.Cells.Add(H2tc1)

                Dim H2tc2 As TableCell = New TableCell
                H2tc2.Text = Format(dt_FCTQty1.Rows(0).Item("N1"), "###,###,##0")
                H2row.Cells.Add(H2tc2)

                Dim H2tc3 As TableCell = New TableCell
                H2tc3.Text = Format(dt_FCTQty1.Rows(0).Item("N2"), "###,###,##0")
                H2row.Cells.Add(H2tc3)

                Dim H2tc4 As TableCell = New TableCell
                H2tc4.Text = Format(dt_FCTQty1.Rows(0).Item("N3"), "###,###,##0")
                H2row.Cells.Add(H2tc4)

                Dim H2tc5 As TableCell = New TableCell
                H2tc5.Text = Format(dt_FCTQty1.Rows(0).Item("N4"), "###,###,##0")
                H2row.Cells.Add(H2tc5)

                Dim H2tc6 As TableCell = New TableCell
                H2tc6.Text = Format(dt_FCTQty1.Rows(0).Item("N5"), "###,###,##0")
                H2row.Cells.Add(H2tc6)

                Dim H2tc7 As TableCell = New TableCell
                H2tc7.Text = Format(dt_FCTQty1.Rows(0).Item("N6"), "###,###,##0")
                H2row.Cells.Add(H2tc7)

                Dim H2tc8 As TableCell = New TableCell
                H2tc8.Text = Format(dt_FCTQty1.Rows(0).Item("N7"), "###,###,##0")
                H2row.Cells.Add(H2tc8)

                Dim H2tc9 As TableCell = New TableCell
                H2tc9.Text = Format(dt_FCTQty1.Rows(0).Item("N8"), "###,###,##0")
                H2row.Cells.Add(H2tc9)

                Dim H2tcA As TableCell = New TableCell
                H2tcA.Text = Format(dt_FCTQty1.Rows(0).Item("N9"), "###,###,##0")
                H2row.Cells.Add(H2tcA)

                Dim H2tcB As TableCell = New TableCell
                H2tcB.Text = Format(dt_FCTQty1.Rows(0).Item("N10"), "###,###,##0")
                H2row.Cells.Add(H2tcB)

                Dim H2tcC As TableCell = New TableCell
                H2tcC.Text = Format(dt_FCTQty1.Rows(0).Item("N11"), "###,###,##0")
                H2row.Cells.Add(H2tcC)

                Dim H2tcD As TableCell = New TableCell
                H2tcD.Text = Format(dt_FCTQty1.Rows(0).Item("N12"), "###,###,##0")
                H2row.Cells.Add(H2tcD)

                Dim H2tcE As TableCell = New TableCell
                H2tcE.Text = Format(dt_FCTQty1.Rows(0).Item("Total"), "###,###,##0")
                H2row.Cells.Add(H2tcE)
            End If
            e.Row.Parent.Controls.Add(H2row)
            '-- ----------------------------------------------------------------
            '-- 差異
            '-- 
            '-- ----------------------------------------------------------------
            Dim H3hd1 As TableCell = New TableCell
            H3hd1.BackColor = Color.LightPink
            H3hd1.Text = "差"
            H3row.Cells.Add(H3hd1)

            Dim H3tc1 As TableCell = New TableCell
            H3tc1.HorizontalAlign = HorizontalAlign.Right
            H3tc1.BackColor = Color.LightPink
            H3tc1.Text = Format(dt_FCTQty.Rows(0).Item("N") - dt_FCTQty1.Rows(0).Item("N"), "###,###,##0")
            H3row.Cells.Add(H3tc1)

            Dim H3tc2 As TableCell = New TableCell
            H3tc2.HorizontalAlign = HorizontalAlign.Right
            H3tc2.BackColor = Color.LightPink
            H3tc2.Text = Format(dt_FCTQty.Rows(0).Item("N1") - dt_FCTQty1.Rows(0).Item("N1"), "###,###,##0")
            H3row.Cells.Add(H3tc2)

            Dim H3tc3 As TableCell = New TableCell
            H3tc3.HorizontalAlign = HorizontalAlign.Right
            H3tc3.BackColor = Color.LightPink
            H3tc3.Text = Format(dt_FCTQty.Rows(0).Item("N2") - dt_FCTQty1.Rows(0).Item("N2"), "###,###,##0")
            H3row.Cells.Add(H3tc3)

            Dim H3tc4 As TableCell = New TableCell
            H3tc4.HorizontalAlign = HorizontalAlign.Right
            H3tc4.BackColor = Color.LightPink
            H3tc4.Text = Format(dt_FCTQty.Rows(0).Item("N3") - dt_FCTQty1.Rows(0).Item("N3"), "###,###,##0")
            H3row.Cells.Add(H3tc4)

            Dim H3tc5 As TableCell = New TableCell
            H3tc5.HorizontalAlign = HorizontalAlign.Right
            H3tc5.BackColor = Color.LightPink
            H3tc5.Text = Format(dt_FCTQty.Rows(0).Item("N4") - dt_FCTQty1.Rows(0).Item("N4"), "###,###,##0")
            H3row.Cells.Add(H3tc5)

            Dim H3tc6 As TableCell = New TableCell
            H3tc6.HorizontalAlign = HorizontalAlign.Right
            H3tc6.BackColor = Color.LightPink
            H3tc6.Text = Format(dt_FCTQty.Rows(0).Item("N5") - dt_FCTQty1.Rows(0).Item("N5"), "###,###,##0")
            H3row.Cells.Add(H3tc6)

            Dim H3tc7 As TableCell = New TableCell
            H3tc7.HorizontalAlign = HorizontalAlign.Right
            H3tc7.BackColor = Color.LightPink
            H3tc7.Text = Format(dt_FCTQty.Rows(0).Item("N6") - dt_FCTQty1.Rows(0).Item("N6"), "###,###,##0")
            H3row.Cells.Add(H3tc7)

            Dim H3tc8 As TableCell = New TableCell
            H3tc8.HorizontalAlign = HorizontalAlign.Right
            H3tc8.BackColor = Color.LightPink
            H3tc8.Text = Format(dt_FCTQty.Rows(0).Item("N7") - dt_FCTQty1.Rows(0).Item("N7"), "###,###,##0")
            H3row.Cells.Add(H3tc8)

            Dim H3tc9 As TableCell = New TableCell
            H3tc9.HorizontalAlign = HorizontalAlign.Right
            H3tc9.BackColor = Color.LightPink
            H3tc9.Text = Format(dt_FCTQty.Rows(0).Item("N8") - dt_FCTQty1.Rows(0).Item("N8"), "###,###,##0")
            H3row.Cells.Add(H3tc9)

            Dim H3tcA As TableCell = New TableCell
            H3tcA.HorizontalAlign = HorizontalAlign.Right
            H3tcA.BackColor = Color.LightPink
            H3tcA.Text = Format(dt_FCTQty.Rows(0).Item("N9") - dt_FCTQty1.Rows(0).Item("N9"), "###,###,##0")
            H3row.Cells.Add(H3tcA)

            Dim H3tcB As TableCell = New TableCell
            H3tcB.HorizontalAlign = HorizontalAlign.Right
            H3tcB.BackColor = Color.LightPink
            H3tcB.Text = Format(dt_FCTQty.Rows(0).Item("N10") - dt_FCTQty1.Rows(0).Item("N10"), "###,###,##0")
            H3row.Cells.Add(H3tcB)

            Dim H3tcC As TableCell = New TableCell
            H3tcC.HorizontalAlign = HorizontalAlign.Right
            H3tcC.BackColor = Color.LightPink
            H3tcC.Text = Format(dt_FCTQty.Rows(0).Item("N11") - dt_FCTQty1.Rows(0).Item("N11"), "###,###,##0")
            H3row.Cells.Add(H3tcC)

            Dim H3tcD As TableCell = New TableCell
            H3tcD.HorizontalAlign = HorizontalAlign.Right
            H3tcD.BackColor = Color.LightPink
            H3tcD.Text = Format(dt_FCTQty.Rows(0).Item("N12") - dt_FCTQty1.Rows(0).Item("N12"), "###,###,##0")
            H3row.Cells.Add(H3tcD)

            Dim H3tcE As TableCell = New TableCell
            H3tcE.HorizontalAlign = HorizontalAlign.Right
            H3tcE.BackColor = Color.LightPink
            H3tcE.Text = Format(dt_FCTQty.Rows(0).Item("Total") - dt_FCTQty1.Rows(0).Item("Total"), "###,###,##0")
            H3row.Cells.Add(H3tcE)
            e.Row.Parent.Controls.Add(H3row)
            '-- ----------------------------------------------------------------
            '-- E_InputSheetBY
            '-- DATA STATUS < 10
            '-- ----------------------------------------------------------------
            Dim H4hd1 As TableCell = New TableCell
            H4hd1.Text = "DATA CHECK STATUS < 10"
            H4row.Cells.Add(H4hd1)

            sql = "SELECT "
            sql &= "sum(CAST(s1 AS float)) as N, "
            sql &= "sum(CAST(t1 AS float)) AS N1, "
            sql &= "sum(CAST(u1 AS float)) AS N2, "
            sql &= "sum(CAST(v1 AS float)) AS N3, "
            sql &= "sum(CAST(w1 AS float)) AS N4, "
            sql &= "sum(CAST(x1 AS float)) AS N5, "
            sql &= "sum(CAST(y1 AS float)) AS N6,"
            sql &= "sum(CAST(z1 AS float)) AS N7, "
            sql &= "sum(CAST(aa1 AS float)) AS N8, "
            sql &= "sum(CAST(ab1 AS float)) AS N9, "
            sql &= "sum(CAST(ac1 AS float)) AS N10, "
            sql &= "sum(CAST(ad1 AS float)) AS N11, "
            sql &= "sum(CAST(ae1 AS float)) AS N12, "
            sql &= "sum(CAST(af1 AS float)) AS Total "
            sql &= "FROM E_InputSheetBY "
            sql &= "WHERE EZ1 < '10' "
            sql &= "  AND Buyer = '" & xBuyer & "' "
            Dim dt_FCTQty2 As DataTable = uDataBase.GetDataTable(sql)
            If dt_FCTQty2.Rows.Count > 0 Then

                Dim H4tc1 As TableCell = New TableCell
                H4tc1.Text = Format(dt_FCTQty2.Rows(0).Item("N"), "###,###,##0")
                H4row.Cells.Add(H4tc1)

                Dim H4tc2 As TableCell = New TableCell
                H4tc2.Text = Format(dt_FCTQty2.Rows(0).Item("N1"), "###,###,##0")
                H4row.Cells.Add(H4tc2)

                Dim H4tc3 As TableCell = New TableCell
                H4tc3.Text = Format(dt_FCTQty2.Rows(0).Item("N2"), "###,###,##0")
                H4row.Cells.Add(H4tc3)

                Dim H4tc4 As TableCell = New TableCell
                H4tc4.Text = Format(dt_FCTQty2.Rows(0).Item("N3"), "###,###,##0")
                H4row.Cells.Add(H4tc4)

                Dim H4tc5 As TableCell = New TableCell
                H4tc5.Text = Format(dt_FCTQty2.Rows(0).Item("N4"), "###,###,##0")
                H4row.Cells.Add(H4tc5)

                Dim H4tc6 As TableCell = New TableCell
                H4tc6.Text = Format(dt_FCTQty2.Rows(0).Item("N5"), "###,###,##0")
                H4row.Cells.Add(H4tc6)

                Dim H4tc7 As TableCell = New TableCell
                H4tc7.Text = Format(dt_FCTQty2.Rows(0).Item("N6"), "###,###,##0")
                H4row.Cells.Add(H4tc7)

                Dim H4tc8 As TableCell = New TableCell
                H4tc8.Text = Format(dt_FCTQty2.Rows(0).Item("N7"), "###,###,##0")
                H4row.Cells.Add(H4tc8)

                Dim H4tc9 As TableCell = New TableCell
                H4tc9.Text = Format(dt_FCTQty2.Rows(0).Item("N8"), "###,###,##0")
                H4row.Cells.Add(H4tc9)

                Dim H4tcA As TableCell = New TableCell
                H4tcA.Text = Format(dt_FCTQty2.Rows(0).Item("N9"), "###,###,##0")
                H4row.Cells.Add(H4tcA)

                Dim H4tcB As TableCell = New TableCell
                H4tcB.Text = Format(dt_FCTQty2.Rows(0).Item("N10"), "###,###,##0")
                H4row.Cells.Add(H4tcB)

                Dim H4tcC As TableCell = New TableCell
                H4tcC.Text = Format(dt_FCTQty2.Rows(0).Item("N11"), "###,###,##0")
                H4row.Cells.Add(H4tcC)

                Dim H4tcD As TableCell = New TableCell
                H4tcD.Text = Format(dt_FCTQty2.Rows(0).Item("N12"), "###,###,##0")
                H4row.Cells.Add(H4tcD)

                Dim H4tcE As TableCell = New TableCell
                H4tcE.Text = Format(dt_FCTQty2.Rows(0).Item("Total"), "###,###,##0")
                H4row.Cells.Add(H4tcE)
            End If
            e.Row.Parent.Controls.Add(H4row)
            '-- ----------------------------------------------------------------
            '-- E_BYFCTSheet
            '-- BUYER FCT(FAS使用)
            '-- ----------------------------------------------------------------
            Dim H5hd1 As TableCell = New TableCell
            H5hd1.Text = "FAS FCT"
            H5row.Cells.Add(H5hd1)

            sql = "SELECT "
            sql &= "sum(N_F) as N, "
            sql &= "sum(N1_F) as N1, "
            sql &= "sum(N2_F) as N2, "
            sql &= "sum(N3_F) as N3, "
            sql &= "sum(N4_F) as N4, "
            sql &= "sum(N5_F) as N5, "
            sql &= "sum(N6_F) as N6, "
            sql &= "sum(N7_F) as N7, "
            sql &= "sum(N8_F) as N8, "
            sql &= "sum(N9_F) as N9, "
            sql &= "sum(N10_F) as N10, "
            sql &= "sum(N11_F) as N11, "
            sql &= "sum(N12_F) as N12, "
            sql &= "sum(TOTAL) as Total "
            sql &= "FROM E_BYFCTSheet "
            sql &= " Where Buyer = '" & xBuyer & "' "
            Dim dt_FCTQty3 As DataTable = uDataBase.GetDataTable(sql)
            If dt_FCTQty3.Rows.Count > 0 Then

                Dim H5tc1 As TableCell = New TableCell
                H5tc1.Text = Format(dt_FCTQty3.Rows(0).Item("N"), "###,###,##0")
                H5row.Cells.Add(H5tc1)

                Dim H5tc2 As TableCell = New TableCell
                H5tc2.Text = Format(dt_FCTQty3.Rows(0).Item("N1"), "###,###,##0")
                H5row.Cells.Add(H5tc2)

                Dim H5tc3 As TableCell = New TableCell
                H5tc3.Text = Format(dt_FCTQty3.Rows(0).Item("N2"), "###,###,##0")
                H5row.Cells.Add(H5tc3)

                Dim H5tc4 As TableCell = New TableCell
                H5tc4.Text = Format(dt_FCTQty3.Rows(0).Item("N3"), "###,###,##0")
                H5row.Cells.Add(H5tc4)

                Dim H5tc5 As TableCell = New TableCell
                H5tc5.Text = Format(dt_FCTQty3.Rows(0).Item("N4"), "###,###,##0")
                H5row.Cells.Add(H5tc5)

                Dim H5tc6 As TableCell = New TableCell
                H5tc6.Text = Format(dt_FCTQty3.Rows(0).Item("N5"), "###,###,##0")
                H5row.Cells.Add(H5tc6)

                Dim H5tc7 As TableCell = New TableCell
                H5tc7.Text = Format(dt_FCTQty3.Rows(0).Item("N6"), "###,###,##0")
                H5row.Cells.Add(H5tc7)

                Dim H5tc8 As TableCell = New TableCell
                H5tc8.Text = Format(dt_FCTQty3.Rows(0).Item("N7"), "###,###,##0")
                H5row.Cells.Add(H5tc8)

                Dim H5tc9 As TableCell = New TableCell
                H5tc9.Text = Format(dt_FCTQty3.Rows(0).Item("N8"), "###,###,##0")
                H5row.Cells.Add(H5tc9)

                Dim H5tcA As TableCell = New TableCell
                H5tcA.Text = Format(dt_FCTQty3.Rows(0).Item("N9"), "###,###,##0")
                H5row.Cells.Add(H5tcA)

                Dim H5tcB As TableCell = New TableCell
                H5tcB.Text = Format(dt_FCTQty3.Rows(0).Item("N10"), "###,###,##0")
                H5row.Cells.Add(H5tcB)

                Dim H5tcC As TableCell = New TableCell
                H5tcC.Text = Format(dt_FCTQty3.Rows(0).Item("N11"), "###,###,##0")
                H5row.Cells.Add(H5tcC)

                Dim H5tcD As TableCell = New TableCell
                H5tcD.Text = Format(dt_FCTQty3.Rows(0).Item("N12"), "###,###,##0")
                H5row.Cells.Add(H5tcD)

                Dim H5tcE As TableCell = New TableCell
                H5tcE.Text = Format(dt_FCTQty3.Rows(0).Item("Total"), "###,###,##0")
                H5row.Cells.Add(H5tcE)
            End If
            e.Row.Parent.Controls.Add(H5row)
            '-- ----------------------------------------------------------------
            '-- 差異
            '-- 
            '-- ----------------------------------------------------------------
            Dim H6hd1 As TableCell = New TableCell
            H6hd1.BackColor = Color.LightPink
            H6hd1.Text = "差"
            H6row.Cells.Add(H6hd1)

            Dim H6tc1 As TableCell = New TableCell
            H6tc1.HorizontalAlign = HorizontalAlign.Right
            H6tc1.BackColor = Color.LightPink
            H6tc1.Text = Format(dt_FCTQty2.Rows(0).Item("N") - dt_FCTQty3.Rows(0).Item("N"), "###,###,##0")
            H6row.Cells.Add(H6tc1)

            Dim H6tc2 As TableCell = New TableCell
            H6tc2.HorizontalAlign = HorizontalAlign.Right
            H6tc2.BackColor = Color.LightPink
            H6tc2.Text = Format(dt_FCTQty2.Rows(0).Item("N1") - dt_FCTQty3.Rows(0).Item("N1"), "###,###,##0")
            H6row.Cells.Add(H6tc2)

            Dim H6tc3 As TableCell = New TableCell
            H6tc3.HorizontalAlign = HorizontalAlign.Right
            H6tc3.BackColor = Color.LightPink
            H6tc3.Text = Format(dt_FCTQty2.Rows(0).Item("N2") - dt_FCTQty3.Rows(0).Item("N2"), "###,###,##0")
            H6row.Cells.Add(H6tc3)

            Dim H6tc4 As TableCell = New TableCell
            H6tc4.HorizontalAlign = HorizontalAlign.Right
            H6tc4.BackColor = Color.LightPink
            H6tc4.Text = Format(dt_FCTQty2.Rows(0).Item("N3") - dt_FCTQty3.Rows(0).Item("N3"), "###,###,##0")
            H6row.Cells.Add(H6tc4)

            Dim H6tc5 As TableCell = New TableCell
            H6tc5.HorizontalAlign = HorizontalAlign.Right
            H6tc5.BackColor = Color.LightPink
            H6tc5.Text = Format(dt_FCTQty2.Rows(0).Item("N4") - dt_FCTQty3.Rows(0).Item("N4"), "###,###,##0")
            H6row.Cells.Add(H6tc5)

            Dim H6tc6 As TableCell = New TableCell
            H6tc6.HorizontalAlign = HorizontalAlign.Right
            H6tc6.BackColor = Color.LightPink
            H6tc6.Text = Format(dt_FCTQty2.Rows(0).Item("N5") - dt_FCTQty3.Rows(0).Item("N5"), "###,###,##0")
            H6row.Cells.Add(H6tc6)

            Dim H6tc7 As TableCell = New TableCell
            H6tc7.HorizontalAlign = HorizontalAlign.Right
            H6tc7.BackColor = Color.LightPink
            H6tc7.Text = Format(dt_FCTQty2.Rows(0).Item("N6") - dt_FCTQty3.Rows(0).Item("N6"), "###,###,##0")
            H6row.Cells.Add(H6tc7)

            Dim H6tc8 As TableCell = New TableCell
            H6tc8.HorizontalAlign = HorizontalAlign.Right
            H6tc8.BackColor = Color.LightPink
            H6tc8.Text = Format(dt_FCTQty2.Rows(0).Item("N7") - dt_FCTQty3.Rows(0).Item("N7"), "###,###,##0")
            H6row.Cells.Add(H6tc8)

            Dim H6tc9 As TableCell = New TableCell
            H6tc9.HorizontalAlign = HorizontalAlign.Right
            H6tc9.BackColor = Color.LightPink
            H6tc9.Text = Format(dt_FCTQty2.Rows(0).Item("N8") - dt_FCTQty3.Rows(0).Item("N8"), "###,###,##0")
            H6row.Cells.Add(H6tc9)

            Dim H6tcA As TableCell = New TableCell
            H6tcA.HorizontalAlign = HorizontalAlign.Right
            H6tcA.BackColor = Color.LightPink
            H6tcA.Text = Format(dt_FCTQty2.Rows(0).Item("N9") - dt_FCTQty3.Rows(0).Item("N9"), "###,###,##0")
            H6row.Cells.Add(H6tcA)

            Dim H6tcB As TableCell = New TableCell
            H6tcB.HorizontalAlign = HorizontalAlign.Right
            H6tcB.BackColor = Color.LightPink
            H6tcB.Text = Format(dt_FCTQty2.Rows(0).Item("N10") - dt_FCTQty3.Rows(0).Item("N10"), "###,###,##0")
            H6row.Cells.Add(H6tcB)

            Dim H6tcC As TableCell = New TableCell
            H6tcC.HorizontalAlign = HorizontalAlign.Right
            H6tcC.BackColor = Color.LightPink
            H6tcC.Text = Format(dt_FCTQty2.Rows(0).Item("N11") - dt_FCTQty3.Rows(0).Item("N11"), "###,###,##0")
            H6row.Cells.Add(H6tcC)

            Dim H6tcD As TableCell = New TableCell
            H6tcD.HorizontalAlign = HorizontalAlign.Right
            H6tcD.BackColor = Color.LightPink
            H6tcD.Text = Format(dt_FCTQty2.Rows(0).Item("N12") - dt_FCTQty3.Rows(0).Item("N12"), "###,###,##0")
            H6row.Cells.Add(H6tcD)

            Dim H6tcE As TableCell = New TableCell
            H6tcE.HorizontalAlign = HorizontalAlign.Right
            H6tcE.BackColor = Color.LightPink
            H6tcE.Text = Format(dt_FCTQty2.Rows(0).Item("Total") - dt_FCTQty3.Rows(0).Item("Total"), "###,###,##0")
            H6row.Cells.Add(H6tcE)
            e.Row.Parent.Controls.Add(H6row)
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then

        End If
    End Sub

End Class
