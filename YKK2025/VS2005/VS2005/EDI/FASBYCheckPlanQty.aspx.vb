Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing


Partial Class FASBYCheckPlanQty
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim xBuyer As String            'Buyer
    Dim xBuyerGroup As String       'BuyerGroup
    Dim UserID As String            'UserID
    Dim xLogID As String            'LogID
    Dim xFun As String              'Fun

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
        Response.Cookies("PGM").Value = "FASBYCheckPlanQty.aspx"    '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        xBuyer = Request.QueryString("pBuyer")
        xBuyerGroup = Request.QueryString("pBuyerGroup")
        UserID = Request.QueryString("pUserID")             'UserID
        xFun = Request.QueryString("pFun")                  'Fun
        xLogID = Request.QueryString("pLogID")              'LogID
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

            Dim H11row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim H21row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim H31row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)

            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim H5row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim H6row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim H7row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim H8row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim H9row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)

            Dim HArow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim HBrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim HCrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim HDrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim HErow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim HFrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)

            Dim HGrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim HHrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim HIrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim HJrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim HKrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim HLrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)

            Dim HMrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim HNrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim HOrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim HProw As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim HQrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim HRrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)

            ' 清除
            e.Row.Cells.Clear()
            DMessage.Text = ""
            '-- ----------------------------------------------------------------
            '-- 客戶 FC (CONVERT)
            '-- ----------------------------------------------------------------
            If xFun = "CONVERT" Or xFun = "FCTPLAN" Or xFun = "LSPLAN" Or xFun = "IPLSPlan" Or xFun = "BYLSPLAN" Or xFun = "EDI" Then
                Dim H1hd1 As TableCell = New TableCell
                H1hd1.Text = "客戶 FC"
                H1row.Cells.Add(H1hd1)

                sql = "SELECT "
                If xBuyer = "FALL-000021" Then
                    sql &= "ISNULL(SUM(CAST(J1 AS float)),0) as N, "
                    sql &= "ISNULL(SUM(CAST(K1 AS float)),0) AS N1, "
                    sql &= "ISNULL(SUM(CAST(L1 AS float)),0) AS N2, "
                    sql &= "ISNULL(SUM(CAST(M1 AS float)),0) AS N3, "
                    sql &= "ISNULL(SUM(CAST(N1 AS float)),0) AS N4, "
                    sql &= "ISNULL(SUM(CAST(O1 AS float)),0) AS N5, "
                    sql &= "ISNULL(SUM(CAST(P1 AS float)),0) AS N6,"
                    sql &= "ISNULL(SUM(CAST(Q1 AS float)),0) AS N7, "
                    sql &= "ISNULL(SUM(CAST(R1 AS float)),0) AS N8, "
                    sql &= "ISNULL(SUM(CAST(S1 AS float)),0) AS N9, "
                    sql &= "ISNULL(SUM(CAST(T1 AS float)),0) AS N10, "
                    sql &= "ISNULL(SUM(CAST(U1 AS float)),0) AS N11, "
                    sql &= "ISNULL(SUM(CAST(V1 AS float)),0) AS N12, "
                    sql &= "ISNULL(SUM(CAST(W1 AS float)),0) AS Total "
                Else
                    If xBuyer = "FALL-TW0371" Or xBuyer = "FALL-TW1741" Then
                        sql &= "ISNULL(SUM(CAST(I1 AS float)),0) as N, "
                        sql &= "ISNULL(SUM(CAST(J1 AS float)),0) AS N1, "
                        sql &= "ISNULL(SUM(CAST(K1 AS float)),0) AS N2, "
                        sql &= "ISNULL(SUM(CAST(L1 AS float)),0) AS N3, "
                        sql &= "ISNULL(SUM(CAST(M1 AS float)),0) AS N4, "
                        sql &= "ISNULL(SUM(CAST(N1 AS float)),0) AS N5, "
                        sql &= "ISNULL(SUM(CAST(O1 AS float)),0) AS N6,"
                        sql &= "ISNULL(SUM(CAST(P1 AS float)),0) AS N7, "
                        sql &= "ISNULL(SUM(CAST(Q1 AS float)),0) AS N8, "
                        sql &= "ISNULL(SUM(CAST(R1 AS float)),0) AS N9, "
                        sql &= "ISNULL(SUM(CAST(S1 AS float)),0) AS N10, "
                        sql &= "ISNULL(SUM(CAST(T1 AS float)),0) AS N11, "
                        sql &= "ISNULL(SUM(CAST(U1 AS float)),0) AS N12, "
                        sql &= "ISNULL(SUM(CAST(V1 AS float)),0) AS Total "
                    Else
                        If xBuyer = "FALL-TW0371T" Then
                            sql &= "ISNULL(SUM(CAST(E1 AS float)),0) as N, "
                            sql &= "ISNULL(SUM(CAST(F1 AS float)),0) AS N1, "
                            sql &= "ISNULL(SUM(CAST(G1 AS float)),0) AS N2, "
                            sql &= "ISNULL(SUM(CAST(H1 AS float)),0) AS N3, "
                            sql &= "ISNULL(SUM(CAST(I1 AS float)),0) AS N4, "
                            sql &= "ISNULL(SUM(CAST(J1 AS float)),0) AS N5, "
                            sql &= "ISNULL(SUM(CAST(K1 AS float)),0) AS N6,"
                            sql &= "ISNULL(SUM(CAST(L1 AS float)),0) AS N7, "
                            sql &= "ISNULL(SUM(CAST(M1 AS float)),0) AS N8, "
                            sql &= "ISNULL(SUM(CAST(N1 AS float)),0) AS N9, "
                            sql &= "ISNULL(SUM(CAST(O1 AS float)),0) AS N10, "
                            sql &= "ISNULL(SUM(CAST(P1 AS float)),0) AS N11, "
                            sql &= "ISNULL(SUM(CAST(Q1 AS float)),0) AS N12, "
                            sql &= "ISNULL(SUM(CAST(R1 AS float)),0) AS Total "
                        Else
                            If xBuyer = "FALL-TP000013" Then
                                sql &= "ISNULL(SUM(CAST(H1 AS float)),0) as N, "
                                sql &= "ISNULL(SUM(CAST(I1 AS float)),0) as N1, "
                                sql &= "ISNULL(SUM(CAST(J1 AS float)),0) AS N2, "
                                sql &= "ISNULL(SUM(CAST(K1 AS float)),0) AS N3, "
                                sql &= "ISNULL(SUM(CAST(L1 AS float)),0) AS N4, "
                                sql &= "ISNULL(SUM(CAST(M1 AS float)),0) AS N5, "
                                sql &= "ISNULL(SUM(CAST(N1 AS float)),0) AS N6, "
                                sql &= "ISNULL(SUM(CAST(O1 AS float)),0) AS N7,"
                                sql &= "ISNULL(SUM(CAST(P1 AS float)),0) AS N8, "
                                sql &= "ISNULL(SUM(CAST(Q1 AS float)),0) AS N9, "
                                sql &= "ISNULL(SUM(CAST(R1 AS float)),0) AS N10, "
                                sql &= "ISNULL(SUM(CAST(S1 AS float)),0) AS N11, "
                                sql &= "ISNULL(SUM(CAST(T1 AS float)),0) AS N12, "
                                sql &= "ISNULL(SUM(CAST(U1 AS float)),0) AS Total "
                            Else
                                If xBuyer = "FALL-VENDOR" Or InStr(xBuyer, "F-VENDOR") > 0 Then
                                    sql &= "ISNULL(SUM(CAST(V1 AS float)),0) as N, "
                                    sql &= "ISNULL(SUM(CAST(W1 AS float)),0) AS N1, "
                                    sql &= "ISNULL(SUM(CAST(X1 AS float)),0) AS N2, "
                                    sql &= "ISNULL(SUM(CAST(Y1 AS float)),0) AS N3, "
                                    sql &= "ISNULL(SUM(CAST(Z1 AS float)),0) AS N4, "
                                    sql &= "ISNULL(SUM(CAST(AA1 AS float)),0) AS N5, "
                                    sql &= "0 AS N6,"
                                    sql &= "0 AS N7, "
                                    sql &= "0 AS N8, "
                                    sql &= "0 AS N9, "
                                    sql &= "0 AS N10, "
                                    sql &= "0 AS N11, "
                                    sql &= "0 AS N12, "
                                    sql &= "ISNULL(SUM(CAST(AB1 AS float)),0) AS Total "
                                Else
                                    If xBuyer = "FALL-TW0655" Then
                                        sql &= "ISNULL(SUM(CAST(G1 AS float)),0) as N, "
                                        sql &= "ISNULL(SUM(CAST(H1 AS float)),0) as N1, "
                                        sql &= "ISNULL(SUM(CAST(I1 AS float)),0) as N2, "
                                        sql &= "ISNULL(SUM(CAST(J1 AS float)),0) AS N3, "
                                        sql &= "ISNULL(SUM(CAST(K1 AS float)),0) AS N4, "
                                        sql &= "ISNULL(SUM(CAST(L1 AS float)),0) AS N5, "
                                        sql &= "ISNULL(SUM(CAST(M1 AS float)),0) AS N6, "
                                        sql &= "ISNULL(SUM(CAST(N1 AS float)),0) AS N7, "
                                        sql &= "ISNULL(SUM(CAST(O1 AS float)),0) AS N8,"
                                        sql &= "ISNULL(SUM(CAST(P1 AS float)),0) AS N9, "
                                        sql &= "ISNULL(SUM(CAST(Q1 AS float)),0) AS N10, "
                                        sql &= "ISNULL(SUM(CAST(R1 AS float)),0) AS N11, "
                                        sql &= "ISNULL(SUM(CAST(S1 AS float)),0) AS N12, "
                                        sql &= "ISNULL(SUM(CAST(T1 AS float)),0) AS Total "
                                    Else
                                        If xBuyer = "FALL-000053" Then
                                            sql &= "ISNULL(SUM(CAST(J1 AS float)),0) as N, "
                                            sql &= "ISNULL(SUM(CAST(K1 AS float)),0) as N1, "
                                            sql &= "ISNULL(SUM(CAST(L1 AS float)),0) as N2, "
                                            sql &= "ISNULL(SUM(CAST(M1 AS float)),0) AS N3, "
                                            sql &= "ISNULL(SUM(CAST(N1 AS float)),0) AS N4, "
                                            sql &= "ISNULL(SUM(CAST(O1 AS float)),0) AS N5, "
                                            sql &= "ISNULL(SUM(CAST(P1 AS float)),0) AS N6, "
                                            sql &= "ISNULL(SUM(CAST(Q1 AS float)),0) AS N7, "
                                            sql &= "ISNULL(SUM(CAST(R1 AS float)),0) AS N8,"
                                            sql &= "ISNULL(SUM(CAST(S1 AS float)),0) AS N9, "
                                            sql &= "ISNULL(SUM(CAST(T1 AS float)),0) AS N10, "
                                            sql &= "ISNULL(SUM(CAST(U1 AS float)),0) AS N11, "
                                            sql &= "ISNULL(SUM(CAST(V1 AS float)),0) AS N12, "
                                            sql &= "ISNULL(SUM(CAST(W1 AS float)),0) AS Total "
                                        Else
                                            If xBuyer = "FALL-000141" Then
                                                sql &= "ISNULL(SUM(CAST(G1 AS float)),0) as N, "
                                                sql &= "ISNULL(SUM(CAST(H1 AS float)),0) as N1, "
                                                sql &= "ISNULL(SUM(CAST(I1 AS float)),0) as N2, "
                                                sql &= "ISNULL(SUM(CAST(J1 AS float)),0) AS N3, "
                                                sql &= "ISNULL(SUM(CAST(K1 AS float)),0) AS N4, "
                                                sql &= "ISNULL(SUM(CAST(L1 AS float)),0) AS N5, "
                                                sql &= "ISNULL(SUM(CAST(M1 AS float)),0) AS N6, "
                                                sql &= "ISNULL(SUM(CAST(N1 AS float)),0) AS N7, "
                                                sql &= "ISNULL(SUM(CAST(O1 AS float)),0) AS N8,"
                                                sql &= "ISNULL(SUM(CAST(P1 AS float)),0) AS N9, "
                                                sql &= "ISNULL(SUM(CAST(Q1 AS float)),0) AS N10, "
                                                sql &= "ISNULL(SUM(CAST(R1 AS float)),0) AS N11, "
                                                sql &= "ISNULL(SUM(CAST(S1 AS float)),0) AS N12, "
                                                sql &= "ISNULL(SUM(CAST(T1 AS float)),0) AS Total "
                                            Else
                                                sql &= "ISNULL(SUM(CAST(M1 AS float)),0) as N, "
                                                sql &= "ISNULL(SUM(CAST(N1 AS float)),0) AS N1, "
                                                sql &= "ISNULL(SUM(CAST(O1 AS float)),0) AS N2, "
                                                sql &= "ISNULL(SUM(CAST(P1 AS float)),0) AS N3, "
                                                sql &= "ISNULL(SUM(CAST(Q1 AS float)),0) AS N4, "
                                                sql &= "ISNULL(SUM(CAST(R1 AS float)),0) AS N5, "
                                                sql &= "ISNULL(SUM(CAST(S1 AS float)),0) AS N6,"
                                                sql &= "ISNULL(SUM(CAST(T1 AS float)),0) AS N7, "
                                                sql &= "ISNULL(SUM(CAST(U1 AS float)),0) AS N8, "
                                                sql &= "ISNULL(SUM(CAST(V1 AS float)),0) AS N9, "
                                                sql &= "ISNULL(SUM(CAST(W1 AS float)),0) AS N10, "
                                                sql &= "ISNULL(SUM(CAST(X1 AS float)),0) AS N11, "
                                                sql &= "ISNULL(SUM(CAST(Y1 AS float)),0) AS N12, "
                                                sql &= "ISNULL(SUM(CAST(Z1 AS float)),0) AS Total "
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                sql &= "FROM E_InputSheet "
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
                '-- PLAN ZIP QTY
                '-- ----------------------------------------------------------------
                Dim H2hd1 As TableCell = New TableCell
                H2hd1.Text = "FCPLAN"
                H2row.Cells.Add(H2hd1)

                sql = "SELECT "
                sql &= "ISNULL(SUM(CAST(N_F AS float)),0) as N, "
                sql &= "ISNULL(SUM(CAST(N1_F AS float)),0) AS N1, "
                sql &= "ISNULL(SUM(CAST(N2_F AS float)),0) AS N2, "
                sql &= "ISNULL(SUM(CAST(N3_F AS float)),0) AS N3, "
                sql &= "ISNULL(SUM(CAST(N4_F AS float)),0) AS N4, "
                sql &= "ISNULL(SUM(CAST(N5_F AS float)),0) AS N5, "
                sql &= "ISNULL(SUM(CAST(N6_F AS float)),0) AS N6,"
                sql &= "ISNULL(SUM(CAST(N7_F AS float)),0) AS N7, "
                sql &= "ISNULL(SUM(CAST(N8_F AS float)),0) AS N8, "
                sql &= "ISNULL(SUM(CAST(N9_F AS float)),0) AS N9, "
                sql &= "ISNULL(SUM(CAST(N10_F AS float)),0) AS N10, "
                sql &= "ISNULL(SUM(CAST(N11_F AS float)),0) AS N11, "
                sql &= "ISNULL(SUM(CAST(N12_F AS float)),0) AS N12, "
                sql &= "ISNULL(SUM(CAST(Total AS float)),0) AS Total "
                sql &= "FROM ForcastPlan "
                sql &= " Where Buyer = '" & xBuyer & "' "
                sql &= " And Y_Level = 0 "
                'TEST   sql &= " And LogID = '" & "20170823105356" & "' "
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
                H3hd1.Text = "[CONVERT]-- 差"
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

                If dt_FCTQty.Rows(0).Item("N") - dt_FCTQty1.Rows(0).Item("N") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N1") - dt_FCTQty1.Rows(0).Item("N1") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N2") - dt_FCTQty1.Rows(0).Item("N2") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N3") - dt_FCTQty1.Rows(0).Item("N3") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N4") - dt_FCTQty1.Rows(0).Item("N4") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N5") - dt_FCTQty1.Rows(0).Item("N5") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N6") - dt_FCTQty1.Rows(0).Item("N6") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N7") - dt_FCTQty1.Rows(0).Item("N7") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N8") - dt_FCTQty1.Rows(0).Item("N8") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N9") - dt_FCTQty1.Rows(0).Item("N9") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N10") - dt_FCTQty1.Rows(0).Item("N10") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N11") - dt_FCTQty1.Rows(0).Item("N11") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N12") - dt_FCTQty1.Rows(0).Item("N12") <> 0 Or _
                dt_FCTQty.Rows(0).Item("Total") - dt_FCTQty1.Rows(0).Item("Total") <> 0 Then
                    If DMessage.Text = "" Then DMessage.Text = "**[注意]發現數量不符，請再次確認客戶需求量QTY.  /CONVERT"
                End If
            End If
            '-- ----------------------------------------------------------------
            '-- 客戶 FC (FCTPLAN)
            '-- ----------------------------------------------------------------
            If xFun = "FCTPLAN" Or xFun = "LSPLAN" Or xFun = "IPLSPlan" Or xFun = "BYLSPLAN" Or xFun = "EDI" Then
                Dim H11hd1 As TableCell = New TableCell
                H11hd1.Text = "客戶 FC"
                H11row.Cells.Add(H11hd1)

                sql = "SELECT "
                If xBuyer = "FALL-000021" Then
                    sql &= "ISNULL(SUM(CAST(J1 AS float)),0) as N, "
                    sql &= "ISNULL(SUM(CAST(K1 AS float)),0) AS N1, "
                    sql &= "ISNULL(SUM(CAST(L1 AS float)),0) AS N2, "
                    sql &= "ISNULL(SUM(CAST(M1 AS float)),0) AS N3, "
                    sql &= "ISNULL(SUM(CAST(N1 AS float)),0) AS N4, "
                    sql &= "ISNULL(SUM(CAST(O1 AS float)),0) AS N5, "
                    sql &= "ISNULL(SUM(CAST(P1 AS float)),0) AS N6,"
                    sql &= "ISNULL(SUM(CAST(Q1 AS float)),0) AS N7, "
                    sql &= "ISNULL(SUM(CAST(R1 AS float)),0) AS N8, "
                    sql &= "ISNULL(SUM(CAST(S1 AS float)),0) AS N9, "
                    sql &= "ISNULL(SUM(CAST(T1 AS float)),0) AS N10, "
                    sql &= "ISNULL(SUM(CAST(U1 AS float)),0) AS N11, "
                    sql &= "ISNULL(SUM(CAST(V1 AS float)),0) AS N12, "
                    sql &= "ISNULL(SUM(CAST(W1 AS float)),0) AS Total "
                Else
                    If xBuyer = "FALL-TW0371" Or xBuyer = "FALL-TW1741" Then
                        sql &= "ISNULL(SUM(CAST(I1 AS float)),0) as N, "
                        sql &= "ISNULL(SUM(CAST(J1 AS float)),0) AS N1, "
                        sql &= "ISNULL(SUM(CAST(K1 AS float)),0) AS N2, "
                        sql &= "ISNULL(SUM(CAST(L1 AS float)),0) AS N3, "
                        sql &= "ISNULL(SUM(CAST(M1 AS float)),0) AS N4, "
                        sql &= "ISNULL(SUM(CAST(N1 AS float)),0) AS N5, "
                        sql &= "ISNULL(SUM(CAST(O1 AS float)),0) AS N6,"
                        sql &= "ISNULL(SUM(CAST(P1 AS float)),0) AS N7, "
                        sql &= "ISNULL(SUM(CAST(Q1 AS float)),0) AS N8, "
                        sql &= "ISNULL(SUM(CAST(R1 AS float)),0) AS N9, "
                        sql &= "ISNULL(SUM(CAST(S1 AS float)),0) AS N10, "
                        sql &= "ISNULL(SUM(CAST(T1 AS float)),0) AS N11, "
                        sql &= "ISNULL(SUM(CAST(U1 AS float)),0) AS N12, "
                        sql &= "ISNULL(SUM(CAST(V1 AS float)),0) AS Total "
                    Else
                        If xBuyer = "FALL-TW0371T" Then
                            sql &= "ISNULL(SUM(CAST(E1 AS float)),0) as N, "
                            sql &= "ISNULL(SUM(CAST(F1 AS float)),0) AS N1, "
                            sql &= "ISNULL(SUM(CAST(G1 AS float)),0) AS N2, "
                            sql &= "ISNULL(SUM(CAST(H1 AS float)),0) AS N3, "
                            sql &= "ISNULL(SUM(CAST(I1 AS float)),0) AS N4, "
                            sql &= "ISNULL(SUM(CAST(J1 AS float)),0) AS N5, "
                            sql &= "ISNULL(SUM(CAST(K1 AS float)),0) AS N6,"
                            sql &= "ISNULL(SUM(CAST(L1 AS float)),0) AS N7, "
                            sql &= "ISNULL(SUM(CAST(M1 AS float)),0) AS N8, "
                            sql &= "ISNULL(SUM(CAST(N1 AS float)),0) AS N9, "
                            sql &= "ISNULL(SUM(CAST(O1 AS float)),0) AS N10, "
                            sql &= "ISNULL(SUM(CAST(P1 AS float)),0) AS N11, "
                            sql &= "ISNULL(SUM(CAST(Q1 AS float)),0) AS N12, "
                            sql &= "ISNULL(SUM(CAST(R1 AS float)),0) AS Total "
                        Else
                            If xBuyer = "FALL-TP000013" Then
                                sql &= "ISNULL(SUM(CAST(H1 AS float)),0) as N, "
                                sql &= "ISNULL(SUM(CAST(I1 AS float)),0) as N1, "
                                sql &= "ISNULL(SUM(CAST(J1 AS float)),0) AS N2, "
                                sql &= "ISNULL(SUM(CAST(K1 AS float)),0) AS N3, "
                                sql &= "ISNULL(SUM(CAST(L1 AS float)),0) AS N4, "
                                sql &= "ISNULL(SUM(CAST(M1 AS float)),0) AS N5, "
                                sql &= "ISNULL(SUM(CAST(N1 AS float)),0) AS N6, "
                                sql &= "ISNULL(SUM(CAST(O1 AS float)),0) AS N7,"
                                sql &= "ISNULL(SUM(CAST(P1 AS float)),0) AS N8, "
                                sql &= "ISNULL(SUM(CAST(Q1 AS float)),0) AS N9, "
                                sql &= "ISNULL(SUM(CAST(R1 AS float)),0) AS N10, "
                                sql &= "ISNULL(SUM(CAST(S1 AS float)),0) AS N11, "
                                sql &= "ISNULL(SUM(CAST(T1 AS float)),0) AS N12, "
                                sql &= "ISNULL(SUM(CAST(U1 AS float)),0) AS Total "
                            Else
                                If xBuyer = "FALL-VENDOR" Or InStr(xBuyer, "F-VENDOR") > 0 Then
                                    sql &= "ISNULL(SUM(CAST(V1 AS float)),0) as N, "
                                    sql &= "ISNULL(SUM(CAST(W1 AS float)),0) AS N1, "
                                    sql &= "ISNULL(SUM(CAST(X1 AS float)),0) AS N2, "
                                    sql &= "ISNULL(SUM(CAST(Y1 AS float)),0) AS N3, "
                                    sql &= "ISNULL(SUM(CAST(Z1 AS float)),0) AS N4, "
                                    sql &= "ISNULL(SUM(CAST(AA1 AS float)),0) AS N5, "
                                    sql &= "0 AS N6,"
                                    sql &= "0 AS N7, "
                                    sql &= "0 AS N8, "
                                    sql &= "0 AS N9, "
                                    sql &= "0 AS N10, "
                                    sql &= "0 AS N11, "
                                    sql &= "0 AS N12, "
                                    sql &= "ISNULL(SUM(CAST(AB1 AS float)),0) AS Total "
                                Else
                                    If xBuyer = "FALL-TW0655" Then
                                        sql &= "ISNULL(SUM(CAST(G1 AS float)),0) as N, "
                                        sql &= "ISNULL(SUM(CAST(H1 AS float)),0) as N1, "
                                        sql &= "ISNULL(SUM(CAST(I1 AS float)),0) as N2, "
                                        sql &= "ISNULL(SUM(CAST(J1 AS float)),0) AS N3, "
                                        sql &= "ISNULL(SUM(CAST(K1 AS float)),0) AS N4, "
                                        sql &= "ISNULL(SUM(CAST(L1 AS float)),0) AS N5, "
                                        sql &= "ISNULL(SUM(CAST(M1 AS float)),0) AS N6, "
                                        sql &= "ISNULL(SUM(CAST(N1 AS float)),0) AS N7, "
                                        sql &= "ISNULL(SUM(CAST(O1 AS float)),0) AS N8,"
                                        sql &= "ISNULL(SUM(CAST(P1 AS float)),0) AS N9, "
                                        sql &= "ISNULL(SUM(CAST(Q1 AS float)),0) AS N10, "
                                        sql &= "ISNULL(SUM(CAST(R1 AS float)),0) AS N11, "
                                        sql &= "ISNULL(SUM(CAST(S1 AS float)),0) AS N12, "
                                        sql &= "ISNULL(SUM(CAST(T1 AS float)),0) AS Total "
                                    Else

                                        If xBuyer = "FALL-000053" Then
                                            sql &= "ISNULL(SUM(CAST(J1 AS float)),0) as N, "
                                            sql &= "ISNULL(SUM(CAST(K1 AS float)),0) as N1, "
                                            sql &= "ISNULL(SUM(CAST(L1 AS float)),0) as N2, "
                                            sql &= "ISNULL(SUM(CAST(M1 AS float)),0) AS N3, "
                                            sql &= "ISNULL(SUM(CAST(N1 AS float)),0) AS N4, "
                                            sql &= "ISNULL(SUM(CAST(O1 AS float)),0) AS N5, "
                                            sql &= "ISNULL(SUM(CAST(P1 AS float)),0) AS N6, "
                                            sql &= "ISNULL(SUM(CAST(Q1 AS float)),0) AS N7, "
                                            sql &= "ISNULL(SUM(CAST(R1 AS float)),0) AS N8,"
                                            sql &= "ISNULL(SUM(CAST(S1 AS float)),0) AS N9, "
                                            sql &= "ISNULL(SUM(CAST(T1 AS float)),0) AS N10, "
                                            sql &= "ISNULL(SUM(CAST(U1 AS float)),0) AS N11, "
                                            sql &= "ISNULL(SUM(CAST(V1 AS float)),0) AS N12, "
                                            sql &= "ISNULL(SUM(CAST(W1 AS float)),0) AS Total "
                                        Else
                                            If xBuyer = "FALL-000141" Then
                                                sql &= "ISNULL(SUM(CAST(G1 AS float)),0) as N, "
                                                sql &= "ISNULL(SUM(CAST(H1 AS float)),0) as N1, "
                                                sql &= "ISNULL(SUM(CAST(I1 AS float)),0) as N2, "
                                                sql &= "ISNULL(SUM(CAST(J1 AS float)),0) AS N3, "
                                                sql &= "ISNULL(SUM(CAST(K1 AS float)),0) AS N4, "
                                                sql &= "ISNULL(SUM(CAST(L1 AS float)),0) AS N5, "
                                                sql &= "ISNULL(SUM(CAST(M1 AS float)),0) AS N6, "
                                                sql &= "ISNULL(SUM(CAST(N1 AS float)),0) AS N7, "
                                                sql &= "ISNULL(SUM(CAST(O1 AS float)),0) AS N8,"
                                                sql &= "ISNULL(SUM(CAST(P1 AS float)),0) AS N9, "
                                                sql &= "ISNULL(SUM(CAST(Q1 AS float)),0) AS N10, "
                                                sql &= "ISNULL(SUM(CAST(R1 AS float)),0) AS N11, "
                                                sql &= "ISNULL(SUM(CAST(S1 AS float)),0) AS N12, "
                                                sql &= "ISNULL(SUM(CAST(T1 AS float)),0) AS Total "
                                            Else
                                                sql &= "ISNULL(SUM(CAST(M1 AS float)),0) as N, "
                                                sql &= "ISNULL(SUM(CAST(N1 AS float)),0) AS N1, "
                                                sql &= "ISNULL(SUM(CAST(O1 AS float)),0) AS N2, "
                                                sql &= "ISNULL(SUM(CAST(P1 AS float)),0) AS N3, "
                                                sql &= "ISNULL(SUM(CAST(Q1 AS float)),0) AS N4, "
                                                sql &= "ISNULL(SUM(CAST(R1 AS float)),0) AS N5, "
                                                sql &= "ISNULL(SUM(CAST(S1 AS float)),0) AS N6,"
                                                sql &= "ISNULL(SUM(CAST(T1 AS float)),0) AS N7, "
                                                sql &= "ISNULL(SUM(CAST(U1 AS float)),0) AS N8, "
                                                sql &= "ISNULL(SUM(CAST(V1 AS float)),0) AS N9, "
                                                sql &= "ISNULL(SUM(CAST(W1 AS float)),0) AS N10, "
                                                sql &= "ISNULL(SUM(CAST(X1 AS float)),0) AS N11, "
                                                sql &= "ISNULL(SUM(CAST(Y1 AS float)),0) AS N12, "
                                                sql &= "ISNULL(SUM(CAST(Z1 AS float)),0) AS Total "
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                sql &= "FROM E_InputSheet "
                sql &= " Where Buyer = '" & xBuyer & "' "
                Dim dt_FCTQty As DataTable = uDataBase.GetDataTable(sql)
                If dt_FCTQty.Rows.Count > 0 Then

                    Dim H11tc1 As TableCell = New TableCell
                    H11tc1.Text = Format(dt_FCTQty.Rows(0).Item("N"), "###,###,##0")
                    H11row.Cells.Add(H11tc1)

                    Dim H11tc2 As TableCell = New TableCell
                    H11tc2.Text = Format(dt_FCTQty.Rows(0).Item("N1"), "###,###,##0")
                    H11row.Cells.Add(H11tc2)

                    Dim H11tc3 As TableCell = New TableCell
                    H11tc3.Text = Format(dt_FCTQty.Rows(0).Item("N2"), "###,###,##0")
                    H11row.Cells.Add(H11tc3)

                    Dim H11tc4 As TableCell = New TableCell
                    H11tc4.Text = Format(dt_FCTQty.Rows(0).Item("N3"), "###,###,##0")
                    H11row.Cells.Add(H11tc4)

                    Dim H11tc5 As TableCell = New TableCell
                    H11tc5.Text = Format(dt_FCTQty.Rows(0).Item("N4"), "###,###,##0")
                    H11row.Cells.Add(H11tc5)

                    Dim H11tc6 As TableCell = New TableCell
                    H11tc6.Text = Format(dt_FCTQty.Rows(0).Item("N5"), "###,###,##0")
                    H11row.Cells.Add(H11tc6)

                    Dim H11tc7 As TableCell = New TableCell
                    H11tc7.Text = Format(dt_FCTQty.Rows(0).Item("N6"), "###,###,##0")
                    H11row.Cells.Add(H11tc7)

                    Dim H11tc8 As TableCell = New TableCell
                    H11tc8.Text = Format(dt_FCTQty.Rows(0).Item("N7"), "###,###,##0")
                    H11row.Cells.Add(H11tc8)

                    Dim H11tc9 As TableCell = New TableCell
                    H11tc9.Text = Format(dt_FCTQty.Rows(0).Item("N8"), "###,###,##0")
                    H11row.Cells.Add(H11tc9)

                    Dim H11tcA As TableCell = New TableCell
                    H11tcA.Text = Format(dt_FCTQty.Rows(0).Item("N9"), "###,###,##0")
                    H11row.Cells.Add(H11tcA)

                    Dim H11tcB As TableCell = New TableCell
                    H11tcB.Text = Format(dt_FCTQty.Rows(0).Item("N10"), "###,###,##0")
                    H11row.Cells.Add(H11tcB)

                    Dim H11tcC As TableCell = New TableCell
                    H11tcC.Text = Format(dt_FCTQty.Rows(0).Item("N11"), "###,###,##0")
                    H11row.Cells.Add(H11tcC)

                    Dim H11tcD As TableCell = New TableCell
                    H11tcD.Text = Format(dt_FCTQty.Rows(0).Item("N12"), "###,###,##0")
                    H11row.Cells.Add(H11tcD)

                    Dim H11tcE As TableCell = New TableCell
                    H11tcE.Text = Format(dt_FCTQty.Rows(0).Item("Total"), "###,###,##0")
                    H11row.Cells.Add(H11tcE)
                End If
                e.Row.Parent.Controls.Add(H11row)
                '-- ----------------------------------------------------------------
                '-- PLAN ZIP QTY
                '-- ----------------------------------------------------------------
                Dim H21hd1 As TableCell = New TableCell
                H21hd1.Text = "FCPLAN"
                H21row.Cells.Add(H21hd1)

                sql = "SELECT "
                sql &= "ISNULL(SUM(CAST(N_F AS float)),0) as N, "
                sql &= "ISNULL(SUM(CAST(N1_F AS float)),0) AS N1, "
                sql &= "ISNULL(SUM(CAST(N2_F AS float)),0) AS N2, "
                sql &= "ISNULL(SUM(CAST(N3_F AS float)),0) AS N3, "
                sql &= "ISNULL(SUM(CAST(N4_F AS float)),0) AS N4, "
                sql &= "ISNULL(SUM(CAST(N5_F AS float)),0) AS N5, "
                sql &= "ISNULL(SUM(CAST(N6_F AS float)),0) AS N6,"
                sql &= "ISNULL(SUM(CAST(N7_F AS float)),0) AS N7, "
                sql &= "ISNULL(SUM(CAST(N8_F AS float)),0) AS N8, "
                sql &= "ISNULL(SUM(CAST(N9_F AS float)),0) AS N9, "
                sql &= "ISNULL(SUM(CAST(N10_F AS float)),0) AS N10, "
                sql &= "ISNULL(SUM(CAST(N11_F AS float)),0) AS N11, "
                sql &= "ISNULL(SUM(CAST(N12_F AS float)),0) AS N12, "
                sql &= "ISNULL(SUM(CAST(Total AS float)),0) AS Total "
                sql &= "FROM ForcastPlan "
                sql &= " Where Buyer = '" & xBuyer & "' "
                sql &= " And Y_Level = 0 "
                'TEST   sql &= " And LogID = '" & "20170823105356" & "' "
                Dim dt_FCTQty1 As DataTable = uDataBase.GetDataTable(sql)
                If dt_FCTQty1.Rows.Count > 0 Then

                    Dim H21tc1 As TableCell = New TableCell
                    H21tc1.Text = Format(dt_FCTQty1.Rows(0).Item("N"), "###,###,##0")
                    H21row.Cells.Add(H21tc1)

                    Dim H21tc2 As TableCell = New TableCell
                    H21tc2.Text = Format(dt_FCTQty1.Rows(0).Item("N1"), "###,###,##0")
                    H21row.Cells.Add(H21tc2)

                    Dim H21tc3 As TableCell = New TableCell
                    H21tc3.Text = Format(dt_FCTQty1.Rows(0).Item("N2"), "###,###,##0")
                    H21row.Cells.Add(H21tc3)

                    Dim H21tc4 As TableCell = New TableCell
                    H21tc4.Text = Format(dt_FCTQty1.Rows(0).Item("N3"), "###,###,##0")
                    H21row.Cells.Add(H21tc4)

                    Dim H21tc5 As TableCell = New TableCell
                    H21tc5.Text = Format(dt_FCTQty1.Rows(0).Item("N4"), "###,###,##0")
                    H21row.Cells.Add(H21tc5)

                    Dim H21tc6 As TableCell = New TableCell
                    H21tc6.Text = Format(dt_FCTQty1.Rows(0).Item("N5"), "###,###,##0")
                    H21row.Cells.Add(H21tc6)

                    Dim H21tc7 As TableCell = New TableCell
                    H21tc7.Text = Format(dt_FCTQty1.Rows(0).Item("N6"), "###,###,##0")
                    H21row.Cells.Add(H21tc7)

                    Dim H21tc8 As TableCell = New TableCell
                    H21tc8.Text = Format(dt_FCTQty1.Rows(0).Item("N7"), "###,###,##0")
                    H21row.Cells.Add(H21tc8)

                    Dim H21tc9 As TableCell = New TableCell
                    H21tc9.Text = Format(dt_FCTQty1.Rows(0).Item("N8"), "###,###,##0")
                    H21row.Cells.Add(H21tc9)

                    Dim H21tcA As TableCell = New TableCell
                    H21tcA.Text = Format(dt_FCTQty1.Rows(0).Item("N9"), "###,###,##0")
                    H21row.Cells.Add(H21tcA)

                    Dim H21tcB As TableCell = New TableCell
                    H21tcB.Text = Format(dt_FCTQty1.Rows(0).Item("N10"), "###,###,##0")
                    H21row.Cells.Add(H21tcB)

                    Dim H21tcC As TableCell = New TableCell
                    H21tcC.Text = Format(dt_FCTQty1.Rows(0).Item("N11"), "###,###,##0")
                    H21row.Cells.Add(H21tcC)

                    Dim H21tcD As TableCell = New TableCell
                    H21tcD.Text = Format(dt_FCTQty1.Rows(0).Item("N12"), "###,###,##0")
                    H21row.Cells.Add(H21tcD)

                    Dim H21tcE As TableCell = New TableCell
                    H21tcE.Text = Format(dt_FCTQty1.Rows(0).Item("Total"), "###,###,##0")
                    H21row.Cells.Add(H21tcE)
                End If
                e.Row.Parent.Controls.Add(H21row)
                '-- ----------------------------------------------------------------
                '-- 差異
                '-- 
                '-- ----------------------------------------------------------------
                Dim H31hd1 As TableCell = New TableCell
                H31hd1.BackColor = Color.LightPink
                H31hd1.Text = "[FCPLAN]-- 差"

                H31row.Cells.Add(H31hd1)

                Dim H31tc1 As TableCell = New TableCell
                H31tc1.HorizontalAlign = HorizontalAlign.Right
                H31tc1.BackColor = Color.LightPink
                H31tc1.Text = Format(dt_FCTQty.Rows(0).Item("N") - dt_FCTQty1.Rows(0).Item("N"), "###,###,##0")
                H31row.Cells.Add(H31tc1)

                Dim H31tc2 As TableCell = New TableCell
                H31tc2.HorizontalAlign = HorizontalAlign.Right
                H31tc2.BackColor = Color.LightPink
                H31tc2.Text = Format(dt_FCTQty.Rows(0).Item("N1") - dt_FCTQty1.Rows(0).Item("N1"), "###,###,##0")
                H31row.Cells.Add(H31tc2)

                Dim H31tc3 As TableCell = New TableCell
                H31tc3.HorizontalAlign = HorizontalAlign.Right
                H31tc3.BackColor = Color.LightPink
                H31tc3.Text = Format(dt_FCTQty.Rows(0).Item("N2") - dt_FCTQty1.Rows(0).Item("N2"), "###,###,##0")
                H31row.Cells.Add(H31tc3)

                Dim H31tc4 As TableCell = New TableCell
                H31tc4.HorizontalAlign = HorizontalAlign.Right
                H31tc4.BackColor = Color.LightPink
                H31tc4.Text = Format(dt_FCTQty.Rows(0).Item("N3") - dt_FCTQty1.Rows(0).Item("N3"), "###,###,##0")
                H31row.Cells.Add(H31tc4)

                Dim H31tc5 As TableCell = New TableCell
                H31tc5.HorizontalAlign = HorizontalAlign.Right
                H31tc5.BackColor = Color.LightPink
                H31tc5.Text = Format(dt_FCTQty.Rows(0).Item("N4") - dt_FCTQty1.Rows(0).Item("N4"), "###,###,##0")
                H31row.Cells.Add(H31tc5)

                Dim H31tc6 As TableCell = New TableCell
                H31tc6.HorizontalAlign = HorizontalAlign.Right
                H31tc6.BackColor = Color.LightPink
                H31tc6.Text = Format(dt_FCTQty.Rows(0).Item("N5") - dt_FCTQty1.Rows(0).Item("N5"), "###,###,##0")
                H31row.Cells.Add(H31tc6)

                Dim H31tc7 As TableCell = New TableCell
                H31tc7.HorizontalAlign = HorizontalAlign.Right
                H31tc7.BackColor = Color.LightPink
                H31tc7.Text = Format(dt_FCTQty.Rows(0).Item("N6") - dt_FCTQty1.Rows(0).Item("N6"), "###,###,##0")
                H31row.Cells.Add(H31tc7)

                Dim H31tc8 As TableCell = New TableCell
                H31tc8.HorizontalAlign = HorizontalAlign.Right
                H31tc8.BackColor = Color.LightPink
                H31tc8.Text = Format(dt_FCTQty.Rows(0).Item("N7") - dt_FCTQty1.Rows(0).Item("N7"), "###,###,##0")
                H31row.Cells.Add(H31tc8)

                Dim H31tc9 As TableCell = New TableCell
                H31tc9.HorizontalAlign = HorizontalAlign.Right
                H31tc9.BackColor = Color.LightPink
                H31tc9.Text = Format(dt_FCTQty.Rows(0).Item("N8") - dt_FCTQty1.Rows(0).Item("N8"), "###,###,##0")
                H31row.Cells.Add(H31tc9)

                Dim H31tcA As TableCell = New TableCell
                H31tcA.HorizontalAlign = HorizontalAlign.Right
                H31tcA.BackColor = Color.LightPink
                H31tcA.Text = Format(dt_FCTQty.Rows(0).Item("N9") - dt_FCTQty1.Rows(0).Item("N9"), "###,###,##0")
                H31row.Cells.Add(H31tcA)

                Dim H31tcB As TableCell = New TableCell
                H31tcB.HorizontalAlign = HorizontalAlign.Right
                H31tcB.BackColor = Color.LightPink
                H31tcB.Text = Format(dt_FCTQty.Rows(0).Item("N10") - dt_FCTQty1.Rows(0).Item("N10"), "###,###,##0")
                H31row.Cells.Add(H31tcB)

                Dim H31tcC As TableCell = New TableCell
                H31tcC.HorizontalAlign = HorizontalAlign.Right
                H31tcC.BackColor = Color.LightPink
                H31tcC.Text = Format(dt_FCTQty.Rows(0).Item("N11") - dt_FCTQty1.Rows(0).Item("N11"), "###,###,##0")
                H31row.Cells.Add(H31tcC)

                Dim H31tcD As TableCell = New TableCell
                H31tcD.HorizontalAlign = HorizontalAlign.Right
                H31tcD.BackColor = Color.LightPink
                H31tcD.Text = Format(dt_FCTQty.Rows(0).Item("N12") - dt_FCTQty1.Rows(0).Item("N12"), "###,###,##0")
                H31row.Cells.Add(H31tcD)

                Dim H31tcE As TableCell = New TableCell
                H31tcE.HorizontalAlign = HorizontalAlign.Right
                H31tcE.BackColor = Color.LightPink
                H31tcE.Text = Format(dt_FCTQty.Rows(0).Item("Total") - dt_FCTQty1.Rows(0).Item("Total"), "###,###,##0")
                H31row.Cells.Add(H31tcE)
                e.Row.Parent.Controls.Add(H31row)

                If dt_FCTQty.Rows(0).Item("N") - dt_FCTQty1.Rows(0).Item("N") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N1") - dt_FCTQty1.Rows(0).Item("N1") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N2") - dt_FCTQty1.Rows(0).Item("N2") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N3") - dt_FCTQty1.Rows(0).Item("N3") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N4") - dt_FCTQty1.Rows(0).Item("N4") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N5") - dt_FCTQty1.Rows(0).Item("N5") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N6") - dt_FCTQty1.Rows(0).Item("N6") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N7") - dt_FCTQty1.Rows(0).Item("N7") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N8") - dt_FCTQty1.Rows(0).Item("N8") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N9") - dt_FCTQty1.Rows(0).Item("N9") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N10") - dt_FCTQty1.Rows(0).Item("N10") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N11") - dt_FCTQty1.Rows(0).Item("N11") <> 0 Or _
                dt_FCTQty.Rows(0).Item("N12") - dt_FCTQty1.Rows(0).Item("N12") <> 0 Or _
                dt_FCTQty.Rows(0).Item("Total") - dt_FCTQty1.Rows(0).Item("Total") <> 0 Then
                    If DMessage.Text = "" Then DMessage.Text = "**[注意]發現數量不符，請再次確認客戶需求量QTY.  /FCTPLAN"
                End If
            End If
                '-- ----------------------------------------------------------------
                '-- SLD QTY
                '-- ----------------------------------------------------------------
                If xFun = "LSPLAN" Or xFun = "IPLSPlan" Or xFun = "BYLSPLAN" Or xFun = "EDI" Then
                    Dim H4hd1 As TableCell = New TableCell
                    If xBuyer = "FALL-TP000013" Then
                        H4hd1.Text = "FCPLAN"
                    Else
                        H4hd1.Text = "SLD-FCPLAN"
                    End If
                    H4row.Cells.Add(H4hd1)

                    sql = "SELECT "
                    sql &= "ISNULL(SUM(CAST(N_F AS float)),0) as N, "
                    sql &= "ISNULL(SUM(CAST(N1_F AS float)),0) AS N1, "
                    sql &= "ISNULL(SUM(CAST(N2_F AS float)),0) AS N2, "
                    sql &= "ISNULL(SUM(CAST(N3_F AS float)),0) AS N3, "
                    sql &= "ISNULL(SUM(CAST(N4_F AS float)),0) AS N4, "
                    sql &= "ISNULL(SUM(CAST(N5_F AS float)),0) AS N5, "
                    sql &= "ISNULL(SUM(CAST(N6_F AS float)),0) AS N6,"
                    sql &= "ISNULL(SUM(CAST(N7_F AS float)),0) AS N7, "
                    sql &= "ISNULL(SUM(CAST(N8_F AS float)),0) AS N8, "
                    sql &= "ISNULL(SUM(CAST(N9_F AS float)),0) AS N9, "
                    sql &= "ISNULL(SUM(CAST(N10_F AS float)),0) AS N10, "
                    sql &= "ISNULL(SUM(CAST(N11_F AS float)),0) AS N11, "
                    sql &= "ISNULL(SUM(CAST(N12_F AS float)),0) AS N12, "
                    sql &= "ISNULL(SUM(CAST(Total AS float)),0) AS Total "
                    sql &= "FROM ForcastPlan "
                    sql &= " Where Buyer = '" & xBuyer & "' "
                    sql &= " And Y_Level = 1 "
                    If xBuyer = "FALL-TP000013" Then
                        sql &= " And Y_A1 = '" & "TP" & "' "
                    Else
                        sql &= " And Y_A1 = '" & "PS" & "' "
                    End If
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
                    '-- LS SLD QTY
                    '-- ----------------------------------------------------------------
                    Dim H5hd1 As TableCell = New TableCell
                    If xBuyer = "FALL-TP000013" Then
                        H5hd1.Text = "LSPLAN"
                    Else
                        H5hd1.Text = "SLD-LSPLAN"
                    End If
                    H5row.Cells.Add(H5hd1)

                    sql = "SELECT "
                    sql &= "ISNULL(SUM(CAST(FS_00 AS float)),0) as N, "
                    sql &= "ISNULL(SUM(CAST(FS_01 AS float)),0) as N1, "
                    sql &= "ISNULL(SUM(CAST(FS_02 AS float)),0) as N2, "
                    sql &= "ISNULL(SUM(CAST(FS_03 AS float)),0) as N3, "
                    sql &= "ISNULL(SUM(CAST(FS_04 AS float)),0) as N4, "
                    sql &= "ISNULL(SUM(CAST(FS_05 AS float)),0) as N5, "
                    sql &= "ISNULL(SUM(CAST(FS_06 AS float)),0) as N6, "
                    sql &= "ISNULL(SUM(CAST(FS_07 AS float)),0) as N7, "
                    sql &= "ISNULL(SUM(CAST(FS_08 AS float)),0) as N8, "
                    sql &= "ISNULL(SUM(CAST(FS_09 AS float)),0) as N9, "
                    sql &= "ISNULL(SUM(CAST(FS_10 AS float)),0) as N10, "
                    sql &= "ISNULL(SUM(CAST(FS_11 AS float)),0) as N11, "
                    sql &= "ISNULL(SUM(CAST(FS_12 AS float)),0) as N12, "
                    sql &= "ISNULL(SUM(CAST(FS_13 AS float)),0) as Total "
                    sql &= "FROM LocalStockPlan "
                    sql &= " Where Buyer = '" & xBuyer & "' "
                    sql &= " And Version = 99 "
                    If xBuyer = "FALL-TP000013" Then
                        sql &= " And GR_08 = '" & "TP" & "' "
                    Else
                        sql &= " And GR_08 = '" & "PS" & "' "
                    End If
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
                    If xBuyer = "FALL-TP000013" Then
                        H6hd1.Text = "[LSPLAN]-- 差"
                    Else
                        H6hd1.Text = "[LSPLAN-SLD]-- 差"
                    End If
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

                    If dt_FCTQty2.Rows(0).Item("N") - dt_FCTQty3.Rows(0).Item("N") <> 0 Or _
                    dt_FCTQty2.Rows(0).Item("N1") - dt_FCTQty3.Rows(0).Item("N1") <> 0 Or _
                    dt_FCTQty2.Rows(0).Item("N2") - dt_FCTQty3.Rows(0).Item("N2") <> 0 Or _
                    dt_FCTQty2.Rows(0).Item("N3") - dt_FCTQty3.Rows(0).Item("N3") <> 0 Or _
                    dt_FCTQty2.Rows(0).Item("N4") - dt_FCTQty3.Rows(0).Item("N4") <> 0 Or _
                    dt_FCTQty2.Rows(0).Item("N5") - dt_FCTQty3.Rows(0).Item("N5") <> 0 Or _
                    dt_FCTQty2.Rows(0).Item("N6") - dt_FCTQty3.Rows(0).Item("N6") <> 0 Or _
                    dt_FCTQty2.Rows(0).Item("N7") - dt_FCTQty3.Rows(0).Item("N7") <> 0 Or _
                    dt_FCTQty2.Rows(0).Item("N8") - dt_FCTQty3.Rows(0).Item("N8") <> 0 Or _
                    dt_FCTQty2.Rows(0).Item("N9") - dt_FCTQty3.Rows(0).Item("N9") <> 0 Or _
                    dt_FCTQty2.Rows(0).Item("N10") - dt_FCTQty3.Rows(0).Item("N10") <> 0 Or _
                    dt_FCTQty2.Rows(0).Item("N11") - dt_FCTQty3.Rows(0).Item("N11") <> 0 Or _
                    dt_FCTQty2.Rows(0).Item("N12") - dt_FCTQty3.Rows(0).Item("N12") <> 0 Or _
                    dt_FCTQty2.Rows(0).Item("Total") - dt_FCTQty3.Rows(0).Item("Total") <> 0 Then
                        If DMessage.Text = "" Then DMessage.Text = "**[注意]發現數量不符，請再次確認客戶需求量QTY.  /LS-SLD"
                    End If
                    '-- ----------------------------------------------------------------
                    '-- CH QTY
                    '-- ----------------------------------------------------------------
                    If xBuyer = "FALL-TP000013" Then
                    Else
                        Dim H7hd1 As TableCell = New TableCell
                        H7hd1.Text = "CH-FCPLAN"
                        H7row.Cells.Add(H7hd1)

                        sql = "SELECT "
                        sql &= "ISNULL(SUM(CAST(N_F AS float)),0) as N, "
                        sql &= "ISNULL(SUM(CAST(N1_F AS float)),0) AS N1, "
                        sql &= "ISNULL(SUM(CAST(N2_F AS float)),0) AS N2, "
                        sql &= "ISNULL(SUM(CAST(N3_F AS float)),0) AS N3, "
                        sql &= "ISNULL(SUM(CAST(N4_F AS float)),0) AS N4, "
                        sql &= "ISNULL(SUM(CAST(N5_F AS float)),0) AS N5, "
                        sql &= "ISNULL(SUM(CAST(N6_F AS float)),0) AS N6,"
                        sql &= "ISNULL(SUM(CAST(N7_F AS float)),0) AS N7, "
                        sql &= "ISNULL(SUM(CAST(N8_F AS float)),0) AS N8, "
                        sql &= "ISNULL(SUM(CAST(N9_F AS float)),0) AS N9, "
                        sql &= "ISNULL(SUM(CAST(N10_F AS float)),0) AS N10, "
                        sql &= "ISNULL(SUM(CAST(N11_F AS float)),0) AS N11, "
                        sql &= "ISNULL(SUM(CAST(N12_F AS float)),0) AS N12, "
                        sql &= "ISNULL(SUM(CAST(Total AS float)),0) AS Total "
                        sql &= "FROM ForcastPlan "
                        sql &= " Where Buyer = '" & xBuyer & "' "
                        sql &= " And Y_Level = 1 "
                        sql &= " And Y_A1 = '" & "CH" & "' "
                        'TEST   sql &= " And LogID = '" & "20170823105356" & "' "
                        Dim dt_FCTQty4 As DataTable = uDataBase.GetDataTable(sql)
                        If dt_FCTQty4.Rows.Count > 0 Then

                            Dim H7tc1 As TableCell = New TableCell
                            H7tc1.Text = Format(dt_FCTQty4.Rows(0).Item("N"), "###,###,##0.0")
                            H7row.Cells.Add(H7tc1)

                            Dim H7tc2 As TableCell = New TableCell
                            H7tc2.Text = Format(dt_FCTQty4.Rows(0).Item("N1"), "###,###,##0.0")
                            H7row.Cells.Add(H7tc2)

                            Dim H7tc3 As TableCell = New TableCell
                            H7tc3.Text = Format(dt_FCTQty4.Rows(0).Item("N2"), "###,###,##0.0")
                            H7row.Cells.Add(H7tc3)

                            Dim H7tc4 As TableCell = New TableCell
                            H7tc4.Text = Format(dt_FCTQty4.Rows(0).Item("N3"), "###,###,##0.0")
                            H7row.Cells.Add(H7tc4)

                            Dim H7tc5 As TableCell = New TableCell
                            H7tc5.Text = Format(dt_FCTQty4.Rows(0).Item("N4"), "###,###,##0.0")
                            H7row.Cells.Add(H7tc5)

                            Dim H7tc6 As TableCell = New TableCell
                            H7tc6.Text = Format(dt_FCTQty4.Rows(0).Item("N5"), "###,###,##0.0")
                            H7row.Cells.Add(H7tc6)

                            Dim H7tc7 As TableCell = New TableCell
                            H7tc7.Text = Format(dt_FCTQty4.Rows(0).Item("N6"), "###,###,##0.0")
                            H7row.Cells.Add(H7tc7)

                            Dim H7tc8 As TableCell = New TableCell
                            H7tc8.Text = Format(dt_FCTQty4.Rows(0).Item("N7"), "###,###,##0.0")
                            H7row.Cells.Add(H7tc8)

                            Dim H7tc9 As TableCell = New TableCell
                            H7tc9.Text = Format(dt_FCTQty4.Rows(0).Item("N8"), "###,###,##0.0")
                            H7row.Cells.Add(H7tc9)

                            Dim H7tcA As TableCell = New TableCell
                            H7tcA.Text = Format(dt_FCTQty4.Rows(0).Item("N9"), "###,###,##0.0")
                            H7row.Cells.Add(H7tcA)

                            Dim H7tcB As TableCell = New TableCell
                            H7tcB.Text = Format(dt_FCTQty4.Rows(0).Item("N10"), "###,###,##0.0")
                            H7row.Cells.Add(H7tcB)

                            Dim H7tcC As TableCell = New TableCell
                            H7tcC.Text = Format(dt_FCTQty4.Rows(0).Item("N11"), "###,###,##0.0")
                            H7row.Cells.Add(H7tcC)

                            Dim H7tcD As TableCell = New TableCell
                            H7tcD.Text = Format(dt_FCTQty4.Rows(0).Item("N12"), "###,###,##0.0")
                            H7row.Cells.Add(H7tcD)

                            Dim H7tcE As TableCell = New TableCell
                            H7tcE.Text = Format(dt_FCTQty4.Rows(0).Item("Total"), "###,###,##0.0")
                            H7row.Cells.Add(H7tcE)
                        End If
                        e.Row.Parent.Controls.Add(H7row)
                        '-- ----------------------------------------------------------------
                        '-- LS CH QTY
                        '-- ----------------------------------------------------------------
                        Dim H8hd1 As TableCell = New TableCell
                        H8hd1.Text = "CH-LSPLAN"
                        H8row.Cells.Add(H8hd1)

                        sql = "SELECT "
                        sql &= "ISNULL(SUM(CAST(FS_00 AS float)),0) as N, "
                        sql &= "ISNULL(SUM(CAST(FS_01 AS float)),0) as N1, "
                        sql &= "ISNULL(SUM(CAST(FS_02 AS float)),0) as N2, "
                        sql &= "ISNULL(SUM(CAST(FS_03 AS float)),0) as N3, "
                        sql &= "ISNULL(SUM(CAST(FS_04 AS float)),0) as N4, "
                        sql &= "ISNULL(SUM(CAST(FS_05 AS float)),0) as N5, "
                        sql &= "ISNULL(SUM(CAST(FS_06 AS float)),0) as N6, "
                        sql &= "ISNULL(SUM(CAST(FS_07 AS float)),0) as N7, "
                        sql &= "ISNULL(SUM(CAST(FS_08 AS float)),0) as N8, "
                        sql &= "ISNULL(SUM(CAST(FS_09 AS float)),0) as N9, "
                        sql &= "ISNULL(SUM(CAST(FS_10 AS float)),0) as N10, "
                        sql &= "ISNULL(SUM(CAST(FS_11 AS float)),0) as N11, "
                        sql &= "ISNULL(SUM(CAST(FS_12 AS float)),0) as N12, "
                        sql &= "ISNULL(SUM(CAST(FS_13 AS float)),0) as Total "
                        sql &= "FROM LocalStockPlan "
                        sql &= " Where Buyer = '" & xBuyer & "' "
                        sql &= " And Version = 99 "
                        sql &= " And GR_08 = '" & "CH" & "' "
                        'TEST   sql &= " And LogID = '" & "20170823105356" & "' "
                        Dim dt_FCTQty5 As DataTable = uDataBase.GetDataTable(sql)
                        If dt_FCTQty5.Rows.Count > 0 Then

                            Dim H8tc1 As TableCell = New TableCell
                            H8tc1.Text = Format(dt_FCTQty5.Rows(0).Item("N"), "###,###,##0.0")
                            H8row.Cells.Add(H8tc1)

                            Dim H8tc2 As TableCell = New TableCell
                            H8tc2.Text = Format(dt_FCTQty5.Rows(0).Item("N1"), "###,###,##0.0")
                            H8row.Cells.Add(H8tc2)

                            Dim H8tc3 As TableCell = New TableCell
                            H8tc3.Text = Format(dt_FCTQty5.Rows(0).Item("N2"), "###,###,##0.0")
                            H8row.Cells.Add(H8tc3)

                            Dim H8tc4 As TableCell = New TableCell
                            H8tc4.Text = Format(dt_FCTQty5.Rows(0).Item("N3"), "###,###,##0.0")
                            H8row.Cells.Add(H8tc4)

                            Dim H8tc5 As TableCell = New TableCell
                            H8tc5.Text = Format(dt_FCTQty5.Rows(0).Item("N4"), "###,###,##0.0")
                            H8row.Cells.Add(H8tc5)

                            Dim H8tc6 As TableCell = New TableCell
                            H8tc6.Text = Format(dt_FCTQty5.Rows(0).Item("N5"), "###,###,##0.0")
                            H8row.Cells.Add(H8tc6)

                            Dim H8tc7 As TableCell = New TableCell
                            H8tc7.Text = Format(dt_FCTQty5.Rows(0).Item("N6"), "###,###,##0.0")
                            H8row.Cells.Add(H8tc7)

                            Dim H8tc8 As TableCell = New TableCell
                            H8tc8.Text = Format(dt_FCTQty5.Rows(0).Item("N7"), "###,###,##0.0")
                            H8row.Cells.Add(H8tc8)

                            Dim H8tc9 As TableCell = New TableCell
                            H8tc9.Text = Format(dt_FCTQty5.Rows(0).Item("N8"), "###,###,##0.0")
                            H8row.Cells.Add(H8tc9)

                            Dim H8tcA As TableCell = New TableCell
                            H8tcA.Text = Format(dt_FCTQty5.Rows(0).Item("N9"), "###,###,##0.0")
                            H8row.Cells.Add(H8tcA)

                            Dim H8tcB As TableCell = New TableCell
                            H8tcB.Text = Format(dt_FCTQty5.Rows(0).Item("N10"), "###,###,##0.0")
                            H8row.Cells.Add(H8tcB)

                            Dim H8tcC As TableCell = New TableCell
                            H8tcC.Text = Format(dt_FCTQty5.Rows(0).Item("N11"), "###,###,##0.0")
                            H8row.Cells.Add(H8tcC)

                            Dim H8tcD As TableCell = New TableCell
                            H8tcD.Text = Format(dt_FCTQty5.Rows(0).Item("N12"), "###,###,##0.0")
                            H8row.Cells.Add(H8tcD)

                            Dim H8tcE As TableCell = New TableCell
                            H8tcE.Text = Format(dt_FCTQty5.Rows(0).Item("Total"), "###,###,##0.0")
                            H8row.Cells.Add(H8tcE)
                        End If
                        e.Row.Parent.Controls.Add(H8row)
                        '-- ----------------------------------------------------------------
                        '-- 差異
                        '-- 
                        '-- ----------------------------------------------------------------
                        Dim H9hd1 As TableCell = New TableCell
                        H9hd1.BackColor = Color.LightPink
                        H9hd1.Text = "[LSPLAN-CH]-- 差"
                        H9row.Cells.Add(H9hd1)

                        Dim H9tc1 As TableCell = New TableCell
                        H9tc1.HorizontalAlign = HorizontalAlign.Right
                        H9tc1.BackColor = Color.LightPink
                        H9tc1.Text = Format(dt_FCTQty4.Rows(0).Item("N") - dt_FCTQty5.Rows(0).Item("N"), "###,###,##0.0")
                        H9row.Cells.Add(H9tc1)

                        Dim H9tc2 As TableCell = New TableCell
                        H9tc2.HorizontalAlign = HorizontalAlign.Right
                        H9tc2.BackColor = Color.LightPink
                        H9tc2.Text = Format(dt_FCTQty4.Rows(0).Item("N1") - dt_FCTQty5.Rows(0).Item("N1"), "###,###,##0.0")
                        H9row.Cells.Add(H9tc2)

                        Dim H9tc3 As TableCell = New TableCell
                        H9tc3.HorizontalAlign = HorizontalAlign.Right
                        H9tc3.BackColor = Color.LightPink
                        H9tc3.Text = Format(dt_FCTQty4.Rows(0).Item("N2") - dt_FCTQty5.Rows(0).Item("N2"), "###,###,##0.0")
                        H9row.Cells.Add(H9tc3)

                        Dim H9tc4 As TableCell = New TableCell
                        H9tc4.HorizontalAlign = HorizontalAlign.Right
                        H9tc4.BackColor = Color.LightPink
                        H9tc4.Text = Format(dt_FCTQty4.Rows(0).Item("N3") - dt_FCTQty5.Rows(0).Item("N3"), "###,###,##0.0")
                        H9row.Cells.Add(H9tc4)

                        Dim H9tc5 As TableCell = New TableCell
                        H9tc5.HorizontalAlign = HorizontalAlign.Right
                        H9tc5.BackColor = Color.LightPink
                        H9tc5.Text = Format(dt_FCTQty4.Rows(0).Item("N4") - dt_FCTQty5.Rows(0).Item("N4"), "###,###,##0.0")
                        H9row.Cells.Add(H9tc5)

                        Dim H9tc6 As TableCell = New TableCell
                        H9tc6.HorizontalAlign = HorizontalAlign.Right
                        H9tc6.BackColor = Color.LightPink
                        H9tc6.Text = Format(dt_FCTQty4.Rows(0).Item("N5") - dt_FCTQty5.Rows(0).Item("N5"), "###,###,##0.0")
                        H9row.Cells.Add(H9tc6)

                        Dim H9tc7 As TableCell = New TableCell
                        H9tc7.HorizontalAlign = HorizontalAlign.Right
                        H9tc7.BackColor = Color.LightPink
                        H9tc7.Text = Format(dt_FCTQty4.Rows(0).Item("N6") - dt_FCTQty5.Rows(0).Item("N6"), "###,###,##0.0")
                        H9row.Cells.Add(H9tc7)

                        Dim H9tc8 As TableCell = New TableCell
                        H9tc8.HorizontalAlign = HorizontalAlign.Right
                        H9tc8.BackColor = Color.LightPink
                        H9tc8.Text = Format(dt_FCTQty4.Rows(0).Item("N7") - dt_FCTQty5.Rows(0).Item("N7"), "###,###,##0.0")
                        H9row.Cells.Add(H9tc8)

                        Dim H9tc9 As TableCell = New TableCell
                        H9tc9.HorizontalAlign = HorizontalAlign.Right
                        H9tc9.BackColor = Color.LightPink
                        H9tc9.Text = Format(dt_FCTQty4.Rows(0).Item("N8") - dt_FCTQty5.Rows(0).Item("N8"), "###,###,##0.0")
                        H9row.Cells.Add(H9tc9)

                        Dim H9tcA As TableCell = New TableCell
                        H9tcA.HorizontalAlign = HorizontalAlign.Right
                        H9tcA.BackColor = Color.LightPink
                        H9tcA.Text = Format(dt_FCTQty4.Rows(0).Item("N9") - dt_FCTQty5.Rows(0).Item("N9"), "###,###,##0.0")
                        H9row.Cells.Add(H9tcA)

                        Dim H9tcB As TableCell = New TableCell
                        H9tcB.HorizontalAlign = HorizontalAlign.Right
                        H9tcB.BackColor = Color.LightPink
                        H9tcB.Text = Format(dt_FCTQty4.Rows(0).Item("N10") - dt_FCTQty5.Rows(0).Item("N10"), "###,###,##0.0")
                        H9row.Cells.Add(H9tcB)

                        Dim H9tcC As TableCell = New TableCell
                        H9tcC.HorizontalAlign = HorizontalAlign.Right
                        H9tcC.BackColor = Color.LightPink
                        H9tcC.Text = Format(dt_FCTQty4.Rows(0).Item("N11") - dt_FCTQty5.Rows(0).Item("N11"), "###,###,##0.0")
                        H9row.Cells.Add(H9tcC)

                        Dim H9tcD As TableCell = New TableCell
                        H9tcD.HorizontalAlign = HorizontalAlign.Right
                        H9tcD.BackColor = Color.LightPink
                        H9tcD.Text = Format(dt_FCTQty4.Rows(0).Item("N12") - dt_FCTQty5.Rows(0).Item("N12"), "###,###,##0.0")
                        H9row.Cells.Add(H9tcD)

                        Dim H9tcE As TableCell = New TableCell
                        H9tcE.HorizontalAlign = HorizontalAlign.Right
                        H9tcE.BackColor = Color.LightPink
                        H9tcE.Text = Format(dt_FCTQty4.Rows(0).Item("Total") - dt_FCTQty5.Rows(0).Item("Total"), "###,###,##0.0")
                        H9row.Cells.Add(H9tcE)
                        e.Row.Parent.Controls.Add(H9row)

                        If Format(dt_FCTQty4.Rows(0).Item("N"), "###,###,##0.0") <> Format(dt_FCTQty5.Rows(0).Item("N"), "###,###,##0.0") Or _
                        Format(dt_FCTQty4.Rows(0).Item("N1"), "###,###,##0.0") <> Format(dt_FCTQty5.Rows(0).Item("N1"), "###,###,##0.0") Or _
                        Format(dt_FCTQty4.Rows(0).Item("N2"), "###,###,##0.0") <> Format(dt_FCTQty5.Rows(0).Item("N2"), "###,###,##0.0") Or _
                        Format(dt_FCTQty4.Rows(0).Item("N3"), "###,###,##0.0") <> Format(dt_FCTQty5.Rows(0).Item("N3"), "###,###,##0.0") Or _
                        Format(dt_FCTQty4.Rows(0).Item("N4"), "###,###,##0.0") <> Format(dt_FCTQty5.Rows(0).Item("N4"), "###,###,##0.0") Or _
                        Format(dt_FCTQty4.Rows(0).Item("N5"), "###,###,##0.0") <> Format(dt_FCTQty5.Rows(0).Item("N5"), "###,###,##0.0") Or _
                        Format(dt_FCTQty4.Rows(0).Item("N6"), "###,###,##0.0") <> Format(dt_FCTQty5.Rows(0).Item("N6"), "###,###,##0.0") Or _
                        Format(dt_FCTQty4.Rows(0).Item("N7"), "###,###,##0.0") <> Format(dt_FCTQty5.Rows(0).Item("N7"), "###,###,##0.0") Or _
                        Format(dt_FCTQty4.Rows(0).Item("N8"), "###,###,##0.0") <> Format(dt_FCTQty5.Rows(0).Item("N8"), "###,###,##0.0") Or _
                        Format(dt_FCTQty4.Rows(0).Item("N9"), "###,###,##0.0") <> Format(dt_FCTQty5.Rows(0).Item("N9"), "###,###,##0.0") Or _
                        Format(dt_FCTQty4.Rows(0).Item("N10"), "###,###,##0.0") <> Format(dt_FCTQty5.Rows(0).Item("N10"), "###,###,##0.0") Or _
                        Format(dt_FCTQty4.Rows(0).Item("N11"), "###,###,##0.0") <> Format(dt_FCTQty5.Rows(0).Item("N11"), "###,###,##0.0") Or _
                        Format(dt_FCTQty4.Rows(0).Item("N12"), "###,###,##0.0") <> Format(dt_FCTQty5.Rows(0).Item("N12"), "###,###,##0.0") Or _
                        Format(dt_FCTQty4.Rows(0).Item("Total"), "###,###,##0.0") <> Format(dt_FCTQty5.Rows(0).Item("Total"), "###,###,##0.0") Then
                            If DMessage.Text = "" Then DMessage.Text = "**[注意]發現數量不符，請再次確認客戶需求量QTY.  /LS-CH"
                        End If
                    End If
                End If
                '-- ----------------------------------------------------------------
                '-- SLD QTY
                '-- ----------------------------------------------------------------
                If xFun = "IPLSPlan" Or xFun = "BYLSPLAN" Or xFun = "EDI" Then
                    Dim HAhd1 As TableCell = New TableCell
                    If xBuyer = "FALL-TP000013" Then
                        HAhd1.Text = "LSPLAN"
                    Else
                        HAhd1.Text = "SLD-LSPLAN"
                    End If
                    HArow.Cells.Add(HAhd1)

                    sql = "SELECT "
                    sql &= "ISNULL(SUM(CAST(FS_00 AS float)),0) as N, "
                    sql &= "ISNULL(SUM(CAST(FS_01 AS float)),0) as N1, "
                    sql &= "ISNULL(SUM(CAST(FS_02 AS float)),0) as N2, "
                    sql &= "ISNULL(SUM(CAST(FS_03 AS float)),0) as N3, "
                    sql &= "ISNULL(SUM(CAST(FS_04 AS float)),0) as N4, "
                    sql &= "ISNULL(SUM(CAST(FS_05 AS float)),0) as N5, "
                    sql &= "ISNULL(SUM(CAST(FS_06 AS float)),0) as N6, "
                    sql &= "ISNULL(SUM(CAST(FS_07 AS float)),0) as N7, "
                    sql &= "ISNULL(SUM(CAST(FS_08 AS float)),0) as N8, "
                    sql &= "ISNULL(SUM(CAST(FS_09 AS float)),0) as N9, "
                    sql &= "ISNULL(SUM(CAST(FS_10 AS float)),0) as N10, "
                    sql &= "ISNULL(SUM(CAST(FS_11 AS float)),0) as N11, "
                    sql &= "ISNULL(SUM(CAST(FS_12 AS float)),0) as N12, "
                    sql &= "ISNULL(SUM(CAST(FS_13 AS float)),0) as Total "
                    sql &= "FROM LocalStockPlan "
                    sql &= " Where Buyer = '" & xBuyer & "' "
                    sql &= " And Version = 98 "
                    sql &= " And GR_01 = '" & "LS" & "' "
                    If xBuyer = "FALL-TP000013" Then
                    Else
                        sql &= " And GR_08 = '" & "PS" & "' "
                    End If
                    Dim dt_FCTQty6 As DataTable = uDataBase.GetDataTable(sql)
                    If dt_FCTQty6.Rows.Count > 0 Then

                        Dim HAtc1 As TableCell = New TableCell
                        HAtc1.Text = Format(dt_FCTQty6.Rows(0).Item("N"), "###,###,##0")
                        HArow.Cells.Add(HAtc1)

                        Dim HAtc2 As TableCell = New TableCell
                        HAtc2.Text = Format(dt_FCTQty6.Rows(0).Item("N1"), "###,###,##0")
                        HArow.Cells.Add(HAtc2)

                        Dim HAtc3 As TableCell = New TableCell
                        HAtc3.Text = Format(dt_FCTQty6.Rows(0).Item("N2"), "###,###,##0")
                        HArow.Cells.Add(HAtc3)

                        Dim HAtc4 As TableCell = New TableCell
                        HAtc4.Text = Format(dt_FCTQty6.Rows(0).Item("N3"), "###,###,##0")
                        HArow.Cells.Add(HAtc4)

                        Dim HAtc5 As TableCell = New TableCell
                        HAtc5.Text = Format(dt_FCTQty6.Rows(0).Item("N4"), "###,###,##0")
                        HArow.Cells.Add(HAtc5)

                        Dim HAtc6 As TableCell = New TableCell
                        HAtc6.Text = Format(dt_FCTQty6.Rows(0).Item("N5"), "###,###,##0")
                        HArow.Cells.Add(HAtc6)

                        Dim HAtc7 As TableCell = New TableCell
                        HAtc7.Text = Format(dt_FCTQty6.Rows(0).Item("N6"), "###,###,##0")
                        HArow.Cells.Add(HAtc7)

                        Dim HAtc8 As TableCell = New TableCell
                        HAtc8.Text = Format(dt_FCTQty6.Rows(0).Item("N7"), "###,###,##0")
                        HArow.Cells.Add(HAtc8)

                        Dim HAtc9 As TableCell = New TableCell
                        HAtc9.Text = Format(dt_FCTQty6.Rows(0).Item("N8"), "###,###,##0")
                        HArow.Cells.Add(HAtc9)

                        Dim HAtcA As TableCell = New TableCell
                        HAtcA.Text = Format(dt_FCTQty6.Rows(0).Item("N9"), "###,###,##0")
                        HArow.Cells.Add(HAtcA)

                        Dim HAtcB As TableCell = New TableCell
                        HAtcB.Text = Format(dt_FCTQty6.Rows(0).Item("N10"), "###,###,##0")
                        HArow.Cells.Add(HAtcB)

                        Dim HAtcC As TableCell = New TableCell
                        HAtcC.Text = Format(dt_FCTQty6.Rows(0).Item("N11"), "###,###,##0")
                        HArow.Cells.Add(HAtcC)

                        Dim HAtcD As TableCell = New TableCell
                        HAtcD.Text = Format(dt_FCTQty6.Rows(0).Item("N12"), "###,###,##0")
                        HArow.Cells.Add(HAtcD)

                        Dim HAtcE As TableCell = New TableCell
                        HAtcE.Text = Format(dt_FCTQty6.Rows(0).Item("Total"), "###,###,##0")
                        HArow.Cells.Add(HAtcE)
                    End If
                    e.Row.Parent.Controls.Add(HArow)
                    '-- ----------------------------------------------------------------
                    '-- LS SLD QTY
                    '-- ----------------------------------------------------------------
                    Dim HBhd1 As TableCell = New TableCell
                    If xBuyer = "FALL-TP000013" Then
                        HBhd1.Text = "IPLSPLAN"
                    Else
                        HBhd1.Text = "SLD-IPLSPLAN"
                    End If
                    HBrow.Cells.Add(HBhd1)

                    sql = "SELECT "
                    sql &= "ISNULL(SUM(CAST(FS_00 AS float)),0) as N, "
                    sql &= "ISNULL(SUM(CAST(FS_01 AS float)),0) as N1, "
                    sql &= "ISNULL(SUM(CAST(FS_02 AS float)),0) as N2, "
                    sql &= "ISNULL(SUM(CAST(FS_03 AS float)),0) as N3, "
                    sql &= "ISNULL(SUM(CAST(FS_04 AS float)),0) as N4, "
                    sql &= "ISNULL(SUM(CAST(FS_05 AS float)),0) as N5, "
                    sql &= "ISNULL(SUM(CAST(FS_06 AS float)),0) as N6, "
                    sql &= "ISNULL(SUM(CAST(FS_07 AS float)),0) as N7, "
                    sql &= "ISNULL(SUM(CAST(FS_08 AS float)),0) as N8, "
                    sql &= "ISNULL(SUM(CAST(FS_09 AS float)),0) as N9, "
                    sql &= "ISNULL(SUM(CAST(FS_10 AS float)),0) as N10, "
                    sql &= "ISNULL(SUM(CAST(FS_11 AS float)),0) as N11, "
                    sql &= "ISNULL(SUM(CAST(FS_12 AS float)),0) as N12, "
                    sql &= "ISNULL(SUM(CAST(FS_13 AS float)),0) as Total "
                    sql &= "FROM LocalStockPlan "
                    sql &= " Where Buyer = '" & xBuyer & "' "
                    sql &= " And Version = 99 "
                    If xBuyer = "FALL-TP000013" Then
                    Else
                        sql &= " And GR_08 = '" & "PS" & "' "
                    End If
                    Dim dt_FCTQty7 As DataTable = uDataBase.GetDataTable(sql)
                    If dt_FCTQty7.Rows.Count > 0 Then

                        Dim HBtc1 As TableCell = New TableCell
                        HBtc1.Text = Format(dt_FCTQty7.Rows(0).Item("N"), "###,###,##0")
                        HBrow.Cells.Add(HBtc1)

                        Dim HBtc2 As TableCell = New TableCell
                        HBtc2.Text = Format(dt_FCTQty7.Rows(0).Item("N1"), "###,###,##0")
                        HBrow.Cells.Add(HBtc2)

                        Dim HBtc3 As TableCell = New TableCell
                        HBtc3.Text = Format(dt_FCTQty7.Rows(0).Item("N2"), "###,###,##0")
                        HBrow.Cells.Add(HBtc3)

                        Dim HBtc4 As TableCell = New TableCell
                        HBtc4.Text = Format(dt_FCTQty7.Rows(0).Item("N3"), "###,###,##0")
                        HBrow.Cells.Add(HBtc4)

                        Dim HBtc5 As TableCell = New TableCell
                        HBtc5.Text = Format(dt_FCTQty7.Rows(0).Item("N4"), "###,###,##0")
                        HBrow.Cells.Add(HBtc5)

                        Dim HBtc6 As TableCell = New TableCell
                        HBtc6.Text = Format(dt_FCTQty7.Rows(0).Item("N5"), "###,###,##0")
                        HBrow.Cells.Add(HBtc6)

                        Dim HBtc7 As TableCell = New TableCell
                        HBtc7.Text = Format(dt_FCTQty7.Rows(0).Item("N6"), "###,###,##0")
                        HBrow.Cells.Add(HBtc7)

                        Dim HBtc8 As TableCell = New TableCell
                        HBtc8.Text = Format(dt_FCTQty7.Rows(0).Item("N7"), "###,###,##0")
                        HBrow.Cells.Add(HBtc8)

                        Dim HBtc9 As TableCell = New TableCell
                        HBtc9.Text = Format(dt_FCTQty7.Rows(0).Item("N8"), "###,###,##0")
                        HBrow.Cells.Add(HBtc9)

                        Dim HBtcA As TableCell = New TableCell
                        HBtcA.Text = Format(dt_FCTQty7.Rows(0).Item("N9"), "###,###,##0")
                        HBrow.Cells.Add(HBtcA)

                        Dim HBtcB As TableCell = New TableCell
                        HBtcB.Text = Format(dt_FCTQty7.Rows(0).Item("N10"), "###,###,##0")
                        HBrow.Cells.Add(HBtcB)

                        Dim HBtcC As TableCell = New TableCell
                        HBtcC.Text = Format(dt_FCTQty7.Rows(0).Item("N11"), "###,###,##0")
                        HBrow.Cells.Add(HBtcC)

                        Dim HBtcD As TableCell = New TableCell
                        HBtcD.Text = Format(dt_FCTQty7.Rows(0).Item("N12"), "###,###,##0")
                        HBrow.Cells.Add(HBtcD)

                        Dim HBtcE As TableCell = New TableCell
                        HBtcE.Text = Format(dt_FCTQty7.Rows(0).Item("Total"), "###,###,##0")
                        HBrow.Cells.Add(HBtcE)
                    End If
                    e.Row.Parent.Controls.Add(HBrow)
                    '-- ----------------------------------------------------------------
                    '-- 差異
                    '-- 
                    '-- ----------------------------------------------------------------
                    Dim HChd1 As TableCell = New TableCell
                    HChd1.BackColor = Color.LightPink
                    If xBuyer = "FALL-TP000013" Then
                        HChd1.Text = "[IPLSPLAN]-- 差"
                    Else
                        HChd1.Text = "[IPLSPLAN-SLD]-- 差"
                    End If
                    HCrow.Cells.Add(HChd1)

                    Dim HCtc1 As TableCell = New TableCell
                    HCtc1.HorizontalAlign = HorizontalAlign.Right
                    HCtc1.BackColor = Color.LightPink
                    HCtc1.Text = Format(dt_FCTQty6.Rows(0).Item("N") - dt_FCTQty7.Rows(0).Item("N"), "###,###,##0")
                    HCrow.Cells.Add(HCtc1)

                    Dim HCtc2 As TableCell = New TableCell
                    HCtc2.HorizontalAlign = HorizontalAlign.Right
                    HCtc2.BackColor = Color.LightPink
                    HCtc2.Text = Format(dt_FCTQty6.Rows(0).Item("N1") - dt_FCTQty7.Rows(0).Item("N1"), "###,###,##0")
                    HCrow.Cells.Add(HCtc2)

                    Dim HCtc3 As TableCell = New TableCell
                    HCtc3.HorizontalAlign = HorizontalAlign.Right
                    HCtc3.BackColor = Color.LightPink
                    HCtc3.Text = Format(dt_FCTQty6.Rows(0).Item("N2") - dt_FCTQty7.Rows(0).Item("N2"), "###,###,##0")
                    HCrow.Cells.Add(HCtc3)

                    Dim HCtc4 As TableCell = New TableCell
                    HCtc4.HorizontalAlign = HorizontalAlign.Right
                    HCtc4.BackColor = Color.LightPink
                    HCtc4.Text = Format(dt_FCTQty6.Rows(0).Item("N3") - dt_FCTQty7.Rows(0).Item("N3"), "###,###,##0")
                    HCrow.Cells.Add(HCtc4)

                    Dim HCtc5 As TableCell = New TableCell
                    HCtc5.HorizontalAlign = HorizontalAlign.Right
                    HCtc5.BackColor = Color.LightPink
                    HCtc5.Text = Format(dt_FCTQty6.Rows(0).Item("N4") - dt_FCTQty7.Rows(0).Item("N4"), "###,###,##0")
                    HCrow.Cells.Add(HCtc5)

                    Dim HCtc6 As TableCell = New TableCell
                    HCtc6.HorizontalAlign = HorizontalAlign.Right
                    HCtc6.BackColor = Color.LightPink
                    HCtc6.Text = Format(dt_FCTQty6.Rows(0).Item("N5") - dt_FCTQty7.Rows(0).Item("N5"), "###,###,##0")
                    HCrow.Cells.Add(HCtc6)

                    Dim HCtc7 As TableCell = New TableCell
                    HCtc7.HorizontalAlign = HorizontalAlign.Right
                    HCtc7.BackColor = Color.LightPink
                    HCtc7.Text = Format(dt_FCTQty6.Rows(0).Item("N6") - dt_FCTQty7.Rows(0).Item("N6"), "###,###,##0")
                    HCrow.Cells.Add(HCtc7)

                    Dim HCtc8 As TableCell = New TableCell
                    HCtc8.HorizontalAlign = HorizontalAlign.Right
                    HCtc8.BackColor = Color.LightPink
                    HCtc8.Text = Format(dt_FCTQty6.Rows(0).Item("N7") - dt_FCTQty7.Rows(0).Item("N7"), "###,###,##0")
                    HCrow.Cells.Add(HCtc8)

                    Dim HCtc9 As TableCell = New TableCell
                    HCtc9.HorizontalAlign = HorizontalAlign.Right
                    HCtc9.BackColor = Color.LightPink
                    HCtc9.Text = Format(dt_FCTQty6.Rows(0).Item("N8") - dt_FCTQty7.Rows(0).Item("N8"), "###,###,##0")
                    HCrow.Cells.Add(HCtc9)

                    Dim HCtcA As TableCell = New TableCell
                    HCtcA.HorizontalAlign = HorizontalAlign.Right
                    HCtcA.BackColor = Color.LightPink
                    HCtcA.Text = Format(dt_FCTQty6.Rows(0).Item("N9") - dt_FCTQty7.Rows(0).Item("N9"), "###,###,##0")
                    HCrow.Cells.Add(HCtcA)

                    Dim HCtcB As TableCell = New TableCell
                    HCtcB.HorizontalAlign = HorizontalAlign.Right
                    HCtcB.BackColor = Color.LightPink
                    HCtcB.Text = Format(dt_FCTQty6.Rows(0).Item("N10") - dt_FCTQty7.Rows(0).Item("N10"), "###,###,##0")
                    HCrow.Cells.Add(HCtcB)

                    Dim HCtcC As TableCell = New TableCell
                    HCtcC.HorizontalAlign = HorizontalAlign.Right
                    HCtcC.BackColor = Color.LightPink
                    HCtcC.Text = Format(dt_FCTQty6.Rows(0).Item("N11") - dt_FCTQty7.Rows(0).Item("N11"), "###,###,##0")
                    HCrow.Cells.Add(HCtcC)

                    Dim HCtcD As TableCell = New TableCell
                    HCtcD.HorizontalAlign = HorizontalAlign.Right
                    HCtcD.BackColor = Color.LightPink
                    HCtcD.Text = Format(dt_FCTQty6.Rows(0).Item("N12") - dt_FCTQty7.Rows(0).Item("N12"), "###,###,##0")
                    HCrow.Cells.Add(HCtcD)

                    Dim HCtcE As TableCell = New TableCell
                    HCtcE.HorizontalAlign = HorizontalAlign.Right
                    HCtcE.BackColor = Color.LightPink
                    HCtcE.Text = Format(dt_FCTQty6.Rows(0).Item("Total") - dt_FCTQty7.Rows(0).Item("Total"), "###,###,##0")
                    HCrow.Cells.Add(HCtcE)
                    e.Row.Parent.Controls.Add(HCrow)

                    If dt_FCTQty6.Rows(0).Item("N") - dt_FCTQty7.Rows(0).Item("N") <> 0 Or _
                    dt_FCTQty6.Rows(0).Item("N1") - dt_FCTQty7.Rows(0).Item("N1") <> 0 Or _
                    dt_FCTQty6.Rows(0).Item("N2") - dt_FCTQty7.Rows(0).Item("N2") <> 0 Or _
                    dt_FCTQty6.Rows(0).Item("N3") - dt_FCTQty7.Rows(0).Item("N3") <> 0 Or _
                    dt_FCTQty6.Rows(0).Item("N4") - dt_FCTQty7.Rows(0).Item("N4") <> 0 Or _
                    dt_FCTQty6.Rows(0).Item("N5") - dt_FCTQty7.Rows(0).Item("N5") <> 0 Or _
                    dt_FCTQty6.Rows(0).Item("N6") - dt_FCTQty7.Rows(0).Item("N6") <> 0 Or _
                    dt_FCTQty6.Rows(0).Item("N7") - dt_FCTQty7.Rows(0).Item("N7") <> 0 Or _
                    dt_FCTQty6.Rows(0).Item("N8") - dt_FCTQty7.Rows(0).Item("N8") <> 0 Or _
                    dt_FCTQty6.Rows(0).Item("N9") - dt_FCTQty7.Rows(0).Item("N9") <> 0 Or _
                    dt_FCTQty6.Rows(0).Item("N10") - dt_FCTQty7.Rows(0).Item("N10") <> 0 Or _
                    dt_FCTQty6.Rows(0).Item("N11") - dt_FCTQty7.Rows(0).Item("N11") <> 0 Or _
                    dt_FCTQty6.Rows(0).Item("N12") - dt_FCTQty7.Rows(0).Item("N12") <> 0 Or _
                    dt_FCTQty6.Rows(0).Item("Total") - dt_FCTQty7.Rows(0).Item("Total") <> 0 Then
                        If DMessage.Text = "" Then DMessage.Text = "**[注意]發現數量不符，請再次確認客戶需求量QTY.  /IMPORTLS-SLD"
                    End If
                    '-- ----------------------------------------------------------------
                    '-- CH QTY
                    '-- ----------------------------------------------------------------
                    If xBuyer = "FALL-TP000013" Then
                    Else
                        Dim HDhd1 As TableCell = New TableCell
                        HDhd1.Text = "CH-LSPLAN"
                        HDrow.Cells.Add(HDhd1)

                        sql = "SELECT "
                        sql &= "ISNULL(SUM(CAST(FS_00 AS float)),0) as N, "
                        sql &= "ISNULL(SUM(CAST(FS_01 AS float)),0) as N1, "
                        sql &= "ISNULL(SUM(CAST(FS_02 AS float)),0) as N2, "
                        sql &= "ISNULL(SUM(CAST(FS_03 AS float)),0) as N3, "
                        sql &= "ISNULL(SUM(CAST(FS_04 AS float)),0) as N4, "
                        sql &= "ISNULL(SUM(CAST(FS_05 AS float)),0) as N5, "
                        sql &= "ISNULL(SUM(CAST(FS_06 AS float)),0) as N6, "
                        sql &= "ISNULL(SUM(CAST(FS_07 AS float)),0) as N7, "
                        sql &= "ISNULL(SUM(CAST(FS_08 AS float)),0) as N8, "
                        sql &= "ISNULL(SUM(CAST(FS_09 AS float)),0) as N9, "
                        sql &= "ISNULL(SUM(CAST(FS_10 AS float)),0) as N10, "
                        sql &= "ISNULL(SUM(CAST(FS_11 AS float)),0) as N11, "
                        sql &= "ISNULL(SUM(CAST(FS_12 AS float)),0) as N12, "
                        sql &= "ISNULL(SUM(CAST(FS_13 AS float)),0) as Total "
                        sql &= "FROM LocalStockPlan "
                        sql &= " Where Buyer = '" & xBuyer & "' "
                        sql &= " And Version = 98 "
                        sql &= " And GR_01 = '" & "LS" & "' "
                        sql &= " And GR_08 = '" & "CH" & "' "
                        Dim dt_FCTQty8 As DataTable = uDataBase.GetDataTable(sql)
                        If dt_FCTQty8.Rows.Count > 0 Then

                            Dim HDtc1 As TableCell = New TableCell
                            HDtc1.Text = Format(dt_FCTQty8.Rows(0).Item("N"), "###,###,##0.0")
                            HDrow.Cells.Add(HDtc1)

                            Dim HDtc2 As TableCell = New TableCell
                            HDtc2.Text = Format(dt_FCTQty8.Rows(0).Item("N1"), "###,###,##0.0")
                            HDrow.Cells.Add(HDtc2)

                            Dim HDtc3 As TableCell = New TableCell
                            HDtc3.Text = Format(dt_FCTQty8.Rows(0).Item("N2"), "###,###,##0.0")
                            HDrow.Cells.Add(HDtc3)

                            Dim HDtc4 As TableCell = New TableCell
                            HDtc4.Text = Format(dt_FCTQty8.Rows(0).Item("N3"), "###,###,##0.0")
                            HDrow.Cells.Add(HDtc4)

                            Dim HDtc5 As TableCell = New TableCell
                            HDtc5.Text = Format(dt_FCTQty8.Rows(0).Item("N4"), "###,###,##0.0")
                            HDrow.Cells.Add(HDtc5)

                            Dim HDtc6 As TableCell = New TableCell
                            HDtc6.Text = Format(dt_FCTQty8.Rows(0).Item("N5"), "###,###,##0.0")
                            HDrow.Cells.Add(HDtc6)

                            Dim HDtc7 As TableCell = New TableCell
                            HDtc7.Text = Format(dt_FCTQty8.Rows(0).Item("N6"), "###,###,##0.0")
                            HDrow.Cells.Add(HDtc7)

                            Dim HDtc8 As TableCell = New TableCell
                            HDtc8.Text = Format(dt_FCTQty8.Rows(0).Item("N7"), "###,###,##0.0")
                            HDrow.Cells.Add(HDtc8)

                            Dim HDtc9 As TableCell = New TableCell
                            HDtc9.Text = Format(dt_FCTQty8.Rows(0).Item("N8"), "###,###,##0.0")
                            HDrow.Cells.Add(HDtc9)

                            Dim HDtcA As TableCell = New TableCell
                            HDtcA.Text = Format(dt_FCTQty8.Rows(0).Item("N9"), "###,###,##0.0")
                            HDrow.Cells.Add(HDtcA)

                            Dim HDtcB As TableCell = New TableCell
                            HDtcB.Text = Format(dt_FCTQty8.Rows(0).Item("N10"), "###,###,##0.0")
                            HDrow.Cells.Add(HDtcB)

                            Dim HDtcC As TableCell = New TableCell
                            HDtcC.Text = Format(dt_FCTQty8.Rows(0).Item("N11"), "###,###,##0.0")
                            HDrow.Cells.Add(HDtcC)

                            Dim HDtcD As TableCell = New TableCell
                            HDtcD.Text = Format(dt_FCTQty8.Rows(0).Item("N12"), "###,###,##0.0")
                            HDrow.Cells.Add(HDtcD)

                            Dim HDtcE As TableCell = New TableCell
                            HDtcE.Text = Format(dt_FCTQty8.Rows(0).Item("Total"), "###,###,##0.0")
                            HDrow.Cells.Add(HDtcE)
                        End If
                        e.Row.Parent.Controls.Add(HDrow)
                        '-- ----------------------------------------------------------------
                        '-- LS CH QTY
                        '-- ----------------------------------------------------------------
                        Dim HEhd1 As TableCell = New TableCell
                        HEhd1.Text = "CH-IPLSPLAN"
                        HErow.Cells.Add(HEhd1)

                        sql = "SELECT "
                        sql &= "ISNULL(SUM(CAST(FS_00 AS float)),0) as N, "
                        sql &= "ISNULL(SUM(CAST(FS_01 AS float)),0) as N1, "
                        sql &= "ISNULL(SUM(CAST(FS_02 AS float)),0) as N2, "
                        sql &= "ISNULL(SUM(CAST(FS_03 AS float)),0) as N3, "
                        sql &= "ISNULL(SUM(CAST(FS_04 AS float)),0) as N4, "
                        sql &= "ISNULL(SUM(CAST(FS_05 AS float)),0) as N5, "
                        sql &= "ISNULL(SUM(CAST(FS_06 AS float)),0) as N6, "
                        sql &= "ISNULL(SUM(CAST(FS_07 AS float)),0) as N7, "
                        sql &= "ISNULL(SUM(CAST(FS_08 AS float)),0) as N8, "
                        sql &= "ISNULL(SUM(CAST(FS_09 AS float)),0) as N9, "
                        sql &= "ISNULL(SUM(CAST(FS_10 AS float)),0) as N10, "
                        sql &= "ISNULL(SUM(CAST(FS_11 AS float)),0) as N11, "
                        sql &= "ISNULL(SUM(CAST(FS_12 AS float)),0) as N12, "
                        sql &= "ISNULL(SUM(CAST(FS_13 AS float)),0) as Total "
                        sql &= "FROM LocalStockPlan "
                        sql &= " Where Buyer = '" & xBuyer & "' "
                        sql &= " And Version = 99 "
                        sql &= " And GR_08 = '" & "CH" & "' "
                        Dim dt_FCTQty9 As DataTable = uDataBase.GetDataTable(sql)
                        If dt_FCTQty9.Rows.Count > 0 Then

                            Dim HEtc1 As TableCell = New TableCell
                            HEtc1.Text = Format(dt_FCTQty9.Rows(0).Item("N"), "###,###,##0.0")
                            HErow.Cells.Add(HEtc1)

                            Dim HEtc2 As TableCell = New TableCell
                            HEtc2.Text = Format(dt_FCTQty9.Rows(0).Item("N1"), "###,###,##0.0")
                            HErow.Cells.Add(HEtc2)

                            Dim HEtc3 As TableCell = New TableCell
                            HEtc3.Text = Format(dt_FCTQty9.Rows(0).Item("N2"), "###,###,##0.0")
                            HErow.Cells.Add(HEtc3)

                            Dim HEtc4 As TableCell = New TableCell
                            HEtc4.Text = Format(dt_FCTQty9.Rows(0).Item("N3"), "###,###,##0.0")
                            HErow.Cells.Add(HEtc4)

                            Dim HEtc5 As TableCell = New TableCell
                            HEtc5.Text = Format(dt_FCTQty9.Rows(0).Item("N4"), "###,###,##0.0")
                            HErow.Cells.Add(HEtc5)

                            Dim HEtc6 As TableCell = New TableCell
                            HEtc6.Text = Format(dt_FCTQty9.Rows(0).Item("N5"), "###,###,##0.0")
                            HErow.Cells.Add(HEtc6)

                            Dim HEtc7 As TableCell = New TableCell
                            HEtc7.Text = Format(dt_FCTQty9.Rows(0).Item("N6"), "###,###,##0.0")
                            HErow.Cells.Add(HEtc7)

                            Dim HEtc8 As TableCell = New TableCell
                            HEtc8.Text = Format(dt_FCTQty9.Rows(0).Item("N7"), "###,###,##0.0")
                            HErow.Cells.Add(HEtc8)

                            Dim HEtc9 As TableCell = New TableCell
                            HEtc9.Text = Format(dt_FCTQty9.Rows(0).Item("N8"), "###,###,##0.0")
                            HErow.Cells.Add(HEtc9)

                            Dim HEtcA As TableCell = New TableCell
                            HEtcA.Text = Format(dt_FCTQty9.Rows(0).Item("N9"), "###,###,##0.0")
                            HErow.Cells.Add(HEtcA)

                            Dim HEtcB As TableCell = New TableCell
                            HEtcB.Text = Format(dt_FCTQty9.Rows(0).Item("N10"), "###,###,##0.0")
                            HErow.Cells.Add(HEtcB)

                            Dim HEtcC As TableCell = New TableCell
                            HEtcC.Text = Format(dt_FCTQty9.Rows(0).Item("N11"), "###,###,##0.0")
                            HErow.Cells.Add(HEtcC)

                            Dim HEtcD As TableCell = New TableCell
                            HEtcD.Text = Format(dt_FCTQty9.Rows(0).Item("N12"), "###,###,##0.0")
                            HErow.Cells.Add(HEtcD)

                            Dim HEtcE As TableCell = New TableCell
                            HEtcE.Text = Format(dt_FCTQty9.Rows(0).Item("Total"), "###,###,##0.0")
                            HErow.Cells.Add(HEtcE)
                        End If
                        e.Row.Parent.Controls.Add(HErow)
                        '-- ----------------------------------------------------------------
                        '-- 差異
                        '-- 
                        '-- ----------------------------------------------------------------
                        Dim HFhd1 As TableCell = New TableCell
                        HFhd1.BackColor = Color.LightPink
                        HFhd1.Text = "[IPLSPLAN-CH]-- 差"
                        HFrow.Cells.Add(HFhd1)

                        Dim HFtc1 As TableCell = New TableCell
                        HFtc1.HorizontalAlign = HorizontalAlign.Right
                        HFtc1.BackColor = Color.LightPink
                        HFtc1.Text = Format(dt_FCTQty8.Rows(0).Item("N") - dt_FCTQty9.Rows(0).Item("N"), "###,###,##0.0")
                        HFrow.Cells.Add(HFtc1)

                        Dim HFtc2 As TableCell = New TableCell
                        HFtc2.HorizontalAlign = HorizontalAlign.Right
                        HFtc2.BackColor = Color.LightPink
                        HFtc2.Text = Format(dt_FCTQty8.Rows(0).Item("N1") - dt_FCTQty9.Rows(0).Item("N1"), "###,###,##0.0")
                        HFrow.Cells.Add(HFtc2)

                        Dim HFtc3 As TableCell = New TableCell
                        HFtc3.HorizontalAlign = HorizontalAlign.Right
                        HFtc3.BackColor = Color.LightPink
                        HFtc3.Text = Format(dt_FCTQty8.Rows(0).Item("N2") - dt_FCTQty9.Rows(0).Item("N2"), "###,###,##0.0")
                        HFrow.Cells.Add(HFtc3)

                        Dim HFtc4 As TableCell = New TableCell
                        HFtc4.HorizontalAlign = HorizontalAlign.Right
                        HFtc4.BackColor = Color.LightPink
                        HFtc4.Text = Format(dt_FCTQty8.Rows(0).Item("N3") - dt_FCTQty9.Rows(0).Item("N3"), "###,###,##0.0")
                        HFrow.Cells.Add(HFtc4)

                        Dim HFtc5 As TableCell = New TableCell
                        HFtc5.HorizontalAlign = HorizontalAlign.Right
                        HFtc5.BackColor = Color.LightPink
                        HFtc5.Text = Format(dt_FCTQty8.Rows(0).Item("N4") - dt_FCTQty9.Rows(0).Item("N4"), "###,###,##0.0")
                        HFrow.Cells.Add(HFtc5)

                        Dim HFtc6 As TableCell = New TableCell
                        HFtc6.HorizontalAlign = HorizontalAlign.Right
                        HFtc6.BackColor = Color.LightPink
                        HFtc6.Text = Format(dt_FCTQty8.Rows(0).Item("N5") - dt_FCTQty9.Rows(0).Item("N5"), "###,###,##0.0")
                        HFrow.Cells.Add(HFtc6)

                        Dim HFtc7 As TableCell = New TableCell
                        HFtc7.HorizontalAlign = HorizontalAlign.Right
                        HFtc7.BackColor = Color.LightPink
                        HFtc7.Text = Format(dt_FCTQty8.Rows(0).Item("N6") - dt_FCTQty9.Rows(0).Item("N6"), "###,###,##0.0")
                        HFrow.Cells.Add(HFtc7)

                        Dim HFtc8 As TableCell = New TableCell
                        HFtc8.HorizontalAlign = HorizontalAlign.Right
                        HFtc8.BackColor = Color.LightPink
                        HFtc8.Text = Format(dt_FCTQty8.Rows(0).Item("N7") - dt_FCTQty9.Rows(0).Item("N7"), "###,###,##0.0")
                        HFrow.Cells.Add(HFtc8)

                        Dim HFtc9 As TableCell = New TableCell
                        HFtc9.HorizontalAlign = HorizontalAlign.Right
                        HFtc9.BackColor = Color.LightPink
                        HFtc9.Text = Format(dt_FCTQty8.Rows(0).Item("N8") - dt_FCTQty9.Rows(0).Item("N8"), "###,###,##0.0")
                        HFrow.Cells.Add(HFtc9)

                        Dim HFtcA As TableCell = New TableCell
                        HFtcA.HorizontalAlign = HorizontalAlign.Right
                        HFtcA.BackColor = Color.LightPink
                        HFtcA.Text = Format(dt_FCTQty8.Rows(0).Item("N9") - dt_FCTQty9.Rows(0).Item("N9"), "###,###,##0.0")
                        HFrow.Cells.Add(HFtcA)

                        Dim HFtcB As TableCell = New TableCell
                        HFtcB.HorizontalAlign = HorizontalAlign.Right
                        HFtcB.BackColor = Color.LightPink
                        HFtcB.Text = Format(dt_FCTQty8.Rows(0).Item("N10") - dt_FCTQty9.Rows(0).Item("N10"), "###,###,##0.0")
                        HFrow.Cells.Add(HFtcB)

                        Dim HFtcC As TableCell = New TableCell
                        HFtcC.HorizontalAlign = HorizontalAlign.Right
                        HFtcC.BackColor = Color.LightPink
                        HFtcC.Text = Format(dt_FCTQty8.Rows(0).Item("N11") - dt_FCTQty9.Rows(0).Item("N11"), "###,###,##0.0")
                        HFrow.Cells.Add(HFtcC)

                        Dim HFtcD As TableCell = New TableCell
                        HFtcD.HorizontalAlign = HorizontalAlign.Right
                        HFtcD.BackColor = Color.LightPink
                        HFtcD.Text = Format(dt_FCTQty8.Rows(0).Item("N12") - dt_FCTQty9.Rows(0).Item("N12"), "###,###,##0.0")
                        HFrow.Cells.Add(HFtcD)

                        Dim HFtcE As TableCell = New TableCell
                        HFtcE.HorizontalAlign = HorizontalAlign.Right
                        HFtcE.BackColor = Color.LightPink
                        HFtcE.Text = Format(dt_FCTQty8.Rows(0).Item("Total") - dt_FCTQty9.Rows(0).Item("Total"), "###,###,##0.0")
                        HFrow.Cells.Add(HFtcE)
                        e.Row.Parent.Controls.Add(HFrow)

                        If Format(dt_FCTQty8.Rows(0).Item("N"), "###,###,##0.0") <> Format(dt_FCTQty9.Rows(0).Item("N"), "###,###,##0.0") Or _
                        Format(dt_FCTQty8.Rows(0).Item("N1"), "###,###,##0.0") <> Format(dt_FCTQty9.Rows(0).Item("N1"), "###,###,##0.0") Or _
                        Format(dt_FCTQty8.Rows(0).Item("N2"), "###,###,##0.0") <> Format(dt_FCTQty9.Rows(0).Item("N2"), "###,###,##0.0") Or _
                        Format(dt_FCTQty8.Rows(0).Item("N3"), "###,###,##0.0") <> Format(dt_FCTQty9.Rows(0).Item("N3"), "###,###,##0.0") Or _
                        Format(dt_FCTQty8.Rows(0).Item("N4"), "###,###,##0.0") <> Format(dt_FCTQty9.Rows(0).Item("N4"), "###,###,##0.0") Or _
                        Format(dt_FCTQty8.Rows(0).Item("N5"), "###,###,##0.0") <> Format(dt_FCTQty9.Rows(0).Item("N5"), "###,###,##0.0") Or _
                        Format(dt_FCTQty8.Rows(0).Item("N6"), "###,###,##0.0") <> Format(dt_FCTQty9.Rows(0).Item("N6"), "###,###,##0.0") Or _
                        Format(dt_FCTQty8.Rows(0).Item("N7"), "###,###,##0.0") <> Format(dt_FCTQty9.Rows(0).Item("N7"), "###,###,##0.0") Or _
                        Format(dt_FCTQty8.Rows(0).Item("N8"), "###,###,##0.0") <> Format(dt_FCTQty9.Rows(0).Item("N8"), "###,###,##0.0") Or _
                        Format(dt_FCTQty8.Rows(0).Item("N9"), "###,###,##0.0") <> Format(dt_FCTQty9.Rows(0).Item("N9"), "###,###,##0.0") Or _
                        Format(dt_FCTQty8.Rows(0).Item("N10"), "###,###,##0.0") <> Format(dt_FCTQty9.Rows(0).Item("N10"), "###,###,##0.0") Or _
                        Format(dt_FCTQty8.Rows(0).Item("N11"), "###,###,##0.0") <> Format(dt_FCTQty9.Rows(0).Item("N11"), "###,###,##0.0") Or _
                        Format(dt_FCTQty8.Rows(0).Item("N12"), "###,###,##0.0") <> Format(dt_FCTQty9.Rows(0).Item("N12"), "###,###,##0.0") Or _
                        Format(dt_FCTQty8.Rows(0).Item("Total"), "###,###,##0.0") <> Format(dt_FCTQty9.Rows(0).Item("Total"), "###,###,##0.0") Then
                            If DMessage.Text = "" Then DMessage.Text = "**[注意]發現數量不符，請再次確認客戶需求量QTY.  /IMPORTLS-CH"
                        End If
                    End If
                End If
                '-- ----------------------------------------------------------------
                '-- SLD QTY
                '-- ----------------------------------------------------------------
                If xFun = "BYLSPLAN" Or xFun = "EDI" Then
                    Dim HGHJ1 As TableCell = New TableCell
                    If xBuyer = "FALL-TP000013" Then
                        HGHJ1.Text = "IPLSPLAN"
                    Else
                        HGHJ1.Text = "SLD-IPLSPLAN"
                    End If
                    HGrow.Cells.Add(HGHJ1)

                    sql = "SELECT "
                    sql &= "ISNULL(SUM(CAST(FS_00 AS float)),0) as N, "
                    sql &= "ISNULL(SUM(CAST(FS_01 AS float)),0) as N1, "
                    sql &= "ISNULL(SUM(CAST(FS_02 AS float)),0) as N2, "
                    sql &= "ISNULL(SUM(CAST(FS_03 AS float)),0) as N3, "
                    sql &= "ISNULL(SUM(CAST(FS_04 AS float)),0) as N4, "
                    sql &= "ISNULL(SUM(CAST(FS_05 AS float)),0) as N5, "
                    sql &= "ISNULL(SUM(CAST(FS_06 AS float)),0) as N6, "
                    sql &= "ISNULL(SUM(CAST(FS_07 AS float)),0) as N7, "
                    sql &= "ISNULL(SUM(CAST(FS_08 AS float)),0) as N8, "
                    sql &= "ISNULL(SUM(CAST(FS_09 AS float)),0) as N9, "
                    sql &= "ISNULL(SUM(CAST(FS_10 AS float)),0) as N10, "
                    sql &= "ISNULL(SUM(CAST(FS_11 AS float)),0) as N11, "
                    sql &= "ISNULL(SUM(CAST(FS_12 AS float)),0) as N12, "
                    sql &= "ISNULL(SUM(CAST(FS_13 AS float)),0) as Total "
                    sql &= "FROM LocalStockPlan "
                    sql &= " WHERE Buyer = '" & xBuyer & "' "
                    sql &= " And Version = 99 "
                    If xBuyer = "FALL-TP000013" Then
                    Else
                        sql &= " And GR_08 = '" & "PS" & "' "
                    End If
                    Dim dt_FCTQtyA As DataTable = uDataBase.GetDataTable(sql)
                    If dt_FCTQtyA.Rows.Count > 0 Then

                        Dim HGtc1 As TableCell = New TableCell
                        HGtc1.Text = Format(dt_FCTQtyA.Rows(0).Item("N"), "###,###,##0")
                        HGrow.Cells.Add(HGtc1)

                        Dim HGtc2 As TableCell = New TableCell
                        HGtc2.Text = Format(dt_FCTQtyA.Rows(0).Item("N1"), "###,###,##0")
                        HGrow.Cells.Add(HGtc2)

                        Dim HGtc3 As TableCell = New TableCell
                        HGtc3.Text = Format(dt_FCTQtyA.Rows(0).Item("N2"), "###,###,##0")
                        HGrow.Cells.Add(HGtc3)

                        Dim HGtc4 As TableCell = New TableCell
                        HGtc4.Text = Format(dt_FCTQtyA.Rows(0).Item("N3"), "###,###,##0")
                        HGrow.Cells.Add(HGtc4)

                        Dim HGtc5 As TableCell = New TableCell
                        HGtc5.Text = Format(dt_FCTQtyA.Rows(0).Item("N4"), "###,###,##0")
                        HGrow.Cells.Add(HGtc5)

                        Dim HGtc6 As TableCell = New TableCell
                        HGtc6.Text = Format(dt_FCTQtyA.Rows(0).Item("N5"), "###,###,##0")
                        HGrow.Cells.Add(HGtc6)

                        Dim HGtc7 As TableCell = New TableCell
                        HGtc7.Text = Format(dt_FCTQtyA.Rows(0).Item("N6"), "###,###,##0")
                        HGrow.Cells.Add(HGtc7)

                        Dim HGtc8 As TableCell = New TableCell
                        HGtc8.Text = Format(dt_FCTQtyA.Rows(0).Item("N7"), "###,###,##0")
                        HGrow.Cells.Add(HGtc8)

                        Dim HGtc9 As TableCell = New TableCell
                        HGtc9.Text = Format(dt_FCTQtyA.Rows(0).Item("N8"), "###,###,##0")
                        HGrow.Cells.Add(HGtc9)

                        Dim HGtcA As TableCell = New TableCell
                        HGtcA.Text = Format(dt_FCTQtyA.Rows(0).Item("N9"), "###,###,##0")
                        HGrow.Cells.Add(HGtcA)

                        Dim HGtcB As TableCell = New TableCell
                        HGtcB.Text = Format(dt_FCTQtyA.Rows(0).Item("N10"), "###,###,##0")
                        HGrow.Cells.Add(HGtcB)

                        Dim HGtcC As TableCell = New TableCell
                        HGtcC.Text = Format(dt_FCTQtyA.Rows(0).Item("N11"), "###,###,##0")
                        HGrow.Cells.Add(HGtcC)

                        Dim HGtcD As TableCell = New TableCell
                        HGtcD.Text = Format(dt_FCTQtyA.Rows(0).Item("N12"), "###,###,##0")
                        HGrow.Cells.Add(HGtcD)

                        Dim HGtcE As TableCell = New TableCell
                        HGtcE.Text = Format(dt_FCTQtyA.Rows(0).Item("Total"), "###,###,##0")
                        HGrow.Cells.Add(HGtcE)
                    End If
                    e.Row.Parent.Controls.Add(HGrow)
                    '-- ----------------------------------------------------------------
                    '-- LS SLD QTY
                    '-- ----------------------------------------------------------------
                    Dim HHHJ1 As TableCell = New TableCell
                    If xBuyer = "FALL-TP000013" Then
                        HHHJ1.Text = "BYLSPLAN"
                    Else
                        HHHJ1.Text = "SLD-BYLSPLAN"
                    End If
                    HHrow.Cells.Add(HHHJ1)

                    sql = "SELECT "
                    sql &= "ISNULL(SUM(CAST(FS_00 AS float)),0) as N, "
                    sql &= "ISNULL(SUM(CAST(FS_01 AS float)),0) as N1, "
                    sql &= "ISNULL(SUM(CAST(FS_02 AS float)),0) as N2, "
                    sql &= "ISNULL(SUM(CAST(FS_03 AS float)),0) as N3, "
                    sql &= "ISNULL(SUM(CAST(FS_04 AS float)),0) as N4, "
                    sql &= "ISNULL(SUM(CAST(FS_05 AS float)),0) as N5, "
                    sql &= "ISNULL(SUM(CAST(FS_06 AS float)),0) as N6, "
                    sql &= "ISNULL(SUM(CAST(FS_07 AS float)),0) as N7, "
                    sql &= "ISNULL(SUM(CAST(FS_08 AS float)),0) as N8, "
                    sql &= "ISNULL(SUM(CAST(FS_09 AS float)),0) as N9, "
                    sql &= "ISNULL(SUM(CAST(FS_10 AS float)),0) as N10, "
                    sql &= "ISNULL(SUM(CAST(FS_11 AS float)),0) as N11, "
                    sql &= "ISNULL(SUM(CAST(FS_12 AS float)),0) as N12, "
                    sql &= "ISNULL(SUM(CAST(FS_13 AS float)),0) as Total "
                    sql &= "FROM BuyerLocalStockPlan "
                    sql &= " Where BuyerGroup = '" & xBuyerGroup & "' "
                    sql &= " And Version = 99 "
                    If xBuyer = "FALL-TP000013" Then
                    Else
                        sql &= " And GR_08 = '" & "PS" & "' "
                    End If
                    Dim dt_FCTQtyB As DataTable = uDataBase.GetDataTable(sql)
                    If dt_FCTQtyB.Rows.Count > 0 Then

                        Dim HHtc1 As TableCell = New TableCell
                        HHtc1.Text = Format(dt_FCTQtyB.Rows(0).Item("N"), "###,###,##0")
                        HHrow.Cells.Add(HHtc1)

                        Dim HHtc2 As TableCell = New TableCell
                        HHtc2.Text = Format(dt_FCTQtyB.Rows(0).Item("N1"), "###,###,##0")
                        HHrow.Cells.Add(HHtc2)

                        Dim HHtc3 As TableCell = New TableCell
                        HHtc3.Text = Format(dt_FCTQtyB.Rows(0).Item("N2"), "###,###,##0")
                        HHrow.Cells.Add(HHtc3)

                        Dim HHtc4 As TableCell = New TableCell
                        HHtc4.Text = Format(dt_FCTQtyB.Rows(0).Item("N3"), "###,###,##0")
                        HHrow.Cells.Add(HHtc4)

                        Dim HHtc5 As TableCell = New TableCell
                        HHtc5.Text = Format(dt_FCTQtyB.Rows(0).Item("N4"), "###,###,##0")
                        HHrow.Cells.Add(HHtc5)

                        Dim HHtc6 As TableCell = New TableCell
                        HHtc6.Text = Format(dt_FCTQtyB.Rows(0).Item("N5"), "###,###,##0")
                        HHrow.Cells.Add(HHtc6)

                        Dim HHtc7 As TableCell = New TableCell
                        HHtc7.Text = Format(dt_FCTQtyB.Rows(0).Item("N6"), "###,###,##0")
                        HHrow.Cells.Add(HHtc7)

                        Dim HHtc8 As TableCell = New TableCell
                        HHtc8.Text = Format(dt_FCTQtyB.Rows(0).Item("N7"), "###,###,##0")
                        HHrow.Cells.Add(HHtc8)

                        Dim HHtc9 As TableCell = New TableCell
                        HHtc9.Text = Format(dt_FCTQtyB.Rows(0).Item("N8"), "###,###,##0")
                        HHrow.Cells.Add(HHtc9)

                        Dim HHtcA As TableCell = New TableCell
                        HHtcA.Text = Format(dt_FCTQtyB.Rows(0).Item("N9"), "###,###,##0")
                        HHrow.Cells.Add(HHtcA)

                        Dim HHtcB As TableCell = New TableCell
                        HHtcB.Text = Format(dt_FCTQtyB.Rows(0).Item("N10"), "###,###,##0")
                        HHrow.Cells.Add(HHtcB)

                        Dim HHtcC As TableCell = New TableCell
                        HHtcC.Text = Format(dt_FCTQtyB.Rows(0).Item("N11"), "###,###,##0")
                        HHrow.Cells.Add(HHtcC)

                        Dim HHtcD As TableCell = New TableCell
                        HHtcD.Text = Format(dt_FCTQtyB.Rows(0).Item("N12"), "###,###,##0")
                        HHrow.Cells.Add(HHtcD)

                        Dim HHtcE As TableCell = New TableCell
                        HHtcE.Text = Format(dt_FCTQtyB.Rows(0).Item("Total"), "###,###,##0")
                        HHrow.Cells.Add(HHtcE)
                    End If
                    e.Row.Parent.Controls.Add(HHrow)
                    '-- ----------------------------------------------------------------
                    '-- 差異
                    '-- 
                    '-- ----------------------------------------------------------------
                    Dim HIHJ1 As TableCell = New TableCell
                    HIHJ1.BackColor = Color.LightPink
                    If xBuyer = "FALL-TP000013" Then
                        HIHJ1.Text = "[BYLSPLAN]-- 差"
                    Else
                        HIHJ1.Text = "[BYLSPLAN-SLD]-- 差"
                    End If
                    HIrow.Cells.Add(HIHJ1)

                    Dim HItc1 As TableCell = New TableCell
                    HItc1.HorizontalAlign = HorizontalAlign.Right
                    HItc1.BackColor = Color.LightPink
                    HItc1.Text = Format(dt_FCTQtyA.Rows(0).Item("N") - dt_FCTQtyB.Rows(0).Item("N"), "###,###,##0")
                    HIrow.Cells.Add(HItc1)

                    Dim HItc2 As TableCell = New TableCell
                    HItc2.HorizontalAlign = HorizontalAlign.Right
                    HItc2.BackColor = Color.LightPink
                    HItc2.Text = Format(dt_FCTQtyA.Rows(0).Item("N1") - dt_FCTQtyB.Rows(0).Item("N1"), "###,###,##0")
                    HIrow.Cells.Add(HItc2)

                    Dim HItc3 As TableCell = New TableCell
                    HItc3.HorizontalAlign = HorizontalAlign.Right
                    HItc3.BackColor = Color.LightPink
                    HItc3.Text = Format(dt_FCTQtyA.Rows(0).Item("N2") - dt_FCTQtyB.Rows(0).Item("N2"), "###,###,##0")
                    HIrow.Cells.Add(HItc3)

                    Dim HItc4 As TableCell = New TableCell
                    HItc4.HorizontalAlign = HorizontalAlign.Right
                    HItc4.BackColor = Color.LightPink
                    HItc4.Text = Format(dt_FCTQtyA.Rows(0).Item("N3") - dt_FCTQtyB.Rows(0).Item("N3"), "###,###,##0")
                    HIrow.Cells.Add(HItc4)

                    Dim HItc5 As TableCell = New TableCell
                    HItc5.HorizontalAlign = HorizontalAlign.Right
                    HItc5.BackColor = Color.LightPink
                    HItc5.Text = Format(dt_FCTQtyA.Rows(0).Item("N4") - dt_FCTQtyB.Rows(0).Item("N4"), "###,###,##0")
                    HIrow.Cells.Add(HItc5)

                    Dim HItc6 As TableCell = New TableCell
                    HItc6.HorizontalAlign = HorizontalAlign.Right
                    HItc6.BackColor = Color.LightPink
                    HItc6.Text = Format(dt_FCTQtyA.Rows(0).Item("N5") - dt_FCTQtyB.Rows(0).Item("N5"), "###,###,##0")
                    HIrow.Cells.Add(HItc6)

                    Dim HItc7 As TableCell = New TableCell
                    HItc7.HorizontalAlign = HorizontalAlign.Right
                    HItc7.BackColor = Color.LightPink
                    HItc7.Text = Format(dt_FCTQtyA.Rows(0).Item("N6") - dt_FCTQtyB.Rows(0).Item("N6"), "###,###,##0")
                    HIrow.Cells.Add(HItc7)

                    Dim HItc8 As TableCell = New TableCell
                    HItc8.HorizontalAlign = HorizontalAlign.Right
                    HItc8.BackColor = Color.LightPink
                    HItc8.Text = Format(dt_FCTQtyA.Rows(0).Item("N7") - dt_FCTQtyB.Rows(0).Item("N7"), "###,###,##0")
                    HIrow.Cells.Add(HItc8)

                    Dim HItc9 As TableCell = New TableCell
                    HItc9.HorizontalAlign = HorizontalAlign.Right
                    HItc9.BackColor = Color.LightPink
                    HItc9.Text = Format(dt_FCTQtyA.Rows(0).Item("N8") - dt_FCTQtyB.Rows(0).Item("N8"), "###,###,##0")
                    HIrow.Cells.Add(HItc9)

                    Dim HItcA As TableCell = New TableCell
                    HItcA.HorizontalAlign = HorizontalAlign.Right
                    HItcA.BackColor = Color.LightPink
                    HItcA.Text = Format(dt_FCTQtyA.Rows(0).Item("N9") - dt_FCTQtyB.Rows(0).Item("N9"), "###,###,##0")
                    HIrow.Cells.Add(HItcA)

                    Dim HItcB As TableCell = New TableCell
                    HItcB.HorizontalAlign = HorizontalAlign.Right
                    HItcB.BackColor = Color.LightPink
                    HItcB.Text = Format(dt_FCTQtyA.Rows(0).Item("N10") - dt_FCTQtyB.Rows(0).Item("N10"), "###,###,##0")
                    HIrow.Cells.Add(HItcB)

                    Dim HItcC As TableCell = New TableCell
                    HItcC.HorizontalAlign = HorizontalAlign.Right
                    HItcC.BackColor = Color.LightPink
                    HItcC.Text = Format(dt_FCTQtyA.Rows(0).Item("N11") - dt_FCTQtyB.Rows(0).Item("N11"), "###,###,##0")
                    HIrow.Cells.Add(HItcC)

                    Dim HItcD As TableCell = New TableCell
                    HItcD.HorizontalAlign = HorizontalAlign.Right
                    HItcD.BackColor = Color.LightPink
                    HItcD.Text = Format(dt_FCTQtyA.Rows(0).Item("N12") - dt_FCTQtyB.Rows(0).Item("N12"), "###,###,##0")
                    HIrow.Cells.Add(HItcD)

                    Dim HItcE As TableCell = New TableCell
                    HItcE.HorizontalAlign = HorizontalAlign.Right
                    HItcE.BackColor = Color.LightPink
                    HItcE.Text = Format(dt_FCTQtyA.Rows(0).Item("Total") - dt_FCTQtyB.Rows(0).Item("Total"), "###,###,##0")
                    HIrow.Cells.Add(HItcE)
                    e.Row.Parent.Controls.Add(HIrow)

                    If dt_FCTQtyA.Rows(0).Item("N") - dt_FCTQtyB.Rows(0).Item("N") <> 0 Or _
                    dt_FCTQtyA.Rows(0).Item("N1") - dt_FCTQtyB.Rows(0).Item("N1") <> 0 Or _
                    dt_FCTQtyA.Rows(0).Item("N2") - dt_FCTQtyB.Rows(0).Item("N2") <> 0 Or _
                    dt_FCTQtyA.Rows(0).Item("N3") - dt_FCTQtyB.Rows(0).Item("N3") <> 0 Or _
                    dt_FCTQtyA.Rows(0).Item("N4") - dt_FCTQtyB.Rows(0).Item("N4") <> 0 Or _
                    dt_FCTQtyA.Rows(0).Item("N5") - dt_FCTQtyB.Rows(0).Item("N5") <> 0 Or _
                    dt_FCTQtyA.Rows(0).Item("N6") - dt_FCTQtyB.Rows(0).Item("N6") <> 0 Or _
                    dt_FCTQtyA.Rows(0).Item("N7") - dt_FCTQtyB.Rows(0).Item("N7") <> 0 Or _
                    dt_FCTQtyA.Rows(0).Item("N8") - dt_FCTQtyB.Rows(0).Item("N8") <> 0 Or _
                    dt_FCTQtyA.Rows(0).Item("N9") - dt_FCTQtyB.Rows(0).Item("N9") <> 0 Or _
                    dt_FCTQtyA.Rows(0).Item("N10") - dt_FCTQtyB.Rows(0).Item("N10") <> 0 Or _
                    dt_FCTQtyA.Rows(0).Item("N11") - dt_FCTQtyB.Rows(0).Item("N11") <> 0 Or _
                    dt_FCTQtyA.Rows(0).Item("N12") - dt_FCTQtyB.Rows(0).Item("N12") <> 0 Or _
                    dt_FCTQtyA.Rows(0).Item("Total") - dt_FCTQtyB.Rows(0).Item("Total") <> 0 Then
                        If DMessage.Text = "" Then DMessage.Text = "**[注意]發現數量不符，請再次確認客戶需求量QTY.  /BYLSPLAN-SLD"
                    End If
                    '-- ----------------------------------------------------------------
                    '-- CH QTY
                    '-- ----------------------------------------------------------------
                    If xBuyer = "FALL-TP000013" Then
                    Else
                        Dim HJHJ1 As TableCell = New TableCell
                        HJHJ1.Text = "CH-IPLSPLAN"
                        HJrow.Cells.Add(HJHJ1)

                        sql = "SELECT "
                        sql &= "ISNULL(SUM(CAST(FS_00 AS float)),0) as N, "
                        sql &= "ISNULL(SUM(CAST(FS_01 AS float)),0) as N1, "
                        sql &= "ISNULL(SUM(CAST(FS_02 AS float)),0) as N2, "
                        sql &= "ISNULL(SUM(CAST(FS_03 AS float)),0) as N3, "
                        sql &= "ISNULL(SUM(CAST(FS_04 AS float)),0) as N4, "
                        sql &= "ISNULL(SUM(CAST(FS_05 AS float)),0) as N5, "
                        sql &= "ISNULL(SUM(CAST(FS_06 AS float)),0) as N6, "
                        sql &= "ISNULL(SUM(CAST(FS_07 AS float)),0) as N7, "
                        sql &= "ISNULL(SUM(CAST(FS_08 AS float)),0) as N8, "
                        sql &= "ISNULL(SUM(CAST(FS_09 AS float)),0) as N9, "
                        sql &= "ISNULL(SUM(CAST(FS_10 AS float)),0) as N10, "
                        sql &= "ISNULL(SUM(CAST(FS_11 AS float)),0) as N11, "
                        sql &= "ISNULL(SUM(CAST(FS_12 AS float)),0) as N12, "
                        sql &= "ISNULL(SUM(CAST(FS_13 AS float)),0) as Total "
                        sql &= "FROM LocalStockPlan "
                        sql &= " Where Buyer = '" & xBuyer & "' "
                        sql &= " And Version = 99 "
                        sql &= " And GR_08 = '" & "CH" & "' "
                        Dim dt_FCTQtyC As DataTable = uDataBase.GetDataTable(sql)
                        If dt_FCTQtyC.Rows.Count > 0 Then

                            Dim HJtc1 As TableCell = New TableCell
                            HJtc1.Text = Format(dt_FCTQtyC.Rows(0).Item("N"), "###,###,##0.0")
                            HJrow.Cells.Add(HJtc1)

                            Dim HJtc2 As TableCell = New TableCell
                            HJtc2.Text = Format(dt_FCTQtyC.Rows(0).Item("N1"), "###,###,##0.0")
                            HJrow.Cells.Add(HJtc2)

                            Dim HJtc3 As TableCell = New TableCell
                            HJtc3.Text = Format(dt_FCTQtyC.Rows(0).Item("N2"), "###,###,##0.0")
                            HJrow.Cells.Add(HJtc3)

                            Dim HJtc4 As TableCell = New TableCell
                            HJtc4.Text = Format(dt_FCTQtyC.Rows(0).Item("N3"), "###,###,##0.0")
                            HJrow.Cells.Add(HJtc4)

                            Dim HJtc5 As TableCell = New TableCell
                            HJtc5.Text = Format(dt_FCTQtyC.Rows(0).Item("N4"), "###,###,##0.0")
                            HJrow.Cells.Add(HJtc5)

                            Dim HJtc6 As TableCell = New TableCell
                            HJtc6.Text = Format(dt_FCTQtyC.Rows(0).Item("N5"), "###,###,##0.0")
                            HJrow.Cells.Add(HJtc6)

                            Dim HJtc7 As TableCell = New TableCell
                            HJtc7.Text = Format(dt_FCTQtyC.Rows(0).Item("N6"), "###,###,##0.0")
                            HJrow.Cells.Add(HJtc7)

                            Dim HJtc8 As TableCell = New TableCell
                            HJtc8.Text = Format(dt_FCTQtyC.Rows(0).Item("N7"), "###,###,##0.0")
                            HJrow.Cells.Add(HJtc8)

                            Dim HJtc9 As TableCell = New TableCell
                            HJtc9.Text = Format(dt_FCTQtyC.Rows(0).Item("N8"), "###,###,##0.0")
                            HJrow.Cells.Add(HJtc9)

                            Dim HJtcA As TableCell = New TableCell
                            HJtcA.Text = Format(dt_FCTQtyC.Rows(0).Item("N9"), "###,###,##0.0")
                            HJrow.Cells.Add(HJtcA)

                            Dim HJtcB As TableCell = New TableCell
                            HJtcB.Text = Format(dt_FCTQtyC.Rows(0).Item("N10"), "###,###,##0.0")
                            HJrow.Cells.Add(HJtcB)

                            Dim HJtcC As TableCell = New TableCell
                            HJtcC.Text = Format(dt_FCTQtyC.Rows(0).Item("N11"), "###,###,##0.0")
                            HJrow.Cells.Add(HJtcC)

                            Dim HJtcD As TableCell = New TableCell
                            HJtcD.Text = Format(dt_FCTQtyC.Rows(0).Item("N12"), "###,###,##0.0")
                            HJrow.Cells.Add(HJtcD)

                            Dim HJtcE As TableCell = New TableCell
                            HJtcE.Text = Format(dt_FCTQtyC.Rows(0).Item("Total"), "###,###,##0.0")
                            HJrow.Cells.Add(HJtcE)
                        End If
                        e.Row.Parent.Controls.Add(HJrow)
                        '-- ----------------------------------------------------------------
                        '-- LS CH QTY
                        '-- ----------------------------------------------------------------
                        Dim HKHJ1 As TableCell = New TableCell
                        HKHJ1.Text = "CH-BYLSPLAN"
                        HKrow.Cells.Add(HKHJ1)

                        sql = "SELECT "
                        sql &= "ISNULL(SUM(CAST(FS_00 AS float)),0) as N, "
                        sql &= "ISNULL(SUM(CAST(FS_01 AS float)),0) as N1, "
                        sql &= "ISNULL(SUM(CAST(FS_02 AS float)),0) as N2, "
                        sql &= "ISNULL(SUM(CAST(FS_03 AS float)),0) as N3, "
                        sql &= "ISNULL(SUM(CAST(FS_04 AS float)),0) as N4, "
                        sql &= "ISNULL(SUM(CAST(FS_05 AS float)),0) as N5, "
                        sql &= "ISNULL(SUM(CAST(FS_06 AS float)),0) as N6, "
                        sql &= "ISNULL(SUM(CAST(FS_07 AS float)),0) as N7, "
                        sql &= "ISNULL(SUM(CAST(FS_08 AS float)),0) as N8, "
                        sql &= "ISNULL(SUM(CAST(FS_09 AS float)),0) as N9, "
                        sql &= "ISNULL(SUM(CAST(FS_10 AS float)),0) as N10, "
                        sql &= "ISNULL(SUM(CAST(FS_11 AS float)),0) as N11, "
                        sql &= "ISNULL(SUM(CAST(FS_12 AS float)),0) as N12, "
                        sql &= "ISNULL(SUM(CAST(FS_13 AS float)),0) as Total "
                        sql &= "FROM BuyerLocalStockPlan "
                        sql &= " Where BuyerGroup = '" & xBuyerGroup & "' "
                        sql &= " And Version = 99 "
                        sql &= " And GR_08 = '" & "CH" & "' "
                        Dim dt_FCTQtyD As DataTable = uDataBase.GetDataTable(sql)
                        If dt_FCTQtyD.Rows.Count > 0 Then

                            Dim HKtc1 As TableCell = New TableCell
                            HKtc1.Text = Format(dt_FCTQtyD.Rows(0).Item("N"), "###,###,##0.0")
                            HKrow.Cells.Add(HKtc1)

                            Dim HKtc2 As TableCell = New TableCell
                            HKtc2.Text = Format(dt_FCTQtyD.Rows(0).Item("N1"), "###,###,##0.0")
                            HKrow.Cells.Add(HKtc2)

                            Dim HKtc3 As TableCell = New TableCell
                            HKtc3.Text = Format(dt_FCTQtyD.Rows(0).Item("N2"), "###,###,##0.0")
                            HKrow.Cells.Add(HKtc3)

                            Dim HKtc4 As TableCell = New TableCell
                            HKtc4.Text = Format(dt_FCTQtyD.Rows(0).Item("N3"), "###,###,##0.0")
                            HKrow.Cells.Add(HKtc4)

                            Dim HKtc5 As TableCell = New TableCell
                            HKtc5.Text = Format(dt_FCTQtyD.Rows(0).Item("N4"), "###,###,##0.0")
                            HKrow.Cells.Add(HKtc5)

                            Dim HKtc6 As TableCell = New TableCell
                            HKtc6.Text = Format(dt_FCTQtyD.Rows(0).Item("N5"), "###,###,##0.0")
                            HKrow.Cells.Add(HKtc6)

                            Dim HKtc7 As TableCell = New TableCell
                            HKtc7.Text = Format(dt_FCTQtyD.Rows(0).Item("N6"), "###,###,##0.0")
                            HKrow.Cells.Add(HKtc7)

                            Dim HKtc8 As TableCell = New TableCell
                            HKtc8.Text = Format(dt_FCTQtyD.Rows(0).Item("N7"), "###,###,##0.0")
                            HKrow.Cells.Add(HKtc8)

                            Dim HKtc9 As TableCell = New TableCell
                            HKtc9.Text = Format(dt_FCTQtyD.Rows(0).Item("N8"), "###,###,##0.0")
                            HKrow.Cells.Add(HKtc9)

                            Dim HKtcA As TableCell = New TableCell
                            HKtcA.Text = Format(dt_FCTQtyD.Rows(0).Item("N9"), "###,###,##0.0")
                            HKrow.Cells.Add(HKtcA)

                            Dim HKtcB As TableCell = New TableCell
                            HKtcB.Text = Format(dt_FCTQtyD.Rows(0).Item("N10"), "###,###,##0.0")
                            HKrow.Cells.Add(HKtcB)

                            Dim HKtcC As TableCell = New TableCell
                            HKtcC.Text = Format(dt_FCTQtyD.Rows(0).Item("N11"), "###,###,##0.0")
                            HKrow.Cells.Add(HKtcC)

                            Dim HKtcD As TableCell = New TableCell
                            HKtcD.Text = Format(dt_FCTQtyD.Rows(0).Item("N12"), "###,###,##0.0")
                            HKrow.Cells.Add(HKtcD)

                            Dim HKtcE As TableCell = New TableCell
                            HKtcE.Text = Format(dt_FCTQtyD.Rows(0).Item("Total"), "###,###,##0.0")
                            HKrow.Cells.Add(HKtcE)
                        End If
                        e.Row.Parent.Controls.Add(HKrow)
                        '-- ----------------------------------------------------------------
                        '-- 差異
                        '-- 
                        '-- ----------------------------------------------------------------
                        Dim HLHJ1 As TableCell = New TableCell
                        HLHJ1.BackColor = Color.LightPink
                        HLHJ1.Text = "[BYLSPLAN-CH]-- 差"
                        HLrow.Cells.Add(HLHJ1)

                        Dim HLtc1 As TableCell = New TableCell
                        HLtc1.HorizontalAlign = HorizontalAlign.Right
                        HLtc1.BackColor = Color.LightPink
                        HLtc1.Text = Format(dt_FCTQtyC.Rows(0).Item("N") - dt_FCTQtyD.Rows(0).Item("N"), "###,###,##0.0")
                        HLrow.Cells.Add(HLtc1)

                        Dim HLtc2 As TableCell = New TableCell
                        HLtc2.HorizontalAlign = HorizontalAlign.Right
                        HLtc2.BackColor = Color.LightPink
                        HLtc2.Text = Format(dt_FCTQtyC.Rows(0).Item("N1") - dt_FCTQtyD.Rows(0).Item("N1"), "###,###,##0.0")
                        HLrow.Cells.Add(HLtc2)

                        Dim HLtc3 As TableCell = New TableCell
                        HLtc3.HorizontalAlign = HorizontalAlign.Right
                        HLtc3.BackColor = Color.LightPink
                        HLtc3.Text = Format(dt_FCTQtyC.Rows(0).Item("N2") - dt_FCTQtyD.Rows(0).Item("N2"), "###,###,##0.0")
                        HLrow.Cells.Add(HLtc3)

                        Dim HLtc4 As TableCell = New TableCell
                        HLtc4.HorizontalAlign = HorizontalAlign.Right
                        HLtc4.BackColor = Color.LightPink
                        HLtc4.Text = Format(dt_FCTQtyC.Rows(0).Item("N3") - dt_FCTQtyD.Rows(0).Item("N3"), "###,###,##0.0")
                        HLrow.Cells.Add(HLtc4)

                        Dim HLtc5 As TableCell = New TableCell
                        HLtc5.HorizontalAlign = HorizontalAlign.Right
                        HLtc5.BackColor = Color.LightPink
                        HLtc5.Text = Format(dt_FCTQtyC.Rows(0).Item("N4") - dt_FCTQtyD.Rows(0).Item("N4"), "###,###,##0.0")
                        HLrow.Cells.Add(HLtc5)

                        Dim HLtc6 As TableCell = New TableCell
                        HLtc6.HorizontalAlign = HorizontalAlign.Right
                        HLtc6.BackColor = Color.LightPink
                        HLtc6.Text = Format(dt_FCTQtyC.Rows(0).Item("N5") - dt_FCTQtyD.Rows(0).Item("N5"), "###,###,##0.0")
                        HLrow.Cells.Add(HLtc6)

                        Dim HLtc7 As TableCell = New TableCell
                        HLtc7.HorizontalAlign = HorizontalAlign.Right
                        HLtc7.BackColor = Color.LightPink
                        HLtc7.Text = Format(dt_FCTQtyC.Rows(0).Item("N6") - dt_FCTQtyD.Rows(0).Item("N6"), "###,###,##0.0")
                        HLrow.Cells.Add(HLtc7)

                        Dim HLtc8 As TableCell = New TableCell
                        HLtc8.HorizontalAlign = HorizontalAlign.Right
                        HLtc8.BackColor = Color.LightPink
                        HLtc8.Text = Format(dt_FCTQtyC.Rows(0).Item("N7") - dt_FCTQtyD.Rows(0).Item("N7"), "###,###,##0.0")
                        HLrow.Cells.Add(HLtc8)

                        Dim HLtc9 As TableCell = New TableCell
                        HLtc9.HorizontalAlign = HorizontalAlign.Right
                        HLtc9.BackColor = Color.LightPink
                        HLtc9.Text = Format(dt_FCTQtyC.Rows(0).Item("N8") - dt_FCTQtyD.Rows(0).Item("N8"), "###,###,##0.0")
                        HLrow.Cells.Add(HLtc9)

                        Dim HLtcA As TableCell = New TableCell
                        HLtcA.HorizontalAlign = HorizontalAlign.Right
                        HLtcA.BackColor = Color.LightPink
                        HLtcA.Text = Format(dt_FCTQtyC.Rows(0).Item("N9") - dt_FCTQtyD.Rows(0).Item("N9"), "###,###,##0.0")
                        HLrow.Cells.Add(HLtcA)

                        Dim HLtcB As TableCell = New TableCell
                        HLtcB.HorizontalAlign = HorizontalAlign.Right
                        HLtcB.BackColor = Color.LightPink
                        HLtcB.Text = Format(dt_FCTQtyC.Rows(0).Item("N10") - dt_FCTQtyD.Rows(0).Item("N10"), "###,###,##0.0")
                        HLrow.Cells.Add(HLtcB)

                        Dim HLtcC As TableCell = New TableCell
                        HLtcC.HorizontalAlign = HorizontalAlign.Right
                        HLtcC.BackColor = Color.LightPink
                        HLtcC.Text = Format(dt_FCTQtyC.Rows(0).Item("N11") - dt_FCTQtyD.Rows(0).Item("N11"), "###,###,##0.0")
                        HLrow.Cells.Add(HLtcC)

                        Dim HLtcD As TableCell = New TableCell
                        HLtcD.HorizontalAlign = HorizontalAlign.Right
                        HLtcD.BackColor = Color.LightPink
                        HLtcD.Text = Format(dt_FCTQtyC.Rows(0).Item("N12") - dt_FCTQtyD.Rows(0).Item("N12"), "###,###,##0.0")
                        HLrow.Cells.Add(HLtcD)

                        Dim HLtcE As TableCell = New TableCell
                        HLtcE.HorizontalAlign = HorizontalAlign.Right
                        HLtcE.BackColor = Color.LightPink
                        HLtcE.Text = Format(dt_FCTQtyC.Rows(0).Item("Total") - dt_FCTQtyD.Rows(0).Item("Total"), "###,###,##0.0")
                        HLrow.Cells.Add(HLtcE)
                        e.Row.Parent.Controls.Add(HLrow)

                        If Format(dt_FCTQtyC.Rows(0).Item("N"), "###,###,##0.0") <> Format(dt_FCTQtyD.Rows(0).Item("N"), "###,###,##0.0") Or _
                        Format(dt_FCTQtyC.Rows(0).Item("N1"), "###,###,##0.0") <> Format(dt_FCTQtyD.Rows(0).Item("N1"), "###,###,##0.0") Or _
                        Format(dt_FCTQtyC.Rows(0).Item("N2"), "###,###,##0.0") <> Format(dt_FCTQtyD.Rows(0).Item("N2"), "###,###,##0.0") Or _
                        Format(dt_FCTQtyC.Rows(0).Item("N3"), "###,###,##0.0") <> Format(dt_FCTQtyD.Rows(0).Item("N3"), "###,###,##0.0") Or _
                        Format(dt_FCTQtyC.Rows(0).Item("N4"), "###,###,##0.0") <> Format(dt_FCTQtyD.Rows(0).Item("N4"), "###,###,##0.0") Or _
                        Format(dt_FCTQtyC.Rows(0).Item("N5"), "###,###,##0.0") <> Format(dt_FCTQtyD.Rows(0).Item("N5"), "###,###,##0.0") Or _
                        Format(dt_FCTQtyC.Rows(0).Item("N6"), "###,###,##0.0") <> Format(dt_FCTQtyD.Rows(0).Item("N6"), "###,###,##0.0") Or _
                        Format(dt_FCTQtyC.Rows(0).Item("N7"), "###,###,##0.0") <> Format(dt_FCTQtyD.Rows(0).Item("N7"), "###,###,##0.0") Or _
                        Format(dt_FCTQtyC.Rows(0).Item("N8"), "###,###,##0.0") <> Format(dt_FCTQtyD.Rows(0).Item("N8"), "###,###,##0.0") Or _
                        Format(dt_FCTQtyC.Rows(0).Item("N9"), "###,###,##0.0") <> Format(dt_FCTQtyD.Rows(0).Item("N9"), "###,###,##0.0") Or _
                        Format(dt_FCTQtyC.Rows(0).Item("N10"), "###,###,##0.0") <> Format(dt_FCTQtyD.Rows(0).Item("N10"), "###,###,##0.0") Or _
                        Format(dt_FCTQtyC.Rows(0).Item("N11"), "###,###,##0.0") <> Format(dt_FCTQtyD.Rows(0).Item("N11"), "###,###,##0.0") Or _
                        Format(dt_FCTQtyC.Rows(0).Item("N12"), "###,###,##0.0") <> Format(dt_FCTQtyD.Rows(0).Item("N12"), "###,###,##0.0") Or _
                        Format(dt_FCTQtyC.Rows(0).Item("Total"), "###,###,##0.0") <> Format(dt_FCTQtyD.Rows(0).Item("Total"), "###,###,##0.0") Then
                            If DMessage.Text = "" Then DMessage.Text = "**[注意]發現數量不符，請再次確認客戶需求量QTY.  /BYLSPLAN-CH"
                        End If
                    End If
                End If

                'START
                '-- ----------------------------------------------------------------
                '-- SLD QTY
                '-- ----------------------------------------------------------------
                If xFun = "EDI" Then
                    Dim HMHP1 As TableCell = New TableCell
                    If xBuyer = "FALL-TP000013" Then
                        HMHP1.Text = "BYLSPLAN"
                    Else
                        HMHP1.Text = "SLD-BYLSPLAN"
                    End If
                    HMrow.Cells.Add(HMHP1)

                    sql = "SELECT "
                    sql &= "ISNULL(SUM(CAST(FS_00 AS float)),0) as N, "
                    sql &= "ISNULL(SUM(CAST(FS_01 AS float)),0) as N1, "
                    sql &= "ISNULL(SUM(CAST(FS_02 AS float)),0) as N2, "
                    sql &= "ISNULL(SUM(CAST(FS_03 AS float)),0) as N3, "
                    sql &= "ISNULL(SUM(CAST(FS_04 AS float)),0) as N4, "
                    sql &= "ISNULL(SUM(CAST(FS_05 AS float)),0) as N5, "
                    sql &= "ISNULL(SUM(CAST(FS_06 AS float)),0) as N6, "
                    sql &= "ISNULL(SUM(CAST(FS_07 AS float)),0) as N7, "
                    sql &= "ISNULL(SUM(CAST(FS_08 AS float)),0) as N8, "
                    sql &= "ISNULL(SUM(CAST(FS_09 AS float)),0) as N9, "
                    sql &= "ISNULL(SUM(CAST(FS_10 AS float)),0) as N10, "
                    sql &= "ISNULL(SUM(CAST(FS_11 AS float)),0) as N11, "
                    sql &= "ISNULL(SUM(CAST(FS_12 AS float)),0) as N12, "
                    sql &= "ISNULL(SUM(CAST(FS_13 AS float)),0) as Total "
                    sql &= "FROM LF_BuyerLocalStockPlan "
                    sql &= " Where BuyerGroup = '" & xBuyerGroup & "' "
                    sql &= " And Version = 99 "
                    If xBuyer = "FALL-TP000013" Then
                    Else
                        sql &= " And GR_08 = '" & "PS" & "' "
                    End If
                    sql &= " And LOGID = '" & xLogID & "' "
                    Dim dt_FCTQtyE As DataTable = uDataBase.GetDataTable(sql)
                    If dt_FCTQtyE.Rows.Count > 0 Then

                        Dim HMtc1 As TableCell = New TableCell
                        HMtc1.Text = Format(dt_FCTQtyE.Rows(0).Item("N"), "###,###,##0")
                        HMrow.Cells.Add(HMtc1)

                        Dim HMtc2 As TableCell = New TableCell
                        HMtc2.Text = Format(dt_FCTQtyE.Rows(0).Item("N1"), "###,###,##0")
                        HMrow.Cells.Add(HMtc2)

                        Dim HMtc3 As TableCell = New TableCell
                        HMtc3.Text = Format(dt_FCTQtyE.Rows(0).Item("N2"), "###,###,##0")
                        HMrow.Cells.Add(HMtc3)

                        Dim HMtc4 As TableCell = New TableCell
                        HMtc4.Text = Format(dt_FCTQtyE.Rows(0).Item("N3"), "###,###,##0")
                        HMrow.Cells.Add(HMtc4)

                        Dim HMtc5 As TableCell = New TableCell
                        HMtc5.Text = Format(dt_FCTQtyE.Rows(0).Item("N4"), "###,###,##0")
                        HMrow.Cells.Add(HMtc5)

                        Dim HMtc6 As TableCell = New TableCell
                        HMtc6.Text = Format(dt_FCTQtyE.Rows(0).Item("N5"), "###,###,##0")
                        HMrow.Cells.Add(HMtc6)

                        Dim HMtc7 As TableCell = New TableCell
                        HMtc7.Text = Format(dt_FCTQtyE.Rows(0).Item("N6"), "###,###,##0")
                        HMrow.Cells.Add(HMtc7)

                        Dim HMtc8 As TableCell = New TableCell
                        HMtc8.Text = Format(dt_FCTQtyE.Rows(0).Item("N7"), "###,###,##0")
                        HMrow.Cells.Add(HMtc8)

                        Dim HMtc9 As TableCell = New TableCell
                        HMtc9.Text = Format(dt_FCTQtyE.Rows(0).Item("N8"), "###,###,##0")
                        HMrow.Cells.Add(HMtc9)

                        Dim HMtcA As TableCell = New TableCell
                        HMtcA.Text = Format(dt_FCTQtyE.Rows(0).Item("N9"), "###,###,##0")
                        HMrow.Cells.Add(HMtcA)

                        Dim HMtcB As TableCell = New TableCell
                        HMtcB.Text = Format(dt_FCTQtyE.Rows(0).Item("N10"), "###,###,##0")
                        HMrow.Cells.Add(HMtcB)

                        Dim HMtcC As TableCell = New TableCell
                        HMtcC.Text = Format(dt_FCTQtyE.Rows(0).Item("N11"), "###,###,##0")
                        HMrow.Cells.Add(HMtcC)

                        Dim HMtcD As TableCell = New TableCell
                        HMtcD.Text = Format(dt_FCTQtyE.Rows(0).Item("N12"), "###,###,##0")
                        HMrow.Cells.Add(HMtcD)

                        Dim HMtcE As TableCell = New TableCell
                        HMtcE.Text = Format(dt_FCTQtyE.Rows(0).Item("Total"), "###,###,##0")
                        HMrow.Cells.Add(HMtcE)
                    End If
                    e.Row.Parent.Controls.Add(HMrow)
                    '-- ----------------------------------------------------------------
                    '-- LS SLD QTY
                    '-- ----------------------------------------------------------------
                    Dim HNHP1 As TableCell = New TableCell
                    If xBuyer = "FALL-TP000013" Then
                        HNHP1.Text = "EDI"
                    Else
                        HNHP1.Text = "SLD-EDI"
                    End If
                    HNrow.Cells.Add(HNHP1)

                    sql = "SELECT "
                    sql &= "ISNULL(SUM(CAST(ProdQty AS float)),0) as Total "
                    sql &= "FROM B_LS2EDIInterface "
                    sql &= " Where BuyerGroup = '" & xBuyerGroup & "' "
                    If xBuyer = "FALL-TP000013" Then
                    Else
                        sql &= " And PartType = '" & "PS" & "' "
                    End If
                    Dim dt_FCTQtyF As DataTable = uDataBase.GetDataTable(sql)
                    If dt_FCTQtyF.Rows.Count > 0 Then

                        Dim HNtc1 As TableCell = New TableCell
                        HNtc1.Text = ""
                        HNrow.Cells.Add(HNtc1)

                        Dim HNtc2 As TableCell = New TableCell
                        HNtc2.Text = ""
                        HNrow.Cells.Add(HNtc2)

                        Dim HNtc3 As TableCell = New TableCell
                        HNtc3.Text = ""
                        HNrow.Cells.Add(HNtc3)

                        Dim HNtc4 As TableCell = New TableCell
                        HNtc4.Text = ""
                        HNrow.Cells.Add(HNtc4)

                        Dim HNtc5 As TableCell = New TableCell
                        HNtc5.Text = ""
                        HNrow.Cells.Add(HNtc5)

                        Dim HNtc6 As TableCell = New TableCell
                        HNtc6.Text = ""
                        HNrow.Cells.Add(HNtc6)

                        Dim HNtc7 As TableCell = New TableCell
                        HNtc7.Text = ""
                        HNrow.Cells.Add(HNtc7)

                        Dim HNtc8 As TableCell = New TableCell
                        HNtc8.Text = ""
                        HNrow.Cells.Add(HNtc8)

                        Dim HNtc9 As TableCell = New TableCell
                        HNtc9.Text = ""
                        HNrow.Cells.Add(HNtc9)

                        Dim HNtcA As TableCell = New TableCell
                        HNtcA.Text = ""
                        HNrow.Cells.Add(HNtcA)

                        Dim HNtcB As TableCell = New TableCell
                        HNtcB.Text = ""
                        HNrow.Cells.Add(HNtcB)

                        Dim HNtcC As TableCell = New TableCell
                        HNtcC.Text = ""
                        HNrow.Cells.Add(HNtcC)

                        Dim HNtcD As TableCell = New TableCell
                        HNtcD.Text = ""
                        HNrow.Cells.Add(HNtcD)

                        Dim HNtcE As TableCell = New TableCell
                        HNtcE.Text = Format(dt_FCTQtyF.Rows(0).Item("Total"), "###,###,##0")
                        HNrow.Cells.Add(HNtcE)
                    End If
                    e.Row.Parent.Controls.Add(HNrow)
                    '-- ----------------------------------------------------------------
                    '-- 差異
                    '-- 
                    '-- ----------------------------------------------------------------
                    Dim HOHP1 As TableCell = New TableCell
                    HOHP1.BackColor = Color.LightPink
                    If xBuyer = "FALL-TP000013" Then
                        HOHP1.Text = "[EDI]-- 差"
                    Else
                        HOHP1.Text = "[EDI-SLD]-- 差"
                    End If
                    HOrow.Cells.Add(HOHP1)

                    Dim HOtc1 As TableCell = New TableCell
                    HOtc1.HorizontalAlign = HorizontalAlign.Right
                    HOtc1.BackColor = Color.LightPink
                    HOtc1.Text = ""
                    HOrow.Cells.Add(HOtc1)

                    Dim HOtc2 As TableCell = New TableCell
                    HOtc2.HorizontalAlign = HorizontalAlign.Right
                    HOtc2.BackColor = Color.LightPink
                    HOtc2.Text = ""
                    HOrow.Cells.Add(HOtc2)

                    Dim HOtc3 As TableCell = New TableCell
                    HOtc3.HorizontalAlign = HorizontalAlign.Right
                    HOtc3.BackColor = Color.LightPink
                    HOtc3.Text = ""
                    HOrow.Cells.Add(HOtc3)

                    Dim HOtc4 As TableCell = New TableCell
                    HOtc4.HorizontalAlign = HorizontalAlign.Right
                    HOtc4.BackColor = Color.LightPink
                    HOtc4.Text = ""
                    HOrow.Cells.Add(HOtc4)

                    Dim HOtc5 As TableCell = New TableCell
                    HOtc5.HorizontalAlign = HorizontalAlign.Right
                    HOtc5.BackColor = Color.LightPink
                    HOtc5.Text = ""
                    HOrow.Cells.Add(HOtc5)

                    Dim HOtc6 As TableCell = New TableCell
                    HOtc6.HorizontalAlign = HorizontalAlign.Right
                    HOtc6.BackColor = Color.LightPink
                    HOtc6.Text = ""
                    HOrow.Cells.Add(HOtc6)

                    Dim HOtc7 As TableCell = New TableCell
                    HOtc7.HorizontalAlign = HorizontalAlign.Right
                    HOtc7.BackColor = Color.LightPink
                    HOtc7.Text = ""
                    HOrow.Cells.Add(HOtc7)

                    Dim HOtc8 As TableCell = New TableCell
                    HOtc8.HorizontalAlign = HorizontalAlign.Right
                    HOtc8.BackColor = Color.LightPink
                    HOtc8.Text = ""
                    HOrow.Cells.Add(HOtc8)

                    Dim HOtc9 As TableCell = New TableCell
                    HOtc9.HorizontalAlign = HorizontalAlign.Right
                    HOtc9.BackColor = Color.LightPink
                    HOtc9.Text = ""
                    HOrow.Cells.Add(HOtc9)

                    Dim HOtcA As TableCell = New TableCell
                    HOtcA.HorizontalAlign = HorizontalAlign.Right
                    HOtcA.BackColor = Color.LightPink
                    HOtcA.Text = ""
                    HOrow.Cells.Add(HOtcA)

                    Dim HOtcB As TableCell = New TableCell
                    HOtcB.HorizontalAlign = HorizontalAlign.Right
                    HOtcB.BackColor = Color.LightPink
                    HOtcB.Text = ""
                    HOrow.Cells.Add(HOtcB)

                    Dim HOtcC As TableCell = New TableCell
                    HOtcC.HorizontalAlign = HorizontalAlign.Right
                    HOtcC.BackColor = Color.LightPink
                    HOtcC.Text = ""
                    HOrow.Cells.Add(HOtcC)

                    Dim HOtcD As TableCell = New TableCell
                    HOtcD.HorizontalAlign = HorizontalAlign.Right
                    HOtcD.BackColor = Color.LightPink
                    HOtcD.Text = ""
                    HOrow.Cells.Add(HOtcD)

                    Dim HOtcE As TableCell = New TableCell
                    HOtcE.HorizontalAlign = HorizontalAlign.Right
                    HOtcE.BackColor = Color.LightPink
                    HOtcE.Text = Format(dt_FCTQtyE.Rows(0).Item("Total") - dt_FCTQtyF.Rows(0).Item("Total"), "###,###,##0")
                    HOrow.Cells.Add(HOtcE)
                    e.Row.Parent.Controls.Add(HOrow)

                    If dt_FCTQtyE.Rows(0).Item("Total") - dt_FCTQtyF.Rows(0).Item("Total") <> 0 Then
                        If DMessage.Text = "" Then DMessage.Text = "**[注意]發現數量不符，請再次確認客戶需求量QTY.  /BYLSPLAN-SLD"
                    End If
                    '-- ----------------------------------------------------------------
                    '-- CH QTY
                    '-- ----------------------------------------------------------------
                    If xBuyer = "FALL-TP000013" Then
                    Else
                        Dim HPHP1 As TableCell = New TableCell
                        HPHP1.Text = "CH-BYLSPLAN"
                        HProw.Cells.Add(HPHP1)

                        sql = "SELECT "
                        sql &= "ISNULL(SUM(CAST(FS_00 AS float)),0) as N, "
                        sql &= "ISNULL(SUM(CAST(FS_01 AS float)),0) as N1, "
                        sql &= "ISNULL(SUM(CAST(FS_02 AS float)),0) as N2, "
                        sql &= "ISNULL(SUM(CAST(FS_03 AS float)),0) as N3, "
                        sql &= "ISNULL(SUM(CAST(FS_04 AS float)),0) as N4, "
                        sql &= "ISNULL(SUM(CAST(FS_05 AS float)),0) as N5, "
                        sql &= "ISNULL(SUM(CAST(FS_06 AS float)),0) as N6, "
                        sql &= "ISNULL(SUM(CAST(FS_07 AS float)),0) as N7, "
                        sql &= "ISNULL(SUM(CAST(FS_08 AS float)),0) as N8, "
                        sql &= "ISNULL(SUM(CAST(FS_09 AS float)),0) as N9, "
                        sql &= "ISNULL(SUM(CAST(FS_10 AS float)),0) as N10, "
                        sql &= "ISNULL(SUM(CAST(FS_11 AS float)),0) as N11, "
                        sql &= "ISNULL(SUM(CAST(FS_12 AS float)),0) as N12, "
                        sql &= "ISNULL(SUM(CAST(FS_13 AS float)),0) as Total "
                        sql &= "FROM LF_BuyerLocalStockPlan "
                        sql &= " Where BuyerGroup = '" & xBuyerGroup & "' "
                        sql &= " And Version = 99 "
                        sql &= " And GR_08 = '" & "CH" & "' "
                        sql &= " And LOGID = '" & xLogID & "' "
                        Dim dt_FCTQtyG As DataTable = uDataBase.GetDataTable(sql)
                        If dt_FCTQtyG.Rows.Count > 0 Then

                            Dim HPtc1 As TableCell = New TableCell
                            HPtc1.Text = Format(dt_FCTQtyG.Rows(0).Item("N"), "###,###,##0.0")
                            HProw.Cells.Add(HPtc1)

                            Dim HPtc2 As TableCell = New TableCell
                            HPtc2.Text = Format(dt_FCTQtyG.Rows(0).Item("N1"), "###,###,##0.0")
                            HProw.Cells.Add(HPtc2)

                            Dim HPtc3 As TableCell = New TableCell
                            HPtc3.Text = Format(dt_FCTQtyG.Rows(0).Item("N2"), "###,###,##0.0")
                            HProw.Cells.Add(HPtc3)

                            Dim HPtc4 As TableCell = New TableCell
                            HPtc4.Text = Format(dt_FCTQtyG.Rows(0).Item("N3"), "###,###,##0.0")
                            HProw.Cells.Add(HPtc4)

                            Dim HPtc5 As TableCell = New TableCell
                            HPtc5.Text = Format(dt_FCTQtyG.Rows(0).Item("N4"), "###,###,##0.0")
                            HProw.Cells.Add(HPtc5)

                            Dim HPtc6 As TableCell = New TableCell
                            HPtc6.Text = Format(dt_FCTQtyG.Rows(0).Item("N5"), "###,###,##0.0")
                            HProw.Cells.Add(HPtc6)

                            Dim HPtc7 As TableCell = New TableCell
                            HPtc7.Text = Format(dt_FCTQtyG.Rows(0).Item("N6"), "###,###,##0.0")
                            HProw.Cells.Add(HPtc7)

                            Dim HPtc8 As TableCell = New TableCell
                            HPtc8.Text = Format(dt_FCTQtyG.Rows(0).Item("N7"), "###,###,##0.0")
                            HProw.Cells.Add(HPtc8)

                            Dim HPtc9 As TableCell = New TableCell
                            HPtc9.Text = Format(dt_FCTQtyG.Rows(0).Item("N8"), "###,###,##0.0")
                            HProw.Cells.Add(HPtc9)

                            Dim HPtcA As TableCell = New TableCell
                            HPtcA.Text = Format(dt_FCTQtyG.Rows(0).Item("N9"), "###,###,##0.0")
                            HProw.Cells.Add(HPtcA)

                            Dim HPtcB As TableCell = New TableCell
                            HPtcB.Text = Format(dt_FCTQtyG.Rows(0).Item("N10"), "###,###,##0.0")
                            HProw.Cells.Add(HPtcB)

                            Dim HPtcC As TableCell = New TableCell
                            HPtcC.Text = Format(dt_FCTQtyG.Rows(0).Item("N11"), "###,###,##0.0")
                            HProw.Cells.Add(HPtcC)

                            Dim HPtcD As TableCell = New TableCell
                            HPtcD.Text = Format(dt_FCTQtyG.Rows(0).Item("N12"), "###,###,##0.0")
                            HProw.Cells.Add(HPtcD)

                            Dim HPtcE As TableCell = New TableCell
                            HPtcE.Text = Format(dt_FCTQtyG.Rows(0).Item("Total"), "###,###,##0.0")
                            HProw.Cells.Add(HPtcE)
                        End If
                        e.Row.Parent.Controls.Add(HProw)
                        '-- ----------------------------------------------------------------
                        '-- LS CH QTY
                        '-- ----------------------------------------------------------------
                        Dim HQHP1 As TableCell = New TableCell
                        HQHP1.Text = "CH-BYLSPLAN"
                        HQrow.Cells.Add(HQHP1)

                        sql = "SELECT "
                        sql &= "ISNULL(SUM(CAST(ProdQty AS float)),0) as Total "
                        sql &= "FROM B_LS2EDIInterface "
                        sql &= " Where BuyerGroup = '" & xBuyerGroup & "' "
                        sql &= " And PartType = '" & "CH" & "' "
                        Dim dt_FCTQtyH As DataTable = uDataBase.GetDataTable(sql)
                        If dt_FCTQtyH.Rows.Count > 0 Then

                            Dim HQtc1 As TableCell = New TableCell
                            HQtc1.Text = ""
                            HQrow.Cells.Add(HQtc1)

                            Dim HQtc2 As TableCell = New TableCell
                            HQtc2.Text = ""
                            HQrow.Cells.Add(HQtc2)

                            Dim HQtc3 As TableCell = New TableCell
                            HQtc3.Text = ""
                            HQrow.Cells.Add(HQtc3)

                            Dim HQtc4 As TableCell = New TableCell
                            HQtc4.Text = ""
                            HQrow.Cells.Add(HQtc4)

                            Dim HQtc5 As TableCell = New TableCell
                            HQtc5.Text = ""
                            HQrow.Cells.Add(HQtc5)

                            Dim HQtc6 As TableCell = New TableCell
                            HQtc6.Text = ""
                            HQrow.Cells.Add(HQtc6)

                            Dim HQtc7 As TableCell = New TableCell
                            HQtc7.Text = ""
                            HQrow.Cells.Add(HQtc7)

                            Dim HQtc8 As TableCell = New TableCell
                            HQtc8.Text = ""
                            HQrow.Cells.Add(HQtc8)

                            Dim HQtc9 As TableCell = New TableCell
                            HQtc9.Text = ""
                            HQrow.Cells.Add(HQtc9)

                            Dim HQtcA As TableCell = New TableCell
                            HQtcA.Text = ""
                            HQrow.Cells.Add(HQtcA)

                            Dim HQtcB As TableCell = New TableCell
                            HQtcB.Text = ""
                            HQrow.Cells.Add(HQtcB)

                            Dim HQtcC As TableCell = New TableCell
                            HQtcC.Text = ""
                            HQrow.Cells.Add(HQtcC)

                            Dim HQtcD As TableCell = New TableCell
                            HQtcD.Text = ""
                            HQrow.Cells.Add(HQtcD)

                            Dim HQtcE As TableCell = New TableCell
                            HQtcE.Text = Format(dt_FCTQtyH.Rows(0).Item("Total"), "###,###,##0.0")
                            HQrow.Cells.Add(HQtcE)
                        End If
                        e.Row.Parent.Controls.Add(HQrow)
                        '-- ----------------------------------------------------------------
                        '-- 差異
                        '-- 
                        '-- ----------------------------------------------------------------
                        Dim HRHP1 As TableCell = New TableCell
                        HRHP1.BackColor = Color.LightPink
                        HRHP1.Text = "[EDI-CH]-- 差"
                        HRrow.Cells.Add(HRHP1)

                        Dim HRtc1 As TableCell = New TableCell
                        HRtc1.HorizontalAlign = HorizontalAlign.Right
                        HRtc1.BackColor = Color.LightPink
                        HRtc1.Text = ""
                        HRrow.Cells.Add(HRtc1)

                        Dim HRtc2 As TableCell = New TableCell
                        HRtc2.HorizontalAlign = HorizontalAlign.Right
                        HRtc2.BackColor = Color.LightPink
                        HRtc2.Text = ""
                        HRrow.Cells.Add(HRtc2)

                        Dim HRtc3 As TableCell = New TableCell
                        HRtc3.HorizontalAlign = HorizontalAlign.Right
                        HRtc3.BackColor = Color.LightPink
                        HRtc3.Text = ""
                        HRrow.Cells.Add(HRtc3)

                        Dim HRtc4 As TableCell = New TableCell
                        HRtc4.HorizontalAlign = HorizontalAlign.Right
                        HRtc4.BackColor = Color.LightPink
                        HRtc4.Text = ""
                        HRrow.Cells.Add(HRtc4)

                        Dim HRtc5 As TableCell = New TableCell
                        HRtc5.HorizontalAlign = HorizontalAlign.Right
                        HRtc5.BackColor = Color.LightPink
                        HRtc5.Text = ""
                        HRrow.Cells.Add(HRtc5)

                        Dim HRtc6 As TableCell = New TableCell
                        HRtc6.HorizontalAlign = HorizontalAlign.Right
                        HRtc6.BackColor = Color.LightPink
                        HRtc6.Text = ""
                        HRrow.Cells.Add(HRtc6)

                        Dim HRtc7 As TableCell = New TableCell
                        HRtc7.HorizontalAlign = HorizontalAlign.Right
                        HRtc7.BackColor = Color.LightPink
                        HRtc7.Text = ""
                        HRrow.Cells.Add(HRtc7)

                        Dim HRtc8 As TableCell = New TableCell
                        HRtc8.HorizontalAlign = HorizontalAlign.Right
                        HRtc8.BackColor = Color.LightPink
                        HRtc8.Text = ""
                        HRrow.Cells.Add(HRtc8)

                        Dim HRtc9 As TableCell = New TableCell
                        HRtc9.HorizontalAlign = HorizontalAlign.Right
                        HRtc9.BackColor = Color.LightPink
                        HRtc9.Text = ""
                        HRrow.Cells.Add(HRtc9)

                        Dim HRtcA As TableCell = New TableCell
                        HRtcA.HorizontalAlign = HorizontalAlign.Right
                        HRtcA.BackColor = Color.LightPink
                        HRtcA.Text = ""
                        HRrow.Cells.Add(HRtcA)

                        Dim HRtcB As TableCell = New TableCell
                        HRtcB.HorizontalAlign = HorizontalAlign.Right
                        HRtcB.BackColor = Color.LightPink
                        HRtcB.Text = ""
                        HRrow.Cells.Add(HRtcB)

                        Dim HRtcC As TableCell = New TableCell
                        HRtcC.HorizontalAlign = HorizontalAlign.Right
                        HRtcC.BackColor = Color.LightPink
                        HRtcC.Text = ""
                        HRrow.Cells.Add(HRtcC)

                        Dim HRtcD As TableCell = New TableCell
                        HRtcD.HorizontalAlign = HorizontalAlign.Right
                        HRtcD.BackColor = Color.LightPink
                        HRtcD.Text = ""
                        HRrow.Cells.Add(HRtcD)

                        Dim HRtcE As TableCell = New TableCell
                        HRtcE.HorizontalAlign = HorizontalAlign.Right
                        HRtcE.BackColor = Color.LightPink
                        HRtcE.Text = Format(dt_FCTQtyG.Rows(0).Item("Total") - dt_FCTQtyH.Rows(0).Item("Total"), "###,###,##0.0")
                        HRrow.Cells.Add(HRtcE)
                        e.Row.Parent.Controls.Add(HRrow)

                        If Format(dt_FCTQtyG.Rows(0).Item("Total"), "###,###,##0.0") <> Format(dt_FCTQtyH.Rows(0).Item("Total"), "###,###,##0.0") Then
                            If DMessage.Text = "" Then DMessage.Text = "**[注意]發現數量不符，請再次確認客戶需求量QTY.  /EDI-CH"
                        End If
                    End If
                    'END
                End If
            End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then

        End If
    End Sub

End Class
