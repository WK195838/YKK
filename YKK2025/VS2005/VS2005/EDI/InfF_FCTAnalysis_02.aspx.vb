Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfF_FCTAnalysis_02
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim gBuyer As String            'Buyer
    Dim gCustCode As String         'Customer Code
    Dim gSeason As String           'Season
    Dim gMonth As String            'Month
    Dim gItem As String             'ITEM
    Dim gNewFCT As String           '含NewFCT
    Dim Version(6) As String        '最新6版本
    Dim Total(6) As Integer         '最新6版本合計數量

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
            ShowFCTDataList()                      '顯示資料
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
        Response.Cookies("PGM").Value = "InfF_FCTAnalysis_02.aspx"  '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")  '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                 '現在日期時間
        gBuyer = Request.QueryString("pBuyer")             'Buyer
        gCustCode = Request.QueryString("pCustCode")       'Customer Code
        gSeason = Request.QueryString("pSeason")           'Season
        gMonth = Request.QueryString("pMonth")             'Month
        gItem = Request.QueryString("pItem")               'Item
        gNewFCT = Request.QueryString("pNewFCT")           '含NewFCT
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
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
    '**(ShowFCTDataList)
    '**     顯示資料
    '**
    '*****************************************************************
    Sub ShowFCTDataList()
        Dim sql As String
        ' 篩選資料
        sql = "SELECT "
        sql &= "a.CustCode As CustCode, b.CustName + '(' + a.CustCode + ')' As Customer, a.Buyer, a.Season, a.Month, a.CustItem, "
        sql &= "a.Color, a.Style, a.Article, a.Part, "
        sql &= "'' As Qty6, '' As Ratio6, "
        sql &= "'' As Qty5, '' As Ratio5, "
        sql &= "'' As Qty4, '' As Ratio4, "
        sql &= "'' As Qty3, '' As Ratio3, "
        sql &= "'' As Qty2, '' As Ratio2, "
        sql &= "'' As Qty1, '' As Ratio1 "
        sql &= "From A_CustomerActual_FCT a, M_NativeVendor b "

        ' T&P NIKE=000013T
        If gBuyer = "000013T" Then
            sql &= "Where 'FALL-TP000013' = b.Buyer "
        Else
            sql &= "Where 'FALL-' + a.Buyer = b.Buyer "
        End If

        sql &= "  And a.CustCode = b.CustCode "
        sql &= "  And a.CustCode  = '" & gCustCode & "' "
        sql &= "  And a.Buyer  = '" & gBuyer & "' "
        sql &= "  And a.Season = '" & gSeason & "' "
        sql &= "  And a.Month  = '" & gMonth & "' "
        sql &= "  And a.CustItem  = '" & gItem & "' "
        ' 是否含 New-FCT
        If gNewFCT = "0" Then
            sql &= "  And a.Version <> '" & "99999999999999" & "' "
        End If
        sql &= "Group by a.CustCode, b.CustName, a.Buyer, a.Season, a.Month, a.CustItem, a.Color, a.Style, a.Article, a.Part "
        sql &= "Order by a.CustCode, b.CustName, a.Buyer, a.Season, a.Month, a.CustItem, a.Color, a.Style, a.Article, a.Part "
        '
        Dim dt_FCTPlan_ITEM As DataTable = uDataBase.GetDataTable(sql)
        If dt_FCTPlan_ITEM.Rows.Count > 0 Then
            ColorGridView.Visible = True
            ColorGridView.DataSource = dt_FCTPlan_ITEM
            ColorGridView.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub ColorGridView_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles ColorGridView.RowDataBound
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            Dim sql As String
            Dim BefQty As Double = 0
            Dim j As Integer = 0
            ' 版本明細展開

            For i As Integer = 1 To 6
                If Version(i) <> "" Then
                    sql = "SELECT Isnull(Sum(FCTQty),0) As Qty From A_CustomerActual_FCT "
                    sql &= "Where Buyer  = '" & gBuyer & "' "
                    sql &= "  And Season = '" & gSeason & "' "
                    sql &= "  And Month  = '" & gMonth & "' "
                    sql &= "  And CustCode = '" & gCustCode & "' "
                    sql &= "  And CustItem = '" & gItem & "' "
                    sql &= "  And Color = '" & DataBinder.Eval(e.Row.DataItem, "Color") & "' "
                    sql &= "  And Style = '" & DataBinder.Eval(e.Row.DataItem, "Style") & "' "
                    sql &= "  And Article = '" & DataBinder.Eval(e.Row.DataItem, "Article") & "' "
                    sql &= "  And Part = '" & DataBinder.Eval(e.Row.DataItem, "Part") & "' "
                    sql &= "  And Version = '" & Version(i) & "' "
                    '
                    Dim dt_FCTQty As DataTable = uDataBase.GetDataTable(sql)
                    If dt_FCTQty.Rows.Count > 0 Then
                        e.Row.Cells(i * 2 + 5).Text = Format(dt_FCTQty.Rows(0).Item("Qty"), "###,###,###")
                        '
                        If i > 1 Then
                            If dt_FCTQty.Rows(0).Item("Qty") > 0 Then
                                e.Row.Cells(i * 2 + 5 - 1).Text = Format(BefQty / dt_FCTQty.Rows(0).Item("Qty") * 100, ".0") + "%"
                            Else
                                e.Row.Cells(i * 2 + 5 - 1).Text = "-"
                            End If
                        End If
                        '
                        BefQty = dt_FCTQty.Rows(0).Item("Qty")
                        Total(i) = Total(i) + dt_FCTQty.Rows(0).Item("Qty")
                    End If
                End If
            Next
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
            ' 合計展開
            Dim row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)

            Dim tc1 As TableCell = New TableCell
            tc1.BackColor = Color.YellowGreen
            tc1.ColumnSpan = 7
            tc1.Text = "合計"
            row.Cells.Add(tc1)

            Dim tc2 As TableCell = New TableCell
            tc2.HorizontalAlign = HorizontalAlign.Right
            tc2.BackColor = Color.YellowGreen
            tc2.Text = Format(Total(1), "###,###,###")
            row.Cells.Add(tc2)

            Dim tc3 As TableCell = New TableCell
            tc3.HorizontalAlign = HorizontalAlign.Right
            tc3.BackColor = Color.YellowGreen
            If Total(2) > 0 Then
                tc3.Text = Format(Total(1) / Total(2) * 100, ".0") + "%"
            Else
                tc3.Text = ""
            End If
            row.Cells.Add(tc3)

            Dim tc4 As TableCell = New TableCell
            tc4.HorizontalAlign = HorizontalAlign.Right
            tc4.BackColor = Color.YellowGreen
            If Total(2) > 0 Then
                tc4.Text = Format(Total(2), "###,###,###")
            Else
                tc4.Text = ""
            End If
            row.Cells.Add(tc4)

            Dim tc5 As TableCell = New TableCell
            tc5.HorizontalAlign = HorizontalAlign.Right
            tc5.BackColor = Color.YellowGreen
            If Total(3) > 0 Then
                tc5.Text = Format(Total(2) / Total(3) * 100, ".0") + "%"
            Else
                tc5.Text = ""
            End If
            row.Cells.Add(tc5)

            Dim tc6 As TableCell = New TableCell
            tc6.HorizontalAlign = HorizontalAlign.Right
            tc6.BackColor = Color.YellowGreen
            If Total(3) > 0 Then
                tc6.Text = Format(Total(3), "###,###,###")
            Else
                tc6.Text = ""
            End If
            row.Cells.Add(tc6)

            Dim tc7 As TableCell = New TableCell
            tc7.HorizontalAlign = HorizontalAlign.Right
            tc7.BackColor = Color.YellowGreen
            If Total(4) > 0 Then
                tc7.Text = Format(Total(3) / Total(4) * 100, ".0") + "%"
            Else
                tc7.Text = ""
            End If
            row.Cells.Add(tc7)

            Dim tc8 As TableCell = New TableCell
            tc8.HorizontalAlign = HorizontalAlign.Right
            tc8.BackColor = Color.YellowGreen
            If Total(4) > 0 Then
                tc8.Text = Format(Total(4), "###,###,###")
            Else
                tc8.Text = ""
            End If
            row.Cells.Add(tc8)

            Dim tc9 As TableCell = New TableCell
            tc9.HorizontalAlign = HorizontalAlign.Right
            tc9.BackColor = Color.YellowGreen
            If Total(5) > 0 Then
                tc9.Text = Format(Total(4) / Total(5) * 100, ".0") + "%"
            Else
                tc9.Text = ""
            End If
            row.Cells.Add(tc9)

            Dim tcA As TableCell = New TableCell
            tcA.HorizontalAlign = HorizontalAlign.Right
            tcA.BackColor = Color.YellowGreen
            If Total(5) > 0 Then
                tcA.Text = Format(Total(5), "###,###,###")
            Else
                tcA.Text = ""
            End If
            row.Cells.Add(tcA)

            Dim tcB As TableCell = New TableCell
            tcB.HorizontalAlign = HorizontalAlign.Right
            tcB.BackColor = Color.YellowGreen
            If Total(6) > 0 Then
                tcB.Text = Format(Total(5) / Total(6) * 100, ".0") + "%"
            Else
                tcB.Text = ""
            End If
            row.Cells.Add(tcB)

            Dim tcC As TableCell = New TableCell
            tcC.HorizontalAlign = HorizontalAlign.Right
            tcC.BackColor = Color.YellowGreen
            If Total(6) > 0 Then
                tcC.Text = Format(Total(6), "###,###,###")
            Else
                tcC.Text = ""
            End If
            row.Cells.Add(tcC)

            Dim tcD As TableCell = New TableCell
            tcD.HorizontalAlign = HorizontalAlign.Right
            tcD.BackColor = Color.YellowGreen
            tcD.Text = ""
            row.Cells.Add(tcD)

            e.Row.Parent.Controls.Add(row)
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯表頭
    '**
    '*****************************************************************
    Protected Sub ColorGridView_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles ColorGridView.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim gv As GridView = sender
            Dim H1row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H2row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()

            ' 表頭-3行
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "F"
            H3row.Cells.Add(H3tc1)

            Dim H3tc2 As TableCell = New TableCell
            H3tc2.Text = "R"
            H3row.Cells.Add(H3tc2)

            Dim H3tc3 As TableCell = New TableCell
            H3tc3.Text = "F"
            H3row.Cells.Add(H3tc3)

            Dim H3tc4 As TableCell = New TableCell
            H3tc4.Text = "R"
            H3row.Cells.Add(H3tc4)

            Dim H3tc5 As TableCell = New TableCell
            H3tc5.Text = "F"
            H3row.Cells.Add(H3tc5)

            Dim H3tc6 As TableCell = New TableCell
            H3tc6.Text = "R"
            H3row.Cells.Add(H3tc6)

            Dim H3tc7 As TableCell = New TableCell
            H3tc7.Text = "F"
            H3row.Cells.Add(H3tc7)

            Dim H3tc8 As TableCell = New TableCell
            H3tc8.Text = "R"
            H3row.Cells.Add(H3tc8)

            Dim H3tc9 As TableCell = New TableCell
            H3tc9.Text = "F"
            H3row.Cells.Add(H3tc9)

            Dim H3tcA As TableCell = New TableCell
            H3tcA.Text = "R"
            H3row.Cells.Add(H3tcA)

            Dim H3tcB As TableCell = New TableCell
            H3tcB.Text = "F"
            H3row.Cells.Add(H3tcB)

            Dim H3tcC As TableCell = New TableCell
            H3tcC.Text = "R"
            H3row.Cells.Add(H3tcC)

            gv.Controls(0).Controls.AddAt(0, H3row)
            ' 表頭-2行
            Dim H2tc1 As TableCell = New TableCell
            H2tc1.Text = ""
            H2tc1.ColumnSpan = 2
            H2row.Cells.Add(H2tc1)

            Dim H2tc2 As TableCell = New TableCell
            H2tc2.Text = ""
            H2tc2.ColumnSpan = 2
            H2row.Cells.Add(H2tc2)

            Dim H2tc3 As TableCell = New TableCell
            H2tc3.Text = ""
            H2tc3.ColumnSpan = 2
            H2row.Cells.Add(H2tc3)

            Dim H2tc4 As TableCell = New TableCell
            H2tc4.Text = ""
            H2tc4.ColumnSpan = 2
            H2row.Cells.Add(H2tc4)

            Dim H2tc5 As TableCell = New TableCell
            H2tc5.Text = ""
            H2tc5.ColumnSpan = 2
            H2row.Cells.Add(H2tc5)

            Dim H2tc6 As TableCell = New TableCell
            H2tc6.Text = ""
            H2tc6.ColumnSpan = 2
            H2row.Cells.Add(H2tc6)

            ' Set 6 個最新 Version 
            Dim sql As String
            sql = "SELECT Top 6 Version From A_CustomerActual "
            sql &= "Where Buyer  = '" & gBuyer & "' "
            sql &= "  And Season = '" & gSeason & "' "
            sql &= "  And Month  = '" & gMonth & "' "
            sql &= "  And FCTQty > 0 "
            ' 是否含 New-FCT
            If gNewFCT = "0" Then
                sql &= "  And Version <> '" & "99999999999999" & "' "
            End If
            sql &= "Group by Version "
            sql &= "Order by Version Desc "
            '
            Dim dt_Version As DataTable = uDataBase.GetDataTable(sql)
            For i As Integer = 1 To dt_Version.Rows.Count
                If i = 1 Then H2tc1.Text = dt_Version.Rows(i - 1).Item("Version")
                If i = 2 Then H2tc2.Text = dt_Version.Rows(i - 1).Item("Version")
                If i = 3 Then H2tc3.Text = dt_Version.Rows(i - 1).Item("Version")
                If i = 4 Then H2tc4.Text = dt_Version.Rows(i - 1).Item("Version")
                If i = 5 Then H2tc5.Text = dt_Version.Rows(i - 1).Item("Version")
                If i = 6 Then H2tc6.Text = dt_Version.Rows(i - 1).Item("Version")
                Version(i) = dt_Version.Rows(i - 1).Item("Version")
            Next

            gv.Controls(0).Controls.AddAt(0, H2row)
            '-----------------------------------------
            ' 表頭-1行
            Dim H1tc1 As TableCell = New TableCell
            H1tc1.Text = "成衣廠"
            H1tc1.RowSpan = 3
            H1row.Cells.Add(H1tc1)

            Dim H1tc1A As TableCell = New TableCell
            H1tc1A.Text = "季"
            H1tc1A.RowSpan = 3
            H1row.Cells.Add(H1tc1A)

            Dim H1tc2 As TableCell = New TableCell
            H1tc2.Text = "ITEM"
            H1tc2.RowSpan = 3
            H1row.Cells.Add(H1tc2)

            Dim H1tc3 As TableCell = New TableCell
            H1tc3.Text = "COLOR"
            H1tc3.RowSpan = 3
            H1row.Cells.Add(H1tc3)

            Dim H1tc4 As TableCell = New TableCell
            H1tc4.Text = "STYLE"
            H1tc4.RowSpan = 3
            H1row.Cells.Add(H1tc4)

            Dim H1tc5 As TableCell = New TableCell
            H1tc5.Text = "INF-1"
            H1tc5.RowSpan = 3
            H1row.Cells.Add(H1tc5)

            Dim H1tc6 As TableCell = New TableCell
            H1tc6.Text = "INF-2"
            H1tc6.RowSpan = 3
            H1row.Cells.Add(H1tc6)

            Dim H1tc7 As TableCell = New TableCell
            H1tc7.Text = Mid(gMonth, 1, 4) + "/" + Mid(gMonth, 5)
            H1tc7.ColumnSpan = 12
            H1row.Cells.Add(H1tc7)

            gv.Controls(0).Controls.AddAt(0, H1row)
            '-----------------------------------------
        End If
    End Sub


End Class
