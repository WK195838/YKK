Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfF_FCTACTAnalysis
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate, NowYY, NowMM As String           '現在日期時間
    Dim Buyer As String             'Buyer
    Dim UserID As String            'UserID
    Dim Month(6) As String          '最新6版本(只上線4個月)
    Dim Version(6) As String        '最後6版本(只上線4個月)
    Dim MMSQLStr As String          '每月最後版本SQL
    Dim ACTTotal(6), FCTTotal(6) As Integer         '最新6版本合計數量
    Dim ACTBTotal, ACTATotal As Integer             'ACT BEF, ACT 合計數量

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
        Response.Cookies("PGM").Value = "InfF_FCTACTAnalysis.aspx"     '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        NowYY = Mid(NowDate, 1, 4)
        NowMM = Mid(NowDate, 5, 2)
        'NowMM = "06"   '測試使用
        Buyer = Request.QueryString("pBuyer")
        If Buyer <> "000013T" And Buyer <> "TW0371T" Then
            Buyer = Mid(Request.QueryString("pBuyer"), 1, 6)
        End If

        UserID = Request.QueryString("pUserID")             'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        DBuyer.ReadOnly = True
        CSTGridView.Visible = False
        For i As Integer = 1 To 6                           '最新6版本
            Month(i) = ""
            Version(i) = ""
            ACTTotal(i) = 0
            FCTTotal(i) = 0
        Next
        ACTBTotal = 0
        ACTATotal = 0
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
        Dim sql As String
        '
        DBuyer.Text = Buyer
        '
        DVersion.Items.Clear()
        DVersion.Items.Add("NIL")
        ' Delete 2015/4/29 JOY
        ' 只限設定 [NIL]
        'sql = "SELECT Top 6 Version From A_CustomerActual "
        'sql &= "Where Buyer  = '" & DBuyer.Text & "' "
        'sql &= "  And Version <> '" & "ZZZZZZZZZZZZZZ" & "' "
        'sql &= "Group by Version "
        'sql &= "Order by Version Desc "
        'Dim dt_Version As DataTable = uDataBase.GetDataTable(sql)
        'If dt_Version.Rows.Count > 0 Then
        '    DVersion.Enabled = True
        '    BInq.Enabled = True
        '    For i As Integer = 0 To dt_Version.Rows.Count - 1
        '        Dim ListItem0 As New ListItem
        '        ListItem0.Text = dt_Version.Rows(i).Item("Version")
        '        ListItem0.Value = dt_Version.Rows(i).Item("Version")
        '        DVersion.Items.Add(ListItem0)
        '    Next
        'Else
        '    DVersion.Enabled = False
        '    BInq.Enabled = False
        'End If
        '
        DSeason.Items.Clear()
        sql = "SELECT Top 18 Season From A_CustomerActual "
        sql &= "Where Buyer  = '" & DBuyer.Text & "' "
        'sql &= "  And Month  = '" & NowYY + NowMM & "' "
        sql &= "Group by Season "
        sql &= "Order by Substring(Season,3,2) + Substring(Season,1,2) Desc "
        Dim dt_Season As DataTable = uDataBase.GetDataTable(sql)
        If dt_Season.Rows.Count > 0 Then
            DSeason.Enabled = True
            BInq.Enabled = True
            For i As Integer = 0 To dt_Season.Rows.Count - 1
                Dim ListItem0 As New ListItem
                ListItem0.Text = dt_Season.Rows(i).Item("Season")
                ListItem0.Value = dt_Season.Rows(i).Item("Season")
                DSeason.Items.Add(ListItem0)
            Next
        Else
            DSeason.Enabled = False
            BInq.Enabled = False
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Inq)
    '**     查詢
    '**
    '*****************************************************************
    Protected Sub BInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BInq.Click
        Dim sql, URLStr As String
        ' Set 6 個最新 Month (只上線4個月)
        For i As Integer = 0 To 3
            If CInt(NowMM) - i > 0 Then
                If CInt(NowMM) - i < 10 Then
                    Month(i + 1) = NowYY + "0" + CStr(CInt(NowMM) - i)
                Else
                    Month(i + 1) = NowYY + CStr(CInt(NowMM) - i)
                End If
            Else
                If CInt(NowMM) + 12 - i < 10 Then
                    Month(i + 1) = CStr(CInt(NowYY) - 1) + "0" + CStr(CInt(NowMM) + 12 - i)
                Else
                    Month(i + 1) = CStr(CInt(NowYY) - 1) + CStr(CInt(NowMM) + 12 - i)
                End If
            End If
        Next
        ' Set 6 個最新 Version  (只上線4個月)
        If DVersion.Text <> "NIL" Then
            ' 指定特定Version
            For i As Integer = 1 To 4
                Version(i) = DVersion.Text
            Next
        Else
            ' 指定Season
            For i As Integer = 1 To 4
                sql = "SELECT Top 1 Version From A_CustomerActual "
                sql &= "Where Buyer  = '" & DBuyer.Text & "' "
                sql &= "  And Season = '" & DSeason.Text & "' "
                sql &= "  And Month  = '" & Month(i) & "' "
                sql &= "  And FCTQty > 0 "
                sql &= "  And Version <> '" & "99999999999999" & "' "
                sql &= "Order by Version Desc "
                Dim dt_Version As DataTable = uDataBase.GetDataTable(sql)
                If dt_Version.Rows.Count > 0 Then
                    Version(i) = dt_Version.Rows(0).Item("Version")
                Else
                    sql = "SELECT Top 1 Version From A_CustomerActual "
                    sql &= "Where Buyer  = '" & DBuyer.Text & "' "
                    sql &= "  And Season = '" & DSeason.Text & "' "
                    sql &= "  And Month  = '" & Month(i) & "' "
                    sql &= "  And Version <> '" & "99999999999999" & "' "
                    sql &= "Order by Version Desc "
                    Dim dt_Version1 As DataTable = uDataBase.GetDataTable(sql)
                    If dt_Version1.Rows.Count > 0 Then
                        Version(i) = dt_Version1.Rows(0).Item("Version")
                    End If
                End If
            Next
        End If
        '
        'URL
        If DLevel.Text = "ITEM" Then
            URLStr = "'InfF_FCTACTAnalysis_01.aspx?' + "
        Else
            URLStr = "'InfF_FCTACTAnalysis_01a.aspx?' + "
        End If
        URLStr = URLStr & "'pBuyer=' + a.Buyer + "
        URLStr = URLStr & "'&pCustCode=' + a.CustCode + "
        URLStr = URLStr & "'&pSeason=' + a.Season + "
        URLStr = URLStr & "'&pMonth=' + '" + NowYY + NowMM + "' + "
        URLStr = URLStr & "'&pVersion=' + '" + DVersion.Text + "'"
        '
        ' 每月最後版本SQL
        For i As Integer = 1 To 4
            If Month(i) <> "" Then
                If i = 1 Then
                    MMSQLStr = "(" + "Month='" + Month(i) + "' And Version='" + Version(i) + "')"
                Else
                    MMSQLStr &= " Or (" + "Month='" + Month(i) + "' And Version='" + Version(i) + "')"
                End If
            End If
        Next
        '
        ' 顯示資料
        sql = "SELECT "
        sql &= "a.CustCode As CustCode, b.CustName + '(' + a.CustCode + ')' As Customer, a.Buyer, a.Season, "
        sql &= "'' As ABQty6, '' As AAQty6, '' As AQty6, '' As FQty6, '' As Ratio6, "
        sql &= "'' As AQty5, '' As FQty5, '' As Ratio5, "
        sql &= "'' As AQty4, '' As FQty4, '' As Ratio4, "
        sql &= "'' As AQty3, '' As FQty3, '' As Ratio3, "
        sql &= "'' As AQty2, '' As FQty2, '' As Ratio2, "
        sql &= "'' As AQty1, '' As FQty1, '' As Ratio1, "
        sql &= URLStr & " As URL "
        sql &= "From A_CustomerActual a, M_NativeVendor b "

        ' T&P NIKE=000013T
        If DBuyer.Text = "000013T" Then
            sql &= "Where 'FALL-TP000013' = b.Buyer "
        Else
            sql &= "Where 'FALL-' + a.Buyer = b.Buyer "
        End If

        sql &= "  And a.CustCode = b.CustCode "
        sql &= "  And a.Buyer  = '" & DBuyer.Text & "' "
        sql &= "  And a.Season = '" & DSeason.Text & "' "
        sql &= "  And ( "
        sql &= MMSQLStr
        sql &= "      ) "
        sql &= "Group by a.CustCode, b.CustName, a.Buyer, a.Season "
        sql &= "Order by a.CustCode, b.CustName, a.Buyer, a.Season "
        '
        Dim dt_FCTPlan As DataTable = uDataBase.GetDataTable(sql)
        If dt_FCTPlan.Rows.Count > 0 Then
            CSTGridView.Visible = True
            CSTGridView.DataSource = dt_FCTPlan
            CSTGridView.DataBind()
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
    Protected Sub CSTGridView_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles CSTGridView.RowDataBound
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            Dim sql As String
            Dim j As Integer = 0
            '
            ' 版本明細展開(只上線4個月)
            For i As Integer = 1 To 4
                If Month(i) <> "" Then
                    ' ACT / FCT 比較
                    sql = "SELECT Top 1 Isnull(Sum(FCTQty),0) As FCTQty, Isnull(Sum(ACTQty),0) As ACTQty From A_CustomerActual "
                    sql &= "Where Buyer  = '" & DBuyer.Text & "' "
                    sql &= "  And Season = '" & DSeason.Text & "' "
                    sql &= "  And CustCode = '" & DataBinder.Eval(e.Row.DataItem, "CustCode") & "' "
                    sql &= "  And Month  = '" & Month(i) & "' "
                    sql &= "  And Version = '" & Version(i) & "' "
                    sql &= "Group by  Buyer, CustCode "
                    sql &= "Order by  Buyer, CustCode "
                    '
                    Dim dt_FCTQty As DataTable = uDataBase.GetDataTable(sql)
                    If dt_FCTQty.Rows.Count > 0 Then
                        e.Row.Cells(i * 3 - 1).Text = Format(dt_FCTQty.Rows(0).Item("ACTQty"), "###,###,###")
                        e.Row.Cells(i * 3).Text = Format(dt_FCTQty.Rows(0).Item("FCTQty"), "###,###,###")
                        j = i * 3
                        '
                        If dt_FCTQty.Rows(0).Item("FCTQty") > 0 Then
                            e.Row.Cells(j + 1).Text = Format(dt_FCTQty.Rows(0).Item("ACTQty") / dt_FCTQty.Rows(0).Item("FCTQty") * 100, ".0") + "%"
                        Else
                            e.Row.Cells(j + 1).Text = ""
                        End If
                        '
                        ACTTotal(i) = ACTTotal(i) + dt_FCTQty.Rows(0).Item("ACTQty")
                        FCTTotal(i) = FCTTotal(i) + dt_FCTQty.Rows(0).Item("FCTQty")
                        '----
                        '(STOP)
                        ' 最後1個月的延後及提前 ACT-ORDER 
                        'If dt_FCTQty.Rows(0).Item("ACTQty") > 0 Then
                        '    If i = 1 Then
                        '        ' 延後 ACT-ORDER
                        '        sql = "SELECT Isnull(Sum(OrderQty),0) As Qty From I_Adidas_ActualOrder "
                        '        sql &= "Where Buyer  = '" & DBuyer.Text & "' "
                        '        sql &= "  And Season = '" & DSeason.Text & "' "
                        '        sql &= "  And CustCode = '" & DataBinder.Eval(e.Row.DataItem, "CustCode") & "' "
                        '        sql &= "  And Month  = '" & Month(i) & "' "
                        '        sql &= "  And Convert(int,substring(Month, 5,2)) > V_BuyMonth "
                        '        '
                        '        Dim dt_ACTBQty As DataTable = uDataBase.GetDataTable(sql)
                        '        If dt_ACTBQty.Rows.Count > 0 Then
                        '            e.Row.Cells(i + 1).Text = Format(dt_ACTBQty.Rows(0).Item("Qty"), "###,###,###")
                        '            ACTBTotal = ACTBTotal + dt_ACTBQty.Rows(0).Item("Qty")
                        '        End If
                        '        ' 提前 ACT-ORDER
                        '        sql = "SELECT Isnull(Sum(OrderQty),0) As Qty From I_Adidas_ActualOrder "
                        '        sql &= "Where Buyer  = '" & DBuyer.Text & "' "
                        '        sql &= "  And Season = '" & DSeason.Text & "' "
                        '        sql &= "  And CustCode = '" & DataBinder.Eval(e.Row.DataItem, "CustCode") & "' "
                        '        sql &= "  And Month  = '" & Month(i) & "' "
                        '        sql &= "  And Convert(int,substring(Month, 5,2)) < V_BuyMonth "
                        '        '
                        '        Dim dt_ACTAQty As DataTable = uDataBase.GetDataTable(sql)
                        '        If dt_ACTAQty.Rows.Count > 0 Then
                        '            e.Row.Cells(i + 2).Text = Format(dt_ACTAQty.Rows(0).Item("Qty"), "###,###,###")
                        '            ACTATotal = ACTATotal + dt_ACTAQty.Rows(0).Item("Qty")
                        '        End If
                        '    End If
                        'End If
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
            tc1.ColumnSpan = 2
            tc1.Text = "合計"
            row.Cells.Add(tc1)

            'STOP-延後及提前
            'Dim tc1A As TableCell = New TableCell
            'tc1A.HorizontalAlign = HorizontalAlign.Right
            'tc1A.BackColor = Color.YellowGreen
            'tc1A.Text = Format(ACTBTotal, "###,###,###")
            'row.Cells.Add(tc1A)

            'Dim tc1B As TableCell = New TableCell
            'tc1B.HorizontalAlign = HorizontalAlign.Right
            'tc1B.BackColor = Color.YellowGreen
            'tc1B.Text = Format(ACTATotal, "###,###,###")
            'row.Cells.Add(tc1B)

            Dim tc2 As TableCell = New TableCell
            tc2.HorizontalAlign = HorizontalAlign.Right
            tc2.BackColor = Color.YellowGreen
            tc2.Text = Format(ACTTotal(1), "###,###,###")
            row.Cells.Add(tc2)

            Dim tc2A As TableCell = New TableCell
            tc2A.HorizontalAlign = HorizontalAlign.Right
            tc2A.BackColor = Color.YellowGreen
            tc2A.Text = Format(FCTTotal(1), "###,###,###")
            row.Cells.Add(tc2A)

            Dim tc3 As TableCell = New TableCell
            tc3.HorizontalAlign = HorizontalAlign.Right
            tc3.BackColor = Color.YellowGreen
            If ACTTotal(1) > 0 Then
                tc3.Text = Format(ACTTotal(1) / FCTTotal(1) * 100, ".0") + "%"
            Else
                tc3.Text = ""
            End If
            row.Cells.Add(tc3)

            Dim tc4 As TableCell = New TableCell
            tc4.HorizontalAlign = HorizontalAlign.Right
            tc4.BackColor = Color.YellowGreen
            tc4.Text = Format(ACTTotal(2), "###,###,###")
            row.Cells.Add(tc4)

            Dim tc4A As TableCell = New TableCell
            tc4A.HorizontalAlign = HorizontalAlign.Right
            tc4A.BackColor = Color.YellowGreen
            tc4A.Text = Format(FCTTotal(2), "###,###,###")
            row.Cells.Add(tc4A)

            Dim tc5 As TableCell = New TableCell
            tc5.HorizontalAlign = HorizontalAlign.Right
            tc5.BackColor = Color.YellowGreen
            If ACTTotal(2) > 0 Then
                tc5.Text = Format(ACTTotal(2) / FCTTotal(2) * 100, ".0") + "%"
            Else
                tc5.Text = ""
            End If
            row.Cells.Add(tc5)

            Dim tc6 As TableCell = New TableCell
            tc6.HorizontalAlign = HorizontalAlign.Right
            tc6.BackColor = Color.YellowGreen
            tc6.Text = Format(ACTTotal(3), "###,###,###")
            row.Cells.Add(tc6)

            Dim tc6A As TableCell = New TableCell
            tc6A.HorizontalAlign = HorizontalAlign.Right
            tc6A.BackColor = Color.YellowGreen
            tc6A.Text = Format(FCTTotal(3), "###,###,###")
            row.Cells.Add(tc6A)

            Dim tc7 As TableCell = New TableCell
            tc7.HorizontalAlign = HorizontalAlign.Right
            tc7.BackColor = Color.YellowGreen
            If ACTTotal(3) > 0 Then
                tc7.Text = Format(ACTTotal(3) / FCTTotal(3) * 100, ".0") + "%"
            Else
                tc7.Text = ""
            End If
            row.Cells.Add(tc7)

            Dim tc8 As TableCell = New TableCell
            tc8.HorizontalAlign = HorizontalAlign.Right
            tc8.BackColor = Color.YellowGreen
            tc8.Text = Format(ACTTotal(4), "###,###,###")
            row.Cells.Add(tc8)

            Dim tc8A As TableCell = New TableCell
            tc8A.HorizontalAlign = HorizontalAlign.Right
            tc8A.BackColor = Color.YellowGreen
            tc8A.Text = Format(FCTTotal(4), "###,###,###")
            row.Cells.Add(tc8A)

            Dim tc9 As TableCell = New TableCell
            tc9.HorizontalAlign = HorizontalAlign.Right
            tc9.BackColor = Color.YellowGreen
            If ACTTotal(4) > 0 Then
                tc9.Text = Format(ACTTotal(4) / FCTTotal(4) * 100, ".0") + "%"
            Else
                tc9.Text = ""
            End If
            row.Cells.Add(tc9)

            'Dim tcA As TableCell = New TableCell
            'tcA.HorizontalAlign = HorizontalAlign.Right
            'tcA.BackColor = Color.YellowGreen
            'tcA.Text = Format(ACTTotal(5), "###,###,###")
            'row.Cells.Add(tcA)

            'Dim tcAA As TableCell = New TableCell
            'tcAA.HorizontalAlign = HorizontalAlign.Right
            'tcAA.BackColor = Color.YellowGreen
            'tcAA.Text = Format(FCTTotal(5), "###,###,###")
            'row.Cells.Add(tcAA)

            'Dim tcB As TableCell = New TableCell
            'tcB.HorizontalAlign = HorizontalAlign.Right
            'tcB.BackColor = Color.YellowGreen
            'If ACTTotal(5) > 0 Then
            '    tcB.Text = Format(ACTTotal(5) / FCTTotal(5) * 100, ".0") + "%"
            'Else
            '    tcB.Text = ""
            'End If
            'row.Cells.Add(tcB)

            'Dim tcC As TableCell = New TableCell
            'tcC.HorizontalAlign = HorizontalAlign.Right
            'tcC.BackColor = Color.YellowGreen
            'tcC.Text = Format(ACTTotal(6), "###,###,###")
            'row.Cells.Add(tcC)

            'Dim tcCA As TableCell = New TableCell
            'tcCA.HorizontalAlign = HorizontalAlign.Right
            'tcCA.BackColor = Color.YellowGreen
            'tcCA.Text = Format(FCTTotal(6), "###,###,###")
            'row.Cells.Add(tcCA)

            'Dim tcD As TableCell = New TableCell
            'tcD.HorizontalAlign = HorizontalAlign.Right
            'tcD.BackColor = Color.YellowGreen
            'If ACTTotal(6) > 0 Then
            '    tcD.Text = Format(ACTTotal(6) / FCTTotal(6) * 100, ".0") + "%"
            'Else
            '    tcD.Text = ""
            'End If
            'row.Cells.Add(tcD)

            e.Row.Parent.Controls.Add(row)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯表頭
    '**
    '*****************************************************************
    Protected Sub CSTGridView_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles CSTGridView.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim gv As GridView = sender
            Dim H1row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H2row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-4行
            ' STOP
            'Dim H4tc0 As TableCell = New TableCell
            'H4tc0.Text = "延後"
            'H4row.Cells.Add(H4tc0)

            'Dim H4tc0A As TableCell = New TableCell
            'H4tc0A.Text = "提前"
            'H4row.Cells.Add(H4tc0A)

            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "A"
            H4row.Cells.Add(H4tc1)

            Dim H4tc1A As TableCell = New TableCell
            H4tc1A.Text = "F"
            H4row.Cells.Add(H4tc1A)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = "R"
            H4row.Cells.Add(H4tc2)

            Dim H4tc3 As TableCell = New TableCell
            H4tc3.Text = "A"
            H4row.Cells.Add(H4tc3)

            Dim H4tc3A As TableCell = New TableCell
            H4tc3A.Text = "F"
            H4row.Cells.Add(H4tc3A)

            Dim H4tc4 As TableCell = New TableCell
            H4tc4.Text = "R"
            H4row.Cells.Add(H4tc4)

            Dim H4tc5 As TableCell = New TableCell
            H4tc5.Text = "A"
            H4row.Cells.Add(H4tc5)

            Dim H4tc5A As TableCell = New TableCell
            H4tc5A.Text = "F"
            H4row.Cells.Add(H4tc5A)

            Dim H4tc6 As TableCell = New TableCell
            H4tc6.Text = "R"
            H4row.Cells.Add(H4tc6)

            Dim H4tc7 As TableCell = New TableCell
            H4tc7.Text = "A"
            H4row.Cells.Add(H4tc7)

            Dim H4tc7A As TableCell = New TableCell
            H4tc7A.Text = "F"
            H4row.Cells.Add(H4tc7A)

            Dim H4tc8 As TableCell = New TableCell
            H4tc8.Text = "R"
            H4row.Cells.Add(H4tc8)

            'Dim H4tc9 As TableCell = New TableCell
            'H4tc9.Text = "A"
            'H4row.Cells.Add(H4tc9)

            'Dim H4tc9A As TableCell = New TableCell
            'H4tc9A.Text = "F"
            'H4row.Cells.Add(H4tc9A)

            'Dim H4tcA As TableCell = New TableCell
            'H4tcA.Text = "R"
            'H4row.Cells.Add(H4tcA)

            'Dim H4tcB As TableCell = New TableCell
            'H4tcB.Text = "A"
            'H4row.Cells.Add(H4tcB)

            'Dim H4tcBA As TableCell = New TableCell
            'H4tcBA.Text = "F"
            'H4row.Cells.Add(H4tcBA)

            'Dim H4tcC As TableCell = New TableCell
            'H4tcC.Text = "R"
            'H4row.Cells.Add(H4tcC)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            ' 表頭-3行
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = Version(1)
            H3tc1.ColumnSpan = 3
            H3row.Cells.Add(H3tc1)

            Dim H3tc2 As TableCell = New TableCell
            H3tc2.Text = Version(2)
            H3tc2.ColumnSpan = 3
            H3row.Cells.Add(H3tc2)

            Dim H3tc3 As TableCell = New TableCell
            H3tc3.Text = Version(3)
            H3tc3.ColumnSpan = 3
            H3row.Cells.Add(H3tc3)

            Dim H3tc4 As TableCell = New TableCell
            H3tc4.Text = Version(4)
            H3tc4.ColumnSpan = 3
            H3row.Cells.Add(H3tc4)

            'Dim H3tc5 As TableCell = New TableCell
            'H3tc5.Text = Version(5)
            'H3tc5.ColumnSpan = 3
            'H3row.Cells.Add(H3tc5)

            'Dim H3tc6 As TableCell = New TableCell
            'H3tc6.Text = Version(6)
            'H3tc6.ColumnSpan = 3
            'H3row.Cells.Add(H3tc6)

            gv.Controls(0).Controls.AddAt(0, H3row)
            '-----------------------------------------
            ' 表頭-2行
            Dim H2tc1 As TableCell = New TableCell
            H2tc1.Text = ""
            H2tc1.ColumnSpan = 3
            H2row.Cells.Add(H2tc1)

            Dim H2tc2 As TableCell = New TableCell
            H2tc2.Text = ""
            H2tc2.ColumnSpan = 3
            H2row.Cells.Add(H2tc2)

            Dim H2tc3 As TableCell = New TableCell
            H2tc3.Text = ""
            H2tc3.ColumnSpan = 3
            H2row.Cells.Add(H2tc3)

            Dim H2tc4 As TableCell = New TableCell
            H2tc4.Text = ""
            H2tc4.ColumnSpan = 3
            H2row.Cells.Add(H2tc4)

            'Dim H2tc5 As TableCell = New TableCell
            'H2tc5.Text = ""
            'H2tc5.ColumnSpan = 3
            'H2row.Cells.Add(H2tc5)

            'Dim H2tc6 As TableCell = New TableCell
            'H2tc6.Text = ""
            'H2tc6.ColumnSpan = 3
            'H2row.Cells.Add(H2tc6)

            ' Set 6 個最新 Month (上線4個月)
            For i As Integer = 1 To 4
                Dim xYY As Integer = CInt(NowYY)
                Dim xMM As Integer = CInt(NowMM)
                Dim xYM As String = NowYY + "/" + NowMM
                xMM = xMM - i + 1
                If xMM <= 0 Then
                    xYY = xYY - 1
                    xMM = xMM + 12
                End If
                If xMM > 9 Then
                    xYM = CStr(xYY) + "/" + CStr(xMM)
                Else
                    xYM = CStr(xYY) + "/" + "0" + CStr(xMM)
                End If
                If i = 1 Then H2tc1.Text = xYM
                If i = 2 Then H2tc2.Text = xYM
                If i = 3 Then H2tc3.Text = xYM
                If i = 4 Then H2tc4.Text = xYM
                'If i = 5 Then H2tc5.Text = xYM
                'If i = 6 Then H2tc6.Text = xYM
                Month(i) = Mid(xYM, 1, 4) + Mid(xYM, 6)
            Next
            gv.Controls(0).Controls.AddAt(0, H2row)

            '-----------------------------------------
            ' 表頭-1行
            Dim H1tc1 As TableCell = New TableCell
            H1tc1.Text = "成衣廠"
            H1tc1.RowSpan = 4
            H1row.Cells.Add(H1tc1)

            Dim H1tc1A As TableCell = New TableCell
            H1tc1A.Text = "季"
            H1tc1A.RowSpan = 4
            H1row.Cells.Add(H1tc1A)

            Dim H1tc2 As TableCell = New TableCell
            H1tc2.Text = "ACT & FCT"
            H1tc2.ColumnSpan = 12
            H1row.Cells.Add(H1tc2)

            gv.Controls(0).Controls.AddAt(0, H1row)
            '-----------------------------------------
        End If
    End Sub


End Class
