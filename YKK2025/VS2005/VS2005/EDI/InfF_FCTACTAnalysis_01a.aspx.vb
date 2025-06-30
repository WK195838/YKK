Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfF_FCTACTAnalysis_01a
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
    Dim gVersion As String          'Version
    Dim UserID As String            'UserID
    Dim Month(6) As String          '最新6版本
    Dim Version(6) As String        '最後6版本
    Dim ACTTotal(6), FCTTotal(6) As Integer         '最新6版本合計數量
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
            ShowItemDataList()                      '顯示資料
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
        Response.Cookies("PGM").Value = "InfF_FCTACTAnalysis_01a.aspx"  '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")  '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                 '現在日期時間
        gBuyer = Request.QueryString("pBuyer")             'Buyer
        gCustCode = Request.QueryString("pCustCode")       'Customer Code
        gSeason = Request.QueryString("pSeason")           'Season
        gMonth = Request.QueryString("pMonth")             'Month
        gVersion = Request.QueryString("pVersion")         'Version
        UserID = Request.QueryString("pUserID")            'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        For i As Integer = 1 To 6                          '最新6版本
            Month(i) = ""
            Version(i) = ""
            ACTTotal(i) = 0
            FCTTotal(i) = 0
        Next
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
    '**(ShowItemDataList)
    '**     顯示資料
    '**
    '*****************************************************************
    Sub ShowItemDataList()
        Dim sql As String
        Dim MMsql As String = ""
        ' Set 4 個最新 Month 
        Dim xYY As String = Mid(gMonth, 1, 4)
        Dim xMM As String = Mid(gMonth, 5, 2)
        For i As Integer = 0 To 3
            If CInt(xMM) - i > 0 Then
                If CInt(xMM) - i < 10 Then
                    Month(i + 1) = xYY + "0" + CStr(CInt(xMM) - i)
                Else
                    Month(i + 1) = xYY + CStr(CInt(xMM) - i)
                End If
            Else
                If CInt(xMM) + 12 - i < 10 Then
                    Month(i + 1) = CStr(CInt(xYY) - 1) + "0" + CStr(CInt(xMM) + 12 - i)
                Else
                    Month(i + 1) = CStr(CInt(xYY) - 1) + CStr(CInt(xMM) + 12 - i)
                End If
            End If
            '
            If i = 0 Then
                MMsql = " a.Month = '" & Month(i + 1) & "' "
            Else
                MMsql = MMsql + " Or a.Month = '" & Month(i + 1) & "' "
            End If
        Next
        ' Set 4 個最新 Version 
        If gVersion <> "NIL" Then
            ' 指定特定Version
            For i As Integer = 1 To 4
                Version(i) = gVersion
            Next
        Else
            ' 指定Season
            For i As Integer = 1 To 4
                sql = "SELECT Top 1 Version From A_CustomerActual "
                sql &= "Where Buyer  = '" & gBuyer & "' "
                sql &= "  And Season = '" & gSeason & "' "
                sql &= "  And Month  = '" & Month(i) & "' "
                sql &= "  And FCTQty > 0 "
                sql &= "  And Version <> '" & "99999999999999" & "' "
                sql &= "Order by Version Desc "
                Dim dt_Version As DataTable = uDataBase.GetDataTable(sql)
                If dt_Version.Rows.Count > 0 Then
                    Version(i) = dt_Version.Rows(0).Item("Version")
                Else
                    sql = "SELECT Top 1 Version From A_CustomerActual "
                    sql &= "Where Buyer  = '" & gBuyer & "' "
                    sql &= "  And Season = '" & gSeason & "' "
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
        ' 篩選資料
        sql = "SELECT "
        sql &= "a.CustCode As CustCode, b.CustName + '(' + a.CustCode + ')' As Customer, a.Buyer, a.Season, a.CustItem, a.KeyData2, "
        sql &= "'' As AQty6, '' As FQty6, '' As Ratio6, "
        sql &= "'' As AQty5, '' As FQty5, '' As Ratio5, "
        sql &= "'' As AQty4, '' As FQty4, '' As Ratio4, "
        sql &= "'' As AQty3, '' As FQty3, '' As Ratio3, "
        sql &= "'' As AQty2, '' As FQty2, '' As Ratio2, "
        sql &= "'' As AQty1, '' As FQty1, '' As Ratio1 "
        sql &= "From A_CustomerActual_Item a, M_NativeVendor b "
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
        sql &= "  And (" + MMsql + ") "

        sql &= "Group by a.CustCode, b.CustName, a.Buyer, a.Season, a.CustItem, a.KeyData2 "
        sql &= "Order by a.CustCode, b.CustName, a.Buyer, a.Season, a.CustItem, a.KeyData2 "
        '
        Dim dt_FCTPlan_ITEM As DataTable = uDataBase.GetDataTable(sql)
        If dt_FCTPlan_ITEM.Rows.Count > 0 Then
            ITEMGridView.Visible = True
            ITEMGridView.DataSource = dt_FCTPlan_ITEM
            ITEMGridView.DataBind()
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
    Protected Sub ITEMGridView_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles ITEMGridView.RowDataBound
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            Dim sql As String
            Dim j As Integer = 0
            Dim PLMTotal As Integer = 0
            ' 版本明細展開
            For i As Integer = 1 To 4
                If Month(i) <> "" Then
                    Dim Hyper1, Hyper2, Hyper3, Hyper4 As New HyperLink
                    Dim URLACT As String = "InfF_FCTACTAnalysis_02.aspx?" + _
                                           "pBuyer=" + gBuyer + _
                                           "&pCustCode=" + DataBinder.Eval(e.Row.DataItem, "CustCode") + _
                                           "&pSeason=" + gSeason + _
                                           "&pItem=" + Replace(DataBinder.Eval(e.Row.DataItem, "CustItem"), "#", "%23") + _
                                           "&pColor=" + DataBinder.Eval(e.Row.DataItem, "KeyData2")

                    Dim URLFCT As String = "InfF_FCTACTAnalysis_02.aspx?" + _
                                           "pBuyer=" + gBuyer + _
                                           "&pCustCode=" + DataBinder.Eval(e.Row.DataItem, "CustCode") + _
                                           "&pSeason=" + gSeason + _
                                           "&pItem=" + Replace(DataBinder.Eval(e.Row.DataItem, "CustItem"), "#", "%23") + _
                                           "&pColor=" + DataBinder.Eval(e.Row.DataItem, "KeyData2") + _
                                           "&pVersion=" + Version(i)
                    '
                    sql = "SELECT Top 1 Isnull(Sum(FCTQty),0) As FCTQty, Isnull(Sum(ACTQty),0) As ACTQty From A_CustomerActual_Item "
                    sql &= "Where Buyer  = '" & gBuyer & "' "
                    sql &= "  And Season = '" & gSeason & "' "
                    sql &= "  And CustCode = '" & DataBinder.Eval(e.Row.DataItem, "CustCode") & "' "
                    sql &= "  And Month  = '" & Month(i) & "' "
                    sql &= "  And CustItem = '" & DataBinder.Eval(e.Row.DataItem, "CustItem") & "' "
                    sql &= "  And KeyData2 = '" & DataBinder.Eval(e.Row.DataItem, "KeyData2") & "' "
                    sql &= "  And Version = '" & Version(i) & "' "
                    sql &= "Group by  Buyer, CustCode, CustItem, KeyData2, Version "
                    sql &= "Order by  Buyer, CustCode, CustItem, KeyData2, Version Desc "
                    '
                    Dim dt_FCTQty As DataTable = uDataBase.GetDataTable(sql)
                    If dt_FCTQty.Rows.Count > 0 Then
                        If i = 1 Then
                            Hyper1.Text = Format(dt_FCTQty.Rows(0).Item("ACTQty"), "###,###,###")
                            Hyper1.NavigateUrl = URLACT + "&pMonth=" + Month(i) + "&pOption=ACT"
                            Hyper1.Target = "_blank"
                            e.Row.Cells(i + 3).Controls.Add(Hyper1)

                            Hyper2.Text = Format(dt_FCTQty.Rows(0).Item("FCTQty"), "###,###,###")
                            Hyper2.NavigateUrl = URLFCT + "&pMonth=" + Month(i) + "&pOption=FCT"
                            Hyper2.Target = "_blank"
                            e.Row.Cells(i + 4).Controls.Add(Hyper2)
                            j = 5
                        Else
                            Hyper3.Text = Format(dt_FCTQty.Rows(0).Item("ACTQty"), "###,###,###")
                            Hyper3.NavigateUrl = URLACT + "&pMonth=" + Month(i) + "&pOption=ACT"
                            Hyper3.Target = "_blank"
                            e.Row.Cells(i * 2 + i + 1).Controls.Add(Hyper3)

                            Hyper4.Text = Format(dt_FCTQty.Rows(0).Item("FCTQty"), "###,###,###")
                            Hyper4.NavigateUrl = URLFCT + "&pMonth=" + Month(i) + "&pOption=FCT"
                            Hyper4.Target = "_blank"
                            e.Row.Cells(i * 2 + i + 2).Controls.Add(Hyper4)
                            j = i * 2 + i + 2
                        End If
                        '
                        If dt_FCTQty.Rows(0).Item("FCTQty") > 0 Then
                            e.Row.Cells(j + 1).Text = Format(dt_FCTQty.Rows(0).Item("ACTQty") / dt_FCTQty.Rows(0).Item("FCTQty") * 100, ".0") + "%"
                        Else
                            e.Row.Cells(j + 1).Text = ""
                        End If
                        '
                        PLMTotal = PLMTotal + dt_FCTQty.Rows(0).Item("FCTQty") + dt_FCTQty.Rows(0).Item("ACTQty")
                        ACTTotal(i) = ACTTotal(i) + dt_FCTQty.Rows(0).Item("ACTQty")
                        FCTTotal(i) = FCTTotal(i) + dt_FCTQty.Rows(0).Item("FCTQty")
                    End If
                End If
            Next
            ' 無任何資料時隱藏 ROW
            If PLMTotal = 0 Then
                e.Row.Visible = False
            End If
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
            ' 合計展開
            Dim row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)

            Dim tc1 As TableCell = New TableCell
            tc1.BackColor = Color.YellowGreen
            tc1.ColumnSpan = 4
            tc1.Text = "合計"
            row.Cells.Add(tc1)

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

            e.Row.Parent.Controls.Add(row)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯表頭
    '**
    '*****************************************************************
    Protected Sub ITEMGridView_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles ITEMGridView.RowCreated
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

            ' Set 6 個最新 Month 
            For i As Integer = 1 To 4
                Dim xYY As Integer = CInt(Mid(gMonth, 1, 4))
                Dim xMM As Integer = CInt(Mid(gMonth, 5, 2))
                Dim xYM As String = gMonth
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

            Dim H1tc1B As TableCell = New TableCell
            H1tc1B.Text = "ITEM"
            H1tc1B.RowSpan = 4
            H1row.Cells.Add(H1tc1B)

            Dim H1tc1C As TableCell = New TableCell
            H1tc1C.Text = "COLOR"
            H1tc1C.RowSpan = 4
            H1row.Cells.Add(H1tc1C)

            Dim H1tc2 As TableCell = New TableCell
            H1tc2.Text = "FCT & ACT"
            H1tc2.ColumnSpan = 12
            H1row.Cells.Add(H1tc2)

            gv.Controls(0).Controls.AddAt(0, H1row)
            '-----------------------------------------
        End If
    End Sub

End Class
