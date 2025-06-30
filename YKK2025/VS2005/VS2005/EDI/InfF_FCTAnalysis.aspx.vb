Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfF_FCTAnalysis
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
        Response.Cookies("PGM").Value = "InfF_FCTAnalysis.aspx"     '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        Buyer = Request.QueryString("pBuyer")
        If Buyer <> "000013T" And Buyer <> "TW0371T" Then
            Buyer = Mid(Request.QueryString("pBuyer"), 1, 6)
        End If
        UserID = Request.QueryString("pUserID")             'UserID
        For i As Integer = 1 To 6                           '最新6版本
            Version(i) = ""
            Total(i) = 0
        Next
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        CSTGridView.Visible = False
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
        Dim sql As String
        ' Buyer
        DBuyer.Text = Buyer
        ' 年
        DYY.Items.Clear()
        For i As Integer = 2020 To 2030
            Dim ListItem1 As New ListItem
            ListItem1.Text = CStr(i)
            ListItem1.Value = CStr(i)
            If i = CInt(Mid(NowDate, 1, 4)) Then
                ListItem1.Selected = True
            End If
            DYY.Items.Add(ListItem1)
        Next
        ' 月
        DMM.Items.Clear()
        For i As Integer = 1 To 12
            Dim ListItem1 As New ListItem
            If i < 10 Then
                ListItem1.Text = "0" + CStr(i)
                ListItem1.Value = "0" + CStr(i)
            Else
                ListItem1.Text = CStr(i)
                ListItem1.Value = CStr(i)
            End If
            If i = CInt(Mid(NowDate, 5, 2)) Then
                ListItem1.Selected = True
            End If
            DMM.Items.Add(ListItem1)
        Next
        ' 季
        DSeason.Items.Clear()
        sql = "SELECT Season From A_CustomerActual "
        sql &= "Where Buyer  = '" & DBuyer.Text & "' "
        sql &= "  And Month  = '" & DYY.Text + DMM.Text & "' "
        sql &= "  And FCTQty > 0 "
        sql &= "Group by Season "
        sql &= "Order by Substring(Season,3,2) + Substring(Season,1,2) Desc "
        '
        Dim dt_Season As DataTable = uDataBase.GetDataTable(sql)
        For i As Integer = 1 To dt_Season.Rows.Count
            Dim ListItem1 As New ListItem
            ListItem1.Text = dt_Season.Rows(i - 1).Item("Season")
            ListItem1.Value = dt_Season.Rows(i - 1).Item("Season")
            DSeason.Items.Add(ListItem1)
        Next
        ' 含New-FCT
        CNewFCT.Checked = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(DYY)
    '**     變更年
    '**
    '*****************************************************************
    Protected Sub DYY_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DYY.SelectedIndexChanged
        Dim sql As String
        ' 季
        DSeason.Items.Clear()
        sql = "SELECT Season From A_CustomerActual "
        sql &= "Where Buyer  = '" & DBuyer.Text & "' "
        sql &= "  And Month  = '" & DYY.Text + DMM.Text & "' "
        sql &= "  And FCTQty > 0 "
        sql &= "Group by Season "
        sql &= "Order by Substring(Season,3,2) + Substring(Season,1,2) Desc "
        '
        Dim dt_Season As DataTable = uDataBase.GetDataTable(sql)
        For i As Integer = 1 To dt_Season.Rows.Count
            Dim ListItem1 As New ListItem
            ListItem1.Text = dt_Season.Rows(i - 1).Item("Season")
            ListItem1.Value = dt_Season.Rows(i - 1).Item("Season")
            DSeason.Items.Add(ListItem1)
        Next
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(DMM)
    '**     變更月
    '**
    '*****************************************************************
    Protected Sub DMM_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DMM.SelectedIndexChanged
        Dim sql As String
        ' 季
        DSeason.Items.Clear()
        sql = "SELECT Season From A_CustomerActual "
        sql &= "Where Buyer  = '" & DBuyer.Text & "' "
        sql &= "  And Month  = '" & DYY.Text + DMM.Text & "' "
        sql &= "  And FCTQty > 0 "
        sql &= "Group by Season "
        sql &= "Order by Substring(Season,3,2) + Substring(Season,1,2) Desc "
        '
        Dim dt_Season As DataTable = uDataBase.GetDataTable(sql)
        For i As Integer = 1 To dt_Season.Rows.Count
            Dim ListItem1 As New ListItem
            ListItem1.Text = dt_Season.Rows(i - 1).Item("Season")
            ListItem1.Value = dt_Season.Rows(i - 1).Item("Season")
            DSeason.Items.Add(ListItem1)
        Next
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Inq)
    '**     查詢
    '**
    '*****************************************************************
    Protected Sub BInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BInq.Click
        Dim sql As String

        ' Set 6 個最新 Version 
        sql = "SELECT Top 6 Version From A_CustomerActual "
        Sql &= "Where Buyer  = '" & DBuyer.Text & "' "
        Sql &= "  And Season = '" & DSeason.Text & "' "
        Sql &= "  And Month  = '" & DYY.Text + DMM.Text & "' "
        sql &= "  And FCTQty > 0 "
        ' 是否含 New-FCT
        If CNewFCT.Checked = False Then
            sql &= "  And Version <> '" & "99999999999999" & "' "
        End If
        sql &= "Group by Version "
        sql &= "Order by Version Desc "
        Dim dt_Version As DataTable = uDataBase.GetDataTable(sql)
        For i As Integer = 1 To dt_Version.Rows.Count
            Version(i) = dt_Version.Rows(i - 1).Item("Version")
        Next

        ' 篩選資料
        sql = "SELECT "
        sql &= "a.CustCode As CustCode, b.CustName + '(' + a.CustCode + ')' As Customer, a.Buyer, a.Season, a.Month, "
        sql &= "'' As Qty6, '' As Ratio6, "
        sql &= "'' As Qty5, '' As Ratio5, "
        sql &= "'' As Qty4, '' As Ratio4, "
        sql &= "'' As Qty3, '' As Ratio3, "
        sql &= "'' As Qty2, '' As Ratio2, "
        sql &= "'' As Qty1, '' As Ratio1, "
        sql &= "'InfF_FCTAnalysis_01.aspx?' + "
        sql &= "'pBuyer=' + a.Buyer + "
        sql &= "'&pCustCode=' + a.CustCode + "
        sql &= "'&pSeason=' + a.Season + "
        sql &= "'&pMonth=' + a.Month + "
        If CNewFCT.Checked Then     ' 是否含 New-FCT
            sql &= "'&pNewFCT=' + '1' "
        Else
            sql &= "'&pNewFCT=' + '0' "
        End If
        sql &= " As URL "
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
        sql &= "  And a.Month  = '" & DYY.Text + DMM.Text & "' "
        sql &= "  And a.FCTQty > 0 "

        If Version(1) <> "" Then
            sql &= "  And ( "
            For i As Integer = 1 To 6
                If Version(i) <> "" Then
                    If i = 1 Then
                        sql &= " a.Version = '" & Version(i) & "' "
                    Else
                        sql &= " Or a.Version = '" & Version(i) & "' "
                    End If
                End If
            Next
            sql &= " ) "
        End If

        sql &= "Group by a.CustCode, b.CustName, a.Buyer, a.Season, a.Month "
        sql &= "Order by a.CustCode, b.CustName, a.Buyer, a.Season, a.Month "
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
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            Dim sql As String
            Dim j As Integer = 0
            Dim BefQty As Double = 0
            ' 版本明細展開
            For i As Integer = 1 To 6
                If Version(i) <> "" Then
                    sql = "SELECT Isnull(Sum(FCTQty),0) As Qty From A_CustomerActual "
                    sql &= "Where Buyer  = '" & DBuyer.Text & "' "
                    sql &= "  And Season = '" & DSeason.Text & "' "
                    sql &= "  And Month  = '" & DYY.Text + DMM.Text & "' "
                    sql &= "  And CustCode = '" & DataBinder.Eval(e.Row.DataItem, "CustCode") & "' "
                    sql &= "  And Version = '" & Version(i) & "' "
                    '
                    Dim dt_FCTQty As DataTable = uDataBase.GetDataTable(sql)
                    If dt_FCTQty.Rows.Count > 0 Then
                        e.Row.Cells(i * 2).Text = Format(dt_FCTQty.Rows(0).Item("Qty"), "###,###,###")
                        '
                        If i > 1 Then
                            If dt_FCTQty.Rows(0).Item("Qty") > 0 Then
                                e.Row.Cells(i * 2 - 1).Text = Format(BefQty / dt_FCTQty.Rows(0).Item("Qty") * 100, ".0") + "%"
                            Else
                                e.Row.Cells(i * 2 - 1).Text = "-"
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
            tc1.ColumnSpan = 2
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
    Protected Sub CSTGridView_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles CSTGridView.RowCreated
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
            H2tc1.Text = Version(1)
            H2tc1.ColumnSpan = 2
            H2row.Cells.Add(H2tc1)

            Dim H2tc2 As TableCell = New TableCell
            H2tc2.Text = Version(2)
            H2tc2.ColumnSpan = 2
            H2row.Cells.Add(H2tc2)

            Dim H2tc3 As TableCell = New TableCell
            H2tc3.Text = Version(3)
            H2tc3.ColumnSpan = 2
            H2row.Cells.Add(H2tc3)

            Dim H2tc4 As TableCell = New TableCell
            H2tc4.Text = Version(4)
            H2tc4.ColumnSpan = 2
            H2row.Cells.Add(H2tc4)

            Dim H2tc5 As TableCell = New TableCell
            H2tc5.Text = Version(5)
            H2tc5.ColumnSpan = 2
            H2row.Cells.Add(H2tc5)

            Dim H2tc6 As TableCell = New TableCell
            H2tc6.Text = Version(6)
            H2tc6.ColumnSpan = 2
            H2row.Cells.Add(H2tc6)

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
            H1tc2.Text = DYY.Text + "/" + DMM.Text + " - FCT & FCT"
            H1tc2.ColumnSpan = 12
            H1row.Cells.Add(H1tc2)

            gv.Controls(0).Controls.AddAt(0, H1row)
            '-----------------------------------------
        End If
    End Sub



End Class
