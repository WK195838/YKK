Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfF_FCTACTVDPAnalysis_01
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
    Dim gLevel As String            'ITEM, ITEM+COLOR
    Dim gVersion As String          'Version
    Dim UserID As String            'UserID
    Dim Month(12) As String         '最新12版本
    Dim Version(12) As String       '最後12版本
    Dim MMSQLStr As String          '每月最後版本SQL
    Dim ACTTotal, FCTTotal As Integer   '最新版本合計數量

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
        Response.Cookies("PGM").Value = "InfF_FCTACTVDPAnalysis_01.aspx"  '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")  '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                 '現在日期時間
        gBuyer = Request.QueryString("pBuyer")             'Buyer
        gCustCode = Request.QueryString("pCustCode")       'Customer Code
        gSeason = Request.QueryString("pSeason")           'Season
        gMonth = Request.QueryString("pMonth")             'Month
        gLevel = Request.QueryString("pLevel")             'ITEM, ITEM+COLOR
        gVersion = Request.QueryString("pVersion")         'Version
        UserID = Request.QueryString("pUserID")            'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        For i As Integer = 1 To 12
            Month(i) = ""
            Version(i) = ""
        Next
        ACTTotal = 0
        FCTTotal = 0
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
        Dim sql, URLStr As String
        ' Set 最新 Month
        sql = "SELECT Month From A_CustomerActual_VDP "
        sql &= "Where Buyer  = '" & gBuyer & "' "
        sql &= "  And Season = '" & gSeason & "' "
        If gMonth <> "ALL" Then
            sql &= "  And Month = '" & gMonth & "' "
        End If
        sql &= "  And Version <> '" & "99999999999999" & "' "
        sql &= "Group by Month "
        sql &= "Order by Month "
        Dim dt_Month As DataTable = uDataBase.GetDataTable(sql)
        For i As Integer = 0 To dt_Month.Rows.Count - 1
            Month(i + 1) = dt_Month.Rows(i).Item("Month")
        Next
        '
        ' Set 最新 Version 
        If gVersion <> "NIL" Then
            ' 指定特定Version
            For i As Integer = 1 To 12
                Version(i) = gVersion
            Next
        Else
            ' 指定Season
            For i As Integer = 1 To 12
                If Month(i) <> "" Then
                    sql = "SELECT Top 1 Version From A_CustomerActual_VDP "
                    sql &= "Where Buyer  = '" & gBuyer & "' "
                    sql &= "  And Season = '" & gSeason & "' "
                    sql &= "  And Month  = '" & Month(i) & "' "
                    sql &= "  And Version <> '" & "99999999999999" & "' "
                    sql &= "  And FCTQty > 0 "
                    sql &= "Order by Version Desc "
                    Dim dt_Version As DataTable = uDataBase.GetDataTable(sql)
                    If dt_Version.Rows.Count > 0 Then
                        Version(i) = dt_Version.Rows(0).Item("Version")
                    Else
                        sql = "SELECT Top 1 Version From A_CustomerActual_VDP "
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
                End If
            Next
        End If
        'URL
        If gLevel = "ITEM" Then
            URLStr = "'InfF_FCTACTVDPAnalysis_02.aspx?' + "
        Else
            URLStr = "'InfF_FCTACTVDPAnalysis_02a.aspx?' + "
        End If
        URLStr = URLStr & "'pBuyer=' + a.Buyer + "
        URLStr = URLStr & "'&pCustCode=' + a.CustCode + "
        URLStr = URLStr & "'&pSeason=' + a.Season + "
        URLStr = URLStr & "'&pMonth=' + a.Month + "
        URLStr = URLStr & "'&pVersion=' + '" + gVersion + "'"
        '
        ' 每月最後版本SQL
        For i As Integer = 1 To 12
            If Month(i) <> "" Then
                If i = 1 Then
                    MMSQLStr = "(" + "Month='" + Month(i) + "' And Version='" + Version(i) + "')"
                Else
                    MMSQLStr &= " Or (" + "Month='" + Month(i) + "' And Version='" + Version(i) + "')"
                End If
            End If
        Next
        '
        ' 篩選資料
        sql = "SELECT "
        sql &= "a.CustCode As CustCode, b.CustName + '(' + a.CustCode + ')' As Customer, a.Buyer, a.Season, a.Month, "
        sql &= "'' As AQty1, '' As FQty1, '' As Ratio1, "
        sql &= URLStr & " As URL "
        sql &= "From A_CustomerActual_Item_VDP a, M_NativeVendor b "
        ' T&P NIKE=000013T
        If gBuyer = "000013T" Then
            sql &= "Where 'FALL-TP000013' = b.Buyer "
        Else
            sql &= "Where 'FALL-' + a.Buyer = b.Buyer "
        End If

        sql &= "  And a.CustCode = b.CustCode "
        sql &= "  And a.Buyer  = '" & gBuyer & "' "
        sql &= "  And a.CustCode  = '" & gCustCode & "' "
        sql &= "  And a.Season = '" & gSeason & "' "
        sql &= "  And ( "
        sql &= MMSQLStr
        sql &= "      ) "
        sql &= "Group by a.CustCode, b.CustName, a.Buyer, a.Season, a.Month "
        sql &= "Order by a.CustCode, b.CustName, a.Buyer, a.Season, a.Month "
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
            Dim sql, xVersion As String
            Dim MMTotal As Integer = 0
            '
            ' 取得Version
            xVersion = ""
            For i As Integer = 1 To 12
                If Month(i) = DataBinder.Eval(e.Row.DataItem, "Month") Then
                    xVersion = Version(i)
                End If
            Next
            '
            ' 版本明細展開
            sql = "SELECT Top 1 Isnull(Sum(FCTQty),0) As FCTQty, Isnull(Sum(ACTQty),0) As ACTQty From A_CustomerActual_Item_VDP "
            sql &= "Where Buyer  = '" & gBuyer & "' "
            sql &= "  And CustCode = '" & gCustCode & "' "
            sql &= "  And Season = '" & gSeason & "' "
            sql &= "  And Month = '" & DataBinder.Eval(e.Row.DataItem, "Month") & "' "
            sql &= "  And Version = '" & xVersion & "' "
            sql &= "Group by  Buyer, CustCode, Month "
            sql &= "Order by  Buyer, CustCode, Month "
            '
            Dim dt_FCTQty As DataTable = uDataBase.GetDataTable(sql)
            If dt_FCTQty.Rows.Count > 0 Then
                e.Row.Cells(3).Text = Format(dt_FCTQty.Rows(0).Item("ACTQty"), "###,###,###")
                e.Row.Cells(4).Text = Format(dt_FCTQty.Rows(0).Item("FCTQty"), "###,###,###")
                '
                If dt_FCTQty.Rows(0).Item("FCTQty") > 0 Then
                    e.Row.Cells(5).Text = Format(dt_FCTQty.Rows(0).Item("ACTQty") / dt_FCTQty.Rows(0).Item("FCTQty") * 100, ".0") + "%"
                Else
                    e.Row.Cells(5).Text = ""
                End If
                '
                ACTTotal = ACTTotal + dt_FCTQty.Rows(0).Item("ACTQty")
                FCTTotal = FCTTotal + dt_FCTQty.Rows(0).Item("FCTQty")
                MMTotal = MMTotal + dt_FCTQty.Rows(0).Item("FCTQty") + dt_FCTQty.Rows(0).Item("ACTQty")
            End If
            ' 無任何資料時隱藏 ROW
            If MMTotal = 0 Then
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
            tc1.ColumnSpan = 3
            tc1.Text = "合計"
            row.Cells.Add(tc1)

            Dim tc2 As TableCell = New TableCell
            tc2.HorizontalAlign = HorizontalAlign.Right
            tc2.BackColor = Color.YellowGreen
            tc2.Text = Format(ACTTotal, "###,###,###")
            row.Cells.Add(tc2)

            Dim tc2A As TableCell = New TableCell
            tc2A.HorizontalAlign = HorizontalAlign.Right
            tc2A.BackColor = Color.YellowGreen
            tc2A.Text = Format(FCTTotal, "###,###,###")
            row.Cells.Add(tc2A)

            Dim tc3 As TableCell = New TableCell
            tc3.HorizontalAlign = HorizontalAlign.Right
            tc3.BackColor = Color.YellowGreen
            If ACTTotal > 0 Then
                tc3.Text = Format(ACTTotal / FCTTotal * 100, ".0") + "%"
            Else
                tc3.Text = ""
            End If
            row.Cells.Add(tc3)

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

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            ' 表頭-3行
            Dim H3tc1 As TableCell = New TableCell
            If gVersion <> "NIL" Then
                H3tc1.Text = Version(1)
            Else
                H3tc1.Text = ""
            End If
            H3tc1.ColumnSpan = 3
            H3row.Cells.Add(H3tc1)

            gv.Controls(0).Controls.AddAt(0, H3row)
            '-----------------------------------------
            ' 表頭-2行
            Dim H2tc1 As TableCell = New TableCell
            H2tc1.Text = ""
            H2tc1.ColumnSpan = 3
            H2row.Cells.Add(H2tc1)

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
            H1tc1B.Text = "月"
            H1tc1B.RowSpan = 4
            H1row.Cells.Add(H1tc1B)

            Dim H1tc2 As TableCell = New TableCell
            H1tc2.Text = "FCT & ACT"
            H1tc2.ColumnSpan = 3
            H1row.Cells.Add(H1tc2)

            gv.Controls(0).Controls.AddAt(0, H1row)
            '-----------------------------------------
        End If
    End Sub
End Class
