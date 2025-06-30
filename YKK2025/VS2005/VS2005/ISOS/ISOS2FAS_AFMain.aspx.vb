Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class ISOS2FAS_AFMain
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate, NowYY, NowMM As String           '現在日期時間
    Dim pBuyer As String            'Buyer
    Dim UserID As String            'UserID
    Dim Month(12) As String         '最新12版本
    Dim Version(12) As String       '最後12版本
    Dim MMSQLStr As String          '每月最後版本SQL
    Dim ACTTotal, FCTTotal As Integer         '最新版本合計數量

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
        Response.Cookies("PGM").Value = "ISOS2FAS_AFMain.aspx"     '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        NowYY = Mid(NowDate, 1, 4)
        NowMM = Mid(NowDate, 5, 2)
        UserID = Request.QueryString("pUserID")             'UserID
        pBuyer = Request.QueryString("pBuyer")
        If pBuyer <> "" Then
            If pBuyer = "000021" Then DBuyer.Text = "TNF"
            If pBuyer = "000001" Then DBuyer.Text = "ADIDAS"
            If pBuyer = "000013" Then DBuyer.Text = "NIKE"
            If pBuyer = "000003" Then DBuyer.Text = "COLUMBIA"
            If pBuyer = "TW1741" Then DBuyer.Text = "LULULEMON"
            If pBuyer = "TW0371" Then DBuyer.Text = "UNDER ARMOUR"
        End If

        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        DBuyer.ReadOnly = True
        CSTGridView.Visible = False
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
        Dim sql As String
        '
        DMM.Items.Clear()
        Dim ListItem As New ListItem
        ListItem.Text = "ALL"
        ListItem.Value = "ALL"
        ListItem.Selected = True
        DMM.Items.Add(ListItem)
        For i As Integer = 1 To 12
            Dim ListItem0 As New ListItem
            If i < 10 Then
                ListItem0.Text = "0" & CStr(i)
                ListItem0.Value = "0" & CStr(i)
            Else
                ListItem0.Text = CStr(i)
                ListItem0.Value = CStr(i)
            End If
            DMM.Items.Add(ListItem0)
        Next
        '
        DSeason.Items.Clear()
        sql = "SELECT Top 4 Season From A_CustomerActual_VDP "
        sql &= "Where Buyer  = '" & pBuyer & "' "
        sql &= "Group by Season "
        sql &= "Order by Substring(Season,3,2) + Substring(Season,1,2) Desc "
        Dim dt_Season As DataTable = uDataBase.GetDataTable(sql)
        If dt_Season.Rows.Count > 0 Then
            DSeason.Enabled = True
            BInq.Enabled = True
            For i As Integer = 0 To dt_Season.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = dt_Season.Rows(i).Item("Season")
                ListItem1.Value = dt_Season.Rows(i).Item("Season")
                DSeason.Items.Add(ListItem1)
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
        ' Set 最新 Month
        sql = "SELECT Month From A_CustomerActual_VDP "
        sql &= "Where Buyer  = '" & pBuyer & "' "
        sql &= "  And Season = '" & DSeason.Text & "' "
        If DMM.Text <> "ALL" Then
            sql &= "  And Month = '" & DMM.Text & "' "
        End If
        sql &= "  And Version <> '" & "99999999999999" & "' "
        sql &= "Group by Month "
        sql &= "Order by Month "
        Dim dt_Month As DataTable = uDataBase.GetDataTable(sql)
        For i As Integer = 0 To dt_Month.Rows.Count - 1
            Month(i + 1) = dt_Month.Rows(i).Item("Month")
        Next
        '
        ' 指定Season
        For i As Integer = 1 To 12
            If Month(i) <> "" Then
                sql = "SELECT Top 1 Version From A_CustomerActual_VDP "
                sql &= "Where Buyer  = '" & pBuyer & "' "
                sql &= "  And Season = '" & DSeason.Text & "' "
                sql &= "  And Month  = '" & Month(i) & "' "
                sql &= "  And Version <> '" & "99999999999999" & "' "
                sql &= "  And FCTQty > 0 "
                sql &= "Order by Version Desc "
                Dim dt_Version As DataTable = uDataBase.GetDataTable(sql)
                If dt_Version.Rows.Count > 0 Then
                    Version(i) = dt_Version.Rows(0).Item("Version")
                Else
                    sql = "SELECT Top 1 Version From A_CustomerActual_VDP "
                    sql &= "Where Buyer  = '" & pBuyer & "' "
                    sql &= "  And Season = '" & DSeason.Text & "' "
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
        '
        'URL
        If DLevel.Text = "ITEM" Then
            URLStr = "'http://10.245.0.3/ISOS/ISOS2FAS_AFItem_01.aspx?' + "
        Else
            URLStr = "'http://10.245.0.3/ISOS/ISOS2FAS_AFItemColor_01.aspx?' + "
        End If
        URLStr = URLStr & "'pBuyer=' + a.Buyer + "
        URLStr = URLStr & "'&pCustCode=' + a.CustCode + "
        URLStr = URLStr & "'&pSeason=' + a.Season + "
        URLStr = URLStr & "'&pMonth=' + '" + DMM.Text + "' + "
        URLStr = URLStr & "'&pLevel=' + '" + DLevel.Text + "' + "
        URLStr = URLStr & "'&pVersion=' + '" + "NIL" + "'"
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
        ' 顯示資料
        sql = "SELECT "
        sql &= "a.CustCode As CustCode, b.CustName + '(' + a.CustCode + ')' As Customer, a.Buyer, a.Season, "
        sql &= "'' As AQty1, '' As FQty1, '' As Ratio1, "
        sql &= URLStr & " As URL "
        sql &= "From A_CustomerActual_VDP a, M_NativeVendor b "
        sql &= "Where 'FALL-' + a.Buyer = b.Buyer "

        sql &= "  And a.CustCode = b.CustCode "
        sql &= "  And a.Buyer  = '" & pBuyer & "' "
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
            '
            ' ACT / FCT 比較
            sql = "SELECT Top 1 Isnull(Sum(FCTQty),0) As FCTQty, Isnull(Sum(ACTQty),0) As ACTQty From A_CustomerActual_VDP "
            sql &= "Where Buyer  = '" & pBuyer & "' "
            sql &= "  And CustCode = '" & DataBinder.Eval(e.Row.DataItem, "CustCode") & "' "
            sql &= "  And Season = '" & DSeason.Text & "' "
            sql &= "  And ( "
            sql &= MMSQLStr
            sql &= "      ) "
            sql &= "Group by  Buyer, CustCode "
            sql &= "Order by  Buyer, CustCode "
            '
            Dim dt_FCTQty As DataTable = uDataBase.GetDataTable(sql)
            If dt_FCTQty.Rows.Count > 0 Then
                e.Row.Cells(2).Text = Format(dt_FCTQty.Rows(0).Item("ACTQty"), "###,###,###")
                e.Row.Cells(3).Text = Format(dt_FCTQty.Rows(0).Item("FCTQty"), "###,###,###")
                '
                If dt_FCTQty.Rows(0).Item("FCTQty") > 0 Then
                    e.Row.Cells(4).Text = Format(dt_FCTQty.Rows(0).Item("ACTQty") / dt_FCTQty.Rows(0).Item("FCTQty") * 100, ".0") + "%"
                    '
                    If (dt_FCTQty.Rows(0).Item("ACTQty") / dt_FCTQty.Rows(0).Item("FCTQty") * 100) >= 130 Or _
                       (dt_FCTQty.Rows(0).Item("ACTQty") / dt_FCTQty.Rows(0).Item("FCTQty") * 100) <= 70 Then
                        e.Row.Cells(4).ForeColor = Color.Red
                    End If
                Else
                    e.Row.Cells(4).Text = ""
                    e.Row.Cells(4).ForeColor = Color.Red
                End If
                '
                ACTTotal = ACTTotal + dt_FCTQty.Rows(0).Item("ACTQty")
                FCTTotal = FCTTotal + dt_FCTQty.Rows(0).Item("FCTQty")
            End If
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
                If (ACTTotal / FCTTotal * 100) >= 130 Or (ACTTotal / FCTTotal * 100) <= 70 Then
                    tc3.ForeColor = Color.Red
                End If
            Else
                tc3.Text = ""
                tc3.ForeColor = Color.Red
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
            H3tc1.Text = Version(1)
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

            Dim H1tc2 As TableCell = New TableCell
            H1tc2.Text = "ACT & FCT"
            H1tc2.ColumnSpan = 3
            H1row.Cells.Add(H1tc2)

            gv.Controls(0).Controls.AddAt(0, H1row)
            '-----------------------------------------
        End If
    End Sub

End Class
