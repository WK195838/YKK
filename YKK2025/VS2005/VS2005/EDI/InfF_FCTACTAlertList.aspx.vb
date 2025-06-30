Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfF_FCTACTAlertList
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
    Dim AlertRatio As String

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
        Response.Cookies("PGM").Value = "InfF_FCTACTAlertList.aspx"     '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        NowYY = Mid(NowDate, 1, 4)
        NowMM = Mid(NowDate, 5, 2)
        Buyer = Request.QueryString("pBuyer")
        If Buyer <> "000013T" And Buyer <> "TW0371T" Then
            Buyer = Mid(Request.QueryString("pBuyer"), 1, 6)
        End If

        UserID = Request.QueryString("pUserID")             'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        DBuyer.ReadOnly = True
        ITEMGridView.Visible = False
        For i As Integer = 1 To 6                           '最新6版本
            Month(i) = ""
            Version(i) = ""
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
        DBuyer.Text = Buyer

        DRatio.Items.Clear()
        For i As Integer = 100 To 150 Step 10
            Dim ListItem1 As New ListItem
            ListItem1.Text = CStr(i)
            ListItem1.Value = CStr(i)
            If i = 100 Then
                ListItem1.Selected = True
            End If
            DRatio.Items.Add(ListItem1)
        Next
        '
        Dim sql As String
        DSeason.Items.Clear()
        sql = "SELECT Season From A_CustomerActual "
        sql &= "Where Buyer  = '" & DBuyer.Text & "' "
        'sql &= "  And Month  = '" & NowYY + NowMM & "' "
        sql &= "Group by Season "
        sql &= "Order by Season "
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
        Dim sql As String
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
        ' Set Alert Ratio 
        AlertRatio = DRatio.Text
        ' 取得資料
        If DLevel.Text = "ITEM" Then
            sql = "SELECT "
            sql &= "CustName + '(' + CustCode + ')' As Customer, '' As Color, * "
            sql &= "From V_A_CustomerActual_Item_01 "
            sql &= "Where Buyer = '" & DBuyer.Text & "' "
            sql &= "  And Season = '" & DSeason.Text & "' "
            sql &= "  And Version <> '" & "ZZZZZZZZZZZZZZ" & "' "

            sql &= "  And ( "
            For i As Integer = 1 To 4       ' (只上線4個月)
                If Month(i) <> "" Then
                    If i = 1 Then
                        sql &= "(Month = '" & Month(i) & "' And Version= '" & Version(i) & "') "
                    Else
                        sql &= " Or "
                        sql &= "(Month = '" & Month(i) & "' And Version= '" & Version(i) & "') "
                    End If
                End If
            Next
            sql &= " ) "
            If DAlertType.Text = "ACT / FCT" Then
                sql &= "  And ACTFCTRatio >= " & AlertRatio & " "
                sql &= "Order by Month Desc, ACTFCTRatio Desc, CustItem, CustCode "
            Else
                sql &= "  And FCTACTRatio >= " & AlertRatio & " "
                sql &= "Order by Month Desc, FCTACTRatio Desc , CustItem, CustCode "
            End If
        Else
            sql = "SELECT "
            sql &= "CustName + '(' + CustCode + ')' As Customer, * "
            sql &= "From V_A_CustomerActual_ItemColor_01 "
            sql &= "Where Buyer = '" & DBuyer.Text & "' "
            sql &= "  And Season = '" & DSeason.Text & "' "
            sql &= "  And Version <> '" & "ZZZZZZZZZZZZZZ" & "' "

            sql &= "  And ( "
            For i As Integer = 1 To 4       ' (只上線4個月)
                If Month(i) <> "" Then
                    If i = 1 Then
                        sql &= "(Month = '" & Month(i) & "' And Version= '" & Version(i) & "') "
                    Else
                        sql &= " Or "
                        sql &= "(Month = '" & Month(i) & "' And Version= '" & Version(i) & "') "
                    End If
                End If
            Next
            sql &= " ) "

            If DAlertType.Text = "ACT / FCT" Then
                sql &= "  And ACTFCTRatio >= " & AlertRatio & " "
                sql &= "Order by Month Desc, ACTFCTRatio Desc, CustItem, Color, CustCode "
            Else
                sql &= "  And FCTACTRatio >= " & AlertRatio & " "
                sql &= "Order by Month Desc, FCTACTRatio Desc, CustItem, Color, CustCode "
            End If

        End If
        '
        Dim dt_FCTPlan As DataTable = uDataBase.GetDataTable(sql)
        If dt_FCTPlan.Rows.Count > 0 Then
            ITEMGridView.Visible = True
            ITEMGridView.DataSource = dt_FCTPlan
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
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            Dim sql As String
            Dim AlertColor As String = ""
            Dim AlertLevel As String = ""
            Dim AlertLen As Integer = 0
            If DAlertType.Text = "ACT / FCT" Then
                AlertLevel = CStr(CInt(DataBinder.Eval(e.Row.DataItem, "ACTFCTRATIO")))
                AlertLen = Len(CStr(CInt(DataBinder.Eval(e.Row.DataItem, "ACTFCTRATIO"))))
            Else
                AlertLevel = CStr(CInt(DataBinder.Eval(e.Row.DataItem, "FCTACTRATIO")))
                AlertLen = Len(CStr(CInt(DataBinder.Eval(e.Row.DataItem, "FCTACTRATIO"))))
            End If
            '
            ' 設有AlertLevel
            For i As Integer = AlertLen + 1 To 5
                AlertLevel = "0" + AlertLevel
            Next
            '
            ' 取得 BackColor
            sql = "SELECT Top 1 Data From M_Referp "
            sql &= "Where Cat = '" & "200" & "' "
            sql &= "  And DKey LIKE '" & "ALERT-COLOR-" & "%' "

            sql &= "  And DKey <= '" & "ALERT-COLOR-" & AlertLevel & "' "
            sql &= "Order By DKey Desc "
            Dim dt_AlertColor As DataTable = uDataBase.GetDataTable(sql)
            If dt_AlertColor.Rows.Count > 0 Then
                AlertColor = Mid(dt_AlertColor.Rows(0).Item("Data"), 1, InStr(dt_AlertColor.Rows(0).Item("Data"), "/") - 1)
            End If
            '
            ' 設定 BackColor
            If AlertColor <> "" Then
                For i As Integer = 0 To 8
                    e.Row.Cells(i).BackColor = System.Drawing.ColorTranslator.FromHtml(AlertColor)
                Next
            End If
            '
            ' 設定資料格式化
            Dim Hyper1, Hyper2 As New HyperLink
            Dim URLACT, URLFCT As String

            If DLevel.Text = "ITEM" Then
                URLACT = "InfF_FCTACTAnalysis_02.aspx?" + _
                         "pBuyer=" + Buyer + _
                         "&pCustCode=" + DataBinder.Eval(e.Row.DataItem, "CustCode") + _
                         "&pSeason=" + DSeason.Text + _
                         "&pItem=" + DataBinder.Eval(e.Row.DataItem, "CustItem")

                URLFCT = "InfF_FCTACTAnalysis_02.aspx?" + _
                         "pBuyer=" + Buyer + _
                         "&pCustCode=" + DataBinder.Eval(e.Row.DataItem, "CustCode") + _
                         "&pSeason=" + DSeason.Text + _
                         "&pItem=" + DataBinder.Eval(e.Row.DataItem, "CustItem") + _
                         "&pVersion=" + DataBinder.Eval(e.Row.DataItem, "Version")
            Else
                URLACT = "InfF_FCTACTAnalysis_02.aspx?" + _
                         "pBuyer=" + Buyer + _
                         "&pCustCode=" + DataBinder.Eval(e.Row.DataItem, "CustCode") + _
                         "&pSeason=" + DSeason.Text + _
                         "&pItem=" + DataBinder.Eval(e.Row.DataItem, "CustItem") + _
                         "&pColor=" + DataBinder.Eval(e.Row.DataItem, "Color")
                URLFCT = "InfF_FCTACTAnalysis_02.aspx?" + _
                         "pBuyer=" + Buyer + _
                         "&pCustCode=" + DataBinder.Eval(e.Row.DataItem, "CustCode") + _
                         "&pSeason=" + DSeason.Text + _
                         "&pItem=" + DataBinder.Eval(e.Row.DataItem, "CustItem") + _
                         "&pColor=" + DataBinder.Eval(e.Row.DataItem, "Color") + _
                         "&pVersion=" + DataBinder.Eval(e.Row.DataItem, "Version")
            End If

            If DAlertType.Text = "ACT / FCT" Then
                Hyper1.Text = Format(DataBinder.Eval(e.Row.DataItem, "ACTQTY"), "###,###,###")
                Hyper1.NavigateUrl = URLACT + "&pMonth=" + DataBinder.Eval(e.Row.DataItem, "Month") + "&pOption=ACT"
                Hyper1.Target = "_blank"
                e.Row.Cells(6).Controls.Add(Hyper1)

                Hyper2.Text = Format(DataBinder.Eval(e.Row.DataItem, "FCTQTY"), "###,###,###")
                Hyper2.NavigateUrl = URLFCT + "&pMonth=" + DataBinder.Eval(e.Row.DataItem, "Month") + "&pOption=FCT"
                Hyper2.Target = "_blank"
                e.Row.Cells(7).Controls.Add(Hyper2)

                e.Row.Cells(8).Text = Format(DataBinder.Eval(e.Row.DataItem, "ACTFCTRATIO"), "#,###0") + "%"
            Else
                Hyper1.Text = Format(DataBinder.Eval(e.Row.DataItem, "FCTQTY"), "###,###,###")
                Hyper1.NavigateUrl = URLFCT + "&pMonth=" + DataBinder.Eval(e.Row.DataItem, "Month") + "&pOption=FCT"
                Hyper1.Target = "_blank"
                e.Row.Cells(6).Controls.Add(Hyper1)

                Hyper2.Text = Format(DataBinder.Eval(e.Row.DataItem, "ACTQTY"), "###,###,###")
                Hyper2.NavigateUrl = URLACT + "&pMonth=" + DataBinder.Eval(e.Row.DataItem, "Month") + "&pOption=ACT"
                Hyper2.Target = "_blank"
                e.Row.Cells(7).Controls.Add(Hyper2)

                e.Row.Cells(8).Text = Format(DataBinder.Eval(e.Row.DataItem, "FCTACTRATIO"), "#,###0") + "%"
            End If

            If DLevel.Text = "ITEM" Then
                e.Row.Cells(1).Visible = False
            End If
        End If
            '
            'Footer
            If e.Row.RowType = DataControlRowType.Footer Then
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
            If DAlertType.Text = "ACT / FCT" Then
                Dim H4tc1 As TableCell = New TableCell
                H4tc1.Text = "A"
                H4row.Cells.Add(H4tc1)

                Dim H4tc1A As TableCell = New TableCell
                H4tc1A.Text = "F"
                H4row.Cells.Add(H4tc1A)
            Else
                Dim H4tc1 As TableCell = New TableCell
                H4tc1.Text = "F"
                H4row.Cells.Add(H4tc1)

                Dim H4tc1A As TableCell = New TableCell
                H4tc1A.Text = "A"
                H4row.Cells.Add(H4tc1A)
            End If

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = "R"
            H4row.Cells.Add(H4tc2)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            ' 表頭-3行
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = ""
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
            H1tc1.Text = "ITEM"
            H1tc1.RowSpan = 4
            H1row.Cells.Add(H1tc1)

            If DLevel.Text <> "ITEM" Then
                Dim H1tc1B As TableCell = New TableCell
                H1tc1B.Text = "COLOR"
                H1tc1B.RowSpan = 4
                H1row.Cells.Add(H1tc1B)
            End If

            Dim H1tc1A As TableCell = New TableCell
            H1tc1A.Text = "成衣廠"
            H1tc1A.RowSpan = 4
            H1row.Cells.Add(H1tc1A)

            Dim H1tc1C As TableCell = New TableCell
            H1tc1C.Text = "季"
            H1tc1C.RowSpan = 4
            H1row.Cells.Add(H1tc1C)

            Dim H1tc1D As TableCell = New TableCell
            H1tc1D.Text = "年月"
            H1tc1D.RowSpan = 4
            H1row.Cells.Add(H1tc1D)

            Dim H1tc1E As TableCell = New TableCell
            H1tc1E.Text = "Version"
            H1tc1E.RowSpan = 4
            H1row.Cells.Add(H1tc1E)

            Dim H1tc2 As TableCell = New TableCell
            If DAlertType.Text = "ACT / FCT" Then
                H1tc2.Text = "ACT & FCT"
            Else
                H1tc2.Text = "FCT & ACT"
            End If
            H1tc2.ColumnSpan = 3
            H1row.Cells.Add(H1tc2)

            gv.Controls(0).Controls.AddAt(0, H1row)
            '-----------------------------------------
        End If
    End Sub
End Class
