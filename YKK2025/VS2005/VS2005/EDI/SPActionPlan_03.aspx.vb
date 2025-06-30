Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class SPActionPlan_03
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim UserID As String            'UserID
    Dim pItem As String
    Dim pColor As String
    Dim pKeep As String
    Dim xDemandQ As Decimal

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
        SetParameter()                          '設定參數
        If Not IsPostBack Then                  'PostBack
            SetDefaultValue()                   '設定初值
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
        Server.ScriptTimeout = 900                          '設定逾時時間
        Response.Cookies("PGM").Value = "SPActionPlan_03.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        UserID = Request.QueryString("pUserID")             'UserID
        pItem = Request.QueryString("pItem")
        pColor = Request.QueryString("pColor")
        pKeep = Request.QueryString("pKeep")

        AtClose.Style("left") = -500 & "px"
        AtClose.Checked = False

        xDemandQ = 0
        '
        '動作按鈕設定
        '
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        '-----------------------------------------------------------------
        '-- Not IsPostBack
        '-----------------------------------------------------------------
        GridView1.Visible = False
        GridView2.Visible = False
        '
    End Sub
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim SQL As String
        '
        SQL = "select "
        SQL = SQL & "Item, ItemName1+' '+ItemName2 as ItemName, Color, Keep, "
        SQL = SQL & "YY, ActionName, "
        '
        SQL = SQL & "[01_Qty], [02_Qty], [03_Qty], [04_Qty], [05_Qty], [06_Qty], "
        SQL = SQL & "[07_Qty], [08_Qty], [09_Qty], [10_Qty], [11_Qty], [12_Qty], "
        SQL = SQL & "0 AS Balance "
        '
        SQL = SQL & "from M_SPBuybackMaterial "

        SQL = SQL & "where Item <> 'ZZZZZZZ' "
        ' code
        If pItem <> "" Then
            SQL = SQL & "and item = '" & pItem & "' "
        End If
        ' Color
        If pColor <> "" Then
            SQL = SQL & "and ltrim(rtrim(Color)) = '" & pColor & "' "
        End If
        ' keep
        If pKeep <> "" Then
            SQL = SQL & "and keep = '" & pKeep & "' "
        End If
        '
        SQL = SQL & "Order by Item, Color, Keep, YY Desc, Action "
        '
        Dim dt As DataTable = uDataBase.GetDataTable(SQL)
        If dt.Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = dt
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView1 編輯資料 TITLE
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Dim i As Integer
        '
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            '
            ' 清除
            e.Row.Cells.Clear()
            '
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '
            '** 4 line *****************
            i = 0
            'Item
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Code"
            tcl(i).BackColor = Color.Blue
            i = i + 1
            'Name
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Name"
            tcl(i).BackColor = Color.Blue
            i = i + 1
            'Color
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Color"
            tcl(i).BackColor = Color.Blue
            i = i + 1
            'keep
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Keep"
            tcl(i).BackColor = Color.Blue
            i = i + 1
            '[YY]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Year"
            tcl(i).BackColor = Color.Purple
            i = i + 1
            '[Action]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Action"
            tcl(i).BackColor = Color.Purple
            i = i + 1
            '-----
            '[01]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "01"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '[02]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "02"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '[03]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "03"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '[04]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "04"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '[05]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "05"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '[06]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "06"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '[07]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "07"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '[08]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "08"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '[09]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "09"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '[10]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "10"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '[11]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "11"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '[12]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "12"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '[Balance]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Balance"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '
        End If
        '
    End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**
    ''**     GridView1 編輯資料 顏色+格式 (1)
    ''**
    ''*****************************************************************
    'Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
    '    Dim i, j As Integer
    '    Dim xFCQ As Decimal
    '    '
    '    'SUM QTY
    '    xFCQ = 0
    '    '
    '    If (e.Row.RowType = DataControlRowType.DataRow) Then

    '        e.Row.Cells(23).Text = ""
    '        '
    '        For i = 0 To 23
    '            If i = 10 Or i = 13 Or i = 16 Or i = 19 Or i = 22 Then
    '                If InStr(e.Row.Cells(i).Text, "/#") > 0 Then
    '                    e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, "#", "<br>")
    '                End If
    '            End If
    '            '
    '            Select Case i
    '                Case 4
    '                    If InStr(e.Row.Cells(i).Text, "LATE-ADDED") > 0 Then
    '                        e.Row.Cells(i).ForeColor = Color.Blue
    '                        e.Row.Cells(i).Font.Bold = True
    '                    End If
    '                Case Is > 6
    '                    If e.Row.Cells(7).Text = "DUMMY" Then
    '                        For j = 7 To 22
    '                            e.Row.Cells(j).Text = ""
    '                            'Select Case j
    '                            '    Case 9, 12, 15, 18, 21
    '                            '        e.Row.Cells(j).ForeColor = Color.White
    '                            '    Case Else
    '                            '        e.Row.Cells(j).ForeColor = Color.Blue
    '                            '        e.Row.Cells(j).Text = "---"
    '                            'End Select
    '                        Next
    '                    End If

    '                    If e.Row.Cells(7).Text = "FC" Or e.Row.Cells(7).Text = "ORDERS" Then
    '                        If e.Row.Cells(7).Text = "FC" Then
    '                            e.Row.Cells(i).BackColor = Color.LemonChiffon
    '                            e.Row.Cells(i).Font.Bold = True
    '                            e.Row.Cells(i).Attributes.Add("style", "border:1px solid black ")
    '                        End If
    '                        If e.Row.Cells(7).Text = "ORDERS" Then
    '                            e.Row.Cells(i).BackColor = Color.PowderBlue
    '                            'e.Row.Cells(i).Font.Bold = True
    '                            e.Row.Cells(i).Attributes.Add("style", "border:1px solid black ")
    '                        End If
    '                    Else
    '                        If i = 23 Then
    '                            e.Row.Cells(i).ForeColor = Color.Red
    '                            e.Row.Cells(i).Font.Bold = True
    '                        End If
    '                    End If
    '                Case Else
    '            End Select
    '        Next

    '        '合計(FCQTY)
    '        If e.Row.Cells(7).Text = "FC" Then
    '            xFCQ = CDbl(e.Row.Cells(9).Text) + CDbl(e.Row.Cells(12).Text) + CDbl(e.Row.Cells(15).Text) + _
    '                   CDbl(e.Row.Cells(18).Text) + CDbl(e.Row.Cells(21).Text)
    '            e.Row.Cells(23).Text = Format(xFCQ, "###,###,##0.00")
    '        End If

    '    End If
    '    '
    '    '---------------------------------------
    '    '
    '    'Footer
    '    If e.Row.RowType = DataControlRowType.Footer Then
    '    End If
    '    '
    'End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView1 合併儲存格 (2)
    '**
    '*****************************************************************
    Protected Sub GridView1_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PreRender
        Dim i As Integer = 1
        Dim idx As Integer = 4
        Dim wRowSpan As Integer = 1
        Dim MergeIdx As Integer = 0
        Dim mySingleRow As GridViewRow
        Dim mergeColumns(10) As Integer
        '
        Dim lnk As HyperLink
        Dim xKey, yKey As String
        '
        mergeColumns(0) = 0
        mergeColumns(1) = 1
        mergeColumns(2) = 2
        mergeColumns(3) = 3
        mergeColumns(4) = 4
        mergeColumns(5) = 99
        mergeColumns(6) = 99
        mergeColumns(7) = 99
        mergeColumns(8) = 99
        mergeColumns(9) = 99
        '
        '合併儲存格
        For MergeIdx = 0 To 9

            i = 1
            For Each mySingleRow In GridView1.Rows          'Read Gridview
                If CInt(mySingleRow.RowIndex) > 0 Then      'Gridview Data Record > 0
                    '
                    If mergeColumns(MergeIdx) <> 99 Then    '有效合併儲存格                    '    '
                        '是否同SPNO : Work Gridview.cell = Gridview1.cell ? 
                        '
                        'lnk = mySingleRow.Cells(4).Controls(0)
                        'xSPNo = lnk.Text.Trim()
                        'lnk = GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(4).Controls(0)
                        'ySPNo = lnk.Text.Trim()
                        xKey = mySingleRow.Cells(0).Text.Trim & mySingleRow.Cells(1).Text.Trim & _
                               mySingleRow.Cells(2).Text.Trim & mySingleRow.Cells(3).Text.Trim & mySingleRow.Cells(4).Text.Trim
                        yKey = GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(0).Text.Trim & GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(1).Text.Trim & _
                               GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(2).Text.Trim & GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(3).Text.Trim & _
                               GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(4).Text.Trim
                        '
                        'If mySingleRow.Cells(mergeColumns(idx)).Text.Trim() = GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(idx)).Text.Trim() Then
                        If xKey = yKey Then
                            '
                            '合併處理
                            If mergeColumns(MergeIdx) < 5 Then  '是否合併欄位
                                GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(MergeIdx)).RowSpan += 1
                                mySingleRow.Cells(mergeColumns(MergeIdx)).Visible = False
                                i = i + 1
                            End If
                        Else
                            '
                            '不合併處理
                            GridView1.Rows(CInt(mySingleRow.RowIndex)).Cells(mergeColumns(MergeIdx)).RowSpan = 1
                            i = 1
                        End If
                    End If
                    '
                Else  'Gridview Data Record <= 0
                    If mergeColumns(MergeIdx) <> 99 Then    '有效合併儲存格
                        mySingleRow.Cells(mergeColumns(MergeIdx)).RowSpan = 1
                    End If
                End If
            Next

        Next
        ''
        ''FC/DEMAND 合併處理
        'i = 1
        'For Each mySingleRow In GridView1.Rows          'Read Gridview
        '    '
        '    If GridView1.Rows(CInt(mySingleRow.RowIndex)).Cells(7).Text.Trim() <> "FC" Then
        '        '
        '        If mySingleRow.Cells(23).Text.Trim() = GridView1.Rows(CInt(mySingleRow.RowIndex) - 1).Cells(23).Text.Trim() And _
        '           GridView1.Rows(CInt(mySingleRow.RowIndex) - 1).Cells(7).Text.Trim() <> "FC" Then
        '            '
        '            '合併處理
        '            GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(23).RowSpan += 1
        '            mySingleRow.Cells(23).Visible = False
        '            i = i + 1
        '        Else
        '            '
        '            '不合併處理
        '            GridView1.Rows(CInt(mySingleRow.RowIndex)).Cells(23).RowSpan = 1
        '            i = 1
        '        End If
        '    Else
        '        i = 1
        '    End If
        '    '
        'Next

        ''合計(DEMAND - QTY)
        'For Each mySingleRow In GridView1.Rows          'Read Gridview
        '    '
        '    If mySingleRow.Cells(7).Text <> "FC" And mySingleRow.Cells(7).Text <> "ORDERS" And mySingleRow.Cells(7).Text <> "" Then
        '        xDemandQ = xDemandQ + CDbl(mySingleRow.Cells(9).Text) + CDbl(mySingleRow.Cells(12).Text) + CDbl(mySingleRow.Cells(15).Text) + CDbl(mySingleRow.Cells(18).Text) + CDbl(mySingleRow.Cells(21).Text)
        '        '
        '        If mySingleRow.Cells(7).Text = "N/A" Then
        '            GridView1.Rows(CInt(mySingleRow.RowIndex) - 3).Cells(23).Text = Format(xDemandQ, "###,###,##0.00")
        '            GridView1.Rows(CInt(mySingleRow.RowIndex) - 3).Cells(23).BackColor = Drawing.Color.LemonChiffon
        '            xDemandQ = 0
        '        End If
        '    End If
        '    '
        'Next
        '
    End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**
    ''**     GridView1 LINK [....]=FC INF.
    ''**
    ''*****************************************************************
    'Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
    '    Dim SQL As String
    '    Dim lnk As HyperLink = GridView1.SelectedRow.Cells(4).Controls(0)
    '    Dim xSPNo As String = Replace(lnk.Text, "&nbsp;", "")
    '    xSPNo = Mid(xSPNo, InStr(xSPNo, "SP"), 99)
    '    '
    '    If InStr(xSPNo, "SP") > 0 Then
    '        xSPNo = Mid(xSPNo, InStr(xSPNo, "SP"), 99)
    '    End If
    '    '
    '    If InStr(xSPNo, "[") > 0 Then
    '        xSPNo = Mid(xSPNo, 1, InStr(xSPNo, "[") - 1)
    '    End If
    '    xSPNo = UCase(xSPNo).Trim
    '    '
    '    SQL = "select "
    '    SQL = SQL & "Item, ItemName1+' '+ItemName2 as ItemName, Color, Keep, "
    '    SQL = SQL & "SPNo, "
    '    SQL = SQL & "'http://10.245.0.205/EDI/SPHttp2File.aspx?pUserID=" & UserID & "&pSPNo=' + SPNo As SPNoURL, "
    '    SQL = SQL & "ActionName, "
    '    SQL = SQL & "N_Content, N_Qty, N_Yobi1, "
    '    SQL = SQL & "N1_Content, N1_Qty, N1_Yobi1, N2_Content, N2_Qty, N2_Yobi1, "
    '    SQL = SQL & "N3_Content, N3_Qty, N3_Yobi1, N4_Content, N4_Qty, N4_Yobi1, "
    '    SQL = SQL & "0 AS Balance "
    '    '
    '    SQL = SQL & "from LF_SPActionPlan "
    '    SQL = SQL & "where Version = 99 "
    '    SQL = SQL & "and SPNO = '" & xSPNo & "' "
    '    ' code
    '    If pItem <> "" Then
    '        SQL = SQL & "and item = '" & pItem & "' "
    '    End If
    '    ' keep
    '    If pKeep <> "" Then
    '        SQL = SQL & "and keep = '" & pKeep & "' "
    '    End If
    '    ' Color
    '    If pColor <> "" Then
    '        SQL = SQL & "and ltrim(rtrim(Color)) = '" & pColor & "' "
    '    End If
    '    '
    '    SQL = SQL & "Group by Item, Color, Keep, SPNo, Action, "
    '    SQL = SQL & "Item, ItemName1, ItemName2, Color, Keep, SPNo, ActionName, "
    '    SQL = SQL & "N_Content, N_Qty, N_Yobi1, "
    '    SQL = SQL & "N1_Content, N1_Qty, N1_Yobi1, N2_Content, N2_Qty, N2_Yobi1, "
    '    SQL = SQL & "N3_Content, N3_Qty, N3_Yobi1, N4_Content, N4_Qty, N4_Yobi1  "
    '    '
    '    SQL = SQL & "Order by Item, Color, Keep, SPNo, Action "
    '    '
    '    Dim dt As DataTable = uDataBase.GetDataTable(SQL)
    '    If dt.Rows.Count > 0 Then
    '        GridView1.Visible = True
    '        GridView1.DataSource = dt
    '        GridView1.DataBind()
    '    Else
    '        uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
    '    End If
    '    '
    '    SQL = "select "
    '    SQL = SQL & "Y_J1, FCTNo, C_Code, C_Color, C_SPECIALREQUEST, C_Season, C_SHORTENLT, C_F1 As C_Insert, C_I1 AS C_Buyer "
    '    SQL = SQL & "from LF_SPForcastPlan "
    '    SQL = SQL & "where Version = 99 "
    '    SQL = SQL & "and Y_J1 = '" & xSPNo & "' "
    '    SQL = SQL & "AND Y_ItemCode = '" & pItem & "' "
    '    SQL = SQL & "AND Y_Color = '" & pColor & "' "
    '    SQL = SQL & "AND C_SHORTENLT = '" & pKeep & "' "
    '    '
    '    SQL = SQL & "And LSNo <> '' "
    '    'SQL = SQL & "And CHARINDEX( "
    '    'SQL = SQL & "      case when C_G1='SLIDER ONLY' then 'PS'"
    '    'SQL = SQL & "           when C_G1='CHAIN ONLY' then 'CH' "
    '    'SQL = SQL & "           else 'ZIP' "
    '    'SQL = SQL & "      end, "
    '    'SQL = SQL & "      'ZIP-' + Y_A1 "
    '    'SQL = SQL & "    ) > 0 "
    '    '
    '    SQL = SQL & "Order by Y_J1, FCTNo, C_Code, C_Color "
    '    '
    '    Dim dt_CustInf As DataTable = uDataBase.GetDataTable(SQL)
    '    If dt_CustInf.Rows.Count > 0 Then
    '        GridView2.Visible = True
    '        GridView2.DataSource = dt_CustInf
    '        GridView2.DataBind()
    '        '
    '        AtCloseCust.Style("left") = 150 & "px"
    '        AtCloseCust.Checked = False
    '        '
    '        AtCloseYKK.Style("left") = -500 & "px"
    '        AtCloseYKK.Checked = False
    '        GridView3.Visible = False
    '    Else
    '        uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
    '    End If
    '    '
    'End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**
    ''**     GridView1 LINK [....]=MATERIAL INF.
    ''**
    ''*****************************************************************
    'Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing
    '    Dim SQL As String
    '    Dim lnk As HyperLink = GridView1.SelectedRow.Cells(4).Controls(0)
    '    Dim xSPNo As String = Replace(lnk.Text, "&nbsp;", "")
    '    '
    '    If InStr(xSPNo, "SP") > 0 Then
    '        xSPNo = Mid(xSPNo, InStr(xSPNo, "SP"), 99)
    '    End If
    '    '
    '    If InStr(xSPNo, "[") > 0 Then
    '        xSPNo = Mid(xSPNo, 1, InStr(xSPNo, "[") - 1)
    '    End If
    '    xSPNo = UCase(xSPNo).Trim
    '    '
    '    SQL = "select "
    '    SQL = SQL & "Item, ItemName1+' '+ItemName2 as ItemName, Color, Keep, "
    '    SQL = SQL & "SPNo, "
    '    SQL = SQL & "'http://10.245.0.205/EDI/SPHttp2File.aspx?pUserID=" & UserID & "&pSPNo=' + SPNo As SPNoURL, "
    '    SQL = SQL & "ActionName, "
    '    SQL = SQL & "N_Content, N_Qty, N_Yobi1, "
    '    SQL = SQL & "N1_Content, N1_Qty, N1_Yobi1, N2_Content, N2_Qty, N2_Yobi1, "
    '    SQL = SQL & "N3_Content, N3_Qty, N3_Yobi1, N4_Content, N4_Qty, N4_Yobi1, "
    '    SQL = SQL & "0 AS Balance "
    '    '
    '    SQL = SQL & "from LF_SPActionPlan "


    '    SQL = SQL & "where Version = 99 "
    '    SQL = SQL & "and SPNO = '" & xSPNo & "' "
    '    ' code
    '    If pItem <> "" Then
    '        SQL = SQL & "and item = '" & pItem & "' "
    '    End If
    '    ' keep
    '    If pKeep <> "" Then
    '        SQL = SQL & "and keep = '" & pKeep & "' "
    '    End If
    '    ' Color
    '    If pColor <> "" Then
    '        SQL = SQL & "and ltrim(rtrim(Color)) = '" & pColor & "' "
    '    End If
    '    '
    '    SQL = SQL & "Group by Item, Color, Keep, SPNo, Action, "
    '    SQL = SQL & "Item, ItemName1, ItemName2, Color, Keep, SPNo, ActionName, "
    '    SQL = SQL & "N_Content, N_Qty, N_Yobi1, "
    '    SQL = SQL & "N1_Content, N1_Qty, N1_Yobi1, N2_Content, N2_Qty, N2_Yobi1, "
    '    SQL = SQL & "N3_Content, N3_Qty, N3_Yobi1, N4_Content, N4_Qty, N4_Yobi1  "
    '    '
    '    SQL = SQL & "Order by Item, Color, Keep, SPNo, Action "
    '    '
    '    Dim dt As DataTable = uDataBase.GetDataTable(SQL)
    '    If dt.Rows.Count > 0 Then
    '        GridView1.Visible = True
    '        GridView1.DataSource = dt
    '        GridView1.DataBind()
    '    Else
    '        uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
    '    End If
    '    '
    '    SQL = "select "
    '    SQL = SQL & "[GR_09],[LSNo],[MinimumStock],[N_ScheProd],[N_OnProd],[N_FreeInv],[N_KeepInv],[Description]  "
    '    SQL = SQL & "from LF_SPLocalStockPlan "
    '    SQL = SQL & "where Version = 99 "
    '    SQL = SQL & "and GR_09 = '" & xSPNo & "' "
    '    SQL = SQL & "AND GR_03 = '" & pItem & "' "
    '    SQL = SQL & "AND GR_07 = '" & pColor & "' "
    '    SQL = SQL & "AND GR_02 = '" & pKeep & "' "
    '    SQL = SQL & "Order by GR_09, LSNo, GR_03, GR_07 "
    '    '
    '    Dim dt_LSInf As DataTable = uDataBase.GetDataTable(SQL)
    '    If dt_LSInf.Rows.Count > 0 Then
    '        GridView3.Visible = True
    '        GridView3.DataSource = dt_LSInf
    '        GridView3.DataBind()
    '        '
    '        AtCloseCust.Style("left") = -500 & "px"
    '        AtCloseCust.Checked = False
    '        '
    '        AtCloseYKK.Style("left") = 150 & "px"
    '        AtCloseYKK.Checked = False
    '        GridView2.Visible = False
    '    Else
    '        uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
    '    End If
    '    '
    'End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(Close Windows Customer)
    ''**     
    ''**
    ''*****************************************************************
    'Protected Sub AtCloseCust_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtCloseCust.CheckedChanged
    '    AtCloseCust.Style("left") = -500 & "px"
    '    AtCloseCust.Checked = True
    '    '
    '    GridView2.Visible = False
    '    '
    '    ShowData()
    '    '
    'End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(Close Windows YKK)
    ''**     
    ''**
    ''*****************************************************************
    'Protected Sub AtCloseYKK_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtCloseYKK.CheckedChanged
    '    AtCloseYKK.Style("left") = -500 & "px"
    '    AtCloseYKK.Checked = True
    '    '
    '    GridView3.Visible = False
    '    '
    '    ShowData()
    '    '
    'End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**
    ''**     GridView2 編輯資料 顏色+格式 (1)
    ''**
    ''*****************************************************************
    'Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
    '    Dim i As Integer
    '    '
    '    If (e.Row.RowType = DataControlRowType.Header) Then

    '        Dim gv As GridView = sender
    '        Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
    '        Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
    '        '
    '        ' 清除
    '        e.Row.Cells.Clear()
    '        '
    '        Dim tcl As TableCellCollection = e.Row.Cells
    '        tcl.Clear() '清除自动生成的表头
    '        '
    '        '** 4 line *****************
    '        i = 0
    '        'SPNo
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "SP No."
    '        tcl(i).BackColor = Color.Blue
    '        i = i + 1
    '        'FC No
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "FC No."
    '        tcl(i).BackColor = Color.Blue
    '        i = i + 1
    '        'Code
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "Code"
    '        tcl(i).BackColor = Color.Purple
    '        i = i + 1
    '        'Color
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "Color"
    '        tcl(i).BackColor = Color.Purple
    '        i = i + 1
    '        'Special
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "Special"
    '        tcl(i).BackColor = Color.Purple
    '        i = i + 1
    '        'L/R INSERT
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "L/R Ins."
    '        tcl(i).BackColor = Color.Purple
    '        i = i + 1
    '        'Season
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "Season"
    '        tcl(i).ColumnSpan = 1
    '        tcl(i).BackColor = Color.Green
    '        i = i + 1
    '        '[KeepCode]
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "Keep"
    '        tcl(i).BackColor = Color.Green
    '        i = i + 1
    '        '[Buyer]
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "Buyer"
    '        tcl(i).BackColor = Color.Green
    '        i = i + 1
    '        '
    '        gv.Controls(0).Controls.AddAt(0, H3row)
    '    End If
    '    '
    'End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**
    ''**     GridView3 編輯資料 顏色+格式 (1)
    ''**
    ''*****************************************************************
    'Protected Sub GridView3_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView3.RowDataBound
    '    Dim i As Integer
    '    '
    '    If (e.Row.RowType = DataControlRowType.Header) Then

    '        Dim gv As GridView = sender
    '        Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
    '        Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
    '        '
    '        ' 清除
    '        e.Row.Cells.Clear()
    '        '
    '        Dim tcl As TableCellCollection = e.Row.Cells
    '        tcl.Clear() '清除自动生成的表头
    '        '
    '        '** 4 line *****************
    '        i = 0
    '        'SPNo
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "SP No."
    '        tcl(i).BackColor = Color.Blue
    '        i = i + 1
    '        'LS No
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "LS No."
    '        tcl(i).BackColor = Color.Blue
    '        i = i + 1
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "Mini. Stock"
    '        tcl(i).BackColor = Color.Purple
    '        i = i + 1
    '        'Sche P.
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "Sche P."
    '        tcl(i).BackColor = Color.Purple
    '        i = i + 1
    '        'On P.
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "On P."
    '        tcl(i).BackColor = Color.Purple
    '        i = i + 1
    '        'Free
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "Free"
    '        tcl(i).ColumnSpan = 1
    '        tcl(i).BackColor = Color.Purple
    '        i = i + 1
    '        'Keep
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "Keep"
    '        tcl(i).BackColor = Color.Purple
    '        i = i + 1
    '        'Descr
    '        tcl.Add(New TableHeaderCell())
    '        tcl(i).Text = "Description."
    '        tcl(i).BackColor = Color.Green
    '        i = i + 1
    '        '
    '        gv.Controls(0).Controls.AddAt(0, H3row)
    '    End If
    '    '
    'End Sub
    '
End Class
