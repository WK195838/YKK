Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class PS_INQ_ListPrice03
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oWaves As New Waves.CommonService

    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim pBuyer As String             'Buyer
    Dim UserID As String            'UserID
    Dim EDLConnect As String = uCommon.GetAppSetting("EDLDB")
    Dim ConnectString As String = uCommon.GetAppSetting("WAVESDB")
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
        Response.Cookies("PGM").Value = "PS_INQ_ListPrice01.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        pBuyer = Request.QueryString("pBuyer")
        UserID = Request.QueryString("pUserID")             'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        DBUYER.ReadOnly = True
        HBuyerCode.Style("left") = -500 & "px"
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        GridView2.Visible = False
        GridView3.Visible = False
        GridView4.Visible = False
        GridView5.Visible = False
        DBUYER.ReadOnly = True
        LItemKey1.ReadOnly = True
        LItemKey2.ReadOnly = True
        LItemKey3.ReadOnly = True
        LItemKey4.ReadOnly = True
        '
        HBuyerCode.Style("left") = -500 & "px"
        BExcel.Style("top") = -500 & "px"
        BSPInq.Style("left") = -500 & "px"
        '
        DINQLISTPRICEFile.Style("left") = -500 & "px"
        DINQLISTPRICEFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "INQLISTPRICE_" & UserID & ".xlsm"
        '
        '動作按鈕設定
        BSPInq.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('INQBUYERITEMExcel')}"
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        DBUYER.Text = ""
        HBuyerCode.Text = ""
        DKEY1.Text = ""
        DKEY2.Text = ""
        '
        If pBuyer <> "" Then
            If pBuyer = "TW1741" Then
                DBUYER.Text = "LULULEMON"
                Label6.Text = "L888(USD)(Exchange Rate=31.5)"
            End If
            If pBuyer = "000151" Then
                DBUYER.Text = "PUMA"
                Label6.Text = "A211(USD)(Exchange Rate=31.5)"
            End If
            '
            If pBuyer = "TW1741" Or pBuyer = "000151" Then
                HBuyerCode.Text = "ISOS-" & pBuyer
            Else
                HBuyerCode.Text = "FALL-" & pBuyer
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '事件-GRIDVIEW EVENT@
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GridView1=Z#, GridView2=P#,GridView3=F#,)
    '**     自創ITEM關鍵字
    '**
    '*****************************************************************
    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Dim xCat As String = Trim(GridView1.SelectedRow.Cells(1).Text.ToUpper)
        Dim xItem As String = Trim(GridView1.SelectedRow.Cells(3).Text.ToUpper)
        Dim xLeadTime As String = Trim(GridView1.SelectedRow.Cells(4).Text.ToUpper)
        Dim RtnCode As String
        '
        RtnCode = GetItemKey(xCat, xItem, xLeadTime)
        If RtnCode <> "1" Then
            DPKEY.Text = RtnCode
            DFKEY.Text = ""
            DKEY1.Text = ""
            DKEY2.Text = ""
            ShowData()
            BGetItem_Click(sender, e)
        End If
    End Sub

    Protected Sub GridView2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView2.SelectedIndexChanged
        Dim xCat As String = Trim(GridView2.SelectedRow.Cells(1).Text.ToUpper)
        Dim xItem As String = Trim(GridView2.SelectedRow.Cells(3).Text.ToUpper)
        Dim xLeadTime As String = Trim(GridView2.SelectedRow.Cells(4).Text.ToUpper)
        Dim RtnCode As String
        '
        RtnCode = GetItemKey(xCat, xItem, xLeadTime)
        If RtnCode <> "1" Then
            'DPKEY.Text = "Z#" & RtnCode
            ShowData()
            BGetItem_Click(sender, e)
        End If
    End Sub

    Protected Sub GridView3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView3.SelectedIndexChanged
        Dim xCat As String = "F#"
        Dim xItem As String = Trim(GridView3.SelectedRow.Cells(0).Text.ToUpper)
        Dim xLeadTime As String = "0"
        Dim RtnCode As String
        '
        If pBuyer = "TW1741" Or pBuyer = "000151" Then
            RtnCode = GetItemKey(xCat, xItem, xLeadTime)
            If RtnCode <> "1" Then
                ShowData()
                BGetItem_Click(sender, e)
            End If
        End If
    End Sub

    Protected Sub GridView4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView4.SelectedIndexChanged
        Dim xCat As String = "T#"
        Dim xItem As String = Trim(GridView4.SelectedRow.Cells(0).Text.ToUpper)
        Dim xLeadTime As String = "0"
        Dim RtnCode As String
        '
        RtnCode = GetItemKey(xCat, xItem, xLeadTime)
        If RtnCode <> "1" Then
            ShowData()
            BGetItem_Click(sender, e)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '事件-BUTTON
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Inq)
    '**     查詢
    '**
    '*****************************************************************
    Protected Sub BInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BInq.Click
        ShowData()
    End Sub
    '*****************************************************************
    '**(BZKEY, BPKEY)
    '**     查詢Z#, P#
    '**
    '*****************************************************************
    Protected Sub BZKEY_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BZKEY.Click
        If pBuyer = "TW1741" Or pBuyer = "000151" Then
            ShowData()
        End If
    End Sub
    Protected Sub BPKEY_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BPKEY.Click
        If pBuyer = "TW1741" Or pBuyer = "000151" Then
            ShowData()
        End If
    End Sub
    '*****************************************************************
    '**(GetItem)
    '**     取得WINGS ITEM
    '**
    '*****************************************************************
    Protected Sub BGetItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGetItem.Click
        On Error GoTo LBL_Error
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim dc As New OleDbCommand
        Dim sql As String
        Dim i As Integer
        Dim xKeyList As Object
        '
        cn.ConnectionString = ConnectString
        '
        ShowData()
        '
        '--SHOW ITEM INF.
        '=========================================================================
        'LULULEMON
        '=========================================================================
        If pBuyer = "TW1741" Then
            '
            sql = "SELECT Top 100 "
            sql = sql + "L888 As PriceInf, BaseLength, InchPrice, ListPrice, "
            sql = sql + "'['+LTRIM(RTRIM(ITEM))+']'+LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) As ItemInf, "
            sql = sql + "LTRIM(RTRIM(ITEM)) As Code, "
            sql = sql + "LTRIM(RTRIM(ITEM_NAME1)) As Name1, "
            sql = sql + "LTRIM(RTRIM(ITEM_NAME2)) As Name2, "
            sql = sql + "LTRIM(RTRIM(ITEM_NAME3)) As Name3 "
            sql = sql + "From V_LLMPRICE "
            '
            sql = sql + "Where LTRIM(RTRIM(ITEM)) <> '" & "@@@@@@@" & "' "
            '
            If InStr(DItemKey1.Text, "/") > 0 Then
                xKeyList = Split(DItemKey1.Text.ToUpper, "/")
                For i = LBound(xKeyList) To UBound(xKeyList)
                    If xKeyList(i) <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & xKeyList(i) & "%' "
                Next
            Else
                If DItemKey1.Text <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DItemKey1.Text & "%' "
            End If
            '
            If InStr(DItemKey2.Text, "/") > 0 Then
                xKeyList = Split(DItemKey2.Text.ToUpper, "/")
                For i = LBound(xKeyList) To UBound(xKeyList)
                    If xKeyList(i) <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & xKeyList(i) & "%' "
                Next
            Else
                If DItemKey2.Text <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DItemKey2.Text & "%' "
            End If
            '
            If InStr(DItemKey3.Text, "/") > 0 Then
                xKeyList = Split(DItemKey3.Text.ToUpper, "/")
                For i = LBound(xKeyList) To UBound(xKeyList)
                    If xKeyList(i) <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & xKeyList(i) & "%' "
                Next
            Else
                If DItemKey3.Text <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DItemKey3.Text.ToUpper & "%' "
            End If
            '
            If InStr(DItemKey4.Text, "/") > 0 Then
                xKeyList = Split(DItemKey4.Text.ToUpper, "/")
                For i = LBound(xKeyList) To UBound(xKeyList)
                    If xKeyList(i) <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & xKeyList(i) & "%' "
                Next
            Else
                If DItemKey4.Text <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DItemKey4.Text.ToUpper & "%' "
            End If
            '
            If InStr(DItemKey5.Text, "/") > 0 Then
                xKeyList = Split(DItemKey5.Text.ToUpper, "/")
                For i = LBound(xKeyList) To UBound(xKeyList)
                    If xKeyList(i) <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & xKeyList(i) & "%' "
                Next
            Else
                If DItemKey5.Text <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DItemKey5.Text.ToUpper & "%' "
            End If
            '
            sql = sql + "And NoDisp <> '1' "
            sql = sql + "And L888 = 'L888(TWD)' "
            sql = sql + "ORDER BY ITEM DESC, ITEM_NAME1, ITEM_NAME2, ITEM_NAME3 "
        End If
        '=========================================================================
        'PUMA
        '=========================================================================
        If pBuyer = "000151" Then
            '
            sql = "SELECT Top 100 "
            sql = sql + "A211 As PriceInf, BaseLength, InchPrice, ListPrice, "
            sql = sql + "'['+LTRIM(RTRIM(ITEM))+']'+LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) As ItemInf, "
            sql = sql + "LTRIM(RTRIM(ITEM)) As Code, "
            sql = sql + "LTRIM(RTRIM(ITEM_NAME1)) As Name1, "
            sql = sql + "LTRIM(RTRIM(ITEM_NAME2)) As Name2, "
            sql = sql + "LTRIM(RTRIM(ITEM_NAME3)) As Name3 "
            sql = sql + "From V_PUMAPRICE "
            '
            sql = sql + "Where LTRIM(RTRIM(ITEM)) <> '" & "@@@@@@@" & "' "
            '
            If InStr(DItemKey1.Text, "/") > 0 Then
                xKeyList = Split(DItemKey1.Text.ToUpper, "/")
                For i = LBound(xKeyList) To UBound(xKeyList)
                    If xKeyList(i) <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & xKeyList(i) & "%' "
                Next
            Else
                If DItemKey1.Text <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DItemKey1.Text & "%' "
            End If
            '
            If InStr(DItemKey2.Text, "/") > 0 Then
                xKeyList = Split(DItemKey2.Text.ToUpper, "/")
                For i = LBound(xKeyList) To UBound(xKeyList)
                    If xKeyList(i) <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & xKeyList(i) & "%' "
                Next
            Else
                If DItemKey2.Text <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DItemKey2.Text & "%' "
            End If
            '
            If InStr(DItemKey3.Text, "/") > 0 Then
                xKeyList = Split(DItemKey3.Text.ToUpper, "/")
                For i = LBound(xKeyList) To UBound(xKeyList)
                    If xKeyList(i) <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & xKeyList(i) & "%' "
                Next
            Else
                If DItemKey3.Text <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DItemKey3.Text.ToUpper & "%' "
            End If
            '
            If InStr(DItemKey4.Text, "/") > 0 Then
                xKeyList = Split(DItemKey4.Text.ToUpper, "/")
                For i = LBound(xKeyList) To UBound(xKeyList)
                    If xKeyList(i) <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & xKeyList(i) & "%' "
                Next
            Else
                If DItemKey4.Text <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DItemKey4.Text.ToUpper & "%' "
            End If
            '
            If InStr(DItemKey5.Text, "/") > 0 Then
                xKeyList = Split(DItemKey5.Text.ToUpper, "/")
                For i = LBound(xKeyList) To UBound(xKeyList)
                    If xKeyList(i) <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & xKeyList(i) & "%' "
                Next
            Else
                If DItemKey5.Text <> "" Then sql = sql + "And LTRIM(RTRIM(ITEM)) + ' ' + LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DItemKey5.Text.ToUpper & "%' "
            End If
            '
            sql = sql + "And NoDisp <> '1' "
            sql = sql + "And A211 = 'A211(TWD)' "
            sql = sql + "ORDER BY ITEM DESC, ITEM_NAME1, ITEM_NAME2, ITEM_NAME3 "
        End If
        '
        'MsgBox(sql)
        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
        DBAdapter1.Fill(ds, "FA000")
        If ds.Tables("FA000").Rows.Count > 0 Then
            GridView5.Visible = True
            GridView5.DataSource = ds
            GridView5.DataBind()
            'Else
            '    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If
        '
        cn.Close()
        '
        GoTo LBL_End
        '##
LBL_Error:
        On Error GoTo -1
        If cn.State = ConnectionState.Open Then cn.Close()
        uJavaScript.PopMsg(Me, "指定條件搜尋不到資料,請確認!")
        '##
LBL_End:
    End Sub
    '*****************************************************************
    '**(Reset)
    '**     Reset
    '**
    '*****************************************************************
    Protected Sub BReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BReset.Click
        DItemKey1.Text = ""
        DItemKey2.Text = ""
        DItemKey3.Text = ""
        DItemKey4.Text = ""
        DItemKey5.Text = ""
        DLeadTime.Text = ""
        DZKEY.Text = ""
        DPKEY.Text = ""
        '
        ShowData()
    End Sub

    '---------------------------------------------------------------------------------------------------
    'GRIDVIEW 表題 & 欄位
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GRIDVIEW1 編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        '=========================================================================
        'LLM
        '=========================================================================
        If pBuyer = "TW1741" Or pBuyer = "000151" Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "D1" & "<BR>" & "Code"
                tcl(0).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "E1" & "<BR>" & "L/T"
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                'e.Row.Cells(4).ForeColor = Color.Red
                For i = 0 To 27
                    e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                Next
                '
                For i = 0 To 2
                    e.Row.Cells(i).Visible = False
                Next
                For i = 5 To 27
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GRIDVIEW2 編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        Dim i As Integer
        '=========================================================================
        'LLM
        '=========================================================================
        If pBuyer = "TW1741" Or pBuyer = "000151" Then
            '
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "D1" & "<BR>" & "Code"
                tcl(0).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "E1" & "<BR>" & "L/T"
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                'e.Row.Cells(4).ForeColor = Color.Red
                For i = 0 To 27
                    e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                Next
                '
                For i = 0 To 2
                    e.Row.Cells(i).Visible = False
                Next
                For i = 5 To 27
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GRIDVIEW3 編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView3_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView3.RowDataBound
        '=========================================================================
        'LLM
        '=========================================================================
        If pBuyer = "TW1741" Or pBuyer = "000151" Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "D1" & "<BR>" & "Code"
                tcl(0).BackColor = Color.Blue
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                e.Row.Cells(0).ForeColor = Color.Red
            End If
        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GRIDVIEW4 編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView4_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView4.RowDataBound
        '=========================================================================
        'LLM
        '=========================================================================
        If pBuyer = "TW1741" Or pBuyer = "000151" Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "D1" & "<BR>" & "Code"
                tcl(0).BackColor = Color.Blue
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                'e.Row.Cells(0).ForeColor = Color.Red
            End If
        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     gridview5 編輯資料
    '**
    '*****************************************************************
    Protected Sub gridview5_RowDataBound1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView5.RowDataBound
        Dim i As Integer
        '
        '=========================================================================
        'LLM
        '=========================================================================
        If pBuyer = "TW1741" Or pBuyer = "000151" Then
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                'tcl.Add(New TableHeaderCell())
                'tcl(0).Text = ""
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "Base" & "<BR>" & "(L)"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "List" & "<BR>" & "(P)"
                tcl(1).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "add 1''" & "<BR>" & "(P)"
                tcl(2).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "Item Inf."
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                e.Row.Cells(2).Text = CStr(CInt(CDbl(e.Row.Cells(2).Text) * 100) / 100)
                e.Row.Cells(3).Text = CStr(CInt(CDbl(e.Row.Cells(3).Text) * 100) / 100)
                e.Row.Cells(4).ForeColor = Color.Red
                e.Row.Cells(0).Visible = False
                For i = 5 To 8
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub

    '---------------------------------------------------------------------------------------------------
    '副PGM
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        On Error GoTo LBL_Error
        Dim cn As New OleDbConnection
        Dim ds, ds1, ds2, ds3 As New DataSet
        Dim dc As New OleDbCommand
        Dim sql As String
        Dim i As Integer
        Dim xKeyList As Object
        '
        cn.ConnectionString = EDLConnect
        '
        '篩選資料
        '=========================================================================
        'LULULEMON
        '=========================================================================
        If pBuyer = "TW1741" Or pBuyer = "000151" Then
            '
            '
            'Z# --
            sql = "SELECT TOP 100 "
            sql &= "'Ｓ' as Mark, "
            sql &= "'' as URL, "
            sql &= "A1, B1, C1, D1, E1, "
            sql &= "F1, G1, H1, I1, J1, "
            sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
            sql &= "U1, V1, W1, X1, Y1, Z1, "
            '
            sql &= "'' as FormMark, "
            sql &= "'' AS FormURL "
            '
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "BUYERITEM-B" & "') "
            sql &= "And Active = '0' "
            sql &= "And B1 IN ('Z#') "
            '
            If DFKEY.Text <> "" Then
                '使用A1~E1, O1~P1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            '
            If InStr(DZKEY.Text, "/") > 0 Then
                xKeyList = Split(DZKEY.Text.ToUpper, "/")
                For i = LBound(xKeyList) To UBound(xKeyList)
                    If xKeyList(i) <> "" Then sql = sql + "And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & xKeyList(i) & "%' "
                Next
            Else
                If DZKEY.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DZKEY.Text & "%' "
            End If
            '
            sql &= " Group by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
            sql &= " Order by E1 DESC,N1 "
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "item")
            If ds.Tables("item").Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = ds
                GridView1.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
            '
            'P# --
            sql = "SELECT TOP 100 "
            sql &= "'Ｓ' as Mark, "
            sql &= "'' as URL, "
            sql &= "A1, B1, C1, D1, E1, "
            sql &= "F1, G1, H1, I1, J1, "
            sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
            sql &= "U1, V1, W1, X1, Y1, Z1, "
            '
            sql &= "'' as FormMark, "
            sql &= "'' AS FormURL "
            '
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "BUYERITEM-B" & "') "
            sql &= "And Active = '0' "
            sql &= "And B1 IN ('P#') "
            '
            If DFKEY.Text <> "" Then
                '使用A1~E1, O1~P1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            '
            If InStr(DPKEY.Text, "/") > 0 Then
                xKeyList = Split(DPKEY.Text.ToUpper, "/")
                For i = LBound(xKeyList) To UBound(xKeyList)
                    If xKeyList(i) <> "" Then sql = sql + "And Replace(A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1,' ','') Like '%" & xKeyList(i) & "%' "
                Next
            Else
                If DPKEY.Text <> "" Then sql &= " And Replace(A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1,' ','') Like '%" & DPKEY.Text & "%' "
            End If
            '
            sql &= " Group by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
            sql &= " Order by E1 DESC,N1 "
            Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
            DBAdapter2.Fill(ds1, "item")
            If ds.Tables("item").Rows.Count > 0 Then
                GridView2.Visible = True
                GridView2.DataSource = ds1
                GridView2.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
            '
            'FINISH# --
            sql = "SELECT TOP 100 "
            sql &= "Finish AS C1 "
            sql &= "From V_OPReportData_DigitalFinish "

            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "AND  Finish <> '" & "" & "' "
            '
            sql &= " Group by Finish "
            sql &= " Order by Finish "
            Dim DBAdapter3 As New OleDbDataAdapter(sql, cn)
            DBAdapter3.Fill(ds2, "item")
            If ds.Tables("item").Rows.Count > 0 Then
                GridView3.Visible = True
                GridView3.DataSource = ds2
                GridView3.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
            '
            'TAPE# --
            sql = "SELECT TOP 100 "
            sql &= "ItemTape AS C1 "
            sql &= "From V_OPReportData_DigitalTape "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "AND ItemTape <> '" & "" & "' "
            '
            sql &= " Group by ItemTape "
            sql &= " Order by ItemTape "
            Dim DBAdapter4 As New OleDbDataAdapter(sql, cn)
            DBAdapter4.Fill(ds3, "item")
            If ds.Tables("item").Rows.Count > 0 Then
                GridView4.Visible = True
                GridView4.DataSource = ds3
                GridView4.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
        End If
        '
        GoTo LBL_End
        '##
LBL_Error:
        On Error GoTo -1
        If cn.State = ConnectionState.Open Then cn.Close()
        uJavaScript.PopMsg(Me, "指定條件搜尋不到資料,請確認!")
        '##
LBL_End:
    End Sub
    '*****************************************************************
    '**
    '**     分解Item Key
    '**
    '*****************************************************************
    Function GetItemKey(ByVal pCat As String, ByVal pItem As String, ByVal pLT As String) As String
        Dim RtnString As String = "1"
        '
        Dim xReplaceStr, xNewStr As String
        Dim xReplaceList, xNewList, xKeyList As Object
        Dim str, xSize, xChainCode, xOther1, xOther2, xOther3 As String
        Dim i As Integer
        '
        'KEYWORD 整理
        str = UCase(Trim(pItem))
        '
        'BUYER別 Z#(TAPE)/P#(SLIDER)
        'LULULEMON
        If pBuyer = "TW1741" Then
            '
            '--TAPE----------------------------------------------
            If pCat = "Z#" Then
                '
                'L/T
                If Not IsNumeric(DLeadTime.Text) Then
                    LLeadTime.Text = "L/T[Z#]"
                    DLeadTime.Text = "0"
                End If
                If IsNumeric(pLT) Then
                    If LLeadTime.Text = "L/T[Z#]" Then
                        LLeadTime.Text = "L/T[Z#]"
                        DLeadTime.Text = pLT
                    Else
                        If CInt(pLT) > CInt(DLeadTime.Text) Then
                            LLeadTime.Text = "L/T[Z#]"
                            DLeadTime.Text = pLT
                        End If
                    End If
                End If
                '---
                '整理KEYWORD
                '[-],[*]以後字串不要
                'If InStr(str, "-") > 0 Then str = Mid(str, 1, InStr(str, "-") - 1)
                'If InStr(str, "*") > 0 Then str = Mid(str, 1, InStr(str, "*") - 1)
                str = Trim(str)
                '
                '置換字串
                '             "1/2/3/4/5/6/7/8/9/10/"
                xReplaceStr = "*REV|*| |-|||||||"
                xNewStr = "_REVERSE|_|_|_|||||||"
                xReplaceList = Split(xReplaceStr, "|")
                xNewList = Split(xNewStr, "|")
                For i = LBound(xReplaceList) To UBound(xReplaceList)
                    str = Replace(str, xReplaceList(i), xNewList(i))
                Next
                '---
                '截取有效字串
                '
                xKeyList = Split(str, "_")
                For i = LBound(xKeyList) To UBound(xKeyList)
                    'MsgBox(xKeyList(i))
                    'Size
                    If i = 0 Then
                        xSize = xKeyList(i)
                        xSize = Mid(xSize, InStr(xSize, "Z#") + 2)
                        xSize = Mid(xSize, 1, InStr(xSize, "YKK") - 1)
                        xSize = Trim(xSize)
                    End If
                    'ChainCode
                    If i = 1 Then
                        xChainCode = xKeyList(i)
                        xChainCode = Trim(xChainCode)
                    End If
                    'Other1
                    If i = 2 Then
                        If IsNumeric(xKeyList(i)) Then
                            xOther1 = ""
                        Else
                            xOther1 = xKeyList(i)
                        End If
                        xOther1 = Trim(xOther1)
                    End If
                    'Other2
                    If i = 3 Then
                        xOther2 = xKeyList(i)
                        xOther2 = Trim(xOther2)
                    End If
                    'Other3
                    If i = 4 Then
                        xOther3 = xKeyList(i)
                        xOther3 = Trim(xOther3)
                    End If
                Next
                '
                If InStr(str, "CHAIN") > 0 Then
                    If CInt(xSize) < 10 Then
                        DItemKey1.Text = "0" & xSize & "/" & xChainCode & "/" & xOther1 & "/" & xOther2 & "/" & xOther3 & "/"
                    Else
                        DItemKey1.Text = xSize & "/" & xChainCode & "/" & xOther1 & "/" & xOther2 & "/" & xOther3 & "/"
                    End If
                Else
                    DItemKey1.Text = xChainCode & "-" & xSize & "/" & xOther1 & "/" & xOther2 & "/" & xOther3 & "/"
                End If
                '
                If DItemKey1.Text <> "" Then
                    RtnString = xSize & xChainCode
                    If InStr(RtnString, "5Y") > 0 Then RtnString = xSize & "M"
                    If InStr(RtnString, "5RC") > 0 Or InStr(RtnString, "8RC") Then
                        RtnString = xSize & "RC"
                    Else
                        RtnString = xSize & Replace(Mid(xChainCode, 1, 1), "R", "M")
                    End If
                End If
                '
            End If
            '--TAPE-END---------------------------------------------
            '
            '--SLIDER----------------------------------------------
            If pCat = "P#" Then
                '
                'L/T
                If Not IsNumeric(DLeadTime.Text) Then
                    LLeadTime.Text = "L/T[P#]"
                    DLeadTime.Text = "0"
                End If
                If IsNumeric(pLT) Then
                    If LLeadTime.Text = "L/T[P#]" Then
                        LLeadTime.Text = "L/T[P#]"
                        DLeadTime.Text = pLT
                    Else
                        If CInt(pLT) > CInt(DLeadTime.Text) Then
                            LLeadTime.Text = "L/T[P#]"
                            DLeadTime.Text = pLT
                        End If
                    End If
                End If
                '---
                '整理KEYWORD
                '
                '刪除無效字串
                '無
                'P#3YKK_03-C DAB2Y634
                '
                '置換字串
                '             "1/2/3/4/5/6/7/8/9/10/"
                xReplaceStr = " C|-C| |-|/||||||"
                xNewStr = "+C|+C|_|_|_||||||"
                xReplaceList = Split(xReplaceStr, "|")
                xNewList = Split(xNewStr, "|")
                For i = LBound(xReplaceList) To UBound(xReplaceList)
                    str = Replace(str, xReplaceList(i), xNewList(i))
                Next
                'MsgBox(str)
                '---
                '截取有效字串
                xKeyList = Split(str, "_")
                For i = LBound(xKeyList) To UBound(xKeyList)
                    'MsgBox(xKeyList(i))
                    'Size
                    If i = 0 Then
                        xSize = xKeyList(i)
                        xSize = Mid(xSize, InStr(xSize, "P#") + 2)
                        xSize = Mid(xSize, 1, InStr(xSize, "YKK") - 1)
                        xSize = Trim(xSize)
                    End If
                    'Other1
                    If i = 2 Then
                        xOther1 = xKeyList(i)
                        xOther1 = Trim(xOther1)
                    End If
                    If i = 3 Then
                        xOther2 = xKeyList(i)
                        xOther2 = Trim(xOther2)
                    End If
                    If i = 4 Then
                        xOther3 = xKeyList(i)
                        xOther3 = Trim(xOther3)
                    End If
                Next
                '
                DItemKey2.Text = xOther1 & "/" & xOther2 & "/" & xOther3 & "/"
                If DItemKey1.Text <> "" Then DItemKey2.Text = Replace(DItemKey2.Text, "DAB2", "DAB")
                '
                If DItemKey2.Text <> "" Then RtnString = xSize
                '
            End If
        End If
        '
        'PUMA
        If pBuyer = "000151" Then
            '
            '--TAPE----------------------------------------------
            If pCat = "Z#" Then
                '
                'L/T
                If Not IsNumeric(DLeadTime.Text) Then
                    LLeadTime.Text = "L/T[Z#]"
                    DLeadTime.Text = "0"
                End If
                If IsNumeric(pLT) Then
                    If LLeadTime.Text = "L/T[Z#]" Then
                        LLeadTime.Text = "L/T[Z#]"
                        DLeadTime.Text = pLT
                    Else
                        If CInt(pLT) > CInt(DLeadTime.Text) Then
                            LLeadTime.Text = "L/T[Z#]"
                            DLeadTime.Text = pLT
                        End If
                    End If
                End If
                '
                '置換字串
                '             "1/2/3/4/5/6/7/8/9/10/"
                str = Trim(str)
                xReplaceStr = "*REV|*| |-|||||||"
                xNewStr = "_REVERSE|_|_|_|||||||"
                xReplaceList = Split(xReplaceStr, "|")
                xNewList = Split(xNewStr, "|")
                For i = LBound(xReplaceList) To UBound(xReplaceList)
                    str = Replace(str, xReplaceList(i), xNewList(i))
                Next
                'MsgBox(str)
                '
                '截取有效字串
                '
                xKeyList = Split(str, "_")
                For i = LBound(xKeyList) To UBound(xKeyList)
                    'MsgBox(xKeyList(i))
                    'I=0 CUST-ITEM 無使用
                    'ChainCode
                    If i = 1 Then
                        xChainCode = xKeyList(i)
                        xChainCode = Trim(xChainCode)
                    End If
                    'Size
                    If i = 2 Then
                        xSize = xKeyList(i)
                        xSize = Trim(xSize)
                    End If
                    'Other1 
                    If i = 3 Then
                        xOther1 = xKeyList(i)
                        xOther1 = Trim(xOther1)
                    End If
                    'Other2
                    If i = 4 Then
                        xOther2 = xKeyList(i)
                        xOther2 = Trim(xOther2)
                    End If
                    'Other3
                    If i = 5 Then
                        xOther3 = xKeyList(i)
                        xOther3 = Trim(xOther3)
                    End If
                Next
                DItemKey1.Text = xChainCode & "/" & xSize & "/" & xOther1 & "/" & xOther2 & "/" & xOther3 & "/"
                '
                If DItemKey1.Text <> "" Then
                    RtnString = xSize & Mid(xChainCode, 1, 1)
                    If CInt(xSize) < 10 Then
                        RtnString = "0" & xSize & Mid(xChainCode, 1, 1)
                    End If
                End If
            End If
            '--TAPE-END---------------------------------------------
            '
            '--SLIDER----------------------------------------------
            If pCat = "P#" Then
                '
                'L/T
                If Not IsNumeric(DLeadTime.Text) Then
                    LLeadTime.Text = "L/T[P#]"
                    DLeadTime.Text = "0"
                End If
                If IsNumeric(pLT) Then
                    If LLeadTime.Text = "L/T[P#]" Then
                        LLeadTime.Text = "L/T[P#]"
                        DLeadTime.Text = pLT
                    Else
                        If CInt(pLT) > CInt(DLeadTime.Text) Then
                            LLeadTime.Text = "L/T[P#]"
                            DLeadTime.Text = pLT
                        End If
                    End If
                End If
                '---
                '整理KEYWORD
                '
                '刪除無效字串
                '無
                'P#3YKK_03-C DAB2Y634
                '
                '置換字串
                '             "1/2/3/4/5/6/7/8/9/10/"
                xReplaceStr = " C|-C| |-|/||||||"
                xNewStr = "+C|+C|_|_|_||||||"
                xReplaceList = Split(xReplaceStr, "|")
                xNewList = Split(xNewStr, "|")
                For i = LBound(xReplaceList) To UBound(xReplaceList)
                    str = Replace(str, xReplaceList(i), xNewList(i))
                Next
                'MsgBox(str)
                '---
                '截取有效字串
                xKeyList = Split(str, "_")
                For i = LBound(xKeyList) To UBound(xKeyList)
                    'MsgBox(xKeyList(i))
                    'Size
                    If i = 0 And i = UBound(xKeyList) Then
                        xSize = xKeyList(i)
                        xSize = Trim(xSize)
                    End If
                    'ChainCode
                    If i = 1 And i = UBound(xKeyList) Then
                        xChainCode = xKeyList(i)
                        xChainCode = Trim(xChainCode)
                    End If
                    'Other1
                    If i = 2 And i = UBound(xKeyList) Then
                        xOther1 = xKeyList(i)
                        xOther1 = Trim(xOther1)
                    End If
                    'Other2
                    If i = 3 And i = UBound(xKeyList) Then
                        xOther2 = xKeyList(i)
                        xOther2 = Trim(xOther2)
                    End If
                    'Other3
                    If i = 4 And i = UBound(xKeyList) Then
                        xOther3 = xKeyList(i)
                        xOther3 = Trim(xOther3)
                    End If
                Next
                DItemKey2.Text = xSize & "/" & xChainCode & "/" & xOther1 & "/" & xOther2 & "/" & xOther3 & "/"
                '
                If DItemKey2.Text <> "" Then RtnString = xSize
            End If
            '
        End If
        '
        'BUYER共通
        If pCat = "F#" Then
            str = UCase(Trim(pItem))
            DItemKey3.Text = str
            '
            If DItemKey3.Text <> "" Then RtnString = "0"
            '
        End If
        '
        If pCat = "T#" Then
            str = UCase(Trim(pItem))
            DItemKey4.Text = str
            '
            If DItemKey4.Text <> "" Then RtnString = "0"
            '
        End If
        '
        Return RtnString
    End Function
    '*****************************************************************
    '**
    '**     Get Search Field
    '**
    '*****************************************************************
    Public Function GetSearchField(ByVal pCmd As String, ByVal pFieldStr As String, ByVal pDataStr As String) As String
        Dim RtnString As String = ""
        Dim str As String
        Dim i As Integer
        Dim FieldNames(), DataNames() As String
        '
        Try
            str = UCase(pCmd)
            FieldNames = pFieldStr.Split("/")
            DataNames = pDataStr.Split("/")
            For i = 0 To DataNames.Length - 1
                If FieldNames(i) <> DataNames(i) Then
                    'MsgBox(FieldNames(i) & ";" & DataNames(i))
                    'MsgBox("B=" & str)
                    str = str.Replace(FieldNames(i), DataNames(i))
                    'MsgBox("A=" & str)
                    If str <> UCase(pCmd) Then Exit For
                End If
            Next
            RtnString = str
        Catch ex As Exception
        End Try
        '
        Return RtnString
    End Function
    '
    '*****************************************************************
    '**
    '**     轉Excel
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        '
        Response.Clear()
        Response.Buffer = True

        Response.AppendHeader("Content-Disposition", "attachment;filename=ItemList.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        '
        ShowData()
        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()
    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub
    '
    '---------------------------------------------------------------------------------------------------


End Class
