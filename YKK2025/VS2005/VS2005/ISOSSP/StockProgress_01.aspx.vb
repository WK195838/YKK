Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.IO

Partial Class StockProgress_01
    Inherits System.Web.UI.Page
    '
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
    Dim oWaves As New WAVES.CommonService

    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim UserID As String            'UserID

    Dim xBalance, iSum, oSum As Double
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
        Response.Cookies("PGM").Value = "StockProgress_01.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        UserID = Request.QueryString("pUserID")             'UserID
        AtCloseHistory.Style("left") = -500 & "px"
        AtCloseHistory.Checked = False
        GridView2.Visible = False
        '
        DKSource.Style("left") = -500 & "px"
        DKPuller.Style("left") = -500 & "px"

        LKPTool.Style("left") = -500 & "px"

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
        If InStr("MK045/IT003/" & "MK043/MK023/", UCase(UserID)) > 0 Then
            '
            LKPTool.Style("left") = 770 & "px"
            LKPTool.NavigateUrl = "http://10.245.0.205/EDI/SPHttp2File.aspx?pFun=KP&pUserID=" & UCase(UserID) & ""
            '
        End If
        '
        LShoppingList.NavigateUrl = "http://10.245.0.205/EDI/SP_ShoppingListInf.aspx?pUserID=" & UCase(UserID) & ""
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Go)
    '**     Find Puller
    '**
    '*****************************************************************
    Protected Sub BFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFind.Click
        AtCloseHistory.Style("left") = -500 & "px"
        AtCloseHistory.Checked = False
        GridView2.Visible = False

        ShowData()
    End Sub

    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim SQL As String
        '
        SQL = "select top 1 [YOBI8L] + [YOBI9L]  as UpdateTime "
        SQL = SQL & "from M_SPStockProgress "
        SQL = SQL & "Order by Keep, ItemName1+' '+ItemName2, Color, Item "
        '
        Dim dtSPStockProgressX As DataTable = uDataBase.GetDataTable(SQL)
        If dtSPStockProgressX.Rows.Count > 0 Then
            DUpdateTime.Text = dtSPStockProgressX.Rows(0)("UpdateTime").ToString.Trim
        End If
        '
        SQL = "select top 150 "
        SQL = SQL & "Item, ItemName1+' '+ItemName2 as ItemName, Color, Keep, "
        SQL = SQL & "ScheProcS, OnProcS, FreeS, KeepS, ScheProcS+OnProcS+FreeS+KeepS as TotalS, "
        SQL = SQL & "N21IH, N21OH, N22IH, N22OH, N23IH, N23OH, N24IH, N24OH, "
        SQL = SQL & "N11IH, N11OH, N12IH, N12OH, N13IH, N13OH, N14IH, N14OH, "
        SQL = SQL & "N01IH, N01OH, N02IH, N02OH, N03IH, N03OH, N04IH, N04OH, "
        SQL = SQL & "'....' as FCL, 'http://10.245.0.205/EDI/SPActionPlan_02.aspx?pItem='+Item+'&pColor='+Color+'&pKeep='+Keep as FCLUrl, "
        SQL = SQL & "'....' as FDM, 'http://10.245.0.205/EDI/SPHttp2File.aspx?pFun=FDM&pSPNo=&pITEM=' + Item + '&pColor=' + Color + '&pKeep=' + Keep As FDMURL, "
        SQL = SQL & "'....' as PRODSTS, 'http://10.245.0.205/EDI/SPHttp2File.aspx?pFun=PROD&pSPNo=&pITEM=' + Item + '&pColor=' + Color + '&pKeep=' + Keep As PRODSTSURL "

        '
        SQL = SQL & "from M_SPStockProgress "
        SQL = SQL & "where keep <> '' "
        ' code
        If DKCode.Text <> "" Then
            SQL = SQL & "and item = '" & DKCode.Text & "' "
        End If
        ' Color
        If DKColor.Text <> "" Then
            SQL = SQL & "and ltrim(rtrim(Color)) = '" & DKColor.Text & "' "
        End If
        ' Keepr
        If DKKeepCode.Text <> "" Then
            SQL = SQL & "and keep = '" & DKKeepCode.Text & "' "
        End If
        ' Other
        If DKOther.Text <> "" And InStr("MK045/IT003/", UCase(DKOther.Text.Trim)) <= 0 Then
            SQL = SQL & "and Item + ItemName1 + ItemName2 + ItemName3 + Color + Keep like '%" & DKOther.Text & "%' "
        End If
        '
        SQL = SQL & "Order by Keep, ItemName1+' '+ItemName2, Color, Item "
        '
        '
        GridView1.Visible = True
        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     History Stock I/O
    '**
    '*****************************************************************
    Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing
        Dim SQL As String
        Dim xItem, xColor, xKeep As String
        xItem = Replace(GridView1.Rows(e.NewEditIndex).Cells(0).Text, "&nbsp;", "")
        xColor = Replace(GridView1.Rows(e.NewEditIndex).Cells(2).Text, "&nbsp;", "")
        xKeep = Replace(GridView1.Rows(e.NewEditIndex).Cells(3).Text, "&nbsp;", "")
        '
        xBalance = 0
        iSum = 0
        oSum = 0
        AtCloseHistory.Style("left") = 744 & "px"
        AtCloseHistory.Checked = False
        '
        SQL = "select top 150 "
        SQL = SQL & "Item, ItemName1+' '+ItemName2 as ItemName, Color, Keep, "
        SQL = SQL & "ScheProcS, OnProcS, FreeS, KeepS, ScheProcS+OnProcS+FreeS+KeepS as TotalS, "
        SQL = SQL & "N21IH, N21OH, N22IH, N22OH, N23IH, N23OH, N24IH, N24OH, "
        SQL = SQL & "N11IH, N11OH, N12IH, N12OH, N13IH, N13OH, N14IH, N14OH, "
        SQL = SQL & "N01IH, N01OH, N02IH, N02OH, N03IH, N03OH, N04IH, N04OH, "
        SQL = SQL & "'....' as FCL, 'http://10.245.0.205/EDI/SPActionPlan_02.aspx?pItem='+Item+'&pColor='+Color+'&pKeep='+Keep as FCLUrl, "
        SQL = SQL & "'....' as FDM, 'http://10.245.0.205/EDI/SPHttp2File.aspx?pFun=FDM&pSPNo=&pITEM=' + Item + '&pColor=' + Color + '&pKeep=' + Keep As FDMURL, "
        SQL = SQL & "'....' as PRODSTS, 'http://10.245.0.205/EDI/SPHttp2File.aspx?pFun=PROD&pSPNo=&pITEM=' + Item + '&pColor=' + Color + '&pKeep=' + Keep As PRODSTSURL "
        '
        SQL = SQL & "from M_SPStockProgress "
        SQL = SQL & "where Item = '" & xItem & "' "
        SQL = SQL & "  and LTRIM(RTRIM(Color)) = '" & Trim(xColor) & "' "
        SQL = SQL & "  and Keep = '" & xKeep & "' "
        SQL = SQL & "  and keep <> '' "
        SQL = SQL & "Order by Keep, ItemName1+' '+ItemName2, Color, Item "
        '
        GridView1.Visible = True
        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()
        '
        SQL = "SELECT "
        SQL &= "ITMC10, CLRC10, KEPC10, "
        '
        SQL &= "ORDN11 + CheckQty as ORDN11, "
        '
        SQL &= "CASE WHEN SUBSTRING(ORDN11,1,2)+CATG11='ORI' THEN 'http://10.245.1.6/ISOSSP/SPWingsOrder.aspx?pUserID=" & UserID & "&pType=ORI&pOrderNo='+ORDN11+'&pItem='+ITMC10+'&pColor='+CLRC10+'&pKeep='+KEPC10+'&pPlanNo='+PP "
        SQL &= "     WHEN SUBSTRING(ORDN11,1,2)+CATG11='ORO' THEN 'http://10.245.1.6/ISOSSP/SPWingsOrder.aspx?pUserID=" & UserID & "&pType=ORO&pOrderNo='+ORDN11+'&pItem='+ITMC10+'&pColor='+CLRC10+'&pKeep='+KEPC10+'&pPlanNo='+PP "
        SQL &= "     WHEN SUBSTRING(ORDN11,1,2)='ST' THEN 'http://10.245.1.6/ISOSSP/SPWingsOrder.aspx?pUserID=" & UserID & "&pType=ST&pItem='+ITMC10+'&pColor='+CLRC10+'&pKeep='+KEPC10 "
        SQL &= "     ELSE '' "
        SQL &= "END AS ORDN11URL, "
        '
        SQL &= "CORN11, OCNU11, "
        SQL &= "case when CATG11='I' then ORRQ11 else 0 end as IORRQ11, "
        SQL &= "case when CATG11='O' then ORRQ11 else 0 end as OORRQ11, "
        SQL &= "0 as BORRQ11, "
        SQL &= "REMK11 "
        '
        ' 管理者
        If InStr("MK045/IT003/", UCase(DKOther.Text.Trim)) <= 0 Or UCase(DKOther.Text.Trim) = "" Then
            SQL = SQL & "FROM V_SPSTOCK10N "
        Else
            SQL = SQL & "FROM V_SPSTOCK10 "
        End If
        '
        SQL = SQL & "where ITMC10 = 'ZZZ' "
        SQL = SQL & "OR ( "
        SQL = SQL & "      ITMC10 = '" & xItem & "' "
        SQL = SQL & "  and LTRIM(RTRIM(CLRC10)) = '" & Trim(xColor) & "' "
        SQL = SQL & "  and KEPC10 = '" & xKeep & "' "
        SQL = SQL & "  and KEPC10 <> '' "
        SQL = SQL & " ) "
        '
        SQL = SQL & "Order by OCNU11, ORDN11, CATG11, ORRQ11 DESC  "
        '
        GridView2.Visible = True
        GridView2.DataSource = uDataBase.GetDataTable(SQL)
        GridView2.DataBind()
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
            'Item
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = "Code"
            tcl(0).BackColor = Color.Blue
            'Name
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "Name"
            tcl(1).BackColor = Color.Blue
            'Color
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Color"
            tcl(2).BackColor = Color.Blue
            'keep
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "Keep"
            tcl(3).BackColor = Color.Blue
            '-----
            'Stock-Sche P.
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "Sche P."
            tcl(4).BackColor = Color.Purple
            'Stock-On P.
            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "On P."
            tcl(5).BackColor = Color.Purple
            'Stock-Free
            tcl.Add(New TableHeaderCell())
            tcl(6).Text = "Free"
            tcl(6).BackColor = Color.Purple
            'Stock-Keep
            tcl.Add(New TableHeaderCell())
            tcl(7).Text = "Keep"
            tcl(7).BackColor = Color.Purple
            'Stock-Total
            tcl.Add(New TableHeaderCell())
            tcl(8).Text = "Total"
            tcl(8).BackColor = Color.Purple
            '-----
            'Link
            'FC/Demand
            i = 8
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "F/D"
            tcl(i + 1).BackColor = Color.GreenYellow
            tcl(i + 1).ForeColor = Color.Black
            i = i + 1
            'FC/Demand(buymonth)
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "F/DM"
            tcl(i + 1).BackColor = Color.GreenYellow
            tcl(i + 1).ForeColor = Color.Black
            i = i + 1
            'P/S Production Status
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "P/S"
            tcl(i + 1).BackColor = Color.GreenYellow
            tcl(i + 1).ForeColor = Color.Black
            i = i + 1
            '3 YEAR IN/OUT
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "STK."
            tcl(i + 1).BackColor = Color.GreenYellow
            tcl(i + 1).ForeColor = Color.Black
            i = i + 1
            '-----
            'YY2-QUAR 10~17
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "I"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "O"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "I"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "O"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "I"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "O"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "I"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "O"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            'YY1-QUAR 18~25
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "I"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "O"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "I"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "O"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "I"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "O"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "I"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "O"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            'YY0-QUAR 26~33
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "I"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "O"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "I"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "O"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "I"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "O"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "I"
            tcl(i + 1).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i + 1).Text = "O"
            tcl(i + 1).BackColor = Color.Green
            '
            '** 3 line *****************
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = ""
            H3tc1.ColumnSpan = 4
            H3tc1.BackColor = Color.Blue
            H3row.Cells.Add(H3tc1)

            Dim H3tc2 As TableCell = New TableCell
            H3tc2.Text = ""
            H3tc2.ColumnSpan = 5
            H3tc2.BackColor = Color.Purple
            H3row.Cells.Add(H3tc2)

            Dim H3tc15 As TableCell = New TableCell
            H3tc15.Text = ""
            H3tc15.ColumnSpan = 4
            H3tc15.BackColor = Color.GreenYellow
            H3row.Cells.Add(H3tc15)
            '-----
            'YY2-QUAR
            Dim H3tc3 As TableCell = New TableCell
            H3tc3.Text = "Q1"
            H3tc3.ColumnSpan = 2
            H3tc3.BackColor = Color.Green
            H3row.Cells.Add(H3tc3)

            Dim H3tc4 As TableCell = New TableCell
            H3tc4.Text = "Q2"
            H3tc4.ColumnSpan = 2
            H3tc4.BackColor = Color.Green
            H3row.Cells.Add(H3tc4)

            Dim H3tc5 As TableCell = New TableCell
            H3tc5.Text = "Q3"
            H3tc5.ColumnSpan = 2
            H3tc5.BackColor = Color.Green
            H3row.Cells.Add(H3tc5)

            Dim H3tc6 As TableCell = New TableCell
            H3tc6.Text = "Q4"
            H3tc6.ColumnSpan = 2
            H3tc6.BackColor = Color.Green
            H3row.Cells.Add(H3tc6)

            'YY1-QUAR
            Dim H3tc7 As TableCell = New TableCell
            H3tc7.Text = "Q1"
            H3tc7.ColumnSpan = 2
            H3tc7.BackColor = Color.Green
            H3row.Cells.Add(H3tc7)

            Dim H3tc8 As TableCell = New TableCell
            H3tc8.Text = "Q2"
            H3tc8.ColumnSpan = 2
            H3tc8.BackColor = Color.Green
            H3row.Cells.Add(H3tc8)

            Dim H3tc9 As TableCell = New TableCell
            H3tc9.Text = "Q3"
            H3tc9.ColumnSpan = 2
            H3tc9.BackColor = Color.Green
            H3row.Cells.Add(H3tc9)

            Dim H3tc10 As TableCell = New TableCell
            H3tc10.Text = "Q4"
            H3tc10.ColumnSpan = 2
            H3tc10.BackColor = Color.Green
            H3row.Cells.Add(H3tc10)

            'YY0-QUAR
            Dim H3tc11 As TableCell = New TableCell
            H3tc11.Text = "Q1"
            H3tc11.ColumnSpan = 2
            H3tc11.BackColor = Color.Green
            H3row.Cells.Add(H3tc11)

            Dim H3tc12 As TableCell = New TableCell
            H3tc12.Text = "Q2"
            H3tc12.ColumnSpan = 2
            H3tc12.BackColor = Color.Green
            H3row.Cells.Add(H3tc12)

            Dim H3tc13 As TableCell = New TableCell
            H3tc13.Text = "Q3"
            H3tc13.ColumnSpan = 2
            H3tc13.BackColor = Color.Green
            H3row.Cells.Add(H3tc13)

            Dim H3tc14 As TableCell = New TableCell
            H3tc14.Text = "Q4"
            H3tc14.ColumnSpan = 2
            H3tc14.BackColor = Color.Green
            H3row.Cells.Add(H3tc14)
            '
            gv.Controls(0).Controls.AddAt(0, H3row)
            '
            '** 2 line *****************
            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "Material"
            H4tc1.ColumnSpan = 4
            H4tc1.BackColor = Color.Black
            H4row.Cells.Add(H4tc1)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = "Stock"
            H4tc2.ColumnSpan = 5
            H4tc2.BackColor = Color.Black
            H4row.Cells.Add(H4tc2)

            Dim H4tc21 As TableCell = New TableCell
            H4tc21.Text = "Link"
            H4tc21.ColumnSpan = 4
            H4tc21.BackColor = Color.Black
            H4row.Cells.Add(H4tc21)
            '-----
            'YY2-QUAR
            Dim H4tc3 As TableCell = New TableCell
            H4tc3.Text = "N-2 Year"
            H4tc3.ColumnSpan = 8
            H4tc3.BackColor = Color.Black
            H4row.Cells.Add(H4tc3)

            Dim H4tc4 As TableCell = New TableCell
            H4tc4.Text = "N-1 Year"
            H4tc4.ColumnSpan = 8
            H4tc4.BackColor = Color.Black
            H4row.Cells.Add(H4tc4)

            Dim H4tc5 As TableCell = New TableCell
            H4tc5.Text = "N Year"
            H4tc5.ColumnSpan = 8
            H4tc5.BackColor = Color.Black
            H4row.Cells.Add(H4tc5)
            '
            gv.Controls(0).Controls.AddAt(0, H4row)
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView1 編輯資料 顏色+格式
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            '
            For i = 0 To 36
                Select Case i
                    Case 0 To 3
                        e.Row.Cells(i).Font.Bold = True
                    Case 4 To 8
                        e.Row.Cells(i).Font.Bold = True
                        e.Row.Cells(i).Text = Format(CDbl(e.Row.Cells(i).Text), "###,###,##0.00")
                    Case 9 To 12
                        'e.Row.Cells(i).ForeColor = Color.Red
                    Case 13 To 36
                        e.Row.Cells(i).Font.Bold = True
                        e.Row.Cells(i).Text = Format(CDbl(e.Row.Cells(i).Text), "###,###,##0.00")
                    Case Else
                End Select

            Next
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
    '**     GridView2 編輯資料 TITLE
    '**
    '*****************************************************************
    Protected Sub GridView2_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowCreated
        '
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            '
            ' 清除
            e.Row.Cells.Clear()
            '
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '
            '** 4 line *****************
            'Item
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = "Code"
            tcl(0).BackColor = Color.Blue
            'Color
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "Color"
            tcl(1).BackColor = Color.Blue
            'keep
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Keep"
            tcl(2).BackColor = Color.Blue
            'OR/ST
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "OR/Stock Tran"
            tcl(3).BackColor = Color.Purple
            'Ref (CUST ORDER NO..)
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "C.Order No"
            tcl(4).BackColor = Color.Purple
            'Data Date
            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "Date"
            tcl(5).BackColor = Color.Purple
            'Input
            tcl.Add(New TableHeaderCell())
            tcl(6).Text = "I"
            tcl(6).BackColor = Color.Green
            'Output
            tcl.Add(New TableHeaderCell())
            tcl(7).Text = "O"
            tcl(7).BackColor = Color.Green
            'Balance
            tcl.Add(New TableHeaderCell())
            tcl(8).Text = "Balance"
            tcl(8).BackColor = Color.Green
            'Reference
            tcl.Add(New TableHeaderCell())
            tcl(9).Text = "Ref."
            tcl(9).BackColor = Color.Black
            '
            gv.Controls(0).Controls.AddAt(0, H3row)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView2 編輯資料 顏色+格式
    '**
    '*****************************************************************
    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        Dim i As Integer
        '
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
            '
            ' STOCK BALANCE 無法解決
            If InStr("MK045/IT003/", UCase(DKOther.Text.Trim)) <= 0 Or UCase(DKOther.Text.Trim) = "" Then
                e.Row.Cells(8).Visible = False
            End If
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            '
            If e.Row.Cells(0).Text = "ZZZ" Then
                For i = 0 To 9
                    e.Row.Cells(i).BackColor = Color.LightCyan
                    Select Case i
                        Case 0
                            e.Row.Cells(i).Text = "Total"
                        Case 6
                            e.Row.Cells(i).Font.Bold = True
                            e.Row.Cells(i).Text = Format(iSum, "###,###,##0.00")
                        Case 7
                            e.Row.Cells(i).Font.Bold = True
                            e.Row.Cells(i).Text = Format(oSum, "###,###,##0.00")
                        Case 8
                            '
                            ' STOCK BALANCE 無法解決
                            If InStr("MK045/IT003/", UCase(DKOther.Text.Trim)) <= 0 Or UCase(DKOther.Text.Trim) = "" Then
                                e.Row.Cells(i).Visible = False
                            End If
                        Case Else
                            e.Row.Cells(i).Text = ""
                    End Select
                Next
            Else
                '
                For i = 0 To 9
                    Select Case i
                        Case 6
                            e.Row.Cells(i).Font.Bold = True
                            e.Row.Cells(i).Text = Format(CDbl(e.Row.Cells(i).Text), "###,###,##0.00")
                            iSum = iSum + CDbl(e.Row.Cells(i).Text)
                            If e.Row.Cells(0).Text = "ZZZ" Then
                                e.Row.Cells(i).Text = Format(iSum, "###,###,##0.00")
                            End If
                        Case 7
                            e.Row.Cells(i).Font.Bold = True
                            e.Row.Cells(i).Text = Format(CDbl(e.Row.Cells(i).Text), "###,###,##0.00")
                            oSum = oSum + CDbl(e.Row.Cells(i).Text)
                            If e.Row.Cells(0).Text = "ZZZ" Then
                                e.Row.Cells(i).Text = Format(oSum, "###,###,##0.00")
                            End If
                        Case 8
                            Select Case e.Row.RowIndex
                                Case 0
                                    xBalance = CDbl(e.Row.Cells(6).Text) - CDbl(e.Row.Cells(7).Text)
                                Case Else
                                    If e.Row.Cells(0).Text = "ZZZ" Then
                                        xBalance = CDbl(e.Row.Cells(6).Text) - CDbl(e.Row.Cells(7).Text)
                                    Else
                                        xBalance = xBalance + CDbl(e.Row.Cells(6).Text) - CDbl(e.Row.Cells(7).Text)
                                    End If
                            End Select
                            e.Row.Cells(i).Font.Bold = True
                            e.Row.Cells(i).Text = Format(xBalance, "###,###,##0.00")
                            '
                            ' STOCK BALANCE 無法解決
                            If InStr("MK045/IT003/", UCase(DKOther.Text.Trim)) <= 0 Or UCase(DKOther.Text.Trim) = "" Then
                                e.Row.Cells(i).Visible = False
                            End If
                        Case Else
                    End Select
                    '
                Next                '
            End If


            '
            Select Case e.Row.RowIndex
                Case 0
                    '
                    ' STOCK BALANCE 無法解決
                    If InStr("MK045/IT003/", UCase(DKOther.Text.Trim)) <= 0 Or UCase(DKOther.Text.Trim) = "" Then
                    Else
                        For i = 0 To 9
                            If i = 9 Then
                                If e.Row.Cells(i).Text = "OPENING STOCK/" Then
                                    e.Row.Cells(5).Text = ""
                                End If
                            End If
                            e.Row.Cells(i).BackColor = Color.LightCyan
                        Next
                    End If
                    '
                Case Else
                    If e.Row.Cells(0).Text = "ZZZ" Then
                        e.Row.Cells(0).Text = ""
                        e.Row.Cells(5).Text = ""
                        For i = 0 To 9
                            e.Row.Cells(i).BackColor = Color.LightCyan
                        Next
                    End If
            End Select
            '
            '
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
    '**(Close IMG Windows)
    '**     
    '**
    '*****************************************************************
    Protected Sub AtCloseHistory_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtCloseHistory.CheckedChanged
        AtCloseHistory.Style("left") = -500 & "px"
        AtCloseHistory.Checked = False
        '
        GridView2.Visible = False
        '
        ShowData()
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Material Plam Page)
    '**     
    '**
    '*****************************************************************
    Protected Sub BMaterial_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BMaterial.Click
        Dim Cmd As String

        Cmd = "<script>" + _
                    "window.open('http://10.245.0.205/EDI/SP_MaterialMenu.aspx?pUserID=" & UserID & "','MaterialPlan','');" + _
              "</script>"

        Response.Write(Cmd)
        '
    End Sub
End Class
