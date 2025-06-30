Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfF_FCTACTVDPAnalysis_04
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
    Dim gCustItem As String         'CustItem
    Dim gItem As String             'Item
    Dim gColor As String            'Color
    Dim UserID As String            'UserID
    Dim ACTTotal As Integer         '合計數量

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
        Response.Cookies("PGM").Value = "InfF_FCTACTVDPAnalysis_04.aspx"  '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")  '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                 '現在日期時間
        gBuyer = Request.QueryString("pBuyer")             'Buyer
        gCustCode = Request.QueryString("pCustCode")       'Customer Code
        gSeason = Request.QueryString("pSeason")           'Season
        gMonth = Request.QueryString("pMonth")             'Month
        gCustItem = Request.QueryString("pCustItem")       'CustItem
        gItem = Request.QueryString("pItem")               'Item
        gColor = Request.QueryString("pColor")             'Color
        UserID = Request.QueryString("pUserID")            'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        ACTTotal = 0
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
        '
        If gBuyer = "000001" Then
            sql = "SELECT "
            sql &= "OrderNo, OrderSubNo, PO, Seqno, CustWavesCode, Season, V_Month, Item, Length, Unit, Color, OrderDate, ReqDate, CompletedDate, OrderQty "
            sql &= "From I_Adidas_ActualOrder "
            sql &= "Where Buyer  = '" & gBuyer & "' "
            sql &= "  And CustCode  = '" & gCustCode & "' "
            sql &= "  And Season = '" & gSeason & "' "
            sql &= "  And V_Month = '" & gMonth & "' "
            sql &= "  And V_PLM = '" & gCustItem & "' "
            sql &= "  And Item = '" & gItem & "' "
            sql &= "  And Color = '" & gColor & "' "
            sql &= "Order by OrderNo, OrderSubNo, Item, Color "
        End If
        '
        If gBuyer = "000013" Then
            sql = "SELECT "
            sql &= "OrderNo, OrderSubNo, PO, Seqno, CustWavesCode, Season, V_Month, Item, Length, Unit, Color, OrderDate, ReqDate, CompletedDate, OrderQty "
            sql &= "From I_Nike_ActualOrder "
            sql &= "Where Buyer  = '" & gBuyer & "' "
            sql &= "  And CustCode  = '" & gCustCode & "' "
            sql &= "  And Season = '" & gSeason & "' "
            sql &= "  And V_Month = '" & gMonth & "' "
            sql &= "  And V_PLM = '" & gCustItem & "' "
            sql &= "  And Item = '" & gItem & "' "
            sql &= "  And Color = '" & gColor & "' "
            sql &= "Order by OrderNo, OrderSubNo, Item, Color "
        End If
        '
        If gBuyer = "000016" Then
            sql = "SELECT "
            sql &= "OrderNo, OrderSubNo, PO, Seqno, CustWavesCode, Season, V_Month, Item, Length, Unit, Color, OrderDate, ReqDate, CompletedDate, OrderQty "
            sql &= "From I_Reebok_ActualOrder "
            sql &= "Where Buyer  = '" & gBuyer & "' "
            sql &= "  And CustCode  = '" & gCustCode & "' "
            sql &= "  And Season = '" & gSeason & "' "
            sql &= "  And V_Month = '" & gMonth & "' "
            sql &= "  And V_PLM = '" & gCustItem & "' "
            sql &= "  And Item = '" & gItem & "' "
            sql &= "  And Color = '" & gColor & "' "
            sql &= "Order by OrderNo, OrderSubNo, Item, Color "
        End If
        '
        If gBuyer = "000021" Then
            sql = "SELECT "
            sql &= "OrderNo, OrderSubNo, PO, Seqno, CustWavesCode, Season, V_Month, Item, Length, Unit, Color, OrderDate, ReqDate, CompletedDate, OrderQty "
            sql &= "From I_TNF_ActualOrder "
            sql &= "Where Buyer  = '" & gBuyer & "' "
            sql &= "  And CustCode  = '" & gCustCode & "' "
            sql &= "  And SeasonGroup = '" & gSeason & "' "
            sql &= "  And V_Month = '" & gMonth & "' "
            sql &= "  And V_PLM = '" & gCustItem & "' "
            sql &= "  And Item = '" & gItem & "' "
            sql &= "  And Color = '" & gColor & "' "
            sql &= "Order by OrderNo, OrderSubNo, Item, Color "
        End If
        '
        If gBuyer = "000003" Then
            sql = "SELECT "
            sql &= "OrderNo, OrderSubNo, PO, Seqno, CustWavesCode, Season, V_Month, Item, Length, Unit, Color, OrderDate, ReqDate, CompletedDate, OrderQty "
            sql &= "From I_COLUMBIA_ActualOrder "
            sql &= "Where Buyer  = '" & gBuyer & "' "
            sql &= "  And CustCode  = '" & gCustCode & "' "
            sql &= "  And SeasonGroup = '" & gSeason & "' "
            sql &= "  And V_Month = '" & gMonth & "' "
            sql &= "  And V_PLM = '" & gCustItem & "' "
            sql &= "  And Item = '" & gItem & "' "
            sql &= "  And Color = '" & gColor & "' "
            sql &= "Order by OrderNo, OrderSubNo, Item, Color "
        End If
        '
        If gBuyer = "TW0371" Then
            sql = "SELECT "
            sql &= "OrderNo, OrderSubNo, PO, Seqno, CustWavesCode, Season, V_Month, Item, Length, Unit, Color, OrderDate, ReqDate, CompletedDate, OrderQty "
            sql &= "From I_UNDERARMOUR_ActualOrder "
            sql &= "Where Buyer  = '" & gBuyer & "' "
            sql &= "  And CustCode  = '" & gCustCode & "' "
            sql &= "  And SeasonGroup = '" & gSeason & "' "
            sql &= "  And V_Month = '" & gMonth & "' "
            sql &= "  And V_PLM = '" & gCustItem & "' "
            sql &= "  And Item = '" & gItem & "' "
            sql &= "  And Color = '" & gColor & "' "
            sql &= "Order by OrderNo, OrderSubNo, Item, Color "
        End If
        '
        If gBuyer = "000013T" Then
            sql = "SELECT "
            sql &= "OrderNo, OrderSubNo, PO, Seqno, CustWavesCode, Season, V_Month, Item, Length, Unit, Color, OrderDate, ReqDate, CompletedDate, OrderQty "
            sql &= "From I_TPNIKE_ActualOrder "
            sql &= "Where Buyer  = '" & "000013" & "' "
            sql &= "  And CustCode  = '" & gCustCode & "' "
            sql &= "  And SeasonGroup = '" & gSeason & "' "
            sql &= "  And V_Month = '" & gMonth & "' "
            'sql &= "  And V_PLM = '" & gCustItem & "' "
            sql &= "  And Item = '" & gItem & "' "
            sql &= "  And Color = '" & gColor & "' "
            sql &= "Order by OrderNo, OrderSubNo, Item, Color "
        End If
        '
        If gBuyer = "TW0371T" Then
            sql = "SELECT "
            sql &= "OrderNo, OrderSubNo, PO, Seqno, CustWavesCode, Season, V_Month, Item, Length, Unit, Color, OrderDate, ReqDate, CompletedDate, OrderQty "
            sql &= "From I_UABAGS_ActualOrder "
            sql &= "Where Buyer  = '" & gBuyer & "' "
            sql &= "  And CustCode  = '" & gCustCode & "' "
            sql &= "  And SeasonGroup = '" & gSeason & "' "
            sql &= "  And V_Month = '" & gMonth & "' "
            sql &= "  And V_PLM = '" & gCustItem & "' "
            sql &= "  And Item = '" & gItem & "' "
            sql &= "  And Color = '" & gColor & "' "
            sql &= "Order by OrderNo, OrderSubNo, Item, Color "
        End If
        '
        If gBuyer = "TW1741" Then
            '不支援 WAVE'S ACTUAL
        End If
        '
        Dim dt_FCTPlan_ITEM As DataTable = uDataBase.GetDataTable(sql)
        If dt_FCTPlan_ITEM.Rows.Count > 0 Then
            OrderGridView.Visible = True
            OrderGridView.DataSource = dt_FCTPlan_ITEM
            OrderGridView.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料(Order GridView)
    '**
    '*****************************************************************
    Protected Sub OrderGridView_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles OrderGridView.RowDataBound
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            e.Row.Cells(14).Text = Format(DataBinder.Eval(e.Row.DataItem, "OrderQty"), "###,###,###")

            ACTTotal = ACTTotal + DataBinder.Eval(e.Row.DataItem, "OrderQty")
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
            ' 合計展開
            Dim row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)

            Dim tc1 As TableCell = New TableCell
            tc1.BackColor = Color.YellowGreen
            tc1.ColumnSpan = 14
            tc1.Text = "合計"
            row.Cells.Add(tc1)

            Dim tc2 As TableCell = New TableCell
            tc2.HorizontalAlign = HorizontalAlign.Right
            tc2.BackColor = Color.YellowGreen
            tc2.Text = Format(ACTTotal, "###,###,###")
            row.Cells.Add(tc2)

            e.Row.Parent.Controls.Add(row)
        End If
    End Sub
End Class
