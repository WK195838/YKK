Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfF_FCTACTAnalysis_03a
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
        Response.Cookies("PGM").Value = "InfF_FCTACTAnalysis_03a.aspx"  '程式名
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
            sql &= "OrderNo, OrderSubNo, PO, Seqno, CustWavesCode, Season, Item, Length, Unit, Color, OrderDate, ReqDate, CompletedDate, OrderQty, V_Month, "
            sql &= "ItemName1 + ' ' + ItemName2 + ' ' + ItemName3 AS ItemName, "
            sql &= "'' AS sItem, '' AS sItemName, '' AS sColor "
            sql &= "From I_Adidas_ActualOrder "
            sql &= "Where Buyer  = '" & gBuyer & "' "
            sql &= "  And CustCode  = '" & gCustCode & "' "
            sql &= "  And Season = '" & gSeason & "' "
            sql &= "  And Month = '" & gMonth & "' "
            sql &= "  And V_PLM = '" & gCustItem & "' "
            sql &= "  And Item = '" & gItem & "' "
            sql &= "  And Color = '" & gColor & "' "
            sql &= "Order by OrderNo, OrderSubNo, Item, Color "
        End If
        '
        If gBuyer = "000013" Then
            sql = "SELECT "
            sql &= "OrderNo, OrderSubNo, PO, Seqno, CustWavesCode, Season, Item, Length, Unit, Color, OrderDate, ReqDate, CompletedDate, OrderQty, V_Month, "
            sql &= "ItemName1 + ' ' + ItemName2 + ' ' + ItemName3 AS ItemName, "
            sql &= "'' AS sItem, '' AS sItemName, '' AS sColor "
            sql &= "From I_Nike_ActualOrder "
            sql &= "Where Buyer  = '" & gBuyer & "' "
            sql &= "  And CustCode  = '" & gCustCode & "' "
            sql &= "  And Season = '" & gSeason & "' "
            sql &= "  And Month = '" & gMonth & "' "
            sql &= "  And V_PLM = '" & gCustItem & "' "
            sql &= "  And Item = '" & gItem & "' "
            sql &= "  And Color = '" & gColor & "' "
            sql &= "Order by OrderNo, OrderSubNo, Item, Color "
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
        Dim i As Integer
        Dim sql As String
        Dim xItem, xItemName, xColor As String
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            e.Row.Cells(15).Text = Format(DataBinder.Eval(e.Row.DataItem, "OrderQty"), "###,###,###")
            '
            xItem = ""
            xItemName = ""
            xColor = ""
            sql = "SELECT "
            sql &= "Item, Color, ItemName1 + ' ' + ItemName2 + ' ' + ItemName3 AS ItemName "
            sql &= "From I_OrderProgress_Slider "
            sql &= "Where OrderNo  = '" & DataBinder.Eval(e.Row.DataItem, "OrderNo") & "' "
            sql &= "  And OrderSubNo = " & CStr(DataBinder.Eval(e.Row.DataItem, "OrderSubNo")) & " "
            sql &= "Order by ItemName1, ItemName2, Color "
            '
            Dim dt_OrderSlider As DataTable = uDataBase.GetDataTable(sql)
            For i = 0 To dt_OrderSlider.Rows.Count - 1
                If xItem = "" Then
                    xItem = dt_OrderSlider.Rows(i).Item("Item")
                    xItemName = dt_OrderSlider.Rows(i).Item("ItemName")
                    xColor = dt_OrderSlider.Rows(i).Item("Color")
                Else
                    xItem = xItem + " / " + dt_OrderSlider.Rows(i).Item("Item")
                    xItemName = xItemName + " / " + dt_OrderSlider.Rows(i).Item("ItemName")
                    xColor = xColor + " / " + dt_OrderSlider.Rows(i).Item("Color")
                End If
            Next
            e.Row.Cells(16).Text = xItem
            e.Row.Cells(17).Text = xItemName
            e.Row.Cells(18).Text = xColor
            '
            ACTTotal = ACTTotal + DataBinder.Eval(e.Row.DataItem, "OrderQty")
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
            ' 合計展開
            Dim row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)

            Dim tc1 As TableCell = New TableCell
            tc1.BackColor = Color.YellowGreen
            tc1.ColumnSpan = 15
            tc1.Text = "合計"
            row.Cells.Add(tc1)

            Dim tc2 As TableCell = New TableCell
            tc2.HorizontalAlign = HorizontalAlign.Right
            tc2.BackColor = Color.YellowGreen
            tc2.Text = Format(ACTTotal, "###,###,###")
            row.Cells.Add(tc2)

            Dim tc3 As TableCell = New TableCell
            tc3.BackColor = Color.YellowGreen
            tc3.ColumnSpan = 3
            tc3.Text = ""
            row.Cells.Add(tc3)

            e.Row.Parent.Controls.Add(row)
        End If
    End Sub
End Class
