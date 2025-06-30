Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfF_FCTACTVDPAnalysis_03
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
    Dim gItem As String             'CustItem
    Dim gColor As String            'Y_Color
    Dim gVersion As String          'Version
    Dim gOption As String           'Option
    Dim UserID As String            'UserID
    Dim ACTTotal, FCTTotal As Integer         '合計數量

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
        Response.Cookies("PGM").Value = "InfF_FCTACTVDPAnalysis_03.aspx"  '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")  '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                 '現在日期時間
        gBuyer = Request.QueryString("pBuyer")             'Buyer
        gCustCode = Request.QueryString("pCustCode")       'Customer Code
        gSeason = Request.QueryString("pSeason")           'Season
        gMonth = Request.QueryString("pMonth")             'Month
        gItem = Request.QueryString("pItem")               'CustItem
        gColor = Request.QueryString("pColor")             'Y_Color
        gVersion = Request.QueryString("pVersion")         'Version
        gOption = Request.QueryString("pOption")           'Option
        UserID = Request.QueryString("pUserID")            'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        FCTGridView.Visible = False
        OrderGridView.Visible = False
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
        If gOption = "FCT" Then
            FCTGridView.Visible = True
        Else
            OrderGridView.Visible = True
        End If
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
        If gOption = "FCT" Then
            sql = "SELECT "
            ' T&P NIKE=000013T
            If gBuyer = "000013T" Then
                sql &= "(Select Top 1 x.CustName From M_NativeVendor x Where x.Buyer = 'FALL-TP000013' And x.CustCode = a.CustCode) + '(' + a.CustCode + ')' As Customer, "
            Else
                sql &= "(Select Top 1 x.CustName From M_NativeVendor x Where x.Buyer = 'FALL-' + a.Buyer And x.CustCode = a.CustCode) + '(' + a.CustCode + ')' As Customer, "
            End If
            sql &= "a.Buyer, a.CustCode, a.Season, a.Month, a.CustItem, a.Color, a.Style, a.Article, a.Part, "
            sql &= "a.KeyData1, a.KeyData2, a.Remark1 + ' ' + a.Remark2 As ItemName, "
            sql &= "a.FCTQty "
            sql &= "From A_CustomerActual_FCT_VDP a "
            sql &= "Where a.Buyer  = '" & gBuyer & "' "
            sql &= "  And a.CustCode  = '" & gCustCode & "' "
            sql &= "  And a.Season = '" & gSeason & "' "
            sql &= "  And a.Month = '" & gMonth & "' "
            sql &= "  And a.CustItem = '" & gItem & "' "
            If gColor <> "" Then
                sql &= "  And a.KeyData2 = '" & gColor & "' "
            End If
            sql &= "  And a.Version = '" & gVersion & "' "
            sql &= "Order by a.Buyer, a.CustCode, a.Season, a.Month, a.CustItem, a.KeyData1, a.KeyData2 "
            '
            Dim dt_FCTPlan_ITEM As DataTable = uDataBase.GetDataTable(sql)
            If dt_FCTPlan_ITEM.Rows.Count > 0 Then
                FCTGridView.Visible = True
                FCTGridView.DataSource = dt_FCTPlan_ITEM
                FCTGridView.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
            End If
        Else
            sql = "SELECT "
            ' T&P NIKE=000013T
            If gBuyer = "000013T" Then
                sql &= "(Select Top 1 x.CustName From M_NativeVendor x Where x.Buyer = 'FALL-TP000013' And x.CustCode = a.CustCode) + '(' + a.CustCode + ')' As Customer, "
            Else
                sql &= "(Select Top 1 x.CustName From M_NativeVendor x Where x.Buyer = 'FALL-' + a.Buyer And x.CustCode = a.CustCode) + '(' + a.CustCode + ')' As Customer, "
            End If
            sql &= "a.Buyer, a.CustCode, a.Season, a.V_Month, a.CustItem, "
            sql &= "a.Item, a.ItemName1 + ' ' + a.ItemName2 As ItemName, a.Color, a.ACTQty, "
            sql &= "'InfF_FCTACTVDPAnalysis_04.aspx?' + "
            sql &= "'pBuyer=' + Buyer + "
            sql &= "'&pCustCode=' + CustCode + "
            sql &= "'&pSeason=' + Season + "
            sql &= "'&pMonth=' + V_Month + "
            sql &= "'&pCustItem=' + Replace(CustItem, '#', '%23') + "
            sql &= "'&pItem=' + Item + "
            sql &= "'&pColor=' + Color "
            sql &= " As URL "
            sql &= "From A_CustomerActual_Order_VDP a "
            sql &= "Where a.Buyer  = '" & gBuyer & "' "
            sql &= "  And a.CustCode  = '" & gCustCode & "' "
            sql &= "  And a.Season = '" & gSeason & "' "
            sql &= "  And a.V_Month = '" & gMonth & "' "
            sql &= "  And a.CustItem = '" & gItem & "' "
            If gColor <> "" Then
                sql &= "  And a.Color = '" & gColor & "' "
            End If
            sql &= "Order by a.Buyer, a.CustCode, a.Season, a.V_Month, a.CustItem, a.Item, a.Color "
            '
            Dim dt_FCTPlan_ITEM As DataTable = uDataBase.GetDataTable(sql)
            If dt_FCTPlan_ITEM.Rows.Count > 0 Then
                OrderGridView.Visible = True
                OrderGridView.DataSource = dt_FCTPlan_ITEM
                OrderGridView.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
            End If
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
            e.Row.Cells(7).Text = Format(DataBinder.Eval(e.Row.DataItem, "ACTQty"), "###,###,###")

            ACTTotal = ACTTotal + DataBinder.Eval(e.Row.DataItem, "ACTQty")
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
            ' 合計展開
            Dim row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)

            Dim tc1 As TableCell = New TableCell
            tc1.BackColor = Color.YellowGreen
            tc1.ColumnSpan = 7
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
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料(FCT GridView)
    '**
    '*****************************************************************
    Protected Sub FCTGridView_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles FCTGridView.RowDataBound
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            e.Row.Cells(11).Text = Format(DataBinder.Eval(e.Row.DataItem, "FCTQty"), "###,###,###")

            FCTTotal = FCTTotal + DataBinder.Eval(e.Row.DataItem, "FCTQty")
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
            ' 合計展開
            Dim row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)

            Dim tc1 As TableCell = New TableCell
            tc1.BackColor = Color.YellowGreen
            tc1.ColumnSpan = 11
            tc1.Text = "合計"
            row.Cells.Add(tc1)

            Dim tc2 As TableCell = New TableCell
            tc2.HorizontalAlign = HorizontalAlign.Right
            tc2.BackColor = Color.YellowGreen
            tc2.Text = Format(FCTTotal, "###,###,###")
            row.Cells.Add(tc2)

            e.Row.Parent.Controls.Add(row)
        End If

    End Sub

End Class
