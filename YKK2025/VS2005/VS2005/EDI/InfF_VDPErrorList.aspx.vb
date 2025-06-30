Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfF_VDPErrorList
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
        Response.Cookies("PGM").Value = "InfF_VDPErrorList.aspx"     '程式名
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
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        DBuyer.ReadOnly = True
        GridView1.Visible = False
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

        DYY.Items.Clear()
        For i As Integer = 2014 To 2024
            Dim ListItem1 As New ListItem
            ListItem1.Text = CStr(i)
            ListItem1.Value = CStr(i)
            If i = CInt(Mid(NowDate, 1, 4)) Then
                ListItem1.Selected = True
            End If
            DYY.Items.Add(ListItem1)
        Next

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
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Inq)
    '**     查詢
    '**
    '*****************************************************************
    Protected Sub BInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BInq.Click
        Dim sql As String
        ' 設定篩選條件
        Dim sqlstr As String = "Season='' or V_Season='' or V_Year<Year(Getdate())-1 or V_Year>Year(Getdate())+1 or V_PLM=''"

        If AtBuyMonth.Checked = True Then
            sqlstr = sqlstr + " or V_BuyMonth<1 or V_BuyMonth>12"
        End If
        If AtStyle.Checked = True Then
            sqlstr = sqlstr + " or V_Style=''"
        End If
        If AtADIDASColor.Checked = True Then
            sqlstr = sqlstr + " or V_TapeColor=''"
        End If
        sqlstr = "( " + sqlstr + " )"
        ' 取得資料
        If DBuyer.Text = "000001" Then
            sql = "SELECT "
            sql &= "Case When CompletedFlag = '1' Then '已出貨' else '未出貨' End As Completed, "
            sql &= "CustWavesCode, OrderNo, Month, Season, V_Season, V_Year, V_PLM, V_BUYMONTH, V_TapeColor, V_Style, OrderDate "
            sql &= "From I_Adidas_ActualOrder "

            sql &= "Where Buyer = '" & DBuyer.Text & "' "
            sql &= "  And Month = '" & DYY.Text + DMM.Text & "' "
            sql &= "  And " & sqlstr & " "
            sql &= "  And ( "
            sql &= "        Select Count(*) From E_VDPErrorList "
            sql &= "        Where E_VDPErrorList.Buyer = '" & DBuyer.Text & "' "
            sql &= "          And E_VDPErrorList.OrderNo = dbo.I_Adidas_ActualOrder.OrderNo "
            sql &= "          And E_VDPErrorList.Finished = '" & "Y" & "' "
            sql &= "      ) = 0 "
            sql &= "Group by CustWavesCode, OrderNo, Month, Season, V_Season, V_Year, V_PLM, V_BUYMONTH, V_TapeColor, V_Style, OrderDate, CompletedFlag "
            sql &= "Order by CustWavesCode, Month, OrderNo "
        End If
        '
        If DBuyer.Text = "000013" Then
            sql = "SELECT "
            sql &= "Case When CompletedFlag = '1' Then '已出貨' else '未出貨' End As Completed, "
            sql &= "CustWavesCode, OrderNo, Month, Season, V_Season, V_Year, V_PLM, V_BUYMONTH, V_TapeColor, V_Style, OrderDate "
            sql &= "From I_Nike_ActualOrder "

            sql &= "Where Buyer = '" & DBuyer.Text & "' "
            sql &= "  And Month = '" & DYY.Text + DMM.Text & "' "
            sql &= "  And " & sqlstr & " "
            sql &= "  And ( "
            sql &= "        Select Count(*) From E_VDPErrorList "
            sql &= "        Where E_VDPErrorList.Buyer = '" & DBuyer.Text & "' "
            sql &= "          And E_VDPErrorList.OrderNo = dbo.I_Nike_ActualOrder.OrderNo "
            sql &= "          And E_VDPErrorList.Finished = '" & "Y" & "' "
            sql &= "      ) = 0 "
            sql &= "Group by CustWavesCode, OrderNo, Month, Season, V_Season, V_Year, V_PLM, V_BUYMONTH, V_TapeColor, V_Style, OrderDate, CompletedFlag "
            sql &= "Order by CustWavesCode, Month, OrderNo "
        End If
        '
        If DBuyer.Text = "000016" Then
            sql = "SELECT "
            sql &= "Case When CompletedFlag = '1' Then '已出貨' else '未出貨' End As Completed, "
            sql &= "CustWavesCode, OrderNo, Month, Season, V_Season, V_Year, V_PLM, V_BUYMONTH, V_TapeColor, V_Style, OrderDate "
            sql &= "From I_Reebok_ActualOrder "

            sql &= "Where Buyer = '" & DBuyer.Text & "' "
            sql &= "  And Month = '" & DYY.Text + DMM.Text & "' "
            sql &= "  And " & sqlstr & " "
            sql &= "  And ( "
            sql &= "        Select Count(*) From E_VDPErrorList "
            sql &= "        Where E_VDPErrorList.Buyer = '" & DBuyer.Text & "' "
            sql &= "          And E_VDPErrorList.OrderNo = dbo.I_Reebok_ActualOrder.OrderNo "
            sql &= "          And E_VDPErrorList.Finished = '" & "Y" & "' "
            sql &= "      ) = 0 "
            sql &= "Group by CustWavesCode, OrderNo, Month, Season, V_Season, V_Year, V_PLM, V_BUYMONTH, V_TapeColor, V_Style, OrderDate, CompletedFlag "
            sql &= "Order by CustWavesCode, Month, OrderNo "
        End If
        '
        If DBuyer.Text = "000021" Then
            sql = "SELECT "
            sql &= "Case When CompletedFlag = '1' Then '已出貨' else '未出貨' End As Completed, "
            sql &= "CustWavesCode, OrderNo, Month, Season, V_Season, V_Year, V_PLM, V_BUYMONTH, V_TapeColor, V_Style, OrderDate "
            sql &= "From I_TNF_ActualOrder "

            sql &= "Where Buyer = '" & DBuyer.Text & "' "
            sql &= "  And Month = '" & DYY.Text + DMM.Text & "' "
            sql &= "  And " & sqlstr & " "
            sql &= "  And ( "
            sql &= "        Select Count(*) From E_VDPErrorList "
            sql &= "        Where E_VDPErrorList.Buyer = '" & DBuyer.Text & "' "
            sql &= "          And E_VDPErrorList.OrderNo = dbo.I_TNF_ActualOrder.OrderNo "
            sql &= "          And E_VDPErrorList.Finished = '" & "Y" & "' "
            sql &= "      ) = 0 "
            sql &= "Group by CustWavesCode, OrderNo, Month, Season, V_Season, V_Year, V_PLM, V_BUYMONTH, V_TapeColor, V_Style, OrderDate, CompletedFlag "
            sql &= "Order by CustWavesCode, Month, OrderNo "
        End If
        '
        If DBuyer.Text = "000003" Then
            sql = "SELECT "
            sql &= "Case When CompletedFlag = '1' Then '已出貨' else '未出貨' End As Completed, "
            sql &= "CustWavesCode, OrderNo, Month, Season, V_Season, V_Year, V_PLM, V_BUYMONTH, V_TapeColor, V_Style, OrderDate "
            sql &= "From I_COLUMBIA_ActualOrder "

            sql &= "Where Buyer = '" & DBuyer.Text & "' "
            sql &= "  And Month = '" & DYY.Text + DMM.Text & "' "
            sql &= "  And " & sqlstr & " "
            sql &= "  And ( "
            sql &= "        Select Count(*) From E_VDPErrorList "
            sql &= "        Where E_VDPErrorList.Buyer = '" & DBuyer.Text & "' "
            sql &= "          And E_VDPErrorList.OrderNo = dbo.I_COLUMBIA_ActualOrder.OrderNo "
            sql &= "          And E_VDPErrorList.Finished = '" & "Y" & "' "
            sql &= "      ) = 0 "
            sql &= "Group by CustWavesCode, OrderNo, Month, Season, V_Season, V_Year, V_PLM, V_BUYMONTH, V_TapeColor, V_Style, OrderDate, CompletedFlag "
            sql &= "Order by CustWavesCode, Month, OrderNo "
        End If
        '
        If DBuyer.Text = "TW0371" Then
            sql = "SELECT "
            sql &= "Case When CompletedFlag = '1' Then '已出貨' else '未出貨' End As Completed, "
            sql &= "CustWavesCode, OrderNo, Month, Season, V_Season, V_Year, V_PLM, V_BUYMONTH, V_TapeColor, V_Style, OrderDate "
            sql &= "From I_UNDERARMOUR_ActualOrder "

            sql &= "Where Buyer = '" & DBuyer.Text & "' "
            sql &= "  And Month = '" & DYY.Text + DMM.Text & "' "
            sql &= "  And " & sqlstr & " "
            sql &= "  And ( "
            sql &= "        Select Count(*) From E_VDPErrorList "
            sql &= "        Where E_VDPErrorList.Buyer = '" & DBuyer.Text & "' "
            sql &= "          And E_VDPErrorList.OrderNo = dbo.I_UNDERARMOUR_ActualOrder.OrderNo "
            sql &= "          And E_VDPErrorList.Finished = '" & "Y" & "' "
            sql &= "      ) = 0 "
            sql &= "Group by CustWavesCode, OrderNo, Month, Season, V_Season, V_Year, V_PLM, V_BUYMONTH, V_TapeColor, V_Style, OrderDate, CompletedFlag "
            sql &= "Order by CustWavesCode, Month, OrderNo "
        End If
        '
        If DBuyer.Text = "000013T" Then
            sql = "SELECT "
            sql &= "Case When CompletedFlag = '1' Then '已出貨' else '未出貨' End As Completed, "
            sql &= "CustWavesCode, OrderNo, Month, Season, V_Season, V_Year, V_PLM, V_BUYMONTH, V_TapeColor, V_Style, OrderDate "
            sql &= "From I_TPNIKE_ActualOrder "

            sql &= "Where Buyer = '" & DBuyer.Text & "' "
            sql &= "  And Month = '" & DYY.Text + DMM.Text & "' "
            sql &= "  And " & sqlstr & " "
            sql &= "  And ( "
            sql &= "        Select Count(*) From E_VDPErrorList "
            sql &= "        Where E_VDPErrorList.Buyer = '" & DBuyer.Text & "' "
            sql &= "          And E_VDPErrorList.OrderNo = dbo.I_TPNIKE_ActualOrder.OrderNo "
            sql &= "          And E_VDPErrorList.Finished = '" & "Y" & "' "
            sql &= "      ) = 0 "
            sql &= "Group by CustWavesCode, OrderNo, Month, Season, V_Season, V_Year, V_PLM, V_BUYMONTH, V_TapeColor, V_Style, OrderDate, CompletedFlag "
            sql &= "Order by CustWavesCode, Month, OrderNo "
        End If
        '
        If DBuyer.Text = "TW0371T" Then
            sql = "SELECT "
            sql &= "Case When CompletedFlag = '1' Then '已出貨' else '未出貨' End As Completed, "
            sql &= "CustWavesCode, OrderNo, Month, Season, V_Season, V_Year, V_PLM, V_BUYMONTH, V_TapeColor, V_Style, OrderDate "
            sql &= "From I_UABAGS_ActualOrder "

            sql &= "Where Buyer = '" & DBuyer.Text & "' "
            sql &= "  And Month = '" & DYY.Text + DMM.Text & "' "
            sql &= "  And " & sqlstr & " "
            sql &= "  And ( "
            sql &= "        Select Count(*) From E_VDPErrorList "
            sql &= "        Where E_VDPErrorList.Buyer = '" & DBuyer.Text & "' "
            sql &= "          And E_VDPErrorList.OrderNo = dbo.I_UABAGS_ActualOrder.OrderNo "
            sql &= "          And E_VDPErrorList.Finished = '" & "Y" & "' "
            sql &= "      ) = 0 "
            sql &= "Group by CustWavesCode, OrderNo, Month, Season, V_Season, V_Year, V_PLM, V_BUYMONTH, V_TapeColor, V_Style, OrderDate, CompletedFlag "
            sql &= "Order by CustWavesCode, Month, OrderNo "
        End If
        '
        If DBuyer.Text = "TW1741" Then
            '無WAVE'S 實績, 不使用
        End If
        '
        Dim dt_ErrorList As DataTable = uDataBase.GetDataTable(sql)
        If dt_ErrorList.Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = dt_ErrorList
            GridView1.DataBind()
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
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            Dim sql As String
            Dim AlertColor As String = ""
            Dim AlertLevel As String = "00100"
            '
            ' 設有AlertLevel
            If DataBinder.Eval(e.Row.DataItem, "Season") = "" Or DataBinder.Eval(e.Row.DataItem, "V_Season") = "" Or _
               DataBinder.Eval(e.Row.DataItem, "V_Year") = 0 Or DataBinder.Eval(e.Row.DataItem, "V_PLM") = "" Then
                AlertLevel = "00100"
            Else
                AlertLevel = "00080"
            End If
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
                For i As Integer = 0 To 10
                    e.Row.Cells(i).BackColor = System.Drawing.ColorTranslator.FromHtml(AlertColor)
                Next
            End If

        End If
    End Sub

End Class
