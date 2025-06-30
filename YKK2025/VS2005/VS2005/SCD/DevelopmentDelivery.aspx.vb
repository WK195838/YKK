Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class DevelopmentDelivery
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDate As String           '現在日期
    Dim NowDateTime As String       '現在日期時間   yyyy/MM/dd HH:mm:ss
    Dim NowDateTime1 As String      '現在日期時間   yyyy/MM/dd HH:mm
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
            SetSearchField()                        '設定搜尋欄位
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
        Server.ScriptTimeout = 900                                                                  '設定逾時時間
        Response.Cookies("PGM").Value = "DevelopmentDelivery.aspx"                             '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDate = Now.ToString("yyyy/MM/dd")                '現在日期
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDateTime1 = Now.ToString("yyyy/MM/dd HH:mm")     '現在日期時間
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
    '**(SetSearchField)
    '**     設定搜尋欄位
    '**
    '*****************************************************************
    Sub SetSearchField()
        DNo.Text = ""
        DDevNo.Text = ""
        DCode.Text = ""
        DBuyer.Text = ""
        DOrderStart.Text = Now.Year.ToString + "/" + Now.Month.ToString + "/01"
        DOrderEnd.Text = NowDate
        DFinishStart.Text = Now.Year.ToString + "/" + Now.Month.ToString + "/01"
        DFinishEnd.Text = ""
        'DFinishEnd.Text = NowDate
        '
        Dim sql As String = ""
        Dim dtFieldData As DataTable
        DSizeNo.Items.Clear()
        DSizeNo.Items.Add("ALL")
        sql = "Select * From M_Referp Where Cat='2002' and DKey='SIZE' Order by Data "
        dtFieldData = uDataBase.GetDataTable(sql)
        For i As Integer = 0 To dtFieldData.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtFieldData.Rows(i).Item("Data")
            ListItem1.Value = dtFieldData.Rows(i).Item("Data")
            DSizeNo.Items.Add(ListItem1)
        Next
        '
        dtFieldData.Clear()
        '
        DItem.Items.Clear()
        DItem.Items.Add("ALL")
        sql = "Select * From M_Referp Where Cat='2002' and DKey='CHAIN' Order by Data "
        dtFieldData = uDataBase.GetDataTable(sql)
        For i As Integer = 0 To dtFieldData.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtFieldData.Rows(i).Item("Data")
            ListItem1.Value = dtFieldData.Rows(i).Item("Data")
            DItem.Items.Add(ListItem1)
        Next
        '現在日期時間
        DNowDateTime.Text = ""
        '至130工程
        DOP130.Text = ""
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BOK_Click)
    '**     OK按鈕按下事件
    '**
    '*****************************************************************
    Protected Sub BOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.Click

        NowDate = Now.ToString("yyyy/MM/dd")                  '現在日期
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日期時間
        NowDateTime1 = Now.ToString("yyyy/MM/dd HH:mm")       '現在日期時間
        '
        DNowDateTime.Text = Now.ToString("yyyy/MM/dd HH:mm") + " 現在"  '現在日期時間
        DOP130.Text = "* 至130工程"                           '至130工程  
        '
        '檢測輸入資料
        Dim ErrCode As Integer = 0
        Dim Message As String = ""
        '委託日
        If ErrCode = 0 Then
            If DOrderStart.Text <> "" Then
                If Not IsDate(DOrderStart.Text) Then ErrCode = 9010
            End If
            If DOrderEnd.Text <> "" Then
                If Not IsDate(DOrderEnd.Text) Then ErrCode = 9010
            End If
        End If
        '希望完成日
        If ErrCode = 0 Then
            If DFinishStart.Text <> "" Then
                If Not IsDate(DFinishStart.Text) Then ErrCode = 9020
            End If
            If DFinishEnd.Text <> "" Then
                If Not IsDate(DFinishEnd.Text) Then ErrCode = 9020
            End If
        End If
        '----異常訊息處理----------------------------------------------------
        If ErrCode <> 0 Then
            If ErrCode = 9010 Then Message = "委託日非日期格式,請確認!"
            If ErrCode = 9020 Then Message = "希望完成日非日期格式,請確認!"
            uJavaScript.PopMsg(Me, Message)
        Else
            ShowData()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowData)
    '**     篩選資料處理
    '**
    '*****************************************************************
    Sub ShowData()
        Dim sql As String = ""
        sql = "SELECT "
        ' FormNo(0), FormSno(1), Sts(2)
        sql &= "FormNo, FormSno, STS, "
        ' 狀態(3)
        sql &= "Case when Sts = 0 Then '開發中' When sts = 1 and  sts_norder =1 then '開發完成(停止受注)'  When sts = 1  Then '開發完成(OK)' When sts = 2 Then '開發完成(NG)'  Else '取消/中止'  End As STSD, "
        ' NO(4), 開發NO(5), CODE(6), BUYER(7), 型別(8), 鏈條(9)
        sql &= "NO, DEVNO, CODENO, AppBuyer AS BUYER, SIZENo AS SIZE, ITEM, "
        ' 委託日(10)
        sql &= "Convert(VarChar, AppDate, 111) AS DATE, "
        ' ViewOP(11)
        sql &= "'@' As ViewOP, "
        ' 希望完成(12)
        sql &= "Convert(VarChar, EXPDEL, 111) AS EDATE, "
        ' 實際完成日(13)
        sql &= "Convert(VarChar, CompletedTime, 111) AS CDATE, "
        ' 預定完成(14)
        sql &= "'' As BDate, "
        ' 實際完成(15)
        sql &= "'' As ADate, "
        ' 預定時間(16)
        sql &= "0 As BDays, "
        ' 實際時間(17)
        sql &= "0 As ADays, "
        ' 進行中工程(18), 進行中工程-預定完成(19)
        sql &= "'' AS OP, '' AS OPBDate, "
        ' NO(URL)(20), ViewOP(URL1)(21)
        sql &= "'' AS URL, '' AS URL1, APPPER,ORNO,Sellvendor  "
        '
        sql &= "FROM  V_CommissionSheet_02 "
        'Sts
        If DSts.SelectedValue <> "ALL" Then
            sql &= "WHERE STS = '" + DSts.SelectedValue + "' "
        Else
            sql &= "WHERE STS <> '9' "
        End If
        'NO
        If DNo.Text <> "" Then
            sql &= " And No LIKE '%" + DNo.Text + "%' "
        End If
        '開發No
        If DDevNo.Text <> "" Then
            sql &= " And DevNo LIKE '%" + DDevNo.Text + "%' "
        End If
        'CodeNo
        If DCode.Text <> "" Then
            sql &= " And CodeNo LIKE '%" + DCode.Text + "%' "
        End If
        'ORNO
        If DORNO.Text <> "" Then
            sql &= " And ORNO LIKE '%" + DORNO.Text + "%' "
        End If
        'Buyer
        If DBuyer.Text <> "" Then
            sql &= " And AppBuyer LIKE '%" + DBuyer.Text + "%' "
        End If
        'Size
        If DSizeNo.SelectedValue <> "ALL" Then
            sql &= " And SizeNo = '" + DSizeNo.SelectedValue + "'"
        End If
        'Chain
        If DItem.SelectedValue <> "ALL" Then
            sql &= " And Item = '" + DItem.SelectedValue + "'"
        End If
        ''委託日
        If DOrderStart.Text <> "" And DOrderEnd.Text <> "" Then
            sql &= " And AppDate >= '" + CDate(DOrderStart.Text).ToString("yyyy/MM/dd") + " 00:00:00" + "' "
            sql &= " And AppDate <= '" + CDate(DOrderEnd.Text).ToString("yyyy/MM/dd") + " 23:59:59" + "' "
        End If
        ''希望完成日
        If DFinishStart.Text <> "" And DFinishEnd.Text <> "" Then
            sql &= " And ExpDel >= '" + CDate(DFinishStart.Text).ToString("yyyy/MM/dd") + " 00:00:00" + "' "
            sql &= " And ExpDel <= '" + CDate(DFinishEnd.Text).ToString("yyyy/MM/dd") + " 23:59:59" + "' "
        End If
        '委託廠商
        If DSellvendor.Text <> "" Then
            sql &= " And Sellvendor LIKE '%" + DSellvendor.Text + "%' "
        End If

        sql &= "ORDER BY CreateTime Desc "
        Dim dtCommissionSheet As DataTable = uDataBase.GetDataTable(sql)
        '設定各欄位資料
        If dtCommissionSheet.Rows.Count > 0 Then
            '
            For i As Integer = 0 To dtCommissionSheet.Rows.Count - 1
                ' ViewOP Mark [@] (11)
                dtCommissionSheet.Rows(i)(11) = GetViewOPMark(dtCommissionSheet.Rows(i).Item("FormNo"), _
                                                               dtCommissionSheet.Rows(i).Item("FormSno"))
                ' 實際完成日(13)
                If dtCommissionSheet.Rows(i)(3) = "開發中" Then
                    dtCommissionSheet.Rows(i)(13) = ""
                End If
                ' BDate 預定完成(14)
                dtCommissionSheet.Rows(i)(14) = GetDeliverDate(dtCommissionSheet.Rows(i).Item("FormNo"), _
                                                               dtCommissionSheet.Rows(i).Item("FormSno"), _
                                                               130, _
                                                               9)
                ' ADate 實際完成(15) 
                dtCommissionSheet.Rows(i)(15) = GetDeliverDate(dtCommissionSheet.Rows(i).Item("FormNo"), _
                                                               dtCommissionSheet.Rows(i).Item("FormSno"), _
                                                               130, _
                                                               1)
                ' BDays 預定時間(16)
                If dtCommissionSheet.Rows(i)(14) <> "" Then
                    Dim xStartDate As DateTime = CDate(dtCommissionSheet.Rows(i)(10).ToString)
                    Dim xEndDate As DateTime = CDate(dtCommissionSheet.Rows(i)(14).ToString)
                    Dim xDays As Integer = DateDiff(DateInterval.Day, xStartDate, xEndDate)
                    Dim xVacationDays As Integer = GetVacationDays("TP1", xStartDate, xEndDate)
                    '
                    dtCommissionSheet.Rows(i)(16) = xDays - xVacationDays
                End If
                ' ADays 實際時間(17)
                If dtCommissionSheet.Rows(i)(15) <> "" Then
                    Dim xStartDate As DateTime = CDate(dtCommissionSheet.Rows(i)(10).ToString)
                    Dim xEndDate As DateTime = CDate(dtCommissionSheet.Rows(i)(15).ToString)
                    Dim xDays As Integer = DateDiff(DateInterval.Day, xStartDate, xEndDate)
                    Dim xVacationDays As Integer = GetVacationDays("TP1", xStartDate, xEndDate)
                    '
                    dtCommissionSheet.Rows(i)(17) = xDays - xVacationDays
                End If
                ' 進行中工程(18), 進行中工程-預定完成(19)
                If dtCommissionSheet.Rows(i).Item("STS") = 0 Then
                    '開發中
                    sql = "SELECT Step, StepName, DecideName, BEndTime From V_WaitHandle_01 "
                    sql &= " Where FormNo = '" & dtCommissionSheet.Rows(i).Item("FormNo") & "'"
                    sql &= "   And FormSno = '" & CStr(dtCommissionSheet.Rows(i).Item("FormSno")) & "'"
                    sql &= "   And Active = '1' "
                    sql &= "   And FlowType <> '0' "
                    Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(sql)
                    If dtWaitHandle.Rows.Count > 0 Then
                        dtCommissionSheet.Rows(i)(18) = "[" + dtWaitHandle.Rows(0).Item("Step").ToString + "]-" + _
                                                        Replace(dtWaitHandle.Rows(0).Item("StepName"), "開發委託_", "") + _
                                                        "(" + dtWaitHandle.Rows(0).Item("DecideName") + ")"
                        dtCommissionSheet.Rows(i)(19) = CDate(dtWaitHandle.Rows(0).Item("BEndTime")).ToString("yyyy/MM/dd HH:mm")
                    End If
                End If
                ' URL NO(20)
                dtCommissionSheet.Rows(i)(20) = "CommissionSheet_02.aspx?pFormNo=" + RTrim(dtCommissionSheet.Rows(i).Item("FormNo")) + "&pFormSno=" + CStr(dtCommissionSheet.Rows(i).Item("FormSno"))
                ' URL1 ViewOP(21)
                dtCommissionSheet.Rows(i)(21) = "DevelopmentDelivery_OP.aspx?pFormNo=" + RTrim(dtCommissionSheet.Rows(i).Item("FormNo")) + "&pFormSno=" + CStr(dtCommissionSheet.Rows(i).Item("FormSno"))
            Next
            '
        End If
        '
        GridView1.DataSource = dtCommissionSheet
        GridView1.DataBind()
        '
        GridView2.DataSource = dtCommissionSheet
        GridView2.DataBind()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetVacationDays)
    '**     取得休假日數
    '**    
    '*****************************************************************
    Public Function GetViewOPMark(ByVal pFormNo As String, ByVal pFormSno As Integer) As String
        Dim RtnString As String = ""
        Dim SQL As String
        Try
            SQL = "Select * From T_WaitHandle "
            SQL &= "Where FormNo  = '" & pFormNo & "' "
            SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
            SQL &= "  And Step    >= '40' "
            SQL &= "Order by Step "
            Dim dt_WaitHandle = uDataBase.GetDataTable(SQL)
            If dt_WaitHandle.Rows.Count > 0 Then
                RtnString = "@"
            End If
        Catch ex As Exception
        End Try
        '
        Return RtnString
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetDeliverDate)
    '**     取得預定完成日
    '**     GetDeliverDate(表單, 單號, 工程, 預定(9) / 實際(1))
    '*****************************************************************
    Public Function GetDeliverDate(ByVal pFormNo As String, _
                                   ByVal pFormSno As Integer, _
                                   ByVal pStep As Integer, _
                                   ByVal pDateType As Integer) As String
        Dim RtnString As String = ""
        Dim SQL As String
        Try
            If pDateType = 9 Then
                SQL = "Select BEndTime From T_WaitHandle "
                SQL &= "Where FormNo  = '" & pFormNo & "' "
                SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                SQL &= "  And Step    = '" & CStr(pStep) & "' "
                'SQL &= "  And FlowType = '9' "
                SQL &= "Order by BEndTime Desc "
                Dim dt_WaitHandle = uDataBase.GetDataTable(SQL)
                If dt_WaitHandle.Rows.Count > 0 Then
                    RtnString = CDate(dt_WaitHandle.Rows(0).Item("BEndTime")).ToString("yyyy/MM/dd")
                End If
            Else
                SQL = "Select AEndTime From T_WaitHandle "
                SQL &= "Where FormNo  = '" & pFormNo & "' "
                SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                SQL &= "  And Step    = '" & CStr(pStep) & "' "
                SQL &= "  And FlowType = '1' "
                SQL &= "Order by AEndTime Desc "
                Dim dt_WaitHandle = uDataBase.GetDataTable(SQL)
                If dt_WaitHandle.Rows.Count > 0 Then
                    RtnString = CDate(dt_WaitHandle.Rows(0).Item("AEndTime")).ToString("yyyy/MM/dd")
                End If
            End If
        Catch ex As Exception
        End Try
        '
        Return RtnString
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetVacationDays)
    '**     取得休假日數
    '**    
    '*****************************************************************
    Public Function GetVacationDays(ByVal pDepo As String, _
                                    ByVal pStartDate As String, _
                                    ByVal pEndDate As String) As Integer
        Dim RtnDays As Integer = 0
        Dim SQL As String
        Try
            SQL = "Select Count(*) As RCount From M_Vacation "
            SQL &= "Where Active = '1' "
            SQL &= "  And Depo  = '" & pDepo & "' "
            SQL &= "  And YMD  >= '" & pStartDate & "' "
            SQL &= "  And YMD  <= '" & pEndDate & "' "
            Dim dt_Vacation = uDataBase.GetDataTable(SQL)
            If dt_Vacation.Rows.Count > 0 Then
                RtnDays = dt_Vacation.Rows(0).Item("RCount")
            End If
        Catch ex As Exception
        End Try
        '
        Return RtnDays
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GridView1_PageIndexChanging)
    '**     換頁處理
    '**
    '*****************************************************************
    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub
    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        NowDate = Now.ToString("yyyy/MM/dd")                  '現在日期
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日期時間
        DNowDateTime.Text = Now.ToString("yyyy/MM/dd HH:mm") + " 現在"  '現在日期時間
        DOP130.Text = "* 至130工程"                           '至130工程  
        '
        ShowData()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GridView1_RowDataBound)
    '**     延遲處理
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            'Dim xBackColor As Boolean = False
            '--開發中--------------------------------------------
            If e.Row.Cells(0).Text.ToString = "開發中" Then
                '工程預定完成日期(17) < 現在日期時間-->顯示紅字/綠底
                If e.Row.Cells(17).Text.ToString <> "&nbsp;" Then
                    If e.Row.Cells(17).Text.ToString < NowDateTime1 Then
                        e.Row.Cells(17).ForeColor = Color.Red
                        e.Row.Cells(17).BackColor = Color.GreenYellow
                        'If Not xBackColor Then e.Row.BackColor = Color.GreenYellow
                    End If
                End If
                '預定完成(12) < 實際完成(13)-->顯示紅字/黃底
                If e.Row.Cells(12).Text.ToString <> "&nbsp;" And e.Row.Cells(13).Text.ToString <> "&nbsp;" Then
                    If e.Row.Cells(12).Text.ToString < e.Row.Cells(13).Text.ToString Then
                        e.Row.Cells(12).ForeColor = Color.Red
                        e.Row.Cells(13).ForeColor = Color.Red
                        e.Row.Cells(12).BackColor = Color.Yellow
                        e.Row.Cells(13).BackColor = Color.Yellow
                        'e.Row.BackColor = Color.Yellow
                    End If
                Else
                    '預定完成(12) < 現在日期-->顯示紅字/黃底
                    If e.Row.Cells(12).Text.ToString <> "&nbsp;" Then
                        If e.Row.Cells(12).Text.ToString < NowDate Then
                            e.Row.Cells(12).ForeColor = Color.Red
                            e.Row.Cells(12).BackColor = Color.Yellow
                            'e.Row.BackColor = Color.Yellow
                            'xBackColor = True
                        End If
                    End If
                End If
                '希望完成(10) < 實際完成(11)-->顯示紅字/粉紅底
                If e.Row.Cells(10).Text.ToString <> "&nbsp;" And e.Row.Cells(11).Text.ToString <> "&nbsp;" Then
                    If e.Row.Cells(10).Text.ToString < e.Row.Cells(11).Text.ToString Then
                        e.Row.Cells(10).ForeColor = Color.Red
                        e.Row.Cells(11).ForeColor = Color.Red
                        e.Row.Cells(10).BackColor = Color.LightPink
                        e.Row.Cells(11).BackColor = Color.LightPink
                        'e.Row.BackColor = Color.LightPink
                    End If
                Else
                    '希望完成(10) < 現在日期時間-->顯示紅字/粉紅底
                    If e.Row.Cells(10).Text.ToString <> "&nbsp;" Then
                        If e.Row.Cells(10).Text.ToString < NowDate Then
                            e.Row.Cells(10).ForeColor = Color.Red
                            e.Row.Cells(10).BackColor = Color.LightPink
                            'e.Row.BackColor = Color.LightPink
                        End If
                    End If
                End If
            End If
            '--開發完成--------------------------------------------
            If e.Row.Cells(0).Text.ToString = "開發完成(OK)" Then
                '預定完成(12) < 實際完成(13)-->顯示紅字/黃底
                If e.Row.Cells(12).Text.ToString <> "&nbsp;" And e.Row.Cells(13).Text.ToString <> "&nbsp;" Then
                    If e.Row.Cells(12).Text.ToString < e.Row.Cells(13).Text.ToString Then
                        e.Row.Cells(12).ForeColor = Color.Red
                        e.Row.Cells(13).ForeColor = Color.Red
                        e.Row.Cells(12).BackColor = Color.Yellow
                        e.Row.Cells(13).BackColor = Color.Yellow
                        'e.Row.BackColor = Color.Yellow
                    End If
                Else
                    '預定完成(12) < 現在日期-->顯示紅字/黃底
                    If e.Row.Cells(12).Text.ToString <> "&nbsp;" Then
                        If e.Row.Cells(12).Text.ToString < NowDate Then
                            e.Row.Cells(12).ForeColor = Color.Red
                            e.Row.Cells(12).BackColor = Color.Yellow
                            'e.Row.BackColor = Color.Yellow
                            'xBackColor = True
                        End If
                    End If
                End If
                '
                '希望完成(10) < 實際完成(11)-->顯示紅字/粉紅底
                If e.Row.Cells(10).Text.ToString <> "&nbsp;" And e.Row.Cells(11).Text.ToString <> "&nbsp;" Then
                    If e.Row.Cells(10).Text.ToString < e.Row.Cells(11).Text.ToString Then
                        e.Row.Cells(10).ForeColor = Color.Red
                        e.Row.Cells(11).ForeColor = Color.Red
                        e.Row.Cells(10).BackColor = Color.LightPink
                        e.Row.Cells(11).BackColor = Color.LightPink
                        'e.Row.BackColor = Color.LightPink
                    End If
                Else
                    '希望完成(10) < 現在日期時間-->顯示紅字/粉紅底
                    If e.Row.Cells(10).Text.ToString <> "&nbsp;" Then
                        If e.Row.Cells(10).Text.ToString < NowDate Then
                            e.Row.Cells(10).ForeColor = Color.Red
                            e.Row.Cells(10).BackColor = Color.LightPink
                            'e.Row.BackColor = Color.LightPink
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Response.AppendHeader("Content-Disposition", "attachment;filename=DeliveryList.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)

        GridView2.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()
    End Sub

End Class
