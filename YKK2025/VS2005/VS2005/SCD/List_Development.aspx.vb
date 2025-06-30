Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing


Partial Class List_Development
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDate As String           '現在日期
    Dim NowDateTime As String       '現在日期時間
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
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
        Response.Cookies("PGM").Value = "List_Development.aspx"                                     '程式名
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")                      '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")                            '工程代碼
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDate = Now.ToString("yyyy/MM/dd")                  '現在日期
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日期時間
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
        DNowDateTime.Text = Now.ToString("yyyy/MM/dd HH:mm") + " 現在"  '現在日期時間
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
        sql &= "FormNo, FormSno, STS, NO, DEVNO, CODENO, ITEM, "     '0~6
        sql &= "Case Sts When 0 Then '開發中' When 1 Then '開發完成(OK)' When 2 Then '開發完成(NG)' Else '取消/中止' End As STSD, "  '7
        sql &= "Convert(VarChar, AppDate, 111) AS Date, "      '8
        sql &= "AppBuyer AS BUYER, "    '9
        sql &= "Convert(VarChar, EXPDEL, 111) AS EDATE, "      '10
        sql &= "SIZENo AS SIZE, "       '11
        sql &= "Case Sts When 0 Then null "     '12
        sql &= "         Else Convert(VarChar, CompletedTime, 111) + ' ' + Convert(VarChar, CompletedTime, 108) "
        sql &= "End As CDATE, "          '12
        sql &= "'' AS OP, "             '13
        sql &= "'' AS BDATE, "          '14    
        sql &= "'' AS URL "             '15
        sql &= "FROM V_CommissionSheet_02 "
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
        '委託日
        If DOrderStart.Text <> "" And DOrderEnd.Text <> "" Then
            sql &= " And AppDate >= '" + CDate(DOrderStart.Text).ToString("yyyy/MM/dd") + " 00:00:00" + "' "
            sql &= " And AppDate <= '" + CDate(DOrderEnd.Text).ToString("yyyy/MM/dd") + " 23:59:59" + "' "
        End If
        '希望完成日
        If DFinishStart.Text <> "" And DFinishEnd.Text <> "" Then
            sql &= " And ExpDel >= '" + CDate(DFinishStart.Text).ToString("yyyy/MM/dd") + " 00:00:00" + "' "
            sql &= " And ExpDel <= '" + CDate(DFinishEnd.Text).ToString("yyyy/MM/dd") + " 23:59:59" + "' "
        End If
        sql &= "ORDER BY CreateTime Desc "
        Dim dtCommissionSheet As DataTable = uDataBase.GetDataTable(sql)
        If dtCommissionSheet.Rows.Count > 0 Then
            '
            For i As Integer = 0 To dtCommissionSheet.Rows.Count - 1
                Dim xOP As String = ""
                Dim xBDate As String = ""
                Dim xURL As String = ""
                '
                If dtCommissionSheet.Rows(i).Item("STS") = 0 Then
                    '開發中
                    sql = "SELECT StepName, DecideName, BEndTime, ViewURL From V_WaitHandle_01 "
                    sql &= " Where FormNo = '" & dtCommissionSheet.Rows(i).Item("FormNo") & "'"
                    sql &= "   And FormSno = '" & CStr(dtCommissionSheet.Rows(i).Item("FormSno")) & "'"
                    sql &= "   And Active = '1' "
                    sql &= "   And FlowType <> '0' "
                    Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(sql)
                    If dtWaitHandle.Rows.Count > 0 Then
                        xOP = dtWaitHandle.Rows(0).Item("StepName") + "(" + dtWaitHandle.Rows(0).Item("DecideName") + ")"
                        xBDate = CDate(dtWaitHandle.Rows(0).Item("BEndTime")).ToString("yyyy/MM/dd HH:mm:ss")
                        xURL = dtWaitHandle.Rows(0).Item("ViewURL")
                    Else
                        xURL = "CommissionSheet_02.aspx?pFormNo=" + RTrim(dtCommissionSheet.Rows(i).Item("FormNo")) + "&pFormSno=" + CStr(dtCommissionSheet.Rows(i).Item("FormSno"))
                    End If
                Else
                    'OK/NG/取消
                    sql = "SELECT ViewURL From V_WaitHandle_01 "
                    sql &= " Where FormNo = '" & dtCommissionSheet.Rows(i).Item("FormNo") & "'"
                    sql &= "   And FormSno = '" & CStr(dtCommissionSheet.Rows(i).Item("FormSno")) & "'"
                    sql &= "   And Step = '1' "
                    sql &= "   And FlowType <> '0' "
                    Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(sql)
                    If dtWaitHandle.Rows.Count > 0 Then
                        xURL = dtWaitHandle.Rows(0).Item("ViewURL")
                    Else
                        xURL = "CommissionSheet_02.aspx?pFormNo=" + RTrim(dtCommissionSheet.Rows(i).Item("FormNo")) + "&pFormSno=" + CStr(dtCommissionSheet.Rows(i).Item("FormSno"))
                    End If
                End If
                '
                dtCommissionSheet.Rows(i)(13) = xOP
                If xBDate <> "" Then dtCommissionSheet.Rows(i)(14) = xBDate
                dtCommissionSheet.Rows(i)(15) = xURL
            Next
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
            Dim xBackColor As Boolean = False
            '--開發中--------------------------------------------
            If e.Row.Cells(0).Text.ToString = "開發中" Then
                '希望完成 < 現在日期-->顯示紅字/粉紅底
                If e.Row.Cells(8).Text.ToString < NowDate Then
                    e.Row.Cells(8).ForeColor = Color.Red
                    e.Row.BackColor = Color.LightPink
                    xBackColor = True
                End If
                '工程預定完成日期時間 < 現在日期時間-->顯示紅字/藍底
                If e.Row.Cells(11).Text.ToString < NowDateTime Then
                    e.Row.Cells(11).ForeColor = Color.Red
                    If Not xBackColor Then e.Row.BackColor = Color.LightGreen
                End If
            End If
            '--開發完成--------------------------------------------
            If e.Row.Cells(0).Text.ToString = "完成" Then
                '希望完成 < 實際完成  
                If e.Row.Cells(8).Text.ToString + " 23:59:59" < e.Row.Cells(9).Text.ToString Then
                    e.Row.Cells(8).ForeColor = Color.Red
                    e.Row.Cells(9).ForeColor = Color.Red
                    e.Row.BackColor = Color.LightPink
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
        Response.AppendHeader("Content-Disposition", "attachment;filename=SCD_DevelopmentList.xls")     '程式別不同
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
