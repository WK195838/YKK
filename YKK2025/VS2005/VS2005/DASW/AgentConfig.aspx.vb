Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class AgentConfig
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDate As String           '現在日期
    Dim NowDateTime As String       '現在日期時間
    Dim UserID As String            '使用者ID
    Dim Calendar As String = ""     '行事曆類別
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
            SetUserInformation()
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
        Server.ScriptTimeout = 900                                              '設定逾時時間
        Response.Cookies("PGM").Value = "AgentConfig.aspx"                      '程式名
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDate = Now.ToString("yyyy/MM/dd")                  '現在日期
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日期時間
        UserID = Request.QueryString("pUserID")               '使用者ID  
        Calendar = oCommon.GetCalendarGroup(UserID)           '取得行事曆類別
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        BFromDate.Attributes("onclick") = "calendarPicker('" + Calendar + "','DFromDate');"  '起始日-選擇輸入
        BToDate.Attributes("onclick") = "calendarPicker('" + Calendar + "','DToDate');"      '截止日-選擇輸入
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetUserInformation)
    '**     設定個人資訊
    '**
    '*****************************************************************
    Sub SetUserInformation()
        '欄位屬性
        DName.ReadOnly = True
        DUserID.ReadOnly = True
        DLastDate.ReadOnly = True
        DLastTime.ReadOnly = True
        DAgentName.ReadOnly = True
        DAgentUserID.ReadOnly = True
        DFromDate.ReadOnly = True
        DToDate.ReadOnly = True
        '欄位內容
        DName.Text = oCommon.GetUserName(UserID)
        DUserID.Text = UserID
        DLastDate.Text = ""
        DLastTime.Text = ""
        DAgentName.Text = ""
        DAgentUserID.Text = ""
        DFromDate.Text = ""
        DToDate.Text = ""
        DDescription.Text = ""
        BSave.Visible = False
        '
        Dim sql As String = ""
        Dim dt As DataTable
        sql = "Select * From M_Agent "
        sql &= "Where UserID = '" + UserID + "' "
        sql = sql + " and formno = '006001'"
        dt = uDataBase.GetDataTable(sql)
        If dt.Rows.Count > 0 Then
            DLastDate.Text = CDate(dt.Rows(0).Item("ModifyTime")).ToString("yyyy/MM/dd")
            DLastTime.Text = CDate(dt.Rows(0).Item("ModifyTime")).ToString("HH:mm:ss")
            DAgentName.Text = dt.Rows(0).Item("AgentName")
            DAgentUserID.Text = dt.Rows(0).Item("AgentID")
            DFromDate.Text = CDate(dt.Rows(0).Item("StartDate")).ToString("yyyy/MM/dd")
            DFromHH.Text = CDate(dt.Rows(0).Item("StartDate")).ToString("HH")
            DFromMM.Text = CDate(dt.Rows(0).Item("StartDate")).ToString("mm")
            DFromSS.Text = CDate(dt.Rows(0).Item("StartDate")).ToString("ss")
            DToDate.Text = CDate(dt.Rows(0).Item("EndDate")).ToString("yyyy/MM/dd")
            DToHH.Text = CDate(dt.Rows(0).Item("EndDate")).ToString("HH")
            DToMM.Text = CDate(dt.Rows(0).Item("EndDate")).ToString("mm")
            DToSS.Text = CDate(dt.Rows(0).Item("EndDate")).ToString("ss")
            BSave.Visible = True
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BSave_Click)
    '**     啟動代理人設定
    '**
    '*****************************************************************
    Protected Sub BSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSave.Click
        Dim ErrCode As Integer = 0
        Dim Message As String = ""
        'Input Data Check
        '起迄日
        If ErrCode = 0 Then
            If CDate(DFromDate.Text) > CDate(DToDate.Text) Then
                ErrCode = 9100
            Else
                If CDate(DFromDate.Text) = CDate(DToDate.Text) And _
                   DFromHH.SelectedValue + DFromMM.SelectedValue + DFromSS.SelectedValue >= DToHH.SelectedValue + DToMM.SelectedValue + DToSS.SelectedValue Then
                    ErrCode = 9200
                End If
            End If
        End If
        '說明
        If ErrCode = 0 Then
            If DDescription.Text = "" Then ErrCode = 9900
        End If
        '異常訊息處理
        If ErrCode <> 0 Then
            If ErrCode = 9100 Then Message = "起迄日期錯誤,請確認!"
            If ErrCode = 9200 Then Message = "起迄日期錯誤,請確認!"
            If ErrCode = 9900 Then Message = "說明需輸入,請確認!"
            uJavaScript.PopMsg(Me, Message)
        End If
        '資料儲存
        If ErrCode = 0 Then
            Try
                Dim sql As String = ""
                Dim dt As DataTable
                Dim StartDate, EndDate As String
                sql = "Select * From M_Agent "
                sql &= "Where UserID = '" + UserID + "' "
                sql &= " and formno = '006001'"
                dt = uDataBase.GetDataTable(Sql)
                If dt.Rows.Count > 0 Then
                    '起迄日重新編成
                    If DFromHH.SelectedValue + DFromMM.SelectedValue + DFromSS.SelectedValue = "000000" Then
                        StartDate = DFromDate.Text + " " + DFromHH.SelectedValue + ":" + DFromMM.SelectedValue + ":" + "10"
                    Else
                        StartDate = DFromDate.Text + " " + DFromHH.SelectedValue + ":" + DFromMM.SelectedValue + ":" + DFromSS.SelectedValue
                    End If
                    If DToHH.SelectedValue + DToMM.SelectedValue + DToSS.SelectedValue = "000000" Then
                        EndDate = DToDate.Text + " " + DToHH.SelectedValue + ":" + DToMM.SelectedValue + ":" + "10"
                    Else
                        EndDate = DToDate.Text + " " + DToHH.SelectedValue + ":" + DToMM.SelectedValue + ":" + DToSS.SelectedValue
                    End If
                    '更新資料
                    sql = "Update M_Agent Set "
                    sql = sql + " AgentID = '" & DAgentUserID.Text & "',"
                    sql = sql + " AgentName = '" & DAgentName.Text & "',"
                    sql = sql + " StartDate = '" & StartDate & "',"
                    sql = sql + " EndDate = '" & EndDate & "',"
                    sql = sql + " Description = '" & DDescription.Text & "',"
                    sql = sql + " ModifyUser = '" & UserID & "',"
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where UserID = '" & UserID & "'"
                    sql = sql + " and formno = '006001'"
                    uDataBase.ExecuteNonQuery(sql)
                    '更新後重新顯示
                    sql = "Select * From M_Agent "
                    sql &= "Where UserID = '" + UserID + "' "
                    sql &= " and formno = '006001'"
                    dt = uDataBase.GetDataTable(sql)
                    If dt.Rows.Count > 0 Then
                        DLastDate.Text = CDate(dt.Rows(0).Item("ModifyTime")).ToString("yyyy/MM/dd")
                        DLastTime.Text = CDate(dt.Rows(0).Item("ModifyTime")).ToString("HH:mm:ss")
                        DAgentName.Text = dt.Rows(0).Item("AgentName")
                        DAgentUserID.Text = dt.Rows(0).Item("AgentID")
                        DFromDate.Text = CDate(dt.Rows(0).Item("StartDate")).ToString("yyyy/MM/dd")
                        DFromHH.Text = CDate(dt.Rows(0).Item("StartDate")).ToString("HH")
                        DFromMM.Text = CDate(dt.Rows(0).Item("StartDate")).ToString("mm")
                        DFromSS.Text = CDate(dt.Rows(0).Item("StartDate")).ToString("ss")
                        DToDate.Text = CDate(dt.Rows(0).Item("EndDate")).ToString("yyyy/MM/dd")
                        DToHH.Text = CDate(dt.Rows(0).Item("EndDate")).ToString("HH")
                        DToMM.Text = CDate(dt.Rows(0).Item("EndDate")).ToString("mm")
                        DToSS.Text = CDate(dt.Rows(0).Item("EndDate")).ToString("ss")
                        BSave.Visible = True
                    End If
                    '
                    uJavaScript.PopMsg(Me, "代理人設定啟動完成")
                End If
            Catch ex As Exception
            End Try

        End If
    End Sub
End Class
