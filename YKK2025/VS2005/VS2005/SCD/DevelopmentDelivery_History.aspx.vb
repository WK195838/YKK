Imports System.Data
Imports System.Drawing

Partial Class DevelopmentDelivery_History
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wFormNo As String
    Dim wFormSno As Integer
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim oCommon As New Common.CommonService
    Dim oSchedule As New Schedule.ScheduleService
    '
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.Cookies("PGM").Value = "DevelopmentDelivery_History.aspx"
        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            DataList()
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        wFormNo = Request.QueryString("pFormNo")
        wFormSno = Request.QueryString("pFormSno")
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選資料
    '**
    '*****************************************************************
    Sub DataList()
        Dim sql As String
        '
        '-----------------------------------------------------------------
        '-- 核定履歷
        '-----------------------------------------------------------------
        sql = "SELECT "
        sql = sql + "FormNo, FormSno, Step, SeqNo, "
        sql = sql + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
        sql = sql + "'' as Description, BStartTimeDesc, BEndTimeDesc, AStartTimeDesc, AEndTimeDesc, WorkID, "
        sql = sql + "'' As ViewOP, '' AS URL, Active, No, "
        sql = sql + "Case Active When 9 Then 'P' When 1 Then 'W' Else '' End As STSD "
        sql = sql + "FROM V_WaitHandle_01 "
        sql = sql + "Where FormNo  = '" & wFormNo & "' "
        sql = sql + "  And FormSno = '" & CStr(wFormSno) & "' "
        sql = sql + "  And (FlowType = '1' or FlowType='2' or FlowType='9' or Step=999)  "
        sql = sql + "Order by Active Desc, Unique_ID Desc "
        Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(sql)
        If dt_WaitHandle.Rows.Count > 0 Then
            For i As Integer = 0 To dt_WaitHandle.Rows.Count - 1
                ' StepNameDesc(9)
                dt_WaitHandle.Rows(i)(9) = "[" + dt_WaitHandle.Rows(i)(2).ToString + "]-" + Replace(dt_WaitHandle.Rows(i)(9), "開發委託_", "")
                '
                ' Description(12)
                ' 預：[2012-04-25 16:00 ~ 2012-04-25 17:00](380),
                ' 實：[2012-04-25 16:00 ~ 2012-04-26 14:10](380)
                Dim xTime As Integer = 0
                ' 預
                xTime = 0
                If InStr(dt_WaitHandle.Rows(i)(14), "無記錄") > 0 Then
                Else
                    dt_WaitHandle.Rows(i)(12) = "預：[" + CDate(dt_WaitHandle.Rows(i)(13)).ToString("yyyy/MM/dd HH:mm") + " ~ " + CDate(dt_WaitHandle.Rows(i)(14)).ToString("yyyy/MM/dd HH:mm")
                    oSchedule.GetDevelopTime(CDate(dt_WaitHandle.Rows(i)(13)).ToString("yyyy/MM/dd HH:mm"), _
                                             CDate(dt_WaitHandle.Rows(i)(14)).ToString("yyyy/MM/dd HH:mm"), _
                                             xTime, _
                                             oCommon.GetCalendarGroup(dt_WaitHandle.Rows(i)(17)))
                    dt_WaitHandle.Rows(i)(12) = dt_WaitHandle.Rows(i)(12) + "  (" + CStr(xTime) + ")]"
                End If
                ' 實
                xTime = 0
                If InStr(dt_WaitHandle.Rows(i)(16), "無記錄") > 0 Then
                Else
                    dt_WaitHandle.Rows(i)(12) = dt_WaitHandle.Rows(i)(12) + ", " + _
                                                "實：[" + CDate(dt_WaitHandle.Rows(i)(15)).ToString("yyyy/MM/dd HH:mm") + " ~ " + CDate(dt_WaitHandle.Rows(i)(16)).ToString("yyyy/MM/dd HH:mm")
                    oSchedule.GetDevelopTime(CDate(dt_WaitHandle.Rows(i)(15)).ToString("yyyy/MM/dd HH:mm"), _
                                             CDate(dt_WaitHandle.Rows(i)(16)).ToString("yyyy/MM/dd HH:mm"), _
                                             xTime, _
                                             oCommon.GetCalendarGroup(dt_WaitHandle.Rows(i)(17)))
                    dt_WaitHandle.Rows(i)(12) = dt_WaitHandle.Rows(i)(12) + "  (" + CStr(xTime) + ")]"
                End If
            Next
        End If
        GridView1.DataSource = dt_WaitHandle
        GridView1.DataBind()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GridView1_RowDataBound)
    '**     延遲處理
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim Pos As Integer = InStr(e.Row.Cells(4).Text.ToString, "],")
            If Pos > 0 Then
                Dim Str1 As String = Mid(e.Row.Cells(4).Text.ToString, 1, Pos)
                Dim Str2 As String = Mid(e.Row.Cells(4).Text.ToString, Pos + 3, Len(e.Row.Cells(4).Text.ToString))
                e.Row.Cells(4).Text = Str1 + "<br/>" + Str2
            End If
        End If
    End Sub


End Class
