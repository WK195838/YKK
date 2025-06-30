Imports System.Data

Partial Class Fix_BEndTime
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim oCommon As New Common.CommonService
    Dim oSchedule As New Schedule.ScheduleService

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900      ' 設定逾時時間
        If Not IsPostBack Then
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(FlowControl)
    '**     簽核處理
    '**
    '*****************************************************************
    Sub FlowControl()
        Dim SQL As String = ""
        Dim xID As Integer = 0

        Dim pFormNo As String = ""
        Dim pNextStep As Integer = 0
        Dim wLevel As String = ""
        Dim wQCLT As Integer = 0
        Dim pLastTime, pStartTime, pEndTime As DateTime
        Dim wNextGateCalendar As String = ""
        Dim pNextGate As String = ""
        Dim NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")

        SQL = "Select * From T_WaitHandle "
        SQL = SQL & "Where BStartTime = BEndTime "
        SQL = SQL & "  And Step >= 10 "
        SQL = SQL & "  And Step <> 999 "
        SQL = SQL & "  And CreateTime >= '" & "2010/6/18 00:00:00" & "' "
        SQL = SQL & "Order by FormNo, FormSno "
        Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(SQL)
        For i As Integer = 0 To dtWaitHandle.Rows.Count - 1
            xID = dtWaitHandle.Rows(i)("Unique_ID")
            pFormNo = dtWaitHandle.Rows(i)("FormNo").ToString
            pNextStep = dtWaitHandle.Rows(i)("Step")
            wLevel = ""
            wQCLT = 0
            pLastTime = dtWaitHandle.Rows(i)("BStartTime")
            pStartTime = Now
            pEndTime = Now
            pNextGate = dtWaitHandle.Rows(i)("DecideID").ToString
            wNextGateCalendar = oCommon.GetCalendarGroup(pNextGate)

            '取得預定開始,完成日程計算(表單號碼,工程號碼,難易度,QC-L/T,現在時間, 預定開始日時, 預定完成日時, 行事曆)
            oSchedule.GetLeadTime(pFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wNextGateCalendar)

            '
            SQL = "Update T_WaitHandle Set "            '更新實際開始日
            SQL &= " BEndTime = '" & pEndTime.ToString("yyyy/MM/dd HH:mm:ss") & "', "
            SQL &= " ModifyUser = '" & "Fix_BEndTime" & "', "
            SQL &= " ModifyTime = '" & NowDateTime & "' "
            SQL &= "Where Unique_ID = '" & xID & "' "
            uDataBase.ExecuteNonQuery(SQL)
        Next
    End Sub

    Function GetLeadTime(ByVal pFormNo As String, _
                                ByVal pStep As Integer, _
                                ByVal pLevel As String, _
                                ByVal pQCLT As Integer, _
                                ByVal pNowDateTime As DateTime, _
                                ByRef pStartTime As DateTime, _
                                ByRef pEndTime As DateTime, _
                                ByVal pDepo As String) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        '現時間(pNowDateTime)分解為->日期及分
        Dim xNowDate As String = pNowDateTime.ToString("yyyy/MM/dd")
        Dim xNowTime As Integer = CInt(pNowDateTime.ToString("HH")) * 60 + CInt(pNowDateTime.ToString("mm"))
        Dim xSeqNo As Integer = 0
        Dim xLeadTime As Integer = 0
        Try
            '取得預定開始日
            pStartTime = pNowDateTime
            '取得工程L/T
            If oCommon.GetQCLeadTime(pFormNo, pStep) = 0 Then       'QCL/T 採用=1, 不採用=0
                If oCommon.GetOPLeadTime(pFormNo, pStep, pLevel) > 0 Then
                    xLeadTime = oCommon.GetOPLeadTime(pFormNo, pStep, pLevel) / 10
                Else
                    If oCommon.GetOPLeadTime(pFormNo, pStep, "0") > 0 Then
                        xLeadTime = oCommon.GetOPLeadTime(pFormNo, pStep, "0") / 10
                    Else
                        xLeadTime = CInt(oCommon.GetSystemData("011", "ReadLeadTime").ToString) / 10
                    End If
                End If
            Else
                xLeadTime = pQCLT / 10
            End If
            'MsgBox("GetLeadTime-LeadTime=[" + CStr(xLeadTime) + "]")
            '取得預定完成日
            pEndTime = pNowDateTime
            SQL = "Select SeqNo, YMD, Hour From M_Calendar "
            SQL &= "Where SeqNo >= '" & CStr(xSeqNo + xLeadTime) & "' "
            SQL &= "  And Active = '1' "
            SQL &= "  And Depo = '" & pDepo & "' "
            SQL &= "Order By SeqNo "
            Dim dt_Calendar As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Calendar.Rows.Count > 0 Then
                pEndTime = CDate(dt_Calendar.Rows(0).Item("YMD")).ToString("yyyy/MM/dd") & " " & _
                                 CStr(Int(dt_Calendar.Rows(0).Item("Hour") / 60)) & ":" & _
                                 CStr(dt_Calendar.Rows(0).Item("Hour") Mod 60) & ":" & _
                                 "00"
            End If
            'MsgBox("GetLeadTime-EndTime=[" + pEndTime + "]")
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        FlowControl()

    End Sub
End Class
