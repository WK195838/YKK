Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb

Partial Class AutoFlowControl
    Inherits System.Web.UI.Page

    'iexplore.exe "http://10.245.1.10/IRW/AutoFlowControl.aspx"
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim sFlow As New Flow.FlowService
    Dim sSchedule As New Schedule.ScheduleService

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900      ' 設定逾時時間
        If Not IsPostBack Then
            DSProgressBar.Text = Now.ToString("yyyy/MM/dd HH:mm:ss") & "    批次自動簽核處理中.........."
            FlowControl()
            DEProgressBar.Text = Now.ToString("yyyy/MM/dd HH:mm:ss") & "    批次自動簽核處理完成........"
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(FlowControl)
    '**     簽核處理
    '**
    '*****************************************************************
    Sub FlowControl()
        Dim RtnCode As Integer = 0
        Dim wCount As Integer = 0
        Dim SQL As String = ""
        Dim wNo As String = ""
        Dim NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")
        '處理 pFun=OK, NG1, NG2, SAVE  
        '動作 pAction=0:OK, 1:NG1, 2:NG2, 3:Save   下一關卡時使用 
        '狀態 pSts=0:未處理, 1:OK, 2:NG1, 3:NG2, 4:已閱讀, 5:被抽單  更新T_Waithandle狀態
        '表單代號,單號, 工程代號,序號, 行事曆, 簽核者,申請者,被代理者,表單Table
        Dim wFun As String = "OK"
        Dim wAction As Integer = 0
        Dim wSts As Integer = 1
        Dim wFormNo As String = ""
        Dim wFormSno As Integer = 0
        Dim wStep As Integer = 10
        Dim wSeqNo As Integer = 1
        Dim wCalendar As String = "TP1"
        Dim wDecideID As String = "MK028"
        Dim wApplyID As String = ""
        Dim wAgentID As String = ""
        Dim wTableName As String = "F_ITEMREGISTERSHEET"
        '
        SQL = "Select * From V_ItemRegisterSheet_03 "
        SQL = SQL & "Where ZIPApply = ZIPSts "
        SQL = SQL & "  And SLDApply = SLDSts "
        SQL = SQL & "  And CHApply = CHSts "

        'SQL = SQL & "  And no in ('06130000538312','06130000538311','06130000538310','06130000538309') "

        SQL = SQL & "Order by No "
        Dim dt1151 As DataTable = uDataBase.GetDataTable(SQL)
        For i As Integer = 0 To dt1151.Rows.Count - 1
            '***單價檢查
            If (dt1151.Rows(i)("PriceApply") > 0) Or _
               (dt1151.Rows(i)("PriceApply") = 0 And _
                dt1151.Rows(i)("A001") = 0 And dt1151.Rows(i)("A206") = 0 And dt1151.Rows(i)("A211") = 0 And _
                dt1151.Rows(i)("A999") = 0 And dt1151.Rows(i)("K206") = 0 And dt1151.Rows(i)("K211") = 0) Then
                '***自動簽核
                wFormNo = dt1151.Rows(i)("FormNo").ToString
                wFormSno = dt1151.Rows(i)("FormSno")
                wApplyID = dt1151.Rows(i)("CreateUser").ToString
                wNo = dt1151.Rows(i)("No").ToString
                SQL = "Select * From T_WaitHandle "
                SQL = SQL & " Where Active = 1 "
                SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
                Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(SQL)
                If dtWaitHandle.Rows.Count > 0 Then
                    RtnCode = oCommon.BatchFlowControl(wFun, wAction, wSts, wFormNo, wFormSno, wStep, wSeqNo, wCalendar, wDecideID, wApplyID, wAgentID, wTableName)
                    '
                    SQL = "INSERT INTO L_BatchApprov ( " & _
                          "FormNo, FormSno, No, Code, Message, " & _
                          "CreateUser, CreateTime) " & _
                    "VALUES("
                    SQL &= "'" & wFormNo & "' ,"
                    SQL &= "'" & CStr(wFormSno) & "' ,"
                    SQL &= "'" & wNo & "' ,"
                    SQL &= "'" & CStr(RtnCode) & "' ,"
                    '
                    If RtnCode = 0 Then SQL &= "'" & "正常" & "' ,"
                    If RtnCode = 9100 Then SQL &= "'" & "[CheckFlow]流程資料更新異常!" & "' ,"
                    If RtnCode = 9200 Then SQL &= "'" & "[GetNextGate]下工程計算異常!" & "' ,"
                    If RtnCode = 9210 Then SQL &= "'" & "[GetNextGate]無下工程管理人!" & "' ,"
                    If RtnCode = 9300 Then SQL &= "'" & "[NextFlow]下一工程資料建置異常!" & "' ,"
                    If RtnCode = 9400 Then SQL &= "'" & "[EndFlow]工程結束資料建置異常(999)!" & "' ,"
                    If RtnCode = 9999 Then SQL &= "'" & "[BatchFlowControl]程式程序異常!" & "' ,"
                    '
                    SQL &= "'" & "IRW-AutoFlowControl" & "' ,"
                    SQL &= "'" & NowDateTime & "' )"
                    uDataBase.ExecuteNonQuery(SQL)
                    '
                    wCount = wCount + 1
                End If

            End If
        Next
        '
        DCount.Text = "處理筆數 ( " & CStr(wCount) & " )"
        '
        '限IE11
        Response.Write("<script>window.open('', '_self', ''); window.close();</script>")
    End Sub

    '**************************************************************************************
    '** 批次簽核(BatchFlowControl)
    '**************************************************************************************
    'BatchFlowControl-Start
    Function xxBatchFlowControl(ByVal pFun As String, _
                    ByVal pAction As Integer, _
                    ByVal pSts As String, _
                    ByVal pFormNo As String, _
                    ByVal pFormSno As Integer, _
                    ByVal pStep As Integer, _
                    ByVal pSeqNo As Integer, _
                    ByVal pDecideCalendar As String, _
                    ByVal pDecideID As String, _
                    ByVal pApplyID As String, _
                    ByVal pAgentID As String, _
                    ByVal pTableName As String) As Integer

        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim wQCLT As Integer = 0 'QC-L/T
        Dim Run As Boolean = True           '是否執行
        Dim RepeatRun As Boolean = False    '是否重覆執行
        Dim MultiJob As Integer = 0         '多人核定
        Dim wLevel As String = ""           '難易度
        Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日時
        '
        Try
            While Run = True
                Run = False     '執行Flag=不執行
                '--取得下一關參數設定---------
                Dim pNextGate(10) As String
                Dim pAgentGate(10) As String
                Dim pNextStep As Integer = 0
                Dim pFlowType As Integer = 0    '0=通知
                Dim pCount As Integer
                '--取得LeadTime參數設定---------
                Dim pCTime, pStartTime, pEndTime As DateTime
                '--取得工程負荷參數設定---------
                Dim pLastTime As DateTime
                Dim pCount1 As Integer
                '--其他參數設定---------
                Dim RtnCode, i As Integer
                Dim NewFormSno As Integer = pFormSno    '表單流水號
                Dim pRunNextStep As Integer = 0         '是否執行計算下一關(會簽)
                Dim wNextGateCalendar As String

                '取得表單流水號或更新交易資料
                If ErrCode = 0 Then
                    If RepeatRun = False Then   '不是通知的重覆執行
                        '更新交易資料

                        fpObj.ModifyTranData(pFun, pSts, pFormNo, pFormSno, pStep, pSeqNo, NowDateTime, pDecideID)

                        '流程資料結束(表單號碼,表單流水號,工程關卡號碼,行事曆,簽核者, 流程結束否(會簽))
                        RtnCode = sFlow.CheckFlow(pFormNo, pFormSno, pStep, pDecideCalendar, pDecideID, pRunNextStep)
                        If RtnCode <> 0 Then
                            ErrCode = 9100
                        End If
                    Else
                        pRunNextStep = 1    '是通知的重覆執行
                    End If
                End If

                '取得下一關
                If ErrCode = 0 And pRunNextStep = 1 Then
                    '指定簽核者User ID
                    Dim wAllocateID As String = ""
                    '取得簽核者
                    '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
                    '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 
                    RtnCode = oCommon.GetNextGate(pFormNo, pStep, pDecideID, pApplyID, pAgentID, wAllocateID, MultiJob, _
                                                  pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)

                    If RtnCode <> 0 Then ErrCode = 9200
                    If pCount = 0 And pNextStep <> 999 Then ErrCode = 9210
                    If ErrCode = 0 Then pAction = 0
                End If

                '建置流程資料
                If ErrCode = 0 And pRunNextStep = 1 Then
                    If pNextStep <> 999 Then
                        For i = 1 To pCount
                            '取得核定者-群組行是曆
                            wNextGateCalendar = oCommon.GetCalendarGroup(pNextGate(i))

                            '取得工程負荷最後日時(核定者, 表單號碼, 工程號碼, 類別(0:通知,1:核定), 開始日時, 最後日時, 件數)
                            sSchedule.GetLastTime(pNextGate(i), pFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)

                            '取得預定開始,完成日程計算(表單號碼,工程號碼,難易度,QC-L/T,現在時間, 預定開始日時, 預定完成日時, 行事曆)
                            sSchedule.GetLeadTime(pFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wNextGateCalendar)

                            '建置流程資料(表單號碼,表單流水號,工程關卡號碼,序號,申請者ID,行事曆,建置者, 簽核者, 被代理者, 申請者, 預定開始日, 預定完成日, 重要性)
                            RtnCode = sFlow.NextFlow(pFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, pDecideID, pNextGate(i), pAgentGate(i), pApplyID, pStartTime, pEndTime, 0)

                            If RtnCode <> 0 Then
                                ErrCode = 9300
                                Exit For
                            End If
                        Next i
                    Else
                        '流程結束(表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者)
                        RtnCode = sFlow.EndFlow(pFormNo, pFormSno, pNextStep, 1, pDecideCalendar, pDecideID, pApplyID)

                        If RtnCode <> 0 Then ErrCode = 9400
                    End If
                End If

                '當工程日程調整
                If ErrCode = 0 Then
                    If RepeatRun = False Then
                        '簽核者,表單號碼,表單流水號,工程關卡號碼,序號,現在日時,難易度,行事曆
                        RtnCode = sSchedule.AdjustSchedule(pDecideID, pFormNo, pFormSno, pStep, pSeqNo, NowDateTime, wLevel, pDecideCalendar)
                    End If
                End If

                '儲存表單資料
                If ErrCode = 0 Then
                    If pNextStep = 999 Then     '工程結束嗎?
                        fpObj.ModifyData(pFun, "1", pFormNo, pFormSno, NowDateTime, pDecideID, pTableName) '更新表單資料
                    Else
                        fpObj.ModifyData(pFun, "0", pFormNo, pFormSno, NowDateTime, pDecideID, pTableName) '更新表單資料 Sts=0(未結)
                    End If

                    '傳送郵件
                    If pNextStep <> 999 Then
                        For i = 1 To pCount
                            oCommon.Send(pDecideID, pNextGate(i), pApplyID, pFormNo, NewFormSno, pNextStep, "FLOW")
                            '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                        Next i
                    Else
                        oCommon.Send(pDecideID, pApplyID, pApplyID, pFormNo, NewFormSno, pNextStep, "END")
                        '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                    End If

                    If pFlowType <> 3 Then
                        MultiJob = 0
                    Else
                        MultiJob = 1
                    End If

                    If (pRunNextStep = 1) And (pFlowType = 0 Or pFlowType = 3) Then
                        Run = True
                        RepeatRun = True
                        pAction = 0
                    End If

                    pStep = pNextStep     '下一工程關卡號碼
                    pFormSno = NewFormSno '下一工程表單流水號
                End If      '儲存表單ErrCode=0
            End While       '重覆執行
        Catch ex As Exception
            ErrCode = 9999
        End Try
        '
        Return ErrCode
    End Function
    'BatchFlowControl-End


End Class
