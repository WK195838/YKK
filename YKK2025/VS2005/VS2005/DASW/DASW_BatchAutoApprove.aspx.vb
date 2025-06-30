
Imports System.Data

Partial Class DASW_BatchAutoApprove
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
    '
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    '
    Dim xFormNo As String
    Dim xFormSno, xStep As Integer
    Dim xDecideID, xAllocateID, xCalendarGr As String
    Dim xRecordCount As Integer
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim w40 As Integer = 0 '判斷寄送mail次數
    Dim w90 As Integer = 0 '判斷寄送mail次數
    Dim w82 As Integer = 0 '判斷寄送mail次數
    Dim w100 As Integer = 0 '判斷寄送mail次數

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900      ' 設定逾時時間
        If Not IsPostBack Then
            FlowControl()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     自動簽核處理
    '**
    '*****************************************************************
    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        DMsg.Text = Now.ToString("yyyy/MM/dd HH:mm:ss") & "  批次自動簽核處理中.........."
        '
        xRecordCount = 0
        xFormNo = "006001"
        xFormSno = CInt(DFormSno.Text)
        xStep = CInt(DStep.Text)
        xDecideID = DDecideID.Text
        xAllocateID = DAllocateID.Text

        xCalendarGr = DcalendarGr.Text
        '
        FlowControl()
        '
        DMsg.Text = Now.ToString("yyyy/MM/dd HH:mm:ss") & "  批次自動簽核完成(筆數=" & CStr(xRecordCount) & ")"
    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(FlowControl)
    '**     簽核處理
    '**
    '*****************************************************************
    Sub FlowControl()
        '處理 pFun=OK, NG1, NG2, SAVE  
        '動作 pAction=0:OK, 1:NG1, 2:NG2, 3:Save   下一關卡時使用 
        '狀態 pSts=0:未處理, 1:OK, 2:NG1, 3:NG2, 4:已閱讀, 5:被抽單  更新T_Waithandle狀態
        '表單代號,單號, 工程代號,序號, 行事曆, 簽核者,申請者,被代理者,表單Table
        Dim wFun As String = ""
        Dim wAction As Integer = 0
        Dim wSts As Integer = 0
        Dim wFormSno As Integer = 0
        Dim wSeqNo As Integer = 1
        Dim wApplyID As String = ""
        Dim wAgentID As String = ""
        Dim wDecideDesc As String = ""
        Dim RtnCode As Integer = 0
        Dim SQL As String = ""
        Dim wNo As String = ""
        Dim wDivision As String = ""
        Dim NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")
        xFormNo = "006001"


        '查詢自動簽核暫存檔
        SQL = "Select * From Q_WaitAutoDisposal  "
        SQL = SQL & " Where formno = '006001' "
        SQL = SQL & " Order by FormSno "
        '
        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)
        For i As Integer = 0 To dtFlow.Rows.Count - 1
            wFormSno = dtFlow.Rows(i)("FormSno")
            wApplyID = dtFlow.Rows(i)("ApplyID").ToString
            xStep = dtFlow.Rows(i)("Step").ToString
            xDecideID = dtFlow.Rows(i)("DecideID").ToString
            wDecideDesc = dtFlow.Rows(i)("DecideDesc").ToString
            wSeqNo = "1"
            ' wDivision = dtFlow.Rows(i)("Division").ToString
            If dtFlow.Rows(i)("Sts").ToString = "1" Then
                wFun = "OK"
                wAction = "0"
                wSts = "1"
            Else
                wFun = "NG1"
                wAction = "1"
                wSts = "2"
            End If

            ' 先檢查 T_WaitHandle 的狀態是為=5
            SQL = " select * from  T_WaitHandle "
            SQL = SQL + " where formno ='006001'  "
            SQL = SQL + " and formsno = '" + CStr(wFormSno) + "'"
            SQL = SQL + " and step='" + CStr(xStep) + "'"
            SQL = SQL + " and seqno='" + CStr(wSeqNo) + "'"
            SQL = SQL + " and sts=5"
            Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(SQL)
            If dtWaitHandle.Rows.Count > 0 Then
                '
                RtnCode = BatchFlowControl(wFun, wAction, wSts, xFormNo, wFormSno, xStep, wSeqNo, xCalendarGr, xDecideID, wApplyID, wAgentID, xAllocateID, wDecideDesc)
                '完成後 存入log 
                SQL = "INSERT INTO L_FinishAutoApprove( " & _
                      "FormNo, FormSno,ApplyId,DecideId,Step,SeqNo,NextDecideID, " & _
                      "message,CreateTime) " & _
                "VALUES("
                SQL &= "'" & xFormNo & "' ,"
                SQL &= "'" & CStr(wFormSno) & "' ,"
                SQL &= "'" & wApplyID & "' ,"
                SQL &= "'" & xDecideID & "' ,"
                SQL &= "" & xStep & " ,"
                SQL &= "" & wSeqNo & " ,"
                '   SQL &= "'" & wDivision & "' ,"                '
                SQL &= "'" & xAllocateID & "' ,"                '


                If RtnCode = 0 Then SQL &= "'" & "正常" & "' ,"
                If RtnCode = 9100 Then SQL &= "'" & "[CheckFlow]流程資料更新異常!" & "' ,"
                If RtnCode = 9200 Then SQL &= "'" & "[GetNextGate]下工程計算異常!" & "' ,"
                If RtnCode = 9210 Then SQL &= "'" & "[GetNextGate]無下工程管理人!" & "' ,"
                If RtnCode = 9300 Then SQL &= "'" & "[NextFlow]下一工程資料建置異常!" & "' ,"
                If RtnCode = 9400 Then SQL &= "'" & "[EndFlow]工程結束資料建置異常(999)!" & "' ,"
                If RtnCode = 9999 Then SQL &= "'" & "[BatchFlowControl]程式程序異常!" & "' ,"
                SQL &= "'" & NowDateTime & "' )"
                uDataBase.ExecuteNonQuery(SQL)
                '自動簽核後刪除暫存table資料
                SQL = "delete  From Q_WaitAutoDisposal"
                SQL = SQL + " where formno = '006001' and "
                SQL = SQL + " formsno= '" + CStr(wFormSno) + "'"
                SQL = SQL + " and applyid='" + wApplyID + "'"
                SQL = SQL + " and step='" + CStr(xStep) + "'"
                uDataBase.ExecuteNonQuery(SQL)
            End If

            'End If
            xRecordCount = xRecordCount + 1
        Next

        'Response.Write("<script>void(window.open('','_parent',''));window.close();</script>")
        'Response.Write("<script>window.opener=null;window.close();</script>")
        'IE8 可以自行關閉網頁
        Response.Write("<script   language=javascript>  window.opener=null;    window.open('','_self');  window.close();</script> ")
        'IE11
        Response.Write("<script>window.open('', '_self', ''); window.close();</script>")







    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BatchFlowControl)
    '**     簽核處理
    '**
    '*****************************************************************
    Function BatchFlowControl(ByVal pFun As String, _
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
                    ByVal pAllocteID As String, _
                    ByVal pDecideDesc As String) As Integer

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

                        ModifyTranData(pFun, pSts, pFormNo, pFormSno, pStep, pSeqNo, NowDateTime, pDecideID, pDecideDesc)

                        '流程資料結束(表單號碼,表單流水號,工程關卡號碼,行事曆,簽核者, 流程結束否(會簽))
                        RtnCode = oFlow.CheckFlow(pFormNo, pFormSno, pStep, pDecideCalendar, pDecideID, pRunNextStep)
                        If RtnCode <> 0 Then
                            ErrCode = 9100
                        End If
                    Else
                        pRunNextStep = 1    '是通知的重覆執行
                    End If
                End If

                '取得下一關
                If ErrCode = 0 And pRunNextStep = 1 Then
                    '-----------------------------
                    '找出簽核者User ID
                    Dim wAllocateID As String = ""

                    If pStep = 80 And pSts = 2 Then
                        wAllocateID = GetRelated(GETDUTYDEPOID(pFormSno))
                    End If
                    If pStep = 82 Then
                        wAllocateID = GETDUTYDEPOID(pFormSno)
                    End If

                    If pStep = 30 And pFun = "NG1" Then
                        wAllocateID = GetRelated(GetRelated(pApplyID)) '提出部門日籍主管
                    End If


                    If pStep = 31 Then
                        wAllocateID = GetRelated(pApplyID)
                    End If


                    '找出簽核者User ID 台藉主管 
                    If pStep = 111 Then  '指定台藉主管 
                        If pAction = 0 Then
                            wAllocateID = GetRelated(pApplyID)
                        End If
                    End If


                    '找出簽核者User ID  
                    If pStep = 112 Then  '指定生管擔當扣帳
                        If pAction = 0 Then
                            wAllocateID = GetUserID(pFormSno)
                        End If
                    End If


                  



                    '取得簽核者
                    '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
                    '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 
                    RtnCode = oCommon.GetNextGate(pFormNo, pStep, pDecideID, pApplyID, pAgentID, wAllocateID, MultiJob, _
                                                  pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)

                    '
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
                            oSchedule.GetLastTime(pNextGate(i), pFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)

                            '取得預定開始,完成日程計算(表單號碼,工程號碼,難易度,QC-L/T,現在時間, 預定開始日時, 預定完成日時, 行事曆)
                            oSchedule.GetLeadTime(pFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wNextGateCalendar)

                            '建置流程資料(表單號碼,表單流水號,工程關卡號碼,序號,申請者ID,行事曆,建置者, 簽核者, 被代理者, 申請者, 預定開始日, 預定完成日, 重要性)
                            RtnCode = oFlow.NextFlow(pFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, pDecideID, pNextGate(i), pAgentGate(i), pApplyID, pStartTime, pEndTime, 0)

                            If RtnCode <> 0 Then
                                ErrCode = 9300
                                Exit For
                            End If
                        Next i
                    Else
                        '流程結束(表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者)
                        RtnCode = oFlow.EndFlow(pFormNo, pFormSno, pNextStep, 1, pDecideCalendar, pDecideID, pApplyID)

                        If RtnCode <> 0 Then ErrCode = 9400
                    End If
                End If

                '當工程日程調整
                If ErrCode = 0 Then
                    If RepeatRun = False Then
                        '簽核者,表單號碼,表單流水號,工程關卡號碼,序號,現在日時,難易度,行事曆
                        RtnCode = oSchedule.AdjustSchedule(pDecideID, pFormNo, pFormSno, pStep, pSeqNo, NowDateTime, wLevel, pDecideCalendar)
                    End If
                End If

                '儲存表單資料
                If ErrCode = 0 Then
                    If pNextStep = 999 Then     '工程結束嗎?
                        ModifyData(pFun, "1", pFormNo, pFormSno, NowDateTime, pDecideID) '更新表單資料
                    Else
                        ModifyData(pFun, "0", pFormNo, pFormSno, NowDateTime, pDecideID) '更新表單資料 Sts=0(未結)
                    End If

                    If pNextStep = 40 Then
                        w40 = w40 + 1
                    End If


                    If pNextStep = 90 Then
                        w90 = w90 + 1
                    End If

                    If pNextStep = 100 Then
                        w100 = w100 + 1
                    End If

                    If pNextStep = 82 Then
                        w82 = w82 + 1
                    End If

                    '傳送郵件
                    If pNextStep <> 999 Then
                        If pSts = 2 Then
                            For i = 1 To pCount
                                oCommon.Send(pDecideID, pNextGate(i), pApplyID, pFormNo, NewFormSno, pNextStep, "FLOW")

                                '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                            Next i
                        ElseIf pNextStep = 40 And w40 = 1 Then
                            For i = 1 To pCount
                                oCommon.Send(pDecideID, pNextGate(i), pApplyID, pFormNo, NewFormSno, pNextStep, "FLOW")

                                '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                            Next i
                        ElseIf pNextStep = 90 And w90 = 1 Then
                            For i = 1 To pCount
                                oCommon.Send(pDecideID, pNextGate(i), pApplyID, pFormNo, NewFormSno, pNextStep, "FLOW")

                                '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                            Next i
                        ElseIf pNextStep = 100 And w100 = 1 Then
                            For i = 1 To pCount
                                oCommon.Send(pDecideID, pNextGate(i), pApplyID, pFormNo, NewFormSno, pNextStep, "FLOW")

                                '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                            Next i
                        ElseIf pNextStep = 110 Or pNextStep = 82 Or pNextStep = 45 Then
                            For i = 1 To pCount
                                oCommon.Send(pDecideID, pNextGate(i), pApplyID, pFormNo, NewFormSno, pNextStep, "FLOW")

                                '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                            Next i
                        End If
                    Else
                        '  oCommon.Send(pDecideID, pApplyID, pApplyID, pFormNo, NewFormSno, pNextStep, "END")
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
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyTranData)
    '**     更新交易資料
    '**
    '*****************************************************************
    Public Function ModifyTranData(ByVal pFun As String, _
                                   ByVal pSts As String, _
                                   ByVal pFormNo As String, _
                                   ByVal pFormSno As Integer, _
                                   ByVal pStep As Integer, _
                                   ByVal pSeqNo As Integer, _
                                   ByVal pNowDateTime As String, _
                                   ByVal pDecideID As String, _
                                   ByVal pDecideDesc As String) As Integer
        Dim RtnCode As Integer = 0
        Dim SQl As String
        SQl = "Update T_WaitHandle Set "
        SQl = SQl + " Active = '" & "0" & "',"
        SQl = SQl + " Sts = '" & pSts & "',"
        If pSts = "1" Then SQl = SQl + " StsDes = 'OK' ,"
        If pSts = "2" Then SQl = SQl + " StsDes = '退回申請者',"
        SQl = SQl + " AEndTime = '" & pNowDateTime & "',"
        SQl = SQl + " CompletedTime = '" & pNowDateTime & "',"
        SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(pDecideDesc) & "',"
        SQl = SQl + " ModifyUser = '" & pDecideID & "',"
        SQl = SQl + " ModifyTime = '" & pNowDateTime & "' "
        SQl = SQl + " Where FormNo  =  '" & pFormNo & "'"
        SQl = SQl + "   And FormSno =  '" & CStr(pFormSno) & "'"
        SQl = SQl + "   And Step    =  '" & CStr(pStep) & "'"
        SQl = SQl + "   And SeqNo   =  '" & CStr(pSeqNo) & "'"
        SQl = SQl + "   And Active =  '1' "
        Try
            uDataBase.ExecuteNonQuery(SQl)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Function ModifyData(ByVal pFun As String, _
                               ByVal pSts As String, _
                               ByVal pFormNo As String, _
                               ByVal pFormSno As Integer, _
                               ByVal pNowDateTime As String, _
                               ByVal pDecideID As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        sql = "Update F_DisposalSheet Set "
        sql &= " Sts = '" & pSts & "',"
        sql &= " CompletedTime = '" & pNowDateTime & "',"
        sql &= " ModifyUser = '" & pDecideID & "',"
        sql &= " ModifyTime = '" & pNowDateTime & "' "
        sql &= " Where FormNo  =  '" & pFormNo & "'"
        sql &= "   And FormSno =  '" & CStr(pFormSno) & "'"
        '
        Try
            uDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        If RtnCode = 0 Then
            sql = "Update F_DisposalSheetDT Set "
            sql &= " Sts = '" & pSts & "',"
            sql &= " CompletedTime = '" & pNowDateTime & "',"
            sql &= " ModifyUser = '" & pDecideID & "',"
            sql &= " ModifyTime = '" & pNowDateTime & "' "
            sql &= " Where FormNo  =  '" & pFormNo & "'"
            sql &= "   And FormSno =  '" & CStr(pFormSno) & "'"
            '
            Try
                uDataBase.ExecuteNonQuery(sql)
            Catch ex As Exception
                RtnCode = 1
            End Try
        End If
        '
        Return RtnCode
        '
    End Function
    '重新取得責任部門關係人
    Function GETDUTYDEPOID(ByVal value As String) As String
        Dim UserId As String = ""
        Dim SQL As String
        SQL = " select * from f_disposalsheet"
        SQL = SQL + " where formsno ='" + value.ToString + "'"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(SQL)
        If Trim(dt.Rows(0)("DutyDepo")) <> "工廠" Then

            If Trim(dt.Rows(0)("DISPOSALTYPE")) = "PN" Or Trim(dt.Rows(0)("DISPOSALTYPE")) = "QF" Or Trim(dt.Rows(0)("DISPOSALTYPE")) = "S&B" Then
                SQL = " select  userid,username  from M_users "
                SQL = SQL + " where username in ("
                SQL = SQL + "  select data from m_referp"
                SQL = SQL + " where cat = 6001 and dkey = 'DUTYDEPO-" + Trim(dt.Rows(0)("DUTYDEPO")) + "'"
                SQL = SQL + " )"
            Else
                SQL = " select  userid,username  from M_users "
                SQL = SQL + " where username in ("
                SQL = SQL + "  select data from m_referp"
                SQL = SQL + " where cat = 6001 and dkey = 'DUTYDEPO-" + Trim(dt.Rows(0)("DUTYDEPO")) + "'"
                SQL = SQL + " )"
            End If

            Dim dt2 As DataTable = uDataBase.GetDataTable(SQL)
            If dt2.Rows.Count > 0 Then
                UserId = dt2.Rows(0).Item("userid")
            End If
        Else

            If Trim(dt.Rows(0)("MKTSIGN")) = 1 Then  ' 需營業簽核
                If Trim(dt.Rows(0)("DISPOSALTYPE")) = "PN" Or Trim(dt.Rows(0)("DISPOSALTYPE")) = "QF" Or Trim(dt.Rows(0)("DISPOSALTYPE")) = "S&B" Then
                    SQL = " select  userid,username  from M_users "
                    SQL = SQL + " where username in ("
                    SQL = SQL + "  select data from m_referp"
                    SQL = SQL + " where cat = 6001 and dkey = 'DUTYDEPO-" + Trim(dt.Rows(0)("DISPOSALTYPE")) + "'"
                    SQL = SQL + " )"
                Else
                    SQL = " select  userid,username  from M_users "
                    SQL = SQL + " where username in ("
                    SQL = SQL + "  select data from m_referp"
                    SQL = SQL + " where cat = 6001 and dkey = 'DUTYDEPO-營業'"
                    SQL = SQL + " )"
                End If

                Dim dt2 As DataTable = uDataBase.GetDataTable(SQL)
                If dt2.Rows.Count > 0 Then
                    UserId = dt2.Rows(0).Item("userid")
                End If
            Else
                SQL = " select ruserid,rusername  from m_condition"
                SQL = SQL + " where userid in ( select ruserid from m_condition"
                SQL = SQL + " where userid = '" + Trim(dt.Rows(0)("Createuser")) + "'"
                SQL = SQL + " and relatedid = 'd')"
                SQL = SQL + " and relatedid = 'd'"
                Dim dt1 As DataTable = uDataBase.GetDataTable(SQL)
                If dt1.Rows.Count > 0 Then
                    UserId = dt1.Rows(0).Item("userid")
                End If
            End If


        End If


        Return UserId
    End Function



    '取得關係人
    Function GetRelated(ByVal userId As String) As String

        Dim sql As String = "select RUserID , RUserName  from M_Related where userid='" & userId & "' and RelatedID='D'"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            NextGate = dt.Rows(0)("RUserID")
        End If
        Return NextGate
    End Function


    Function GetUserID(ByVal FormSno As String) As String
        Dim sql As String
        sql = " select userid from ("
        sql = sql & " select pcname from f_disposalsheet"
        sql = sql & " where FormSno = " & FormSno + ""
        sql = sql & " )a,m_users b"
        sql = sql & " where(a.pcname = b.username)"
     
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            NextGate = dt.Rows(0)("Userid")
        End If
        Return NextGate

    End Function


 
   
End Class
