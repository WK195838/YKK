Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO

<WebService(Namespace:="http://tempuri.org/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class CommonService
     Inherits System.Web.Services.WebService

    '**************************************************************************************
    '** 全域變數
    '**************************************************************************************
    Dim DBObj As New ForProject
    Dim uDataBase As Utility.DataBase = DBObj.GetDataBaseObj()

    '**************************************************************************************
    '** 封存檔還原(RecoveryArchiveFile)
    '**************************************************************************************
    'RecoveryArchiveFile-Start
    <WebMethod()> _
    Public Function RecoveryArchiveFile(ByVal pHost As String, ByVal pPath As String, ByVal pFile As String) As Integer
        Dim RtnCode As Integer = 0
        Dim xSourceFile As String = pHost & pPath & pFile
        Dim xTargetFile As String = pHost & pPath & pFile
        Dim xArchiveFile As String = "\\10.245.1.9\SyncData\inetpub\wwwroot\" & pPath & pFile
        '
        '\\10.245.1.10\www$\WorkFlow\Document\000001
        '\\10.245.1.9\SyncData\inetpub\wwwroot\WorkFlow\Document\000001
        '
        Try
            'MsgBox("[source][" & xSourceFile & "]")
            If Not File.Exists(xSourceFile) Then
                File.Copy(xArchiveFile, xTargetFile)
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'RecoveryArchiveFile-End

    '**************************************************************************************
    '** 可使用代理簽核(UseAgentApprov)
    '**************************************************************************************
    'UseAgentApprov-Start
    <WebMethod()> _
    Public Function UseAgentApprov(ByVal pFormNo As String, ByVal pStep As Integer, ByVal pFun As String) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        Try
            If pFun = "SAVE" Then RtnCode = 1
            '
            If RtnCode = 0 Then
                SQL = "Select * From M_AgentApprovTask "
                SQL &= "Where FormNo = '" & pFormNo & "' "
                SQL &= "  And Step   = '" & CStr(pStep) & "' "
                SQL &= "  And Active = '1' "
                Dim dt_AgentApprov As DataTable = uDataBase.GetDataTable(SQL)
                If dt_AgentApprov.Rows.Count <= 0 Then
                    RtnCode = 1
                End If
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'UseAgentApprov-End

    '**************************************************************************************
    '** 列入標準代理簽核序列檔(RunAgentApprov)
    '**************************************************************************************
    'pFun		    pFun=OK, NG1, NG2, SAVE  
    'pAction	    pAction=0:OK, 1:NG1, 2:NG2, 3:Save   下一關卡時使用 
    'pSts		    pSts=0:未處理, 1:OK, 2:NG1, 3:NG2, 4:已閱讀, 5:被抽單  更新T_Waithandle狀態
    'pStsDes	    pSts說明
    'pFormNo	    表單代號
    'pFormSno	    單號
    'pStep		    工程代號
    'pSeqNo		    序號
    'pDecideCalendar行事曆
    'pDecideID	    簽核者	
    'pApplyID	    申請者
    'pAgentID	    被代理者
    'pAllocteID	    指定簽核者
    'pDecideDesc    承認說明
    'pTableName	表單Table
    'RunAgentApprov-Start
    <WebMethod()> _
    Public Function RunAgentApprov(ByVal pFun As String, _
                    ByVal pAction As Integer, _
                    ByVal pSts As String, _
                    ByVal pStsDes As String, _
                    ByVal pFormNo As String, _
                    ByVal pFormSno As Integer, _
                    ByVal pStep As Integer, _
                    ByVal pSeqNo As Integer, _
                    ByVal pDecideCalendar As String, _
                    ByVal pDecideID As String, _
                    ByVal pApplyID As String, _
                    ByVal pAgentID As String, _
                    ByVal pAllocteID As String, _
                    ByVal pDecideDesc As String, _
                    ByVal pTableName As String, _
                    ByVal pUserIP As String) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")
        Try
            SQL = "Select * From M_AgentApprovTask "
            SQL &= "Where FormNo = '" & pFormNo & "' "
            SQL &= "  And Step   = '" & CStr(pStep) & "' "
            SQL &= "  And Active = '1' "
            Dim dt_AgentApprov As DataTable = uDataBase.GetDataTable(SQL)
            If dt_AgentApprov.Rows.Count > 0 Then
                '
                '2020/12/9 移至執行代理簽核處理(AgentApprovStart)
                '*
                '簽核流程狀態=6 代理簽核處理
                'SQL = "update T_WaitHandle "
                'SQL = SQL + "set sts = 6, "
                'SQL = SQL + "stsdes ='代理簽核', "
                'SQL = SQL + "ModifyUser =DecideID, "
                'SQL = SQL + "ModifyTime =getdate() "
                'SQL = SQL + "where DecideID = '" & pDecideID & "' "
                'SQL = SQL + "AND FormNo = '" & pFormNo & "' "
                'SQL = SQL + "and FormSNo = " & CStr(pFormSno) & " "
                'SQL = SQL + "and Step = " + CStr(pStep) & " "
                'SQL = SQL + "and SeqNo = " + CStr(pSeqNo) & " "
                'SQL = SQL + "and Active = '1' and flowtype=1  And (Sts = '0'  Or  Sts = '4')"
                'uDataBase.ExecuteNonQuery(SQL)
                '
                '列入代理簽核序列檔
                SQL = "Insert into Q_AgentApprov "
                SQL &= "( "
                SQL &= "FormNo, FormSno, Step, SeqNo, Fun, "
                SQL &= "Action, Sts, StsDes, DecideCalendar, DecideID, ApplyID, "
                SQL &= "AgentID, AllocteID, DecideDesc, TableName, UserIP, CreateTime "
                SQL &= ") "
                SQL &= "VALUES( "
                SQL &= "'" & pFormNo & "', "
                SQL &= " " & CStr(pFormSno) & ", "
                SQL &= " " & CStr(pStep) & ", "
                SQL &= " " & CStr(pSeqNo) & ", "
                SQL &= "'" & pFun & "', "
                SQL &= " " & CStr(pAction) & ", "
                SQL &= "'" & pSts & "', "
                SQL &= "'" & pStsDes & "', "
                SQL &= "'" & pDecideCalendar & "', "
                SQL &= "'" & pDecideID & "', "
                SQL &= "'" & pApplyID & "', "
                SQL &= "'" & pAgentID & "', "
                SQL &= "'" & pAllocteID & "', "
                SQL &= "N'" & pDecideDesc & "', "
                SQL &= "'" & pTableName & "', "
                SQL &= "'" & pUserIP & "', "
                SQL &= "'" & NowDateTime & "' "             '製作時間
                SQL &= ") "
                uDataBase.ExecuteNonQuery(SQL)
            Else
                RtnCode = 1
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'RunAgentApprov-End

    '**************************************************************************************
    '** 執行代理簽核處理(AgentApprovStart)
    '**************************************************************************************
    'AgentApprovStart-Start
    <WebMethod()> _
    Public Function AgentApprovStart() As Integer
        Dim RtnCode As Integer = 0
        Dim SQL, Msg As String
        Dim xNextDecideID As String = ""
        Dim xNextStep As Integer = 0
        Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")
        Dim xErr As Boolean
        Try
            SQL = "Select * From Q_AgentApprov "
            'TEST-START
            'SQL = SQL & "Where FormNo = '" & "003101" & "' "
            'SQL = SQL & "And FormSno = " & "1680" & " "
            'END
            SQL = SQL & "Order by Unique_ID "
            Dim dtAgentApprov As DataTable = uDataBase.GetDataTable(SQL)
            For i As Integer = 0 To dtAgentApprov.Rows.Count - 1
                xErr = False
                '
                SQL = " select * from  T_WaitHandle "
                SQL = SQL + "where FormNo = '" & dtAgentApprov.Rows(i)("FormNo").ToString & "' "
                SQL = SQL + "  and FormSno = " & dtAgentApprov.Rows(i)("FormSno").ToString & " "
                SQL = SQL + "  and Step    = " & dtAgentApprov.Rows(i)("Step").ToString & " "
                SQL = SQL + "  and seqno   = " & dtAgentApprov.Rows(i)("SeqNo").ToString & " "
                SQL = SQL + "  and DecideID = '" & dtAgentApprov.Rows(i)("DecideID").ToString & "' "
                SQL = SQL + "  and Active = '1' and flowtype=1  And (Sts = '0'  Or  Sts = '4')"
                Dim dtWaitHandle1 As DataTable = uDataBase.GetDataTable(SQL)
                If dtWaitHandle1.Rows.Count > 0 Then
                    '設定簽核流程狀態=6 代理簽核處理(異常時可追蹤使用)
                    SQL = "update T_WaitHandle "
                    SQL = SQL + "set sts = 6, "
                    SQL = SQL + "stsdes ='代理簽核', "
                    SQL = SQL + "ModifyUser = '" & dtAgentApprov.Rows(i)("DecideID").ToString & "', "
                    SQL = SQL + "ModifyTime =getdate() "
                    SQL = SQL + "where DecideID = '" & dtAgentApprov.Rows(i)("DecideID").ToString & "' "
                    SQL = SQL + "AND FormNo = '" & dtAgentApprov.Rows(i)("FormNo").ToString & "' "
                    SQL = SQL + "and FormSNo = " & dtAgentApprov.Rows(i)("FormSno").ToString & " "
                    SQL = SQL + "and Step = " + dtAgentApprov.Rows(i)("Step").ToString & " "
                    SQL = SQL + "and SeqNo = " + dtAgentApprov.Rows(i)("SeqNo").ToString & " "
                    SQL = SQL + "and Active = '1' and flowtype=1  And (Sts = '0'  Or  Sts = '4')"
                    uDataBase.ExecuteNonQuery(SQL)
                Else
                    SQL = "INSERT INTO L_AgentApprov " & _
                              "Select FormNo,FormSno,Step,SeqNo,Fun, " & _
                              "Action,Sts,StsDes,DecideCalendar,DecideID,ApplyID, " & _
                              "AgentID,AllocteID,DecideDesc,TableName,UserIP,CreateTime, " & _
                              "'" & xNextDecideID & "', " & _
                              " " & xNextStep & " , " & _
                              "'" & "無效資料" & "', " & _
                              "'" & NowDateTime & "' " & _
                              "From Q_AgentApprov Where Unique_ID = " & dtAgentApprov.Rows(i)("Unique_ID").ToString & " "
                    uDataBase.ExecuteNonQuery(SQL)
                    '
                    '刪除代理簽核序列檔
                    SQL = "delete  From Q_AgentApprov"
                    SQL = SQL + " where Unique_ID = " & dtAgentApprov.Rows(i)("Unique_ID").ToString & " "
                    uDataBase.ExecuteNonQuery(SQL)
                    '
                    xErr = True
                End If
                '
                If xErr = False Then
                    '
                    '簽核流程狀態是為=6
                    SQL = " select * from  T_WaitHandle "
                    SQL = SQL + "where FormNo = '" & dtAgentApprov.Rows(i)("FormNo").ToString & "' "
                    SQL = SQL + "  and FormSno = " & dtAgentApprov.Rows(i)("FormSno").ToString & " "
                    SQL = SQL + "  and Step    = " & dtAgentApprov.Rows(i)("Step").ToString & " "
                    SQL = SQL + "  and seqno   = " & dtAgentApprov.Rows(i)("SeqNo").ToString & " "
                    SQL = SQL + "  and sts = 6 "
                    Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(SQL)
                    If dtWaitHandle.Rows.Count > 0 Then
                        '
                        '執行代理簽核
                        RtnCode = AgentApprov(dtAgentApprov.Rows(i)("Fun").ToString, _
                                                dtAgentApprov.Rows(i)("Action"), _
                                                dtAgentApprov.Rows(i)("Sts").ToString, _
                                                dtAgentApprov.Rows(i)("StsDes").ToString, _
                                                dtAgentApprov.Rows(i)("FormNo").ToString, _
                                                dtAgentApprov.Rows(i)("FormSno"), _
                                                dtAgentApprov.Rows(i)("Step"), _
                                                dtAgentApprov.Rows(i)("SeqNo"), _
                                                dtAgentApprov.Rows(i)("DecideCalendar").ToString, _
                                                dtAgentApprov.Rows(i)("DecideID").ToString, _
                                                dtAgentApprov.Rows(i)("ApplyID").ToString, _
                                                dtAgentApprov.Rows(i)("AgentID").ToString, _
                                                dtAgentApprov.Rows(i)("AllocteID").ToString, _
                                                dtAgentApprov.Rows(i)("DecideDesc").ToString, _
                                                dtAgentApprov.Rows(i)("TableName").ToString, _
                                                xNextDecideID, xNextStep)
                        '
                        '代理簽核LOG
                        If xNextDecideID = "" Then
                            If xNextStep = 999 Then
                                RtnCode = 0
                            Else
                                RtnCode = 9210
                            End If
                        End If
                        '
                        If RtnCode = 0 Then Msg = "正常"
                        If RtnCode = 9100 Then Msg = "[CheckFlow]流程資料更新異常!"
                        If RtnCode = 9200 Then Msg = "[GetNextGate]下工程計算異常!"
                        If RtnCode = 9210 Then Msg = "[GetNextGate]無下工程管理人!"
                        If RtnCode = 9300 Then Msg = "[NextFlow]下一工程資料建置異常!"
                        If RtnCode = 9400 Then Msg = "[EndFlow]工程結束資料建置異常(999)!"
                        If RtnCode = 9999 Then Msg = "[BatchFlowControl]程式程序異常!"
                        SQL = "INSERT INTO L_AgentApprov " & _
                                  "Select FormNo,FormSno,Step,SeqNo,Fun, " & _
                                  "Action,Sts,StsDes,DecideCalendar,DecideID,ApplyID, " & _
                                  "AgentID,AllocteID,DecideDesc,TableName,UserIP,CreateTime, " & _
                                  "'" & xNextDecideID & "', " & _
                                  " " & xNextStep & " , " & _
                                  "'" & Msg & "', " & _
                                  "'" & NowDateTime & "' " & _
                                  "From Q_AgentApprov Where Unique_ID = " & dtAgentApprov.Rows(i)("Unique_ID").ToString & " "
                        uDataBase.ExecuteNonQuery(SQL)
                        '
                        '刪除代理簽核序列檔
                        SQL = "delete  From Q_AgentApprov"
                        SQL = SQL + " where Unique_ID = " & dtAgentApprov.Rows(i)("Unique_ID").ToString & " "
                        uDataBase.ExecuteNonQuery(SQL)
                        '
                        '發送異常MAIL
                        If RtnCode <> 0 Then
                            '
                            '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                            Send(dtAgentApprov.Rows(i)("DecideID").ToString, "IT013", dtAgentApprov.Rows(i)("ApplyID").ToString, _
                                    dtAgentApprov.Rows(i)("FormNo").ToString, dtAgentApprov.Rows(i)("FormSno"), dtAgentApprov.Rows(i)("Step"), "AAPPROV")
                            '
                            Send(dtAgentApprov.Rows(i)("DecideID").ToString, "IT003", dtAgentApprov.Rows(i)("ApplyID").ToString, _
                                    dtAgentApprov.Rows(i)("FormNo").ToString, dtAgentApprov.Rows(i)("FormSno"), dtAgentApprov.Rows(i)("Step"), "AAPPROV")
                        End If
                    Else
                        '
                        '重新再設定代理簽核
                        SQL = "update T_WaitHandle "
                        SQL = SQL + "set sts = 6, "
                        SQL = SQL + "stsdes ='代理簽核', "
                        SQL = SQL + "ModifyUser = '" & dtAgentApprov.Rows(i)("DecideID").ToString & "', "
                        SQL = SQL + "ModifyTime =getdate() "
                        SQL = SQL + "where DecideID = '" & dtAgentApprov.Rows(i)("DecideID").ToString & "' "
                        SQL = SQL + "AND FormNo = '" & dtAgentApprov.Rows(i)("FormNo").ToString & "' "
                        SQL = SQL + "and FormSNo = " & dtAgentApprov.Rows(i)("FormSno").ToString & " "
                        SQL = SQL + "and Step = " + dtAgentApprov.Rows(i)("Step").ToString & " "
                        SQL = SQL + "and SeqNo = " + dtAgentApprov.Rows(i)("SeqNo").ToString & " "
                        SQL = SQL + "and Active = '1' and flowtype=1  And (Sts = '0'  Or  Sts = '4')"
                        uDataBase.ExecuteNonQuery(SQL)
                        '
                        '通知系統管理人
                        Send(dtAgentApprov.Rows(i)("DecideID").ToString, "IT013", dtAgentApprov.Rows(i)("ApplyID").ToString, _
                                dtAgentApprov.Rows(i)("FormNo").ToString, dtAgentApprov.Rows(i)("FormSno"), dtAgentApprov.Rows(i)("Step"), "AAPPROV-1")
                        '
                        Send(dtAgentApprov.Rows(i)("DecideID").ToString, "IT003", dtAgentApprov.Rows(i)("ApplyID").ToString, _
                                dtAgentApprov.Rows(i)("FormNo").ToString, dtAgentApprov.Rows(i)("FormSno"), dtAgentApprov.Rows(i)("Step"), "AAPPROV-1")
                    End If
                    '
                End If
                '
            Next
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'AgentApprovStart-End

    '**************************************************************************************
    '** 標準代理簽核(AgentApprov)
    '**************************************************************************************
    'AgentApprov-Start
    <WebMethod()> _
    Public Function AgentApprov(ByVal pFun As String, _
                    ByVal pAction As Integer, _
                    ByVal pSts As String, _
                    ByVal pStsDes As String, _
                    ByVal pFormNo As String, _
                    ByVal pFormSno As Integer, _
                    ByVal pStep As Integer, _
                    ByVal pSeqNo As Integer, _
                    ByVal pDecideCalendar As String, _
                    ByVal pDecideID As String, _
                    ByVal pApplyID As String, _
                    ByVal pAgentID As String, _
                    ByVal pAllocteID As String, _
                    ByVal pDecideDesc As String, _
                    ByVal pTableName As String, _
                    ByRef pNextDecideID As String, _
                    ByRef pNextStep1 As Integer) As Integer
        Dim sFlow As New FlowService
        Dim sSchedule As New ScheduleService

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
                        DBObj.BAModifyTranData(pFun, pSts, pFormNo, pFormSno, pStep, pSeqNo, NowDateTime, pDecideID, pDecideDesc, pStsDes)

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
                    Dim wAllocateID As String = pAllocteID
                    '取得簽核者
                    '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
                    '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 
                    RtnCode = GetNextGate(pFormNo, pStep, pDecideID, pApplyID, pAgentID, wAllocateID, MultiJob, _
                                                  pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)

                    If RtnCode <> 0 Then ErrCode = 9200
                    If pCount = 0 And pNextStep <> 999 Then ErrCode = 9210
                    If ErrCode = 0 Then pAction = 0
                End If

                '建置流程資料
                If ErrCode = 0 And pRunNextStep = 1 Then
                    '代理簽核LOG (ADD)
                    pNextStep1 = pNextStep

                    If pNextStep <> 999 Then
                        For i = 1 To pCount
                            '代理簽核LOG (ADD)
                            pNextDecideID = pNextGate(1)

                            '取得核定者-群組行是曆
                            wNextGateCalendar = GetCalendarGroup(pNextGate(i))

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
                        DBObj.ModifyData(pFun, "1", pFormNo, pFormSno, NowDateTime, pDecideID, pTableName) '更新表單資料
                    Else
                        DBObj.ModifyData(pFun, "0", pFormNo, pFormSno, NowDateTime, pDecideID, pTableName) '更新表單資料 Sts=0(未結)
                    End If

                    '傳送郵件
                    If pNextStep <> 999 Then
                        For i = 1 To pCount
                            Send(pDecideID, pNextGate(i), pApplyID, pFormNo, NewFormSno, pNextStep, "FLOW")
                            '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                        Next i
                    Else
                        Send(pDecideID, pApplyID, pApplyID, pFormNo, NewFormSno, pNextStep, "END")
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
    'AgentApprov-End

    '**************************************************************************************
    '** 批次簽核(BatchFlowControl)
    '**************************************************************************************
    'BatchFlowControl-Start
    <WebMethod()> _
    Public Function BatchFlowControl(ByVal pFun As String, _
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
        Dim sFlow As New FlowService
        Dim sSchedule As New ScheduleService

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
                        DBObj.ModifyTranData(pFun, pSts, pFormNo, pFormSno, pStep, pSeqNo, NowDateTime, pDecideID)

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
                    RtnCode = GetNextGate(pFormNo, pStep, pDecideID, pApplyID, pAgentID, wAllocateID, MultiJob, _
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
                            wNextGateCalendar = GetCalendarGroup(pNextGate(i))

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
                        DBObj.ModifyData(pFun, "1", pFormNo, pFormSno, NowDateTime, pDecideID, pTableName) '更新表單資料
                    Else
                        DBObj.ModifyData(pFun, "0", pFormNo, pFormSno, NowDateTime, pDecideID, pTableName) '更新表單資料 Sts=0(未結)
                    End If

                    '傳送郵件
                    If pNextStep <> 999 Then
                        For i = 1 To pCount
                            Send(pDecideID, pNextGate(i), pApplyID, pFormNo, NewFormSno, pNextStep, "FLOW")
                            '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                        Next i
                    Else
                        Send(pDecideID, pApplyID, pApplyID, pFormNo, NewFormSno, pNextStep, "END")
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

    '**************************************************************************************
    '** 傳送郵件(執行外部程式)(SendMail)--停用中
    '**************************************************************************************
    'SendMail-Start
    <WebMethod()> _
    Public Function SendMail() As Integer
        Dim RtnCode As Integer = 0
        Try
            '停用中
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'SendMail-End

    '**************************************************************************************
    '** 產生郵件資料至序列檔(Send)
    '**************************************************************************************
    'Send-Start
    <WebMethod()> _
    Public Function Send(ByVal pFrom As String, _
                         ByVal pTo As String, _
                         ByVal pApplyID As String, _
                         ByVal pFormNo As String, _
                         ByVal pFormSno As Integer, _
                         ByVal pStep As Integer, _
                         ByVal pType As String) As Integer
        Dim RtnCode As Integer = 0
        Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")
        Dim SQL As String
        Dim xFlowType As String = ""
        Dim xStepname As String = ""
        Dim xNo As String = ""
        Dim xFormName As String = ""
        Dim xFromMail As String = ""
        Dim xFromName As String = ""
        Dim xToMail As String = ""
        Dim xToName As String = ""
        Dim xApplyName As String = ""
        Dim xCCMail As String = ""
        Try
            '取得訊息類別名稱,工程名稱
            If UCase(pType) = "FLOW" Then
                SQL = "Select FlowType, StepName From M_Flow "
                SQL &= "Where FormNo = '" & pFormNo & "' "
                SQL &= "  And Step   = '" & CStr(pStep) & "' "
                SQL &= "  And Active = '1' "
                Dim dt_Flow As DataTable = uDataBase.GetDataTable(SQL)
                If dt_Flow.Rows.Count > 0 Then
                    Select Case dt_Flow.Rows(0).Item("FlowType")
                        Case Is = 0
                            xFlowType = "通知"
                        Case Is = 1
                            xFlowType = "核定"
                        Case Is = 2
                            xFlowType = "申請"
                        Case Is = 3
                            xFlowType = "核定"
                    End Select
                    xStepname = dt_Flow.Rows(0).Item("StepName").ToString
                End If
            End If
            If UCase(pType) = "END" Then xFlowType = "結束"
            If UCase(pType) = "DELAY" Then xFlowType = "延遲"
            If UCase(pType) = "CANCELSTEP" Then xFlowType = "抽單-單一流程"
            If UCase(pType) = "CANCELALL" Then xFlowType = "抽單-整張單"
            '##
            '##AgentApprov-Start
            If UCase(pType) = "AAPPROV" Then xFlowType = "代簽-異常(L_AgentApprov)"
            If UCase(pType) = "AAPPROV-1" Then xFlowType = "代簽-異常(WAITHANDLE.sts<>6)"
            '##AgentApprov-End
            '##
            '委託No.
            SQL = "SELECT No FROM T_CommissionNo "
            SQL &= "Where FormNo  = '" & pFormNo & "' "
            SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
            Dim dt_Mail As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Mail.Rows.Count > 0 Then
                xNo = dt_Mail.Rows(0).Item("No").ToString
            End If
            '表單名稱
            SQL = "SELECT FormName FROM M_Form "
            SQL &= "Where FormNo  = '" & pFormNo & "' "
            dt_Mail = uDataBase.GetDataTable(SQL)
            If dt_Mail.Rows.Count > 0 Then
                xFormName = dt_Mail.Rows(0).Item("FormName").ToString
            End If
            '寄件者郵件帳號,姓名
            SQL = "Select Mail, UserName From M_Users "
            SQL &= "Where UserID  = '" & pFrom & "' "
            dt_Mail = uDataBase.GetDataTable(SQL)
            If dt_Mail.Rows.Count > 0 Then
                xFromMail = dt_Mail.Rows(0).Item("Mail").ToString
                xFromName = dt_Mail.Rows(0).Item("UserName").ToString
            End If
            '收件者郵件帳號,姓名
            SQL = "Select Mail, UserName From M_Users "
            SQL &= "Where UserID  = '" & pTo & "' "
            dt_Mail = uDataBase.GetDataTable(SQL)
            If dt_Mail.Rows.Count > 0 Then
                xToMail = dt_Mail.Rows(0).Item("Mail").ToString
                xToName = dt_Mail.Rows(0).Item("UserName").ToString
            End If
            '申請者姓名
            xApplyName = GetUserName(pApplyID)
            'cc者
            xCCMail = GetSystemData("013", "ADMADR")
            If GetSystemData("013", "USERADR1") <> "" Then
                xCCMail = xCCMail & ";" & GetSystemData("013", "USERADR1")
            End If
            If GetSystemData("013", "USERADR2") <> "" Then
                xCCMail = xCCMail & ";" & GetSystemData("013", "USERADR2")
            End If
            '產生郵件資料至序列檔
            SQL = "Insert into Q_WaitSend "
            SQL &= "( "
            SQL &= "Sts, FromID, FromMail, FromName, ToID, "
            SQL &= "ToMail, ToName, CCMail, "
            SQL &= "FormNo, FormSno, FormName, Step, StepName, "
            SQL &= "ApplyID, ApplyName, MSG, MSGName, No, "
            SQL &= "CreateTime "
            SQL &= ") "
            SQL &= "VALUES( "
            SQL &= "'0' ,"                              '控制狀態(0:完成,1:待處理)
            SQL &= "'" & pFrom & "' ,"                  '寄件者ID
            SQL &= "'" & xFromMail & "' ,"              '寄件者郵件帳號
            SQL &= "'" & xFromName & "' ,"              '寄件者姓名
            SQL &= "'" & pTo & "' ,"                    '收件者ID
            SQL &= "'" & xToMail & "' ,"                '收件者郵件帳號
            SQL &= "'" & xToName & "' ,"                '收件者姓名
            SQL &= "'" & xCCMail & "' ,"                'cc者郵件帳號
            SQL &= "'" & pFormNo & "' ,"                '表單No
            SQL &= "'" & CStr(pFormSno) & "' ,"         '表單序號
            SQL &= "'" & xFormName & "' ,"              '表單名稱
            SQL &= "'" & CStr(pStep) & "' ,"            '工程No
            SQL &= "'" & xStepname & "' ,"              '工程名稱
            SQL &= "'" & pApplyID & "' ,"               '申請者ID
            SQL &= "'" & xApplyName & "' ,"             '申請者姓名
            SQL &= "'" & pType & "' ,"                  '訊息類別代碼
            SQL &= "'" & xFlowType & "' ,"              '訊息類別名稱
            SQL &= "'" & xNo & "' ,"                    '委託No
            SQL &= "'" & NowDateTime & "' "             '製作時間
            SQL &= ") "
            uDataBase.ExecuteNonQuery(SQL)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'Send-End
    '**************************************************************************************
    '** 取得下工程簽核者(GetNextGate)
    '**************************************************************************************
    'GetNextGate-Start
    <WebMethod()> _
    Public Function GetNextGate(ByVal pFormNo As String, _
                                ByVal pStep As Integer, _
                                ByVal pUserID As String, _
                                ByVal pApplyID As String, _
                                ByVal pAgentID As String, _
                                ByVal pAllocateID As String, _
                                ByVal pMultiJob As Integer, _
                                ByRef pNextStep As Integer, _
                                ByRef pNextGate() As String, _
                                ByRef pAgentGate() As String, _
                                ByRef pCount As Integer, _
                                ByRef pFlowType As Integer, _
                                ByVal pAction As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")
        Dim SQL As String
        Dim i, xNextStep As Integer
        Dim xRelatedID, xID, xUserID As String
        '特殊處理[S001]-參數
        Dim S001_NextStep As Integer = pNextStep
        '設定初始值
        pNextStep = 0       '下工程No.
        pCount = 0          '下工程簽核人數
        pFlowType = 1       '處理方法(0:核定,1:通知,2:申請)
        For i = 1 To 10     '下工程簽核者
            pNextGate(i) = ""
            pAgentGate(i) = ""
        Next
        Try
            'MsgBox("Get NextStep-In")
            '取得下工程No.
            SQL = "Select RelatedType, RelatedID, FlowTypeStep, NextStep, JumpStep, Condition, ConditionIdx From M_Flow "
            SQL &= "Where FormNo = '" & pFormNo & "' "
            SQL &= "  And Step   = '" & CStr(pStep) & "' "
            SQL &= "  And Active = '1' "
            SQL &= "  And Action = '" & CStr(pAction) & "' "
            SQL &= "Order by Action "
            Dim dt_Flow As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Flow.Rows.Count > 0 Then
                '標準方式
                xNextStep = dt_Flow.Rows(0).Item("NextStep")                           '##下一工程No.
                '多人核定方式
                If pMultiJob = 1 Then xNextStep = dt_Flow.Rows(0).Item("FlowTypeStep") '##下一工程No.
                '自訂條件方式
                If dt_Flow.Rows(0).Item("Condition").ToString <> "" Then        '有自訂條件
                    'MsgBox("[" + dt_Flow.Rows(0).Item("Condition").ToString + "]")
                    '----------------------------------------------------------------
                    '特殊處理[S001]-跳至指定工程(pMultiJob=FormSno, pNextStep=指定工程)  
                    If dt_Flow.Rows(0).Item("Condition").ToString = "[S]-001" Then
                        SQL = "Select Top 1 Step, DecideID From T_WaitHandle "
                        SQL &= "Where FormNo  =  '" & pFormNo & "' "
                        SQL &= "  And FormSno =  '" & CStr(pMultiJob) & "' "         '因沒欄位可使用所以使用MultiJob
                        SQL &= "  And Step    = '" & CStr(S001_NextStep) & "' "
                        SQL &= "  And Active  = '0' "
                        SQL &= "  And FlowType <> '0' "
                        SQL &= "Order by Unique_ID Desc "
                        Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(SQL)
                        If dtWaitHandle.Rows.Count > 0 Then
                            xNextStep = dtWaitHandle.Rows(0).Item("Step")          '下一工程No.
                            pAllocateID = dtWaitHandle.Rows(0).Item("DecideID")    '下一工程者
                        End If
                    End If
                    '--非特殊處理--------------------------------------------------------------
                    If InStr(dt_Flow.Rows(0).Item("Condition").ToString, "[S]-") <= 0 Then
                        '--關係人自訂條件
                        If dt_Flow.Rows(0).Item("RelatedType") = 0 Then             '關係人
                            xRelatedID = dt_Flow.Rows(0).Item("RelatedID").ToString '關係人ID
                        Else                                                        '其他
                            xRelatedID = ""
                        End If
                        '關係人ID=台北(A)
                        If (xRelatedID = "A") And (pApplyID <> pUserID) Then
                            If pAgentID <> "" Then                                  '是否為代理
                                xID = pAgentID                                      '被代理者
                            Else
                                xID = pUserID                                       '簽核者
                            End If
                            SQL = "Select UserID From M_Condition "
                            SQL &= "Where UserID = '" & pApplyID & "' "
                            SQL &= "  And RUserID   = '" & xID & "' "
                            SQL &= "  And " & dt_Flow.Rows(0).Item("Condition").ToString
                            SQL &= "  And RelatedID = '" & xRelatedID & "' "
                            SQL &= "  And Active = '1' "
                            Dim dt_Condition As DataTable = uDataBase.GetDataTable(SQL)
                            If dt_Condition.Rows.Count > 0 Then xNextStep = dt_Flow.Rows(0).Item("JumpStep") '##下一工程No.
                        Else
                            '關係人ID=中壢(B)
                            If dt_Flow.Rows(0).Item("ConditionIdx") = 0 Then        '自訂條件對象(0=簽核者, 1=申請者)
                                xID = pUserID
                            Else
                                xID = pApplyID
                            End If
                            SQL = "Select UserID From M_Condition "
                            SQL &= "Where UserID = '" & xID & "' "
                            SQL &= "  And " & dt_Flow.Rows(0).Item("Condition").ToString
                            SQL &= "  And RelatedID = '" & xRelatedID & "' "
                            SQL &= "  And Active = '1' "
                            Dim dt_Condition As DataTable = uDataBase.GetDataTable(SQL)
                            If dt_Condition.Rows.Count > 0 Then xNextStep = dt_Flow.Rows(0).Item("JumpStep") '##下一工程No.
                        End If
                    End If
                End If
            Else
                RtnCode = 1
            End If
            'MsgBox("Get NextStep-Out" & Chr(13) & "NextStep=[" + CStr(xNextStep) + "]")
            '取得下工程簽核者及人數
            If RtnCode = 0 And xNextStep > 0 And xNextStep < 999 Then   'xNextStep=0(無下工程), =999(結束工程)
                SQL = "Select SignImage, RelatedID, RelatedType, FlowType From M_Flow "
                SQL &= "Where FormNo = '" & pFormNo & "' "
                SQL &= "  And Step   = '" & CStr(xNextStep) & "' "
                SQL &= "  And Active = '1' "
                SQL &= "  And Action = '0' "
                SQL &= "Order by Action "
                dt_Flow = uDataBase.GetDataTable(SQL)
                If dt_Flow.Rows.Count > 0 Then
                    pNextStep = xNextStep                               '##下一關No
                    pFlowType = dt_Flow.Rows(0).Item("FlowType")        '##處理方法(0:核定,1:通知,2:申請)
                    If dt_Flow.Rows(0).Item("SignImage") = 0 Then       '0:標準, 1:強制跳關主
                        '關主類別:0:關係人, 1:使用者, 2:群組, 3:申請者, 4:無
                        Select Case dt_Flow.Rows(0).Item("RelatedType")
                            Case Is = 0                                 '關係人
                                If pAgentID <> "" Then                              '是否為代理
                                    xID = pAgentID                                  '被代理者
                                Else
                                    xID = pUserID                                   '簽核者
                                End If
                                SQL = "Select RelatedID, UserID, RUserID, RRUserID From M_Related "
                                SQL &= "Where RelatedID = '" & dt_Flow.Rows(0).Item("RelatedID").ToString & "' "
                                SQL &= "  And Active = '1' "
                                If dt_Flow.Rows(0).Item("RelatedID").ToString = "A" And pApplyID <> pUserID And pApplyID <> pAgentID Then
                                    SQL &= "  And UserID  = '" & pApplyID & "' "
                                    SQL &= "  And RUserID = '" & xID & "' "
                                Else
                                    SQL &= "  And UserID  = '" & xID & "' "
                                End If
                            Case Is = 1                                 '使用者
                                SQL = "Select UserID From M_Users "
                                SQL &= "Where UserID = '" & dt_Flow.Rows(0).Item("RelatedID").ToString & "' "
                                SQL &= "  And Active = '1' "
                            Case Is = 2                                 '群組
                                SQL = "Select UserID From M_Group "
                                SQL &= "Where GroupID = '" & dt_Flow.Rows(0).Item("RelatedID").ToString & "' "
                                SQL &= "  And Active = '1' "
                            Case Is = 3                                 '申請者
                                SQL = "Select UserID From M_Users "
                                SQL &= "Where UserID = '" & pApplyID & "' "
                                SQL &= "  And Active = '1' "
                            Case Is = 4                                 '無
                        End Select
                        Dim dt_NextGate As DataTable = uDataBase.GetDataTable(SQL)
                        '
                        For i = 0 To dt_NextGate.Rows.Count - 1
                            If dt_Flow.Rows(0).Item("RelatedType") = 0 Then     '關係人
                                '台北(A)
                                If dt_NextGate.Rows(i).Item("RelatedID").ToString = "A" Then
                                    If (pApplyID <> pUserID) And (pApplyID <> pAgentID) Then
                                        xUserID = dt_NextGate.Rows(i).Item("RRUserID").ToString
                                    Else
                                        xUserID = dt_NextGate.Rows(i).Item("RUserID").ToString
                                    End If
                                End If
                                '中壢(B)
                                If dt_NextGate.Rows(i).Item("RelatedID").ToString = "B" Then
                                    xUserID = dt_NextGate.Rows(i).Item("RUserID").ToString
                                End If
                                '報廢關係人
                                If dt_NextGate.Rows(i).Item("RelatedID").ToString = "D" Then
                                    xUserID = dt_NextGate.Rows(i).Item("RUserID").ToString
                                End If
                                'MOD-START BY JOY 2023/1/16
                                '共通(E)
                                If dt_NextGate.Rows(i).Item("RelatedID").ToString = "E" Then
                                    xUserID = dt_NextGate.Rows(i).Item("RUserID").ToString
                                End If
                                'MOD-END
                                '
                                'MOD-START BY Jessica 2023/5/23
                                '出差及清算
                                If dt_NextGate.Rows(i).Item("RelatedID").ToString = "F" Then
                                    xUserID = dt_NextGate.Rows(i).Item("RUserID").ToString
                                End If
                                'MOD-END
                            Else        '非關係人
                                xUserID = dt_NextGate.Rows(i).Item("UserID").ToString
                            End If
                            '取得代理人
                            If GetAgentID(pFormNo, xUserID, NowDateTime) <> "" Then
                                pAgentGate(i + 1) = xUserID          '##被代理者
                                xUserID = GetAgentID(pFormNo, xUserID, NowDateTime)
                            End If
                            '取得原職務人(兼職)
                            'Modify-Start By Joy 2009/12/3
                            'If GetMainJobID(xUserID) <> "" Then
                            '    xUserID = GetMainJobID(xUserID)
                            '    '取得代理人
                            '    If GetAgentID(pFormNo, xUserID, NowDateTime) <> "" Then
                            '        pAgentGate(i + 1) = xUserID     '##被代理者
                            '        xUserID = GetAgentID(pFormNo, xUserID, NowDateTime)
                            '    End If
                            'End If
                            '
                            If GetMainJobID(xUserID) <> "" Then
                                pAgentGate(i + 1) = xUserID     '##兼職
                                xUserID = GetMainJobID(xUserID)
                            End If
                            'Modify-End
                            pNextGate(i + 1) = xUserID          '##簽核者
                        Next
                        pCount = dt_NextGate.Rows.Count     '##人數
                    Else        '1:強制跳關主
                        xUserID = pAllocateID
                        '取得代理人
                        If GetAgentID(pFormNo, xUserID, NowDateTime) <> "" Then
                            pAgentGate(1) = xUserID          '##被代理者
                            xUserID = GetAgentID(pFormNo, xUserID, NowDateTime)
                        End If
                        '取得原職務人(兼職)
                        'Modify-Start By Joy 2009/12/3
                        'If GetMainJobID(xUserID) <> "" Then
                        '    xUserID = GetMainJobID(xUserID)
                        '    '取得代理人
                        '    If GetAgentID(pFormNo, xUserID, NowDateTime) <> "" Then
                        '        pAgentGate(1) = xUserID      '##被代理者
                        '        xUserID = GetAgentID(pFormNo, xUserID, NowDateTime)
                        '    End If
                        'End If
                        '
                        If GetMainJobID(xUserID) <> "" Then
                            pAgentGate(1) = xUserID     '##兼職
                            xUserID = GetMainJobID(xUserID)
                        End If
                        'Modify-End
                        pNextGate(1) = xUserID              '##簽核者
                        pCount = 1                          '##人數
                    End If
                Else
                    RtnCode = 1
                End If
            Else
                If xNextStep >= 999 Then pNextStep = xNextStep
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        'MsgBox("NextGate=[" + pNextGate(1) + "]" + Chr(13) + "AgentGate=[" + pAgentGate(1) + "]")
        Return RtnCode
    End Function
    'GetNextGate-End
    '**************************************************************************************
    '** 取得原職務人(GetMainJobID)
    '**************************************************************************************
    'GetMainJobID-Start
    <WebMethod()> _
    Public Function GetMainJobID(ByVal pUser As String) As String
        Dim RtnString As String = ""
        Dim SQL As String
        Try
            SQL = "Select Pluralism, OriUserID From M_Users "
            SQL &= "Where Active  = '1' "
            SQL &= "  And UserID  = '" & pUser & "' "
            Dim dt_Users As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Users.Rows.Count > 0 Then
                If dt_Users.Rows(0).Item("Pluralism") = 1 Then
                    RtnString = dt_Users.Rows(0).Item("OriUserID").ToString
                End If
            End If
        Catch ex As Exception
            RtnString = ""
        End Try
        '
        Return RtnString
    End Function
    'GetMainJobID-End
    '**************************************************************************************
    '** 取得代理人(GetAgentID)
    '**************************************************************************************
    'GetAgentID-Start
    <WebMethod()> _
    Public Function GetAgentID(ByVal pFormNo As String, ByVal pUser As String, ByVal pNowDateTime As String) As String
        Dim RtnString As String = ""
        Dim SQL As String
        Try
            SQL = "Select AgentID From M_Agent "
            SQL &= "Where Active  = '1' "
            SQL &= "  And UserID  = '" & pUser & "' "
            SQL &= "  And AllForm = '0' "
            SQL &= "  And StartDate <= '" & pNowDateTime & "' "
            SQL &= "  And EndDate >= '" & pNowDateTime & "' "
            Dim dt_Agent As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Agent.Rows.Count > 0 Then
                RtnString = dt_Agent.Rows(0).Item("AgentID").ToString
            Else
                SQL = "Select AgentID From M_Agent "
                SQL &= "Where Active  = '1' "
                SQL &= "  And UserID  = '" & pUser & "' "
                SQL &= "  And AllForm = '1' "
                SQL &= "  And FormNo  = '" & pFormNo & "' "
                SQL &= "  And StartDate <= '" & pNowDateTime & "' "
                SQL &= "  And EndDate >= '" & pNowDateTime & "' "
                dt_Agent = uDataBase.GetDataTable(SQL)
                If dt_Agent.Rows.Count > 0 Then
                    RtnString = dt_Agent.Rows(0).Item("AgentID").ToString
                End If
            End If
        Catch ex As Exception
            RtnString = ""
        End Try
        '
        Return RtnString
    End Function
    'GetAgentID-End
    '**************************************************************************************
    '** 取得表單可使用序號(GetFormSeqNo)
    '**************************************************************************************
    'GetFormSeqNo-Start
    <WebMethod()> _
    Public Function GetFormSeqNo(ByVal pFormNo As String, ByRef pSeqNo As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        Dim xTableName As String
        Dim xSeqNo As Integer
        Dim xHaveSeqNo As Boolean = False
        Try
            '--2016/8/20 變更   not check T_WaitHandle
            '
            SQL = "Select FormSno, TableName1 From M_Form "
            SQL &= "Where FormNo  = '" & pFormNo & "' "
            SQL &= "  And Active = '1' "
            Dim dt_FormSeq As DataTable = uDataBase.GetDataTable(SQL)
            If dt_FormSeq.Rows.Count > 0 Then
                xTableName = "F_" & dt_FormSeq.Rows(0).Item("TableName1").ToString
                xSeqNo = dt_FormSeq.Rows(0).Item("FormSno")
                Do Until xHaveSeqNo = True
                    SQL = "Select FormNo From " & xTableName
                    SQL &= " Where FormNo  = '" & pFormNo & "' "
                    SQL &= "   And FormSno = '" & CStr(xSeqNo) & "' "
                    Dim dt_Form = uDataBase.GetDataTable(SQL)
                    If dt_Form.Rows.Count <= 0 Then xHaveSeqNo = True
                    If xHaveSeqNo = False Then xSeqNo = xSeqNo + 1
                Loop
            Else
                RtnCode = 1
            End If
            '更新表單可使用序號
            If RtnCode = 0 And xHaveSeqNo = True Then
                SQL = "Update M_Form Set "
                SQL &= " FormSno = '" & CStr(xSeqNo + 1) & "' "
                SQL &= "Where FormNo  = '" & pFormNo & "' "
                uDataBase.ExecuteNonQuery(SQL)
                pSeqNo = xSeqNo
            End If
            '
            '--2016/8/20 變更前
            '
            'SQL = "Select FormSno, TableName1 From M_Form "
            'SQL &= "Where FormNo  = '" & pFormNo & "' "
            'SQL &= "  And Active = '1' "
            'Dim dt_FormSeq As DataTable = uDataBase.GetDataTable(SQL)
            'If dt_FormSeq.Rows.Count > 0 Then
            '    xTableName = "F_" & dt_FormSeq.Rows(0).Item("TableName1").ToString
            '    xSeqNo = dt_FormSeq.Rows(0).Item("FormSno")
            '    Do Until xHaveSeqNo = True
            '        SQL = "Select FormNo From T_WaitHandle "
            '        SQL &= " Where FormNo  = '" & pFormNo & "' "
            '        SQL &= "   And FormSno = '" & CStr(xSeqNo) & "' "
            '        Dim dt_WaitHandle = uDataBase.GetDataTable(SQL)
            '        If dt_WaitHandle.Rows.Count <= 0 Then
            '            SQL = "Select FormNo From " & xTableName
            '            SQL &= " Where FormNo  = '" & pFormNo & "' "
            '            SQL &= "   And FormSno = '" & CStr(xSeqNo) & "' "
            '            Dim dt_Form = uDataBase.GetDataTable(SQL)
            '            If dt_Form.Rows.Count <= 0 Then xHaveSeqNo = True
            '        End If
            '        If xHaveSeqNo = False Then xSeqNo = xSeqNo + 1
            '    Loop
            'Else
            '    RtnCode = 1
            'End If
            ''更新表單可使用序號
            'If RtnCode = 0 And xHaveSeqNo = True Then
            '    SQL = "Update M_Form Set "
            '    SQL &= " FormSno = '" & CStr(xSeqNo + 1) & "' "
            '    SQL &= "Where FormNo  = '" & pFormNo & "' "
            '    uDataBase.ExecuteNonQuery(SQL)
            '    pSeqNo = xSeqNo
            'End If
        Catch ex As Exception
            RtnCode = 2
        End Try
        '
        Return RtnCode
    End Function
    'GetFormSeqNo-End
    '**************************************************************************************
    '** 取得表單欄位屬性(GetFieldAttribute)
    '**************************************************************************************
    'GetFieldAttribute-Start
    <WebMethod()> _
    Public Function GetFieldAttribute(ByVal pFormNo As String, ByVal pStep As Integer, _
                                      ByRef pFieldName() As String, _
                                      ByRef pAttribute() As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        Dim i As Integer = 0
        Dim xFieldName, xAttribute As String
        Try
            '設定初始值
            For i = 1 To 60
                pFieldName(i) = ""
                pAttribute(i) = 0
            Next
            '取得表單欄位屬性
            SQL = "Select * From M_FormField "
            SQL &= "Where FormNo = '" & pFormNo & "' "
            SQL &= "  And Step   = '" & CStr(pStep) & "' "
            SQL &= "  And Active = '1' "
            Dim dt_FormField As DataTable = uDataBase.GetDataTable(SQL)
            If dt_FormField.Rows.Count > 0 Then
                For i = 1 To 60
                    xFieldName = "FieldName" + CStr(i)
                    xAttribute = "Attribute" + CStr(i)
                    pFieldName(i) = dt_FormField.Rows(0).Item(xFieldName).ToString
                    pAttribute(i) = dt_FormField.Rows(0).Item(xAttribute)
                Next
            Else
                RtnCode = 1
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'GetFieldAttribute-End
    '**************************************************************************************
    '** 委託No.檢查(CommissionNo)
    '**************************************************************************************
    'CommissionNo-Start
    <WebMethod()> _
    Public Function CommissionNo(ByVal pFormNo As String, ByVal pFormSno As Integer, ByVal pStep As Integer, ByVal pNo As String) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL, xTableName As String
        Try
            If GetTableName(pFormNo) <> "" Then
                xTableName = GetTableName(pFormNo)
                '檢查表單之委託No.是否有重覆
                If pFormSno = 0 Then
                    SQL = "SELECT No From " & xTableName
                    SQL &= " Where No = '" & pNo & "' "
                    Dim dt_FormTable As DataTable = uDataBase.GetDataTable(SQL)
                    If dt_FormTable.Rows.Count > 0 Then
                        RtnCode = 1
                    End If
                Else
                    SQL = "SELECT No From " & xTableName
                    SQL &= " Where No = '" & pNo & "' "
                    SQL &= "   And FormSno <> '" & CStr(pFormSno) & "' "
                    Dim dt_FormTable As DataTable = uDataBase.GetDataTable(SQL)
                    If dt_FormTable.Rows.Count > 0 Then
                        RtnCode = 1
                    End If
                End If
            Else
                RtnCode = 1
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'CommissionNo-End
    '**************************************************************************************
    '** 取得工程LeadTime(GetOPLeadTime)
    '**************************************************************************************
    'GetOPLeadTime-Start
    <WebMethod()> _
    Public Function GetOPLeadTime(ByVal pFormNo As String, ByVal pStep As Integer, ByVal pLevel As String) As Integer
        Dim RtnValue As Integer = 0
        Dim SQL As String
        Try
            SQL = "Select sum(Hours) as Hours From M_LeadTime "
            SQL &= "Where FormNo = '" & pFormNo & "' "
            SQL &= "  And Step   = '" & CStr(pStep) & "' "
            SQL &= "  And Level  = '" & pLevel & "' "
            SQL &= "  And Active = '1' "
            Dim dt_LeadTime As DataTable = uDataBase.GetDataTable(SQL)
            If dt_LeadTime.Rows.Count > 0 Then
                RtnValue = dt_LeadTime.Rows(0).Item("Hours")
            End If
        Catch ex As Exception
            RtnValue = 0
        End Try
        '
        'MsgBox("time=" + CStr(RtnValue))
        Return RtnValue
    End Function
    'GetOPLeadTime-End
    '**************************************************************************************
    '** 取得系統參數(GetSystemData)
    '**************************************************************************************
    'GetSystemData-Start
    <WebMethod()> _
    Public Function GetSystemData(ByVal pCat As String, ByVal pDKey As String) As String
        Dim RtnString As String = ""
        Dim SQL As String
        Try
            SQL = "Select Data From M_Referp "
            SQL &= "Where Cat  = '" & pCat & "' "
            SQL &= "  And DKey = '" & pDKey & "' "
            Dim dt_Referp As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Referp.Rows.Count > 0 Then
                RtnString = dt_Referp.Rows(0).Item("Data").ToString
            End If
        Catch ex As Exception
            RtnString = ""
        End Try
        '
        Return RtnString
    End Function
    'GetSystemData-End
    '**************************************************************************************
    '** 取得QCLeadTime(GetQCLeadTime)
    '**************************************************************************************
    'GetQCLeadTime-Start
    <WebMethod()> _
    Public Function GetQCLeadTime(ByVal pFormNo As String, ByVal pStep As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        Try
            SQL = "Select LeadTime From M_Flow "
            SQL &= "Where Active  = '1' "
            SQL &= "  And FormNo  = '" & pFormNo & "' "
            SQL &= "  And Step    = '" & CStr(pStep) & "' "
            Dim dt_Flow As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Flow.Rows.Count > 0 Then
                RtnCode = dt_Flow.Rows(0).Item("LeadTime")
            End If
        Catch ex As Exception
            RtnCode = 0
        End Try
        '
        Return RtnCode
    End Function
    'GetQCLeadTime-End
    '**************************************************************************************
    '** 取得表單檔案名稱(GetTableName)
    '**************************************************************************************
    'GetTableName-Start
    <WebMethod()> _
    Public Function GetTableName(ByVal pFormNo As String) As String
        Dim RtnString As String = ""
        Dim SQL As String
        Try
            SQL = "Select TableName1 From M_Form "
            SQL &= "Where Active  = '1' "
            SQL &= "  And FormNo  = '" & pFormNo & "' "
            Dim dt_Form As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Form.Rows.Count > 0 Then
                RtnString = "F_" & dt_Form.Rows(0).Item("TableName1").ToString
            End If
        Catch ex As Exception
            RtnString = ""
        End Try
        '
        Return RtnString
    End Function
    'GetTableName-End
    '**************************************************************************************
    '** 取得使用者姓名(GetUserName)
    '**************************************************************************************
    'GetUserName-Start
    <WebMethod()> _
    Public Function GetUserName(ByVal pUser As String) As String
        Dim RtnString As String = ""
        Dim SQL As String
        Try
            SQL = "Select UserName From M_Users "
            SQL &= "Where UserID  = '" & pUser & "' "
            Dim dt_Users As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Users.Rows.Count > 0 Then
                RtnString = dt_Users.Rows(0).Item("UserName").ToString
            End If
        Catch ex As Exception
            RtnString = ""
        End Try
        '
        Return RtnString
    End Function
    'GetUserName-End
    '**************************************************************************************
    '** 取得使用者ID(GetUserID)
    '**************************************************************************************
    'GetUserID-Start
    <WebMethod()> _
    Public Function GetUserID(ByVal pUserName As String) As String
        Dim RtnString As String = ""
        Dim SQL As String
        Try
            SQL = "Select UserID From M_Users "
            SQL &= "Where UserName  = N'" & pUserName & "' "
            Dim dt_Users As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Users.Rows.Count > 0 Then
                RtnString = dt_Users.Rows(0).Item("UserID").ToString
            End If
        Catch ex As Exception
            RtnString = ""
        End Try
        '
        Return RtnString
    End Function
    'GetUserID-End
    '**************************************************************************************
    '** 取得群組行事曆(GetCalendarGroup)  Add-2009/11/20 by Joy
    '**************************************************************************************
    'GetCalendarGroup-Start
    <WebMethod()> _
    Public Function GetCalendarGroup(ByVal pUser As String) As String
        Dim RtnString As String = ""
        Dim SQL As String
        Try
            SQL = "Select DepoID, EmpID From M_Users "
            SQL &= "Where UserID  = '" & pUser & "' "
            Dim dt_Users As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Users.Rows.Count > 0 Then
                '
                SQL = "Select CalendarGroup From M_EMP "
                SQL &= "Where Com_Code = '" & dt_Users.Rows(0).Item("DepoID").ToString & "' "
                SQL &= "  And ID       = '" & dt_Users.Rows(0).Item("EmpID").ToString & "' "
                Dim dt_EMP As DataTable = uDataBase.GetDataTable(SQL)
                If dt_EMP.Rows.Count > 0 Then
                    RtnString = dt_EMP.Rows(0).Item("CalendarGroup").ToString
                Else
                    If dt_Users.Rows(0).Item("DepoID").ToString() = "01" Or _
                       dt_Users.Rows(0).Item("DepoID").ToString() = "60" Or _
                       dt_Users.Rows(0).Item("DepoID").ToString() = "65" Then RtnString = "CL2"
                    '
                    If dt_Users.Rows(0).Item("DepoID").ToString() = "10" Or _
                       dt_Users.Rows(0).Item("DepoID").ToString() = "11" Then RtnString = "TP1"
                End If
                '
            End If
        Catch ex As Exception
            RtnString = ""
        End Try
        '
        Return RtnString
    End Function
    'GetCalendarGroup-End
    '**************************************************************************************
    '** 取得工作ID(GetWorkID)
    '**************************************************************************************
    'GetWorkID-Start
    <WebMethod()> _
    Public Function GetWorkID(ByVal pUser As String) As String
        Dim RtnString As String = "999999"
        Dim SQL As String
        Try
            SQL = "Select WorkID From M_Users "
            SQL &= "Where UserID  = '" & pUser & "' "
            Dim dt_Users As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Users.Rows.Count > 0 Then
                RtnString = dt_Users.Rows(0).Item("WorkID").ToString
            End If
        Catch ex As Exception
            RtnString = "999999"
        End Try
        '
        Return RtnString
    End Function
    'GetWorkID-End
    '**************************************************************************************
    '** 取得任一簽核(GetSignType)   0:否, 1:是
    '**************************************************************************************
    'GetSignType-Start
    <WebMethod()> _
    Public Function GetSignType(ByVal pFormNo As String, ByVal pStep As Integer) As Integer
        Dim RtnCode As Integer = 1
        Dim SQL As String
        Try
            SQL = "Select EveryOne From M_Flow "  '取得簽核種類   任一簽核 0:否, 1:是
            SQL &= "Where Active = '1' "
            SQL &= "  And FormNo = '" & pFormNo & "' "
            SQL &= "  And Step   = '" & CStr(pStep) & "' "
            Dim dt_Flow As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Flow.Rows.Count > 0 Then
                RtnCode = dt_Flow.Rows(0).Item("EveryOne")
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'GetSignType-End
    '**************************************************************************************
    '** 取得流程種類(0:通知,1:核定,2:申請)(GetFlowType)  
    '**************************************************************************************
    'GetFlowType-Start
    <WebMethod()> _
    Public Function GetFlowType(ByVal pFormNo As String, ByVal pStep As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        Try
            SQL = "Select FlowType From M_Flow "
            SQL &= "Where FormNo  = '" & pFormNo & "' "
            SQL &= "  And Step    = '" & CStr(pStep) & "' "
            SQL &= "  And Active  = '" & "1" & "' "
            Dim dt_Flow As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Flow.Rows.Count > 0 Then
                RtnCode = dt_Flow.Rows(0).Item("FlowType")
                If RtnCode = 3 Then RtnCode = 1
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'GetFlowType-End
    '**************************************************************************************
    '** 取得表單申請時間(GetApplyTime)  
    '**************************************************************************************
    'GetApplyTime-Start
    <WebMethod()> _
    Public Function GetApplyTime(ByVal pFormNo As String, ByVal pFormSno As Integer, ByVal pNowDateTime As String) As String
        Dim RtnString As String = pNowDateTime
        Dim SQL As String
        Try
            SQL = "Select ApplyTime From T_WaitHandle "   '申請時間
            SQL &= "Where FormNo  = '" & pFormNo & "' "
            SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
            SQL &= "  And Step    < '10' "
            SQL &= "Order by Step, Seqno "
            Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
            If dt_WaitHandle.Rows.Count > 0 Then
                RtnString = CDate(dt_WaitHandle.Rows(0).Item("ApplyTime")).ToString("yyyy/MM/dd HH:mm:ss")
            End If
        Catch ex As Exception
            RtnString = pNowDateTime
        End Try
        '
        Return RtnString
    End Function
    'GetApplyTime-End
    '**************************************************************************************
    '** 取得核定履歷(GetDecideHistory)  
    '**************************************************************************************
    'GetDecideHistory-Start
    <WebMethod()> _
    Public Function GetDecideHistory(ByVal pFormNo As String, ByVal pFormSno As Integer) As String
        Dim RtnString As String = ""
        Dim SQL As String
        Try
            SQL = "Select DecideHistory From T_WaitHandle "
            SQL &= "Where FormNo  = '" & pFormNo & "' "
            SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
            SQL &= "Order by Unique_ID Desc "
            Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
            If dt_WaitHandle.Rows.Count > 0 Then
                RtnString = dt_WaitHandle.Rows(0).Item("DecideHistory").ToString
            End If
        Catch ex As Exception
            RtnString = ""
        End Try
        '
        Return RtnString
    End Function
    'GetDecideHistory-End
    '**************************************************************************************
    '** 建置待檢查資料(WriteWaitCheck)
    '**************************************************************************************
    'WriteWaitCheck-Start
    <WebMethod()> _
    Public Function WriteWaitCheck(ByVal pFormNo As String, ByVal pFormSno As Integer, ByVal pNowDateTime As String) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        Try
            SQL = "SELECT FormNo FROM Q_WaitCheck "
            SQL &= "Where Active  = '1' "
            SQL &= "  And FormNo  = '" & pFormNo & "' "
            SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
            SQL &= "ORDER BY FormNo, FormSno "
            Dim dt_WaitCheck As DataTable = uDataBase.GetDataTable(SQL)
            If dt_WaitCheck.Rows.Count <= 0 Then
                SQL = "Insert into Q_WaitCheck "
                SQL &= "( "
                SQL &= "Active, FormNo, FormSno, Description, CreateTime, ModifyTime "
                SQL &= ") "
                SQL &= "VALUES( "
                SQL &= "'1' ,"                              '控制狀態(0:完成,1:待處理)
                SQL &= "'" & pFormNo & "' ,"                '表單代號
                SQL &= "'" & CStr(pFormSno) & "' ,"         '表單編號
                SQL &= "'" & "" & "' ,"                     '處理狀態
                SQL &= "'" & pNowDateTime & "' ,"           '製作時間
                SQL &= "'" & pNowDateTime & "' "            '修改時間
                SQL &= ") "
                uDataBase.ExecuteNonQuery(SQL)
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'WriteWaitCheck-End
    '**************************************************************************************
    '** 是否可使用此表單(CanUseSheet)
    '** Add-Start 2011/12/2 對應:濫用資料修改委託書 / 對策:1人限1張(Active=1)
    '**************************************************************************************
    'CanUseSheet-Start
    <WebMethod()> _
    Public Function CanUseSheet(ByVal pFormNo As String, ByVal pUserID As String) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        Try
            If pFormNo = "800001" Then
                SQL = "SELECT FormSno FROM F_ModifyDataSheet "
            End If
            SQL &= "Where Sts = '0' "
            SQL &= "  And FormNo  = '" & pFormNo & "' "
            SQL &= "  And CreateUser = '" & pUserID & "' "
            Dim dt_Sheet As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Sheet.Rows.Count > 0 Then
                RtnCode = 1
                RtnCode = 0
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'CanUseSheet-End

    '2012/4/3 納期對應
    'Start
    '**************************************************************************************
    '** 取得流程負荷(0:無,1:有)(GetFlowLoading)  
    '**************************************************************************************
    'GetFlowLoading-Start
    <WebMethod()> _
    Public Function GetFlowLoading(ByVal pFormNo As String, ByVal pStep As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        Try
            SQL = "Select Loading From M_Flow "
            SQL &= "Where FormNo  = '" & pFormNo & "' "
            SQL &= "  And Step    = '" & CStr(pStep) & "' "
            SQL &= "  And Active  = '" & "1" & "' "
            Dim dt_Flow As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Flow.Rows.Count > 0 Then
                RtnCode = dt_Flow.Rows(0).Item("Loading")
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'GetFlowLoading-End

    '**************************************************************************************
    '** 取得NextAction-預定納期用(0:OK,1:NG1,2:NG2,3:SAVE)(GetNextAction)  
    '**************************************************************************************
    'GetNextAction-Start
    <WebMethod()> _
    Public Function GetNextAction(ByVal pFormNo As String, ByVal pStep As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        Try
            SQL = "Select NextAction From M_Flow "
            SQL &= "Where Active  = '1' "
            SQL &= "  And FormNo  = '" & pFormNo & "' "
            SQL &= "  And Step    = '" & CStr(pStep) & "' "
            Dim dt_Flow As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Flow.Rows.Count > 0 Then
                RtnCode = dt_Flow.Rows(0).Item("NextAction")
            End If
        Catch ex As Exception
            RtnCode = 0
        End Try
        '
        Return RtnCode
    End Function


    '**************************************************************************************
    '** 安排預定工程(ArrangFlowControl)  
    '**************************************************************************************
    'ArrangFlowControl-Start
    <WebMethod()> _
    Public Function ArrangFlowControl(ByVal pFormNo As String, _
                                      ByVal pFormSno As Integer, _
                                      ByVal pStep As Integer, _
                                      ByVal pLevel As String, _
                                      ByVal pApplyID As String, _
                                      ByVal pUserID As String, _
                                      ByVal pAllocateID As String, _
                                      ByVal pAction As Integer, _
                                      ByRef pNextStep As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        '
        Try
            Dim sFlow As New FlowService
            Dim sSchedule As New ScheduleService
            '
            Dim xAction As Integer = pAction
            Dim xUserID As String = pUserID
            '
            Dim xStep, xNextStep, xFlowType, xCount As Integer
            Dim xNextGate(10), xAgentGate(10) As String
            Dim xLastTime, xStartTime, xEndTime As DateTime
            Dim xNextGateCalendar As String
            Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")
            '
            xFlowType = 0
            xStep = pStep
            While xFlowType <> 1
                GetNextGate(pFormNo, xStep, xUserID, pApplyID, "", pAllocateID, 0, _
                                     xNextStep, xNextGate, xAgentGate, xCount, xFlowType, xAction)
                '
                xUserID = xNextGate(1)
                xStep = xNextStep
            End While
            '
            xNextGateCalendar = GetCalendarGroup(xUserID)
            '
            sSchedule.GetLastTimeCustom(xUserID, pFormNo, pFormSno, xStep, xFlowType, NowDateTime, 1, xLastTime, xCount)
            '
            sSchedule.GetLeadTime(pFormNo, xStep, pLevel, 0, xLastTime, xStartTime, xEndTime, xNextGateCalendar)
            '
            '建製下工程流程資料
            SQL = "Insert into T_WaitHandle "
            SQL &= "( "
            SQL &= "Active, FormNo, FormSno, Step, SeqNo, "
            SQL &= "Sts, FlowType, Important, ApplyTime, ReceiptTime, "
            SQL &= "ReadTimeLimit, BStartTime, BEndTime, AStartTime, ApplyID, "
            SQL &= "ApplyName, DecideID, DecideName, ReasonCode, Reason, "
            SQL &= "ReasonDesc, DecideDesc, CreateUser, CreateTime, EveryOne, "
            SQL &= "WorkID, StsDes, DecideHistory, AgentID, AgentName "
            SQL &= ") "
            SQL &= "VALUES( "
            SQL &= "'9' ,"                              '控制狀態(0:完成,1:待處理,9:預定工程)
            SQL &= "'" & pFormNo & "' ,"                '表單代號
            SQL &= "'" & CStr(pFormSno) & "' ,"         '表單代號
            SQL &= "'" & CStr(xStep) & "' ,"            '關卡代號
            SQL &= "'" & CStr(1) & "' ,"                '序號
            '1~5
            SQL &= "'" & "0" & "' ,"                    '處理狀態(0:未處理,1:核准/完工,2:駁回/NG,3:已閱讀)
            SQL &= "'" & "9" & "' ,"                    '流程種類(0:通知,1:核定,2:申請,3:多流程核定,9:預定工程)
            SQL &= "'" & "0" & "' ,"                    '重要性(0:一般,1:重要)
            SQL &= "'" & GetApplyTime(pFormNo, pFormSno, NowDateTime).ToString & "' ,"  '申請時間
            SQL &= "'" & NowDateTime & "' ,"                                            '收件時間
            '6~10
            SQL &= "'" & sSchedule.GetReadTimeLimit(NowDateTime, xNextGateCalendar).ToString & "' ,"    '閱讀期限
            SQL &= "'" & xStartTime.ToString("yyyy/MM/dd HH:mm:ss") & "' ,"                             '預定開始時間
            SQL &= "'" & xEndTime.ToString("yyyy/MM/dd HH:mm:ss") & "' ,"                               '預定完成時間
            SQL &= "'" & xStartTime.ToString("yyyy/MM/dd HH:mm:ss") & "' ,"                             '實際開始時間
            SQL &= "'" & pApplyID & "' ,"                           '申請者代號
            '11~15
            SQL &= "N'" & GetUserName(pApplyID).ToString & "' ,"    '申請者
            SQL &= "'" & xUserID & "' ,"                            '核定者代號
            SQL &= "N'" & GetUserName(xUserID).ToString & "' ,"     '核定者
            SQL &= "'" & "" & "' ,"                                 '延遲原因類別代碼
            SQL &= "'" & "" & "' ,"                                 '延遲原因類別名稱
            '16~20
            SQL &= "'" & "" & "' ,"                                 '延遲說明
            SQL &= "'" & "" & "' ,"                                 '核定說明
            SQL &= "'" & pUserID & "' ,"                            '建置者
            SQL &= "'" & NowDateTime & "' ,"                        '建置時間
            SQL &= "'" & GetSignType(pFormNo, pStep).ToString & "' ,"       '取得簽核種類   任一簽核 0:否, 1:是
            '21~25
            SQL &= "'" & GetWorkID(xUserID).ToString & "' ,"                '工作ID
            SQL &= "'" & "未處理" & "' ,"                           '狀態說明
            SQL &= "'" & "" & "' ,"                                 '核定者履歷
            SQL &= "'" & "" & "' ,"                                 '被代理者代號
            SQL &= "'" & "" & "' "                                  '被代理者
            SQL &= ") "
            uDataBase.ExecuteNonQuery(SQL)
            'Return NextStep
            pNextStep = xStep
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'ArrangFlowControl-End

    'End


End Class
