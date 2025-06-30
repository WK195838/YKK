Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient

<WebService(Namespace:="http://tempuri.org/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class FlowService
     Inherits System.Web.Services.WebService

    '**************************************************************************************
    '** 全域變數
    '**************************************************************************************
    Dim DBObj As New ForProject
    Dim uDataBase As Utility.DataBase = DBObj.GetDataBaseObj()
    Dim sCommon As New CommonService
    Dim sSchedule As New ScheduleService
    '**************************************************************************************
    '** 建置新流程(NewFlow)
    '**************************************************************************************
    'NewFlow-Start
    <WebMethod()> _
    Public Function NewFlow(ByVal pFormNo As String, _
                            ByVal pFormSno As Integer, _
                            ByVal pStep As Integer, _
                            ByVal pSeqNo As Integer, _
                            ByVal pDepo As String, _
                            ByVal pUser As String, ByVal pApplyID As String) As Integer
        Dim RtnCode As Integer = 0
        Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")
        Dim SQL As String
        Try
            '--Add Start By Joy 2016/8/20
            sCommon.WriteWaitCheck(pFormNo, pFormSno, NowDateTime)    '建置待檢查資料
            '
            SQL = "Select Unique_ID From T_WaitHandle "
            SQL &= "Where Active  = '0' "
            SQL &= "  And FormNo  = '" & pFormNo & "' "
            SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
            SQL &= "  And Step    = '" & CStr(pStep) & "' "
            SQL &= "  And SeqNo   = '" & CStr(pSeqNo) & "' "
            Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
            If dt_WaitHandle.Rows.Count <= 0 Then
                '
                '新增流程
                SQL = "Insert into T_WaitHandle "
                SQL &= "( "
                SQL &= "Active, FormNo, FormSno, Step, SeqNo, "
                SQL &= "Sts, FlowType, Important, ApplyTime, ReceiptTime, "
                SQL &= "ReadTimeLimit, FirstReadTime, LastReadTime, BStartTime, BEndTime, "
                SQL &= "AStartTime, AEndTime, CompletedTime, ApplyID, ApplyName, "
                SQL &= "DecideID, DecideName, ReasonCode, Reason, ReasonDesc, "
                SQL &= "DecideDesc, CreateUser, CreateTime, ModifyUser, ModifyTime, "
                SQL &= "EveryOne, WorkID, StsDes, DecideHistory, AgentID, "
                SQL &= "AgentName "
                SQL &= ") "
                SQL &= "VALUES( "
                SQL &= "'0' ,"                              '控制狀態(0:完成,1:待處理)
                SQL &= "'" & pFormNo & "' ,"                '表單代號
                SQL &= "'" & CStr(pFormSno) & "' ,"         '表單代號
                SQL &= "'" & CStr(pStep) & "' ,"            '關卡代號
                SQL &= "'" & CStr(pSeqNo) & "' ,"           '序號
                '1~5
                SQL &= "'" & "1" & "' ,"                    '處理狀態(0:未處理,1:核准/完工,2:駁回/NG,3:已閱讀)
                SQL &= "'" & "2" & "' ,"                    '流程種類(0:通知,1:核定,2:申請)
                SQL &= "'" & "0" & "' ,"                    '重要性(0:一般,1:重要)
                SQL &= "'" & NowDateTime & "' ,"            '申請時間
                SQL &= "'" & NowDateTime & "' ,"            '收件時間
                '6~10
                SQL &= "'" & NowDateTime & "' ,"            '閱讀期限
                SQL &= "'" & NowDateTime & "' ,"            '第一次閱讀時間
                SQL &= "'" & NowDateTime & "' ,"            '最後閱讀時間
                SQL &= "'" & NowDateTime & "' ,"            '預定開始時間
                SQL &= "'" & NowDateTime & "' ,"            '預定完成時間
                '11~15
                SQL &= "'" & NowDateTime & "' ,"            '實際開始時間
                SQL &= "'" & NowDateTime & "' ,"            '實際完成時間
                SQL &= "'" & NowDateTime & "' ,"            '完成時間
                SQL &= "'" & pApplyID & "' ,"               '申請者代號
                SQL &= "N'" & sCommon.GetUserName(pApplyID).ToString & "' ,"              '申請者
                '16~20
                SQL &= "'" & pApplyID & "' ,"               '核定者代號
                SQL &= "N'" & sCommon.GetUserName(pApplyID).ToString & "' ,"              '核定者
                SQL &= "'" & "" & "' ,"                     '延遲原因類別代碼
                SQL &= "'" & "" & "' ,"                     '延遲原因類別名稱
                SQL &= "'" & "" & "' ,"                     '延遲說明
                '21~25
                SQL &= "'" & "" & "' ,"                     '核定說明
                SQL &= "'" & pUser & "' ,"                  '建置者
                SQL &= "'" & NowDateTime & "' ,"            '建置時間
                SQL &= "'" & "" & "' ,"                     '修改者
                SQL &= "'" & NowDateTime & "' ,"            '修改時間
                '26~30
                SQL &= "'" & sCommon.GetSignType(pFormNo, pStep).ToString & "' ,"       '取得簽核種類   任一簽核 0:否, 1:是
                SQL &= "'" & sCommon.GetWorkID(pApplyID).ToString & "' ,"               '工作ID
                SQL &= "'" & "申請" & "' ,"                 '狀態說明
                SQL &= "'" & pApplyID & "' ,"               '核定者履歷
                SQL &= "'" & "" & "' ,"                     '被代理者代號
                SQL &= "'" & "" & "' "                      '被代理者
                SQL &= ") "
                uDataBase.ExecuteNonQuery(SQL)
                '
            End If
            '--Add End

        Catch ex As Exception
            RtnCode = 1
        End Try
        Return RtnCode
    End Function
    'NewFlow-End
    '**************************************************************************************
    '** 建置結束流程(EndFlow)
    '**************************************************************************************
    'EndFlow-Start
    <WebMethod()> _
    Public Function EndFlow(ByVal pFormNo As String, _
                            ByVal pFormSno As Integer, _
                            ByVal pStep As Integer, _
                            ByVal pSeqNo As Integer, _
                            ByVal pDepo As String, _
                            ByVal pUser As String, _
                            ByVal pApplyID As String) As Integer
        Dim RtnCode As Integer = 0
        Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")
        Dim SQL As String
        Try
            '--Add Start By Joy 2016/8/20
            sCommon.WriteWaitCheck(pFormNo, pFormSno, NowDateTime)    '建置待檢查資料
            '
            SQL = "Select Unique_ID From T_WaitHandle "
            SQL &= "Where Active  = '0' "
            SQL &= "  And FormNo  = '" & pFormNo & "' "
            SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
            SQL &= "  And Step    = '" & CStr(pStep) & "' "
            SQL &= "  And SeqNo   = '" & CStr(pSeqNo) & "' "
            Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
            If dt_WaitHandle.Rows.Count <= 0 Then
                '
                '新增流程
                SQL = "Insert into T_WaitHandle "
                SQL &= "( "
                SQL &= "Active, FormNo, FormSno, Step, SeqNo, "
                SQL &= "Sts, FlowType, Important, ApplyTime, ReceiptTime, "
                SQL &= "ReadTimeLimit, FirstReadTime, LastReadTime, BStartTime, BEndTime, "
                SQL &= "AStartTime, AEndTime, CompletedTime, ApplyID, ApplyName, "
                SQL &= "DecideID, DecideName, ReasonCode, Reason, ReasonDesc, "
                SQL &= "DecideDesc, CreateUser, CreateTime, ModifyUser, ModifyTime, "
                SQL &= "EveryOne, WorkID, StsDes, DecideHistory, AgentID, "
                SQL &= "AgentName "
                SQL &= ") "
                SQL &= "VALUES( "
                SQL &= "'0' ,"                              '控制狀態(0:完成,1:待處理)
                SQL &= "'" & pFormNo & "' ,"                '表單代號
                SQL &= "'" & CStr(pFormSno) & "' ,"         '表單代號
                SQL &= "'" & CStr(pStep) & "' ,"            '關卡代號
                SQL &= "'" & CStr(pSeqNo) & "' ,"           '序號
                '1~5
                SQL &= "'" & "1" & "' ,"                    '處理狀態(0:未處理,1:核准/完工,2:駁回/NG,3:已閱讀)
                SQL &= "'" & "0" & "' ,"                    '流程種類(0:通知,1:核定,2:申請)
                SQL &= "'" & "0" & "' ,"                    '重要性(0:一般,1:重要)
                SQL &= "'" & sCommon.GetApplyTime(pFormNo, pFormSno, NowDateTime).ToString & "' ,"  '申請時間
                SQL &= "'" & NowDateTime & "' ,"            '收件時間
                '6~10
                SQL &= "'" & "" & "' ,"                     '閱讀期限
                SQL &= "'" & "" & "' ,"                     '第一次閱讀時間
                SQL &= "'" & "" & "' ,"                     '最後閱讀時間
                SQL &= "'" & "" & "' ,"                     '預定開始時間
                SQL &= "'" & "" & "' ,"                     '預定完成時間
                '11~15
                SQL &= "'" & "" & "' ,"                     '實際開始時間
                SQL &= "'" & "" & "' ,"                     '實際完成時間
                SQL &= "'" & NowDateTime & "' ,"            '完成時間
                SQL &= "'" & pApplyID & "' ,"               '申請者代號
                SQL &= "N'" & sCommon.GetUserName(pApplyID).ToString & "' ,"              '申請者
                '16~20
                SQL &= "'" & "" & "' ,"                     '核定者代號
                SQL &= "'" & "" & "' ,"                     '核定者
                SQL &= "'" & "" & "' ,"                     '延遲原因類別代碼
                SQL &= "'" & "" & "' ,"                     '延遲原因類別名稱
                SQL &= "'" & "" & "' ,"                     '延遲說明
                '21~25
                SQL &= "'" & "" & "' ,"                     '核定說明
                SQL &= "'" & pUser & "' ,"                  '建置者
                SQL &= "'" & NowDateTime & "' ,"            '建置時間
                SQL &= "'" & "" & "' ,"                     '修改者
                SQL &= "'" & NowDateTime & "' ,"            '修改時間
                '26~30
                SQL &= "'" & sCommon.GetSignType(pFormNo, pStep).ToString & "' ,"       '取得簽核種類   任一簽核 0:否, 1:是
                SQL &= "'" & "" & "' ,"                     '工作ID
                SQL &= "'" & "結束" & "' ,"                 '申請
                SQL &= "'" & "" & "' ,"                     '核定者履歷
                SQL &= "'" & "" & "' ,"                     '被代理者代號
                SQL &= "'" & "" & "' "                      '被代理者
                SQL &= ") "
                uDataBase.ExecuteNonQuery(SQL)
                '
            End If
            '--Add End
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'EndFlow-End
    '**************************************************************************************
    '** 檢測流程資料(CheckFlow)
    '**************************************************************************************
    'CheckFlow-Start
    <WebMethod()> _
    Public Function CheckFlow(ByVal pFormNo As String, _
                              ByVal pFormSno As Integer, _
                              ByVal pStep As Integer, _
                              ByVal pDepo As String, _
                              ByVal pUser As String, _
                              ByRef pEnd As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")
        Dim SQL As String
        pEnd = 0
        Try
            sCommon.WriteWaitCheck(pFormNo, pFormSno, NowDateTime)                    '建置待檢查資料
            '
            '取得本關資訊
            SQL = "Select EveryOne, EveryOneStep From M_Flow "
            SQL &= "Where Active = '1' "
            SQL &= "  And Action = '0' "
            SQL &= "  And FormNo = '" & pFormNo & "' "
            SQL &= "  And Step   = '" & CStr(pStep) & "' "
            SQL &= "Order by Action "
            Dim dt_Flow As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Flow.Rows.Count > 0 Then
                If dt_Flow.Rows(0).Item("EveryOne").ToString = "1" Then             '任一簽核=是
                    If dt_Flow.Rows(0).Item("EveryOneStep").ToString = "0" Then     '同一工程
                        SQL = "Select Unique_ID From T_WaitHandle "
                        SQL &= "Where Active  = '1' "
                        SQL &= "  And FormNo  = '" & pFormNo & "' "
                        SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                        SQL &= "  And Step    = '" & CStr(pStep) & "' "
                        Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
                        If dt_WaitHandle.Rows.Count > 0 Then                        '有其他同一工程待簽更新為已核定
                            SQL = "Update T_WaitHandle Set "
                            SQL &= " Active = '" & "0" & "',"
                            SQL &= " Sts = '" & "1" & "',"
                            SQL &= " AEndTime = '" & NowDateTime & "',"
                            SQL &= " CompletedTime = '" & NowDateTime & "',"
                            SQL &= " ModifyUser = '" & pUser & "',"
                            SQL &= " ModifyTime = '" & NowDateTime & "',"
                            SQL &= " StsDes = '" & "完成" & "' "
                            SQL &= "Where Active  = '1' "
                            SQL &= "  And FormNo  = '" & pFormNo & "' "
                            SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                            SQL &= "  And Step    = '" & CStr(pStep) & "' "
                            uDataBase.ExecuteNonQuery(SQL)
                        End If
                        pEnd = 1
                    Else                                                            '不同工程
                        SQL = "Select Unique_ID From T_WaitHandle "           '有同一工程待簽更新為已核定
                        SQL &= "Where Active  = '1' "
                        SQL &= "  And FormNo  = '" & pFormNo & "' "
                        SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                        SQL &= "  And Step    = '" & CStr(pStep) & "' "
                        Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
                        If dt_WaitHandle.Rows.Count > 0 Then
                            SQL = "Update T_WaitHandle Set "
                            SQL &= " Active = '" & "0" & "',"
                            SQL &= " Sts = '" & "1" & "',"
                            SQL &= " AEndTime = '" & NowDateTime & "',"
                            SQL &= " CompletedTime = '" & NowDateTime & "',"
                            SQL &= " ModifyUser = '" & pUser & "',"
                            SQL &= " ModifyTime = '" & NowDateTime & "',"
                            SQL &= " StsDes = '" & "完成" & "' "
                            SQL &= "Where Active  = '1' "
                            SQL &= "  And FormNo  = '" & pFormNo & "' "
                            SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                            SQL &= "  And Step    = '" & CStr(pStep) & "' "
                            uDataBase.ExecuteNonQuery(SQL)
                        End If
                        '有其他不同工程待簽更新為已核定
                        SQL = "Select Unique_ID From T_WaitHandle "
                        SQL &= "Where Active  = '1' "
                        SQL &= "  And FormNo  = '" & pFormNo & "' "
                        SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                        SQL &= "  And Step    = '" & dt_Flow.Rows(0).Item("EveryOneStep").ToString & "' "
                        dt_WaitHandle = uDataBase.GetDataTable(SQL)
                        If dt_WaitHandle.Rows.Count > 0 Then
                            SQL = "Update T_WaitHandle Set "
                            SQL &= " Active = '" & "0" & "',"
                            SQL &= " Sts = '" & "1" & "',"
                            SQL &= " AEndTime = '" & NowDateTime & "',"
                            SQL &= " CompletedTime = '" & NowDateTime & "',"
                            SQL &= " ModifyUser = '" & pUser & "',"
                            SQL &= " ModifyTime = '" & NowDateTime & "',"
                            SQL &= " StsDes = '" & "完成" & "' "
                            SQL &= "Where Active  = '1' "
                            SQL &= "  And FormNo  = '" & pFormNo & "' "
                            SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                            SQL &= "  And Step    = '" & dt_Flow.Rows(0).Item("EveryOneStep").ToString & "' "
                            uDataBase.ExecuteNonQuery(SQL)
                        End If
                        pEnd = 1
                    End If
                Else                                                            '任一簽核=否
                    If dt_Flow.Rows(0).Item("EveryOneStep").ToString = "0" Then     '同一工程
                        SQL = "Select Unique_ID From T_WaitHandle "
                        SQL &= "Where Active  = '1' "
                        SQL &= "  And FormNo  = '" & pFormNo & "' "
                        SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                        SQL &= "  And Step    = '" & CStr(pStep) & "' "
                        Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
                        If dt_WaitHandle.Rows.Count <= 0 Then                        '有其他同一工程待簽更新為已核定
                            pEnd = 1
                        End If
                    Else                                                            '不同工程
                        '有其他不同工程待簽更新為已核定
                        SQL = "Select Unique_ID From T_WaitHandle "
                        SQL &= "Where Active  = '1' "
                        SQL &= "  And FormNo  = '" & pFormNo & "' "
                        SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                        SQL &= "  And Step    = '" & dt_Flow.Rows(0).Item("EveryOneStep").ToString & "' "
                        Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
                        If dt_WaitHandle.Rows.Count <= 0 Then
                            pEnd = 1
                        End If

                        'If pFormNo = "002002" Then
                        '    'SCD專用
                        '    SQL = "Select Unique_ID From T_WaitHandle "
                        '    SQL &= "Where Active  = '1' "
                        '    SQL &= "Where FormNo  = '" & pFormNo & "' "
                        '    SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                        '    SQL &= "  And Step    = '" & dt_Flow.Rows(0).Item("EveryOneStep").ToString & "' "
                        '    Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
                        '    If dt_WaitHandle.Rows.Count > 0 Then
                        '        pEnd = 1
                        '    End If
                        'Else
                        '    '標準
                        '    SQL = "Select Unique_ID From T_WaitHandle "
                        '    SQL &= "Where Active  = '1' "
                        '    SQL &= "  And FormNo  = '" & pFormNo & "' "
                        '    SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                        '    SQL &= "  And Step    = '" & dt_Flow.Rows(0).Item("EveryOneStep").ToString & "' "
                        '    Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
                        '    If dt_WaitHandle.Rows.Count > 0 Then
                        '        pEnd = 1
                        '    End If
                        'End If
                    End If
                End If
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'CheckFlow-End
    '**************************************************************************************
    '** 更新流程(UpdateFlow)
    '**************************************************************************************
    'UpdateFlow-Start
    <WebMethod()> _
    Public Function UpdateFlow(ByVal pFormNo As String, _
                               ByVal pFormSno As Integer, _
                               ByVal pStep As Integer, _
                               ByVal pSeqNo As Integer, _
                               ByVal pDepo As String, _
                               ByVal pUser As String) As Integer
        Dim RtnCode As Integer = 0
        Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")
        Dim SQL As String
        Dim xHistory As String = ""
        Dim xEnd As Integer = 0
        Try
            sCommon.WriteWaitCheck(pFormNo, pFormSno, NowDateTime)    '建置待檢查資料
            '
            '取得流程資料
            SQL = "Select Unique_ID, FirstReadTime, DecideHistory, FlowType From T_WaitHandle "
            SQL &= "Where Active  = '1' "
            SQL &= "  And FormNo  = '" & pFormNo & "' "
            SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
            SQL &= "  And Step    = '" & CStr(pStep) & "' "
            SQL &= "  And SeqNo   = '" & CStr(pSeqNo) & "' "
            Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
            If dt_WaitHandle.Rows.Count > 0 Then
                '是否首次閱讀
                If IsDBNull(dt_WaitHandle.Rows(0).Item("FirstReadTime")) Or dt_WaitHandle.Rows(0).Item("FirstReadTime").ToString < "1901/01/01 00:00:00" Then
                    SQL = "Update T_WaitHandle Set "                        '更新已閱讀,首次/最終閱讀時間
                    SQL &= " Sts = '" & "4" & "',"
                    SQL &= " FirstReadTime = '" & NowDateTime & "',"
                    SQL &= " LastReadTime = '" & NowDateTime & "',"
                    SQL &= " ModifyUser = '" & pUser & "',"
                    SQL &= " ModifyTime = '" & NowDateTime & "',"
                    SQL &= " StsDes = '" & "已閱讀" & "' "
                    SQL &= "Where Unique_ID = '" & dt_WaitHandle.Rows(0).Item("Unique_ID").ToString & "' "
                    uDataBase.ExecuteNonQuery(SQL)
                Else
                    SQL = "Update T_WaitHandle Set "                        '更新最終閱讀時間
                    SQL &= " Sts = '" & "4" & "',"
                    SQL &= " LastReadTime = '" & NowDateTime & "',"
                    SQL &= " ModifyUser = '" & pUser & "',"
                    SQL &= " ModifyTime = '" & NowDateTime & "',"
                    SQL &= " StsDes = '" & "已閱讀" & "' "
                    SQL &= "Where Unique_ID = '" & dt_WaitHandle.Rows(0).Item("Unique_ID").ToString & "' "
                    uDataBase.ExecuteNonQuery(SQL)
                End If
                '是否是通知
                If dt_WaitHandle.Rows(0).Item("FlowType").ToString = "0" Then
                    SQL = "Update T_WaitHandle Set "                        '更新結束流程
                    SQL &= " Active = '" & "0" & "',"
                    SQL &= " Sts = '" & "1" & "',"
                    SQL &= " AEndTime = '" & NowDateTime & "',"
                    SQL &= " CompletedTime = '" & NowDateTime & "',"
                    SQL &= " ModifyUser = '" & pUser & "',"
                    SQL &= " ModifyTime = '" & NowDateTime & "',"
                    SQL &= " StsDes = '" & "完成" & "' "
                    SQL &= "Where Unique_ID = '" & dt_WaitHandle.Rows(0).Item("Unique_ID").ToString & "' "
                    uDataBase.ExecuteNonQuery(SQL)
                    '更新簽核履歷
                    xHistory = dt_WaitHandle.Rows(0).Item("DecideHistory").ToString
                    If InStr(1, xHistory, pUser) = 0 Then xHistory = xHistory + "," + pUser
                    SQL = "Update T_WaitHandle Set "
                    SQL &= " DecideHistory = '" & xHistory & "' "
                    SQL &= "Where FormNo  = '" & pFormNo & "' "
                    SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                    uDataBase.ExecuteNonQuery(SQL)
                    '檢測流程資料(CheckFlow)
                    CheckFlow(pFormNo, pFormSno, pStep, pDepo, pUser, xEnd)
                End If
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'UpdateFlow-End
    '**************************************************************************************
    '** 建置下工程流程(NextFlow)
    '**************************************************************************************
    'NextFlow-Start
    <WebMethod()> _
    Public Function NextFlow(ByVal pFormNo As String, _
                             ByVal pFormSno As Integer, _
                             ByVal pStep As Integer, _
                             ByVal pSeqNo As Integer, _
                             ByVal pDepo As String, _
                             ByVal pUser As String, _
                             ByVal pNextUser As String, _
                             ByVal pAgentID As String, _
                             ByVal pApplyID As String, _
                             ByVal pStartTime As DateTime, _
                             ByVal pEndTime As DateTime, _
                             ByVal pImportant As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")
        Dim SQL As String
        Try
            Dim xHistory As String = sCommon.GetDecideHistory(pFormNo, pFormSno).ToString   '核定者履歷
            If InStr(1, xHistory, pUser) = 0 Then xHistory = xHistory + "," + pUser
            '
            '--Add Start By Joy 2016/8/20
            sCommon.WriteWaitCheck(pFormNo, pFormSno, NowDateTime)    '建置待檢查資料
            '
            SQL = "Select Unique_ID From T_WaitHandle "
            SQL &= "Where Active  = '1' "
            SQL &= "  And FormNo  = '" & pFormNo & "' "
            SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
            SQL &= "  And Step    = '" & CStr(pStep) & "' "
            SQL &= "  And SeqNo   = '" & CStr(pSeqNo) & "' "
            Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
            If dt_WaitHandle.Rows.Count <= 0 Then
                '
                SQL = "Update T_WaitHandle Set "
                SQL &= " DecideHistory = '" & xHistory & "' "
                SQL &= "Where FormNo  = '" & pFormNo & "' "
                SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                uDataBase.ExecuteNonQuery(SQL)
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
                SQL &= "'1' ,"                              '控制狀態(0:完成,1:待處理)
                SQL &= "'" & pFormNo & "' ,"                '表單代號
                SQL &= "'" & CStr(pFormSno) & "' ,"         '表單代號
                SQL &= "'" & CStr(pStep) & "' ,"            '關卡代號
                SQL &= "'" & CStr(pSeqNo) & "' ,"           '序號
                '1~5
                SQL &= "'" & "0" & "' ,"                    '處理狀態(0:未處理,1:核准/完工,2:駁回/NG,3:已閱讀)
                SQL &= "'" & sCommon.GetFlowType(pFormNo, pStep).ToString & "' ,"   '流程種類(0:通知,1:核定,2:申請)
                SQL &= "'" & CStr(pImportant) & "' ,"       '重要性(0:一般,1:重要)
                SQL &= "'" & sCommon.GetApplyTime(pFormNo, pFormSno, NowDateTime).ToString & "' ,"  '申請時間
                SQL &= "'" & NowDateTime & "' ,"            '收件時間
                '6~10
                SQL &= "'" & sSchedule.GetReadTimeLimit(NowDateTime, pDepo).ToString & "' ,"    '閱讀期限
                SQL &= "'" & pStartTime.ToString("yyyy/MM/dd HH:mm:ss") & "' ,"             '預定開始時間
                SQL &= "'" & pEndTime.ToString("yyyy/MM/dd HH:mm:ss") & "' ,"               '預定完成時間
                SQL &= "'" & pStartTime.ToString("yyyy/MM/dd HH:mm:ss") & "' ,"             '實際開始時間
                SQL &= "'" & pApplyID & "' ,"               '申請者代號
                '11~15
                SQL &= "N'" & sCommon.GetUserName(pApplyID).ToString & "' ,"     '申請者
                SQL &= "'" & pNextUser & "' ,"              '核定者代號
                SQL &= "N'" & sCommon.GetUserName(pNextUser).ToString & "' ,"    '核定者
                SQL &= "'" & "" & "' ,"                     '延遲原因類別代碼
                SQL &= "'" & "" & "' ,"                     '延遲原因類別名稱
                '16~20
                SQL &= "'" & "" & "' ,"                     '延遲說明
                SQL &= "'" & "" & "' ,"                     '核定說明
                SQL &= "'" & pUser & "' ,"                  '建置者
                SQL &= "'" & NowDateTime & "' ,"            '建置時間
                SQL &= "'" & sCommon.GetSignType(pFormNo, pStep).ToString & "' ,"       '取得簽核種類   任一簽核 0:否, 1:是
                '21~25
                SQL &= "'" & sCommon.GetWorkID(pNextUser).ToString & "' ,"              '工作ID
                SQL &= "'" & "未處理" & "' ,"               '狀態說明
                SQL &= "'" & xHistory & "' ,"               '核定者履歷
                SQL &= "'" & pAgentID & "' ,"               '被代理者代號
                SQL &= "N'" & sCommon.GetUserName(pAgentID).ToString & "' "              '被代理者
                SQL &= ") "
                uDataBase.ExecuteNonQuery(SQL)
                '
            End If
            '--Add End
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'NextFlow-End
    '
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    'Add-Start by Joy  2012/7/31 新納期對應
    '
    '**************************************************************************************
    '** 刪除預定工程(Active=9 and FlowType=9)(DeleteArrangFlow)
    '**************************************************************************************
    'DeleteArrangFlow-Start
    <WebMethod()> _
    Public Function DeleteArrangFlow(ByVal pFormNo As String, _
                                     ByVal pFormSno As Integer, _
                                     ByVal pNextStep As Integer, _
                                     ByVal pSeqNo As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")
        Dim SQL As String
        Try
            sCommon.WriteWaitCheck(pFormNo, pFormSno, NowDateTime)    '建置待檢查資料

            SQL = "Select * From T_WaitHandle "
            SQL &= "Where FormNo  = '" & pFormNo & "' "
            SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
            SQL &= "  And Step = '" & CStr(pNextStep) & "' "
            SQL &= "  And SeqNo = '" & CStr(pSeqNo) & "' "
            SQL &= "  And Active = '1' "
            Dim dt_WaitHandle = uDataBase.GetDataTable(SQL)
            If dt_WaitHandle.Rows.Count > 0 Then
                SQL = "Delete From T_WaitHandle "
                SQL &= "Where FormNo  = '" & pFormNo & "' "
                SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                SQL &= "  And Step = '" & CStr(pNextStep) & "' "
                SQL &= "  And SeqNo = '" & CStr(pSeqNo) & "' "
                SQL &= "  And Active = '9' "
                uDataBase.ExecuteNonQuery(SQL)
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'DeleteArrangFlow-End
    'Add-End
End Class
