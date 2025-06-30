Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient

<WebService(Namespace:="http://tempuri.org/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class ScheduleService
    Inherits System.Web.Services.WebService

    '**************************************************************************************
    '** 全域變數
    '**************************************************************************************
    Dim DBObj As New ForProject
    Dim uDataBase As Utility.DataBase = DBObj.GetDataBaseObj()
    Dim sCommon As New CommonService
    '**************************************************************************************
    '** 取得最後預定完成及待處理件數(GetLastTime)
    '   2014/10/9 DTMW對應版
    '       負荷考量下增加接受 [同時處理多件數]
    '       原:只限單件  ---> 新:支援單日同時處理多件數
    '**************************************************************************************
    'GetLastTime-Start
    '
    <WebMethod()> _
    Public Function GetLastTime(ByVal pUser As String, _
                                ByVal pFormNo As String, _
                                ByVal pStep As Integer, _
                                ByVal pFlowType As Integer, _
                                ByVal pNowDateTime As DateTime, _
                                ByRef pLastTime As DateTime, _
                                ByRef pCount1 As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        Try
            '設定初始值
            pLastTime = pNowDateTime
            pCount1 = 0
            '取得WorkID
            If sCommon.GetWorkID(pUser).ToString <> "" Then
                '取得待處理件數
                SQL = "Select Count(*) As RecCount From T_WaitHandle "
                SQL &= "Where WorkID = '" & sCommon.GetWorkID(pUser).ToString & "' "
                SQL &= "  And SeqNo  = '1' "
                SQL &= "  And Active = '1' "
                Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
                If dt_WaitHandle.Rows.Count > 0 Then
                    pCount1 = dt_WaitHandle.Rows(0).Item("RecCount")
                End If
                '取得最後預定完成
                If pFlowType <> 0 Then
                    SQL = "Select Loading, Record From M_Flow "
                    SQL &= "Where FormNo = '" & pFormNo & "' "
                    SQL &= "  And Step   = '" & CStr(pStep) & "' "
                    SQL &= "  And Active = '1' "
                    Dim dt_Flow As DataTable = uDataBase.GetDataTable(SQL)
                    If dt_Flow.Rows.Count > 0 Then

                     

                        '負荷考量
                        If dt_Flow.Rows(0).Item("Loading").ToString = "1" Then
                            '單件處理
                            If dt_Flow.Rows(0).Item("Record").ToString = "1" Then
                                '取得最後待處理工程預定完成時間
                                SQL = "Select BEndTime From T_WaitHandle "
                                SQL &= "Where WorkID   = '" & sCommon.GetWorkID(pUser).ToString & "' "
                                SQL &= "  And FlowType = '1' "
                                SQL &= " ) "
                                SQL &= "Order by BEndTime Desc "
                                dt_WaitHandle = uDataBase.GetDataTable(SQL)
                                If dt_WaitHandle.Rows.Count > 0 Then
                                    pLastTime = CDate(dt_WaitHandle.Rows(0).Item("BEndTime").ToString)
                                End If
                            Else
                                '20141223修改 JESSICA

                                '單日同時處理多件

                                Dim xStartDate As String = ""
                                Dim xStartDateTime As DateTime = CDate(pNowDateTime.ToString("yyyy/MM/dd") + " " + "08:00:00")
                                '1.AM/PM取得下一工作日
                                If CInt(pNowDateTime.ToString("HH")) >= 12 Then
                                    xStartDateTime = CDate(pNowDateTime.AddDays(1).ToString("yyyy/MM/dd") + " " + "08:00:00")
                                End If
                                '2.預定開始日的待處件數 >= 同時處理件數
                                ' For i As Integer = 0 To 99
                                xStartDate = CDate(GetStartTime(xStartDateTime.ToString("yyyy/MM/dd HH:mm:ss"), sCommon.GetCalendarGroup(pUser))).ToString("yyyy/MM/dd")                                '


                                ' 取那些表單件數需算在一起
                                ' SQL = " select  data  from  m_referp where cat = '017' and Dkey = 'ALL' "
                              
                                '2018/02/27 修改UA&KIPPLING CAPA Jessica

                                ' 取那些表單件數需算在一起 
                                If pFormNo = "005011" Then
                                    If CStr(pStep) = "45" Then  'SLD
                                        SQL = " select  data  from  m_referp where cat = '017' and Dkey = 'SLD' "
                                    ElseIf CStr(pStep) = "40" Or CStr(pStep) = "43" Then
                                        SQL = " select  data  from  m_referp where cat = '017' and Dkey = 'VF' "
                                    Else
                                        SQL = " select  data  from  m_referp where cat = '017' and Dkey = 'DYE' "
                                    End If

                                Else
                                    SQL = " select  data  from  m_referp where cat = '017' and Dkey = 'DYE' "
                                End If
                                Dim dt_Mreferp As DataTable = uDataBase.GetDataTable(SQL)
                                dt_Mreferp = uDataBase.GetDataTable(SQL)

                                Dim DataString As String = ""
                                If dt_Mreferp.Rows.Count > 0 Then
                                    DataString = dt_Mreferp.Rows(0).Item("data")
                                End If

                                '4.增加單日同時處理多數的字串找出最近時間的一筆資料的開始日期及時間
                                SQL = "Select max(BStartTime)BStartTime,max(convert(char(10),BStartTime,111))BStartDate From T_WaitHandle "
                                SQL &= "Where WorkID   = '" & sCommon.GetWorkID(pUser).ToString & "' "
                                SQL &= "  And FlowType = '1' "
                                SQL &= "  And Active   = '1' "
                                SQL &= "  And FormNo+'-'+ LTrim(RTrim(Str(Step))) IN (" + DataString + ")"
                                dt_WaitHandle = uDataBase.GetDataTable(SQL)
                                If dt_WaitHandle.Rows.Count > 0 Then
                                    If dt_WaitHandle.Rows(0).Item("BStartDate").ToString <> "" Then
                                        If xStartDate <= dt_WaitHandle.Rows(0).Item("BStartDate").ToString Then '如果xStartDateTime比最大的日期小才修改
                                            xStartDate = dt_WaitHandle.Rows(0).Item("BStartDate")
                                            xStartDateTime = CDate(dt_WaitHandle.Rows(0).Item("BStartDate") + " " + "08:00:00")
                                        End If

                                    End If
                                End If




                                '5.負荷調整程式
                                SQL = "Select Count(*) As RecCount From T_WaitHandle "
                                SQL &= "Where WorkID = '" & sCommon.GetWorkID(pUser).ToString & "' "
                                SQL &= "  And Convert(varchar(20),BStartTime,111) = '" & xStartDate & "' "
                                SQL &= "  And SeqNo  = '1' "
                                SQL &= "  And Active = '1' "
                                SQL &= "  And FormNo+'-'+ LTrim(RTrim(Str(Step))) IN (" + DataString + ")"

                                Dim dt_WaitRecord As DataTable = uDataBase.GetDataTable(SQL)
                                If dt_WaitRecord.Rows.Count > 0 Then
                                    If dt_WaitRecord.Rows(0).Item("RecCount") < dt_Flow.Rows(0).Item("Record") Then
                                        pLastTime = CDate(xStartDate + " " + "07:55:00")
                                        '改變 pNowDateTime讓條件不成立
                                        pNowDateTime = CDate(xStartDate + " " + "07:00:00")
                                    Else
                                        pLastTime = CDate(xStartDateTime.AddDays(1).ToString("yyyy/MM/dd") + " " + "07:55:00")
                                    End If
                                Else
                                    '如果都沒有資料就以當天的日期
                                    pLastTime = CDate(pNowDateTime.ToString("yyyy/MM/dd") + " " + "07:55:00")
                                    '改變 pNowDateTime讓條件不成立
                                    pNowDateTime = CDate(pNowDateTime.ToString("yyyy/MM/dd") + " " + "07:00:00")
                                End If


                                '20141223修改 JESSICA
                                '
                                '  If Finished = True Then
                                '    Exit For
                                'Else
                                'xStartDateTime = CDate(xStartDate + " " + "08:00:00")
                                'End If
                                ' Next
                            End If
                        End If
                    End If
                End If
            End If
            If pLastTime <= pNowDateTime Then pLastTime = pNowDateTime
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    '--------------------------------------------------------
    '---2014/10/9 DTMW對應前原始程式                              -----
    '--------------------------------------------------------
    'Public Function GetLastTime(ByVal pUser As String, _
    '                            ByVal pFormNo As String, _
    '                            ByVal pStep As Integer, _
    '                            ByVal pFlowType As Integer, _
    '                            ByVal pNowDateTime As DateTime, _
    '                            ByRef pLastTime As DateTime, _
    '                            ByRef pCount1 As Integer) As Integer
    '    Dim RtnCode As Integer = 0
    '    Dim SQL As String
    '    Try
    '        '設定初始值
    '        pLastTime = pNowDateTime
    '        pCount1 = 0
    '        '取得WorkID
    '        If sCommon.GetWorkID(pUser).ToString <> "" Then
    '            '取得待處理件數
    '            SQL = "Select Count(*) As RecCount From T_WaitHandle "
    '            SQL &= "Where WorkID = '" & sCommon.GetWorkID(pUser).ToString & "' "
    '            SQL &= "  And SeqNo  = '1' "
    '            SQL &= "  And Active = '1' "
    '            Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
    '            If dt_WaitHandle.Rows.Count > 0 Then
    '                pCount1 = dt_WaitHandle.Rows(0).Item("RecCount")
    '            End If
    '            '取得最後預定完成
    '            If pFlowType <> 0 Then
    '                SQL = "Select Loading, Record From M_Flow "
    '                SQL &= "Where FormNo = '" & pFormNo & "' "
    '                SQL &= "  And Step   = '" & CStr(pStep) & "' "
    '                SQL &= "  And Active = '1' "
    '                Dim dt_Flow As DataTable = uDataBase.GetDataTable(SQL)
    '                If dt_Flow.Rows.Count > 0 Then
    '                    If dt_Flow.Rows(0).Item("Loading").ToString = "1" Then          '負荷考量
    '                        '**不對應多件處理
    '                        'If dt_Flow.Rows(0).Item("Record").ToString = "1" Then      '單件處理,
    '                        '取得最後待處理工程預定完成時間
    '                        SQL = "Select BEndTime From T_WaitHandle "
    '                        SQL &= "Where WorkID   = '" & sCommon.GetWorkID(pUser).ToString & "' "
    '                        SQL &= "  And FlowType = '1' "
    '                        SQL &= "  And Active   = '1' "
    '                        SQL &= "Order by BEndTime Desc "
    '                        dt_WaitHandle = uDataBase.GetDataTable(SQL)
    '                        If dt_WaitHandle.Rows.Count > 0 Then
    '                            pLastTime = CDate(dt_WaitHandle.Rows(0).Item("BEndTime").ToString)
    '                        End If
    '                        'Else                                                        
    '                        '多件同時處理-->不對應
    '                        'End If
    '                    End If
    '                End If
    '            End If
    '        End If
    '        If pLastTime <= pNowDateTime Then pLastTime = pNowDateTime
    '    Catch ex As Exception
    '        RtnCode = 1
    '    End Try
    '    '
    '    Return RtnCode
    'End Function
    'GetLastTime-End
    '**************************************************************************************
    '** 取得開發所花時間(GetDevelopTime)
    '**************************************************************************************
    'GetDevelopTime-Start
    <WebMethod()> _
    Public Function GetDevelopTime(ByVal pStartTime As String, ByVal pEndTime As String, ByRef pDevelopTime As Integer, ByVal pDepo As String) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        Dim xStartSeqNo As Integer = 0
        Dim xEndSeqNo As Integer = 0
        Try
            '開始日分解->日期及分
            Dim xNowDate As String = CDate(pStartTime).ToString("yyyy/MM/dd")
            Dim xNowTime As Integer = CInt(CDate(pStartTime).ToString("HH")) * 60 + CInt(CDate(pStartTime).ToString("mm"))
            '開始日序號
            SQL = "Select SeqNo From M_Calendar "
            SQL &= "Where YMD  = '" & xNowDate & "' "
            SQL &= "  And Hour >= '" & CStr(xNowTime) & "' "
            SQL &= "  And Active = '1' "
            SQL &= "  And Depo = '" & pDepo & "' "
            SQL &= "Order By YMD, Hour "
            Dim dt_Calendar As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Calendar.Rows.Count > 0 Then
                xStartSeqNo = dt_Calendar.Rows(0).Item("SeqNo")
            Else
                SQL = "Select SeqNo From M_Calendar "
                SQL &= "Where YMD  > '" & xNowDate & "' "
                SQL &= "  And Active = '1' "
                SQL &= "  And Depo = '" & pDepo & "' "
                SQL &= "Order By YMD, Hour "
                dt_Calendar = uDataBase.GetDataTable(SQL)
                If dt_Calendar.Rows.Count > 0 Then
                    xStartSeqNo = dt_Calendar.Rows(0).Item("SeqNo")
                End If
            End If
            '開始日分解->日期及分
            xNowDate = CDate(pEndTime).ToString("yyyy/MM/dd")
            xNowTime = CInt(CDate(pEndTime).ToString("HH")) * 60 + CInt(CDate(pEndTime).ToString("mm"))
            '開始日序號
            SQL = "Select SeqNo From M_Calendar "
            SQL &= "Where YMD  = '" & xNowDate & "' "
            SQL &= "  And Hour >= '" & CStr(xNowTime) & "' "
            SQL &= "  And Active = '1' "
            SQL &= "  And Depo = '" & pDepo & "' "
            SQL &= "Order By YMD, Hour "
            dt_Calendar = uDataBase.GetDataTable(SQL)
            If dt_Calendar.Rows.Count > 0 Then
                xEndSeqNo = dt_Calendar.Rows(0).Item("SeqNo")
            Else
                SQL = "Select SeqNo From M_Calendar "
                SQL &= "Where YMD  > '" & xNowDate & "' "
                SQL &= "  And Active = '1' "
                SQL &= "  And Depo = '" & pDepo & "' "
                SQL &= "Order By YMD, Hour "
                dt_Calendar = uDataBase.GetDataTable(SQL)
                If dt_Calendar.Rows.Count > 0 Then
                    xEndSeqNo = dt_Calendar.Rows(0).Item("SeqNo")
                End If
            End If
            pDevelopTime = (xEndSeqNo - xStartSeqNo) * 10
            If pDevelopTime < 0 Then pDevelopTime = 0
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'GetDevelopTime-End
    '**************************************************************************************
    '** 日程調整(AdjustSchedule)
    '**************************************************************************************
    'AdjustSchedule-Start
    <WebMethod()> _
    Public Function AdjustSchedule(ByVal pUser As String, _
                                   ByVal pFormNo As String, _
                                   ByVal pFormSno As Integer, _
                                   ByVal pStep As Integer, _
                                   ByVal pSeqNo As Integer, _
                                   ByVal pLastTime As String, _
                                   ByVal pLevel As String, _
                                   ByVal pDepo As String) As Integer
        Dim RtnCode As Integer = 0
        '****** 停止使用-Start ******
        'Dim xUnique_ID, xQCLT, i As Integer
        'Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")
        'Dim xNowDateTime, xWorkID, xLevel, xTableName, xBaseEndTime, xStartTime, xEndTime As String
        'Dim SQL As String
        'Try
        '    If pFormNo >= "000001" And pFormNo <= "001000" Then
        '        SQL = "Select Unique_ID, DecideID, Loading, AEndTime From V_WaitHandle_01 "
        '        SQL &= "Where FormNo  = '" & pFormNo & "' "
        '        SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
        '        SQL &= "  And Step = '" & CStr(pStep) & "' "
        '        SQL &= "  And SeqNo = '" & CStr(pSeqNo) & "' "
        '        SQL &= "Order by DecideID "
        '        Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
        '        If dt_WaitHandle.Rows.Count > 0 Then
        '            xUnique_ID = dt_WaitHandle.Rows(0).Item("Unique_ID")
        '            xEndTime = dt_WaitHandle.Rows(0).Item("AEndTime").ToString
        '            xWorkID = sCommon.GetWorkID(dt_WaitHandle.Rows(0).Item("DecideID")).ToString
        '            If xWorkID <> "999999" And xWorkID <> "" Then
        '                '採用實際完成日基準
        '                SQL = "Select AEndTime From T_WaitHandle "      '取得同表單上一工程實際完成日
        '                SQL &= "Where Active  = '0' "
        '                SQL &= "  And FlowType <> '0' "
        '                SQL &= "  And FormNo  = '" & pFormNo & "' "
        '                SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
        '                SQL &= "  And Step < '" & CStr(pStep) & "' "
        '                SQL &= "Order by Step Desc, SeqNo Desc "
        '                dt_WaitHandle = uDataBase.GetDataTable(SQL)
        '                If dt_WaitHandle.Rows.Count > 0 Then
        '                    xBaseEndTime = dt_WaitHandle.Rows(0).Item("AEndTime").ToString
        '                End If
        '                SQL = "Select AEndTime From T_WaitHandle "      '取得工作ID上一工程實際完成日
        '                SQL &= "Where Active  = '0' "
        '                SQL &= "  And WorkID  = '" & xWorkID & "' "
        '                SQL &= "  And FlowType <> '0' "
        '                SQL &= "  And AEndTime < '" & xEndTime & "' "
        '                SQL &= "Order by AEndTime Desc "
        '                dt_WaitHandle = uDataBase.GetDataTable(SQL)
        '                If dt_WaitHandle.Rows.Count > 0 Then
        '                    If dt_WaitHandle.Rows(0).Item("AEndTime").ToString > xBaseEndTime Then
        '                        xBaseEndTime = dt_WaitHandle.Rows(0).Item("AEndTime").ToString
        '                    End If
        '                    SQL = "Update T_WaitHandle Set "            '更新實際開始日
        '                    SQL &= " AStartTime = '" & GetStartTime(xBaseEndTime, pDepo) & "', "
        '                    SQL &= " ModifyUser = '" & "ASchedule" & "', "
        '                    SQL &= " ModifyTime = '" & NowDateTime & "' "
        '                    SQL &= "Where Unique_ID = '" & xUnique_ID & "' "
        '                    uDataBase.ExecuteNonQuery(SQL)
        '                End If
        '                '納期調整(實際開始/實際完成日)
        '                xNowDateTime = NowDateTime
        '                SQL = "Select FormNo, FormSno, Step, Unique_ID From T_WaitHandle "
        '                SQL &= "Where Active  = '1' "
        '                SQL &= "  And WorkID  = '" & xWorkID & "' "
        '                SQL &= "  And FlowType <> '0' "
        '                SQL &= "  And ASchedule = '1' "
        '                SQL &= "Order by BStartTime "
        '                dt_WaitHandle = uDataBase.GetDataTable(SQL)
        '                For i = 0 To dt_WaitHandle.Rows.Count - 1
        '                    xTableName = sCommon.GetTableName(dt_WaitHandle.Rows(i).Item("FormNo").ToString)
        '                    If dt_WaitHandle.Rows(i).Item("FormNo").ToString = "000001" Or dt_WaitHandle.Rows(i).Item("FormNo").ToString = "000002" Then
        '                        SQL = "Select Sample As Level, "" as QCLT  From F_" & xTableName
        '                    Else
        '                        If dt_WaitHandle.Rows(i).Item("FormNo").ToString = "000012" Then
        '                            SQL = "Select "" As Level, "" as QCLT  From F_" & xTableName
        '                        Else
        '                            SQL = "Select Level As Level, QCLT as QCLT From F_" & xTableName
        '                        End If
        '                    End If
        '                    SQL &= "Where FormNo  = '" & dt_WaitHandle.Rows(i).Item("FormNo").ToString & "' "
        '                    SQL &= "  And FormSno = '" & dt_WaitHandle.Rows(i).Item("FormSno").ToString & "' "
        '                    Dim dt_FormTable As DataTable = uDataBase.GetDataTable(SQL)
        '                    If dt_FormTable.Rows.Count > 0 Then
        '                        xLevel = dt_FormTable.Rows(0).Item("Level").ToString
        '                        If UCase(dt_FormTable.Rows(0).Item("Level").ToString) = "YES" Then
        '                            xLevel = "Z"
        '                        Else
        '                            If UCase(dt_FormTable.Rows(0).Item("Level").ToString) = "NO" Then
        '                                xLevel = "0"
        '                            End If
        '                        End If
        '                        If IsNumeric(dt_FormTable.Rows(0).Item("QCLT")) Then
        '                            xQCLT = dt_FormTable.Rows(0).Item("QCLT")
        '                        End If
        '                        If GetLeadTime(dt_FormTable.Rows(0).Item("FormNo").ToString, dt_FormTable.Rows(0).Item("Step"), xLevel, xQCLT, xNowDateTime, xStartTime, xEndTime, pDepo) = 0 Then
        '                            SQL = "Update T_WaitHandle Set "
        '                            SQL &= " BStartTime = '" & xStartTime & "',"
        '                            SQL &= " BEndTime = '" & xEndTime & "',"
        '                            SQL &= " AStartTime = '" & xStartTime & "',"
        '                            SQL &= " ModifyUser = '" & "ASchedule" & "',"
        '                            SQL &= " ModifyTime = '" & NowDateTime & "' "
        '                            SQL &= "Where Unique_ID = '" & dt_WaitHandle.Rows(i).Item("Unique_ID").ToString & "' "
        '                            uDataBase.ExecuteNonQuery(SQL)
        '                            xNowDateTime = xEndTime
        '                        End If
        '                    End If
        '                Next
        '            End If
        '        End If
        '    End If
        'Catch ex As Exception
        '    RtnCode = 1
        'End Try
        '****** 停止使用-End ******
        '+++++++
        Return RtnCode
    End Function
    'AdjustSchedule-End
    '**************************************************************************************
    '** 取得開始時間(GetStartTime)
    '**************************************************************************************
    'GetStartTime-Start
    <WebMethod()> _
    Public Function GetStartTime(ByVal pNowDateTime As String, ByVal pDepo As String) As String
        Dim RtnString As String = pNowDateTime
        Dim SQL As String
        '現時間(pNowDateTime)分解為->日期及分
        Dim xNowDate As String = CDate(pNowDateTime).ToString("yyyy/MM/dd")
        Dim xNowTime As Integer = CInt(CDate(pNowDateTime).ToString("HH")) * 60 + CInt(CDate(pNowDateTime).ToString("mm"))
        Try
            SQL = "Select SeqNo, YMD, Hour From M_Calendar "
            SQL &= "Where YMD  = '" & xNowDate & "' "
            SQL &= "  And Hour > '" & CStr(xNowTime) & "' "
            SQL &= "  And Active = '1' "
            SQL &= "  And Depo = '" & pDepo & "' "
            SQL &= "Order By YMD, Hour "
            Dim dt_Calendar As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Calendar.Rows.Count > 0 Then
                RtnString = dt_Calendar.Rows(0).Item("YMD").ToString("yyyy / MM / dd") & " " & _
                            CStr(Int(dt_Calendar.Rows(0).Item("Hour") / 60)) & ":" & _
                            CStr(dt_Calendar.Rows(0).Item("Hour") Mod 60) & ":" & _
                            "00"
                RtnString = RtnString.ToString("yyyy/mm/dd hh:mm:ss")
            Else
                SQL = "Select SeqNo, YMD, Hour From M_Calendar "
                SQL &= "Where YMD  > '" & xNowDate & "' "
                SQL &= "  And Active = '1' "
                SQL &= "  And Depo = '" & pDepo & "' "
                SQL &= "Order By YMD, Hour "
                dt_Calendar = uDataBase.GetDataTable(SQL)
                If dt_Calendar.Rows.Count > 0 Then
                    RtnString = dt_Calendar.Rows(0).Item("YMD").ToString("yyyy / MM / dd") & " " & _
                                CStr(Int(dt_Calendar.Rows(0).Item("Hour") / 60)) & ":" & _
                                CStr(dt_Calendar.Rows(0).Item("Hour") Mod 60) & ":" & _
                                "00"
                    RtnString = RtnString.ToString("yyyy/mm/dd hh:mm:ss")
                End If
            End If
        Catch ex As Exception
            RtnString = pNowDateTime
        End Try
        '
        Return RtnString
    End Function
    'GetStartTime-End
    '**************************************************************************************
    '** 取得工程LeadTime(GetLeadTime)
    '**************************************************************************************
    'GetLeadTime-Start
    <WebMethod()> _
    Public Function GetLeadTime(ByVal pFormNo As String, _
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
            If sCommon.GetSystemData("011", "JobLeadTime").ToString <> "" Then      '加上工程間等待時間
                xNowTime = xNowTime + CInt(sCommon.GetSystemData("011", "JobLeadTime").ToString)
            End If
            SQL = "Select SeqNo, YMD, Hour From M_Calendar "
            SQL &= "Where YMD  = '" & xNowDate & "' "
            SQL &= "  And Hour > '" & CStr(xNowTime) & "' "
            SQL &= "  And Active = '1' "
            SQL &= "  And Depo = '" & pDepo & "' "
            SQL &= "Order By YMD, Hour "
            Dim dt_Calendar As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Calendar.Rows.Count > 0 Then
                xSeqNo = dt_Calendar.Rows(0).Item("SeqNo")
                pStartTime = CDate(dt_Calendar.Rows(0).Item("YMD")).ToString("yyyy/MM/dd") & " " & _
                                   CStr(Int(dt_Calendar.Rows(0).Item("Hour") / 60)) & ":" & _
                                   CStr(dt_Calendar.Rows(0).Item("Hour") Mod 60) & ":" & _
                                   "00"
            Else
                SQL = "Select SeqNo, YMD, Hour From M_Calendar "
                SQL &= "Where YMD  > '" & xNowDate & "' "
                SQL &= "  And Active = '1' "
                SQL &= "  And Depo = '" & pDepo & "' "
                SQL &= "Order By YMD, Hour "
                dt_Calendar = uDataBase.GetDataTable(SQL)
                If dt_Calendar.Rows.Count > 0 Then
                    xSeqNo = dt_Calendar.Rows(0).Item("SeqNo")
                    pStartTime = CDate(dt_Calendar.Rows(0).Item("YMD")).ToString("yyyy/MM/dd") & " " & _
                                       CStr(Int(dt_Calendar.Rows(0).Item("Hour") / 60)) & ":" & _
                                       CStr(dt_Calendar.Rows(0).Item("Hour") Mod 60) & ":" & _
                                       "00"
                End If
            End If
            'MsgBox("xNowDate=[" + xNowDate + "]")
            'MsgBox("xNowTime=[" + CStr(xNowTime) + "]")
            'MsgBox("GetLeadTime-StartTime=[" + pStartTime + "]")
            '取得工程L/T
            If sCommon.GetQCLeadTime(pFormNo, pStep) = 0 Then       'QCL/T 採用=1, 不採用=0
                If sCommon.GetOPLeadTime(pFormNo, pStep, pLevel) > 0 Then
                    xLeadTime = sCommon.GetOPLeadTime(pFormNo, pStep, pLevel) / 10
                Else
                    If sCommon.GetOPLeadTime(pFormNo, pStep, "0") > 0 Then
                        xLeadTime = sCommon.GetOPLeadTime(pFormNo, pStep, "0") / 10
                    Else
                        xLeadTime = CInt(sCommon.GetSystemData("011", "ReadLeadTime").ToString) / 10
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
            dt_Calendar = uDataBase.GetDataTable(SQL)
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
    'GetLeadTime-End
    '**************************************************************************************
    '** 取得閱讀期限(GetReadTimeLimit)
    '**************************************************************************************
    'GetReadTimeLimit-Start
    <WebMethod()> _
    Public Function GetReadTimeLimit(ByVal pNowDateTime As String, ByVal pDepo As String) As String
        Dim RtnString As String = pNowDateTime
        Dim SQL As String
        '現時間(pNowDateTime)分解為->日期及分
        Dim xNowDate As String = CDate(pNowDateTime).ToString("yyyy/MM/dd")
        Dim xNowTime As Integer = CInt(CDate(pNowDateTime).ToString("HH")) * 60 + CInt(CDate(pNowDateTime).ToString("mm"))
        Try
            'MsgBox("date=" + xNowDate)
            'MsgBox("time=" + Str(xNowTime))
            '取得讀取開始日
            Dim xSeqNo As Integer = 0
            SQL = "Select SeqNo From M_Calendar "
            SQL &= "Where YMD  = '" & xNowDate & "' "
            SQL &= "  And Hour > '" & CStr(xNowTime) & "' "
            SQL &= "  And Active = '1' "
            SQL &= "  And Depo = '" & pDepo & "' "
            SQL &= "Order By YMD, Hour "
            Dim dt_Calendar As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Calendar.Rows.Count > 0 Then
                xSeqNo = dt_Calendar.Rows(0).Item("SeqNo")
            Else
                SQL = "Select SeqNo From M_Calendar "
                SQL &= "Where YMD  > '" & xNowDate & "' "
                SQL &= "  And Active = '1' "
                SQL &= "  And Depo = '" & pDepo & "' "
                SQL &= "Order By YMD, Hour "
                dt_Calendar = uDataBase.GetDataTable(SQL)
                If dt_Calendar.Rows.Count > 0 Then
                    xSeqNo = dt_Calendar.Rows(0).Item("SeqNo")
                End If
            End If
            'MsgBox("seqno=" + CStr(xSeqNo))
            '取得限制時間
            Dim xReadLeadTime As Integer = Int(CInt(sCommon.GetSystemData("011", "ReadLeadTime").ToString) / 10)
            'MsgBox("ReadLeadTime=" + CStr(xReadLeadTime))
            '取得讀取完成日
            SQL = "Select SeqNo, YMD, Hour From M_Calendar "
            SQL &= "Where SeqNo  >= '" & CStr(xSeqNo + xReadLeadTime) & "' "
            SQL &= "  And Active = '1' "
            SQL &= "  And Depo = '" & pDepo & "' "
            SQL &= "Order By SeqNo "
            dt_Calendar = uDataBase.GetDataTable(SQL)
            If dt_Calendar.Rows.Count > 0 Then
                RtnString = CDate(dt_Calendar.Rows(0).Item("YMD")).ToString("yyyy/MM/dd") & " " & _
                            CStr(Int(dt_Calendar.Rows(0).Item("Hour") / 60)) & ":" & _
                            CStr(dt_Calendar.Rows(0).Item("Hour") Mod 60) & ":" & _
                            "00"
                'MsgBox("Return=" + RtnString)
            End If
        Catch ex As Exception
            RtnString = pNowDateTime
        End Try
        '
        Return RtnString
    End Function
    'GetReadTimeLimit-End
    '**************************************************************************************
    '** 取得閱讀時間(GetReadLeadTime)
    '**************************************************************************************
    'GetReadLeadTime-Start
    <WebMethod()> _
    Public Function GetReadLeadTime(ByVal pNowDateTime As String, _
                                    ByRef pStartTime As String, _
                                    ByRef pEndTime As String, _
                                    ByVal pDepo As String) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        '現時間(pNowDateTime)分解為->日期及分
        Dim xNowDate As String = CDate(pNowDateTime).ToString("yyyy/MM/dd")
        Dim xNowTime As Integer = CInt(CDate(pNowDateTime).ToString("HH")) * 60 + CInt(CDate(pNowDateTime).ToString("mm"))
        Try
            '取得讀取開始日
            Dim xSeqNo As Integer = 0
            SQL = "Select SeqNo, YMD, Hour From M_Calendar "
            SQL &= "Where YMD  = '" & xNowDate & "' "
            SQL &= "  And Hour > '" & CStr(xNowTime) & "' "
            SQL &= "  And Active = '1' "
            SQL &= "  And Depo = '" & pDepo & "' "
            SQL &= "Order By YMD, Hour "
            Dim dt_Calendar As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Calendar.Rows.Count > 0 Then
                xSeqNo = dt_Calendar.Rows(0).Item("SeqNo")
                pStartTime = dt_Calendar.Rows(0).Item("YMD").ToString("yyyy/MM/dd") & " " & _
                             CStr(Int(dt_Calendar.Rows(0).Item("Hour") / 60)) & ":" & _
                             CStr(dt_Calendar.Rows(0).Item("Hour") Mod 60) & ":" & _
                             "00"
                pStartTime = pStartTime.ToString("yyyy/mm/dd hh:mm:ss")
            Else
                SQL = "Select SeqNo, YMD, Hour From M_Calendar "
                SQL &= "Where YMD  > '" & xNowDate & "' "
                SQL &= "  And Active = '1' "
                SQL &= "  And Depo = '" & pDepo & "' "
                SQL &= "Order By YMD, Hour "
                dt_Calendar = uDataBase.GetDataTable(SQL)
                If dt_Calendar.Rows.Count > 0 Then
                    xSeqNo = dt_Calendar.Rows(0).Item("SeqNo")
                    pStartTime = dt_Calendar.Rows(0).Item("YMD").ToString("yyyy/MM/dd") & " " & _
                                 CStr(Int(dt_Calendar.Rows(0).Item("Hour") / 60)) & ":" & _
                                 CStr(dt_Calendar.Rows(0).Item("Hour") Mod 60) & ":" & _
                                 "00"
                    pStartTime = pStartTime.ToString("yyyy/mm/dd hh:mm:ss")
                Else
                    RtnCode = 1
                End If
            End If
            '取得讀取完成日
            If RtnCode = 0 Then
                '取得限制時間
                Dim xReadLeadTime As Integer = Int(CInt(sCommon.GetSystemData("011", "ReadLeadTime").ToString) / 10)
                '取得讀取完成日
                SQL = "Select SeqNo, YMD, Hour From M_Calendar "
                SQL &= "Where SeqNo  >= '" & CStr(xSeqNo + xReadLeadTime) & "' "
                SQL &= "  And Active = '1' "
                SQL &= "  And Depo = '" & pDepo & "' "
                SQL &= "Order By SeqNo "
                dt_Calendar = uDataBase.GetDataTable(SQL)
                If dt_Calendar.Rows.Count > 0 Then
                    pEndTime = dt_Calendar.Rows(0).Item("YMD").ToString("yyyy/MM/dd") & " " & _
                               CStr(Int(dt_Calendar.Rows(0).Item("Hour") / 60)) & ":" & _
                               CStr(dt_Calendar.Rows(0).Item("Hour") Mod 60) & ":" & _
                               "00"
                    pEndTime = pEndTime.ToString("yyyy/mm/dd hh:mm:ss")
                Else
                    RtnCode = 1
                End If
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'GetReadLeadTime-End
    '
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    'Add-Start by Joy  2012/7/31 新納期對應
    '
    '**************************************************************************************
    '** 取得最後預定完成及待處理件數-客戶導向(GetLastTimeCustom)
    '**************************************************************************************
    'GetLastTimeCustom-Start
    <WebMethod()> _
    Public Function GetLastTimeCustom(ByVal pUser As String, _
                                      ByVal pFormNo As String, _
                                      ByVal pFormSno As Integer, _
                                      ByVal pNextStep As Integer, _
                                      ByVal pFlowType As Integer, _
                                      ByVal pNowDateTime As DateTime, _
                                      ByVal pSimulation As Integer, _
                                      ByRef pLastTime As DateTime, _
                                      ByRef pCount1 As Integer) As Integer
        'pSimulation=0 不含模擬 / pSimulation=1 含模擬 
        Dim RtnCode As Integer = 0
        Dim xStep As String
        Dim SQL As String
        Try
            '設定初始值
            pLastTime = pNowDateTime
            pCount1 = 0
            '取得WorkID
            If sCommon.GetWorkID(pUser).ToString <> "" Then
                '取得待處理件數
                SQL = "Select Count(*) As RecCount From T_WaitHandle "
                SQL &= "Where WorkID = '" & sCommon.GetWorkID(pUser).ToString & "' "
                SQL &= "  And FormNo = '" & pFormNo & "' "
                SQL &= "  And SeqNo  = '1' "
                SQL &= "  And Active = '1' "
                Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
                If dt_WaitHandle.Rows.Count > 0 Then
                    pCount1 = dt_WaitHandle.Rows(0).Item("RecCount")
                End If
                '
                '取得最後預定完成
                If pFlowType <> 0 Then
                    '取得上工程預定完成時間
                    SQL = "Select Step, BEndTime From T_WaitHandle "
                    SQL &= "Where FormNo  = '" & pFormNo & "' "
                    SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                    If pSimulation = 0 Then
                        SQL &= "  And FlowType = '1' "
                    Else
                        SQL &= "  And (FlowType = '1' or FlowType = '9') "
                    End If
                    SQL &= "Order by Unique_ID Desc "
                    dt_WaitHandle = uDataBase.GetDataTable(SQL)
                    If dt_WaitHandle.Rows.Count > 0 Then
                        xStep = dt_WaitHandle.Rows(0).Item("Step").ToString
                        pLastTime = CDate(dt_WaitHandle.Rows(0).Item("BEndTime").ToString)
                        '
                        '多流程核定工程
                        SQL = "Select FlowTypeStep From M_Flow "
                        SQL &= "Where FormNo = '" & pFormNo & "' "
                        SQL &= "  And Step   = '" & xStep & "' "
                        SQL &= "  And FlowTypeStep <> '" & "0" & "' "
                        SQL &= "  And Active = '1' "
                        Dim dt_Flow1 As DataTable = uDataBase.GetDataTable(SQL)
                        If dt_Flow1.Rows.Count > 0 Then
                            '
                            '取得多流程核定工程-上工程預定完成時間
                            SQL = "Select Step, BEndTime From T_WaitHandle "
                            SQL &= "Where FormNo  = '" & pFormNo & "' "
                            SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                            SQL &= "  And Step    < '" & xStep & "' "
                            If pSimulation = 0 Then
                                SQL &= "  And FlowType = '1' "
                            Else
                                SQL &= "  And (FlowType = '1' or FlowType = '9') "
                            End If
                            SQL &= "Order by Unique_ID Desc "
                            dt_WaitHandle = uDataBase.GetDataTable(SQL)
                            If dt_WaitHandle.Rows.Count > 0 Then
                                pLastTime = CDate(dt_WaitHandle.Rows(0).Item("BEndTime").ToString)
                            End If
                        End If
                    End If
                    '
                    '負荷工程
                    SQL = "Select Loading, Record From M_Flow "
                    SQL &= "Where FormNo = '" & pFormNo & "' "
                    SQL &= "  And Step   = '" & CStr(pNextStep) & "' "
                    SQL &= "  And Active = '1' "
                    Dim dt_Flow As DataTable = uDataBase.GetDataTable(SQL)
                    If dt_Flow.Rows.Count > 0 Then
                        If dt_Flow.Rows(0).Item("Loading").ToString = "1" Then
                            '考慮負荷
                            '取得最後待處理工程預定完成時間
                            SQL = "Select BEndTime From T_WaitHandle "
                            SQL &= "Where WorkID   = '" & sCommon.GetWorkID(pUser).ToString & "' "
                            SQL &= "  And FormNo   = '" & pFormNo & "' "
                            If pSimulation = 0 Then
                                SQL &= "  And FlowType = '1' "
                                SQL &= "  And Active   = '1' "
                            Else
                                SQL &= "  And (FlowType = '1' or FlowType = '9') "
                                SQL &= "  And (Active   = '1' or Active   = '9') "
                            End If
                            SQL &= "Order by BEndTime Desc "
                            dt_WaitHandle = uDataBase.GetDataTable(SQL)
                            If dt_WaitHandle.Rows.Count > 0 Then
                                '最後待處理工程時間>=上工程預定完成時間
                                If CDate(dt_WaitHandle.Rows(0).Item("BEndTime").ToString) >= pLastTime Then
                                    pLastTime = CDate(dt_WaitHandle.Rows(0).Item("BEndTime").ToString)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'GetLastTimeCustom-End
    '
    '**************************************************************************************
    '** 取得實際開始時間(GetActualStartTime)
    '**************************************************************************************
    'GetActualStartTime-Start
    <WebMethod()> _
    Public Function GetActualStartTime(ByVal pFormNo As String, _
                                       ByVal pFormSno As Integer, _
                                       ByVal pLoading As Integer, _
                                       ByVal pWorkID As String, _
                                       ByVal pDepo As String, _
                                       ByRef pAStartTime As String) As Integer
        Dim RtnCode As Integer = 0
        Dim xAEndTime As String = ""
        Dim SQL As String
        Try
            If pLoading = 1 Then
                ' 負荷=有
                ' 取得實際完成時間(同工程最後1筆)
                SQL = "Select AEndTime From T_WaitHandle "
                SQL &= "Where WorkID   = '" & pWorkID & "' "
                SQL &= "  And FormNo   = '" & pFormNo & "' "
                SQL &= "  And FlowType = '1' "
                SQL &= "  And Active   = '0' "
                SQL &= "Order by AEndTime Desc "
                Dim dt_AEndTime As DataTable = uDataBase.GetDataTable(SQL)
                If dt_AEndTime.Rows.Count > 0 Then
                    xAEndTime = dt_AEndTime.Rows(0).Item("AEndTime").ToString
                Else
                    ' 取得實際完成時間(上工程)
                    SQL = "Select AEndTime From T_WaitHandle "
                    SQL &= "Where FormNo  = '" & pFormNo & "' "
                    SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                    SQL &= "  And FlowType = '1' "
                    SQL &= "  And Active   = '0' "
                    SQL &= "Order by AEndTime Desc "
                    dt_AEndTime = uDataBase.GetDataTable(SQL)
                    If dt_AEndTime.Rows.Count > 0 Then
                        xAEndTime = dt_AEndTime.Rows(0).Item("AEndTime").ToString
                    End If
                End If
            Else
                ' 負荷=無
                ' 取得實際完成時間(上工程)
                SQL = "Select AEndTime From T_WaitHandle "
                SQL &= "Where FormNo  = '" & pFormNo & "' "
                SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                SQL &= "  And FlowType = '1' "
                SQL &= "  And Active   = '0' "
                SQL &= "Order by AEndTime Desc "
                Dim dt_AEndTime As DataTable = uDataBase.GetDataTable(SQL)
                If dt_AEndTime.Rows.Count > 0 Then
                    xAEndTime = dt_AEndTime.Rows(0).Item("AEndTime").ToString
                End If
            End If
            If xAEndTime <> "" Then
                ' 取得最後實際完成時間(整數)
                ' 現時間分解為->日期及分
                Dim xNowDate As String = CDate(xAEndTime).ToString("yyyy/MM/dd")
                Dim xNowTime As Integer = CInt(CDate(xAEndTime).ToString("HH")) * 60 + CInt(CDate(xAEndTime).ToString("mm"))
                '
                SQL = "Select SeqNo, YMD, Hour From M_Calendar "
                SQL &= "Where YMD  = '" & xNowDate & "' "
                SQL &= "  And Hour > '" & CStr(xNowTime) & "' "
                SQL &= "  And Active = '1' "
                SQL &= "  And Depo = '" & pDepo & "' "
                SQL &= "Order By YMD, Hour "
                Dim dt_Calendar As DataTable = uDataBase.GetDataTable(SQL)
                If dt_Calendar.Rows.Count > 0 Then
                    pAStartTime = CDate(dt_Calendar.Rows(0).Item("YMD")).ToString("yyyy/MM/dd")
                    pAStartTime = pAStartTime & " " & _
                                  CStr(Int(dt_Calendar.Rows(0).Item("Hour") / 60)) & ":" & _
                                  CStr(dt_Calendar.Rows(0).Item("Hour") Mod 60) & ":" & _
                                  "00"
                Else
                    SQL = "Select SeqNo, YMD, Hour From M_Calendar "
                    SQL &= "Where YMD  > '" & xNowDate & "' "
                    SQL &= "  And Active = '1' "
                    SQL &= "  And Depo = '" & pDepo & "' "
                    SQL &= "Order By YMD, Hour "
                    dt_Calendar = uDataBase.GetDataTable(SQL)
                    If dt_Calendar.Rows.Count > 0 Then
                        pAStartTime = CDate(dt_Calendar.Rows(0).Item("YMD")).ToString("yyyy/MM/dd")
                        pAStartTime = pAStartTime & " " & _
                                      CStr(Int(dt_Calendar.Rows(0).Item("Hour") / 60)) & ":" & _
                                      CStr(dt_Calendar.Rows(0).Item("Hour") Mod 60) & ":" & _
                                      "00"
                    Else
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
    'GetActualStartTime-End

    '**************************************************************************************
    '** 調整最後日時(AdjustLeadTime)
    '**************************************************************************************
    'AdjustLeadTime-Start

    <WebMethod()> _
    Public Function AdjustLeadTime(ByVal pFormNo As String, _
                                   ByVal pFormSno As Integer, _
                                   ByVal pNextStep As Integer, _
                                   ByVal pSeqNo As Integer, _
                                   ByVal pDepo As String, _
                                   ByVal pNowDateTime As DateTime, _
                                   ByRef pStartTime As DateTime, _
                                   ByRef pEndTime As DateTime) As Integer
        Dim RtnCode As Integer = 0
        Dim xBTime As String = ""
        Dim SQL As String
        Try
            SQL = "Select NGStep From M_Flow "
            SQL &= "Where FormNo = '" & pFormNo & "' "
            SQL &= "  And Step   = '" & CStr(pNextStep) & "' "
            SQL &= "  And Active = '1' "
            Dim dt_Flow As DataTable = uDataBase.GetDataTable(SQL)
            If dt_Flow.Rows.Count > 0 Then
                '
                ' 判斷下工程是否NG工程(NGStep 0:否, 1:是)	
                If dt_Flow.Rows(0).Item("NGStep").ToString = "1" Then
                    ' NG工程=是
                    ' 取得工程最後預定完成時間
                    SQL = "Select BStartTime From T_WaitHandle "
                    SQL &= "Where FormNo  =  '" & pFormNo & "' "
                    SQL &= "  And FormSno =  '" & CStr(pFormSno) & "' "
                    SQL &= "  And FlowType = '1' "
                    SQL &= "Order by BStartTime Desc "
                    Dim dt_WaitHandle = uDataBase.GetDataTable(SQL)
                    If dt_WaitHandle.Rows.Count > 0 Then
                        xBTime = dt_WaitHandle.Rows(0).Item("BStartTime").ToString
                    End If
                    '
                    pStartTime = CDate(xBTime)
                    pEndTime = CDate(xBTime)
                Else
                    ' NG工程=否
                    ' 判斷下工程是否重新處理
                    SQL = "Select BStartTime, BEndTime From T_WaitHandle "
                    SQL &= "Where FormNo  = '" & pFormNo & "' "
                    SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                    SQL &= "  And Step    = '" & CStr(pNextStep) & "' "
                    SQL &= "  And SeqNo   = '" & CStr(pSeqNo) & "' "
                    SQL &= "Order by Unique_ID Desc "
                    Dim dt_WaitHandle = uDataBase.GetDataTable(SQL)
                    If dt_WaitHandle.Rows.Count > 0 Then
                        pStartTime = CDate(dt_WaitHandle.Rows(0).Item("BStartTime").ToString)
                        pEndTime = CDate(dt_WaitHandle.Rows(0).Item("BEndTime").ToString)
                    Else
                        If pNextStep >= 40 And pNextStep <= 110 Then
                            ' 不是重新處理但是試作工程
                            SQL = "Select BStartTime, BEndTime From T_WaitHandle "
                            SQL &= "Where FormNo  = '" & pFormNo & "' "
                            SQL &= "  And FormSno = '" & CStr(pFormSno) & "' "
                            SQL &= "  And Step >= '" & CStr(40) & "' "
                            SQL &= "  And Step <= '" & CStr(110) & "' "
                            SQL &= "Order by Step Desc "
                            dt_WaitHandle = uDataBase.GetDataTable(SQL)
                            If dt_WaitHandle.Rows.Count > 0 Then
                                pStartTime = CDate(dt_WaitHandle.Rows(0).Item("BStartTime").ToString)
                                pEndTime = CDate(dt_WaitHandle.Rows(0).Item("BEndTime").ToString)
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'AdjustLeadTime-End

    '
    'Add-End

End Class
