
Partial Class _Default
    Inherits System.Web.UI.Page

    Dim oCommon As New CommonService
    Dim oFlow As New FlowService
    Dim oSchedule As New ScheduleService

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Not Me.IsPostBack Then
            '     Dim ws As New localhost.WFSService
            '     Dim return_answer As String = ws.HelloWorld
            '    TextBox1.Text = return_answer
        End If
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGetFormSeqNo.Click
        Dim RtnCode As Integer = 0
        Dim NewFormSno As Integer = 0
        RtnCode = oCommon.GetFormSeqNo(DFormNo.Text, NewFormSno)
        If RtnCode = 0 Then
            DRtnCode.Text = "Code=[" + CStr(RtnCode) + "];Sno=['" + CStr(NewFormSno) + "]'"
        Else
            DRtnCode.Text = "Code=[" + CStr(RtnCode) + "];Sno=['" + CStr(NewFormSno) + "]'"
        End If
    End Sub

    Protected Sub BNextFlow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNextFlow.Click
        Dim RtnCode As Integer = 0

        RtnCode = oFlow.NextFlow(DFormNo.Text, CInt(DFormSno.Text), CInt(DStep.Text), CInt(DSeqno.Text), DDepo.Text, DUser.Text, _
                                 DNextUser.Text, DAgentID.Text, DApplyID.Text, CDate(DStartTime.Text), CDate(DEndTime.Text), CInt(DImportant.Text))
        '表單號碼,表單流水號,工程關卡號碼,序號,申請者ID,台北,
        '建置者, 簽核者, 被代理者, 申請者, 預定開始日, 預定完成日, 重要性

        DRtnCode.Text = "Code=[" + CStr(RtnCode) + "]'"
    End Sub

    Protected Sub Button1_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim RtnCode As Integer = 0
        Dim pNextGate(10) As String
        Dim pAgentGate(10) As String
        Dim pNextStep As Integer = 0
        Dim pFlowType As Integer = 0    '0=通知
        Dim pCount As Integer

        'RtnCode = oCommon.GetNextGate("003105", CInt("1"), "IT004", "IT004", "", "", CInt("0"), _
        '                              pNextStep, pNextGate, pAgentGate, pCount, pFlowType, CInt("0"))

        RtnCode = oCommon.GetNextGate(DFormNo.Text, CInt(DStep.Text), DUser.Text, DApplyID.Text, DAgentID.Text, DAllocateID.Text, CInt(DMultiJob.Text), _
                                      pNextStep, pNextGate, pAgentGate, pCount, pFlowType, CInt(DAction.Text))
        '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
        '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 

        DRtnCode.Text = "Code=[" + CStr(RtnCode) + "],Next=[" + pNextGate(1) + "]"

    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim RtnCode As Integer = 0
        Dim pStartTime, pEndTime As DateTime

        oSchedule.GetLeadTime(DFormNo.Text, CInt(DStep.Text), DLevel.Text, CInt(DQCLT.Text), CDate(DStartTime.Text), pStartTime, pEndTime, DDepo.Text)
        '表單號碼,工程號碼,難易度,QC-L/T,現在時間, 預定開始日時, 預定完成日時, 行事曆

        DRtnCode.Text = "Code=[" + CStr(RtnCode) + "],Start=[" + pStartTime + "],End=[" + pEndTime + "]"

    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim RtnCode As String

        RtnCode = oSchedule.GetReadTimeLimit(DStartTime.Text, DDepo.Text)
        '現在日,行事曆

        DRtnCode.Text = "ReadtimeLimit=[" + RtnCode + "]"
    End Sub

    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim RtnCode As String

        RtnCode = oCommon.WriteWaitCheck(DFormNo.Text, CInt(DFormSno.Text), DStartTime.Text)
        '現在日,行事曆

        DRtnCode.Text = CStr(RtnCode)

    End Sub

    Protected Sub Button5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim RtnCode As Integer

        RtnCode = oCommon.Send(DFrom.Text, DTo.Text, DApplyID.Text, DFormNo.Text, CInt(DFormSno.Text), CInt(DStep.Text), DMsgType.Text)
        '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別

        DRtnCode.Text = CStr(RtnCode)
    End Sub

    Protected Sub Button6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim RtnCode As Integer
        RtnCode = oCommon.CommissionNo(DFormNo.Text, CInt(DFormSno.Text), CInt(DStep.Text), DNo.Text)
        '表單號碼, 表單流水號, 工程, 委託書No
        DRtnCode.Text = CStr(RtnCode)
    End Sub

    Protected Sub Button7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim RtnCode As String
        RtnCode = oCommon.GetCalendarGroup(DApplyID.Text)
        DRtnCode.Text = CStr(RtnCode)
    End Sub

    Protected Sub BNewFlow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNewFlow.Click
        Dim RtnCode As Integer = 0

        RtnCode = oFlow.NewFlow(DFormNo.Text, CInt(DFormSno.Text), CInt(DStep.Text), CInt(DSeqno.Text), DDepo.Text, DUser.Text, _
                                 DApplyID.Text)

        DRtnCode.Text = "Code=[" + CStr(RtnCode) + "]'"

    End Sub

    Protected Sub BEndFlow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BEndFlow.Click
        Dim RtnCode As Integer = 0

        RtnCode = oFlow.NewFlow(DFormNo.Text, CInt(DFormSno.Text), CInt(DStep.Text), CInt(DSeqno.Text), DDepo.Text, DUser.Text, _
                                 DApplyID.Text)

        DRtnCode.Text = "Code=[" + CStr(RtnCode) + "]'"

    End Sub

    Protected Sub Button8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button8.Click

        Dim RtnCode As Integer = 0

        '**        pFun=OK, NG1, NG2, SAVE  
        '**        pAction=0:OK, 1:NG1, 2:NG2, 3:Save   下一關卡時使用 
        '**        pSts=0:未處理, 1:OK, 2:NG1, 3:NG2, 4:已閱讀, 5:被抽單  更新T_Waithandle狀態
        '**             '                '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
        '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 

        '處理 pFun=OK, NG1, NG2, SAVE  
        '動作 pAction=0:OK, 1:NG1, 2:NG2, 3:Save   下一關卡時使用 
        '狀態 pSts=0:未處理, 1:OK, 2:NG1, 3:NG2, 4:已閱讀, 5:被抽單  更新T_Waithandle狀態
        '表單代號,單號, 工程代號,序號, 行事曆, 簽核者,申請者,被代理者,表單Table
        RtnCode = oCommon.BatchFlowControl("OK", 0, 1, "001151", 7, 10, 1, "TP1", "MK028", "TC010", "", "F_ITEMREGISTERSHEET")

        DRtnCode.Text = "Code=[" + CStr(RtnCode) + "]'"

    End Sub

    Protected Sub Button9_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim RtnCode As Integer = 0
        Dim wDevelopTime As Integer = 0

        RtnCode = oSchedule.GetDevelopTime(DStartTime.Text, DEndTime.Text, wDevelopTime, DDepo.Text)

        DRtnCode.Text = "Code=[" + CStr(RtnCode) + "]'"
        DErrorCode.Text = "Time=[" + CStr(wDevelopTime) + "]'"

    End Sub

    Protected Sub Button10_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim RtnCode As Integer = 0
        Dim wAStartTime As String = ""

        RtnCode = oSchedule.GetActualStartTime(DFormNo.Text, _
                                               CInt(DFormSno.Text), _
                                               1, _
                                               oCommon.GetWorkID(DUser.Text), _
                                               oCommon.GetCalendarGroup(DUser.Text), _
                                               wAStartTime)

        DRtnCode.Text = "Code=[" + CStr(RtnCode) + "]"
        DErrorCode.Text = "Time=[" + wAStartTime + "]"
    End Sub

    Protected Sub Button11_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim RtnCode As Integer = 0
        Dim xCount As Integer = 0
        Dim wLastTime As DateTime

        RtnCode = oSchedule.GetLastTimeCustom(DNextUser.Text, _
                                              DFormNo.Text, _
                                              CInt(DFormSno.Text), _
                                              40, _
                                              1, _
                                              CDate("2012/4/11 17:00:00"), _
                                              1, _
                                              wLastTime, _
                                              xCount)

        DRtnCode.Text = "Code=[" + CStr(RtnCode) + "]"
        DErrorCode.Text = "Count=[" + CStr(xCount) + "]Time=[" + CStr(wLastTime) + "]"
    End Sub

    Protected Sub Button12_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button12.Click

        'oCommon.ArrangFlowControl("002002", 5289, 20, "it003", "ea011", "tex002", 0)

    End Sub

    Protected Sub BUPDATEFC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BUPDATEFC.Click
        Dim RtnCode As Integer = 0

        'RtnCode = oCommon.UpdateFCData("AD-SLT^1746207^TN697^^^2017/12^SL022^", "1900934.00^1052737249200.00^")

        DRtnCode.Text = "Code=[" + CStr(RtnCode) + "]"
    End Sub

    Protected Sub Button14_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim RtnCode As Integer = 0

        RtnCode = oCommon.AgentApprovStart

        DRtnCode.Text = "Code=[" + CStr(RtnCode) + "]"

    End Sub

    Protected Sub BRecoveryFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BRecoveryFile.Click
        Dim RtnCode As Integer = 0
        '\\10.245.1.10\www$\WorkFlow\Document\000001
        '\\10.245.1.9\SyncData\inetpub\wwwroot\WorkFlow\Document\000001
        RtnCode = oCommon.RecoveryArchiveFile("\\10.245.1.10\www$\", "WorkFlow\Document\000003\", "4962-Map-20131218113028-Noname.jpg")

        'RtnCode = oCommon.RecoveryArchiveFile("\\10.245.1.10\www$\", "WorkFlow\Document\000003\", dt_MapFile.Rows(i)("MapFile").ToString)


        DRtnCode.Text = "Code=[" + CStr(RtnCode) + "]"
        'rtncode=0 正常 / 1異常

    End Sub

    Protected Sub Button13_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button13.Click

        Dim RtnCode As Integer = 0



        DRtnCode.Text = "Code=[" + CStr(RtnCode) + "]"

    End Sub
End Class
