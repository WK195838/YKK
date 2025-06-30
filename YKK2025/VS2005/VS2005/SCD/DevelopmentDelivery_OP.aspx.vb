Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class DevelopmentDelivery_OP
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDate As String           '現在日期
    Dim NowDateTime As String       '現在日期時間
    Dim FormNo As String            '表單
    Dim Formsno As Integer          '單號
    Dim StepIdx As Integer          '工程Index
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oSchedule As New Schedule.ScheduleService
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        SetPageAttribute()                          '設定畫面
        If Not IsPostBack Then                      'PostBack
            ShowData()                              '顯示資料 
            SetPageStepAttribute(StepIdx)           '設定畫面工程數
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
        Response.Cookies("PGM").Value = "DevelopmentDelivery_OP.aspx"                          '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDate = Now.ToString("yyyy/MM/dd")                  '現在日期
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日期時間
        FormNo = Request.QueryString("pFormNo")
        Formsno = Request.QueryString("pFormSno")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPageAttribute)
    '**     設定畫面
    '**
    '*****************************************************************
    Sub SetPageAttribute()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowData)
    '**     顯示資料
    '**
    '*****************************************************************
    Sub ShowData()
        Dim xNo As String = ""
        Dim xStep As Integer = 0
        Dim xWorkID As String = ""
        Dim xBeforeStep As Integer = 0
        Dim sql As String
        '
        StepIdx = 0
        sql = "Select FormNo, FormSno, Step, WorkID From T_WaitHandle "
        sql &= "Where FormNo  = '" & FormNo & "' "
        sql &= "  And FormSno = '" & CStr(Formsno) & "' "
        sql &= "  And Step Between 40 and 130 "
        sql &= "  And (Step < 120 or Step > 129) "
        sql &= "  And (FlowType = '1' or FlowType = '9') "
        sql &= "Order by FormNo, FormSno, Step, Unique_ID, WorkID "
        Dim dt_WaitHandle = uDataBase.GetDataTable(sql)
        '
        If dt_WaitHandle.Rows.Count > 0 Then
            For i As Integer = 0 To dt_WaitHandle.Rows.Count - 1
                If xStep <> dt_WaitHandle.Rows(i).Item("Step") Or xWorkID <> dt_WaitHandle.Rows(i).Item("WorkID") Then
                    '
                    StepIdx = StepIdx + 1
                    If xBeforeStep <> dt_WaitHandle.Rows(i).Item("Step") Then
                        '無重覆Step
                        ShowStepData(StepIdx, FormNo, Formsno, dt_WaitHandle.Rows(i).Item("Step"), dt_WaitHandle.Rows(i).Item("WorkID"), 0)
                    Else        '重覆Step
                        ShowStepData(StepIdx, FormNo, Formsno, dt_WaitHandle.Rows(i).Item("Step"), dt_WaitHandle.Rows(i).Item("WorkID"), 1)
                    End If
                    xBeforeStep = dt_WaitHandle.Rows(i).Item("Step")
                    '
                End If
                xStep = dt_WaitHandle.Rows(i).Item("Step")
                xWorkID = dt_WaitHandle.Rows(i).Item("WorkID")
            Next
            LHistory.NavigateUrl = "DevelopmentDelivery_History.aspx?pFormNo=" + FormNo + "&pFormSno=" + CStr(Formsno)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowStepData)
    '**     顯示工程資料
    '**
    '*****************************************************************
    Sub ShowStepData(ByVal pIdx As Integer, ByVal pFormno As String, ByVal pFormSno As Integer, ByVal pStep As Integer, ByVal pWID As String, ByVal pRepeat As Integer)
        Dim xNo As String = ""
        '
        Select Case pIdx
            Case 1
                DOP1Step.Text = GetActiveStep(pFormno, pFormSno, pStep)
                DOP1Name.Text = GetManufactureOPInf(pFormno, pFormSno, pStep, pWID, "OP1")
                DOP1Count.Text = GetManufactureCount(pFormno, pFormSno, pStep, DOP1Name.Text)
                DOP1Yotei.Text = "預定"
                DOP1Jisseki.Text = "實績"
                If pRepeat = 0 Then
                    DOP1BDate.Text = GetBDate(pFormno, pFormSno, pStep, pWID)
                    DOP1BTime.Text = GetBTime(pFormno, pFormSno, pStep, pWID)
                Else
                    DOP1BDate.Text = ""
                    DOP1BTime.Text = ""
                End If
                DOP1ADate.Text = GetADate(pFormno, pFormSno, pStep, pStep, pWID)
                DOP1ATime.Text = GetATime(pFormno, pFormSno, pStep, pStep, pWID)
                '
                If HaveManufactureHistory(pFormno, pFormSno, pStep, pWID) = 1 Then
                    LOP1Detail.Visible = True
                    LOP1Detail.NavigateUrl = "DevelopmentDelivery_OPHistory.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pStep=" + CStr(pStep) + "&pOPCode=" + fpObj.GetOPCode("CODE-" & DOP1Name.Text)
                Else
                    LOP1Detail.Style("Left") = 9999
                End If
                '
                LOP1Loading.Visible = True
                LOP1Loading.NavigateUrl = "DevelopmentDelivery_Load.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pWID=" + pWID + "&pStartTime=" + Mid(GetBDate(pFormno, pFormSno, pStep, pWID), 1, 10)
            Case 2
                DOP2Step.Text = GetActiveStep(pFormno, pFormSno, pStep)
                DOP2Name.Text = GetManufactureOPInf(pFormno, pFormSno, pStep, pWID, "OP2")
                DOP2Count.Text = GetManufactureCount(pFormno, pFormSno, pStep, DOP2Name.Text)
                DOP2Yotei.Text = "預定"
                DOP2Jisseki.Text = "實績"
                If pRepeat = 0 Then
                    DOP2BDate.Text = GetBDate(pFormno, pFormSno, pStep, pWID)
                    DOP2BTime.Text = GetBTime(pFormno, pFormSno, pStep, pWID)
                Else
                    DOP2BDate.Text = ""
                    DOP2BTime.Text = ""
                End If
                DOP2ADate.Text = GetADate(pFormno, pFormSno, pStep, pStep, pWID)
                DOP2ATime.Text = GetATime(pFormno, pFormSno, pStep, pStep, pWID)
                '
                If HaveManufactureHistory(pFormno, pFormSno, pStep, pWID) = 1 Then
                    LOP2Detail.Visible = True
                    LOP2Detail.NavigateUrl = "DevelopmentDelivery_OPHistory.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pStep=" + CStr(pStep) + "&pOPCode=" + fpObj.GetOPCode("CODE-" & DOP2Name.Text)
                Else
                    LOP2Detail.Style("Left") = 9999
                End If
                '
                LOP2Loading.Visible = True
                LOP2Loading.NavigateUrl = "DevelopmentDelivery_Load.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pWID=" + pWID + "&pStartTime=" + Mid(GetBDate(pFormno, pFormSno, pStep, pWID), 1, 10)
            Case 3
                DOP3Step.Text = GetActiveStep(pFormno, pFormSno, pStep)
                DOP3Name.Text = GetManufactureOPInf(pFormno, pFormSno, pStep, pWID, "OP3")
                DOP3Count.Text = GetManufactureCount(pFormno, pFormSno, pStep, DOP3Name.Text)
                DOP3Yotei.Text = "預定"
                DOP3Jisseki.Text = "實績"

                If pRepeat = 0 Then
                    DOP3BDate.Text = GetBDate(pFormno, pFormSno, pStep, pWID)
                    DOP3BTime.Text = GetBTime(pFormno, pFormSno, pStep, pWID)
                Else
                    DOP3BDate.Text = ""
                    DOP3BTime.Text = ""
                End If
                DOP3ADate.Text = GetADate(pFormno, pFormSno, pStep, pStep, pWID)
                DOP3ATime.Text = GetATime(pFormno, pFormSno, pStep, pStep, pWID)
                '
                If HaveManufactureHistory(pFormno, pFormSno, pStep, pWID) = 1 Then
                    LOP3Detail.Visible = True
                    LOP3Detail.NavigateUrl = "DevelopmentDelivery_OPHistory.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pStep=" + CStr(pStep) + "&pOPCode=" + fpObj.GetOPCode("CODE-" & DOP3Name.Text)
                Else
                    LOP3Detail.Style("Left") = 9999
                End If
                '
                LOP3Loading.Visible = True
                LOP3Loading.NavigateUrl = "DevelopmentDelivery_Load.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pWID=" + pWID + "&pStartTime=" + Mid(GetBDate(pFormno, pFormSno, pStep, pWID), 1, 10)
            Case 4
                DOP4Step.Text = GetActiveStep(pFormno, pFormSno, pStep)
                DOP4Name.Text = GetManufactureOPInf(pFormno, pFormSno, pStep, pWID, "OP4")
                DOP4Count.Text = GetManufactureCount(pFormno, pFormSno, pStep, DOP4Name.Text)
                DOP4Yotei.Text = "預定"
                DOP4Jisseki.Text = "實績"

                If pRepeat = 0 Then
                    DOP4BDate.Text = GetBDate(pFormno, pFormSno, pStep, pWID)
                    DOP4BTime.Text = GetBTime(pFormno, pFormSno, pStep, pWID)
                Else
                    DOP4BDate.Text = ""
                    DOP4BTime.Text = ""
                End If
                DOP4ADate.Text = GetADate(pFormno, pFormSno, pStep, pStep, pWID)
                DOP4ATime.Text = GetATime(pFormno, pFormSno, pStep, pStep, pWID)
                '
                If HaveManufactureHistory(pFormno, pFormSno, pStep, pWID) = 1 Then
                    LOP4Detail.Visible = True
                    LOP4Detail.NavigateUrl = "DevelopmentDelivery_OPHistory.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pStep=" + CStr(pStep) + "&pOPCode=" + fpObj.GetOPCode("CODE-" & DOP4Name.Text)
                Else
                    LOP4Detail.Style("Left") = 9999
                End If
                '
                LOP4Loading.Visible = True
                LOP4Loading.NavigateUrl = "DevelopmentDelivery_Load.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pWID=" + pWID + "&pStartTime=" + Mid(GetBDate(pFormno, pFormSno, pStep, pWID), 1, 10)
            Case 5
                DOP5Step.Text = GetActiveStep(pFormno, pFormSno, pStep)
                DOP5Name.Text = GetManufactureOPInf(pFormno, pFormSno, pStep, pWID, "OP5")
                DOP5Count.Text = GetManufactureCount(pFormno, pFormSno, pStep, DOP5Name.Text)
                DOP5Yotei.Text = "預定"
                DOP5Jisseki.Text = "實績"

                If pRepeat = 0 Then
                    DOP5BDate.Text = GetBDate(pFormno, pFormSno, pStep, pWID)
                    DOP5BTime.Text = GetBTime(pFormno, pFormSno, pStep, pWID)
                Else
                    DOP5BDate.Text = ""
                    DOP5BTime.Text = ""
                End If
                DOP5ADate.Text = GetADate(pFormno, pFormSno, pStep, pStep, pWID)
                DOP5ATime.Text = GetATime(pFormno, pFormSno, pStep, pStep, pWID)
                '
                If HaveManufactureHistory(pFormno, pFormSno, pStep, pWID) = 1 Then
                    LOP5Detail.Visible = True
                    LOP5Detail.NavigateUrl = "DevelopmentDelivery_OPHistory.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pStep=" + CStr(pStep) + "&pOPCode=" + fpObj.GetOPCode("CODE-" & DOP5Name.Text)
                Else
                    LOP5Detail.Style("Left") = 9999
                End If
                '
                LOP5Loading.Visible = True
                LOP5Loading.NavigateUrl = "DevelopmentDelivery_Load.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pWID=" + pWID + "&pStartTime=" + Mid(GetBDate(pFormno, pFormSno, pStep, pWID), 1, 10)
            Case 6
                DOP6Step.Text = GetActiveStep(pFormno, pFormSno, pStep)
                DOP6Name.Text = GetManufactureOPInf(pFormno, pFormSno, pStep, pWID, "OP6")
                DOP6Count.Text = GetManufactureCount(pFormno, pFormSno, pStep, DOP6Name.Text)
                DOP6Yotei.Text = "預定"
                DOP6Jisseki.Text = "實績"

                If pRepeat = 0 Then
                    DOP6BDate.Text = GetBDate(pFormno, pFormSno, pStep, pWID)
                    DOP6BTime.Text = GetBTime(pFormno, pFormSno, pStep, pWID)
                Else
                    DOP6BDate.Text = ""
                    DOP6BTime.Text = ""
                End If
                DOP6ADate.Text = GetADate(pFormno, pFormSno, pStep, pStep, pWID)
                DOP6ATime.Text = GetATime(pFormno, pFormSno, pStep, pStep, pWID)
                '
                If HaveManufactureHistory(pFormno, pFormSno, pStep, pWID) = 1 Then
                    LOP6Detail.Visible = True
                    LOP6Detail.NavigateUrl = "DevelopmentDelivery_OPHistory.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pStep=" + CStr(pStep) + "&pOPCode=" + fpObj.GetOPCode("CODE-" & DOP6Name.Text)
                Else
                    LOP6Detail.Style("Left") = 9999
                End If
                '
                LOP6Loading.Visible = True
                LOP6Loading.NavigateUrl = "DevelopmentDelivery_Load.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pWID=" + pWID + "&pStartTime=" + Mid(GetBDate(pFormno, pFormSno, pStep, pWID), 1, 10)
            Case 7
                DOP7Step.Text = GetActiveStep(pFormno, pFormSno, pStep)
                DOP7Name.Text = GetManufactureOPInf(pFormno, pFormSno, pStep, pWID, "OP7")
                DOP7Count.Text = GetManufactureCount(pFormno, pFormSno, pStep, DOP7Name.Text)
                DOP7Yotei.Text = "預定"
                DOP7Jisseki.Text = "實績"

                If pRepeat = 0 Then
                    DOP7BDate.Text = GetBDate(pFormno, pFormSno, pStep, pWID)
                    DOP7BTime.Text = GetBTime(pFormno, pFormSno, pStep, pWID)
                Else
                    DOP7BDate.Text = ""
                    DOP7BTime.Text = ""
                End If
                DOP7ADate.Text = GetADate(pFormno, pFormSno, pStep, pStep, pWID)
                DOP7ATime.Text = GetATime(pFormno, pFormSno, pStep, pStep, pWID)
                '
                If HaveManufactureHistory(pFormno, pFormSno, pStep, pWID) = 1 Then
                    LOP7Detail.Visible = True
                    LOP7Detail.NavigateUrl = "DevelopmentDelivery_OPHistory.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pStep=" + CStr(pStep) + "&pOPCode=" + fpObj.GetOPCode("CODE-" & DOP7Name.Text)
                Else
                    LOP7Detail.Style("Left") = 9999
                End If
                '
                LOP7Loading.Visible = True
                LOP7Loading.NavigateUrl = "DevelopmentDelivery_Load.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pWID=" + pWID + "&pStartTime=" + Mid(GetBDate(pFormno, pFormSno, pStep, pWID), 1, 10)
            Case 8
                DOP8Step.Text = GetActiveStep(pFormno, pFormSno, pStep)
                DOP8Name.Text = GetManufactureOPInf(pFormno, pFormSno, pStep, pWID, "OP8")
                DOP8Count.Text = GetManufactureCount(pFormno, pFormSno, pStep, DOP8Name.Text)
                DOP8Yotei.Text = "預定"
                DOP8Jisseki.Text = "實績"

                If pRepeat = 0 Then
                    DOP8BDate.Text = GetBDate(pFormno, pFormSno, pStep, pWID)
                    DOP8BTime.Text = GetBTime(pFormno, pFormSno, pStep, pWID)
                Else
                    DOP8BDate.Text = ""
                    DOP8BTime.Text = ""
                End If
                DOP8ADate.Text = GetADate(pFormno, pFormSno, pStep, pStep, pWID)
                DOP8ATime.Text = GetATime(pFormno, pFormSno, pStep, pStep, pWID)
                '
                If HaveManufactureHistory(pFormno, pFormSno, pStep, pWID) = 1 Then
                    LOP8Detail.Visible = True
                    LOP8Detail.NavigateUrl = "DevelopmentDelivery_OPHistory.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pStep=" + CStr(pStep) + "&pOPCode=" + fpObj.GetOPCode("CODE-" & DOP8Name.Text)
                Else
                    LOP8Detail.Style("Left") = 9999
                End If
                '
                LOP8Loading.Visible = True
                LOP8Loading.NavigateUrl = "DevelopmentDelivery_Load.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pWID=" + pWID + "&pStartTime=" + Mid(GetBDate(pFormno, pFormSno, pStep, pWID), 1, 10)
            Case Else
                DOP9Step.Text = GetActiveStep(pFormno, pFormSno, pStep)
                DOP9Name.Text = GetManufactureOPInf(pFormno, pFormSno, pStep, pWID, "OP9")
                DOP9Count.Text = GetManufactureCount(pFormno, pFormSno, pStep, DOP9Name.Text)
                DOP9Yotei.Text = "預定"
                DOP9Jisseki.Text = "實績"

                If pRepeat = 0 Then
                    DOP9BDate.Text = GetBDate(pFormno, pFormSno, pStep, pWID)
                    DOP9BTime.Text = GetBTime(pFormno, pFormSno, pStep, pWID)
                Else
                    DOP9BDate.Text = ""
                    DOP9BTime.Text = ""
                End If
                DOP9ADate.Text = GetADate(pFormno, pFormSno, pStep, pStep, pWID)
                DOP9ATime.Text = GetATime(pFormno, pFormSno, pStep, pStep, pWID)
                '
                If HaveManufactureHistory(pFormno, pFormSno, pStep, pWID) = 1 Then
                    LOP9Detail.Visible = True
                    LOP9Detail.NavigateUrl = "DevelopmentDelivery_OPHistory.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pStep=" + CStr(pStep) + "&pOPCode=" + fpObj.GetOPCode("CODE-" & DOP9Name.Text)
                Else
                    LOP9Detail.Style("Left") = 9999
                End If
                '
                LOP9Loading.Visible = True
                LOP9Loading.NavigateUrl = "DevelopmentDelivery_Load.aspx?pFormNo=" + pFormno + "&pFormSno=" + CStr(pFormSno) + "&pWID=" + pWID + "&pStartTime=" + Mid(GetBDate(pFormno, pFormSno, pStep, pWID), 1, 10)
        End Select
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetFlowLoading)
    '**     取得負荷狀態
    '**
    '*****************************************************************
    Function GetFlowLoading(ByVal pOPName As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        '
        sql = "Select * From M_Referp "
        sql &= "Where Cat  = '2002' "
        sql &= "  And DKey = '" & "OPLOAD-" & pOPName & "' "
        Dim dt_Referp = uDataBase.GetDataTable(sql)
        If dt_Referp.Rows.Count > 0 Then
            RtnCode = 1
        End If
        '
        Return RtnCode
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetActiveStep)
    '**     取得執行工程
    '**
    '*****************************************************************
    Function GetActiveStep(ByVal pFormNo As String, ByVal pFormSno As Integer, ByVal pStep As Integer) As String
        Dim RtnString As String = ""
        Dim sql As String
        '
        sql = "Select * From T_WaitHandle  "
        sql &= "Where FormNo  = '" & pFormNo & "' "
        sql &= "  And FormSno = '" & CStr(pFormSno) & "' "
        sql &= "  And Step = '" & CStr(pStep) & "' "
        sql &= "  And Active = '1' "
        Dim dt_WaitHandle = uDataBase.GetDataTable(sql)
        If dt_WaitHandle.Rows.Count > 0 Then
            RtnString = "* " + CStr(pStep)
        Else
            RtnString = CStr(pStep)
        End If
        '
        Return RtnString
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetCommissionInf)
    '**     取得開發委託資料
    '**
    '*****************************************************************
    Function GetCommissionInf(ByVal pFormNo As String, ByVal pFormSno As Integer, ByVal pField As String) As String
        Dim RtnString As String = ""
        Dim sql As String
        '
        sql = "Select * From V_CommissionSheet_02 "
        sql &= "Where FormNo  = '" & pFormNo & "' "
        sql &= "  And FormSno = '" & CStr(pFormSno) & "' "
        Dim dt_Commission = uDataBase.GetDataTable(sql)
        If dt_Commission.Rows.Count > 0 Then
            RtnString = dt_Commission.Rows(0).Item(pField).ToString
        End If
        '
        Return RtnString
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetManufactureOPInf)
    '**     取得製造委託資料
    '**
    '*****************************************************************
    Function GetManufactureOPInf(ByVal pFormNo As String, ByVal pFormSno As Integer, ByVal pStep As Integer, ByVal pWID As String, ByVal pOP As String) As String
        Dim RtnString As String = ""
        Dim sql As String
        '
        If pStep <> 130 Then
            sql = "Select OP From FS_ManufactureHistory "
            sql &= "Where FormNo  = '" & pFormNo & "' "
            sql &= "  And FormSno = '" & CStr(pFormSno) & "' "
            sql &= "  And Step    = '" & CStr(pStep) & "' "
            sql &= "  And CreateUser = '" & pWID & "' "
            Dim dt_ManufactureHistory = uDataBase.GetDataTable(sql)
            If dt_ManufactureHistory.Rows.Count > 0 Then
                RtnString = dt_ManufactureHistory.Rows(0).Item("OP").ToString
            Else
                sql = "Select No From F_CommissionSheet "
                sql &= "Where FormNo  = '" & pFormNo & "' "
                sql &= "  And FormSno = '" & CStr(pFormSno) & "' "
                Dim dt_CommissionSheet = uDataBase.GetDataTable(sql)
                If dt_CommissionSheet.Rows.Count > 0 Then
                    sql = "Select OP1, OP2, OP3, OP4, OP5, OP6, OP7, OP8 From FS_ManufactureSheet "
                    sql &= "Where No  = '" & dt_CommissionSheet.Rows(0).Item("No") & "' "
                    Dim dt_ManufactureSheet = uDataBase.GetDataTable(sql)
                    If dt_ManufactureSheet.Rows.Count > 0 Then
                        RtnString = dt_ManufactureSheet.Rows(0).Item(pOP).ToString
                    End If
                End If
            End If
        Else
            RtnString = "準備品質檢測"
        End If
        '
        Return RtnString
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetManufactureCount)
    '**     取得試作次數
    '**
    '*****************************************************************
    Function GetManufactureCount(ByVal pFormNo As String, ByVal pFormSno As Integer, ByVal pStep As Integer, ByVal pOP As String) As String
        Dim RtnString As String = ""
        Dim sql As String
        '
        sql = "select Count(*) As RCount from FS_ManufactureHistory "
        sql &= "where formno = '" + pFormNo + "' "
        sql &= "  and formsno ='" + CStr(pFormSno) + "' "
        sql &= "  and step = '" + CStr(pStep) + "' "
        sql &= "  and OPCode = '" + fpObj.GetOPCode("CODE-" & pOP) + "' "
        Dim dt_ManufactureHistory = uDataBase.GetDataTable(sql)
        If dt_ManufactureHistory.Rows(0).Item("RCount") > 1 Then
            RtnString = "(" + dt_ManufactureHistory.Rows(0).Item("RCount").ToString + ")"
        End If
        '
        Return RtnString
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(HaveManufactureHistory)
    '**     有製造履歷表
    '**
    '*****************************************************************
    Function HaveManufactureHistory(ByVal pFormNo As String, ByVal pFormSno As Integer, ByVal pStep As Integer, ByVal pWID As String) As Integer
        Dim RtnCode As String = 0
        Dim sql As String
        '
        sql = "Select * From FS_ManufactureHistory "
        sql &= "Where FormNo  = '" & pFormNo & "' "
        sql &= "  And FormSno = '" & CStr(pFormSno) & "' "
        sql &= "  And Step    = '" & CStr(pStep) & "' "
        sql &= "  And CreateUser = '" & pWID & "' "
        Dim dt_ManufactureHistory = uDataBase.GetDataTable(sql)
        If dt_ManufactureHistory.Rows.Count > 0 Then
            RtnCode = 1
        End If
        '
        Return RtnCode
    End Function


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetBDate)
    '**     取得預定完成日
    '**
    '*****************************************************************
    Function GetBDate(ByVal pFormNo As String, ByVal pFormSno As Integer, ByVal pStep As Integer, ByVal pWID As String) As String
        Dim RtnString As String = ""
        Dim sql As String
        '
        sql = "Select BEndTime From T_WaitHandle "
        sql &= "Where FormNo  = '" & pFormNo & "' "
        sql &= "  And FormSno = '" & CStr(pFormSno) & "' "
        sql &= "  And Step    = '" & CStr(pStep) & "' "
        sql &= "  And WorkID  = '" & pWID & "' "
        sql &= "Order by BEndTime "
        Dim dt_WaitHandle = uDataBase.GetDataTable(sql)
        If dt_WaitHandle.Rows.Count > 0 Then
            RtnString = CDate(dt_WaitHandle.Rows(0).Item("BEndTime")).ToString("yyyy/MM/dd HH:mm")
        End If
        '
        Return RtnString
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetBDate)
    '**     取得預定時間
    '**
    '*****************************************************************
    Function GetBTime(ByVal pFormNo As String, ByVal pFormSno As Integer, ByVal pStep As Integer, ByVal pWID As String) As String
        Dim xBTime As Integer = 0
        Dim RtnString As String = ""
        Dim sql As String
        '
        sql = "Select WorkID, BStartTime, BEndTime From T_WaitHandle "
        sql &= "Where FormNo  = '" & pFormNo & "' "
        sql &= "  And FormSno = '" & CStr(pFormSno) & "' "
        sql &= "  And Step    = '" & CStr(pStep) & "' "
        sql &= "  And WorkID  = '" & pWID & "' "
        sql &= "Order by BEndTime "
        Dim dt_WaitHandle = uDataBase.GetDataTable(sql)
        If dt_WaitHandle.Rows.Count > 0 Then
            oSchedule.GetDevelopTime(CDate(dt_WaitHandle.Rows(0).Item("BStartTime")).ToString("yyyy/MM/dd HH:mm"), _
                                     CDate(dt_WaitHandle.Rows(0).Item("BEndTime")).ToString("yyyy/MM/dd HH:mm"), _
                                     xBTime, _
                                     oCommon.GetCalendarGroup(dt_WaitHandle.Rows(0).Item("WorkID")))
            RtnString = CStr(xBTime)
        End If
        '
        Return RtnString
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetADate)
    '**     取得實際完成日
    '**
    '*****************************************************************
    Function GetADate(ByVal pFormNo As String, ByVal pFormSno As Integer, ByVal pStartStep As Integer, ByVal pEndStep As Integer, ByVal pWID As String) As String
        Dim RtnString As String = ""
        Dim sql As String
        '
        sql = "Select AEndTime From T_WaitHandle "
        sql &= "Where FormNo  = '" & pFormNo & "' "
        sql &= "  And FormSno = '" & CStr(pFormSno) & "' "
        sql &= "  And Step   >= '" & CStr(pStartStep) & "' "
        sql &= "  And Step   <= '" & CStr(pEndStep) & "' "
        sql &= "  And WorkID  = '" & pWID & "' "
        sql &= "Order by AEndTime Desc "
        Dim dt_WaitHandle = uDataBase.GetDataTable(sql)
        If dt_WaitHandle.Rows.Count > 0 Then
            If Not IsDBNull(dt_WaitHandle.Rows(0).Item("AEndTime")) Then
                RtnString = CDate(dt_WaitHandle.Rows(0).Item("AEndTime")).ToString("yyyy/MM/dd HH:mm")
            End If
        End If
        '
        Return RtnString
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetATime)
    '**     取得實績時間
    '**
    '*****************************************************************
    Function GetATime(ByVal pFormNo As String, ByVal pFormSno As Integer, ByVal pStartStep As Integer, ByVal pEndStep As Integer, ByVal pWID As String) As String
        Dim xATime As Integer = 0
        Dim xATimeSum As Integer = 0
        Dim RtnString As String = ""
        Dim sql As String
        '
        sql = "Select WorkID, AStartTime, AEndTime From T_WaitHandle "
        sql &= "Where FormNo  = '" & pFormNo & "' "
        sql &= "  And FormSno = '" & CStr(pFormSno) & "' "
        sql &= "  And Step   >= '" & CStr(pStartStep) & "' "
        sql &= "  And Step   <= '" & CStr(pEndStep) & "' "
        sql &= "  And WorkID  = '" & pWID & "' "
        sql &= "Order by AEndTime "
        Dim dt_WaitHandle = uDataBase.GetDataTable(sql)
        For i As Integer = 0 To dt_WaitHandle.Rows.Count - 1
            If Not IsDBNull(dt_WaitHandle.Rows(i).Item("AEndTime")) Then
                oSchedule.GetDevelopTime(CDate(dt_WaitHandle.Rows(i).Item("AStartTime")).ToString("yyyy/MM/dd HH:mm"), _
                                         CDate(dt_WaitHandle.Rows(i).Item("AEndTime")).ToString("yyyy/MM/dd HH:mm"), _
                                         xATime, _
                                         oCommon.GetCalendarGroup(dt_WaitHandle.Rows(0).Item("WorkID")))
                xATimeSum = xATimeSum + xATime
            End If
        Next
        If xATimeSum < 0 Then xATimeSum = 0
        RtnString = CStr(xATimeSum)
        '
        Return RtnString
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPageStepAttribute)
    '**     設定首頁工程數屬性
    '**
    '*****************************************************************
    Sub SetPageStepAttribute(ByVal pIdx As Integer)
        Dim xcom As String = ""
        Dim ximg As System.Web.UI.WebControls.Image
        Dim xlink As System.Web.UI.WebControls.HyperLink
        Dim xtxt As TextBox
        For i As Integer = 1 To 9
            If i <= pIdx Then
                'Image
                xcom = "DOP" + CStr(i) + "Image"
                ximg = Me.form1.FindControl(xcom)
                ximg.Visible = True
                'Step
                xcom = "DOP" + CStr(i) + "Step"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = True
                'Name
                xcom = "DOP" + CStr(i) + "Name"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = True
                'Count
                xcom = "DOP" + CStr(i) + "Count"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = True
                'Yotei
                xcom = "DOP" + CStr(i) + "Yotei"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = True
                'Jisseki
                xcom = "DOP" + CStr(i) + "Jisseki"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = True
                'Detail
                xcom = "LOP" + CStr(i) + "Detail"
                xlink = Me.form1.FindControl(xcom)
                xlink.Visible = True
                If xlink.Style("Left") = "9999" Then
                    xlink.Visible = False
                Else
                    xlink.Visible = True
                End If
                'Loading
                xcom = "LOP" + CStr(i) + "Loading"
                xlink = Me.form1.FindControl(xcom)
                If xlink.Style("Left") = "9999" Then
                    xlink.Visible = False
                Else
                    xlink.Visible = True
                End If
                'BDate
                xcom = "DOP" + CStr(i) + "BDate"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = True
                'BTime
                xcom = "DOP" + CStr(i) + "BTime"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = True
                'ADate
                xcom = "DOP" + CStr(i) + "ADate"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = True
                'ATime
                xcom = "DOP" + CStr(i) + "ATime"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = True
            Else
                'Image
                xcom = "DOP" + CStr(i) + "Image"
                ximg = Me.form1.FindControl(xcom)
                ximg.Visible = False
                'Step
                xcom = "DOP" + CStr(i) + "Step"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = False
                'Name
                xcom = "DOP" + CStr(i) + "Name"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = False
                'Count
                xcom = "DOP" + CStr(i) + "Count"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = False
                'Yotei
                xcom = "DOP" + CStr(i) + "Yotei"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = False
                'Jisseki
                xcom = "DOP" + CStr(i) + "Jisseki"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = False
                'Detail
                xcom = "LOP" + CStr(i) + "Detail"
                xlink = Me.form1.FindControl(xcom)
                xlink.Visible = False
                'Loading
                xcom = "LOP" + CStr(i) + "Loading"
                xlink = Me.form1.FindControl(xcom)
                xlink.Visible = False
                'BDate
                xcom = "DOP" + CStr(i) + "BDate"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = False
                'BTime
                xcom = "DOP" + CStr(i) + "BTime"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = False
                'ADate
                xcom = "DOP" + CStr(i) + "ADate"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = False
                'ATime
                xcom = "DOP" + CStr(i) + "ATime"
                xtxt = Me.form1.FindControl(xcom)
                xtxt.Visible = False
            End If
        Next
    End Sub

End Class
