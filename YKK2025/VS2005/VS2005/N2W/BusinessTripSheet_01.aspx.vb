Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
Imports System.IO
Imports System
Imports System.Web.UI
Imports System.Text
Imports System.Web.Configuration
Imports System.Data.Common
Imports System.Web.Security
Imports System.Web.UI.HtmlControls



Partial Class BusinessTripSheet_01
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim FieldName(60) As String     '各欄位
    Dim Attribute(60) As Integer    '各欄位屬性    
    Dim Top As String               '預設的控制項位置
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wStep As Integer            '工程代碼
    Dim wbFormSno As Integer        '連續起單-表單流水號
    Dim wbStep As Integer           '連續起單-工程代碼
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID
    Dim wAgentID As String          '被代理人ID
    Dim NowDateTime As String       '現在日期時間
    Dim wNextGate As String         '下一關人
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    Dim wUserIP As String = ""
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    Dim wUserName As String = ""            '姓名代理用
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    Dim OpenDir1 As String = ""
    Dim OpenDir2 As String = ""
    Dim AttachFileName As String
    Dim sourceDir, backupDir As String




    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'BASDate.Attributes("onclick") = "calendarPicker('DCheckDate');"

        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置
        SetControlPosition()    ' 設定控制項位置
        GetData()
        If Not Me.IsPostBack Then   '不是PostBack
            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()     '表單功能按鈕顯示
            NewAttachFilePath()
            If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
                ShowFormData()      '顯示表單資料
                UpdateTranFile()    '更新交易資料
                TopPosition()
                SetControlPosition()    ' 設定控制項位置

            End If
            SetPopupFunction()      '設定彈出視窗事件


        Else
            ShowSheetFunction()     '表單功能按鈕顯示
            ' fpObj.GetDeafultNextGate(DNextGate, DDivision.Text,Request.QueryString("pUserID")) ' 設定預設的簽核者
            ShowSheetField("Post")  '表單欄位顯示及欄位輸入檢查

            ShowMessage()           '上傳資料檢查及顯示訊息

            '上傳資料檢查及顯示訊息

        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     按鈕及RequestedField的Top位置
    '**
    '*****************************************************************
    Sub TopPosition()
        If wFormSno > 0 And wStep > 3 Then      '判斷是否[簽核]
            Top = 1000
            'Top = GridView1.Height.Value + 500

            Dim SQL As String = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(SQL)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0)("BEndTime").ToString < NowDateTime Then
                    If DDelay.Visible = True Then
                        Top = 1000
                        'Top = GridView1.Height.Value + 400
                    End If
                End If
            End If
        Else
            Top = 1000
            'Top = GridView1.Height.Value + 400
        End If


    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     '上傳資料檢查及顯示訊息
    '**
    '*****************************************************************
    Sub ShowMessage()
        Dim Message As String = ""
        If Message <> "" Then
            Message = "下列所設定的附加檔案將會遺失 (" & Message & ") " & ",請重新設定!"
            Response.Write(YKK.ShowMessage(Message))
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetControlPosition)
    '**    設定按鈕,說明,延遲,核定履歷的位置
    '**
    '*****************************************************************
    Sub SetControlPosition()
        If DDescSheet.Visible Then                                      ' 說明
            DDescSheet.Style("top") = Top - 150 & "px"
            DDecideDesc.Style("top") = Top - 150 + 6 & "px"
            Top = Top - 80
        End If

        If DDelay.Visible Then                                          ' 延遲說明
            DDelay.Style("top") = Top & "px"
            DReasonCode.Style("top") = Top + 9 & "px"
            DReason.Style("top") = Top + 6 & "px"
            DReasonDesc.Style("top") = Top + 39 & "px"
            Top += 161
        End If

        BSAVE.Style("top") = Top + 20 & "px"
        BNG1.Style("top") = Top + 20 & "px"
        BNG2.Style("top") = Top + 20 & "px"
        BOK.Style("top") = Top + 20 & "px"
        LMemo.Style("top") = Top + 50 & "px"

        Top = Top + 100
        If GridView2.Rows.Count > 0 Then                                ' 核定履歷
            DHistoryLabel.Style("top") = Top & "px"
            Top += 20
            GridView2.Style("top") = Top & "px"
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日時
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
        wStep = Request.QueryString("pStep")        '工程代碼
        wbFormSno = Request.QueryString("pFormSno")  '連續起單用流水號
        wbStep = Request.QueryString("pStep")        '連續起單用工程代碼
        wSeqNo = Request.QueryString("pSeqNo")      '序號
        wApplyID = Request.QueryString("pApplyID")  '申請者ID
        wAgentID = Request.QueryString("pAgentID")  '被代理人ID
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        wDecideCalendar = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))
        wUserIP = Request.ServerVariables("REMOTE_ADDR")
        '
        Response.Cookies("PGM").Value = "BusinessTripSheet_01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼





    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub UpdateTranFile()
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, Request.QueryString("pUserID"))
        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetFunction)
    '**     表單功能按鈕顯示
    '**
    '*****************************************************************
    Sub ShowSheetFunction()
        Dim wDelay As Integer = 0
        Dim sql As String = ""
        sql = "Select * From M_Flow "
        sql &= " Where Active = 1 "
        sql &= "   And FormNo =  '" & wFormNo & "'"
        sql &= "   And Step   =  '" & wStep & "'"
        Dim dtFlow As DataTable = uDataBase.GetDataTable(sql)
        If dtFlow.Rows.Count > 0 Then
            '電子簽章未使用
            If dtFlow.Rows(0)("SignImage") = 1 Then
            Else
            End If
            '附加檔案未使用(由FormField中設定)
            If dtFlow.Rows(0)("Attach") = 1 Then
            Else
            End If
            '儲存按鈕
            If dtFlow.Rows(0)("SaveFun") = 1 Then
                BSAVE.Visible = True
                BSAVE.Text = dtFlow.Rows(0)("SaveDesc")
                'BSAVE.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BSAVE.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
            Else
                BSAVE.Visible = False
            End If
            'NG-1按鈕
            If dtFlow.Rows(0)("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Text = dtFlow.Rows(0)("NGDesc1")
                'BNG1.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BNG1.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wNGSts1 = dtFlow.Rows(0)("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2按鈕
            If dtFlow.Rows(0)("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Text = dtFlow.Rows(0)("NGDesc2")
                'BNG2.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BNG2.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wNGSts2 = dtFlow.Rows(0)("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK按鈕
            If dtFlow.Rows(0)("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Text = dtFlow.Rows(0)("OKDesc")
                'BOK.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BOK.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wOKSts = dtFlow.Rows(0)("OKSts") + 1
                If wStep = 1 Or wStep = 500 Then
                    LMemo.Visible = True
                End If
            Else
                BOK.Visible = False
                LMemo.Visible = False
            End If
            '遲納管理
            If dtFlow.Rows(0)("Delay") = 1 Then
                wDelay = 1
            End If
        End If
        '
        If wFormSno > 0 And wStep > 3 Then    '判斷是否[簽核]
            sql = "Select * From T_WaitHandle "
            sql = sql & " Where Active = 1 "
            sql = sql & "   And FormNo =  '" & wFormNo & "'"
            sql = sql & "   And FormSno =  '" & CStr(wFormSno) & "'"
            sql = sql & "   And Step   =  '" & CStr(wStep) & "'"
            Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(sql)
            If dtWaitHandle.Rows.Count > 0 Then
                '說明設定
                Top = 1000
                Top = GridView1.Height.Value + 450

                DDescSheet.Visible = True
                DDecideDesc.Visible = True
                If dtWaitHandle.Rows(0)("FlowType") = 0 Then
                    DDecideDesc.BackColor = Color.LightGray
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "異常：需輸入說明")
                End If
                '遲納管理
                If wDelay = 1 Then
                    If dtWaitHandle.Rows(0)("BEndTime") < NowDateTime Then
                        DDelay.Visible = True   '延遲Sheet
                        DReasonCode.Visible = True     '延遲理由代碼
                        DReasonCode.BackColor = Color.GreenYellow
                        DReason.Visible = True         '延遲理由
                        DReason.BackColor = Color.GreenYellow
                        DReasonDesc.Visible = True     '延遲其他說明
                        ShowRequiredFieldValidator("DReasonCodeRqd", "DReasonCode", "異常：需輸入延遲理由")
                        ShowRequiredFieldValidator("DReasonRqd", "DReason", "異常：需輸入延遲理由")
                        Top = Top + 110
                    Else
                        DDelay.Visible = False  '延遲Sheet
                        DReasonCode.Visible = False     '延遲理由代碼
                        DReason.Visible = False         '延遲理由
                        DReasonDesc.Visible = False     '延遲其他說明
                    End If
                End If

            End If
        Else
            Top = 1000
            'Top = GridView1.Height.Value + 400

            'Sheet隱藏
            DDescSheet.Visible = False  '說明Sheet
            DDelay.Visible = False      '延遲Sheet
            '欄位隱藏
            DDecideDesc.Visible = False    '說明
            DReasonCode.Visible = False    '延遲理由代碼0
            DReason.Visible = False        '延遲理由
            DReasonDesc.Visible = False    '延遲其他說明
            DHistoryLabel.Visible = False  '核定履歷
        End If
        '按鈕及超連結設值

        If wStep > 1 Then
            Top = 1000
        Else
            Top = 600
        End If
        BSAVE.Style("top") = Top & "px"
        BNG1.Style("top") = Top & "px"
        BNG2.Style("top") = Top & "px"
        BOK.Style("top") = Top & "px"
        LMemo.Style("top") = Top + 30 & "px"

        DHistoryLabel.Style("top") = Top + 40 & "px"
        GridView2.Style("top") = Top + 40 + 16 & "px"
        '簽核,通知 移除控制設定

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
    '**     表單欄位顯示及欄位輸入檢查
    '**
    '*****************************************************************
    Sub ShowSheetField(ByVal pPost As String)
        '建置欄位及屬性陣列
        oCommon.GetFieldAttribute(wFormNo, wStep, FieldName, Attribute)
        '表單號碼,工程關卡號碼,欄位名,欄位屬性

        SetFieldAttribute(pPost)     '表單各欄位屬性及欄位輸入檢查等設定
    End Sub
    '*****************************************************************
    '**(SetFieldAttribute)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        Dim SQL As String
        Dim InputData1 As String
        Dim InputData2 As String
        Dim InputData3 As String


        InputData1 = D1.Text
        InputData2 = D2.Text
        InputData3 = DQC.Text



        'SQL = " Select a.* From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a ,m_users b"
        'SQL = SQL & "  where a.empname =b.username"
        'SQL = SQL & " and UserID = '" & wApplyID & "'"
        'SQL = SQL & "   And Active = '1' "
        'Dim DBUser As DataTable = uDataBase.GetDataTable(SQL)


        'DDivision1.Text = DBUser.Rows(0).Item("DepName")
        'DDivisionCode1.Text = DBUser.Rows(0).Item("DepID")
        'DEmpID1.Text = DBUser.Rows(0).Item("EmpID")
        'DEmpName1.Text = DBUser.Rows(0).Item("EmpName")
        'DJobTitle1.Text = DBUser.Rows(0).Item("JobTitle")
        Try

            SQL = " Select a.* From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a ,m_users b"
            SQL = SQL & "  where a.empname =b.username"
            SQL = SQL & " and UserID = '" & wApplyID & "'"
            SQL = SQL & "   And Active = '1' "
            Dim DBUser As DataTable = uDataBase.GetDataTable(SQL)



            If DBUser.Rows(0).Item("EmpName") = "陳怡君" Then
                SQL = " Select a.* From m_wfsemp a,m_users b"
                SQL = SQL & "  where a.empname =b.username and depid = hrwdivid "
                SQL = SQL & " and UserID = '" & wApplyID & "'"
                SQL = SQL & "   And Active = '1' "
                Dim DBUser1 As DataTable = uDataBase.GetDataTable(SQL)
                DDivision1.Text = DBUser1.Rows(0).Item("DepName")
                DDivisionCode1.Text = DBUser1.Rows(0).Item("DepID")
                DEmpID1.Text = DBUser1.Rows(0).Item("EmpID")
                DEmpName1.Text = DBUser1.Rows(0).Item("EmpName")
                DJobTitle1.Text = DBUser1.Rows(0).Item("JobTitle")
            Else
                DDivision1.Text = DBUser.Rows(0).Item("DepName")
                DDivisionCode1.Text = DBUser.Rows(0).Item("DepID")
                DEmpID1.Text = DBUser.Rows(0).Item("EmpID")
                DEmpName1.Text = DBUser.Rows(0).Item("EmpName")
                DJobTitle1.Text = DBUser.Rows(0).Item("JobTitle")

            End If

        Catch

            uJavaScript.PopMsg(Me, wApplyID + "-無此員工資料，請連絡電腦室")
        End Try


        'If InputData3 <> "" Then
        '    SQL = " select * from F_ComplaintOutSheet where no = '" & InputData3 & "'"
        '    Dim DBQC As DataTable = uDataBase.GetDataTable(SQL)
        '    DQCNO1.Text = DBQC.Rows(0).Item("NO")
        '    DQCNO.Text = DBQC.Rows(0).Item("QCNO")
        'End If



        'No
        Select Case FindFieldInf("No")
            Case 0  '顯示
                DNo.BackColor = Color.LightGray
                DNo.Visible = True
                DNo.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DNo.Visible = True
                DNo.BackColor = Color.GreenYellow
                DNo.ReadOnly = False
                ShowRequiredFieldValidator("DNoRqd", "DNo", "異常：需輸入Ｎｏ")

            Case 2  '修改
                DNo.Visible = True
                DNo.BackColor = Color.Yellow
                DNo.ReadOnly = False

            Case Else   '隱藏
                DNo.Visible = False
        End Select

        'Date
        Select Case FindFieldInf("Date")
            Case 0  '顯示
                DDate.BackColor = Color.LightGray
                DDate.Visible = True
                DDate.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DDate.Visible = True
                DDate.BackColor = Color.GreenYellow
                DDate.ReadOnly = False
                ShowRequiredFieldValidator("DDateRqd", "DDate", "異常：需輸入申請日期")

            Case 2  '修改
                DDate.Visible = True
                DDate.BackColor = Color.Yellow
                DDate.ReadOnly = False

            Case Else   '隱藏
                DDate.Visible = False

        End Select
        If pPost = "New" Then DDate.Text = Now.ToString("yyyy/MM/dd") '現在日時



        'Division1
        Select Case FindFieldInf("Division1")
            Case 0  '顯示
                DDivision1.BackColor = Color.LightGray
                DDivision1.Visible = True
                DDivision1.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DDivision1.Visible = True
                DDivision1.BackColor = Color.GreenYellow
                DDivision1.ReadOnly = False
                ShowRequiredFieldValidator("DDivision1Rqd", "DDivision1", "異常：需輸入申請者")

            Case 2  '修改
                DDivision1.Visible = True
                DDivision1.BackColor = Color.Yellow
                DDivision1.ReadOnly = False

            Case Else   '隱藏
                DDivision1.Visible = False

        End Select


        'Division2
        Select Case FindFieldInf("Division2")
            Case 0  '顯示
                DDivision2.BackColor = Color.LightGray
                DDivision2.Visible = True
                DDivision2.Attributes.Add("readonly", "true")
                CDivision2.Visible = False
                BDivision2.Visible = False
            Case 1  '修改+檢查
                DDivision2.Visible = True
                DDivision2.BackColor = Color.GreenYellow
                DDivision2.ReadOnly = True
                ShowRequiredFieldValidator("DDivision2Rqd", "DDivision2", "異常：需輸入受款者")
                CDivision2.Visible = True
                BDivision2.Visible = True

            Case 2  '修改
                DDivision2.Visible = True
                DDivision2.BackColor = Color.Yellow
                DDivision2.ReadOnly = False
                CDivision2.Visible = True
                BDivision2.Visible = True

            Case Else   '隱藏
                DDivision2.Visible = False
                CDivision2.Visible = False
                BDivision2.Visible = False
        End Select



        'Division3
        Select Case FindFieldInf("Division3")
            Case 0  '顯示
                DDivision3.BackColor = Color.LightGray
                DDivision3.Visible = True
                DDivision3.Attributes.Add("readonly", "true")
                DCDivision3.Visible = False
                BDivision3.Visible = False
            Case 1  '修改+檢查
                DDivision3.Visible = True
                DDivision3.BackColor = Color.GreenYellow
                DDivision3.ReadOnly = True
                ShowRequiredFieldValidator("DDivision3Rqd", "DDivision3", "異常：需輸入費用歸屬")
                DCDivision3.Visible = True
                BDivision3.Visible = True
            Case 2  '修改
                DDivision3.Visible = True
                DDivision3.BackColor = Color.Yellow
                DDivision3.ReadOnly = False
                DCDivision3.Visible = True
                BDivision3.Visible = True

            Case Else   '隱藏
                DDivision3.Visible = False
                DCDivision3.Visible = False
                BDivision3.Visible = False

        End Select



        'Division4
        Select Case FindFieldInf("Division4")
            Case 0  '顯示
                DDivision4.BackColor = Color.LightGray
                DDivision4.Visible = True
                DDivision4.Attributes.Add("readonly", "true")
                BDivision4.Visible = False
            Case 1  '修改+檢查
                DDivision4.Visible = True
                DDivision4.BackColor = Color.GreenYellow
                DDivision4.ReadOnly = True
                ShowRequiredFieldValidator("DDivision4Rqd", "DDivision4", "異常：需輸入出差者")
                BDivision4.Visible = True
            Case 2  '修改
                DDivision4.Visible = True
                DDivision4.BackColor = Color.Yellow
                DDivision4.ReadOnly = False
                BDivision4.Visible = True
            Case Else   '隱藏
                DDivision4.Visible = False
                BDivision4.Visible = False
        End Select




        'EmpName1
        Select Case FindFieldInf("EmpName1")
            Case 0  '顯示
                DEmpName1.BackColor = Color.LightGray
                DEmpName1.Visible = True
                DEmpName1.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DEmpName1.Visible = True
                DEmpName1.BackColor = Color.GreenYellow
                DEmpName1.ReadOnly = False
                ShowRequiredFieldValidator("DEmpName1Rqd", "DEmpName1", "異常：需輸入申請人")

            Case 2  '修改
                DEmpName1.Visible = True
                DEmpName1.BackColor = Color.Yellow
                DEmpName1.ReadOnly = False

            Case Else   '隱藏
                DEmpName1.Visible = False
        End Select

        'EmpName2
        Select Case FindFieldInf("EmpName2")
            Case 0  '顯示
                DEmpName2.BackColor = Color.LightGray
                DEmpName2.Visible = True
                DEmpName2.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DEmpName2.Visible = True
                DEmpName2.BackColor = Color.GreenYellow
                DEmpName2.ReadOnly = True
                ShowRequiredFieldValidator("DEmpName2Rqd", "DEmpName2", "異常：需輸入受款者")

            Case 2  '修改
                DEmpName2.Visible = True
                DEmpName2.BackColor = Color.Yellow
                DEmpName2.ReadOnly = False

            Case Else   '隱藏
                DEmpName2.Visible = False
        End Select

        'EmpName4
        Select Case FindFieldInf("EmpName4")
            Case 0  '顯示
                DEmpName4.BackColor = Color.LightGray
                DEmpName4.Visible = True
                DEmpName4.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DEmpName4.Visible = True
                DEmpName4.BackColor = Color.GreenYellow
                DEmpName4.ReadOnly = True
                ShowRequiredFieldValidator("DEmpName4Rqd", "DEmpName4", "異常：需輸入出差者")

            Case 2  '修改
                DEmpName4.Visible = True
                DEmpName4.BackColor = Color.Yellow
                DEmpName4.ReadOnly = False

            Case Else   '隱藏
                DEmpName4.Visible = False
        End Select


        'DivisionCode1
        Select Case FindFieldInf("DivisionCode1")
            Case 0  '顯示
                DDivisionCode1.BackColor = Color.LightGray
                DDivisionCode1.Visible = True
                DDivisionCode1.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DDivisionCode1.Visible = True
                DDivisionCode1.BackColor = Color.GreenYellow
                DDivisionCode1.ReadOnly = False
                ShowRequiredFieldValidator("DDivisionCode1Rqd", "DDivisionCode1", "異常：需輸入申請人")

            Case 2  '修改
                DDivisionCode1.Visible = True
                DDivisionCode1.BackColor = Color.Yellow
                DDivisionCode1.ReadOnly = False

            Case Else   '隱藏
                DDivisionCode1.Visible = False
        End Select



        'DivisionCode2
        Select Case FindFieldInf("DivisionCode2")
            Case 0  '顯示
                DDivisionCode2.BackColor = Color.LightGray
                DDivisionCode2.Visible = True
                DDivisionCode2.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DDivisionCode2.Visible = True
                DDivisionCode2.BackColor = Color.GreenYellow
                DDivisionCode2.ReadOnly = True
                ShowRequiredFieldValidator("DDivisionCode2Rqd", "DDivisionCode2", "異常：需輸入受款者")

            Case 2  '修改
                DDivisionCode2.Visible = True
                DDivisionCode2.BackColor = Color.Yellow
                DDivisionCode2.ReadOnly = False

            Case Else   '隱藏
                DDivisionCode2.Visible = False
        End Select


        'Divisioncode3
        Select Case FindFieldInf("DivisionCode3")
            Case 0  '顯示
                DDivisionCode3.BackColor = Color.LightGray
                DDivisionCode3.Visible = True
                DDivisionCode3.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DDivisionCode3.Visible = True
                DDivisionCode3.BackColor = Color.GreenYellow
                DDivisionCode3.ReadOnly = True
                ShowRequiredFieldValidator("DDivisioncode3Rqd", "DDivisioncode3", "異常：需輸入費用歸屬")

            Case 2  '修改
                DDivisionCode3.Visible = True
                DDivisionCode3.BackColor = Color.Yellow
                DDivisionCode3.ReadOnly = False

            Case Else   '隱藏
                DDivisionCode3.Visible = False
        End Select



        'DivisionCode4
        Select Case FindFieldInf("DivisionCode4")
            Case 0  '顯示
                DDivisionCode4.BackColor = Color.LightGray
                DDivisionCode4.Visible = True
                DDivisionCode4.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DDivisionCode4.Visible = True
                DDivisionCode4.BackColor = Color.GreenYellow
                DDivisionCode4.ReadOnly = True
                ShowRequiredFieldValidator("DDivisionCode4Rqd", "DDivisionCode4", "異常：需輸入出差者")

            Case 2  '修改
                DDivisionCode4.Visible = True
                DDivisionCode4.BackColor = Color.Yellow
                DDivisionCode4.ReadOnly = False

            Case Else   '隱藏
                DDivisionCode4.Visible = False
        End Select




        'EmpID1
        Select Case FindFieldInf("EmpID1")
            Case 0  '顯示
                DEmpID1.BackColor = Color.LightGray
                DEmpID1.Visible = True
                DEmpID1.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DEmpID1.Visible = True
                DEmpID1.BackColor = Color.GreenYellow
                DEmpID1.ReadOnly = False
                ShowRequiredFieldValidator("DEmpID1Rqd", "DEmpID1", "異常：需輸入申請人")

            Case 2  '修改
                DEmpID1.Visible = True
                DEmpID1.BackColor = Color.Yellow
                DEmpID1.ReadOnly = False

            Case Else   '隱藏
                DEmpID1.Visible = False
        End Select

        'EmpID2
        Select Case FindFieldInf("EmpID2")
            Case 0  '顯示
                DEmpID2.BackColor = Color.LightGray
                DEmpID2.Visible = True
                DEmpID2.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DEmpID2.Visible = True
                DEmpID2.BackColor = Color.GreenYellow
                DEmpID2.ReadOnly = True
                ShowRequiredFieldValidator("DEmpID2Rqd", "DEmpID2", "異常：需輸入出差者")

            Case 2  '修改
                DEmpID2.Visible = True
                DEmpID2.BackColor = Color.Yellow
                DEmpID2.ReadOnly = False

            Case Else   '隱藏
                DEmpID2.Visible = False
        End Select


        'EmpID4
        Select Case FindFieldInf("EmpID4")
            Case 0  '顯示
                DEmpID4.BackColor = Color.LightGray
                DEmpID4.Visible = True
                DEmpID4.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DEmpID4.Visible = True
                DEmpID4.BackColor = Color.GreenYellow
                DEmpID4.ReadOnly = True
                ShowRequiredFieldValidator("DEmpID4Rqd", "DEmpID4", "異常：需輸入出差者")

            Case 2  '修改
                DEmpID4.Visible = True
                DEmpID4.BackColor = Color.Yellow
                DEmpID4.ReadOnly = False

            Case Else   '隱藏
                DEmpID4.Visible = False
        End Select


        'Jobtitle1
        Select Case FindFieldInf("Jobtitle1")
            Case 0  '顯示
                DJobTitle1.BackColor = Color.LightGray
                DJobTitle1.Visible = True
                DJobTitle1.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DJobTitle1.Visible = True
                DJobTitle1.BackColor = Color.GreenYellow
                DJobTitle1.ReadOnly = False
                ShowRequiredFieldValidator("DJobtitle1Rqd", "DJobtitle1", "異常：需輸入申請人")

            Case 2  '修改
                DJobTitle1.Visible = True
                DJobTitle1.BackColor = Color.Yellow
                DJobTitle1.ReadOnly = False

            Case Else   '隱藏
                DJobTitle1.Visible = False
        End Select


        'Jobtitle2
        Select Case FindFieldInf("JobTitle2")
            Case 0  '顯示
                DJobTitle2.BackColor = Color.LightGray
                DJobTitle2.Visible = True
                DJobTitle2.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DJobTitle2.Visible = True
                DJobTitle2.BackColor = Color.GreenYellow
                DJobTitle2.ReadOnly = True
                ShowRequiredFieldValidator("DJobtitle2Rqd", "DJobtitle2", "異常：需輸入受款者")

            Case 2  '修改
                DJobTitle2.Visible = True
                DJobTitle2.BackColor = Color.Yellow
                DJobTitle2.ReadOnly = False

            Case Else   '隱藏
                DJobTitle2.Visible = False
        End Select

        'JobTitle4
        Select Case FindFieldInf("JobTitle4")
            Case 0  '顯示
                DJobTitle4.BackColor = Color.LightGray
                DJobTitle4.Visible = True
                DJobTitle4.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DJobTitle4.Visible = True
                DJobTitle4.BackColor = Color.GreenYellow
                DJobTitle4.ReadOnly = True
                ShowRequiredFieldValidator("DJobTitle4Rqd", "DJobTitle4", "異常：需輸入出差者")

            Case 2  '修改
                DJobTitle4.Visible = True
                DJobTitle4.BackColor = Color.Yellow
                DJobTitle4.ReadOnly = False

            Case Else   '隱藏
                DJobTitle4.Visible = False
        End Select




        'QCNO
        Select Case FindFieldInf("QCNO")
            Case 0  '顯示
                DQCNO.BackColor = Color.LightGray
                DQCNO.Visible = True
                DQCNO.Attributes.Add("readonly", "true")

                DQCNO1.BackColor = Color.LightGray
                DQCNO1.Visible = True
                DQCNO1.Attributes.Add("readonly", "true")
                BQCNO.Visible = False
            Case 1  '修改+檢查
                DQCNO.Visible = True
                DQCNO.BackColor = Color.GreenYellow
                DQCNO.ReadOnly = True
                ShowRequiredFieldValidator("DQCNORqd", "DQCNO", "異常：需輸入客訴編號")

                DQCNO1.Visible = True
                DQCNO1.BackColor = Color.GreenYellow
                DQCNO1.ReadOnly = True
                BQCNO.Visible = True
            Case 2  '修改
                DQCNO.Visible = True
                DQCNO.BackColor = Color.Yellow
                DQCNO.ReadOnly = True

                DQCNO1.Visible = True
                DQCNO1.BackColor = Color.Yellow
                DQCNO1.ReadOnly = True
                BQCNO.Visible = True

            Case Else   '隱藏
                DQCNO.Visible = False
                DQCNO1.Visible = False
                BQCNO.Visible = False
        End Select


        'PhoneNo
        Select Case FindFieldInf("PhoneNo")
            Case 0  '顯示
                DPhoneNo.BackColor = Color.LightGray
                DPhoneNo.Visible = True
                DPhoneNo.Attributes.Add("readonly", "true")
                DChkPhone.Enabled = False

            Case 1  '修改+檢查
                DPhoneNo.Visible = True
                DPhoneNo.BackColor = Color.GreenYellow
                DPhoneNo.ReadOnly = True
                ShowRequiredFieldValidator("DPhoneNoRqd", "DPhoneNo", "異常：需輸入門號")
                DChkPhone.Enabled = True

            Case 2  '修改
                DPhoneNo.Visible = True
                DPhoneNo.BackColor = Color.Yellow
                DPhoneNo.ReadOnly = False
                DChkPhone.Enabled = True
            Case Else   '隱藏
                DPhoneNo.Visible = False
                DChkPhone.Enabled = False

        End Select



        'PassDate
        Select Case FindFieldInf("PassDate")
            Case 0  '顯示
                DPassDate.BackColor = Color.LightGray
                DPassDate.Visible = True
                DPassDate.Attributes.Add("readonly", "true")
                BPassDate.Visible = False

            Case 1  '修改+檢查
                DPassDate.Visible = True
                DPassDate.BackColor = Color.GreenYellow
                ' DPassDate.ReadOnly = True
                ShowRequiredFieldValidator("DPassDateRqd", "DPassDate", "異常：需輸入設照有效日")
                BPassDate.Visible = True

            Case 2  '修改
                DPassDate.Visible = True
                DPassDate.BackColor = Color.Yellow
                'DPassDate.ReadOnly = False
                BPassDate.Visible = True
            Case Else   '隱藏
                DPassDate.Visible = False

                BPassDate.Visible = False

        End Select
        If pPost = "New" Then DPassDate.Text = Now.ToString("yyyy/MM/dd") '現在日時



        'InsurancePay
        Select Case FindFieldInf("InsurancePay")
            Case 0  '顯示

                DchkInsurance.Enabled = False


            Case 1  '修改+檢查

                ' ShowRequiredFieldValidator("DInsurancePayRqd", "DInsurancePay", "異常：海外需打勾")
                DchkInsurance.Enabled = True

            Case 2  '修改


                DchkInsurance.Enabled = True

            Case Else   '隱藏

                DchkInsurance.Enabled = False

        End Select




        'chkVisa
        Select Case FindFieldInf("chkVisa")
            Case 0  '顯示

                DchkVisa.Enabled = False

            Case 1  '修改+檢查

                '  ShowRequiredFieldValidator("DchkVisaRqd", "DchkVisa", "異常：需輸入出差地簽証")
                DchkVisa.Enabled = True
            Case 2  '修改
                DchkVisa.Enabled = True
            Case Else   '隱藏
                DchkVisa.Enabled = False

        End Select




        'AirTickes
        Select Case FindFieldInf("AirTickets")
            Case 0  '顯示
                DAirTickets.BackColor = Color.LightGray
                DAirTickets.Visible = True
                DAirTickets.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DAirTickets.Visible = True
                DAirTickets.BackColor = Color.GreenYellow

                ShowRequiredFieldValidator("DAirTicketsRqd", "DAirTickets", "異常：需輸入機票總額")

            Case 2  '修改
                DAirTickets.Visible = True
                DAirTickets.BackColor = Color.Yellow

            Case Else   '隱藏
                DAirTickets.Visible = False
        End Select



        'Object
        Select Case FindFieldInf("Object")
            Case 0  '顯示
                DObject.BackColor = Color.LightGray
                DObject.Visible = True
                DObject.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DObject.Visible = True
                DObject.BackColor = Color.GreenYellow

                ShowRequiredFieldValidator("DObjectRqd", "DObject", "異常：需輸入目的")

            Case 2  '修改
                DObject.Visible = True
                DObject.BackColor = Color.Yellow

            Case Else   '隱藏
                DObject.Visible = False
        End Select



        'Location
        Select Case FindFieldInf("Location")
            Case 0  '顯示
                DLocation.BackColor = Color.LightGray
                DLocation.Visible = True
                DLocation.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DLocation.Visible = True
                DLocation.BackColor = Color.GreenYellow

                ShowRequiredFieldValidator("DLocationRqd", "DLocation", "異常：需輸入地區/訪問")

            Case 2  '修改
                DLocation.Visible = True
                DLocation.BackColor = Color.Yellow

            Case Else   '隱藏
                DLocation.Visible = False
        End Select




        'Remark
        Select Case FindFieldInf("Remark")
            Case 0  '顯示
                DRemark.BackColor = Color.LightGray
                DRemark.Visible = True
                DRemark.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DRemark.Visible = True
                DRemark.BackColor = Color.GreenYellow

                ShowRequiredFieldValidator("DRemarkRqd", "DRemark", "異常：需同行者/備註/其它")

            Case 2  '修改
                DRemark.Visible = True
                DRemark.BackColor = Color.Yellow

            Case Else   '隱藏
                DRemark.Visible = False
        End Select




        'COMDATE
        Select Case FindFieldInf("SDate")
            Case 0  '顯示
                DSDate.BackColor = Color.LightGray
                DSDate.Visible = True
                DSDate.Attributes.Add("readonly", "true")
                BSDate.Visible = False
            Case 1  '修改+檢查
                DSDate.Visible = True
                DSDate.BackColor = Color.GreenYellow
                DSDate.ReadOnly = False
                BSDate.Visible = True
                ShowRequiredFieldValidator("DSDateRqd", "DSDate", "異常：需輸入預定日程")
            Case 2  '修改
                DSDate.Visible = True
                DSDate.BackColor = Color.Yellow
                DSDate.ReadOnly = False
                BSDate.Visible = True
            Case Else   '隱藏
                DSDate.Visible = False
                BSDate.Visible = False
        End Select



        'COMDATE
        Select Case FindFieldInf("EDate")
            Case 0  '顯示
                DEDate.BackColor = Color.LightGray
                DEDate.Visible = True
                DEDate.Attributes.Add("readonly", "true")
                BEDate.Visible = False
            Case 1  '修改+檢查
                DEDate.Visible = True
                DEDate.BackColor = Color.GreenYellow
                DEDate.ReadOnly = False
                BEDate.Visible = True
                ShowRequiredFieldValidator("DEdateRqd", "DEdate", "異常：需輸入預定日程")
            Case 2  '修改
                DEDate.Visible = True
                DEDate.BackColor = Color.Yellow
                DEDate.ReadOnly = False
                BEDate.Visible = True
            Case Else   '隱藏
                DEDate.Visible = False
                BEDate.Visible = False
        End Select



        'AddFly
        Select Case FindFieldInf("AddFly")
            Case 0  '顯示
                BAddFly.Visible = False
                PAddFly.Visible = True
            Case 1  '修改+檢查
                BAddFly.Visible = True
                PAddFly.Visible = True

            Case 2  '修改
                BAddFly.Visible = True
                PAddFly.Visible = True

            Case Else   '隱藏
                BAddFly.Visible = False
                PAddFly.Visible = False

        End Select

        'Add
        Select Case FindFieldInf("Add")
            Case 0  '顯示
                BAdd.Visible = False
                PAdd.Visible = True
            Case 1  '修改+檢查
                BAdd.Visible = True
                PAdd.Visible = True

            Case 2  '修改
                BAdd.Visible = True
                PAdd.Visible = True

            Case Else   '隱藏
                BAdd.Visible = False
                PAdd.Visible = False

        End Select




        Select Case FindFieldInf("ATTACHFILE1")
            Case 0  '顯示            

                DAttachfile1.Visible = True



            Case 1  '修改+檢查
                DAttachfile1.Visible = True

            Case 2  '修改
                DAttachfile1.Visible = True


            Case Else   '隱藏
                DAttachfile1.Visible = False

        End Select



    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()



        Dim SQL As String
        SQL = "Select * From F_BusinessTripSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'" '
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dtData.Rows(0).Item("No")                         'No
            DDate.Text = dtData.Rows(0).Item("Date")                         'No 
            DDivision1.Text = dtData.Rows(0).Item("Division1")
            DDivision2.Text = dtData.Rows(0).Item("Division2")
            DDivision3.Text = dtData.Rows(0).Item("Division3")
            DDivision4.Text = dtData.Rows(0).Item("Division4")
            DDivisionCode1.Text = dtData.Rows(0).Item("Divisioncode1")
            DDivisionCode2.Text = dtData.Rows(0).Item("Divisioncode2")
            DDivisionCode3.Text = dtData.Rows(0).Item("Divisioncode3")
            DDivisionCode4.Text = dtData.Rows(0).Item("Divisioncode4")
            DEmpID1.Text = dtData.Rows(0).Item("EmpID1")
            DEmpID2.Text = dtData.Rows(0).Item("EmpID2")
            DEmpID4.Text = dtData.Rows(0).Item("EmpID4")
            DEmpName1.Text = dtData.Rows(0).Item("EmpName1")
            DEmpName2.Text = dtData.Rows(0).Item("EmpName2")
            DEmpName4.Text = dtData.Rows(0).Item("EmpName4")
            DJobTitle1.Text = dtData.Rows(0).Item("Jobtitle1")
            DJobTitle2.Text = dtData.Rows(0).Item("Jobtitle2")
            DJobTitle4.Text = dtData.Rows(0).Item("Jobtitle4")
            DPassDate.Text = dtData.Rows(0).Item("PassDate")
            DQCNO.Text = dtData.Rows(0).Item("QCNO")
            DQCNO1.Text = dtData.Rows(0).Item("QCNO1")
            If dtData.Rows(0).Item("ChkPhone") = 1 Then
                DChkPhone.Checked = True
            Else
                DChkPhone.Checked = False
            End If

            DPhoneNo.Text = dtData.Rows(0).Item("PhoneNo")

            If dtData.Rows(0).Item("chkInsurance") = 1 Then
                DchkInsurance.Checked = True
            Else
                DchkInsurance.Checked = False
            End If
            DInsurancePay.Text = dtData.Rows(0).Item("InsurancePay")

            If dtData.Rows(0).Item("chkVisa") = 1 Then
                DchkVisa.Checked = True
            Else
                DchkVisa.Checked = False
            End If
            DAirTickets.Text = dtData.Rows(0).Item("AirTickets")
            DObject.Text = dtData.Rows(0).Item("Object")
            DLocation.Text = dtData.Rows(0).Item("Location")
            DSDate.Text = dtData.Rows(0).Item("SDate")
            DEDate.Text = dtData.Rows(0).Item("EDate")
            DRemark.Text = dtData.Rows(0).Item("Remark")


        
            ' 客訴連結
            SQL = " select * from F_ComplaintOutSheet"
            SQL = SQL + "  where no ='" & dtData.Rows(0).Item("QCNO1") & "'"
            Dim Ddata1 As DataTable = uDataBase.GetDataTable(SQL)
            If Ddata1.Rows.Count > 0 Then
                DQCNO.Text = Ddata1.Rows(0).Item("QCNO")
                LQCNo.Visible = True
                LQCNo.Text = "Link"
                '  LQCNo.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutSheet_02.aspx?&pFormno=003109&pFormsno=" + Trim(Ddata1.Rows(0).Item("formsno"))
                LQCNo.NavigateUrl = "ComplaintOutSheet_02.aspx?&pFormno=003109&pFormsno=" + Trim(Ddata1.Rows(0).Item("formsno"))
            End If



            '明細
            SQL = " Select  Unique_ID,No, Type, Appoint,convert(char(10),SDate,111)+' '+stime1+':'+stime2 as SDate,"
            SQL = SQL + " convert(char(10),EDate,111)+' '+etime1+':'+etime2 as EDate,"
            SQL = SQL + " Days,Money,Currency, FlyInf,SFly,EFly,HotelInf, Remark  from  F_BusinessTripSheetDT where 1=1 "
            SQL = SQL + " and NO = '" + DNo.Text + "'"
            SQL = SQL + " order by unique_id "
            GridView1.DataSource = uDataBase.GetDataTable(SQL)
            GridView1.DataBind()
            If GridView1.Rows.Count > 0 Then
                DHaveData.Text = "1"
            Else
                DHaveData.Text = "0"
            End If

            '總務備註


            SQL = " Select Unique_ID,No,FlyType,Pay, Money1,Currency1,Remark1,convert(char(10),SDate,112)Sdate,SNo "
            SQL = SQL + "   from  F_BusinessFlyDT  where 1=1 "
            SQL = SQL + " and NO = '" + DNo.Text + "'"
            SQL = SQL + " order by unique_id "

            GridView3.DataSource = uDataBase.GetDataTable(SQL)
            GridView3.DataBind()


 

            '主檔資料
            SQL = " select  data,replace(data,'/','\')data1  from M_referp"
            SQL = SQL + " where cat = '3114'"
            SQL = SQL + " and dkey ='AttachfilePath1'"
            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
            Dim OpenDir As String = ""
            If DBAdapter3.Rows.Count > 0 Then
                OpenDir = "file://" + DBAdapter3.Rows(0).Item("Data") + DNo.Text + "/AttachFile"
            End If
            Dim OpenDir2 As String = ""
            OpenDir2 = "file://" + DBAdapter3.Rows(0).Item("Data") + DNo.Text + "/AttachFile"
            DAttachfile1.Attributes.Add("onclick", "window.open('" + OpenDir + "','_blank');return false;")


   
            '交易資料
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"
            Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(SQL)
            If dtWaitHandle.Rows.Count > 0 Then
                DDecideDesc.Text = dtWaitHandle.Rows(0)("DecideDesc")           '說明

                If dtWaitHandle.Rows(0)("BEndTime") < NowDateTime Then
                    If DDelay.Visible = True Then
                        SetFieldData("ReasonCode", dtWaitHandle.Rows(0)("ReasonCode"))    '延遲理由代碼
                        If dtWaitHandle.Rows(0)("ReasonCode") = "" Then
                            SetFieldData("Reason", DReasonCode.SelectedValue)    '延遲理由
                            DReasonDesc.Text = ""   '延遲其他說明
                        Else
                            DReason.Text = dtWaitHandle.Rows(0)("Reason")  '延遲理由
                            DReasonDesc.Text = dtWaitHandle.Rows(0)("ReasonDesc")     '延遲其他說明
                        End If
                    End If
                End If
            End If


            If DDecideDesc.Text = "" Then
                DDecideDesc.Text = "OK."
            End If

            '核定履歷資料
            SQL = "SELECT "
            SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
            SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
            SQL = SQL + "AEndTimeDesc As Description, "
            SQL = SQL + "URL "
            SQL = SQL + "FROM V_WaitHandle_01 "
            SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
            SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
            SQL = SQL + "Order by CreateTime Desc, Step Desc, SeqNo Desc "
            GridView2.DataSource = uDataBase.GetDataTable(SQL)
            GridView2.DataBind()
        End If
    End Sub
    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim sql As String = ""
        Dim idx As Integer = FindFieldInf(pFieldName)
        Dim i As Integer
        ''擔當者及部門 

        'sql = "Select Divname,Username From M_Users "
        'sql = sql & " Where UserID = '" & wApplyID & "'"
        'sql = sql & "   And Active = '1' "
        'Dim DBUser As DataTable = uDataBase.GetDataTable(sql)


        'DDepName.Text = DBUser.Rows(0).Item("Divname")
        'DAppID.Text = DBUser.Rows(0).Item("Username")









        '延遲理由代碼
        If pFieldName = "ReasonCode" Then
            DReasonCode.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DReasonCode.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='014' Order by DKey, Data "
                Dim dtReasonCode As DataTable = uDataBase.GetDataTable(sql)
                For i = 0 To dtReasonCode.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReasonCode.Rows(i)("DKey")
                    ListItem1.Value = dtReasonCode.Rows(i)("DKey")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DReasonCode.Items.Add(ListItem1)
                Next
            End If
        End If
        '延遲理由
        If pFieldName = "Reason" Then
            sql = "Select * From M_Referp Where Cat='014' and DKey= '" & DReasonCode.SelectedValue & "' Order by Data "
            Dim dtReason As DataTable = uDataBase.GetDataTable(sql)
            For i = 0 To dtReason.Rows.Count - 1
                DReason.Text = dtReason.Rows(i)("Data")
            Next
        End If
    End Sub

    '*****************************************************************
    '**(ShowRequiredFieldValidator)
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator

        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.Text = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        If wStep > 1 Then
            Top = 1000
        Else
            Top = 550
        End If
        rqdVal.Style.Add("Top", Top + 50 & "px")
        rqdVal.Style.Add("Left", "8px")
        rqdVal.Style.Add("Height", "20px")
        rqdVal.Style.Add("Width", "250px")
        rqdVal.Style.Add("POSITION", "absolute")
        Me.Form.Controls.Add(rqdVal)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(FindFieldInf)
    '**     尋找表單欄位屬性
    '**
    '*****************************************************************
    Function FindFieldInf(ByVal pFieldName As String) As Integer
        Dim Run As Boolean
        Dim i As Integer
        Run = True
        FindFieldInf = 9
        While i <= 60 And Run
            If FieldName(i) = pFieldName Then
                FindFieldInf = Attribute(i)
                Run = False
            End If
            i = i + 1
        End While
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()


        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"
        '
        BPassDate.Attributes("onclick") = "calendarPicker('Form1.DPassDate');"


        If wFormSno <> 0 Then
            DTNo.Text = DNo.Text
        Else
            DTNo.Text = Now.ToString("yyyyMMddHHmmss") '虛擬單號
        End If
        BDivision2.Attributes.Add("onClick", "GetEMP1()") '受款人
        BDivision4.Attributes.Add("onClick", "GetEMP2()") '代理人
        '
        BDivision3.Attributes.Add("onClick", "GetDivision()")

        BQCNO.Attributes.Add("onClick", "GetQCNo()") '取客訴編號

        ' BAdd.Attributes.Add("onClick", "AddItem1(" + DTNo.Text + "）")
        'BAddFly.Attributes.Add("onClick", "AddFly(" + DTNo.Text + "）")


        ' BAdd.Attributes("onclick") = "window.open('AddItemListBC.aspx?pUserID=" & Request.QueryString("pUserID") & "&pDTNo=" & DTNo.Text & "','','height=600,width=1000,menubar=no,location=no');"
        BAdd.Attributes("onclick") = "window.open('AddItemListBC.aspx?pUserID=" & Request.QueryString("pUserID") & "&pDTNo=" & DTNo.Text & "&pwStep=" & wStep & "','newwindow','height=600,width=1200,menubar=no,location=no');"

        BAddFly.Attributes("onclick") = "window.open('AddItemListFLY.aspx?pUserID=" & Request.QueryString("pUserID") & "&pDTNo=" & DTNo.Text & "&pwStep=" & wStep & "','','height=600,width=1200,menubar=no,location=no');"




    End Sub
    '##-1
    '##AgentApprovProc-Start
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AgentApprovProc)
    '**     流程控制
    '**        pFun=OK, NG1, NG2, SAVE  
    '**        pAction=0:OK, 1:NG1, 2:NG2, 3:Save   下一關卡時使用 
    '**        pSts=0:未處理, 1:OK, 2:NG1, 3:NG2, 4:已閱讀, 5:被抽單  更新T_Waithandle狀態
    '**     
    '*****************************************************************
    Sub AgentApprovProc(ByVal pFun As String, ByVal pAction As Integer, ByVal pSts As String)
        Dim ErrCode As Integer = 0
        '------------------------------------------------
        '處理button說明
        'StsDes
        Dim wStsDes As String = ""
        If pSts = "1" Then wStsDes = BOK.Text
        If pSts = "2" Then wStsDes = BNG1.Text
        If pSts = "3" Then wStsDes = BNG2.Text
        '--
        '------------------------------------------------
        '指定下工程簽核者
        'AllocteID
        Dim wAllocateID As String = ""
        '--
        '------------------------------------------------
        '主表單
        'TabelName
        Dim wTableName As String = "F_BusinessTripSheet"
        '--
        '------------------------------------------------
        '代理簽核
        'RunAgentApprov
        ErrCode = oCommon.RunAgentApprov(pFun, _
                                            pAction, _
                                            pSts, _
                                            wStsDes, _
                                            wFormNo, _
                                            wFormSno, _
                                            wStep, _
                                            wSeqNo, _
                                            wDecideCalendar, _
                                            Request.QueryString("pUserID"), _
                                            wApplyID, _
                                            wAgentID, _
                                            wAllocateID, _
                                            uCommon.ReplaceString(DDecideDesc.Text), _
                                            wTableName, _
                                            wUserIP)
        '--
        '------------------------------------------------
        '表單資料
        If ErrCode = 0 Then
            ModifyData("AGENTAPPROVE", "0")
            AddCommissionNo(wFormNo, wFormSno)
            '
            Dim URL As String = "http://10.245.1.10/WorkFlow/WaitHandle.aspx?pUserID=" + Request.QueryString("pUserID")
            Response.Redirect(URL)
        Else
            EnabledButton()   '起動Button運作
            uJavaScript.PopMsg(Me, "代理簽核異常,請確認或連絡系統人員!")
        End If

    End Sub
    '##AgentApprovProc-End
    '##

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(FlowControl)
    '**     流程控制
    '**        pFun=OK, NG1, NG2, SAVE  
    '**        pAction=0:OK, 1:NG1, 2:NG2, 3:Save   下一關卡時使用 
    '**        pSts=0:未處理, 1:OK, 2:NG1, 3:NG2, 4:已閱讀, 5:被抽單  更新T_Waithandle狀態
    '**     
    '*****************************************************************
    Sub FlowControl(ByVal pFun As String, ByVal pAction As Integer, ByVal pSts As String)
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim Message As String = ""
        Dim wQCLT As Integer = 0 'QC-L/T
        Dim Run As Boolean = True           '是否執行
        Dim RepeatRun As Boolean = False    '是否重覆執行
        Dim MultiJob As Integer = 0         '多人核定
        Dim wLevel As String = ""           '難易度
        '
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
            Dim NewFormSno As Integer = wFormSno    '表單流水號
            Dim pRunNextStep As Integer = 0         '是否執行計算下一關(會簽)

            '取得表單流水號或更新交易資料
            If ErrCode = 0 Then
                If wFormSno = 0 And wStep < 3 Then    '判斷是否起單
                    '取得表單流水號
                    RtnCode = oCommon.GetFormSeqNo(wFormNo, NewFormSno) '表單號碼, 表單流水號
                    If RtnCode <> 0 Then
                        ErrCode = 9110
                    Else
                        '申請流程資料建置(表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者)
                        oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                        '設定委託No
                        DNo.Text = SetNo(NewFormSno)
                    End If
                    pRunNextStep = 1
                Else
                    If RepeatRun = False Then   '不是通知的重覆執行
                        '更新交易資料
                        ModifyTranData(pFun, pSts)
                        '流程資料結束(表單號碼,表單流水號,工程關卡號碼,行事曆,簽核者, 流程結束否(會簽))
                        RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, Request.QueryString("pUserID"), pRunNextStep)

                        If RtnCode <> 0 Then
                            ErrCode = 9120
                        End If
                    Else
                        pRunNextStep = 1    '是通知的重覆執行
                    End If
                End If
            End If

            '取得下一關
            If ErrCode = 0 And pRunNextStep = 1 Then
                '指定簽核者User ID
                Dim wAllocateID As String = ""
                '取得簽核者
                '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
                '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 



                RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, wAgentID, wAllocateID, MultiJob, _
                                                  pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)






                If RtnCode <> 0 Then ErrCode = 9130
                If pCount = 0 And pNextStep <> 999 Then ErrCode = 9131
                If ErrCode = 0 Then pAction = 0
            End If

            '建置流程資料
            If ErrCode = 0 And pRunNextStep = 1 Then
                If pNextStep <> 999 Then
                    wNextGate = ""
                    For i = 1 To pCount
                        '取得下一關人員(訊息時使用)
                        If wNextGate = "" Then
                            wNextGate = pNextGate(i)
                        Else
                            wNextGate = "," & pNextGate(i)
                        End If
                        '取得核定者-群組行是曆
                        wNextGateCalendar = oCommon.GetCalendarGroup(pNextGate(i))

                        '取得工程負荷最後日時(核定者, 表單號碼, 工程號碼, 類別(0:通知,1:核定), 開始日時, 最後日時, 件數)
                        oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)

                        '取得預定開始,完成日程計算(表單號碼,工程號碼,難易度,QC-L/T,現在時間, 預定開始日時, 預定完成日時, 行事曆)
                        oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wNextGateCalendar)

                        '建置流程資料(表單號碼,表單流水號,工程關卡號碼,序號,申請者ID,行事曆,建置者, 簽核者, 被代理者, 申請者, 預定開始日, 預定完成日, 重要性)
                        RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)

                        If RtnCode <> 0 Then
                            ErrCode = 9150
                            Exit For
                        End If
                    Next i
                Else
                    '流程結束(表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者)
                    RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)

                    If RtnCode <> 0 Then ErrCode = 9160
                End If
            End If
            '當工程日程調整
            If ErrCode = 0 Then
                If RepeatRun = False Then
                    '簽核者,表單號碼,表單流水號,工程關卡號碼,序號,現在日時,難易度,行事曆
                    RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDecideCalendar)
                End If
            End If
            '儲存表單資料
            If ErrCode = 0 Then
                If wFormSno = 0 And wStep < 3 Then    '判斷是否起單
                    If NewFormSno <> 0 Then
                        AppendData(pFun, NewFormSno)  'Insert
                        AddCommissionNo(wFormNo, NewFormSno)
                    End If
                Else
                    If pNextStep = 999 Then     '工程結束嗎?
                        If pFun = "OK" Then ModifyData(pFun, CStr(wOKSts)) '更新表單資料
                        If pFun = "NG1" Then ModifyData(pFun, CStr(wNGSts1))
                        If pFun = "NG2" Then ModifyData(pFun, CStr(wNGSts2))
                    Else
                        ModifyData(pFun, "0")         '更新表單資料 Sts=0(未結)
                    End If
                    AddCommissionNo(wFormNo, wFormSno)
                End If
                '傳送郵件
                If pNextStep <> 999 Then
                    For i = 1 To pCount
                        oCommon.Send(Request.QueryString("pUserID"), pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "FLOW")
                        '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                    Next i
                Else
                    oCommon.Send(Request.QueryString("pUserID"), wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "END")
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

                wStep = pNextStep     '下一工程關卡號碼
                wFormSno = NewFormSno '下一工程表單流水號
            Else
                EnabledButton()   '起動Button運作
                If ErrCode = 9110 Then Message = "取得表單流水號計算異常,請確認或連絡系統人員!"
                If ErrCode = 9120 Then Message = "流程資料更新異常,請確認或連絡系統人員!"
                If ErrCode = 9130 Then Message = "下工程計算異常,請確認或連絡系統人員!"
                If ErrCode = 9131 Then Message = "無下工程管理人,請確認或連絡系統人員!"
                If ErrCode = 9140 Then Message = "工程預定開始及完成日計算異常,請確認或連絡系統人員!"
                If ErrCode = 9150 Then Message = "下一工程資料建置異常,請確認或連絡系統人員!"
                If ErrCode = 9160 Then Message = "工程結束資料建置異常(999),請確認或連絡系統人員!"
                uJavaScript.PopMsg(Me, Message)
            End If      '儲存表單ErrCode=0
        End While       '重覆執行

        If ErrCode = 0 Then
            '--郵件傳送---------
            'oCommon.SendMail()
            '--郵件傳送---------
            oCommon.SendMail()
            CheckTempFile()
            '
            Dim URL As String = ""
            If wStep = 10 Then
                'ADD-START BY JOY 2020/12/27

                'ADD-END
                '
                URL = "BusinessTripSheet_03.aspx?pformno=" & wFormNo & "&pformsno=" & wFormSno
                'URL = "http://10.245.1.10/N2W/BusinessTripSheet_03.aspx?pformno=" & wFormNo & "&pformsno=" & wFormSno

                '——原視窗保留，另外新增一個新頁面;
                'Response.Write("<script>window.open('" + URL + "','_blank');</script>")

                'URL = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                '                                            "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID

                Response.Redirect(URL)


            Else

                URL = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                               "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID

                Response.Redirect(URL)
            End If

        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(CheckTempFile)
    '**     同步化 F_BusinessTripSheetDT & F_BusinessTripSheetDTTemp
    '**
    '*****************************************************************
    Sub CheckTempFile()
        Dim sql As String = "Exec sp_FundingReplication "
        uDataBase.ExecuteNonQuery(sql)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     新增表單資料
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)

        Dim sql As String = ""
        sql = " Insert into F_BusinessTripSheet (Sts, CompletedTime, FormNo, FormSno,Date,"
        sql &= " NO,Division1,DivisionCode1,EmpID1,EmpName1,JobTitle1,"
        sql &= " Division2,DivisionCode2,EmpID2,EmpName2,JobTitle2,"
        sql &= " Division3,DivisionCode3,"
        sql &= " Division4,DivisionCode4,EmpID4,EmpName4,JobTitle4,"
        sql &= " PassDate,PreMoney,QCNO,QCNO1,chkPhone,PhoneNo,chkInsurance,InsurancePay,chkVisa,AirTickets,"
        sql &= " Object,Location,Sdate,Edate,Remark,CloseAccNo,"
        sql &= " CreateUser, CreateTime, ModifyUser, ModifyTime) "
        sql &= "VALUES( "
        sql &= " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        sql &= " '" + NowDateTime + "', "        '結案日
        sql &= " '003114', "                     '表單代號
        sql &= " '" + CStr(NewFormSno) + "', "   '表單流水號
        sql = sql + " N'" + DDate.Text + "', "                '日期

        sql = sql + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(Trim(DDivision1.Text)) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DDivisionCode1.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DEmpID1.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DEmpName1.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DJobTitle1.Text) + "', "   'NO  1

        sql = sql + " N'" + YKK.ReplaceString(Trim(DDivision2.Text)) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DDivisionCode2.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DEmpID2.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DEmpName2.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DJobTitle2.Text) + "', "   'NO  1

        sql = sql + " N'" + YKK.ReplaceString(Trim(DDivision3.Text)) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DDivisionCode3.Text) + "', "   'NO  1

        sql = sql + " N'" + YKK.ReplaceString(Trim(DDivision4.Text)) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DDivisionCode4.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DEmpID4.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DEmpName4.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DJobTitle4.Text) + "', "   'NO  1

        sql = sql + " N'" + YKK.ReplaceString(DPassDate.Text) + "', "   'NO  1
        sql = sql + " '', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DQCNO.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DQCNO1.Text) + "', "   'NO  1
        If DChkPhone.Checked = True Then
            sql = sql + " 1, "   'NO  1
        Else
            sql = sql + " 0, "   'NO  1
        End If
        sql = sql + " N'" + YKK.ReplaceString(DPhoneNo.Text) + "', "   'NO  1
        If DchkInsurance.Checked = True Then
            sql = sql + " 1, "   'NO  1
        Else
            sql = sql + " 0, "   'NO  1
        End If
        sql = sql + " N'" + YKK.ReplaceString(DInsurancePay.Text) + "', "   'NO  1
        If DchkVisa.Checked = True Then
            sql = sql + " 1, "   'NO  1
        Else
            sql = sql + " 0, "   'NO  1
        End If
        sql = sql + " N'" + YKK.ReplaceString(DAirTickets.Text) + "', "   'NO  1

        sql = sql + " N'" + YKK.ReplaceString(DObject.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DLocation.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DSDate.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DEDate.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DRemark.Text) + "', "   'NO  1
        sql = sql + " '', "   'CloseAccNo

        sql = sql + " N'" + Request.QueryString("pUserID") & "' ,"
        sql = sql + " N'" + NowDateTime & "' ,"
        sql = sql + " N'" + Request.QueryString("pUserID") & "' ,"
        sql = sql + " N'" + NowDateTime & "' )"
        uDataBase.ExecuteNonQuery(sql)





        ''存入明細


        sql = " insert into F_BusinessTripSheetdt (No,TypeNo,Type,Appoint,Sdate,Edate,STime1,STime2,ETime1,ETime2,Days,Money,Currency,Pay, FlyInf,SFly,EFly,HotelInf,Remark,CreateUser, CreateTime, ModifyUser, ModifyTime) "
        sql = sql + " select N'" + YKK.ReplaceString(DNo.Text) + "',"
        sql = sql + " TypeNo,Type,Appoint,Sdate,Edate,STime1,STime2,ETime1,ETime2,Days,Money,Currency,Pay, FlyInf,SFly,EFly,HotelInf,Remark,'" + Request.QueryString("pUserID") + "',"
        sql = sql + " getdate(),'" + Request.QueryString("pUserID") + "',getdate() from  F_BusinessTripSheetDTTemp"
        sql = sql + "  where 1=1 and No='" + DTNo.Text + "'"
        uDataBase.ExecuteNonQuery(sql)

        If wStep = 1 Then
            sql = " delete from F_BusinessTripSheetdt where No='" + DTNo.Text + "'"
            uDataBase.ExecuteNonQuery(sql)
        End If



        '主檔資料
        sql = " select  data,replace(data,'/','\')data1  from M_referp"
        sql = sql + " where cat = '3114'"
        sql = sql + " and dkey ='AttachfilePath' "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

        If DBAdapter1.Rows.Count > 0 Then
            sourceDir = "\\" + DBAdapter1.Rows(0).Item("Data1") + D3.Text   '來源
        End If

        '主檔資料
        sql = " select  data,replace(data,'/','\')data1  from M_referp"
        sql = sql + " where cat = '3114'"
        sql = sql + " and dkey ='AttachfilePath1' "
        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(sql)

        If DBAdapter2.Rows.Count > 0 Then
            backupDir = "\\" + DBAdapter2.Rows(0).Item("Data1") + DNo.Text      '目的  
        End If

        '複製前先確認是否有原本的資料夾再刪除
        If Directory.Exists(backupDir) Then
            Directory.Delete(backupDir, True)
        End If
        CopyDir(sourceDir, backupDir)
        NewAttachFilePath()



        ''主檔資料
        'sql = " select  data,replace(data,'/','\')data1  from M_referp"
        'sql = sql + " where cat = '3110'"
        'sql = sql + " and dkey ='AttachfilePath' "
        'Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

        'If DBAdapter1.Rows.Count > 0 Then
        '    sourceDir = "\\" + DBAdapter1.Rows(0).Item("Data1") + D3.Text   '來源
        'End If

        ''主檔資料
        'sql = " select  data,replace(data,'/','\')data1  from M_referp"
        'sql = sql + " where cat = '3110'"
        'sql = sql + " and dkey ='AttachfilePath1' "
        'Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(sql)

        'If DBAdapter2.Rows.Count > 0 Then
        '    backupDir = "\\" + DBAdapter2.Rows(0).Item("Data1") + DNo.Text      '目的  
        'End If


        ''sourceDir = "\\10.245.1.61\wfs$\N2W\003109Temp\" + D3.Text   '來源
        ''backupDir = "\\10.245.1.61\wfs$\N2W\003109\" + DNO.Text      '目的     

        ''複製前先確認是否有原本的資料夾再刪除
        'If Directory.Exists(backupDir) Then
        '    Directory.Delete(backupDir, True)
        'End If


        'CopyDir(sourceDir, backupDir)
        'NewAttachFilePath()




    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim sql As String

        sql = " Update F_BusinessTripSheet"
        sql = sql + " Set "

        '##-2
        '##AgentApprov-Start
        'If pFun <> "SAVE" Then
        '    sql = sql + " Sts = '" & pSts & "',"
        '    sql = sql + " CompletedTime = '" & NowDateTime & "',"
        'End If
        '
        If pFun <> "SAVE" And pFun <> "AGENTAPPROVE" Then
            sql = sql + " Sts = '" & pSts & "',"
            sql = sql + " CompletedTime = '" & NowDateTime & "',"
        End If
        '##AgentApprov-End
        '##
        sql = sql + " Date = N'" & DDate.Text & "',"
        sql = sql + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"
        sql = sql + " Division1 = N'" & YKK.ReplaceString(Trim(DDivision1.Text)) & "',"
        sql = sql + " Division2 = N'" & YKK.ReplaceString(Trim(DDivision2.Text)) & "',"
        sql = sql + " Division3 = N'" & YKK.ReplaceString(Trim(DDivision3.Text)) & "',"
        sql = sql + " Division4 = N'" & YKK.ReplaceString(Trim(DDivision4.Text)) & "',"
        sql = sql + " DivisionCode1 = N'" & YKK.ReplaceString(DDivisionCode1.Text) & "',"
        sql = sql + " DivisionCode2 = N'" & YKK.ReplaceString(DDivisionCode2.Text) & "',"
        sql = sql + " DivisionCode3 = N'" & YKK.ReplaceString(DDivisionCode3.Text) & "',"
        sql = sql + " DivisionCode4 = N'" & YKK.ReplaceString(DDivisionCode4.Text) & "',"
        sql = sql + " EmpID1 = N'" & YKK.ReplaceString(DEmpID1.Text) & "',"
        sql = sql + " EmpID2 = N'" & YKK.ReplaceString(DEmpID2.Text) & "',"
        sql = sql + " EmpID4 = N'" & YKK.ReplaceString(DEmpID4.Text) & "',"
        sql = sql + " EmpName1 = N'" & YKK.ReplaceString(DEmpName1.Text) & "',"
        sql = sql + " EmpName2 = N'" & YKK.ReplaceString(DEmpName2.Text) & "',"
        sql = sql + " EmpName4 = N'" & YKK.ReplaceString(DEmpName4.Text) & "',"
        sql = sql + " JobTitle1 = N'" & YKK.ReplaceString(DJobTitle1.Text) & "',"
        sql = sql + " JobTitle2 = N'" & YKK.ReplaceString(DJobTitle2.Text) & "',"
        sql = sql + " JobTitle4 = N'" & YKK.ReplaceString(DJobTitle4.Text) & "',"
        sql = sql + " PassDate = N'" & YKK.ReplaceString(DPassDate.Text) & "',"
        sql = sql + " QCNO = N'" & YKK.ReplaceString(DQCNO.Text) & "',"
        sql = sql + " QCNO1 = N'" & YKK.ReplaceString(DQCNO1.Text) & "',"
        If DChkPhone.Checked = True Then
            sql = sql + " Chkphone = 1,"
        Else
            sql = sql + " Chkphone = 0,"
        End If
        sql = sql + " PhoneNo = N'" & YKK.ReplaceString(DPhoneNo.Text) & "',"
        If DchkInsurance.Checked = True Then
            sql = sql + "chkInsurance = 1,"
        Else
            sql = sql + " chkInsurance = 0,"
        End If
        sql = sql + " InsurancePay = N'" & YKK.ReplaceString(DInsurancePay.Text) & "',"

        If DchkVisa.Checked = True Then
            sql = sql + " chkVisa = 1,"
        Else
            sql = sql + " chkVisa = 0,"
        End If
        sql = sql + " AirTickets = N'" & YKK.ReplaceString(DAirTickets.Text) & "',"
        sql = sql + " Object = N'" & YKK.ReplaceString(DObject.Text) & "',"
        sql = sql + " Location = N'" & YKK.ReplaceString(DLocation.Text) & "',"
        sql = sql + " SDate = N'" & YKK.ReplaceString(DSDate.Text) & "',"
        sql = sql + " EDate = N'" & YKK.ReplaceString(DEDate.Text) & "',"
        sql = sql + " Remark = N'" & YKK.ReplaceString(DRemark.Text) & "',"

     

        sql &= " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        sql &= " ModifyTime = '" & NowDateTime & "' "
        sql &= " Where FormNo  =  '" & wFormNo & "'"
        sql &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        uDataBase.ExecuteNonQuery(sql)




        ''存入明細


        ' delete from F_BusinessFlyDT  
        ' insert into F_BusinessFlyDT 
        ' select  *from F_BusinessFlyDTTemp  





    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Add T_CommissionNo)
    '**     追加交易資料和委託單對照表
    '**
    '*****************************************************************
    Sub AddCommissionNo(ByVal pFormNo As String, ByVal pFormSno As Integer)
        Dim SQl As String
        SQl = "Select * From T_CommissionNo "
        SQl = SQl & " Where FormNo =  '" & pFormNo & "'"
        SQl = SQl & "   And FormSno =  '" & CStr(pFormSno) & "'"
        Dim dtCommissionNo As DataTable = uDataBase.GetDataTable(SQl)

        If dtCommissionNo.Rows.Count <= 0 Then
            If DNo.Text <> "" Then
                SQl = "Insert into T_CommissionNo "
                SQl = SQl + "( "
                SQl = SQl + "FormNo, FormSno, No, MapNo, CreateUser, CreateTime "    '1~5
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '" + pFormNo + "', "
                SQl = SQl + " '" + CStr(pFormSno) + "', "
                SQl = SQl + " '" + DNo.Text + "', "
                SQl = SQl + " '" + "" + "', "
                SQl = SQl + " '" + Request.QueryString("pUserID") + "', "
                SQl = SQl + " '" + NowDateTime + "' "
                SQl = SQl + " ) "
                uDataBase.ExecuteNonQuery(SQl)
            End If
        Else
            If DNo.Text <> "" Then
                If DNo.Text <> dtCommissionNo.Rows(0)("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNo.Text & "',"
                    SQl = SQl + " MapNo = '" & "" & "',"
                    SQl = SQl + " CreateUser = '" & Request.QueryString("pUserID") & "',"
                    SQl = SQl + " CreateTime = '" & NowDateTime & "' "
                    SQl = SQl + " Where FormNo  =  '" & pFormNo & "'"
                    SQl = SQl + "   And FormSno =  '" & CStr(pFormSno) & "'"
                    uDataBase.ExecuteNonQuery(SQl)
                End If
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyTranData)
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub ModifyTranData(ByVal pFun As String, ByVal pSts As String)
        Dim SQl As String
        If pFun <> "SAVE" Then      '<> Save
            SQl = "Update T_WaitHandle Set "
            SQl = SQl + " Active = '" & "0" & "',"
            SQl = SQl + " Sts = '" & pSts & "',"
            If pSts = "1" Then SQl = SQl + " StsDes = '" & BOK.Text & "',"
            If pSts = "2" Then SQl = SQl + " StsDes = '" & BNG1.Text & "',"
            If pSts = "3" Then SQl = SQl + " StsDes = '" & BNG2.Text & "',"

            SQl = SQl + " AEndTime = '" & NowDateTime & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
            If DDelay.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & uCommon.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step    =  '" & CStr(wStep) & "'"
            SQl = SQl + "   And SeqNo   =  '" & CStr(wSeqNo) & "'"
            SQl = SQl + "   And Active =  '1' "
        Else
            SQl = "Update T_WaitHandle Set "
            If DReasonCode.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & uCommon.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step =  '" & CStr(wStep) & "'"
            SQl = SQl + "   And SeqNo =  '" & CStr(wSeqNo) & "'"
        End If
        uDataBase.ExecuteNonQuery(SQl)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     儲存按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BSAVE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSAVE.Click

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-2按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BNG2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNG2.Click
        If InputDataOK(2) Then
            DisabledButton()   '停止Button運作
            '##-3
            '##AgentApprov-Start
            If oCommon.UseAgentApprov(wFormNo, wStep, "NG2") = 0 Then
                AgentApprovProc("NG2", 2, "3")
            Else
                FlowControl("NG2", 2, "3")
            End If
            '##AgentApprov-End
            '##
        Else
            EnabledButton()   '起動Button運作
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-1按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BNG1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNG1.Click
        If InputDataOK(1) Then
            DisabledButton()   '停止Button運作
            '##-4
            '##AgentApprov-Start
            If oCommon.UseAgentApprov(wFormNo, wStep, "NG1") = 0 Then
                AgentApprovProc("NG1", 1, "2")
            Else
                FlowControl("NG1", 1, "2")
            End If
            '##AgentApprov-End
            '##
        Else
            EnabledButton()   '起動Button運作
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.Click
        If InputDataOK(1) Then
            DisabledButton()   '停止Button運作
            FlowControl("OK", 0, "1")
        Else
            EnabledButton()   '起動Button運作
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     停止Button(OK, Ng1, NG2, Save)運作 
    '**
    '*****************************************************************
    Private Sub DisabledButton()
        BOK.Enabled = False
        BNG1.Enabled = False
        BNG2.Enabled = False
        BSAVE.Enabled = False
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     起動Button(OK, Ng1, NG2, Save)運作 
    '**
    '*****************************************************************
    Private Sub EnabledButton()
        BOK.Enabled = True
        BNG1.Enabled = True
        BNG2.Enabled = True
        BSAVE.Enabled = True
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編製委託No 
    '**
    '*****************************************************************
    Function SetNo(ByVal Seq As Integer) As String
        Dim Str As String = ""
        Dim Str1 As String = ""
        Dim Str2 As String = ""
        Dim Str3 As String = ""
        Dim i As Integer

        'Set當日日期
        Str3 = Mid(CStr(DateTime.Now.Year), 3, 2)  '年

        Str2 = CStr(DateTime.Now.Month)  '月
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        Str = Str2
        Str2 = CStr(DateTime.Now.Day)    '日
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        Str = Str3 + Str + Str2
        'Set單號
        Str1 = CStr(Seq)
        For i = Str1.Length To 4 - 1
            Str1 = "0" + Str1
        Next

        SetNo = Str + Str1
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetDataStatus)
    '**     取得表單進度狀態
    '**
    '*****************************************************************
    Sub GetDataStatus()
        Dim SQL As String
        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)
        If dtFlow.Rows.Count > 0 Then
            'NG-1按鈕
            If dtFlow.Rows(0)("NGFun1") = 1 Then
                wNGSts1 = dtFlow.Rows(0)("NGSts1") + 1
            End If
            'NG-2按鈕
            If dtFlow.Rows(0)("NGFun2") = 1 Then
                wNGSts2 = dtFlow.Rows(0)("NGSts2") + 1
            End If
            'OK按鈕
            If dtFlow.Rows(0)("OKFun") = 1 Then
                wOKSts = dtFlow.Rows(0)("OKSts") + 1
            End If
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     檢查上傳檔案
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As FileUpload) As Integer
        Dim fileExtension As String     '宣告一個變數存放檔案格式(副檔名)
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".ai", ".xls", ".xlsx", ".doc", ".docx", ".ppt"}   '定義允許的檔案格式
        Dim i As Integer

        UPFileIsNormal = 0

        fileExtension = IO.Path.GetExtension(UPFile.PostedFile.FileName).ToLower   '取得檔案格式
        For i = 0 To allowedExtensions.Length - 1           '逐一檢查允許的格式中是否有上傳的格式
            If fileExtension = allowedExtensions(i) Then
                UPFileIsNormal = 0
                Exit For
            Else
                UPFileIsNormal = 9020
            End If
        Next
        'Check上傳檔案Size
        If UPFileIsNormal = 0 Then
            If UPFile.PostedFile.ContentLength <= 1000 * 1024 Then
                UPFileIsNormal = 0
            Else
                UPFileIsNormal = 9038
            End If
        End If
    End Function

    '*****************************************************************
    '**
    '**     判斷是否可繼續執行(驗證資料)
    '**
    '*****************************************************************
    Function InputDataOK(ByVal pAction As Integer) As Boolean
        Dim isOK As Boolean = False
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim Message As String = ""
        Dim SQL As String

        If wStep = 1 Then
            '明細
            SQL = " Select  Unique_ID,No, Type, Appoint,convert(char(10),SDate,111)+' '+stime1+':'+stime2 as SDate,"
            SQL = SQL + " convert(char(10),EDate,111)+' '+etime1+':'+etime2 as EDate,"
            SQL = SQL + " Days,Money,Currency, FlyInf,SFly,EFly,HotelInf, Remark  from  F_BusinessTripSheetDTTemp where 1=1 "
            SQL = SQL + " and NO = '" + DTNo.Text + "'"
            SQL = SQL + " order by unique_id "
            GridView1.DataSource = uDataBase.GetDataTable(SQL)
            GridView1.DataBind()
            If GridView1.Rows.Count > 0 Then
                DHaveData.Text = "1"
            Else
                DHaveData.Text = "0"
            End If

        End If


        '如果細項沒有資料不能送出
        If ErrCode = 0 Then
            If DHaveData.Text = 0 Then
                ErrCode = 9052
            End If

        End If


        '如果細項沒有資料不能送出
        If ErrCode = 0 Then
            If DchkInsurance.Checked = False Then
                ErrCode = 9053
            End If

        End If


        'Check延遲理由
        If ErrCode = 0 Then
            If DReasonCode.Visible = True Then
                If DReasonCode.SelectedValue = "99" Then
                    If DReasonDesc.Text = "" Then ErrCode = 9040
                End If
            End If
        End If

        '檢查委託書No
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                ErrCode = oCommon.CommissionNo("003114", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9050
                End If
            End If
        End If

        If ErrCode = 0 Then
            ' 檢查受款人職稱ID
            ' 如果受款人及跟申請人的主管相同就不能申請，但總經理例外
            If GetRelated(Request.QueryString("pUserID")) = GetUerID(DEmpName2.Text) Then
                If DJobTitle2.Text <> "總經理" Then
                    'ErrCode = 9051
                End If
            End If
        End If







        If ErrCode <> 0 Then
            If ErrCode = 9010 Then Message = "上傳檔案格式錯誤,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案格式有誤,請確認!"
            If ErrCode = 9038 Then Message = "上傳檔案Size有誤,請確認!"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9050 Then Message = "委託書No.重覆,請確認委託書No.!"
            If ErrCode = 9051 Then Message = "受款者與簽核者相同!"
            If ErrCode = 9052 Then Message = "細項沒有資料不能送出!"
            If ErrCode = 9053 Then Message = "海外保險需打勾!"
     
            uJavaScript.PopMsg(Me, Message)
        Else
            isOK = True
        End If
        '
        Return isOK
    End Function

    '取得關係人
    Function GetRelated(ByVal userId As String) As String

        Dim sql As String = "select RUserID,RRUserID,USERID  from M_Related where userid='" & userId & "' and RelatedID='B'"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then

            NextGate = dt.Rows(0)("RUserID")

        End If
        Return NextGate
    End Function


    '取得關係人
    Function GetUerID(ByVal Username As String) As String

        Dim sql As String = "select UserID from M_users where username='" & Username + "'"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            NextGate = dt.Rows(0)("UserID")
        End If
        Return NextGate
    End Function

    Protected Sub CDivision3_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DCDivision3.CheckedChanged
        If DCDivision3.Checked Then
            If DEmpID2.Text <> "" Then
                Dim SQL As String

                SQL = " Select  * From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a "
                SQL = SQL & "  where  EmpID = '" & DEmpID2.Text & "'"


                Dim DBUser As DataTable = uDataBase.GetDataTable(SQL)

                If DBUser.Rows(0).Item("EmpName") = "陳怡君" Then
                    SQL = " Select a.* From m_wfsemp a,m_users b"
                    SQL = SQL & "  where a.empname =b.username and depid = hrwdivid "
                    SQL = SQL & " and UserID = '" & wApplyID & "'"
                    SQL = SQL & "   And Active = '1' "
                    Dim DBUser1 As DataTable = uDataBase.GetDataTable(SQL)

                    DDivision3.Text = DBUser1.Rows(0).Item("CostName")
                    DDivisionCode3.Text = DBUser1.Rows(0).Item("CostID")
                Else

                    DDivision3.Text = DBUser.Rows(0).Item("CostName")
                    DDivisionCode3.Text = DBUser.Rows(0).Item("CostID")
                End If
            Else

            End If
           
        End If

    End Sub

    Protected Sub CDivision2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CDivision2.CheckedChanged
        If CDivision2.Checked Then
            DDivision2.Text = DDivision1.Text
            DEmpName2.Text = DEmpName1.Text
            DDivisionCode2.Text = DDivisionCode1.Text
            DEmpID2.Text = DEmpID1.Text
            DJobTitle2.Text = DJobTitle1.Text
        End If


    End Sub

    Sub GetData()
        Dim SQL As String

        ' Dim DBDataSet1 As New DataSet
        ' Dim OleDbConnection1 As New OleDbConnection
        ' OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        SQL = " Select  Unique_ID,No, Type, Appoint,convert(char(10),SDate,111)+' '+stime1+':'+stime2 as SDate,"
        SQL = SQL + " convert(char(10),EDate,111)+' '+etime1+':'+etime2 as EDate,"
        SQL = SQL + " Days,Money,Currency,Pay, FlyInf,SFly,EFly,HotelInf,Remark  from  F_BusinessTripSheetDT where 1=1 "
        SQL = SQL + " and NO = '" + DTNo.Text + "'"
        SQL = SQL + " order by typeno,sdate "

        '  OleDbConnection1.Open()
        '  Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        '  DBAdapter1.Fill(DBDataSet1, "Getata")
        ' GridView1.DataSource = DBDataSet1
        'GridView1.DataBind()

        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()
 

        '  OleDbConnection1.Close()

        SQL = " Select Unique_ID,No,FlyType,Pay, Money1,Currency1,Remark1,convert(char(10),SDate,112)Sdate,SNo "
        SQL = SQL + "   from  F_BusinessFlyDT  where 1=1 "
        SQL = SQL + " and NO = '" + DTNo.Text + "'"
        SQL = SQL + " order by unique_id "

        GridView3.DataSource = uDataBase.GetDataTable(SQL)
        GridView3.DataBind()

        'SQL = "Select isnull(sum(Amt),0) as Amt  from F_BusinessTripSheetdtTemp "
        'SQL = SQL & " Where DTNO =  '" & DTNo.Text & "'"
        'SQL = SQL & " AND Amt > 0 "
        'Dim dtExpData1 As DataTable = uDataBase.GetDataTable(SQL)
        'If dtExpData1.Rows(0).Item("Amt") <> 0 Then
        '    DApplyAmt.Text = String.Format("{0:N0}", Val(dtExpData1.Rows(0).Item("Amt")))
        'End If
        ''
        'SQL = "Select isnull(sum(Amt),0)*-1 as Amt  from F_BusinessTripSheetdtTemp "
        'SQL = SQL & " Where DTNO =  '" & DTNo.Text & "'"
        'SQL = SQL & " AND Amt < 0 "
        'Dim dtExpData2 As DataTable = uDataBase.GetDataTable(SQL)
        'If dtExpData2.Rows(0).Item("Amt") <> 0 Then
        '    DDebitAmt.Text = String.Format("{0:N0}", Val(dtExpData2.Rows(0).Item("Amt")))
        'End If
        ''



    End Sub



    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定附檔路徑 1
    '**
    '*****************************************************************
    Sub NewAttachFilePath()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '3114'"
        SQL = SQL + " and dkey ='AttachfilePath'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If


        If D3.Text = "" Then
            D3.Text = Now.ToString("yyyyMMddHHmmss")
        End If



        OpenDir1 = OpenDir1 + D3.Text + "/AttachFile"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + D3.Text + "\AttachFile"
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If
        '開啟附檔資料夾路徑
        DAttachfile1.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")

    End Sub


    '複製資料夾
    Public Shared Sub CopyDir(ByVal srcPath As String, ByVal aimPath As String)
        Try
            ' 檢查目標目錄是否以目錄分割字元結束如果不是則新增之
            If aimPath(aimPath.Length - 1) <> Path.DirectorySeparatorChar Then
                aimPath += Path.DirectorySeparatorChar
            End If
            ' 判斷目標目錄是否存在如果不存在則新建之
            If Not Directory.Exists(aimPath) Then
                Directory.CreateDirectory(aimPath)
            End If
            ' 得到源目錄的檔案列表，該裡面是包含檔案以及目錄路徑的一個數組
            ' 如果你指向copy目標檔案下面的檔案而不包含目錄請使用下面的方法
            ' string[] fileList = Directory.GetFiles(srcPath);
            Dim fileList() As String = Directory.GetFileSystemEntries(srcPath)
            ' 遍歷所有的檔案和目錄
            Dim file As String
            For Each file In fileList
                ' 先當作目錄處理如果存在這個目錄就遞迴Copy該目錄下面的檔案
                If Directory.Exists(file) Then
                    CopyDir(file, aimPath + Path.GetFileName(file))
                Else
                    System.IO.File.Copy(file, aimPath + Path.GetFileName(file), True)
                End If

                ' 否則直接Copy檔案

            Next
        Catch
            Console.WriteLine("無法複製!")
        End Try
    End Sub

    Protected Sub D1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles D1.TextChanged
        Dim SQL As String


        'SQL = " Select a.* From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a ,m_users b"
        'SQL = SQL & "  where a.empname =b.username"
        'SQL = SQL & " and a.EMPID = '" & D1.Text & "'"
        'SQL = SQL & "   And Active = '1' "

        SQL = " Select  * From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a "
        SQL = SQL & "  where  EmpID = '" & D1.Text & "'"



        'SQL = " Select a.* From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a ,m_users b"
        'SQL = SQL & "  where a.empname =b.username"
        'SQL = SQL & " and UserID = '" & wApplyID & "'"
        'SQL = SQL & "   And Active = '1' "


        Dim DBUser As DataTable = uDataBase.GetDataTable(SQL)


        DDivision2.Text = DBUser.Rows(0).Item("DepName")
        DDivisionCode2.Text = DBUser.Rows(0).Item("DepID")
        DEmpID2.Text = DBUser.Rows(0).Item("EmpID")
        DEmpName2.Text = DBUser.Rows(0).Item("EmpName")
        DJobTitle2.Text = DBUser.Rows(0).Item("JobTitle")


    End Sub

    Protected Sub D2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles D2.TextChanged
        Dim SQL As String

 
        SQL = " Select  distinct CostID,CostName  From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a "
        SQL = SQL & "  where  CostID = '" & D2.Text & "'"


        Dim DBUser As DataTable = uDataBase.GetDataTable(SQL)


        DDivision3.Text = DBUser.Rows(0).Item("CostName")
        DDivisionCode3.Text = DBUser.Rows(0).Item("CostID")


    End Sub

    Protected Sub D4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles D4.TextChanged
        Dim SQL As String



        '  SQL = " Select a.* From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a ,m_users b"
        '  SQL = SQL & "  where a.empname =b.username"
        '  SQL = SQL & " and a.EMPID = '" & D4.Text & "'"
        '  SQL = SQL & "   And Active = '1' "

        SQL = " Select  * From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a "
        SQL = SQL & "  where  EmpID = '" & D4.Text & "'"


        Dim DBUser As DataTable = uDataBase.GetDataTable(SQL)


        DDivision4.Text = DBUser.Rows(0).Item("DepName")
        DDivisionCode4.Text = DBUser.Rows(0).Item("DepID")
        DEmpID4.Text = DBUser.Rows(0).Item("EmpID")
        DEmpName4.Text = DBUser.Rows(0).Item("EmpName")
        DJobTitle4.Text = DBUser.Rows(0).Item("JobTitle")

    End Sub


    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        e.Row.Cells(0).Visible = False
    End Sub


    Protected Sub GridView3_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView3.RowCreated
        e.Row.Cells(0).Visible = False
    End Sub

    Protected Sub DQCNO_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DQCNO.TextChanged
      
    End Sub

    Protected Sub DQC_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DQC.TextChanged
        Dim sql As String
        sql = " select * from F_ComplaintOutSheet where no = '" & DQC.Text & "'"
        Dim DBQC As DataTable = uDataBase.GetDataTable(sql)
        DQCNO1.Text = DBQC.Rows(0).Item("NO")
        DQCNO.Text = DBQC.Rows(0).Item("QCNO")
    End Sub

    Protected Sub DQCNO1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DQCNO1.TextChanged

    End Sub
 

 
End Class
