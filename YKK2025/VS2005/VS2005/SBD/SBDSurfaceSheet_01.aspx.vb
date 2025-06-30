Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
 


Partial Class SBDSurfaceSheet_01
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
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID
    Dim wAgentID As String          '被代理人ID
    Dim NowDateTime As String       '現在日期時間
    Dim wNextGate As String         '下一關人
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    'Modify-Start by Joy 2009/11/20(2010行事曆對應)
    'Dim wDepo As String = "CL"      '台北行事曆(CL->中壢, TP->台北)
    '
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    'Modify-End

    Dim wUserName As String = ""    '姓名代理用
    Dim HolidayList As New List(Of Integer) '用以記錄假日的欄位索引值
    Dim indexList As New List(Of Integer)   '用以記錄不屬於選取月份的欄位索引值
    Dim DateList As New List(Of String)     '用以記錄所選取的一周日期

    ''' <summary>
    ''' 以下為共用函式的宣告
    ''' </summary>
    ''' <remarks></remarks>
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    Dim aAgentID As String
    Dim Makemap1, Makemap2, Makemap3, Makemap4, Makemap5, Makemap6 As Integer



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '設定共用參數
        '  TopPosition()           '按鈕及RequestedField的Top位置
        SetControlPosition()    ' 設定控制項位置
        If Not Me.IsPostBack Then   '不是PostBack
            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()     '表單功能按鈕顯示

            If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
                ShowFormData()      '顯示表單資料
                UpdateTranFile()    '更新交易資料
                SetControlPosition()    ' 設定控制項位置
            End If
            SetPopupFunction()      '設定彈出視窗事件

        Else
            ShowSheetFunction()     '表單功能按鈕顯示
            ' fpObj.GetDeafultNextGate(DNextGate, DDivision.Text, Request.QueryString("pUserID")) ' 設定預設的簽核者
            ShowSheetField("Post")  '表單欄位顯示及欄位輸入檢查

            ShowMessage()           '上傳資料檢查及顯示訊息

        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日時
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
        wStep = Request.QueryString("pStep")        '工程代碼
        wSeqNo = Request.QueryString("pSeqNo")      '序號
        wApplyID = Request.QueryString("pApplyID")  '申請者ID
        wAgentID = Request.QueryString("pAgentID")  '被代理人ID
        'Add-Start by Joy 2009/11/20(2010行事曆對應)
        '申請者-群組行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        wDecideCalendar = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))
        'Add-End

        Response.Cookies("PGM").Value = "SBDSurfaceSheet_01.aspx"
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
        'Modify-Start by Joy 2009/11/20(2010行事曆對應)
        'oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDepo, Request.QueryString("pUserID"))
        '
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, Request.QueryString("pUserID"))
        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者
        'Modify-End
    End Sub



    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim SQL As String
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("SBDSufaceFilePath")



        SQL = "Select * From F_SBDSufaceSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then

            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            DAppDate.Text = DBAdapter1.Rows(0).Item("AppDate")              'AppDate
            DDivision.Text = DBAdapter1.Rows(0).Item("Division")              'Division
            DAppper.Text = DBAdapter1.Rows(0).Item("Appper")              'Appper
            SetFieldData("Buyer", DBAdapter1.Rows(0).Item("Buyer"))    'Buyer
            DSellVendor.Text = DBAdapter1.Rows(0).Item("SellVendor")              'Vendor
            DReqDelDate.Value = DBAdapter1.Rows(0).Item("ReqDelDate")              '希望交期
            DReqQty.Text = DBAdapter1.Rows(0).Item("ReqQty")              '希望交期
            DCode.Text = DBAdapter1.Rows(0).Item("Code")              '品名
            SetFieldData("AttachSample", DBAdapter1.Rows(0).Item("AttachSample"))    '附樣
            DORNO.Text = DBAdapter1.Rows(0).Item("ORNO")              'ORNO

            If Mid(DBAdapter1.Rows(0).Item("OrderTime").ToString, 1, 4) = "1900" Then
                DOrderTime.Value = ""
            Else
                DOrderTime.Value = DBAdapter1.Rows(0).Item("OrderTime")               '下單時間
            End If


            DDevReason.Text = DBAdapter1.Rows(0).Item("DevReason")              '開發理由
            DColor.Text = DBAdapter1.Rows(0).Item("Color")              'color code
            SetFieldData("SampleQty", DBAdapter1.Rows(0).Item("SampleQty"))    '附樣
            DEnglishName.Text = DBAdapter1.Rows(0).Item("EnglishName")              '英文名稱
            SetFieldData("Supplier", DBAdapter1.Rows(0).Item("Supplier"))    '外注廠商
            SetFieldData("AllowSample", DBAdapter1.Rows(0).Item("AllowSample"))    '限度樣品
            DCap.Text = DBAdapter1.Rows(0).Item("Cap")              '日產能

            If Mid(DBAdapter1.Rows(0).Item("BFinalDate").ToString, 1, 4) = "1900" Then
                DBFinalDate.Value = ""
            Else
                DBFinalDate.Value = DBAdapter1.Rows(0).Item("BFinalDate")                '預定完成日
            End If




            DSchedule.Text = DBAdapter1.Rows(0).Item("Schedule")              '基礎日程

            If Trim(DBAdapter1.Rows(0).Item("QCReqFile")) <> "" Then
                LQCReqFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("QCReqFile")  '品質依賴書
                LQCReqFile.Visible = True
            Else
                LQCReqFile.Visible = False
            End If

            If Trim(DBAdapter1.Rows(0).Item("QAFile")) <> "" Then
                LQAFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("QAFile")  '品測報告書
                LQAFile.Visible = True
            Else
                LQAFile.Visible = False
            End If


            If Mid(DBAdapter1.Rows(0).Item("QCDate").ToString, 1, 4) = "1900" Then
                DQCDate.Value = ""
            Else
                DQCDate.Value = DBAdapter1.Rows(0).Item("QCDate")          '品質判定日期
            End If




            SetFieldData("QCResult", DBAdapter1.Rows(0).Item("QCResult"))    '檢測結果
            DQCRemark.Text = DBAdapter1.Rows(0).Item("QCRemark")              '品質備註


            If Trim(DBAdapter1.Rows(0).Item("ManufFlowFile")) <> "" Then
                LManufFlowFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ManufFlowFile")  '製造流程
                LManufFlowFile.Visible = True
            Else
                LManufFlowFile.Visible = False
            End If

            If Trim(DBAdapter1.Rows(0).Item("OPManualFile")) <> "" Then
                LOPManualFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("OPManualFile")  '作業流程
                LOPManualFile.Visible = True
            Else
                LOPManualFile.Visible = False
            End If


            If Trim(DBAdapter1.Rows(0).Item("ForcastFile")) <> "" Then
                LForcastFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ForcastFile")  ' 報價單
                LForcastFile.Visible = True
            Else
                LForcastFile.Visible = False
            End If


            If Trim(DBAdapter1.Rows(0).Item("ContactFile")) <> "" Then
                LContactFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContactFile")  '切結書
                LContactFile.Visible = True
            Else
                LContactFile.Visible = False
            End If

            If Trim(DBAdapter1.Rows(0).Item("CustSampleFile")) <> "" Then
                LCustSampleFile.ImageUrl = Path & DBAdapter1.Rows(0).Item("CustSampleFile")  '切結書
                LCustSampleFile.Visible = True
            Else
                LCustSampleFile.Visible = False
            End If

            If Trim(DBAdapter1.Rows(0).Item("FinalSampleFile")) <> "" Then
                LFinalSampleFile.ImageUrl = Path & DBAdapter1.Rows(0).Item("FinalSampleFile")  '切結書
                LFinalSampleFile.Visible = True
            Else
                LFinalSampleFile.Visible = False
            End If




            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"

            Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)

            If DBAdapter2.Rows.Count > 0 Then
                DDecideDesc.Text = DBAdapter2.Rows(0).Item("DecideDesc")       '說明


                If DBAdapter2.Rows(0).Item("BEndTime") < NowDateTime Then
                    If DDelay.Visible = True Then
                        SetFieldData("ReasonCode", DBAdapter2.Rows(0).Item("ReasonCode"))    '延遲理由代碼
                        If DBAdapter2.Rows(0).Item("ReasonCode") = "" Then
                            SetFieldData("Reason", DReasonCode.SelectedValue)    '延遲理由
                            DReasonDesc.Text = ""   '延遲其他說明
                        Else
                            DReason.Text = DBAdapter2.Rows(0).Item("Reason")  '延遲理由
                            DReasonDesc.Text = DBAdapter2.Rows(0).Item("ReasonDesc")     '延遲其他說明
                        End If
                    End If
                End If

            End If
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
        SQL = SQL + "Order by Unique_ID Desc "
        Dim dtWaitHandle1 As DataTable = uDataBase.GetDataTable(SQL)

        'Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter4.Fill(DBDataSet9, "DecideHistory")
        GridView2.DataSource = dtWaitHandle1
        GridView2.DataBind()

        'DB連結關閉
        'OleDbConnection1.Close()

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetFunction)
    '**     表單功能按鈕顯示
    '**
    '*****************************************************************
    Sub ShowSheetFunction()
        Dim wDelay As Integer = 0
        Dim SQL As String

        'DB連結設定

        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)


        If DBAdapter1.Rows.Count > 0 Then
            '電子簽章未使用
            If DBAdapter1.Rows(0).Item("SignImage") = 1 Then
            Else
            End If
            '附加檔案未使用(由FormField中設定)
            If DBAdapter1.Rows(0).Item("Attach") = 1 Then
            Else
            End If
            '儲存按鈕
            If DBAdapter1.Rows(0).Item("SaveFun") = 1 Then
                BSAVE.Visible = True
                BSAVE.Value = DBAdapter1.Rows(0).Item("SaveDesc")
                BSAVE.Attributes("onclick") = "Button('SAVE', '" + BSAVE.Value + "');"
            Else
                BSAVE.Visible = False
            End If
            'NG-1按鈕
            If DBAdapter1.Rows(0).Item("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Value = DBAdapter1.Rows(0).Item("NGDesc1")
                BNG1.Attributes("onclick") = "Button('NG1', '" + BNG1.Value + "');"
                wNGSts1 = DBAdapter1.Rows(0).Item("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2按鈕
            If DBAdapter1.Rows(0).Item("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Value = DBAdapter1.Rows(0).Item("NGDesc2")
                BNG2.Attributes("onclick") = "Button('NG2', '" + BNG2.Value + "');"
                wNGSts2 = DBAdapter1.Rows(0).Item("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK按鈕
            If DBAdapter1.Rows(0).Item("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Value = DBAdapter1.Rows(0).Item("OKDesc")
                BOK.Attributes("onclick") = "Button('OK', '" + BOK.Value + "');"
                wOKSts = DBAdapter1.Rows(0).Item("OKSts") + 1
            Else
                BOK.Visible = False
            End If
            '遲納管理
            If DBAdapter1.Rows(0).Item("Delay") = 1 Then
                wDelay = 1
            End If
        End If

        If wFormSno > 0 And wStep > 2 Then      '判斷是否[簽核]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)

            If DBAdapter2.Rows.Count > 0 Then

                '遲納管理
                If wDelay = 1 Then
                    If DBAdapter2.Rows(0).Item("BEndTime") < NowDateTime Then
                        DDelay.Visible = True   '延遲Sheet
                    Else
                        DDelay.Visible = False  '延遲Sheet
                    End If
                End If
                If DDelay.Visible = True Then
                    DReasonCode.Visible = True     '延遲理由代碼
                    DReasonCode.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonCodeRqd", "DReasonCode", "異常：需輸入延遲理由")
                    DReason.Visible = True         '延遲理由
                    DReason.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonRqd", "DReason", "異常：需輸入延遲理由")
                    DReasonDesc.Visible = True     '延遲其他說明
                    Top = 1100
                Else
                    DReasonCode.Visible = False     '延遲理由代碼
                    DReason.Visible = False         '延遲理由
                    DReasonDesc.Visible = False     '延遲其他說明
                    Top = 1000
                End If



                '欄位顯示
                DDecideDesc.Visible = True      '說明
                '說明需輸入
                If DBAdapter2.Rows(0).Item("FlowType") = 0 Then
                    DDecideDesc.BackColor = Color.LightGray
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "異常：需輸入說明")
                End If

                '按鈕位置
                BNG1.Style.Add("Top", Top)      'NG按鈕
                BNG2.Style.Add("Top", Top)     'NG按鈕
                BSAVE.Style.Add("Top", Top)    '儲存按鈕
                BOK.Style.Add("Top", Top)      'OK按鈕

            End If
        Else
            Top = 800
            'Sheet隱藏

            DDelay.Visible = False      '延遲Sheet
            '欄位隱藏
            DDecideDesc.Visible = False     '說明
            DDescSheet.Visible = False
            DReasonCode.Visible = False     '延遲理由代碼
            DReason.Visible = False         '延遲理由
            DReasonDesc.Visible = False     '延遲其他說明
            DHistoryLabel.Visible = False  '核定履歷
            '按鈕位置
            BNG1.Style.Add("Top", Top)      'NG按鈕
            BNG2.Style.Add("Top", Top)     'NG按鈕
            BSAVE.Style.Add("Top", Top)    '儲存按鈕
            BOK.Style.Add("Top", Top)      'OK按鈕

        End If

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        '日期選擇
        BDate1.Attributes.Add("onClick", "calendarPicker('DOrderTime')")
        BDate2.Attributes.Add("onClick", "calendarPicker('DBFinalDate')")
        BDate3.Attributes.Add("onClick", "calendarPicker('DQCDate')")
        BDate4.Attributes.Add("onClick", "calendarPicker('DReqDelDate')")

        'BDate1.Attributes("onclick") = "calendarPicker('Form1.DMapDate');"
        'BDate2.Attributes("onclick") = "calendarPicker('Form1.DSampleDate');"
        'BDate3.Attributes("onclick") = "calendarPicker('Form1.DHalfFinishDdate');"

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     延遲理由代碼點選後事件
    '**
    '*****************************************************************
    Private Sub DReasonCode_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DReasonCode.SelectedIndexChanged
        SetFieldData("Reason", DReasonCode.SelectedValue)    '延遲理由
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     按鈕及RequestedField的Top位置
    '**
    '*****************************************************************
    Sub TopPosition()
        Dim SQL As String


        '按鈕及RequestedField的Top位置
        If wFormSno > 0 And wStep > 2 Then      '判斷是否[簽核]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

            If DBAdapter1.Rows.Count > 0 Then
                If DBAdapter1.Rows(0).Item("BEndTime") >= NowDateTime Then
                    Top = 1100
                Else
                    If DDelay.Visible = True Then
                        Top = 1600
                    Else
                        Top = 1100
                    End If
                End If
            End If
        Else
            Top = 696
        End If
        '----


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
    '**(ShowSheetField)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '表單各欄位屬性及欄位輸入檢查等設定

        '客戶樣品圖
        Select Case FindFieldInf("CustSampleFile")
            Case 0  '顯示
                DCustSampleFile.Visible = False
                DCustSampleFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DCustSampleFileRqd", "DCustSampleFile", "異常：需輸入客戶樣品圖")
                DCustSampleFile.Visible = True
                DCustSampleFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DCustSampleFile.Visible = True
                DCustSampleFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DCustSampleFile.Visible = False
        End Select
        LCustSampleFile.Visible = False

        '最終樣品圖
        Select Case FindFieldInf("FinalSampleFile")
            Case 0  '顯示
                DFinalSampleFile.Visible = False
                DFinalSampleFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DFinalSampleFileRqd", "DFinalSampleFile", "異常：需輸入最終樣品圖")
                DFinalSampleFile.Visible = True
                DFinalSampleFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DFinalSampleFile.Visible = True
                DFinalSampleFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DFinalSampleFile.Visible = False
        End Select
        LFinalSampleFile.Visible = False

        'No
        Select Case FindFieldInf("No")
            Case 0  '顯示
                DNo.BackColor = Color.LightGray
                DNo.ReadOnly = True
                DNo.Visible = True
            Case 1  '修改+檢查
                DNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DnoRqd", "Dno", "異常：需輸入Ｎｏ")
                DNo.Visible = True
            Case 2  '修改
                DNo.BackColor = Color.Yellow
                DNo.Visible = True
            Case Else   '隱藏
                DNo.Visible = False
        End Select
        If pPost = "New" Then DNo.Text = ""

        '日期
        Select Case FindFieldInf("AppDate")
            Case 0  '顯示
                DAppDate.BackColor = Color.LightGray
                DAppDate.ReadOnly = True
                DAppDate.Visible = True

            Case 1  '修改+檢查
                DAppDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAppDateRqd", "DDate", "異常：需輸入日期")
                DAppDate.Visible = True

            Case 2  '修改
                DAppDate.BackColor = Color.Yellow
                DAppDate.Visible = True

            Case Else   '隱藏
                DAppDate.Visible = False

        End Select
        If pPost = "New" Then DAppDate.Text = Now.ToString("yyyy/MM/dd") '現在日時

        '修改部門
        Select Case FindFieldInf("Division")
            Case 0  '顯示
                DDivision.BackColor = Color.LightGray
                DDivision.ReadOnly = True
                DDivision.Visible = True
            Case 1  '修改+檢查
                DDivision.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDivisionRqd", "DDivision", "異常：需輸入部門")
                DDivision.Visible = True
            Case 2  '修改
                DDivision.BackColor = Color.Yellow
                DDivision.Visible = True
            Case Else   '隱藏
                DDivision.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Division", "ZZZZZZ")

        '擔當
        Select Case FindFieldInf("Appper")
            Case 0  '顯示
                DAppper.BackColor = Color.LightGray
                DAppper.Visible = True
                DAppper.ReadOnly = True
            Case 1  '修改+檢查
                DAppper.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAppperRqd", "DAppper", "異常：需輸入擔當")
                DAppper.Visible = True
            Case 2  '修改
                DAppper.BackColor = Color.Yellow
                DAppper.Visible = True
            Case Else   '隱藏
                DAppper.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Appper", "ZZZZZZ")



        'buyer
        Select Case FindFieldInf("Buyer")
            Case 0  '顯示
                DBuyer.BackColor = Color.LightGray
                DBuyer.Visible = True

            Case 1  '修改+檢查
                DBuyer.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerRqd", "DBuyer", "異常：需輸入Buyer")
                DBuyer.Visible = True
            Case 2  '修改
                DBuyer.BackColor = Color.Yellow
                DBuyer.Visible = True
            Case Else   '隱藏
                DBuyer.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Buyer", "ZZZZZZ")


        '委託廠商
        Select Case FindFieldInf("SellVendor")
            Case 0  '顯示
                DSellVendor.BackColor = Color.LightGray
                DSellVendor.ReadOnly = True
                DSellVendor.Visible = True
            Case 1  '修改+檢查
                DSellVendor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSellVendorRqd", "DSellVendor", "異常：需輸入委託廠商")
                DSellVendor.Visible = True
            Case 2  '修改
                DSellVendor.BackColor = Color.Yellow
                DSellVendor.Visible = True
            Case Else   '隱藏
                DSellVendor.Visible = False
        End Select
        If pPost = "New" Then DSellVendor.Text = ""






        '希望交期
        Select Case FindFieldInf("ReqDelDate")
            Case 0  '顯示
                DReqDelDate.Visible = True
                DReqDelDate.Style.Add("background-color", "lightgrey")
                DReqDelDate.Attributes.Add("readonly", "true")
                BDate4.Disabled = True

            Case 1  '修改+檢查
                DReqDelDate.Visible = True
                DReqDelDate.Style.Add("background-color", "greenyellow")
                DReqDelDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DReqDelDateRqd", "DReqDelDate", "異常：需輸入希望交期")
                BDate4.Disabled = False

            Case 2  '修改
                DReqDelDate.Visible = True
                DReqDelDate.Style.Add("background-color", "yellow")
                DReqDelDate.Attributes.Add("readonly", "true")
                BDate4.Disabled = False

            Case Else   '隱藏
                DReqDelDate.Visible = False
                BDate4.Disabled = True

        End Select
        If pPost = "New" Then DReqDelDate.Value = ""


        '大貨預估量
        Select Case FindFieldInf("ReqQty")
            Case 0  '顯示
                DReqQty.BackColor = Color.LightGray
                DReqQty.Visible = True
                DReqQty.ReadOnly = True
            Case 1  '修改+檢查
                DReqQty.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DReqQtyRqd", "DReqQty", "異常：需輸入大貨預估量")
                DReqQty.Visible = True
            Case 2  '修改
                DReqQty.BackColor = Color.Yellow
                DReqQty.Visible = True
            Case Else   '隱藏
                DReqQty.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ReqQty", "ZZZZZZ")



        '品名
        Select Case FindFieldInf("Code")
            Case 0  '顯示
                DCode.BackColor = Color.LightGray
                DCode.Visible = True
                DCode.ReadOnly = True
            Case 1  '修改+檢查
                DCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCodeRqd", "DCode", "異常：需輸入品名")
                DCode.Visible = True
            Case 2  '修改
                DCode.BackColor = Color.Yellow
                DCode.Visible = True
            Case Else   '隱藏
                DCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Code", "ZZZZZZ")



        '附樣
        Select Case FindFieldInf("AttachSample")
            Case 0  '顯示
                DAttachSample.BackColor = Color.LightGray
                DAttachSample.Visible = True

            Case 1  '修改+檢查
                DAttachSample.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAttachSampleRqd", "DAttachSample", "異常：需輸入附樣")
                DAttachSample.Visible = True
            Case 2  '修改
                DAttachSample.BackColor = Color.Yellow
                DAttachSample.Visible = True
            Case Else   '隱藏
                DAttachSample.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AttachSample", "ZZZZZZ")



        'ORNO
        Select Case FindFieldInf("ORNO")
            Case 0  '顯示
                DORNO.BackColor = Color.LightGray
                DORNO.Visible = True
                DORNO.ReadOnly = True
            Case 1  '修改+檢查
                DORNO.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DORNORqd", "DORNO", "異常：需輸入ORNO")
                DORNO.Visible = True
            Case 2  '修改
                DORNO.BackColor = Color.Yellow
                DORNO.Visible = True
            Case Else   '隱藏
                DORNO.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ORNO", "ZZZZZZ")


        '下單時間
        Select Case FindFieldInf("OrderTime")
            Case 0  '顯示
                DOrderTime.Visible = True
                DOrderTime.Style.Add("background-color", "lightgrey")
                DOrderTime.Attributes.Add("readonly", "true")
                BDate1.Disabled = True

            Case 1  '修改+檢查
                DOrderTime.Visible = True
                DOrderTime.Style.Add("background-color", "greenyellow")
                DOrderTime.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DOrderTimeRqd", "DOrderTime", "異常：需輸入下單時間")
                BDate1.Disabled = False

            Case 2  '修改
                DOrderTime.Visible = True
                DOrderTime.Style.Add("background-color", "yellow")
                DOrderTime.Attributes.Add("readonly", "true")
                BDate1.Disabled = False

            Case Else   '隱藏
                DOrderTime.Visible = False
                BDate1.Disabled = True

        End Select
        If pPost = "New" Then DOrderTime.Value = ""



        '開發理由
        Select Case FindFieldInf("DevReason")
            Case 0  '顯示
                DDevReason.BackColor = Color.LightGray
                DDevReason.Visible = True
                DDevReason.ReadOnly = True
            Case 1  '修改+檢查
                DDevReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDevReasonRqd", "DDevReason", "異常：需輸入開發理由")
                DDevReason.Visible = True
            Case 2  '修改
                DDevReason.BackColor = Color.Yellow
                DDevReason.Visible = True
            Case Else   '隱藏
                DDevReason.Visible = False
        End Select
        If pPost = "New" Then DDevReason.Text = ""




        'Color Code
        Select Case FindFieldInf("Color")
            Case 0  '顯示
                DColor.BackColor = Color.LightGray
                DColor.Visible = True
                DColor.ReadOnly = True
            Case 1  '修改+檢查
                DColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DColorRqd", "DColor", "異常：需輸入Color Code")
                DColor.Visible = True
            Case 2  '修改
                DColor.BackColor = Color.Yellow
                DColor.Visible = True
            Case Else   '隱藏
                DColor.Visible = False
        End Select
        If pPost = "New" Then DColor.Text = ""


        '樣品需求量
        Select Case FindFieldInf("SampleQty")
            Case 0  '顯示
                DSampleQty.BackColor = Color.LightGray
                DSampleQty.Visible = True

            Case 1  '修改+檢查
                DSampleQty.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSampleQtyRqd", "DSampleQty", "異常：需輸入樣品需求量")
                DSampleQty.Visible = True
            Case 2  '修改
                DSampleQty.BackColor = Color.Yellow
                DSampleQty.Visible = True
            Case Else   '隱藏
                DSampleQty.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SampleQty", "ZZZZZZ")



        '英文名稱
        Select Case FindFieldInf("EnglishName")
            Case 0  '顯示
                DEnglishName.BackColor = Color.LightGray
                DEnglishName.Visible = True
                DEnglishName.ReadOnly = True
            Case 1  '修改+檢查
                DEnglishName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEnglishNameRqd", "DEnglishName", "異常：需輸入Color Code")
                DEnglishName.Visible = True
            Case 2  '修改
                DEnglishName.BackColor = Color.Yellow
                DEnglishName.Visible = True
            Case Else   '隱藏
                DEnglishName.Visible = False
        End Select
        If pPost = "New" Then DEnglishName.Text = ""


        '外注廠商
        Select Case FindFieldInf("Supplier")
            Case 0  '顯示
                DSupplier.BackColor = Color.LightGray
                DSupplier.Visible = True

            Case 1  '修改+檢查
                DSupplier.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSupplierRqd", "DSupplier", "異常：需輸入外注廠商")
                DSupplier.Visible = True
            Case 2  '修改
                DSupplier.BackColor = Color.Yellow
                DSupplier.Visible = True
            Case Else   '隱藏
                DSupplier.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Supplier", "ZZZZZZ")



        '限度樣品
        Select Case FindFieldInf("AllowSample")
            Case 0  '顯示
                DAllowSample.BackColor = Color.LightGray
                DAllowSample.Visible = True

            Case 1  '修改+檢查
                DAllowSample.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAllowSampleRqd", "DAllowSample", "異常：需輸入限度樣品")
                DAllowSample.Visible = True
            Case 2  '修改
                DAllowSample.BackColor = Color.Yellow
                DAllowSample.Visible = True
            Case Else   '隱藏
                DAllowSample.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AllowSample", "ZZZZZZ")



        '日產能
        Select Case FindFieldInf("Cap")
            Case 0  '顯示
                DCap.BackColor = Color.LightGray
                DCap.Visible = True
                DCap.ReadOnly = True
            Case 1  '修改+檢查
                DCap.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCapRqd", "DCap", "異常：需輸入日產能")
                DCap.Visible = True
            Case 2  '修改
                DCap.BackColor = Color.Yellow
                DCap.Visible = True
            Case Else   '隱藏
                DCap.Visible = False
        End Select
        If pPost = "New" Then DCap.Text = ""


        '預定完成日間
        Select Case FindFieldInf("BFinalDate")
            Case 0  '顯示
                DBFinalDate.Visible = True
                DBFinalDate.Style.Add("background-color", "lightgrey")
                DBFinalDate.Attributes.Add("readonly", "true")
                BDate2.Disabled = True

            Case 1  '修改+檢查
                DBFinalDate.Visible = True
                DBFinalDate.Style.Add("background-color", "greenyellow")
                DBFinalDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DBFinalDateRqd", "DBFinalDate", "異常：需輸入預定完成日間")
                BDate2.Disabled = False

            Case 2  '修改
                DBFinalDate.Visible = True
                DBFinalDate.Style.Add("background-color", "yellow")
                DBFinalDate.Attributes.Add("readonly", "true")
                BDate2.Disabled = False

            Case Else   '隱藏
                DBFinalDate.Visible = False
                BDate2.Disabled = True

        End Select
        If pPost = "New" Then DBFinalDate.Value = ""


        '基準日程
        Select Case FindFieldInf("Schedule")
            Case 0  '顯示
                DSchedule.BackColor = Color.LightGray
                DSchedule.Visible = True
                DSchedule.ReadOnly = True
            Case 1  '修改+檢查
                DSchedule.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DScheduleRqd", "DSchedule", "異常：需輸入基準日程")
                DSchedule.Visible = True
            Case 2  '修改
                DSchedule.BackColor = Color.Yellow
                DSchedule.Visible = True
            Case Else   '隱藏
                DSchedule.Visible = False
        End Select
        If pPost = "New" Then DSchedule.Text = ""



        '品質依賴書
        Select Case FindFieldInf("QCReqFile")

            Case 0  '顯示
                DQCReqFile.Visible = False
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQCReqFileRqd", "DQCReqFile", "異常：需輸入草圖")
                DQCReqFile.Visible = True
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DQCReqFile.Visible = True
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DQCReqFile.Visible = False
        End Select
        LQCReqFile.Visible = False


        '品測報告書
        Select Case FindFieldInf("QAFile")

            Case 0  '顯示
                DQAFile.Visible = False
                DQAFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQAFileRqd", "DQAFile", "異常：需輸入品測報告書")
                DQAFile.Visible = True
                DQAFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DQAFile.Visible = True
                DQAFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DQAFile.Visible = False
        End Select
        LQAFile.Visible = False



        '品質判定時間
        Select Case FindFieldInf("QCDate")
            Case 0  '顯示
                DQCDate.Visible = True
                DQCDate.Style.Add("background-color", "lightgrey")
                DQCDate.Attributes.Add("readonly", "true")
                BDate3.Disabled = True

            Case 1  '修改+檢查
                DQCDate.Visible = True
                DQCDate.Style.Add("background-color", "greenyellow")
                DQCDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DQCDateRqd", "DQCDate", "異常：需輸入品質判定時間")
                BDate3.Disabled = False

            Case 2  '修改
                DQCDate.Visible = True
                DQCDate.Style.Add("background-color", "yellow")
                DQCDate.Attributes.Add("readonly", "true")
                BDate3.Disabled = False

            Case Else   '隱藏
                DQCDate.Visible = False
                BDate3.Disabled = True

        End Select
        If pPost = "New" Then DQCDate.Value = ""



        '檢測結果
        Select Case FindFieldInf("QCResult")
            Case 0  '顯示
                DQCResult.BackColor = Color.LightGray
                DQCResult.Visible = True
            Case 1  '修改+檢查
                DQCResult.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCResultRqd", "DQCResult", "異常：需輸入檢測結果")
                DQCResult.Visible = True
            Case 2  '修改
                DQCResult.BackColor = Color.Yellow
                DQCResult.Visible = True
            Case Else   '隱藏
                DQCResult.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCResult", "ZZZZZZ")


        '品測備註 
        Select Case FindFieldInf("QCRemark")
            Case 0  '顯示
                DQCRemark.BackColor = Color.LightGray
                DQCRemark.Visible = True
                DQCRemark.ReadOnly = True
            Case 1  '修改+檢查
                DQCRemark.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCRemarkRqd", "DQCRemark", "異常：需輸入品測備註")
                DQCRemark.Visible = True
            Case 2  '修改
                DQCRemark.BackColor = Color.Yellow
                DQCRemark.Visible = True
            Case Else   '隱藏
                DQCRemark.Visible = False
        End Select
        If pPost = "New" Then DQCRemark.Text = ""


        '製造流程表
        Select Case FindFieldInf("ManufFlowFile")

            Case 0  '顯示
                DManufFlowFile.Visible = False
                DManufFlowFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DManufFlowFileRqd", "DManufFlowFile", "異常：需輸入製造流程表")
                DManufFlowFile.Visible = True
                DManufFlowFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DManufFlowFile.Visible = True
                DManufFlowFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DManufFlowFile.Visible = False
        End Select
        LManufFlowFile.Visible = False


        '作業標準書
        Select Case FindFieldInf("OPManualFile")

            Case 0  '顯示
                DOPManualFile.Visible = False
                DOPManualFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DOPManualFileRqd", "DOPManualFile", "異常：需輸入作業標準書")
                DOPManualFile.Visible = True
                DOPManualFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DOPManualFile.Visible = True
                DOPManualFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DOPManualFile.Visible = False
        End Select
        LOPManualFile.Visible = False






        '報價單
        Select Case FindFieldInf("ForcastFile")
            Case 0  '顯示
                DForcastFile.Visible = False
                DForcastFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DForcastFile.Attributes.Add("readonly", "true")
                LForcastFile.Visible = True
                LForcastFile.BackColor = Color.LightGray
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DForcastFileRqd", "DForcastFile", "異常：需輸入報價單")
                DForcastFile.Visible = True
                DForcastFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DForcastFile.Visible = True
                DForcastFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DForcastFile.Visible = False
        End Select
        LForcastFile.Visible = False


        '切結書
        Select Case FindFieldInf("ContactFile")
            Case 0  '顯示
                DContactFile.Visible = False
                DContactFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DContactFile.Attributes.Add("readonly", "true")
                LContactFile.Visible = True
                LContactFile.BackColor = Color.LightGray

            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DContactFileRqd", "DContactFile", "異常：需輸入切結書")
                DContactFile.Visible = True
                DContactFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DContactFile.Visible = True
                DContactFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DContactFile.Visible = False
        End Select
        LContactFile.Visible = False




    End Sub
    '*****************************************************************
    '**(ShowSheetField)
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator

        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.ErrorMessage = pMessage
        rqdVal.Display = ValidatorDisplay.None              ' 因在頁面上加入ValidationSummary , 故驗證控制項統一顯示
        rqdVal.Style.Add("Top", Top + 300 & "px")
        rqdVal.Style.Add("Left", "8")
        rqdVal.Style.Add("Height", "20")
        rqdVal.Style.Add("Width", "250")
        rqdVal.Style.Add("POSITION", "absolute")
        Me.Form.Controls.Add(rqdVal)
    End Sub

    '*****************************************************************
    '**(ShowSheetField)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim DBDataSet1 As New DataSet

        Dim SQL As String
        Dim idx As Integer
        Dim i As Integer


        idx = FindFieldInf(pFieldName)

        '擔當者及部門 

        SQL = "Select Divname,Username From M_Users "
        SQL = SQL & " Where UserID = '" & wApplyID & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBUser As DataTable = uDataBase.GetDataTable(SQL)
        DAppper.Text = DBUser.Rows(0).Item("Username")
        DDivision.Text = DBUser.Rows(0).Item("Divname")

        'buyer
        If pFieldName = "Buyer" Then
            DBuyer.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBuyer.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'Buyer' "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DBuyer.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBuyer.Items.Add(ListItem1)
                Next
            End If
        End If

        '附樣
        If pFieldName = "AttachSample" Then
            DAttachSample.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAttachSample.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'AttachSample' "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DAttachSample.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAttachSample.Items.Add(ListItem1)
                Next
            End If
        End If

        '樣品需求量
        If pFieldName = "SampleQty" Then
            DSampleQty.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSampleQty.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'SampleQty' "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DSampleQty.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSampleQty.Items.Add(ListItem1)
                Next
            End If
        End If




        '外注廠商
        If pFieldName = "Supplier" Then
            DSupplier.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSupplier.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'Supplier'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DSupplier.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSupplier.Items.Add(ListItem1)
                Next
            End If
        End If



        '限度樣品
        If pFieldName = "AllowSample" Then
            DAllowSample.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAllowSample.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'AllowSample'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DAllowSample.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAllowSample.Items.Add(ListItem1)
                Next
            End If
        End If


        'QC 檢測結果
        If pFieldName = "QCResult" Then
            DQCResult.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCResult.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'QCResult'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DQCResult.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCResult.Items.Add(ListItem1)
                Next
            End If
        End If





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
                SQL = "Select * From M_Referp Where Cat='014' Order by DKey, Data "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("DKey")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("DKey")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DReasonCode.Items.Add(ListItem1)
                Next
            End If
        End If

        '延遲理由
        If pFieldName = "Reason" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DReason.Text = pName
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='014' and DKey= '" & DReasonCode.SelectedValue & "' Order by Data "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

                For i = 0 To DBAdapter1.Rows.Count - 1
                    DReason.Text = DBAdapter1.Rows(i).Item("Data")
                Next
            End If
        End If



    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
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
    '**
    '**     儲存按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BSAVE_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSAVE.ServerClick
        If Request.Cookies("RunBSAVE").Value = True Then
            ' If InputCheck() = 0 Then
            DisabledButton()   '停止Button運作
            Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
            Dim Message As String = ""

            'Check理由
            If ErrCode = 0 Then
                If DReasonCode.Visible = True Then
                    If DReasonCode.SelectedValue = "99" Then
                        If DReasonDesc.Text = "" Then ErrCode = 9210
                    End If
                End If
            End If
            '儲存資料
            If ErrCode = 0 Then
                ModifyData("SAVE", "0")           '更新表單資料 Sts=0(未結)
                ModifyTranData("SAVE", "0")       '更新交易資料
            Else
                If ErrCode = 9210 Then Message = "延遲理由為其他時需填寫說明,請確認!"
                Response.Write(YKK.ShowMessage(Message))
            End If      '上傳檔案ErrCode=0

            If ErrCode = 0 Then

                Dim URL As String = uCommon.GetAppSetting("RedirectURL") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                    "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
            End If
            'End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BOK_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
        If Request.Cookies("RunBOK").Value = True Then
            DisabledButton()   '停止Button運作
            FlowControl("OK", 0, "1")
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG1按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BNG1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG1.ServerClick
        If Request.Cookies("RunBNG1").Value = True Then
            DisabledButton()   '停止Button運作
            FlowControl("NG1", 1, "2")
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG2按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BNG2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG2.ServerClick
        If Request.Cookies("RunBNG2").Value = True Then
            DisabledButton()   '停止Button運作
            FlowControl("NG2", 2, "3")
        End If
    End Sub

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

        GetDataStatus()  '取得此關卡各按鈕的Data Status



        'Check延遲理由
        If ErrCode = 0 Then
            If DReasonCode.Visible = True Then
                If DReasonCode.SelectedValue = "99" Then
                    If DReasonDesc.Text = "" Then ErrCode = 9040
                End If
            End If
        End If

        '--檢查委託書No---------
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                ErrCode = oCommon.CommissionNo("003002", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9060
                End If
            End If
        End If

        If ErrCode = 0 Then
            Dim Run As Boolean = True           '是否執行
            Dim RepeatRun As Boolean = False    '是否重覆執行
            Dim MultiJob As Integer = 0         '多人核定
            Dim wLevel As String = ""           '樣品難易度

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
                Dim i, RtnCode As Integer
                Dim NewFormSno As Integer = wFormSno    '表單流水號
                Dim pRunNextStep As Integer = 0         '是否執行計算下一關(會簽)
                Dim SQL As String

                '取得表單流水號或更新交易資料
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then      '判斷是否起單
                        '取得表單流水號
                        RtnCode = oCommon.GetFormSeqNo(wFormNo, NewFormSno) '表單號碼, 表單流水號
                        If RtnCode <> 0 Then
                            ErrCode = 9110
                        Else

                            '申請流程資料建置
                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                            '
                            oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                            '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者
                            'Modify-End
                        End If
                        pRunNextStep = 1
                    Else
                        If RepeatRun = False Then   '不是通知的重覆執行
                            '更新交易資料
                            ModifyTranData(pFun, pSts)

                            '流程資料結束
                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDepo, Request.QueryString("pUserID"), pRunNextStep)
                            '
                            RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, Request.QueryString("pUserID"), pRunNextStep)
                            '表單號碼,表單流水號,工程關卡號碼,行事曆,簽核者, 流程結束否(會簽)
                            'Modify-End

                            If RtnCode <> 0 Then ErrCode = 9120
                        Else
                            pRunNextStep = 1    '是通知的重覆執行
                        End If
                    End If
                End If

                '取得下一關
                If ErrCode = 0 And pRunNextStep = 1 Then

                    Dim wAllocateID As String = ""
                    RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, wAgentID, wAllocateID, MultiJob, _
                                                  pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
                    '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
                    '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 

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

                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)
                            'oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wDepo)
                            'RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wDepo, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '
                            '取得核定者-群組行是曆
                            wNextGateCalendar = oCommon.GetCalendarGroup(pNextGate(i))

                            '取得工程負荷最後日時
                            oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)
                            '核定者, 表單號碼, 工程號碼, 類別(0:通知,1:核定), 開始日時, 最後日時, 件數

                            '取得預定開始,完成日程計算
                            oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wNextGateCalendar)
                            '表單號碼,工程號碼,難易度,QC-L/T,現在時間, 預定開始日時, 預定完成日時, 行事曆

                            '建置流程資料
                            RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '表單號碼,表單流水號,工程關卡號碼,序號,申請者ID,行事曆,建置者, 簽核者, 被代理者, 申請者, 預定開始日, 預定完成日, 重要性
                            'Modify-End

                            If RtnCode <> 0 Then
                                ErrCode = 9150
                                Exit For
                            End If
                        Next i
                    Else
                        'Modify-Start by 2009/11/20(2010行事曆對應)
                        'RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                        '
                        RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者
                        'Modify-End

                        If RtnCode <> 0 Then ErrCode = 9160
                    End If
                End If
                '當工程日程調整
                If ErrCode = 0 Then
                    If RepeatRun = False Then
                        'Modify-Start by 2009/11/20(2010行事曆對應)
                        'RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDepo)
                        '
                        RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDecideCalendar)
                        '簽核者,表單號碼,表單流水號,工程關卡號碼,序號,現在日時,難易度,行事曆
                        'Modify-End
                    End If
                End If
                '儲存表單資料
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then      '判斷是否起單
                        If NewFormSno <> 0 Then
                            AppendData(pFun, NewFormSno)  'Insert
                            AddCommissionNo(wFormNo, NewFormSno)
                        End If  'pSeqno <> 0
                    Else    '判斷是否起單
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
                    Response.Write(YKK.ShowMessage(Message))
                End If      '儲存表單ErrCode=0
            End While       '重覆執行

            If ErrCode = 0 Then
                '--郵件傳送---------
                oCommon.SendMail()
                Dim URL As String = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                                 "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
            End If
        Else
            EnabledButton()   '起動Button運作
            If ErrCode = 9010 Then Message = "上傳檔案格式錯誤,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案Size超過1024KB,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案沒指定,請確認!"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9050 Then Message = "材質為其他時需填寫說明,請確認!"
            If ErrCode = 9060 Then Message = "委託書No.重覆,請確認委託書No.!"
            If ErrCode = 9070 Then Message = "需填入工程待處理部門及工程待處理者.!"
            Response.Write(YKK.ShowMessage(Message))
        End If      '上傳檔案ErrCode=0
    End Sub

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
    '**(AppendData)
    '**     新增表單資料
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)
        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                   System.Configuration.ConfigurationManager.AppSettings("SBDSufaceFilePath")

        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String


        SQl = "Insert into F_SBDSufaceSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, "
        SQl = SQl + "No, Appdate, Division,APPPER,Buyer,"                  '1~5
        SQl = SQl + "SellVendor,ReqDelDate,ReqQty,Code,AttachSample, "               '6~10
        SQl = SQl + "ORNO,OrderTime,DevReason,Color,SampleQty, "                 '11~15
        SQl = SQl + "EnglishName,Supplier,AllowSample,Cap,BFinalDate, "   '16~20
        SQl = SQl + "Schedule,QCReqFile,QAFile,QCDate,QCResult, "                                                  '21~25
        SQl = SQl + "QCRemark,ManufFlowFile,OPManualFile,ForcastFile,ContactFile, "            '26~30
        SQl = SQl + "CustSampleFile,FinalSampleFile, "
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "


        SQl = SQl + ")  "

        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        SQl = SQl + " '" + NowDateTime + "', "        '結案日
        SQl = SQl + " '003002', "                     '表單代號
        SQl = SQl + " '" + CStr(NewFormSno) + "', "   '表單流水號

        SQl = SQl + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO
        SQl = SQl + " '" + DAppDate.Text + "', "                '日期
        SQl = SQl + " '" + DDivision.Text + "', "   '部門
        SQl = SQl + " '" + DAppper.Text + "', "     '擔當
        SQl = SQl + " '" + DBuyer.SelectedValue + "', "  'buyer

        SQl = SQl + " '" + DSellVendor.Text + "', "    '委託廠商
        SQl = SQl + " '" + DReqDelDate.Value + "', "      '希望交期
        SQl = SQl + " N'" + YKK.ReplaceString(DReqQty.Text) + "', "   '大貨預估量
        SQl = SQl + " '" + DCode.Text + "', "     '品名
        SQl = SQl + " '" + DAttachSample.SelectedValue + "', "  '附樣


        SQl = SQl + " '" + DORNO.Text + "', "    'ORNO
        SQl = SQl + " '" + DOrderTime.Value + "',"   '下單時間
        SQl = SQl + " '" + DDevReason.Text + "', "   '開發理由
        SQl = SQl + " '" + DColor.Text + "', "   'color code
        SQl = SQl + " '" + DSampleQty.SelectedValue + "', "   '樣品需求量


        SQl = SQl + " '" + DEnglishName.Text + "', "    '英文名稱
        SQl = SQl + " '" + DSupplier.SelectedValue + "',"   '外注廠商
        SQl = SQl + " '" + DAllowSample.SelectedValue + "', "   '限度樣品
        SQl = SQl + " '" + DCap.Text + "', "   '日產能
        SQl = SQl + " '" + DBFinalDate.Value + "', "   '預定完成日



        SQl = SQl + " '" + DSchedule.Text + "', "    '基準日程
 
        FileName = ""
        If DQCReqFile.Visible Then
            If DQCReqFile.PostedFile.FileName <> "" Then
                ' Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                '                              CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
                'Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                'System.Configuration.ConfigurationManager.AppSettings("TimeOffSheetPath")

                FileName = CStr(NewFormSno) & "-" & "QCReqFile" & "-" & UploadDateTime & "-" & Right(DQCReqFile.PostedFile.FileName, InStr(StrReverse(DQCReqFile.PostedFile.FileName), "\") - 1)

                DQCReqFile.PostedFile.SaveAs(Path + FileName)

            Else
                FileName = ""
            End If
        End If
        SQl = SQl + " '" + FileName + "'," '品質依賴書 


        FileName = ""
        If DQAFile.Visible Then
            If DQAFile.PostedFile.FileName <> "" Then
                ' Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                '                              CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
                'Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                'System.Configuration.ConfigurationManager.AppSettings("TimeOffSheetPath")

                FileName = CStr(NewFormSno) & "-" & "QAFile" & "-" & UploadDateTime & "-" & Right(DQAFile.PostedFile.FileName, InStr(StrReverse(DQAFile.PostedFile.FileName), "\") - 1)

                DQAFile.PostedFile.SaveAs(Path + FileName)

            Else
                FileName = ""
            End If
        End If
        SQl = SQl + " '" + FileName + "'," '品測報告書 

        SQl = SQl + " '" + DQCDate.Value + "', "    '測試日期
        SQl = SQl + " '" + DQCResult.SelectedValue + "', "    '檢測結果 



        SQl = SQl + " '" + DQCRemark.Text + "', "    '檢測備註 
        FileName = ""
        If DManufFlowFile.Visible Then
            If DManufFlowFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ManufFlowFile" & "-" & UploadDateTime & "-" & Right(DManufFlowFile.PostedFile.FileName, InStr(StrReverse(DManufFlowFile.PostedFile.FileName), "\") - 1)
                DManufFlowFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If
        SQl = SQl + " '" + FileName + "'," '製造流程表



        FileName = ""
        If DOPManualFile.Visible Then
            If DOPManualFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "OPManualFile" & "-" & UploadDateTime & "-" & Right(DOPManualFile.PostedFile.FileName, InStr(StrReverse(DOPManualFile.PostedFile.FileName), "\") - 1)
                DOPManualFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If
        SQl = SQl + " '" + FileName + "'," '作業標準書


        FileName = ""
        If DForcastFile.Visible Then
            If DForcastFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ForcastFile" & "-" & UploadDateTime & "-" & Right(DForcastFile.PostedFile.FileName, InStr(StrReverse(DForcastFile.PostedFile.FileName), "\") - 1)
                DForcastFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If
        SQl = SQl + " '" + FileName + "',"     '報價單

     
        FileName = ""
        If DContactFile.Visible Then
            If DContactFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ContactFile" & "-" & UploadDateTime & "-" & Right(DContactFile.PostedFile.FileName, InStr(StrReverse(DContactFile.PostedFile.FileName), "\") - 1)
                DContactFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '切結書

        FileName = ""
        If DCustSampleFile.Visible Then
            If DCustSampleFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "CustSampleFile" & "-" & UploadDateTime & "-" & Right(DCustSampleFile.PostedFile.FileName, InStr(StrReverse(DCustSampleFile.PostedFile.FileName), "\") - 1)
                DCustSampleFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '顧客樣品圖


        FileName = ""
        If DFinalSampleFile.Visible Then
            If DFinalSampleFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "FinalSampleFile" & "-" & UploadDateTime & "-" & Right(DFinalSampleFile.PostedFile.FileName, InStr(StrReverse(DFinalSampleFile.PostedFile.FileName), "\") - 1)
                DFinalSampleFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '最終樣品圖



        SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '作成者
        SQl = SQl + " '" + NowDateTime + "', "       '作成時間
        SQl = SQl + " '" + "" + "', "                       '修改者
        SQl = SQl + " '" + NowDateTime + "' "       '修改時間
        SQl = SQl + " ) "
        uDataBase.ExecuteNonQuery(SQl)

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
               System.Configuration.ConfigurationManager.AppSettings("SBDSufaceFilePath")
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String


        SQl = "Update F_SBDSufaceSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"
        SQl = SQl + " AppDate = '" & DAppDate.Text & "',"
        SQl = SQl + " Division = '" & DDivision.Text & "',"
        SQl = SQl + " APPPER = '" & DAppper.Text & "',"
        SQl = SQl + " Buyer = '" & DBuyer.SelectedValue & "',"
        SQl = SQl + " SellVendor = '" & DSellVendor.Text & "',"

        SQl = SQl + " ReqDelDate = '" & DReqDelDate.Value & "',"
        SQl = SQl + " ReqQty = '" & DReqQty.Text & "',"

        SQl = SQl + " Code = '" & DCode.Text & "',"
        SQl = SQl + " AttachSample = '" & DAttachSample.SelectedValue & "',"
        SQl = SQl + " ORNO = '" & DORNO.Text & "',"
        SQl = SQl + " DevReason = '" & DDevReason.Text & "',"
        SQl = SQl + " Color = '" & DColor.Text & "',"
        SQl = SQl + " SampleQty = '" & DSampleQty.Text & "',"
        SQl = SQl + " EnglishName= '" & DEnglishName.Text & "',"

        SQl = SQl + " Supplier= '" & DSupplier.SelectedValue & "',"
        SQl = SQl + " AllowSample= '" & DAllowSample.Text & "',"
        SQl = SQl + " Cap = '" & DCap.Text & "',"

        SQl = SQl + " BFinalDate= '" & DBFinalDate.Value & "',"
        SQl = SQl + " Schedule= '" & DSchedule.Text & "',"


        FileName = ""
        If DQCReqFile.Visible Then
            If DQCReqFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "QCReqFile" & "-" & UploadDateTime & "-" & Right(DQCReqFile.PostedFile.FileName, InStr(StrReverse(DQCReqFile.PostedFile.FileName), "\") - 1)
                DQCReqFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " QCReqFile= N'" + YKK.ReplaceString(FileName) + "',"           '品質依賴書
            Else
                FileName = ""
            End If
        End If


        FileName = ""
        If DQAFile.Visible Then
            If DQAFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "QAFile" & "-" & UploadDateTime & "-" & Right(DQAFile.PostedFile.FileName, InStr(StrReverse(DQAFile.PostedFile.FileName), "\") - 1)
                DQAFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " QAFile= N'" + YKK.ReplaceString(FileName) + "',"           '品測報告書
            Else
                FileName = ""
            End If
        End If




        SQl = SQl + " QCDate= '" & DQCDate.Value & "',"                      '檢測日期
        SQl = SQl + " QCResult= '" & DQCResult.SelectedValue & "',"                               '檢測結果
        SQl = SQl + " QCRemark= '" & DQCRemark.Text & "',"                                 '檢測備註

        FileName = ""
        If DManufFlowFile.Visible Then
            If DManufFlowFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "ManufFlowFile" & "-" & UploadDateTime & "-" & Right(DManufFlowFile.PostedFile.FileName, InStr(StrReverse(DManufFlowFile.PostedFile.FileName), "\") - 1)
                DManufFlowFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " ManufFlowFile= N'" + YKK.ReplaceString(FileName) + "',"           '製造流程圖
            Else
                FileName = ""
            End If
        End If


        FileName = ""
        If DOPManualFile.Visible Then
            If DOPManualFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "OPManualFile" & "-" & UploadDateTime & "-" & Right(DOPManualFile.PostedFile.FileName, InStr(StrReverse(DOPManualFile.PostedFile.FileName), "\") - 1)
                DOPManualFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " OPManualFile= N'" + YKK.ReplaceString(FileName) + "',"           '作業標準書
            Else
                FileName = ""
            End If
        End If




        FileName = ""
        If DForcastFile.Visible Then
            If DForcastFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "ForcastFile" & "-" & UploadDateTime & "-" & Right(DForcastFile.PostedFile.FileName, InStr(StrReverse(DForcastFile.PostedFile.FileName), "\") - 1)
                DForcastFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " ForcastFile= N'" + YKK.ReplaceString(FileName) + "',"                 '報價單
            Else
                FileName = ""
            End If
        End If



        FileName = ""
        If DContactFile.Visible Then
            If DContactFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "ContactFile" & "-" & UploadDateTime & "-" & Right(DContactFile.PostedFile.FileName, InStr(StrReverse(DContactFile.PostedFile.FileName), "\") - 1)
                DContactFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " ContactFile= N'" + YKK.ReplaceString(FileName) + "',"                     '切結書
            Else
                FileName = ""
            End If
        End If


        FileName = ""
        If DCustSampleFile.Visible Then
            If DCustSampleFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "CustSampleFile" & "-" & UploadDateTime & "-" & Right(DCustSampleFile.PostedFile.FileName, InStr(StrReverse(DCustSampleFile.PostedFile.FileName), "\") - 1)
                DCustSampleFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " CustSampleFile= N'" + YKK.ReplaceString(FileName) + "',"
            Else
                FileName = ""
            End If
        End If



        FileName = ""
        If DFinalSampleFile.Visible Then
            If DFinalSampleFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "FinalSampleFile" & "-" & UploadDateTime & "-" & Right(DFinalSampleFile.PostedFile.FileName, InStr(StrReverse(DFinalSampleFile.PostedFile.FileName), "\") - 1)
                DFinalSampleFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " FinalSampleFile= N'" + YKK.ReplaceString(FileName) + "',"
            Else
                FileName = ""
            End If
        End If





        SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
        SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
        SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
        uDataBase.ExecuteNonQuery(SQl)
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyTranData)
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub ModifyTranData(ByVal pFun As String, ByVal pSts As String)
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim SQl As String


        If pFun <> "SAVE" Then      '<> Save
            SQl = "Update T_WaitHandle Set "
            SQl = SQl + " Active = '" & "0" & "',"
            SQl = SQl + " Sts = '" & pSts & "',"
            If pSts = "1" Then SQl = SQl + " StsDes = '" & BOK.Value & "',"
            If pSts = "2" Then SQl = SQl + " StsDes = '" & BNG1.Value & "',"
            If pSts = "3" Then SQl = SQl + " StsDes = '" & BNG2.Value & "',"
            SQl = SQl + " AEndTime = '" & NowDateTime & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
            If DDelay.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(DDecideDesc.Text) & "',"
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
            SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(DDecideDesc.Text) & "',"
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
    '**(Add T_CommissionNo)
    '**     追加交易資料和委託單對照表
    '**
    '*****************************************************************
    Sub AddCommissionNo(ByVal pFormNo As String, ByVal pFormSno As Integer)
        Dim SQl As String


        SQl = "Select * From T_CommissionNo "
        SQl = SQl & " Where FormNo =  '" & pFormNo & "'"
        SQl = SQl & "   And FormSno =  '" & CStr(pFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQl)

        If DBAdapter1.Rows.Count <= 0 Then
            If DNo.Text <> "" Then
                SQl = "Insert into T_CommissionNo "
                SQl = SQl + "( "
                SQl = SQl + "FormNo, FormSno, No, CreateUser, CreateTime "    '1~5
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '" + pFormNo + "', "
                SQl = SQl + " '" + CStr(pFormSno) + "', "
                SQl = SQl + " '" + DNo.Text + "', "
                SQl = SQl + " '" + Request.QueryString("pUserID") + "', "
                SQl = SQl + " '" + NowDateTime + "' "
                SQl = SQl + " ) "
                uDataBase.ExecuteNonQuery(SQl)
            End If
        Else
            If DNo.Text <> "" Then
                If DNo.Text <> DBAdapter1.Rows(0).Item("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNo.Text & "',"
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
    '**
    '**     檢查上傳檔案
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As HtmlInputFile) As Integer
        Dim fileExtension As String     '宣告一個變數存放檔案格式(副檔名)
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".ai", ".xls", ".doc", ".ppt"}   '定義允許的檔案格式
        Dim i As Integer

        UPFileIsNormal = 0

        fileExtension = IO.Path.GetExtension(UPFile.PostedFile.FileName).ToLower   '取得檔案格式
        For i = 0 To allowedExtensions.Length - 1           '逐一檢查允許的格式中是否有上傳的格式
            If fileExtension = allowedExtensions(i) Then
                UPFileIsNormal = 0
                Exit For
            Else
                UPFileIsNormal = 9010
            End If
        Next
        'Check上傳檔案Size
        If UPFileIsNormal = 0 Then
            If UPFile.PostedFile.ContentLength <= 1000 * 1024 Then
                UPFileIsNormal = 0
            Else
                UPFileIsNormal = 9020
            End If
        End If

        'If UPFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
        'Check上傳檔案格式
        'Else
        'UPFileIsNormal = 9030
        'End If
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     停止Button(OK, Ng1, NG2, Save)運作 
    '**
    '*****************************************************************
    Private Sub DisabledButton()
        BOK.Disabled = True
        BNG1.Disabled = True
        BNG2.Disabled = True
        BSAVE.Disabled = True
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     起動Button(OK, Ng1, NG2, Save)運作 
    '**
    '*****************************************************************
    Private Sub EnabledButton()
        BOK.Disabled = False
        BNG1.Disabled = False
        BNG2.Disabled = False
        BSAVE.Disabled = False
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetControlPosition)
    '**    設定按鈕,說明,延遲,核定履歷的位置
    '**
    '*****************************************************************
    Sub SetControlPosition()
        TopPosition()
        If DDescSheet.Visible Then                                      ' 說明
            DDescSheet.Style("top") = Top - 250 & "px"
            DDecideDesc.Style("top") = Top - 250 + 6 & "px"
            Top = Top - 80
        End If

        If DDelay.Visible Then                                          ' 延遲說明
            DDelay.Style("top") = Top & "px"
            DReasonCode.Style("top") = Top + 9 & "px"
            DReason.Style("top") = Top + 6 & "px"
            DReasonDesc.Style("top") = Top + 39 & "px"
            Top += 161
        End If

        BSAVE.Style("top") = Top + 10 & "px"
        BNG1.Style("top") = Top + 10 & "px"
        BNG2.Style("top") = Top + 10 & "px"
        BOK.Style("top") = Top + 10 & "px"

        Top += 48

        If GridView2.Rows.Count > 0 Then                                ' 核定履歷
            DHistoryLabel.Style("top") = Top & "px"
            Top += 20
            GridView2.Style("top") = Top & "px"
        End If
    End Sub



    Protected Sub DNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DNo.TextChanged

    End Sub

    Protected Sub DBFinalDate_ServerChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles DBFinalDate.ServerChange

    End Sub
End Class

