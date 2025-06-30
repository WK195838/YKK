Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb

Partial Class RegisterSheet_01
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
    Dim wCode As String          'ItemCode
    Dim NowDateTime As String       '現在日期時間
    Dim wNextGate As String         '下一關人
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
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

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900      ' 設定逾時時間
        SetParameter()                  ' 設定共用參數
        TopPosition()                   ' 設定top
        SetPopupFunction()              ' 設定彈出視窗事件
        ShowSheetField("New")           ' 表單欄位顯示及欄位輸入檢查
        ShowSheetFunction()             ' 表單功能按鈕顯示
        If Not IsPostBack Then
            'jessica 2021/11/25 
            If wCode <> "" Then
                GetItemCode(wCode)
            End If
            If wFormSno > 0 And wStep > 3 Then    '判斷是否[簽核]
                ShowFormData()          ' 顯示表單資料
                UpdateTranFile()        ' 更新交易資料
            End If
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
            Top = 1050
            Dim SQL As String = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(SQL)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0)("BEndTime").ToString < NowDateTime Then
                    If DDelay.Visible = True Then
                        Top = 1078
                    End If
                End If
            End If
        Else
            Top = 970
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
        wCode = Request.QueryString("pCode")  'Itemcode
      

        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        wDecideCalendar = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))
        '
        Response.Cookies("PGM").Value = "ItemRegisterSheet_01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼


    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     取得ITMECODE 資料
    '**
    '*****************************************************************
    Sub GetItemCode(ByVal pCode As String)

        Dim xCode As String = pCode
        Dim SQL As String
        '
        SQL = "SELECT "
        SQL = SQL + "ITEM, ITEM_NAME1,ITEM_NAME2, ITEM_NAME3, "
        SQL = SQL + "SIZE, CHAINTYPE, CLASS, TAPE_, SLIDER, FINISH, "
        SQL = SQL + "SLIDER2, FINISH2, SPECIAL1, SPECIAL2, SPECIAL3, SPECIAL4, "
        SQL = SQL + "SPECIAL5, SPECIAL6, ST1CA0, ST2CA0, ST3CA0, "
        SQL = SQL + "ST4CA0, ST5CA0, ST6CA0, ST7CA0, NODISP, FMLCA0 "
        SQL = SQL + "From MST_ITEM "
        SQL = SQL + "Where ITEM = '" & xCode & "' "
        Dim dtITEM As DataTable = uDataBase.GetDataTable(SQL)
        If dtITEM.Rows.Count > 0 Then
            DRCode.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ITEM")))
            DRItemName1.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ITEM_NAME1")))
            DRItemName2.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ITEM_NAME2")))
            DRItemName3.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ITEM_NAME3")))
            DRSize.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SIZE")))
            DRChain.Text = LTrim(RTrim(dtITEM.Rows(0).Item("CHAINTYPE")))
            DRClass.Text = LTrim(RTrim(dtITEM.Rows(0).Item("CLASS")))
            DRTape.Text = LTrim(RTrim(dtITEM.Rows(0).Item("TAPE_")))
            DRSlider1.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SLIDER")))
            DRFinish1.Text = LTrim(RTrim(dtITEM.Rows(0).Item("FINISH")))
            DRSlider2.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SLIDER2")))
            DRFinish2.Text = LTrim(RTrim(dtITEM.Rows(0).Item("FINISH2")))
            DRSRequest1.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL1")))
            DRSRequest2.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL2")))
            DRSRequest3.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL3")))
            DRSRequest4.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL4")))
            DRSRequest5.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL5")))
            DRSRequest6.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL6")))
            DRFamily.Text = LTrim(RTrim(dtITEM.Rows(0).Item("FMLCA0")))
            DRST1.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ST1CA0")))
            DRST2.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ST2CA0")))
            DRST3.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ST3CA0")))
            DRST4.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ST4CA0")))
            DRST5.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ST5CA0")))
            DRST6.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ST6CA0")))
            DRST7.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ST7CA0")))
            DRNoDisplay.Text = LTrim(RTrim(dtITEM.Rows(0).Item("NODISP")))

            DSize.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SIZE")))
            DChain.Text = LTrim(RTrim(dtITEM.Rows(0).Item("CHAINTYPE")))
            DClass.Text = LTrim(RTrim(dtITEM.Rows(0).Item("CLASS")))
            DTape.Text = LTrim(RTrim(dtITEM.Rows(0).Item("TAPE_")))
            DSlider1.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SLIDER")))
            DFinish1.Text = LTrim(RTrim(dtITEM.Rows(0).Item("FINISH")))
            DSlider2.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SLIDER2")))
            DFinish2.Text = LTrim(RTrim(dtITEM.Rows(0).Item("FINISH2")))
            DSRequest1.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL1")))
            DSRequest2.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL2")))
            DSRequest3.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL3")))
            DSRequest4.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL4")))
            DSRequest5.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL5")))
            DSRequest6.Text = LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL6")))
            DFamily.Text = LTrim(RTrim(dtITEM.Rows(0).Item("FMLCA0")))
            DST1.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ST1CA0")))
            DST2.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ST2CA0")))
            DST3.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ST3CA0")))
            DST4.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ST4CA0")))
            DST5.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ST5CA0")))
            DST6.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ST6CA0")))
            DST7.Text = LTrim(RTrim(dtITEM.Rows(0).Item("ST7CA0")))
        End If
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
            Else
                BOK.Visible = False
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
                Top = 1050
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
            Top = 970
            'Sheet隱藏
            DDescSheet.Visible = False  '說明Sheet
            DDelay.Visible = False      '延遲Sheet
            '欄位隱藏
            DDecideDesc.Visible = False    '說明
            DReasonCode.Visible = False    '延遲理由代碼
            DReason.Visible = False        '延遲理由
            DReasonDesc.Visible = False    '延遲其他說明
            DHistoryLabel.Visible = False  '核定履歷
        End If
        '按鈕及超連結設值
        If Not IsPostBack Then
            LAttachfile1.Visible = False
        End If
        BSAVE.Style("top") = Top & "px"
        BNG1.Style("top") = Top & "px"
        BNG2.Style("top") = Top & "px"
        BOK.Style("top") = Top & "px"
        DHistoryLabel.Style("top") = Top + 32 & "px"
        GridView2.Style("top") = Top + 32 + 16 & "px"
        '簽核,通知 移除控制設定
        If (wFormSno > 0 And wStep > 3) And wStep <> 500 Then    '判斷是否[簽核]
            DStatus.Value = "OK"
            BFindItem.Visible = False
            BCheckItem.Visible = False
            DMark1.Visible = False
        End If
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

        Dim SQL1 As String
        Dim InputData1 As String
        Dim InputData2 As String


        InputData1 = D1.Text
        InputData2 = D2.Text


        '帶入CUSTOMER
        If InputData1 <> "" Then

            SQL1 = " Select  * from  MST_Custmer where 1=1 "
            If InputData1 <> "" Then
                SQL1 = SQL1 + " and ( custmer like '%" + InputData1 + "%' or name_c like '%" + InputData1 + "%')"
            End If
            SQL1 = SQL1 + " order by custmer,name_c "

            Dim DBData As DataTable = uDataBase.GetDataTable(SQL1)

            DCustomer.Text = DBData.Rows(0).Item("Name_C")
            DCustomerCode.Text = DBData.Rows(0).Item("Custmer")
        End If

        '帶入BUYER
        If InputData2 <> "" Then


            SQL1 = " Select  * from  MST_Buyer where 1=1 "
            If InputData2 <> "" Then
                SQL1 = SQL1 + " and ( buyer like '%" + InputData2 + "%' or buyer_name like '%" + InputData2 + "%')"
            End If

            SQL1 = SQL1 + " order by buyer_name,buyer "
            Dim DBData As DataTable = uDataBase.GetDataTable(SQL1)

            DBuyer.Text = DBData.Rows(0).Item("Buyer_Name")
            DBuyerCode.Text = DBData.Rows(0).Item("Buyer")

            DMemo.Text = ""
            DA999.Checked = False
            DA001.Checked = False
            If DBuyer.Text <> "" Then
                '20190722:   JESSICA()
                If DST4.Text = "1" Then
                    SQL1 = "SELECT  * FROM m_REFERP"
                    SQL1 = SQL1 + " WHERE CAT = '1151'"
                    SQL1 = SQL1 + " AND DKEY  in ('A999','A001')"
                    SQL1 = SQL1 + " AND DATA ='" + DBuyerCode.Text + "'"
                    Dim dtPrice As DataTable = uDataBase.GetDataTable(SQL1)
                    If dtPrice.Rows.Count > 0 Then
                        If dtPrice.Rows(0)("DKEY").ToString = "A999" Then
                            DA999.Checked = True
                            DA001.Checked = False
                        ElseIf dtPrice.Rows(0)("DKEY").ToString = "A001" Then
                            DA001.Checked = True
                            DA999.Checked = False
                        End If

                        DMemo.Text = "請確認若為開發中產品(SLS-ANY、CH-ANY、TAPE-ANY、ILYTEMP、I-YTEMP、ILTP***、I-TP***)不可勾選A001或A999。"
                        '  uJavaScript.PopMsg(Me, "請確認若為開發中產品(SLS-ANY、CH-ANY、TAPE-ANY、ILYTEMP、I-YTEMP、ILTP***、I-TP***)不可勾選A001或A999。")
                    Else
                        DA999.Checked = False
                        DA001.Checked = False
                    End If
                Else
                    DA999.Checked = False
                    DA001.Checked = False

                    DMemo.Text = ""
                End If
            End If
            D2.Text = ""
        End If

        If pPost = "New" Then
            DDate.Text = CDate(NowDateTime).ToString("yyyy/MM/dd")
            Dim sql As String = "Select e.Com_Code,e.Com_Name,e.Dep_Code,e.Dep_Name,e.[ID],e.[Name],e.Job_Title_Code,e.Job_Title from M_Users u inner join M_Emp e on u.EmpID = e.[ID] and u.DepoID = e.Com_Code where u.UserID='" & wApplyID & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(sql)
            '取得申請者資訊
            If dt.Rows.Count > 0 Then
                DDivision.Text = dt.Rows(0)("Dep_Name").ToString.Trim
                DName.Text = dt.Rows(0)("Name").ToString.Trim
                DJobTitle.Text = dt.Rows(0)("Job_Title").ToString.Trim
            End If
        End If

        'No
        Select Case FindFieldInf("No")
            Case 0  '顯示
                DNo.BackColor = Color.White
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

        'Name
        Select Case FindFieldInf("Name")
            Case 0  '顯示
                DName.BackColor = Color.LightGray
                DName.Visible = True
                DName.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DName.Visible = True
                DName.BackColor = Color.GreenYellow
                DName.ReadOnly = False
                ShowRequiredFieldValidator("DNameRqd", "DName", "異常：需輸入姓名")
            Case 2  '修改
                DName.Visible = True
                DName.BackColor = Color.Yellow
                DName.ReadOnly = False
            Case Else   '隱藏
                DName.Visible = False
        End Select

        'Division
        'Select Case FindFieldInf("Division")
        Select Case FindFieldInf("Name")
            Case 0  '顯示
                DDivision.BackColor = Color.LightGray
                DDivision.Visible = True
                DDivision.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DDivision.Visible = True
                DDivision.BackColor = Color.GreenYellow
                DDivision.ReadOnly = False
                ShowRequiredFieldValidator("DDivisionRqd", "DDivision", "異常：需輸入部門")
            Case 2  '修改
                DDivision.Visible = True
                DDivision.BackColor = Color.Yellow
                DDivision.ReadOnly = False
            Case Else   '隱藏
                DDivision.Visible = False
        End Select

        'JobTitle
        'Select Case FindFieldInf("JobTitle")
        Select Case FindFieldInf("Name")
            Case 0  '顯示
                DJobTitle.BackColor = Color.LightGray
                DJobTitle.Visible = True
                DJobTitle.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DJobTitle.Visible = True
                DJobTitle.BackColor = Color.GreenYellow
                DJobTitle.ReadOnly = False
                ShowRequiredFieldValidator("DJobTitleRqd", "DJobTitle", "異常：需輸入職稱")
            Case 2  '修改
                DJobTitle.Visible = True
                DJobTitle.BackColor = Color.Yellow
                DJobTitle.ReadOnly = False
            Case Else   '隱藏
                DJobTitle.Visible = False
        End Select

        'SPDPerson
        Select Case FindFieldInf("SPDPerson")
            Case 0  '顯示
                DSPDPerson.BackColor = Color.LightGray
                DSPDPerson.Visible = True
            Case 1  '修改+檢查
                DSPDPerson.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSPDPersonRqd", "DSPDPerson", "異常：需輸入開發擔當")
                DSPDPerson.Visible = True
            Case 2  '修改
                DSPDPerson.BackColor = Color.Yellow
                DSPDPerson.Visible = True
            Case Else   '隱藏
                DSPDPerson.Visible = False
        End Select
        If Not IsPostBack Then SetFieldData("SPDPerson", "ZZZZZZ")

        'ReRegister
        Select Case FindFieldInf("ReRegister")
            Case 0  '顯示
                DReRegister.BackColor = Color.LightGray
                DReRegister.Visible = True
            Case 1  '修改+檢查
                DReRegister.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DReRegisterRqd", "DReRegister", "異常：需輸入繼續申請")
                DReRegister.Visible = True
            Case 2  '修改
                DReRegister.BackColor = Color.Yellow
                DReRegister.Visible = True
            Case Else   '隱藏
                DReRegister.Visible = False
        End Select

        'RCode
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRCode.BackColor = Color.LightGray
                DRCode.Visible = True
                DRCode.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRCode.Visible = True
                DRCode.BackColor = Color.GreenYellow
                DRCode.ReadOnly = False
                ShowRequiredFieldValidator("DRCodeRqd", "DRCode", "異常：需輸入Code")
            Case 2  '修改
                DRCode.Visible = True
                DRCode.BackColor = Color.Yellow
                DRCode.ReadOnly = False
            Case Else   '隱藏
                DRCode.Visible = False
        End Select

        'RItemName1
        'Select Case FindFieldInf("RItemName1")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRItemName1.BackColor = Color.LightGray
                DRItemName1.Visible = True
                DRItemName1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRItemName1.Visible = True
                DRItemName1.BackColor = Color.GreenYellow
                DRItemName1.ReadOnly = False
                ShowRequiredFieldValidator("DRItemName1Rqd", "DRItemName1", "異常：需輸入Item Name-1")
            Case 2  '修改
                DRItemName1.Visible = True
                DRItemName1.BackColor = Color.Yellow
                DRItemName1.ReadOnly = False
            Case Else   '隱藏
                DRItemName1.Visible = False
        End Select

        'RItemName2
        'Select Case FindFieldInf("RItemName2")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRItemName2.BackColor = Color.LightGray
                DRItemName2.Visible = True
                DRItemName2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRItemName2.Visible = True
                DRItemName2.BackColor = Color.GreenYellow
                DRItemName2.ReadOnly = False
                ShowRequiredFieldValidator("DRItemName2Rqd", "DRItemName2", "異常：需輸入Item Name-2")
            Case 2  '修改
                DRItemName2.Visible = True
                DRItemName2.BackColor = Color.Yellow
                DRItemName2.ReadOnly = False
            Case Else   '隱藏
                DRItemName2.Visible = False
        End Select

        'RItemName3
        'Select Case FindFieldInf("RItemName3")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRItemName3.BackColor = Color.LightGray
                DRItemName3.Visible = True
                DRItemName3.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRItemName3.Visible = True
                DRItemName3.BackColor = Color.GreenYellow
                DRItemName3.ReadOnly = False
                ShowRequiredFieldValidator("DRItemName3Rqd", "DRItemName3", "異常：需輸入Item Name-3")
            Case 2  '修改
                DRItemName3.Visible = True
                DRItemName3.BackColor = Color.Yellow
                DRItemName3.ReadOnly = False
            Case Else   '隱藏
                DRItemName3.Visible = False
        End Select

        'RSize
        'Select Case FindFieldInf("RSize")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRSize.BackColor = Color.LightGray
                DRSize.Visible = True
                DRSize.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRSize.Visible = True
                DRSize.BackColor = Color.GreenYellow
                DRSize.ReadOnly = False
                ShowRequiredFieldValidator("DRSizeRqd", "DRSize", "異常：需輸入Size")
            Case 2  '修改
                DRSize.Visible = True
                DRSize.BackColor = Color.Yellow
                DRSize.ReadOnly = False
            Case Else   '隱藏
                DRSize.Visible = False
        End Select

        'RChain
        'Select Case FindFieldInf("RChain")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRChain.BackColor = Color.LightGray
                DRChain.Visible = True
                DRChain.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRChain.Visible = True
                DRChain.BackColor = Color.GreenYellow
                DRChain.ReadOnly = False
                ShowRequiredFieldValidator("DRChainRqd", "DRChain", "異常：需輸入Chain")
            Case 2  '修改
                DRChain.Visible = True
                DRChain.BackColor = Color.Yellow
                DRChain.ReadOnly = False
            Case Else   '隱藏
                DRChain.Visible = False
        End Select

        'RClass
        'Select Case FindFieldInf("RClass")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRClass.BackColor = Color.LightGray
                DRClass.Visible = True
                DRClass.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRClass.Visible = True
                DRClass.BackColor = Color.GreenYellow
                DRClass.ReadOnly = False
                ShowRequiredFieldValidator("DRClassRqd", "DRClass", "異常：需輸入Class")
            Case 2  '修改
                DRClass.Visible = True
                DRClass.BackColor = Color.Yellow
                DRClass.ReadOnly = False
            Case Else   '隱藏
                DRClass.Visible = False
        End Select

        'RTape
        'Select Case FindFieldInf("RTape")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRTape.BackColor = Color.LightGray
                DRTape.Visible = True
                DRTape.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRTape.Visible = True
                DRTape.BackColor = Color.GreenYellow
                DRTape.ReadOnly = False
                ShowRequiredFieldValidator("DRTapeRqd", "DRTape", "異常：需輸入Tape")
            Case 2  '修改
                DRTape.Visible = True
                DRTape.BackColor = Color.Yellow
                DRTape.ReadOnly = False
            Case Else   '隱藏
                DRTape.Visible = False
        End Select

        'RSlider1
        'Select Case FindFieldInf("RSlider1")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRSlider1.BackColor = Color.LightGray
                DRSlider1.Visible = True
                DRSlider1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRSlider1.Visible = True
                DRSlider1.BackColor = Color.GreenYellow
                DRSlider1.ReadOnly = False
                ShowRequiredFieldValidator("DRSlider1Rqd", "DRSlider1", "異常：需輸入Slider-1")
            Case 2  '修改
                DRSlider1.Visible = True
                DRSlider1.BackColor = Color.Yellow
                DRSlider1.ReadOnly = False
            Case Else   '隱藏
                DRSlider1.Visible = False
        End Select

        'RFinish1
        'Select Case FindFieldInf("RFinish1")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRFinish1.BackColor = Color.LightGray
                DRFinish1.Visible = True
                DRFinish1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRFinish1.Visible = True
                DRFinish1.BackColor = Color.GreenYellow
                DRFinish1.ReadOnly = False
                ShowRequiredFieldValidator("DRFinish1Rqd", "DRFinish1", "異常：需輸入Finish-1")
            Case 2  '修改
                DRFinish1.Visible = True
                DRFinish1.BackColor = Color.Yellow
                DRFinish1.ReadOnly = False
            Case Else   '隱藏
                DRFinish1.Visible = False
        End Select

        'RSlider2
        'Select Case FindFieldInf("RSlider2")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRSlider2.BackColor = Color.LightGray
                DRSlider2.Visible = True
                DRSlider2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRSlider2.Visible = True
                DRSlider2.BackColor = Color.GreenYellow
                DRSlider2.ReadOnly = False
                ShowRequiredFieldValidator("DRSlider2Rqd", "DRSlider2", "異常：需輸入Slider-2")
            Case 2  '修改
                DRSlider2.Visible = True
                DRSlider2.BackColor = Color.Yellow
                DRSlider2.ReadOnly = False
            Case Else   '隱藏
                DRSlider2.Visible = False
        End Select

        'RFinish2
        'Select Case FindFieldInf("RFinish2")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRFinish2.BackColor = Color.LightGray
                DRFinish2.Visible = True
                DRFinish2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRFinish2.Visible = True
                DRFinish2.BackColor = Color.GreenYellow
                DRFinish2.ReadOnly = False
                ShowRequiredFieldValidator("DRFinish2Rqd", "DRFinish2", "異常：需輸入Finish-2")
            Case 2  '修改
                DRFinish2.Visible = True
                DRFinish2.BackColor = Color.Yellow
                DRFinish2.ReadOnly = False
            Case Else   '隱藏
                DRFinish2.Visible = False
        End Select

        'RSRequest1
        'Select Case FindFieldInf("RSRequest1")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRSRequest1.BackColor = Color.LightGray
                DRSRequest1.Visible = True
                DRSRequest1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRSRequest1.Visible = True
                DRSRequest1.BackColor = Color.GreenYellow
                DRSRequest1.ReadOnly = False
                ShowRequiredFieldValidator("DRSRequest1Rqd", "DRSRequest1", "異常：需輸入特殊要求-1")
            Case 2  '修改
                DRSRequest1.Visible = True
                DRSRequest1.BackColor = Color.Yellow
                DRSRequest1.ReadOnly = False
            Case Else   '隱藏
                DRSRequest1.Visible = False
        End Select

        'RSRequest2
        'Select Case FindFieldInf("RSRequest2")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRSRequest2.BackColor = Color.LightGray
                DRSRequest2.Visible = True
                DRSRequest2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRSRequest2.Visible = True
                DRSRequest2.BackColor = Color.GreenYellow
                DRSRequest2.ReadOnly = False
                ShowRequiredFieldValidator("DRSRequest2Rqd", "DRSRequest2", "異常：需輸入特殊要求-2")
            Case 2  '修改
                DRSRequest2.Visible = True
                DRSRequest2.BackColor = Color.Yellow
                DRSRequest2.ReadOnly = False
            Case Else   '隱藏
                DRSRequest2.Visible = False
        End Select

        'RSRequest3
        'Select Case FindFieldInf("RSRequest3")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRSRequest3.BackColor = Color.LightGray
                DRSRequest3.Visible = True
                DRSRequest3.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRSRequest3.Visible = True
                DRSRequest3.BackColor = Color.GreenYellow
                DRSRequest3.ReadOnly = False
                ShowRequiredFieldValidator("DRSRequest3Rqd", "DRSRequest3", "異常：需輸入特殊要求-3")
            Case 2  '修改
                DRSRequest3.Visible = True
                DRSRequest3.BackColor = Color.Yellow
                DRSRequest3.ReadOnly = False
            Case Else   '隱藏
                DRSRequest3.Visible = False
        End Select

        'RSRequest4
        'Select Case FindFieldInf("RSRequest4")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRSRequest4.BackColor = Color.LightGray
                DRSRequest4.Visible = True
                DRSRequest4.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRSRequest4.Visible = True
                DRSRequest4.BackColor = Color.GreenYellow
                DRSRequest4.ReadOnly = False
                ShowRequiredFieldValidator("DRSRequest4Rqd", "DRSRequest4", "異常：需輸入特殊要求-4")
            Case 2  '修改
                DRSRequest4.Visible = True
                DRSRequest4.BackColor = Color.Yellow
                DRSRequest4.ReadOnly = False
            Case Else   '隱藏
                DRSRequest4.Visible = False
        End Select

        'RSRequest5
        'Select Case FindFieldInf("RSRequest5")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRSRequest5.BackColor = Color.LightGray
                DRSRequest5.Visible = True
                DRSRequest5.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRSRequest5.Visible = True
                DRSRequest5.BackColor = Color.GreenYellow
                DRSRequest5.ReadOnly = False
                ShowRequiredFieldValidator("DRSRequest5Rqd", "DRSRequest5", "異常：需輸入特殊要求-5")
            Case 2  '修改
                DRSRequest5.Visible = True
                DRSRequest5.BackColor = Color.Yellow
                DRSRequest5.ReadOnly = False
            Case Else   '隱藏
                DRSRequest5.Visible = False
        End Select

        'RSRequest6
        'Select Case FindFieldInf("RSRequest6")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRSRequest6.BackColor = Color.LightGray
                DRSRequest6.Visible = True
                DRSRequest6.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRSRequest6.Visible = True
                DRSRequest6.BackColor = Color.GreenYellow
                DRSRequest6.ReadOnly = False
                ShowRequiredFieldValidator("DRSRequest6Rqd", "DRSRequest6", "異常：需輸入特殊要求-6")
            Case 2  '修改
                DRSRequest6.Visible = True
                DRSRequest6.BackColor = Color.Yellow
                DRSRequest6.ReadOnly = False
            Case Else   '隱藏
                DRSRequest6.Visible = False
        End Select

        'RFamily
        'Select Case FindFieldInf("RFamily")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRFamily.BackColor = Color.LightGray
                DRFamily.Visible = True
                DRFamily.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRFamily.Visible = True
                DRFamily.BackColor = Color.GreenYellow
                DRFamily.ReadOnly = False
                ShowRequiredFieldValidator("DRFamilyRqd", "DRFamily", "異常：需輸入Family Code")
            Case 2  '修改
                DRFamily.Visible = True
                DRFamily.BackColor = Color.Yellow
                DRFamily.ReadOnly = False
            Case Else   '隱藏
                DRFamily.Visible = False
        End Select

        'RST1
        'Select Case FindFieldInf("RST1")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRST1.BackColor = Color.LightGray
                DRST1.Visible = True
                DRST1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRST1.Visible = True
                DRST1.BackColor = Color.GreenYellow
                DRST1.ReadOnly = False
                ShowRequiredFieldValidator("DRST1Rqd", "DRST1", "異常：需輸入統計區分-1")
            Case 2  '修改
                DRST1.Visible = True
                DRST1.BackColor = Color.Yellow
                DRST1.ReadOnly = False
            Case Else   '隱藏
                DRST1.Visible = False
        End Select

        'RST2
        'Select Case FindFieldInf("RST2")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRST2.BackColor = Color.LightGray
                DRST2.Visible = True
                DRST2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRST2.Visible = True
                DRST2.BackColor = Color.GreenYellow
                DRST2.ReadOnly = False
                ShowRequiredFieldValidator("DRST2Rqd", "DRST2", "異常：需輸入統計區分-2")
            Case 2  '修改
                DRST2.Visible = True
                DRST2.BackColor = Color.Yellow
                DRST2.ReadOnly = False
            Case Else   '隱藏
                DRST2.Visible = False
        End Select

        'RST3
        'Select Case FindFieldInf("RST3")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRST3.BackColor = Color.LightGray
                DRST3.Visible = True
                DRST3.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRST3.Visible = True
                DRST3.BackColor = Color.GreenYellow
                DRST3.ReadOnly = False
                ShowRequiredFieldValidator("DRST3Rqd", "DRST3", "異常：需輸入統計區分-3")
            Case 2  '修改
                DRST3.Visible = True
                DRST3.BackColor = Color.Yellow
                DRST3.ReadOnly = False
            Case Else   '隱藏
                DRST3.Visible = False
        End Select

        'RST4
        'Select Case FindFieldInf("RST4")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRST4.BackColor = Color.LightGray
                DRST4.Visible = True
                DRST4.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRST4.Visible = True
                DRST4.BackColor = Color.GreenYellow
                DRST4.ReadOnly = False
                ShowRequiredFieldValidator("DRST4Rqd", "DRST4", "異常：需輸入統計區分-4")
            Case 2  '修改
                DRST4.Visible = True
                DRST4.BackColor = Color.Yellow
                DRST4.ReadOnly = False
            Case Else   '隱藏
                DRST4.Visible = False
        End Select

        'RST5
        'Select Case FindFieldInf("RST5")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRST5.BackColor = Color.LightGray
                DRST5.Visible = True
                DRST5.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRST5.Visible = True
                DRST5.BackColor = Color.GreenYellow
                DRST5.ReadOnly = False
                ShowRequiredFieldValidator("DRST5Rqd", "DRST5", "異常：需輸入統計區分-5")
            Case 2  '修改
                DRST5.Visible = True
                DRST5.BackColor = Color.Yellow
                DRST5.ReadOnly = False
            Case Else   '隱藏
                DRST5.Visible = False
        End Select

        'RST6
        'Select Case FindFieldInf("RST6")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRST6.BackColor = Color.LightGray
                DRST6.Visible = True
                DRST6.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRST6.Visible = True
                DRST6.BackColor = Color.GreenYellow
                DRST6.ReadOnly = False
                ShowRequiredFieldValidator("DRST6Rqd", "DRST6", "異常：需輸入統計區分-6")
            Case 2  '修改
                DRST6.Visible = True
                DRST6.BackColor = Color.Yellow
                DRST6.ReadOnly = False
            Case Else   '隱藏
                DRST6.Visible = False
        End Select

        'RST7
        'Select Case FindFieldInf("RST7")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRST7.BackColor = Color.LightGray
                DRST7.Visible = True
                DRST7.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRST7.Visible = True
                DRST7.BackColor = Color.GreenYellow
                DRST7.ReadOnly = False
                ShowRequiredFieldValidator("DRST7Rqd", "DRST7", "異常：需輸入統計區分-7")
            Case 2  '修改
                DRST7.Visible = True
                DRST7.BackColor = Color.Yellow
                DRST7.ReadOnly = False
            Case Else   '隱藏
                DRST7.Visible = False
        End Select

        'RNoDisplay
        'Select Case FindFieldInf("RNoDisplay")
        Select Case FindFieldInf("RCode")
            Case 0  '顯示
                DRNoDisplay.BackColor = Color.LightGray
                DRNoDisplay.Visible = True
                DRNoDisplay.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRNoDisplay.Visible = True
                DRNoDisplay.BackColor = Color.GreenYellow
                DRNoDisplay.ReadOnly = False
                ShowRequiredFieldValidator("DRNoDisplayRqd", "DRNoDisplay", "異常：需輸入No Display")
            Case 2  '修改
                DRNoDisplay.Visible = True
                DRNoDisplay.BackColor = Color.Yellow
                DRNoDisplay.ReadOnly = False
            Case Else   '隱藏
                DRNoDisplay.Visible = False
        End Select

        'PriceVersion(A001/A206/A211/A999/K206/K211)
        Select Case FindFieldInf("PriceVersion")
            Case 0  '顯示
                DA001.BackColor = Color.LightGray
                DA001.Visible = True
                DA001.Enabled = False
                DA206.BackColor = Color.LightGray
                DA206.Visible = True
                DA206.Enabled = False
                DA211.BackColor = Color.LightGray
                DA211.Visible = True
                DA211.Enabled = False
                DA999.BackColor = Color.LightGray
                DA999.Visible = True
                DA999.Enabled = False
                DK206.BackColor = Color.LightGray
                DK206.Visible = True
                DK206.Enabled = False
                DK211.BackColor = Color.LightGray
                DK211.Visible = True
                DK211.Enabled = False
            Case 1  '修改+檢查
                DA001.Visible = True
                DA001.BackColor = Color.GreenYellow
                'ShowRequiredFieldValidator("DA001Rqd", "DA001", "異常：需輸入單價版本")
                DA206.Visible = True
                DA206.BackColor = Color.GreenYellow
                DA206.Enabled = False
                'ShowRequiredFieldValidator("DA206Rqd", "DA206", "異常：需輸入單價版本")
                DA211.Visible = True
                DA211.BackColor = Color.GreenYellow
                'ShowRequiredFieldValidator("DA211Rqd", "DA211", "異常：需輸入單價版本")
                DA999.Visible = True
                DA999.BackColor = Color.GreenYellow
                'ShowRequiredFieldValidator("DA999Rqd", "DA999", "異常：需輸入單價版本")
                DK206.Visible = True
                DK206.BackColor = Color.GreenYellow
                DK206.Enabled = False
                'ShowRequiredFieldValidator("DK206Rqd", "DK206", "異常：需輸入單價版本")
                DK211.Visible = True
                DK211.BackColor = Color.GreenYellow
                'ShowRequiredFieldValidator("DK211Rqd", "DK211", "異常：需輸入單價版本")
            Case 2  '修改
                DA001.Visible = True
                DA001.BackColor = Color.Yellow
                DA206.Visible = True
                DA206.BackColor = Color.Yellow
                DA206.Enabled = False
                DA211.Visible = True
                DA211.BackColor = Color.Yellow
                DA999.Visible = True
                DA999.BackColor = Color.Yellow
                DK206.Visible = True
                DK206.BackColor = Color.Yellow
                DK206.Enabled = False
                DK211.Visible = True
                DK211.BackColor = Color.Yellow
            Case Else   '隱藏
                DA001.Visible = False
                DA206.Visible = False
                DA211.Visible = False
                DA999.Visible = False
                DK206.Visible = False
                DK211.Visible = False
        End Select
        '----------------------------------------------------------------
        'Code
        Select Case FindFieldInf("Code")
            Case 0  '顯示
                DCode.BackColor = Color.LightGray
                DCode.Visible = True
                DCode.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCode.Visible = True
                DCode.BackColor = Color.GreenYellow
                DCode.ReadOnly = False
                ShowRequiredFieldValidator("DCodeRqd", "DCode", "異常：需輸入Code")
            Case 2  '修改
                DCode.Visible = True
                DCode.BackColor = Color.Yellow
                DCode.ReadOnly = False
            Case Else   '隱藏
                DCode.Visible = False
        End Select

        'ItemName1
        Select Case FindFieldInf("ItemName1")
            Case 0  '顯示
                DItemName1.BackColor = Color.LightGray
                DItemName1.Visible = True
                DItemName1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DItemName1.Visible = True
                DItemName1.BackColor = Color.GreenYellow
                DItemName1.ReadOnly = False
                ShowRequiredFieldValidator("DItemName1Rqd", "DItemName1", "異常：需輸入Item Name-1")
            Case 2  '修改
                DItemName1.Visible = True
                DItemName1.BackColor = Color.Yellow
                DItemName1.ReadOnly = False
            Case Else   '隱藏
                DItemName1.Visible = False
        End Select

        'ItemName2
        Select Case FindFieldInf("ItemName2")
            Case 0  '顯示
                DItemName2.BackColor = Color.LightGray
                DItemName2.Visible = True
                DItemName2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DItemName2.Visible = True
                DItemName2.BackColor = Color.GreenYellow
                DItemName2.ReadOnly = False
                ShowRequiredFieldValidator("DItemName2Rqd", "DItemName2", "異常：需輸入Item Name-2")
            Case 2  '修改
                DItemName2.Visible = True
                DItemName2.BackColor = Color.Yellow
                DItemName2.ReadOnly = False
            Case Else   '隱藏
                DItemName2.Visible = False
        End Select

        'ItemName3
        Select Case FindFieldInf("ItemName3")
            Case 0  '顯示
                DItemName3.BackColor = Color.LightGray
                DItemName3.Visible = True
                DItemName3.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DItemName3.Visible = True
                DItemName3.BackColor = Color.GreenYellow
                DItemName3.ReadOnly = False
                ShowRequiredFieldValidator("DItemName3Rqd", "DItemName3", "異常：需輸入Item Name-3")
            Case 2  '修改
                DItemName3.Visible = True
                DItemName3.BackColor = Color.Yellow
                DItemName3.ReadOnly = False
            Case Else   '隱藏
                DItemName3.Visible = False
        End Select

        'Size
        Select Case FindFieldInf("Size")
            Case 0  '顯示
                DSize.BackColor = Color.LightGray
                DSize.Visible = True
                DSize.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSize.Visible = True
                DSize.BackColor = Color.GreenYellow
                DSize.ReadOnly = False
                ShowRequiredFieldValidator("DSizeRqd", "DSize", "異常：需輸入Size")
            Case 2  '修改
                DSize.Visible = True
                DSize.BackColor = Color.Yellow
                DSize.ReadOnly = False
            Case Else   '隱藏
                DSize.Visible = False
        End Select

        'Chain
        Select Case FindFieldInf("Chain")
            Case 0  '顯示
                DChain.BackColor = Color.LightGray
                DChain.Visible = True
                DChain.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DChain.Visible = True
                DChain.BackColor = Color.GreenYellow
                DChain.ReadOnly = False
                ShowRequiredFieldValidator("DChainRqd", "DChain", "異常：需輸入Chain")
            Case 2  '修改
                DChain.Visible = True
                DChain.BackColor = Color.Yellow
                DChain.ReadOnly = False
            Case Else   '隱藏
                DChain.Visible = False
        End Select

        'Class
        Select Case FindFieldInf("Class")
            Case 0  '顯示
                DClass.BackColor = Color.LightGray
                DClass.Visible = True
                DClass.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DClass.Visible = True
                DClass.BackColor = Color.GreenYellow
                DClass.ReadOnly = False
                ShowRequiredFieldValidator("DClassRqd", "DClass", "異常：需輸入Class")
            Case 2  '修改
                DClass.Visible = True
                DClass.BackColor = Color.Yellow
                DClass.ReadOnly = False
            Case Else   '隱藏
                DClass.Visible = False
        End Select

        'Tape
        Select Case FindFieldInf("Tape")
            Case 0  '顯示
                DTape.BackColor = Color.LightGray
                DTape.Visible = True
                DTape.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTape.Visible = True
                DTape.BackColor = Color.GreenYellow
                DTape.ReadOnly = False
                ShowRequiredFieldValidator("DTapeRqd", "DTape", "異常：需輸入Tape")
            Case 2  '修改
                DTape.Visible = True
                DTape.BackColor = Color.Yellow
                DTape.ReadOnly = False
            Case Else   '隱藏
                DTape.Visible = False
        End Select

        'Slider1
        Select Case FindFieldInf("Slider1")
            Case 0  '顯示
                DSlider1.BackColor = Color.LightGray
                DSlider1.Visible = True
                DSlider1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSlider1.Visible = True
                DSlider1.BackColor = Color.GreenYellow
                DSlider1.ReadOnly = False
                ShowRequiredFieldValidator("DSlider1Rqd", "DSlider1", "異常：需輸入Slider-1")
            Case 2  '修改
                DSlider1.Visible = True
                DSlider1.BackColor = Color.Yellow
                DSlider1.ReadOnly = False
            Case Else   '隱藏
                DSlider1.Visible = False
        End Select

        'Finish1
        Select Case FindFieldInf("Finish1")
            Case 0  '顯示
                DFinish1.BackColor = Color.LightGray
                DFinish1.Visible = True
                DFinish1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DFinish1.Visible = True
                DFinish1.BackColor = Color.GreenYellow
                DFinish1.ReadOnly = False
                ShowRequiredFieldValidator("DFinish1Rqd", "DFinish1", "異常：需輸入Finish-1")
            Case 2  '修改
                DFinish1.Visible = True
                DFinish1.BackColor = Color.Yellow
                DFinish1.ReadOnly = False
            Case Else   '隱藏
                DFinish1.Visible = False
        End Select

        'Slider2
        Select Case FindFieldInf("Slider2")
            Case 0  '顯示
                DSlider2.BackColor = Color.LightGray
                DSlider2.Visible = True
                DSlider2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSlider2.Visible = True
                DSlider2.BackColor = Color.GreenYellow
                DSlider2.ReadOnly = False
                ShowRequiredFieldValidator("DSlider2Rqd", "DSlider2", "異常：需輸入Slider-2")
            Case 2  '修改
                DSlider2.Visible = True
                DSlider2.BackColor = Color.Yellow
                DSlider2.ReadOnly = False
            Case Else   '隱藏
                DSlider2.Visible = False
        End Select

        'Finish2
        Select Case FindFieldInf("Finish2")
            Case 0  '顯示
                DFinish2.BackColor = Color.LightGray
                DFinish2.Visible = True
                DFinish2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DFinish2.Visible = True
                DFinish2.BackColor = Color.GreenYellow
                DFinish2.ReadOnly = False
                ShowRequiredFieldValidator("DFinish2Rqd", "DFinish2", "異常：需輸入Finish-2")
            Case 2  '修改
                DFinish2.Visible = True
                DFinish2.BackColor = Color.Yellow
                DFinish2.ReadOnly = False
            Case Else   '隱藏
                DFinish2.Visible = False
        End Select

        'SRequest1
        Select Case FindFieldInf("SRequest1")
            Case 0  '顯示
                DSRequest1.BackColor = Color.LightGray
                DSRequest1.Visible = True
                DSRequest1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSRequest1.Visible = True
                DSRequest1.BackColor = Color.GreenYellow
                DSRequest1.ReadOnly = False
                ShowRequiredFieldValidator("DSRequest1Rqd", "DSRequest1", "異常：需輸入特殊要求-1")
            Case 2  '修改
                DSRequest1.Visible = True
                DSRequest1.BackColor = Color.Yellow
                DSRequest1.ReadOnly = False
            Case Else   '隱藏
                DSRequest1.Visible = False
        End Select

        'SRequest2
        Select Case FindFieldInf("SRequest2")
            Case 0  '顯示
                DSRequest2.BackColor = Color.LightGray
                DSRequest2.Visible = True
                DSRequest2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSRequest2.Visible = True
                DSRequest2.BackColor = Color.GreenYellow
                DSRequest2.ReadOnly = False
                ShowRequiredFieldValidator("DSRequest2Rqd", "DSRequest2", "異常：需輸入特殊要求-2")
            Case 2  '修改
                DSRequest2.Visible = True
                DSRequest2.BackColor = Color.Yellow
                DSRequest2.ReadOnly = False
            Case Else   '隱藏
                DSRequest2.Visible = False
        End Select

        'SRequest3
        Select Case FindFieldInf("SRequest3")
            Case 0  '顯示
                DSRequest3.BackColor = Color.LightGray
                DSRequest3.Visible = True
                DSRequest3.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSRequest3.Visible = True
                DSRequest3.BackColor = Color.GreenYellow
                DSRequest3.ReadOnly = False
                ShowRequiredFieldValidator("DSRequest3Rqd", "DSRequest3", "異常：需輸入特殊要求-3")
            Case 2  '修改
                DSRequest3.Visible = True
                DSRequest3.BackColor = Color.Yellow
                DSRequest3.ReadOnly = False
            Case Else   '隱藏
                DSRequest3.Visible = False
        End Select

        'SRequest4
        Select Case FindFieldInf("SRequest4")
            Case 0  '顯示
                DSRequest4.BackColor = Color.LightGray
                DSRequest4.Visible = True
                DSRequest4.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSRequest4.Visible = True
                DSRequest4.BackColor = Color.GreenYellow
                DSRequest4.ReadOnly = False
                ShowRequiredFieldValidator("DSRequest4Rqd", "DSRequest4", "異常：需輸入特殊要求-4")
            Case 2  '修改
                DSRequest4.Visible = True
                DSRequest4.BackColor = Color.Yellow
                DSRequest4.ReadOnly = False
            Case Else   '隱藏
                DSRequest4.Visible = False
        End Select

        'SRequest5
        Select Case FindFieldInf("SRequest5")
            Case 0  '顯示
                DSRequest5.BackColor = Color.LightGray
                DSRequest5.Visible = True
                DSRequest5.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSRequest5.Visible = True
                DSRequest5.BackColor = Color.GreenYellow
                DSRequest5.ReadOnly = False
                ShowRequiredFieldValidator("DSRequest5Rqd", "DSRequest5", "異常：需輸入特殊要求-5")
            Case 2  '修改
                DSRequest5.Visible = True
                DSRequest5.BackColor = Color.Yellow
                DSRequest5.ReadOnly = False
            Case Else   '隱藏
                DSRequest5.Visible = False
        End Select

        'SRequest6
        Select Case FindFieldInf("SRequest6")
            Case 0  '顯示
                DSRequest6.BackColor = Color.LightGray
                DSRequest6.Visible = True
                DSRequest6.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSRequest6.Visible = True
                DSRequest6.BackColor = Color.GreenYellow
                DSRequest6.ReadOnly = False
                ShowRequiredFieldValidator("DSRequest6Rqd", "DSRequest6", "異常：需輸入特殊要求-6")
            Case 2  '修改
                DSRequest6.Visible = True
                DSRequest6.BackColor = Color.Yellow
                DSRequest6.ReadOnly = False
            Case Else   '隱藏
                DSRequest6.Visible = False
        End Select

        'Family
        Select Case FindFieldInf("Family")
            Case 0  '顯示
                DFamily.BackColor = Color.LightGray
                DFamily.Visible = True
                DFamily.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DFamily.Visible = True
                DFamily.BackColor = Color.GreenYellow
                DFamily.ReadOnly = False
                ShowRequiredFieldValidator("DFamilyRqd", "DFamily", "異常：需輸入Family Code")
            Case 2  '修改
                DFamily.Visible = True
                DFamily.BackColor = Color.Yellow
                DFamily.ReadOnly = False
            Case Else   '隱藏
                DFamily.Visible = False
        End Select

        'ST1
        Select Case FindFieldInf("ST1")
            Case 0  '顯示
                DST1.BackColor = Color.LightGray
                DST1.Visible = True
                DST1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DST1.Visible = True
                DST1.BackColor = Color.GreenYellow
                DST1.ReadOnly = False
                ShowRequiredFieldValidator("DST1Rqd", "DST1", "異常：需輸入統計區分-1")
            Case 2  '修改
                DST1.Visible = True
                DST1.BackColor = Color.Yellow
                DST1.ReadOnly = False
            Case Else   '隱藏
                DST1.Visible = False
        End Select

        'ST2
        Select Case FindFieldInf("ST2")
            Case 0  '顯示
                DST2.BackColor = Color.LightGray
                DST2.Visible = True
                DST2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DST2.Visible = True
                DST2.BackColor = Color.GreenYellow
                DST2.ReadOnly = False
                ShowRequiredFieldValidator("DST2Rqd", "DST2", "異常：需輸入統計區分-2")
            Case 2  '修改
                DST2.Visible = True
                DST2.BackColor = Color.Yellow
                DST2.ReadOnly = False
            Case Else   '隱藏
                DST2.Visible = False
        End Select

        'ST3
        Select Case FindFieldInf("ST3")
            Case 0  '顯示
                DST3.BackColor = Color.LightGray
                DST3.Visible = True
                DST3.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DST3.Visible = True
                DST3.BackColor = Color.GreenYellow
                DST3.ReadOnly = False
                ShowRequiredFieldValidator("DST3Rqd", "DST3", "異常：需輸入統計區分-3")
            Case 2  '修改
                DST3.Visible = True
                DST3.BackColor = Color.Yellow
                DST3.ReadOnly = False
            Case Else   '隱藏
                DST3.Visible = False
        End Select

        'ST4
        Select Case FindFieldInf("ST4")
            Case 0  '顯示
                DST4.BackColor = Color.LightGray
                DST4.Visible = True
                DST4.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DST4.Visible = True
                DST4.BackColor = Color.GreenYellow
                DST4.ReadOnly = False
                ShowRequiredFieldValidator("DST4Rqd", "DST4", "異常：需輸入統計區分-4")
            Case 2  '修改
                DST4.Visible = True
                DST4.BackColor = Color.Yellow
                DST4.ReadOnly = False
            Case Else   '隱藏
                DST4.Visible = False
        End Select

        'ST5
        Select Case FindFieldInf("ST5")
            Case 0  '顯示
                DST5.BackColor = Color.LightGray
                DST5.Visible = True
                DST5.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DST5.Visible = True
                DST5.BackColor = Color.GreenYellow
                DST5.ReadOnly = False
                ShowRequiredFieldValidator("DST5Rqd", "DST5", "異常：需輸入統計區分-5")
            Case 2  '修改
                DST5.Visible = True
                DST5.BackColor = Color.Yellow
                DST5.ReadOnly = False
            Case Else   '隱藏
                DST5.Visible = False
        End Select

        'ST6
        Select Case FindFieldInf("ST6")
            Case 0  '顯示
                DST6.BackColor = Color.LightGray
                DST6.Visible = True
                DST6.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DST6.Visible = True
                DST6.BackColor = Color.GreenYellow
                DST6.ReadOnly = False
                ShowRequiredFieldValidator("DST6Rqd", "DST6", "異常：需輸入統計區分-6")
            Case 2  '修改
                DST6.Visible = True
                DST6.BackColor = Color.Yellow
                DST6.ReadOnly = False
            Case Else   '隱藏
                DST6.Visible = False
        End Select

        'ST7
        Select Case FindFieldInf("ST7")
            Case 0  '顯示
                DST7.BackColor = Color.LightGray
                DST7.Visible = True
                DST7.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DST7.Visible = True
                DST7.BackColor = Color.GreenYellow
                DST7.ReadOnly = False
                ShowRequiredFieldValidator("DST7Rqd", "DST7", "異常：需輸入統計區分-7")
            Case 2  '修改
                DST7.Visible = True
                DST7.BackColor = Color.Yellow
                DST7.ReadOnly = False
            Case Else   '隱藏
                DST7.Visible = False
        End Select

        'NoDisplay
        Select Case FindFieldInf("NoDisplay")
            Case 0  '顯示
                DNoDisplay.BackColor = Color.LightGray
                DNoDisplay.Visible = True
                DNoDisplay.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DNoDisplay.Visible = True
                DNoDisplay.BackColor = Color.GreenYellow
                DNoDisplay.ReadOnly = False
                ShowRequiredFieldValidator("DNoDisplayRqd", "DNoDisplay", "異常：需輸入No Display")
            Case 2  '修改
                DNoDisplay.Visible = True
                DNoDisplay.BackColor = Color.Yellow
                DNoDisplay.ReadOnly = False
            Case Else   '隱藏
                DNoDisplay.Visible = False
        End Select

        'PriceDescr
        DPriceDescr.BackColor = Color.LightGray
        DPriceDescr.Visible = True
        DPriceDescr.Attributes.Add("readonly", "true")
        'PriceNo
        DPriceNo.BackColor = Color.LightGray
        DPriceNo.Visible = True
        DPriceNo.Attributes.Add("readonly", "true")

        'Remark
        Select Case FindFieldInf("Remark")
            Case 0  '顯示
                DRemark.BackColor = Color.LightGray
                DRemark.Visible = True
                DRemark.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRemark.Visible = True
                DRemark.BackColor = Color.GreenYellow
                DRemark.ReadOnly = False
                ShowRequiredFieldValidator("DRemarkRqd", "DRemark", "異常：需輸入備註")
            Case 2  '修改
                DRemark.Visible = True
                DRemark.BackColor = Color.Yellow
                DRemark.ReadOnly = False
            Case Else   '隱藏
                DRemark.Visible = False
        End Select

        'SLDPrice
        Select Case FindFieldInf("SLDPrice")
            Case 0  '顯示
                DSLDPrice.BackColor = Color.LightGray
                DSLDPrice.Visible = True
                DSLDPrice.Attributes.Add("readonly", "true")
                DRemark.Width = 479
            Case 1  '修改+檢查
                DSLDPrice.Visible = True
                DSLDPrice.BackColor = Color.GreenYellow
                DSLDPrice.ReadOnly = False
                ShowRequiredFieldValidator("DSLDPriceRqd", "DSLDPrice", "異常：需輸入拉頭單價")
                DRemark.Width = 479
            Case 2  '修改
                DSLDPrice.Visible = True
                DSLDPrice.BackColor = Color.Yellow
                DSLDPrice.ReadOnly = False
                DRemark.Width = 479
            Case Else   '隱藏
                DSLDPrice.Visible = False
                DRemark.Width = 583
        End Select

        'Attachfile1
        Select Case FindFieldInf("Attachfile1")
            Case 0  '顯示
                DAttachfile1.Visible = False
                DAttachfile1.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DAttachfile1Rqd", "DAttachfile1", "異常：需輸入附件")
                DAttachfile1.Visible = True
                DAttachfile1.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DAttachfile1.Visible = True
                DAttachfile1.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DAttachfile1.Visible = False
        End Select


        'Customer
        Select Case FindFieldInf("Customer")
            Case 0  '顯示
                DCustomer.BackColor = Color.LightGray
                DCustomer.ReadOnly = True
                DCustomer.Visible = True
                BCustomer.Visible = False

            Case 1  '修改+檢查
                DCustomer.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerRqd", "DCustomer", "異常：需輸入客戶名稱")
                DCustomer.Visible = True
                DCustomer.ReadOnly = True
                BCustomer.Visible = True

            Case 2  '修改
                DCustomer.BackColor = Color.Yellow
                DCustomer.Visible = True
                BCustomer.Visible = True

            Case Else   '隱藏
                DCustomer.Visible = False
                BCustomer.Visible = False

        End Select
        If pPost = "New" Then SetFieldData("Customer", "ZZZZZZ")


        'CustomerCode
        Select Case FindFieldInf("CustomerCode")
            Case 0  '顯示
                DCustomerCode.BackColor = Color.LightGray
                DCustomerCode.ReadOnly = True
                DCustomerCode.Visible = True
            Case 1  '修改+檢查
                DCustomerCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerCodeRqd", "DCustomerCode", "異常：需輸入客戶名稱")
                DCustomerCode.Visible = True
                DCustomer.ReadOnly = True
            Case 2  '修改
                DCustomerCode.BackColor = Color.Yellow
                DCustomerCode.Visible = True
            Case Else   '隱藏
                DCustomerCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CustomerCode", "ZZZZZZ")



        'Buyer
        Select Case FindFieldInf("Buyer")
            Case 0  '顯示
                DBuyer.BackColor = Color.LightGray
                DBuyer.ReadOnly = True
                DBuyer.Visible = True
                BBuyer.Visible = False
            Case 1  '修改+檢查
                DBuyer.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerRqd", "DBuyer", "異常：需輸入Buyer")
                DBuyer.Visible = True
                DBuyer.ReadOnly = True
                BBuyer.Visible = True
            Case 2  '修改
                DBuyer.BackColor = Color.Yellow
                DBuyer.Visible = True
                BBuyer.Visible = True
            Case Else   '隱藏
                DBuyer.Visible = False
                BBuyer.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Buyer", "ZZZZZZ")


        'BuyerCode
        Select Case FindFieldInf("BuyerCode")
            Case 0  '顯示
                DBuyerCode.BackColor = Color.LightGray
                DBuyerCode.ReadOnly = True
                DBuyerCode.Visible = True
            Case 1  '修改+檢查
                DBuyerCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerCodeRqd", "DBuyerCode", "異常：需輸入Buyer")
                DBuyerCode.Visible = True
                DBuyerCode.ReadOnly = True
            Case 2  '修改
                DBuyerCode.BackColor = Color.Yellow
                DBuyerCode.Visible = True
            Case Else   '隱藏
                DBuyerCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BuyerCode", "ZZZZZZ")


        'ForUse
        Select Case FindFieldInf("ForUse")
            Case 0  '顯示
                DForUse.BackColor = Color.LightGray
                DForUse.Visible = True
            Case 1  '修改+檢查
                DForUse.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DForUseRqd", "DForUse", "異常：需輸入用途")
                DForUse.Visible = True
            Case 2  '修改
                DForUse.BackColor = Color.Yellow
                DForUse.Visible = True
            Case Else   '隱藏
                DForUse.Visible = False
        End Select
        If Not IsPostBack Then SetFieldData("ForUse", "ZZZZZZ")




    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("ItemRegisterFilePath")
        Dim SQL As String
        SQL = "Select * From F_ItemRegisterSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtItemRegisterSheet As DataTable = uDataBase.GetDataTable(SQL)
        If dtItemRegisterSheet.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dtItemRegisterSheet.Rows(0).Item("No")                           ' No
            DDate.Text = dtItemRegisterSheet.Rows(0).Item("Date")                       ' 申請日
            DName.Text = dtItemRegisterSheet.Rows(0).Item("Name")                       ' 申請人姓名
            DJobTitle.Text = dtItemRegisterSheet.Rows(0).Item("JobTitle")               ' 職稱
            DDivision.Text = dtItemRegisterSheet.Rows(0).Item("Division")               ' 部門
            DRCode.Text = dtItemRegisterSheet.Rows(0).Item("RCode")                     ' R Code
            DRItemName1.Text = dtItemRegisterSheet.Rows(0).Item("RItemName1")           ' R Item Name-1
            DRItemName2.Text = dtItemRegisterSheet.Rows(0).Item("RItemName2")           ' R Item Name-2
            DRItemName3.Text = dtItemRegisterSheet.Rows(0).Item("RItemName3")           ' R Item Name-3
            DRSize.Text = dtItemRegisterSheet.Rows(0).Item("RSize")                     ' R Size
            DRChain.Text = dtItemRegisterSheet.Rows(0).Item("RChain")                   ' R Chain
            DRClass.Text = dtItemRegisterSheet.Rows(0).Item("RClass")                   ' R Class
            DRTape.Text = dtItemRegisterSheet.Rows(0).Item("RTape")                     ' R Tape
            DRSlider1.Text = dtItemRegisterSheet.Rows(0).Item("RSlider1")               ' R Slider1
            DRFinish1.Text = dtItemRegisterSheet.Rows(0).Item("RFinish1")               ' R Finish1
            DRSlider2.Text = dtItemRegisterSheet.Rows(0).Item("RSlider2")               ' R Slider2
            DRFinish2.Text = dtItemRegisterSheet.Rows(0).Item("RFinish2")               ' R Finish2
            DRSRequest1.Text = dtItemRegisterSheet.Rows(0).Item("RSRequest1")           ' R 特殊要求1
            DRSRequest2.Text = dtItemRegisterSheet.Rows(0).Item("RSRequest2")           ' R 特殊要求2
            DRSRequest3.Text = dtItemRegisterSheet.Rows(0).Item("RSRequest3")           ' R 特殊要求3
            DRSRequest4.Text = dtItemRegisterSheet.Rows(0).Item("RSRequest4")           ' R 特殊要求4
            DRSRequest5.Text = dtItemRegisterSheet.Rows(0).Item("RSRequest5")           ' R 特殊要求5
            DRSRequest6.Text = dtItemRegisterSheet.Rows(0).Item("RSRequest6")           ' R 特殊要求6
            DRFamily.Text = dtItemRegisterSheet.Rows(0).Item("RFamily")                 ' R Family Code
            DRST1.Text = dtItemRegisterSheet.Rows(0).Item("RST1")                       ' R 統計區分1
            DRST2.Text = dtItemRegisterSheet.Rows(0).Item("RST2")                       ' R 統計區分2
            DRST3.Text = dtItemRegisterSheet.Rows(0).Item("RST3")                       ' R 統計區分3
            DRST4.Text = dtItemRegisterSheet.Rows(0).Item("RST4")                       ' R 統計區分4
            DRST5.Text = dtItemRegisterSheet.Rows(0).Item("RST5")                       ' R 統計區分5
            DRST6.Text = dtItemRegisterSheet.Rows(0).Item("RST6")                       ' R 統計區分6
            DRST7.Text = dtItemRegisterSheet.Rows(0).Item("RST7")                       ' R 統計區分7
            DRNoDisplay.Text = dtItemRegisterSheet.Rows(0).Item("RNoDisplay")           ' R No Display
            ' PriceVersion
            DA001.Checked = False
            DA206.Checked = False
            DA211.Checked = False
            DA999.Checked = False
            DK206.Checked = False
            DK211.Checked = False
            If dtItemRegisterSheet.Rows(0).Item("A001") = 1 Then DA001.Checked = True
            If dtItemRegisterSheet.Rows(0).Item("A206") = 1 Then DA206.Checked = True
            If dtItemRegisterSheet.Rows(0).Item("A211") = 1 Then DA211.Checked = True
            If dtItemRegisterSheet.Rows(0).Item("A999") = 1 Then DA999.Checked = True
            If dtItemRegisterSheet.Rows(0).Item("K206") = 1 Then DK206.Checked = True
            If dtItemRegisterSheet.Rows(0).Item("K211") = 1 Then DK211.Checked = True
            DCode.Text = dtItemRegisterSheet.Rows(0).Item("Code")                       ' Code
            DItemName1.Text = dtItemRegisterSheet.Rows(0).Item("ItemName1")             ' Item Name-1
            DItemName2.Text = dtItemRegisterSheet.Rows(0).Item("ItemName2")             ' Item Name-2
            DItemName3.Text = dtItemRegisterSheet.Rows(0).Item("ItemName3")             ' Item Name-3
            DSize.Text = dtItemRegisterSheet.Rows(0).Item("Size")                       ' Size
            DChain.Text = dtItemRegisterSheet.Rows(0).Item("Chain")                     ' Chain
            DClass.Text = dtItemRegisterSheet.Rows(0).Item("Class")                     ' Class
            DTape.Text = dtItemRegisterSheet.Rows(0).Item("Tape")                       ' Tape
            DSlider1.Text = dtItemRegisterSheet.Rows(0).Item("Slider1")                 ' Slider1
            DFinish1.Text = dtItemRegisterSheet.Rows(0).Item("Finish1")                 ' Finish1
            DSlider2.Text = dtItemRegisterSheet.Rows(0).Item("Slider2")                 ' Slider2
            DFinish2.Text = dtItemRegisterSheet.Rows(0).Item("Finish2")                 ' Finish2
            DSRequest1.Text = dtItemRegisterSheet.Rows(0).Item("SRequest1")             ' 特殊要求1
            DSRequest2.Text = dtItemRegisterSheet.Rows(0).Item("SRequest2")             ' 特殊要求2
            DSRequest3.Text = dtItemRegisterSheet.Rows(0).Item("SRequest3")             ' 特殊要求3
            DSRequest4.Text = dtItemRegisterSheet.Rows(0).Item("SRequest4")             ' 特殊要求4
            DSRequest5.Text = dtItemRegisterSheet.Rows(0).Item("SRequest5")             ' 特殊要求5
            DSRequest6.Text = dtItemRegisterSheet.Rows(0).Item("SRequest6")             ' 特殊要求6
            DFamily.Text = dtItemRegisterSheet.Rows(0).Item("Family")                   ' Family Code
            DST1.Text = dtItemRegisterSheet.Rows(0).Item("ST1")                         ' 統計區分1
            DST2.Text = dtItemRegisterSheet.Rows(0).Item("ST2")                         ' 統計區分2
            DST3.Text = dtItemRegisterSheet.Rows(0).Item("ST3")                         ' 統計區分3
            DST4.Text = dtItemRegisterSheet.Rows(0).Item("ST4")                         ' 統計區分4
            DST5.Text = dtItemRegisterSheet.Rows(0).Item("ST5")                         ' 統計區分5
            DST6.Text = dtItemRegisterSheet.Rows(0).Item("ST6")                         ' 統計區分6
            DST7.Text = dtItemRegisterSheet.Rows(0).Item("ST7")                         ' 統計區分7
            DNoDisplay.Text = dtItemRegisterSheet.Rows(0).Item("NoDisplay")             ' No Display
            If dtItemRegisterSheet.Rows(0).Item("PriceApply") = 0 Then                  ' Price Descriptioon / PriceNo
                DPriceDescr.Text = ""
                DPriceNo.Text = ""
            End If
            If dtItemRegisterSheet.Rows(0).Item("PriceApply") = 1 Then                  ' Price Descriptioon / PriceNo
                DPriceDescr.Text = "非賣品不設定/進口成品/手動計算"
                DPriceNo.Text = ""
            End If
            If dtItemRegisterSheet.Rows(0).Item("PriceApply") = 2 Then                  ' Price Descriptioon / PriceNo
                DPriceDescr.Text = "已設定"
                DPriceNo.Text = dtItemRegisterSheet.Rows(0).Item("PriceNo")
            End If
            DSLDPrice.Text = dtItemRegisterSheet.Rows(0).Item("SLDPrice")               ' SLD Price
            DRemark.Text = dtItemRegisterSheet.Rows(0).Item("Remark")                   ' Remark
            If dtItemRegisterSheet.Rows(0).Item("Attachfile1") <> "" Then
                LAttachfile1.NavigateUrl = Path & dtItemRegisterSheet.Rows(0).Item("Attachfile1")   '附件
                LAttachfile1.Visible = True
            Else
                LAttachfile1.Visible = False
            End If
            DBuyer.Text = dtItemRegisterSheet.Rows(0).Item("Buyer")                         ' BUYER
            DBuyerCode.Text = dtItemRegisterSheet.Rows(0).Item("BuyerCode")                         ' BUYER
            DCustomer.Text = dtItemRegisterSheet.Rows(0).Item("Customer")                         ' Customer
            DCustomerCode.Text = dtItemRegisterSheet.Rows(0).Item("CustomerCode")                         ' Customer
            SetFieldData("ForUse", dtItemRegisterSheet.Rows(0).Item("ForUse"))    '用途

            '設定修改後Key-Data
            DStatusItem.Value = DSize.Text.ToUpper + "!" + _
                                DChain.Text.ToUpper + "!" + _
                                DClass.Text.ToUpper + "!" + _
                                DTape.Text.ToUpper + "!" + _
                                DSlider1.Text.ToUpper + "!" + _
                                DFinish1.Text.ToUpper + "!" + _
                                DSlider2.Text.ToUpper + "!" + _
                                DFinish2.Text.ToUpper + "!" + _
                                DFinish2.Text.ToUpper + "!" + _
                                DSRequest1.Text.ToUpper + "!" + _
                                DSRequest2.Text.ToUpper + "!" + _
                                DSRequest3.Text.ToUpper + "!" + _
                                DSRequest4.Text.ToUpper + "!" + _
                                DSRequest5.Text.ToUpper + "!" + _
                                DSRequest6.Text
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

        '開發擔當
        If pFieldName = "SPDPerson" Then
            DSPDPerson.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSPDPerson.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='1151' and DKey='SPDPERSON' Order by DKey, Data "
                Dim dtPerson As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtPerson.Rows.Count - 1
                    Dim SPD As String() = dtPerson.Rows(i)("Data").ToString.Split("/")
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = SPD(0)
                    ListItem1.Value = SPD(1)
                    DSPDPerson.Items.Add(ListItem1)
                Next
            End If
        End If

        '用途
        If pFieldName = "ForUse" Then
            DForUse.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DForUse.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='1151' and DKey='FORUSE' Order by DKey, Data "
                Dim dtFORUSE As DataTable = uDataBase.GetDataTable(sql)
                DForUse.Items.Add("")
                For i As Integer = 0 To dtFORUSE.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFORUSE.Rows(i)("Data")
                    ListItem1.Value = dtFORUSE.Rows(i)("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DForUse.Items.Add(ListItem1)
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
                sql = "Select * From M_Referp Where Cat='014' Order by DKey, Data "
                Dim dtReasonCode As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtReasonCode.Rows.Count - 1
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
            For i As Integer = 0 To dtReason.Rows.Count - 1
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
        '
        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.Text = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", Top & "px")
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
        BFindItem.Attributes("onclick") = "window.open('FindItemPage.aspx','FindItemPage','status=0,toolbar=0,width=620,height=650,resizable=yes,scrollbars=yes');"
        DA001.Attributes.Add("onclick", "PriceVersion('A001')")    ' A001-PriceVersion
        DA206.Attributes.Add("onclick", "PriceVersion('A206')")    ' A206-PriceVersion
        DA211.Attributes.Add("onclick", "PriceVersion('A211')")    ' A211-PriceVersion
        DA999.Attributes.Add("onclick", "PriceVersion('A999')")    ' A999-PriceVersion
        DK206.Attributes.Add("onclick", "PriceVersion('K206')")    ' K206-PriceVersion
        DK211.Attributes.Add("onclick", "PriceVersion('K211')")    ' K211-PriceVersion
        BCustomer.Attributes.Add("onClick", "GetCustomer()") '找客戶
        BBuyer.Attributes.Add("onClick", "GetBuyer()") '找buyer
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BCheckItem)
    '**     1.輸入資料檢查
    '**     2.特殊要求重整
    '**     3.組合Item Name
    '**     4.檢查是否已申請之Item
    '**     5.檢查是否已存在之Item
    '**
    '*****************************************************************
    Protected Sub BCheckItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BCheckItem.Click
        '1.輸入資料檢查
        Dim InputError As Boolean = False
        If DataError() Then InputError = True
        '
        If Not InputError Then
            '2.特殊要求重整
            Dim str As String = DSRequest1.Text & "!" & _
                                DSRequest2.Text & "!" & _
                                DSRequest3.Text & "!" & _
                                DSRequest4.Text & "!" & _
                                DSRequest5.Text & "!" & _
                                DSRequest6.Text
            Dim xSRequest As String()
            xSRequest = fpObj.SPRequestSort(str).Split("!")
            '
            DSRequest1.Text = ""
            DSRequest2.Text = ""
            DSRequest3.Text = ""
            DSRequest4.Text = ""
            DSRequest5.Text = ""
            DSRequest6.Text = ""
            Dim j As Integer = 0
            For i As Integer = 0 To xSRequest.Length - 1
                If xSRequest(i) <> "" Then
                    Select Case j
                        Case 0
                            DSRequest1.Text = xSRequest(i)
                        Case 1
                            DSRequest2.Text = xSRequest(i)
                        Case 2
                            DSRequest3.Text = xSRequest(i)
                        Case 3
                            DSRequest4.Text = xSRequest(i)
                        Case 4
                            DSRequest5.Text = xSRequest(i)
                        Case Else
                            DSRequest6.Text = xSRequest(i)
                    End Select
                    j = j + 1
                End If
            Next
            '3.組合Item Name
            DItemName1.Text = ""
            DItemName2.Text = ""
            DItemName3.Text = ""
            'ItemName-1
            str = DSize.Text & "!" & DChain.Text & "!" & DClass.Text & "!" & DTape.Text & "!" & _
                  DSlider1.Text & "!" & DFinish1.Text & "!" & DSlider2.Text & "!" & DFinish2.Text & "!" & DFamily.Text

            DItemName1.Text = fpObj.SetItemName1(DClass.Text, str)

            If DClass.Text = "PS" And DST4.Text = "3" And DST5.Text = "1" And DST6.Text = "P" Then  'SLD
                DItemName1.Text = DItemName1.Text + " " + "PARTS"
            End If

            If DClass.Text = "CH" Or DClass.Text = "T" Then    'CHAIN
                If DST7.Text = "1" Then DItemName1.Text = DItemName1.Text & " " & "NAT"
                If DST7.Text = "2" Then DItemName1.Text = DItemName1.Text & " " & "SET"
                If DST7.Text = "3" Then
                    If InStr(DRItemName1.Text + " " + DRItemName2.Text + " " + DRItemName3.Text, " NAT") Then
                        DItemName1.Text = DItemName1.Text & " " & "NAT"
                    Else
                        DItemName1.Text = DItemName1.Text & " " & "DYED"
                    End If
                End If
            End If
            '特殊要求
            str = DSRequest1.Text & "!" & _
                  DSRequest2.Text & "!" & _
                  DSRequest3.Text & "!" & _
                  DSRequest4.Text & "!" & _
                  DSRequest5.Text & "!" & _
                  DSRequest6.Text
            xSRequest = str.Split("!")
            'ItemName1 > 35時處理
            str = ""
            If DItemName1.Text.Length > 35 Then
                If Mid(DItemName1.Text, 36, 1) = " " Then
                    str = Mid(DItemName1.Text, 37, DItemName1.Text.Length - 36)
                Else
                    str = Mid(DItemName1.Text, 36, DItemName1.Text.Length - 35)
                End If
                DItemName1.Text = Mid(DItemName1.Text, 1, 35)
            End If
            'ItemName2及3處理
            For i As Integer = 0 To xSRequest.Length - 1
                If xSRequest(i) <> "" Then
                    If str.Length + xSRequest(i).Length + 1 > 34 Then
                        DItemName2.Text = str
                        str = xSRequest(i)
                    Else
                        If str = "" Then
                            str = xSRequest(i)
                        Else
                            str = str & " " & xSRequest(i)
                        End If
                    End If
                End If
            Next
            If DItemName2.Text = "" Then
                DItemName2.Text = str
            Else
                DItemName3.Text = str
            End If
            '4.檢查是否已申請之Item
            Dim SQL As String = "Select No From F_ItemRegisterSheet " & _
                                "Where (Sts='0' or Sts='1') " & _
                                "  And ItemName1 = '" & DItemName1.Text.Trim & "' " & _
                                "  And ItemName2 = '" & DItemName2.Text.Trim & "' " & _
                                "  And ItemName3 = '" & DItemName3.Text.Trim & "' "
            Dim dtItemRegister As DataTable = uDataBase.GetDataTable(SQL)
            If dtItemRegister.Rows.Count > 0 Then
                If DNo.Text <> dtItemRegister.Rows(0)("No").ToString.Trim Then
                    DCode.Text = dtItemRegister.Rows(0)("No").ToString.Trim
                    DStatus.Value = "NG"
                    uJavaScript.PopMsg(Me, "[ItemRegister-1101]:此Item資料已經申請,請確認 ! ")
                Else
                    DStatus.Value = "OK"
                End If
            Else
                DStatus.Value = "OK"
            End If
            '5.檢查是否已存在之Item
            If DStatus.Value = "OK" Then
                DCode.Text = ""

                ' Modify-Start 2013/5/15
                'SQL = "SELECT ITEM From MST_ITEM "
                'SQL = SQL + "WHERE ITEM_NAME1 = '" & DItemName1.Text.Trim & "' "
                'SQL = SQL + "  AND ITEM_NAME2 = '" & DItemName2.Text.Trim & "' "
                'SQL = SQL + "  AND ITEM_NAME3 = '" & DItemName3.Text.Trim & "' "

                SQL = "SELECT ITEM From MST_ITEM "
                SQL = SQL + "WHERE Rtrim(Rtrim(ITEM_NAME1) + ' ' + Rtrim(ITEM_NAME2) + ' ' + Rtrim(ITEM_NAME3)) = '" & _
                                   Trim(DItemName1.Text.Trim & " " & DItemName2.Text.Trim & " " & DItemName3.Text.Trim) & "' "
                ' Modify-End

                Dim dtITEM As DataTable = uDataBase.GetDataTable(SQL)
                If dtITEM.Rows.Count > 0 Then
                    DCode.Text = dtITEM.Rows(0).Item("ITEM")
                    DStatus.Value = "NG"
                    uJavaScript.PopMsg(Me, "[ItemRegister-1102]:此Item資料已經存在(Waves),請確認 ! ")
                Else
                    DStatus.Value = "OK"
                End If
            End If
            '
            If DStatus.Value = "OK" Then
                DStatusItem.Value = DSize.Text + "!" + _
                                    DChain.Text + "!" + _
                                    DClass.Text + "!" + _
                                    DTape.Text + "!" + _
                                    DSlider1.Text + "!" + _
                                    DFinish1.Text + "!" + _
                                    DSlider2.Text + "!" + _
                                    DFinish2.Text + "!" + _
                                    DFinish2.Text + "!" + _
                                    DSRequest1.Text + "!" + _
                                    DSRequest2.Text + "!" + _
                                    DSRequest3.Text + "!" + _
                                    DSRequest4.Text + "!" + _
                                    DSRequest5.Text + "!" + _
                                    DSRequest6.Text
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(DataError)
    '**     檢查輸入資料
    '**
    '*****************************************************************
    Public Function DataError() As Boolean
        Dim uError As Boolean = False
        '
        DSize.Text = DSize.Text.ToUpper()
        DChain.Text = DChain.Text.ToUpper()
        DClass.Text = DClass.Text.ToUpper()
        DTape.Text = DTape.Text.ToUpper()
        DSlider1.Text = DSlider1.Text.ToUpper()
        DFinish1.Text = DFinish1.Text.ToUpper()
        DSlider2.Text = DSlider2.Text.ToUpper()
        DFinish2.Text = DFinish2.Text.ToUpper()
        DSRequest1.Text = DSRequest1.Text.ToUpper()
        DSRequest2.Text = DSRequest2.Text.ToUpper()
        DSRequest3.Text = DSRequest3.Text.ToUpper()
        DSRequest4.Text = DSRequest4.Text.ToUpper()
        DSRequest5.Text = DSRequest5.Text.ToUpper()
        DSRequest6.Text = DSRequest6.Text.ToUpper()
        '
        If Not uError Then
            If DSize.Text <> "" Then
                If InStr(DSize.Text, " ") > 0 Then uError = True
                If Not uError Then If Len(DSize.Text) > 2 Then uError = True
                If Not uError Then uError = CheckData("Size", DSize.Text)
                If uError Then uJavaScript.PopMsg(Me, "[" + DSize.Text + "]:Size ?,請確認 ! ")
            End If
        End If
        If Not uError Then
            If DChain.Text <> "" Then
                If InStr(DChain.Text, " ") > 0 Then uError = True
                If Not uError Then If Len(DChain.Text) > 6 Then uError = True
                If Not uError Then uError = CheckData("Chain", DChain.Text)
                If uError Then uJavaScript.PopMsg(Me, "[" + DChain.Text + "]:Chain Code ?,請確認 ! ")
            End If
        End If
        If Not uError Then
            If DClass.Text <> "" Then
                If InStr(DClass.Text, " ") > 0 Then uError = True
                If Not uError Then If Len(DClass.Text) > 2 Then uError = True
                If Not uError Then uError = CheckData("Class", DClass.Text)
                If uError Then uJavaScript.PopMsg(Me, "[" + DClass.Text + "]:Class ?,請確認 ! ")
            End If
        End If
        '-----------------------------------------------------
        If Not uError Then
            If DTape.Text <> "" Then
                If InStr(DTape.Text, " ") > 0 Then uError = True
                If Not uError Then If Len(DTape.Text) > 5 Then uError = True
                If uError Then uJavaScript.PopMsg(Me, "[" + DTape.Text + "]:Tape ? ,請確認 ! ")
            End If
        End If
        If Not uError Then
            If DSlider1.Text <> "" Then
                If InStr(DSlider1.Text, " ") > 0 Then uError = True
                If Not uError Then If Len(DSlider1.Text) > 10 Then uError = True
                If uError Then uJavaScript.PopMsg(Me, "[" + DSlider1.Text + "]:Slider-1 ? ,請確認 ! ")
            End If
        End If
        If Not uError Then
            If DFinish1.Text <> "" Then
                If InStr(DFinish1.Text, " ") > 0 Then uError = True
                If Not uError Then If Len(DFinish1.Text) > 3 Then uError = True
                If uError Then uJavaScript.PopMsg(Me, "[" + DFinish1.Text + "]:Finish-1 ? ,請確認 ! ")
            End If
        End If
        If Not uError Then
            If DSlider2.Text <> "" Then
                If InStr(DSlider2.Text, " ") > 0 Then uError = True
                If Not uError Then If Len(DSlider2.Text) > 10 Then uError = True
                If uError Then uJavaScript.PopMsg(Me, "[" + DSlider2.Text + "]:Slider-2 ? ,請確認 ! ")
            End If
        End If
        If Not uError Then
            If DFinish2.Text <> "" Then
                If InStr(DFinish2.Text, " ") > 0 Then uError = True
                If Not uError Then If Len(DFinish2.Text) > 3 Then uError = True
                If uError Then uJavaScript.PopMsg(Me, "[" + DFinish2.Text + "]:Finish-2 ? ,請確認 ! ")
            End If
        End If
        '-----------------------------------------------------
        If Not uError Then
            If DSRequest1.Text <> "" Then
                If InStr(DSRequest1.Text, " ") > 0 Then uError = True
                If Not uError Then If Len(DSRequest1.Text) > 10 Then uError = True
                If Not uError Then uError = CheckData("SR1", DSRequest1.Text)
                If uError Then uJavaScript.PopMsg(Me, "[" + DSRequest1.Text + "]:特殊要求-1 ?,請確認 ! ")
            End If
        End If
        If Not uError Then
            If DSRequest2.Text <> "" Then
                If InStr(DSRequest2.Text, " ") > 0 Then uError = True
                If Not uError Then If Len(DSRequest2.Text) > 10 Then uError = True
                If Not uError Then uError = CheckData("SR2", DSRequest2.Text)
                If uError Then uJavaScript.PopMsg(Me, "[" + DSRequest2.Text + "]:特殊要求-2 ?,請確認 ! ")
            End If
        End If
        If Not uError Then
            If DSRequest3.Text <> "" Then
                If InStr(DSRequest3.Text, " ") > 0 Then uError = True
                If Not uError Then If Len(DSRequest3.Text) > 10 Then uError = True
                If Not uError Then uError = CheckData("SR3", DSRequest3.Text)
                If uError Then uJavaScript.PopMsg(Me, "[" + DSRequest3.Text + "]:特殊要求-3 ?,請確認 ! ")
            End If
        End If
        If Not uError Then
            If DSRequest4.Text <> "" Then
                If InStr(DSRequest4.Text, " ") > 0 Then uError = True
                If Not uError Then If Len(DSRequest4.Text) > 10 Then uError = True
                If Not uError Then uError = CheckData("SR4", DSRequest4.Text)
                If uError Then uJavaScript.PopMsg(Me, "[" + DSRequest4.Text + "]:特殊要求-4 ?,請確認 ! ")
            End If
        End If
        If Not uError Then
            If DSRequest5.Text <> "" Then
                If InStr(DSRequest5.Text, " ") > 0 Then uError = True
                If Not uError Then If Len(DSRequest5.Text) > 10 Then uError = True
                If Not uError Then uError = CheckData("SR5", DSRequest5.Text)
                If uError Then uJavaScript.PopMsg(Me, "[" + DSRequest5.Text + "]:特殊要求-5 ?,請確認 ! ")
            End If
        End If
        If Not uError Then
            If DSRequest6.Text <> "" Then
                If InStr(DSRequest6.Text, " ") > 0 Then uError = True
                If Not uError Then If Len(DSRequest6.Text) > 10 Then uError = True
                If Not uError Then uError = CheckData("SR6", DSRequest6.Text)
                If uError Then uJavaScript.PopMsg(Me, "[" + DSRequest6.Text + "]:特殊要求-6 ?,請確認 ! ")
            End If
        End If
        '
        Return uError
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(CheckData)
    '**     檢查輸入資料
    '**
    '*****************************************************************
    '*********************************************************************
    Public Function CheckData(ByVal pCat As String, ByVal pData As String) As Boolean
        Dim uError As Boolean = False
        Dim SQL As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
        '
        cn.ConnectionString = ConnectString
        '
        If pData <> "" Then
            SQL = "SELECT CN1I09 FROM WAVEDLIB.C0900 "
            If pCat = "Size" Then SQL = SQL + "WHERE DGRC09 = 'SIZC' AND DDTC09 = '" & pData & "' "
            If pCat = "Chain" Then SQL = SQL + "WHERE DGRC09 = 'CHNC' AND DDTC09 = '" & pData & "' "
            If pCat = "Class" Then SQL = SQL + "WHERE DGRC09 = 'CLSC' AND DDTC09 = '" & pData & "' "
            If pCat = "SR1" Then SQL = SQL + "WHERE DGRC09 = 'SF1C' AND DDTC09 = '" & pData & "' "
            If pCat = "SR2" Then SQL = SQL + "WHERE DGRC09 = 'SF1C' AND DDTC09 = '" & pData & "' "
            If pCat = "SR3" Then SQL = SQL + "WHERE DGRC09 = 'SF1C' AND DDTC09 = '" & pData & "' "
            If pCat = "SR4" Then SQL = SQL + "WHERE DGRC09 = 'SF1C' AND DDTC09 = '" & pData & "' "
            If pCat = "SR5" Then SQL = SQL + "WHERE DGRC09 = 'SF1C' AND DDTC09 = '" & pData & "' "
            If pCat = "SR6" Then SQL = SQL + "WHERE DGRC09 = 'SF1C' AND DDTC09 = '" & pData & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, cn)
            DBAdapter1.Fill(ds, "C0900")
            If ds.Tables("C0900").Rows.Count <= 0 Then uError = True
            cn.Close()
        Else
            If pCat = "Size" Or pCat = "Chain" Or pCat = "Class" Then uError = True
        End If
        '
        Return uError
    End Function
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
                Dim wAllocateID As String = DSPDPerson.SelectedValue

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
            oCommon.SendMail()
            '
            If wbFormSno = 0 And wbStep < 3 Then    '判斷是否起單
                If DReRegister.Checked = True Then
                    uJavaScript.PopMsg(Me, "已成功送出您所填寫之申請單，請繼續申請！")
                    EnabledButton()                             '起動Button運作
                    DStatus.Value = "NG"
                    DNo.Text = ""
                    wFormSno = wbFormSno
                    wStep = wbStep
                Else
                    Dim URL As String = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                        "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                    Response.Redirect(URL)
                End If
            Else
                Dim URL As String = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                    "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
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
        Dim Path As String = Server.MapPath(uCommon.GetAppSetting("ItemRegisterFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim sql As String = ""
        sql = "INSERT INTO F_ItemRegisterSheet ( " & _
              "Sts, CompletedTime, FormNo, FormSno, " & _
              "No, Date, Name, JobTitle, Division, " & _
              "RCode, RItemName1, RItemName2, RItemName3, RSize, RChain, RClass, RTape, RSlider1, RFinish1, RSlider2, RFinish2, " & _
              "RSRequest1, RSRequest2, RSRequest3, RSRequest4, RSRequest5, RSRequest6, RFamily, " & _
              "RST1, RST2, RST3, RST4, RST5, RST6, RST7, RNoDisplay, " & _
              "Code, ItemName1, ItemName2, ItemName3, Size, Chain, Class, Tape, Slider1, Finish1, Slider2, Finish2, " & _
              "SRequest1, SRequest2, SRequest3, SRequest4, SRequest5, SRequest6, Family, " & _
              "ST1, ST2, ST3, ST4, ST5, ST6, ST7, NoDisplay, " & _
              "SLDPrice, Remark, Attachfile1, A001, A206, A211, A999, K206, K211,Buyer,BuyerCode,Customer,Customercode,ForUse, " & _
              "CreateUser, CreateTime, ModifyUser, ModifyTime) " & _
        "VALUES("
        sql &= "'0' ,"
        sql &= "'" & NowDateTime & "' ,"
        sql &= "'" & wFormNo & "' ,"
        sql &= "'" & CStr(NewFormSno) & "' ,"
        sql &= "'" & DNo.Text & "' ,"
        sql &= "'" & DDate.Text & "' ,"
        sql &= "'" & DName.Text & "' ,"
        sql &= "'" & DJobTitle.Text & "' ,"
        sql &= "'" & DDivision.Text & "' ,"
        '
        sql &= "'" & DRCode.Text & "' ,"
        sql &= "'" & DRItemName1.Text & "' ,"
        sql &= "'" & DRItemName2.Text & "' ,"
        sql &= "'" & DRItemName3.Text & "' ,"
        sql &= "'" & DRSize.Text & "' ,"
        sql &= "'" & DRChain.Text & "' ,"
        sql &= "'" & DRClass.Text & "' ,"
        sql &= "'" & DRTape.Text & "' ,"
        sql &= "'" & DRSlider1.Text & "' ,"
        sql &= "'" & DRFinish1.Text & "' ,"
        sql &= "'" & DRSlider2.Text & "' ,"
        sql &= "'" & DRFinish2.Text & "' ,"
        sql &= "'" & DRSRequest1.Text & "' ,"
        sql &= "'" & DRSRequest2.Text & "' ,"
        sql &= "'" & DRSRequest3.Text & "' ,"
        sql &= "'" & DRSRequest4.Text & "' ,"
        sql &= "'" & DRSRequest5.Text & "' ,"
        sql &= "'" & DRSRequest6.Text & "' ,"
        sql &= "'" & DRFamily.Text & "' ,"
        sql &= "'" & DRST1.Text & "' ,"
        sql &= "'" & DRST2.Text & "' ,"
        sql &= "'" & DRST3.Text & "' ,"
        sql &= "'" & DRST4.Text & "' ,"
        sql &= "'" & DRST5.Text & "' ,"
        sql &= "'" & DRST6.Text & "' ,"
        sql &= "'" & DRST7.Text & "' ,"
        sql &= "'" & DRNoDisplay.Text & "' ,"
        '
        sql &= "'" & DCode.Text & "' ,"
        sql &= "'" & DItemName1.Text & "' ,"
        sql &= "'" & DItemName2.Text & "' ,"
        sql &= "'" & DItemName3.Text & "' ,"
        sql &= "'" & DSize.Text & "' ,"
        sql &= "'" & DChain.Text & "' ,"
        sql &= "'" & DClass.Text & "' ,"
        sql &= "'" & DTape.Text & "' ,"
        sql &= "'" & DSlider1.Text & "' ,"
        sql &= "'" & DFinish1.Text & "' ,"
        sql &= "'" & DSlider2.Text & "' ,"
        sql &= "'" & DFinish2.Text & "' ,"
        sql &= "'" & DSRequest1.Text & "' ,"
        sql &= "'" & DSRequest2.Text & "' ,"
        sql &= "'" & DSRequest3.Text & "' ,"
        sql &= "'" & DSRequest4.Text & "' ,"
        sql &= "'" & DSRequest5.Text & "' ,"
        sql &= "'" & DSRequest6.Text & "' ,"
        sql &= "'" & DFamily.Text & "' ,"
        sql &= "'" & DST1.Text & "' ,"
        sql &= "'" & DST2.Text & "' ,"
        sql &= "'" & DST3.Text & "' ,"
        sql &= "'" & DST4.Text & "' ,"
        sql &= "'" & DST5.Text & "' ,"
        sql &= "'" & DST6.Text & "' ,"
        sql &= "'" & DST7.Text & "' ,"
        sql &= "'" & DNoDisplay.Text & "' ,"
        sql &= "'" & DSLDPrice.Text & "' ,"
        '
        sql &= "N'" & DRemark.Text & "' ,"
        Dim FileName As String = ""
        If DAttachfile1.Visible Then
            If DAttachfile1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Attachfile1" & "-" & UploadDateTime & "-" & Right(DAttachfile1.PostedFile.FileName, InStr(StrReverse(DAttachfile1.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "Attachfile1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DAttachfile1.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DAttachfile1.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "
        '
        If DA001.Checked = True Then
            sql &= "1 ,"
        Else
            sql &= "0 ,"
        End If
        If DA206.Checked = True Then
            sql &= "1 ,"
        Else
            sql &= "0 ,"
        End If
        If DA211.Checked = True Then
            sql &= "1 ,"
        Else
            sql &= "0 ,"
        End If
        If DA999.Checked = True Then
            sql &= "1 ,"
        Else
            sql &= "0 ,"
        End If
        If DK206.Checked = True Then
            sql &= "1 ,"
        Else
            sql &= "0 ,"
        End If
        If DK211.Checked = True Then
            sql &= "1 ,"
        Else
            sql &= "0 ,"
        End If

        '20160128 Jessica Modify
        sql &= "'" & YKK.ReplaceString(DBuyer.Text) & "' ,"
        sql &= "'" & DBuyerCode.Text & "' ,"
        sql &= "'" & YKK.ReplaceString(DCustomer.Text) & "' ,"
        sql &= "'" & DCustomerCode.Text & "' ,"
        sql &= "'" & DForUse.SelectedValue & "' ,"

        '
        sql &= "'" & Request.QueryString("pUserID") & "' ,"
        sql &= "'" & NowDateTime & "' ,"
        sql &= "'" & Request.QueryString("pUserID") & "' ,"
        sql &= "'" & NowDateTime & "' )"

        uDataBase.ExecuteNonQuery(sql)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim Path As String = Server.MapPath(uCommon.GetAppSetting("ItemRegisterFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim sql As String

        sql = "Update F_ItemRegisterSheet Set "
        If pFun <> "SAVE" Then
            sql &= " Sts = '" & pSts & "',"
            sql &= " CompletedTime = '" & NowDateTime & "',"
        End If
        '
        If DA001.Checked = True Then
            sql &= " A001 = 1 ,"
        Else
            sql &= " A001 = 0 ,"
        End If
        If DA206.Checked = True Then
            sql &= " A206 = 1 ,"
        Else
            sql &= " A206 = 0 ,"
        End If
        If DA211.Checked = True Then
            sql &= " A211 = 1 ,"
        Else
            sql &= " A211 = 0 ,"
        End If
        If DA999.Checked = True Then
            sql &= " A999 = 1 ,"
        Else
            sql &= " A999 = 0 ,"
        End If
        If DK206.Checked = True Then
            sql &= " K206 = 1 ,"
        Else
            sql &= " K206 = 0 ,"
        End If
        If DK211.Checked = True Then
            sql &= " K211 = 1 ,"
        Else
            sql &= " K211 = 0 ,"
        End If
        '
        sql &= " Code = '" & DCode.Text & "',"
        sql &= " ItemName1 = '" & DItemName1.Text & "',"
        sql &= " ItemName2 = '" & DItemName2.Text & "',"
        sql &= " ItemName3 = '" & DItemName3.Text & "',"
        sql &= " Size = '" & DSize.Text & "',"
        sql &= " Chain = '" & DChain.Text & "',"
        sql &= " Class = '" & DClass.Text & "',"
        sql &= " Tape = '" & DTape.Text & "',"
        sql &= " Slider1 = '" & DSlider1.Text & "',"
        sql &= " Finish1 = '" & DFinish1.Text & "',"
        sql &= " Slider2 = '" & DSlider2.Text & "',"
        sql &= " Finish2 = '" & DFinish2.Text & "',"
        sql &= " SRequest1 = '" & DSRequest1.Text & "',"
        sql &= " SRequest2 = '" & DSRequest2.Text & "',"
        sql &= " SRequest3 = '" & DSRequest3.Text & "',"
        sql &= " SRequest4 = '" & DSRequest4.Text & "',"
        sql &= " SRequest5 = '" & DSRequest5.Text & "',"
        sql &= " SRequest6 = '" & DSRequest6.Text & "',"
        sql &= " Family = '" & DFamily.Text & "',"
        sql &= " ST1 = '" & DST1.Text & "',"
        sql &= " ST2 = '" & DST2.Text & "',"
        sql &= " ST3 = '" & DST3.Text & "',"
        sql &= " ST4 = '" & DST4.Text & "',"
        sql &= " ST5 = '" & DST5.Text & "',"
        sql &= " ST6 = '" & DST6.Text & "',"
        sql &= " ST7 = '" & DST7.Text & "',"
        sql &= " NoDisplay = '" & DNoDisplay.Text & "',"
        sql &= " SLDPrice = '" & DSLDPrice.Text & "',"

        If wStep = 500 And pFun = "OK" Then  '退回重選要更新參考ITEM
            sql &= " RCode = '" & DRCode.Text & "',"
            sql &= " RItemName1 = '" & DRItemName1.Text & "',"
            sql &= " RItemName2 = '" & DRItemName2.Text & "',"
            sql &= " RItemName3 = '" & DRItemName3.Text & "',"
            sql &= " RSize = '" & DRSize.Text & "',"
            sql &= " RChain = '" & DRChain.Text & "',"
            sql &= " RClass = '" & DRClass.Text & "',"
            sql &= " RTape = '" & DRTape.Text & "',"
            sql &= " RSlider1 = '" & DRSlider1.Text & "',"
            sql &= " RFinish1 = '" & DRFinish1.Text & "',"
            sql &= " RSlider2 = '" & DRSlider2.Text & "',"
            sql &= " RFinish2 = '" & DRFinish2.Text & "',"
            sql &= " RSRequest1 = '" & DRSRequest1.Text & "',"
            sql &= " RSRequest2 = '" & DRSRequest2.Text & "',"
            sql &= " RSRequest3 = '" & DRSRequest3.Text & "',"
            sql &= " RSRequest4 = '" & DRSRequest4.Text & "',"
            sql &= " RSRequest5 = '" & DRSRequest5.Text & "',"
            sql &= " RSRequest6 = '" & DRSRequest6.Text & "',"
            sql &= " RFamily = '" & DRFamily.Text & "',"
            sql &= " RST1 = '" & DRST1.Text & "',"
            sql &= " RST2 = '" & DRST2.Text & "',"
            sql &= " RST3 = '" & DRST3.Text & "',"
            sql &= " RST4 = '" & DRST4.Text & "',"
            sql &= " RST5 = '" & DRST5.Text & "',"
            sql &= " RST6 = '" & DRST6.Text & "',"
            sql &= " RST7 = '" & DRST7.Text & "',"
            sql &= " RNoDisplay = '" & DRNoDisplay.Text & "',"


        End If
      
        sql &= " Remark = N'" & DRemark.Text & "',"
        '
        Dim FileName As String = ""
        If DAttachfile1.Visible Then
            If DAttachfile1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Attachfile1" & "-" & UploadDateTime & "-" & Right(DAttachfile1.PostedFile.FileName, InStr(StrReverse(DAttachfile1.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Attachfile1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DAttachfile1.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DAttachfile1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql &= " AttachFile1 = '" & FileName & "',"
            End If
        End If

        'Jessia 20160128 Modify
        sql &= " Buyer = N'" & YKK.ReplaceString(DBuyer.Text) & "',"
        sql &= " BuyerCode = N'" & DBuyerCode.Text & "',"
        sql &= " Customer = N'" & YKK.ReplaceString(DCustomer.Text) & "',"
        sql &= " CustomerCode = N'" & DCustomerCode.Text & "',"
        sql &= " ForUse = N'" & DForUse.SelectedValue & "',"
        '
        sql &= " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        sql &= " ModifyTime = '" & NowDateTime & "' "
        sql &= " Where FormNo  =  '" & wFormNo & "'"
        sql &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        '
        uDataBase.ExecuteNonQuery(sql)
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
        If InputDataOK(3) Then
            DisabledButton()   '停止Button運作
            ModifyData("SAVE", "0")           '更新表單資料 Sts=0(未結)
            ModifyTranData("SAVE", "0")       '更新交易資料
            Dim URL As String = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=S&pFormNo=" & wFormNo & "&pStep=" & wStep & _
                                                "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
            Response.Redirect(URL)
        Else
            EnabledButton()   '起動Button運作
        End If
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
            FlowControl("NG2", 2, "3")
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
            FlowControl("NG1", 1, "2")
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
        If InputDataOK(0) Then
            DisabledButton()   '停止Button運作
            FlowControl("OK", 0, "1")
        Else
            EnabledButton()   '起動Button運作
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     判斷是否可繼續執行(驗證資料)
    '**
    '*****************************************************************
    Function InputDataOK(ByVal pAction As Integer) As Boolean
        Dim isOK As Boolean = False
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim Message As String = ""

        'Item Inf.檢測結果
        If DStatus.Value <> "OK" Then
            '**Modify-Start 2012/7/5
            'If wStep = 500 And pAction = 2 Then
            'Else
            '    ErrCode = 9010
            'End If
            If wStep = 500 Then
            Else
                ErrCode = 9010
            End If
            '**Modify-End
        End If
        'Check最後檢測OKItem
        If ErrCode = 0 Then
            If wStep = 10 And pAction = 3 Then  '儲存
            Else
                If DStatusItem.Value <> DSize.Text.ToUpper + "!" + _
                                        DChain.Text.ToUpper + "!" + _
                                        DClass.Text.ToUpper + "!" + _
                                        DTape.Text.ToUpper + "!" + _
                                        DSlider1.Text.ToUpper + "!" + _
                                        DFinish1.Text.ToUpper + "!" + _
                                        DSlider2.Text.ToUpper + "!" + _
                                        DFinish2.Text.ToUpper + "!" + _
                                        DFinish2.Text.ToUpper + "!" + _
                                        DSRequest1.Text.ToUpper + "!" + _
                                        DSRequest2.Text.ToUpper + "!" + _
                                        DSRequest3.Text.ToUpper + "!" + _
                                        DSRequest4.Text.ToUpper + "!" + _
                                        DSRequest5.Text.ToUpper + "!" + _
                                        DSRequest6.Text.ToUpper Then
                    ErrCode = 9060
                End If
            End If
        End If
        'Check最終資料(雙空白+ItemName2(35)+Sort)
        If ErrCode = 0 Then
            Dim str As String = DItemName1.Text + DItemName2.Text + DItemName3.Text
            If InStr(str, "  ") > 0 Then ErrCode = 9070
            If ErrCode = 0 Then If Len(DItemName2.Text) > 34 Then ErrCode = 9071
        End If
        'Check上傳附件Size及格式
        If ErrCode = 0 Then
            If DAttachfile1.Visible Then
                If DAttachfile1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DAttachfile1)
                End If
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
                ErrCode = oCommon.CommissionNo("001151", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9050
                End If
            End If
        End If

        If ErrCode <> 0 Then
            If ErrCode = 9010 Then Message = "Item資料有誤,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案格式有誤,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案Size有誤,請確認!"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9050 Then Message = "委託書No.重覆,請確認委託書No.!"
            If ErrCode = 9060 Then Message = "修改後未重新檢測資料,請確認!"
            If ErrCode = 9070 Then Message = "發現有空白重覆,請重新檢測資料,請確認!"
            If ErrCode = 9071 Then Message = "Item Name(2)字串過長(>34),請重新檢測資料,請確認!"
            If ErrCode = 9072 Then Message = "發現特殊要求中有未排序資料,請重新檢測資料,請確認!"
            uJavaScript.PopMsg(Me, Message)
        Else
            isOK = True
        End If
        '
        Return isOK
    End Function
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
        Dim i As Integer

        'Set當日日期
        Str2 = CStr(DateTime.Now.Month)  '月
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        Str = Str2
        Str2 = CStr(DateTime.Now.Day)    '日
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        Str = Str + Str2
        'Set單號
        Str1 = CStr(Seq)
        For i = Str1.Length To 10 - 1
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
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".ai", ".xls", ".doc", ".ppt", ".xlsx"}   '定義允許的檔案格式
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
                UPFileIsNormal = 9030
            End If
        End If
    End Function
   

End Class
