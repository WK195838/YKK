Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.IO

Partial Class ItemRegisterSheet_03
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
    Dim LimitedId, strNewRequestBfSort, sWarnMsg As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900      ' 設定逾時時間
        SetParameter()                  ' 設定共用參數
        TopPosition()                   ' 設定top
        SetPopupFunction()              ' 設定彈出視窗事件
        ShowSheetField("New")           ' 表單欄位顯示及欄位輸入檢查
        ShowSheetFunction()             ' 表單功能按鈕顯示
        If Not IsPostBack Then

            '2024/07/08 DOrderDate改為客戶了解使用 (Emily)
            'ADD-START ISOS-2308 PJ
            '受注日計算
            'DOrderDate.Text = DateAdd("m", 4, CDate(NowDateTime).ToString("yyyy/MM/dd"))
            'DOrderDate.Text = Mid(DOrderDate.Text, 1, 7) & "/1"
            'DOrderDate.Text = Replace(DOrderDate.Text, "//", "/")
            'DOrderDate.Text = DateAdd("d", -1, CDate(DOrderDate.Text))
            'HOrderDate.Text = DOrderDate.Text
            '
            'ADD-END

            'jessica 2021/11/25 
            If wCode <> "" Then
                GetItemCode(wCode)
            End If
            ' NOTICE PAGE
            Response.Write("<script>window.open('http://10.245.1.6/IRW/IRWNoticePage.aspx','','height=1000,width=800,menubar=no,location=no,resizable=yes,scrollbars=yes');</script>")
            '
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
            Top = 1080
            Dim SQL As String = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(SQL)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0)("BEndTime").ToString < NowDateTime Then
                    If DDelay.Visible = True Then
                        Top = 1108
                    End If
                End If
            End If
        Else
            Top = 1000
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
        '
        'JOY-ADD
        'DSPDNOUrl.Style("left") = -500 & "px"

        DSUITABLECHECK.Style("left") = -500 & "px"
        DITEMSUITABLEFile.Style("left") = -500 & "px"
        DITEMSUITABLEFile.Text = "\\10.245.0.192\program$\ISOS\DataPrepare\" + "ITEMSUITABLEIRW_" & wApplyID & ".xlsm"
        BSUITABLE.Attributes("onclick") = "CheckAttribute('ITEMSUITABLEExcel')"

        'ADD-START ISOS-2308 PJ
        '販賣實績-中止
        '欄位隱藏
        DMark9.Style("left") = -500 & "px"
        HOrderDate.Style("left") = -500 & "px"
        DSUITABLECHECK.Text = "OK"
        '
        'ADD-END

        'ADD-START LIMITEDITEM
        If UCase(Request.QueryString("pUserID")) <> "IT003" And _
           UCase(Request.QueryString("pUserID")) <> "MK028" And UCase(Request.QueryString("pUserID")) <> "MK045" Then
            BSPDNO.Style("left") = -500 & "px"
        End If
        'ADD-END
        '
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
                Top = 1080
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
            Top = 954
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
        If Not IsPostBack Then
            LAttachfile2.Visible = False
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
            DCustUnd.Visible = False
            DMark1.Visible = False
            BSUITABLE.Visible = False
            DMark9.Visible = False
        End If
        'EDX追加
        If Not IsPostBack Then
            LEDXLink1.Visible = False
            LEDXRegister.Visible = False
            DEDXComment1.Visible = False
            DEDXComment2.Visible = False
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
    '**     表單欄位顯示及欄位輸入檢查
    '**
    '*****************************************************************
    Sub ShowSheetField(ByVal pPost As String)
        '建置欄位及屬性陣列

        '2022/08/04 針對以下三人從流程2開始 
        If wStep = 1 Then
            If Request.QueryString("pUserID") = "mk028" Or Request.QueryString("pUserID") = "mk019" Or Request.QueryString("pUserID") = "mk035" Then
                wStep = 2
            End If
        End If

        'ADD-START ISOS-2308 PJ
        '違反者採用流程2開始
        ' xx MK057 林依儒 --> MK010 劉明德
        ' xx td003 梁菀筑 --> MK010 劉明德
        ' xx sl057 李宜樺 --> MK014 吳松惠
        ' xx tn006 李偉群 --> tn002 史文雄
        ' xx sl056 張家綱 --> MK014 吳松惠
        ' xx mk038 游佳蓁 --> MK010 劉明德 --> MK003 蔡逸星
        '
        'tc006	蘇宗毅 --> tn002 史文雄
        'mk058	莊若楨 --> MK010 劉明德 
        '250208
        'tc006	蘇宗毅	2	mk003	蔡逸星
        'tc023	廖彥凱	1	tn002	史文雄
        'mk006	周宇婕	1	mk010	劉明德
        'sl044	張凱鈞	1	mk013	劉玉慧
        '
        If wStep = 1 Then
            If Request.QueryString("pUserID") = "tc006" Then
                wStep = 888
            End If
            'If Request.QueryString("pUserID") = "tc023" Then
            '    wStep = 3
            'End If
            'If Request.QueryString("pUserID") = "mk006" Then
            '    wStep = 3
            'End If
            'If Request.QueryString("pUserID") = "sl044" Then
            '    wStep = 3
            'End If
        End If
        'ADD-END

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
            '
            'ADD-START 2311-BUYER&ITEM限定
            If D1.Text <> DCustomerCode.Text Then
            End If
            If LimitedItemError("CUSTOMER") = True Then
                DStatus.Value = "NG"
                DCustomer.Text = ""
                DBuyer.Text = ""
                DCustomerCode.Text = ""
                DBuyerCode.Text = ""
                D1.Text = ""
                D2.Text = ""
            Else
                DStatus.Value = "OK"
            End If

            'ADD-END
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
            DA999.Enabled = True
            DA001.Enabled = True
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
                        '判斷若特殊要求有開發中產品keyword，則取消勾選單價A001/A999
                        SetDevPrice()
                        '  uJavaScript.PopMsg(Me, "請確認若為開發中產品(SLS-ANY、CH-ANY、TAPE-ANY、ILYTEMP、I-YTEMP、ILTP***、I-TP***)不可勾選A001或A999。")
                    Else
                        DA999.Checked = False
                        DA001.Checked = False
                    End If
                Else
                    DA999.Checked = False
                    DA001.Checked = False
                    '統計區分4<>1，不可勾選單價A001/A999
                    DA999.Enabled = False
                    DA001.Enabled = False

                    DMemo.Text = ""
                End If
            End If
            '
            'ADD-START 2311-BUYER&ITEM限定
            If D2.Text <> DBuyerCode.Text Then
            End If
            If LimitedItemError("BUYER") = True Then
                DStatus.Value = "NG"
                DCustomer.Text = ""
                DBuyer.Text = ""
                DCustomerCode.Text = ""
                DBuyerCode.Text = ""
                D1.Text = ""
                D2.Text = ""
            Else
                DStatus.Value = "OK"
            End If
            'ADD-END
            '
            'PERFORMACE-START
            'ADD-START ISOS-2308 PJ
            'EDX CHECK
            'If EDXNotFound("BUYER") = True Then
            '    DStatus.Value = "NG"
            '    uJavaScript.PopMsg(Me, "搜尋不到[EDX]，請確認是否已完成 ! ")
            'Else
            '    DStatus.Value = "OK"
            'End If
            'ADD-END
            'PERFORMACE-END
            '
            '2023052:    jessica(繼續申請)
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

        'Attachfile2
        Select Case FindFieldInf("Attachfile2")
            Case 0  '顯示
                DAttachfile2.Visible = False
                DAttachfile2.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DAttachfile2Rqd", "DAttachfile2", "異常：需輸入憑證")
                DAttachfile2.Visible = True
                DAttachfile2.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DAttachfile2.Visible = True
                DAttachfile2.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DAttachfile2.Visible = False
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


        'SPDNO
        Select Case FindFieldInf("SPDNO")
            Case 0  '顯示
                DSPDNO.BackColor = Color.LightGray
                DSPDNO.ReadOnly = True
                DSPDNO.Visible = True

            Case 1  '修改+檢查
                DSPDNO.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSPDNORqd", "DSPDNO", "異常：需輸入SPDNO")
                DSPDNO.Visible = True
                DSPDNO.ReadOnly = True

            Case 2  '修改
                DSPDNO.BackColor = Color.Yellow
                DSPDNO.Visible = True

            Case Else   '隱藏
                DSPDNO.Visible = False

        End Select

        'ADD-START ISOS-2308 PJ
        'SPDNO
        Select Case FindFieldInf("OrderDate")
            Case 0  '顯示
                DOrderDate.BackColor = Color.LightGray
                DOrderDate.ReadOnly = True
                DOrderDate.Visible = True

            Case 1  '修改+檢查
                DOrderDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOrderDateRqd", "DOrderDate", "異常：需輸入預定受注日")
                DOrderDate.Visible = True
            Case 2  '修改
                DOrderDate.BackColor = Color.Yellow
                DOrderDate.Visible = True

            Case Else   '隱藏
                DOrderDate.Visible = False

        End Select
        '
        'ADD-END

        '精準檢測
        '
        DSUITABLECHECK.Visible = True
        DSUITABLECHECK.BackColor = Color.GreenYellow
        DSUITABLECHECK.ReadOnly = False

        'SPDNO
        '
        '  DSPDNO.BackColor = Color.LightGray
        'DSPDNO.Visible = True
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
            DSPDNO.Text = dtItemRegisterSheet.Rows(0).Item("SPDNO")                     ' SPDNO
          

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

            If dtItemRegisterSheet.Rows(0).Item("Attachfile2") <> "" Then
                LAttachfile2.NavigateUrl = Path & dtItemRegisterSheet.Rows(0).Item("Attachfile2")   '憑證
                LAttachfile2.Visible = True
            Else
                LAttachfile2.Visible = False
            End If


            DBuyer.Text = dtItemRegisterSheet.Rows(0).Item("Buyer")                         ' BUYER
            DBuyerCode.Text = dtItemRegisterSheet.Rows(0).Item("BuyerCode")                         ' BUYER
            DCustomer.Text = dtItemRegisterSheet.Rows(0).Item("Customer")                         ' Customer
            DCustomerCode.Text = dtItemRegisterSheet.Rows(0).Item("CustomerCode")                         ' Customer
            SetFieldData("ForUse", dtItemRegisterSheet.Rows(0).Item("ForUse"))    '用途
            '
            'ADD-START ISOS-2308 PJ
            'Order Date
            DOrderDate.Text = dtItemRegisterSheet.Rows(0).Item("OrderDate")     ' Order Date
            DSPDNO.Text = dtItemRegisterSheet.Rows(0).Item("SPDNo")             ' SPD No
            DSPDNOUrl.Text = dtItemRegisterSheet.Rows(0).Item("SPDUrl")         ' SPD URL
            '
            'ADD-END
            '
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

            '在20工程強制先檢查ITEM 20220826 JESSICA
            If wStep = 20 Then
                CheckItem()
            End If

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
    '**     6.模具報廢之Item
    '**
    '*****************************************************************
    Protected Sub BCheckItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BCheckItem.Click
        '20230502
        '上次備存
        If DCustomerCode.Text <> "" Or DBuyerCode.Text <> "" Then
            HCustomer.Text = DCustomer.Text
            HBuyer.Text = DBuyer.Text
            HCustomerCode.Text = DCustomerCode.Text
            HBuyerCode.Text = DBuyerCode.Text
        End If
        '換ITEM清空
        DCustomer.Text = ""
        DBuyer.Text = ""
        DCustomerCode.Text = ""
        DBuyerCode.Text = ""
        D1.Text = ""
        D2.Text = ""
        DLimitedItem.Text = ""
        DWarnMsg.Value = ""
        DCustUnd.Checked = False
        DCustUnd.Disabled = True
        DOrderDate.Text = ""
        DProdReady.Text = ""
        'ADD-START 20250527
        DItemInf.Text = ""
        DNewItemInf.Text = ""
        'ADD-END
        '
        'DISABLE Customer & Buyer BUTTON
        BCustomer.Visible = True
        BBuyer.Visible = True
        BLoadCustomerBuyer.Visible = True
        '
        '特別要求縮碼-ADD-START 2025/02/27 
        If Left(DSRequest1.Text, 3) = "S2I" Then
            Dim SQL1 As String
            SQL1 = "SELECT * "
            SQL1 = SQL1 & "From M_SpecialShortIRW "
            SQL1 = SQL1 & "WHERE ShortNo = '" & DSRequest1.Text.Trim & "' "
            SQL1 = SQL1 & "ORDER BY ShortNo "
            Dim dtS2I As DataTable = uDataBase.GetDataTable(SQL1)
            If dtS2I.Rows.Count > 0 Then
                DSRequest1.Text = dtS2I.Rows(0).Item("Special1")
                DSRequest2.Text = dtS2I.Rows(0).Item("Special2")
                DSRequest3.Text = dtS2I.Rows(0).Item("Special3")
                DSRequest4.Text = dtS2I.Rows(0).Item("Special4")
                DSRequest5.Text = dtS2I.Rows(0).Item("Special5")
                DSRequest6.Text = dtS2I.Rows(0).Item("Special6")
            End If
        End If
        '特別要求縮碼-ADD-END
        '
        '紀錄排序前的特殊要求資料，用來做Blue-Sign檢查
        strNewRequestBfSort = GetFullName(DSRequest1.Text) & "!" & _
                       GetFullName(DSRequest2.Text) & "!" & _
                       GetFullName(DSRequest3.Text) & "!" & _
                       GetFullName(DSRequest4.Text) & "!" & _
                       GetFullName(DSRequest5.Text) & "!" & _
                       GetFullName(DSRequest6.Text) & "!"

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
            'JOYJOY-TEST-241019
            'DStatus.Value = "OK"

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
            'JOYJOY-TEST-241019
            'DStatus.Value = "OK"

            '6.模具報廢之Item (JOY-ADD)
            'Slider-1
            If DStatus.Value = "OK" Then
                If Trim(DSlider1.Text) <> "" Then
                    '
                    If GetPuller(DSlider1.Text) <> "" Then
                        SQL = "SELECT Body, Puller, Size + '/' + Body + '/' + Puller As ErrMsg From V_IRWMoldCancel "
                        SQL = SQL & "WHERE Active = 1 "
                        SQL = SQL & "AND Size = '" & Trim(DSize.Text) & "' "
                        SQL = SQL & "AND Puller = '" & Trim(DSPDNOUrl.Text) & "' "
                        'SQL = SQL & "AND '" & Trim(DSlider1.Text) & "' like '%' + Puller + '%' "
                        SQL = SQL & "Order by SIZE, BODY, PULLER "
                        Dim dtMold1 As DataTable = uDataBase.GetDataTable(SQL)
                        If dtMold1.Rows.Count > 0 Then
                            For i As Integer = 0 To dtMold1.Rows.Count - 1
                                'MsgBox(dtMold1.Rows(i)("errmsg"))
                                If Mid(Trim(DSlider1.Text), 1, Len(dtMold1.Rows(i)("Body"))) = dtMold1.Rows(i)("Body") Then
                                    DStatus.Value = "NG"
                                    uJavaScript.PopMsg(Me, "[ItemRegister-1103]:此Item(Slider1)相關模具已經廢除,請確認 ! ")
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
            End If
            'Slider-2
            If DStatus.Value = "OK" Then
                If Trim(DSlider2.Text) <> "" Then
                    '
                    If GetPuller(DSlider2.Text) <> "" Then
                        SQL = "SELECT Body, Puller, Size + '/' + Body + '/' + Puller As ErrMsg From V_IRWMoldCancel "
                        SQL = SQL & "WHERE Active = 1 "
                        SQL = SQL & "AND Size = '" & Trim(DSize.Text) & "' "
                        SQL = SQL & "AND Puller = '" & Trim(DSPDNOUrl.Text) & "' "
                        SQL = SQL & "Order by SIZE, BODY, PULLER "
                        Dim dtMold2 As DataTable = uDataBase.GetDataTable(SQL)
                        If dtMold2.Rows.Count > 0 Then
                            For i As Integer = 0 To dtMold2.Rows.Count - 1
                                'MsgBox(dtMold2.Rows(i)("errmsg"))
                                If Mid(Trim(DSlider2.Text), 1, Len(dtMold2.Rows(i)("Body"))) = dtMold2.Rows(i)("Body") Then
                                    DStatus.Value = "NG"
                                    uJavaScript.PopMsg(Me, "[ItemRegister-1103]:此Item(Slider2)相關模具已經廢除,請確認 ! ")
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
            End If
            '-------------
            '檢查Blue-Sign，新Item的特殊要求有BN-ANTI時，參考Item必須要有N-ANTI
            If DStatus.Value = "OK" Then
                If (DClass.Text.Trim) <> "CH" Then
                    '
                    If InStr(strNewRequestBfSort, "BN-ANTI!") > 0 Then
                        '
                        Dim strNewRequest As String
                        Dim strRefRequest As String
                        strRefRequest = GetFullName(DRSRequest1.Text) & "!" & _
                           GetFullName(DRSRequest2.Text) & "!" & _
                           GetFullName(DRSRequest3.Text) & "!" & _
                           GetFullName(DRSRequest4.Text) & "!" & _
                           GetFullName(DRSRequest5.Text) & "!" & _
                           GetFullName(DRSRequest6.Text) & "!"
                        If InStr(strRefRequest, "N-ANTI!") = 0 Then
                            DLimitedItem.Text = "參考Item非N-ANTI!"
                            uJavaScript.PopMsg(Me, "參考Item非N-ANTI!")
                            '將Item資料存入table以便日後統計
                            AppendLimitedItem(False)
                            DStatus.Value = "NG"
                        Else
                            '判斷SIZE~特殊要求-6要相同(除N-ANTI跟BN-ANTI外)
                            strNewRequest = Replace(strNewRequestBfSort, "BN-ANTI!", "")
                            strRefRequest = Replace(strRefRequest, "N-ANTI!", "")
                            strNewRequest = DSize.Text + "!" + _
                                        DChain.Text + "!" + _
                                        DClass.Text + "!" + _
                                        DTape.Text + "!" + _
                                        DSlider1.Text + "!" + _
                                        DFinish1.Text + "!" + _
                                        DSlider2.Text + "!" + _
                                        DFinish2.Text + "!" + _
                                        strNewRequest
                            strRefRequest = DRSize.Text + "!" + _
                                        DRChain.Text + "!" + _
                                        DRClass.Text + "!" + _
                                        DRTape.Text + "!" + _
                                        DRSlider1.Text + "!" + _
                                        DRFinish1.Text + "!" + _
                                        DRSlider2.Text + "!" + _
                                        DRFinish2.Text + "!" + _
                                        strRefRequest
                            If Trim(Replace(strNewRequest.ToUpper(), " ", "")) <> Trim(Replace(strRefRequest.ToUpper(), " ", "")) Then
                                DLimitedItem.Text = "除了BN-ANTI 其他欄位品名需同參考ITEM!" & Chr(13) & strNewRequest & " : " & strRefRequest
                                uJavaScript.PopMsg(Me, "除了BN-ANTI 其他欄位品名需同參考ITEM!")
                                '將Item資料存入table以便日後統計
                                AppendLimitedItem(False)
                                DStatus.Value = "NG"
                            End If
                        End If
                        '
                    End If
                    '
                End If
            End If
            '--------------
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
            '
            'MODIFY-START ISOS-2308 PJ
            '適用ITEM
            '----------------
            'ITEM限定/BUYER限定 
            'MODIFY-END
            If DStatus.Value = "OK" Then
                '
                Dim UploadDateTime As String = DateTime.Now.ToString("yyyyMMddHHmmss")
                '
                '[適用ITEM] CLEAR
                SQL = "UPDATE F_ItemSuitableSheet SET Sts = 1, "
                SQL = SQL & "RMDSIM = '" & "未執行" & "', "
                SQL = SQL & "ModifyUser = '" & Request.QueryString("pUserID") & "', "
                SQL = SQL & "ModifyTime = '" & NowDateTime & "' "
                SQL = SQL & "WHERE Sts = 0 "
                SQL = SQL & "AND CreateUser = '" & Request.QueryString("pUserID") & "' "
                uDataBase.ExecuteNonQuery(SQL)
                '
                '[適用ITEM]資料存入
                SQL = "INSERT INTO F_ItemSuitableSheet ( " & _
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
                SQL &= "'0' ,"
                SQL &= "'" & NowDateTime & "' ,"
                SQL &= "'" & "Z01151" & "' ,"
                SQL &= "'" & CStr(1) & "' ,"
                SQL &= "'" & UploadDateTime & "' ,"
                SQL &= "'" & DDate.Text & "' ,"
                SQL &= "'" & DName.Text & "' ,"
                SQL &= "'" & DJobTitle.Text & "' ,"
                SQL &= "'" & DDivision.Text & "' ,"
                '
                SQL &= "'" & DRCode.Text & "' ,"
                SQL &= "'" & DRItemName1.Text & "' ,"
                SQL &= "'" & DRItemName2.Text & "' ,"
                SQL &= "'" & DRItemName3.Text & "' ,"
                SQL &= "'" & DRSize.Text & "' ,"
                SQL &= "'" & DRChain.Text & "' ,"
                SQL &= "'" & DRClass.Text & "' ,"
                SQL &= "'" & DRTape.Text & "' ,"
                SQL &= "'" & DRSlider1.Text & "' ,"
                SQL &= "'" & DRFinish1.Text & "' ,"
                SQL &= "'" & DRSlider2.Text & "' ,"
                SQL &= "'" & DRFinish2.Text & "' ,"
                SQL &= "'" & DRSRequest1.Text & "' ,"
                SQL &= "'" & DRSRequest2.Text & "' ,"
                SQL &= "'" & DRSRequest3.Text & "' ,"
                SQL &= "'" & DRSRequest4.Text & "' ,"
                SQL &= "'" & DRSRequest5.Text & "' ,"
                SQL &= "'" & DRSRequest6.Text & "' ,"
                SQL &= "'" & DRFamily.Text & "' ,"
                SQL &= "'" & DRST1.Text & "' ,"
                SQL &= "'" & DRST2.Text & "' ,"
                SQL &= "'" & DRST3.Text & "' ,"
                SQL &= "'" & DRST4.Text & "' ,"
                SQL &= "'" & DRST5.Text & "' ,"
                SQL &= "'" & DRST6.Text & "' ,"
                SQL &= "'" & DRST7.Text & "' ,"
                SQL &= "'" & DRNoDisplay.Text & "' ,"
                '
                SQL &= "'" & DCode.Text & "' ,"
                SQL &= "'" & DItemName1.Text & "' ,"
                SQL &= "'" & DItemName2.Text & "' ,"
                SQL &= "'" & DItemName3.Text & "' ,"
                SQL &= "'" & DSize.Text & "' ,"
                SQL &= "'" & DChain.Text & "' ,"
                SQL &= "'" & DClass.Text & "' ,"
                SQL &= "'" & DTape.Text & "' ,"
                SQL &= "'" & DSlider1.Text & "' ,"
                SQL &= "'" & DFinish1.Text & "' ,"
                SQL &= "'" & DSlider2.Text & "' ,"
                SQL &= "'" & DFinish2.Text & "' ,"
                SQL &= "'" & DSRequest1.Text & "' ,"
                SQL &= "'" & DSRequest2.Text & "' ,"
                SQL &= "'" & DSRequest3.Text & "' ,"
                SQL &= "'" & DSRequest4.Text & "' ,"
                SQL &= "'" & DSRequest5.Text & "' ,"
                SQL &= "'" & DSRequest6.Text & "' ,"
                SQL &= "'" & DFamily.Text & "' ,"
                SQL &= "'" & DST1.Text & "' ,"
                SQL &= "'" & DST2.Text & "' ,"
                SQL &= "'" & DST3.Text & "' ,"
                SQL &= "'" & DST4.Text & "' ,"
                SQL &= "'" & DST5.Text & "' ,"
                SQL &= "'" & DST6.Text & "' ,"
                SQL &= "'" & DST7.Text & "' ,"
                SQL &= "'" & DNoDisplay.Text & "' ,"
                SQL &= "'" & DSLDPrice.Text & "' ,"
                '
                SQL &= "N'" & DRemark.Text & "' ,"
                SQL = SQL + " '" + "" + "', "
                '
                If DA001.Checked = True Then
                    SQL &= "1 ,"
                Else
                    SQL &= "0 ,"
                End If
                If DA206.Checked = True Then
                    SQL &= "1 ,"
                Else
                    SQL &= "0 ,"
                End If
                If DA211.Checked = True Then
                    SQL &= "1 ,"
                Else
                    SQL &= "0 ,"
                End If
                If DA999.Checked = True Then
                    SQL &= "1 ,"
                Else
                    SQL &= "0 ,"
                End If
                If DK206.Checked = True Then
                    SQL &= "1 ,"
                Else
                    SQL &= "0 ,"
                End If
                If DK211.Checked = True Then
                    SQL &= "1 ,"
                Else
                    SQL &= "0 ,"
                End If

                '20160128 Jessica Modify
                SQL &= "'" & YKK.ReplaceString(DBuyer.Text) & "' ,"
                SQL &= "'" & DBuyerCode.Text & "' ,"
                SQL &= "'" & YKK.ReplaceString(DCustomer.Text) & "' ,"
                SQL &= "'" & DCustomerCode.Text & "' ,"
                SQL &= "'" & DForUse.SelectedValue & "' ,"

                '
                SQL &= "'" & Request.QueryString("pUserID") & "' ,"
                SQL &= "'" & NowDateTime & "' ,"
                SQL &= "'" & Request.QueryString("pUserID") & "' ,"
                SQL &= "'" & NowDateTime & "' )"
                uDataBase.ExecuteNonQuery(SQL)

                'MODIFY-START ISOS-2308 PJ
                '販賣實績-中止
                'DSUITABLECHECK.Text = ""
                DSUITABLECHECK.Text = "OK"
                '
                '生產NG-檢測(JOY)
                'If ItemNotProd("ITEM") Then
                '    '
                '    'DSPDNO.Text = ""
                DSPDNO.Text = "NOT-WINGSITEM"
                '    DSPDNOUrl.Text = ""
                '    '
                '    '生產NG
                '    Dim ColorStr, KeepStr, xSize, p1, p2 As String

                '    If CDbl(DSize.Text) > 9 Then
                '        xSize = DSize.Text
                '    Else
                '        xSize = Replace(DSize.Text, "0", "")
                '    End If
                '    '
                '    If DSlider1.Text <> "" Then
                '        KeepStr = ""
                '        ColorStr = ""
                '        str = DSlider1.Text
                '        j = Len(DSlider1.Text)
                '        '**
                '        Do Until j < 0 Or ColorStr = DSlider1.Text
                '            '** 

                '            If Mid(str, j, 1) >= "0" And Mid(str, j, 1) <= "9" Then
                '                KeepStr = Mid(str, 1, j)
                '                Exit Do
                '            Else
                '                ColorStr = Mid(str, j, 1) & ColorStr
                '            End If
                '            j = j - 1
                '        Loop
                '        '**
                '        If ColorStr = DSlider1.Text Then KeepStr = ColorStr
                '        '**
                '        p1 = xSize & "-" & DFamily.Text & "-" & KeepStr
                '        p2 = xSize & "-" & DFamily.Text & "-" & DSlider1.Text
                '        '
                '        '2022/8/5 ADD-START
                '        p1 = Replace(p1, "3-C-DSV", "3-C-DSB")
                '        p1 = Replace(p1, "5-CZ-DA7", "5-CN-DA8")
                '        p1 = Replace(p1, "5-CZ-DAG7", "5-CN-DAG8")
                '        '
                '        p1 = Replace(p1, "5-CI", "5-CN")
                '        p1 = Replace(p1, "5-CZ", "5-CN")
                '        p1 = Replace(p1, "5-CS", "5-CN")
                '        p1 = Replace(p1, "5-Y", "5-M")
                '        '
                '        p1 = Replace(p1, "3-CZ", "3-C")
                '        p1 = Replace(p1, "3-CS", "3-C")
                '        'ADD-END
                '        '
                '        'MODIFY-START LIMITEDITEM
                '        'BSPDNO.Attributes("onclick") = "window.open('" & _
                '        '   "FindSPDPage.aspx?pUserID=" & wApplyID & "&pSLDKey1=" & p2 & "&pSLDKey2=" & DSlider1.Text & "&pFINKey1=" & p1 & "&pFINKey2=" & KeepStr & _
                '        '   "','FindSPDPage','resizable=yes,scrollbars=yes');"
                '        'MODIFY-END
                '    End If
                'Else
                '    '不是生產NG
                '    DSPDNO.Text = "WINGSITEM"
                'End If
                'MODIFY-END
                '
                '
                'ADD-START 2311-BUYER&ITEM限定
                If LimitedItemError("ITEM") = True Then
                    DStatus.Value = "NG"
                    DCustomer.Text = ""
                    DBuyer.Text = ""
                    DCustomerCode.Text = ""
                    DBuyerCode.Text = ""
                    D1.Text = ""
                    D2.Text = ""
                    '
                    'DISABLE Customer & Buyer BUTTON
                    BCustomer.Visible = False
                    BBuyer.Visible = False
                    BLoadCustomerBuyer.Visible = False
                Else
                    DStatus.Value = "OK"
                End If
            End If

            '20240705-增加Warning的檢查 (Emily)
            'DStatus.Value = "OK"
            If DStatus.Value = "OK" Then
                LimitedItemError("WARNING")
                'Warning，不做DStatus.Value設定，由sWarnMsg=""及DCustUnd勾選來判斷是否OK
            End If
            '
            'JOY-ADD 20241203 joyjoy
            DPerformaceUp.Text = ""
            If DStatus.Value = "OK" Then
                '
                If PerformaceUp("IMPFINISH") = True Then
                    DPerformaceUp.Text = "IMPFINISH"
                    DPerformaceUp.Style("left") = -500 & "px"
                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[進口品檢測-Start]"
                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "ITEM = 進口品"
                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[進口品-End]"
                End If
                '
            End If
            'JOY-END
            '
            'JOY-ADD 20240800
            '--20240808-Slider10字縮寫判定
            If DStatus.Value = "OK" Then
                '
                If DPerformaceUp.Text <> "IMPFINISH" Then
                    '
                    If LimitedItemOther("SLIDERSHORT") = True Then
                        DStatus.Value = "NG"
                        DCustomer.Text = ""
                        DBuyer.Text = ""
                        DCustomerCode.Text = ""
                        DBuyerCode.Text = ""
                        D1.Text = ""
                        D2.Text = ""
                        '
                        'DISABLE Customer & Buyer BUTTON
                        BCustomer.Visible = False
                        BBuyer.Visible = False
                        BLoadCustomerBuyer.Visible = False
                        '---
                        'DStatus.Value = "OK"
                    Else
                        DStatus.Value = "OK"
                    End If
                    '
                Else
                    DStatus.Value = "OK"
                End If
                '
            End If
            '
            '--20240819-生產可行性
            If DStatus.Value = "OK" Then
                '
                If LimitedItemOther("PRODREADY") = True Then
                    DProdReady.Text = "NG"
                Else
                    DProdReady.Text = "OK"
                End If
                '
                'TEST
                'DProdReady.Text = "OK"
                DStatus.Value = "OK"
                '
            End If
            '
            'DProdReady.Text = "OK"
            '--20240820-表面處理
            'If DStatus.Value = "OK" And DProdReady.Text = "OK" Then
            If DStatus.Value = "OK" Then
                '
                If DPerformaceUp.Text <> "IMPFINISH" Then
                    '
                    If LimitedItemOther("FINISH") = True Then
                        DStatus.Value = "NG"
                        DCustomer.Text = ""
                        DBuyer.Text = ""
                        DCustomerCode.Text = ""
                        DBuyerCode.Text = ""
                        D1.Text = ""
                        D2.Text = ""
                        '
                        'DISABLE Customer & Buyer BUTTON
                        BCustomer.Visible = False
                        BBuyer.Visible = False
                        BLoadCustomerBuyer.Visible = False
                        '---
                        'DStatus.Value = "OK"
                    Else
                        DStatus.Value = "OK"
                    End If
                    '
                Else
                    DStatus.Value = "OK"
                End If
                '
            End If
            '
            '--20240907-P-PACK
            'If DStatus.Value = "OK" And DProdReady.Text = "OK" Then
            If DStatus.Value = "OK" Then
                '
                If DPerformaceUp.Text <> "IMPFINISH" Then
                    '
                    If LimitedItemOther("PPACK") = True Then
                        DStatus.Value = "NG"
                        DCustomer.Text = ""
                        DBuyer.Text = ""
                        DCustomerCode.Text = ""
                        DBuyerCode.Text = ""
                        D1.Text = ""
                        D2.Text = ""
                        '
                        'DISABLE Customer & Buyer BUTTON
                        BCustomer.Visible = False
                        BBuyer.Visible = False
                        BLoadCustomerBuyer.Visible = False
                        '---
                        'DStatus.Value = "OK"
                    Else
                        DStatus.Value = "OK"
                    End If
                    '
                Else
                    DStatus.Value = "OK"
                End If
                '
            End If
            '
            '--20240907-TWIST4
            'If DStatus.Value = "OK" And DProdReady.Text = "OK" Then
            If DStatus.Value = "OK" Then
                '
                If DPerformaceUp.Text <> "IMPFINISH" Then
                    '
                    If LimitedItemOther("TWIST4") = True Then
                        DStatus.Value = "NG"
                        DCustomer.Text = ""
                        DBuyer.Text = ""
                        DCustomerCode.Text = ""
                        DBuyerCode.Text = ""
                        D1.Text = ""
                        D2.Text = ""
                        '
                        'DISABLE Customer & Buyer BUTTON
                        BCustomer.Visible = False
                        BBuyer.Visible = False
                        BLoadCustomerBuyer.Visible = False
                        '---
                        'DStatus.Value = "OK"
                    Else
                        DStatus.Value = "OK"
                    End If
                    '
                Else
                    DStatus.Value = "OK"
                End If
                '
            End If
            '
            '--20240907-SCD
            'If DStatus.Value = "OK" And DProdReady.Text = "OK" Then
            If DStatus.Value = "OK" Then
                '
                If DPerformaceUp.Text <> "IMPFINISH" Then
                    '
                    If LimitedItemOther("SCD") = True Then
                        DStatus.Value = "NG"
                        DCustomer.Text = ""
                        DBuyer.Text = ""
                        DCustomerCode.Text = ""
                        DBuyerCode.Text = ""
                        D1.Text = ""
                        D2.Text = ""
                        '
                        'DISABLE Customer & Buyer BUTTON
                        BCustomer.Visible = False
                        BBuyer.Visible = False
                        BLoadCustomerBuyer.Visible = False
                        '---
                        'DStatus.Value = "OK"
                    Else
                        DStatus.Value = "OK"
                    End If
                    '
                Else
                    DStatus.Value = "OK"
                End If
                '
            End If
            '
            '--20241009-特別要求縮寫判定
            If DStatus.Value = "OK" Then
                '
                If DPerformaceUp.Text <> "IMPFINISH" Then
                    '
                    If LimitedItemOther("SPECIALSHORT") = True Then
                        DStatus.Value = "NG"
                        DCustomer.Text = ""
                        DBuyer.Text = ""
                        DCustomerCode.Text = ""
                        DBuyerCode.Text = ""
                        D1.Text = ""
                        D2.Text = ""
                        '
                        BCustomer.Visible = False
                        BBuyer.Visible = False
                        BLoadCustomerBuyer.Visible = False
                        '---
                        'DStatus.Value = "OK"
                    Else
                        DStatus.Value = "OK"
                    End If
                    '
                Else
                    DStatus.Value = "OK"
                End If
                '
            End If
            '
            '--20241025-檢針
            If DStatus.Value = "OK" Then
                '
                If DPerformaceUp.Text <> "IMPFINISH" Then
                    '
                    If LimitedItemOther("SLIDERKENSIN") = True Then
                        DStatus.Value = "NG"
                        DCustomer.Text = ""
                        DBuyer.Text = ""
                        DCustomerCode.Text = ""
                        DBuyerCode.Text = ""
                        D1.Text = ""
                        D2.Text = ""
                        '
                        'DISABLE Customer & Buyer BUTTON
                        BCustomer.Visible = False
                        BBuyer.Visible = False
                        BLoadCustomerBuyer.Visible = False
                        '---
                        'DStatus.Value = "OK"
                    Else
                        DStatus.Value = "OK"
                    End If

                Else
                    DStatus.Value = "OK"
                End If
                '
            End If
            '
            '--20241109-型別
            If DStatus.Value = "OK" Then
                '
                If DPerformaceUp.Text <> "IMPFINISH" Then
                    '
                    If LimitedItemOther("SLIDERSPEC") = True Then
                        DStatus.Value = "NG"
                        DCustomer.Text = ""
                        DBuyer.Text = ""
                        DCustomerCode.Text = ""
                        DBuyerCode.Text = ""
                        D1.Text = ""
                        D2.Text = ""
                        '
                        'DISABLE Customer & Buyer BUTTON
                        BCustomer.Visible = False
                        BBuyer.Visible = False
                        BLoadCustomerBuyer.Visible = False
                        '---
                        'DStatus.Value = "OK"
                    Else
                        DStatus.Value = "OK"
                    End If
                    '
                Else
                    DStatus.Value = "OK"
                End If
                '
            End If
            '
            '--20241109-MFGAP
            If DStatus.Value = "OK" Then
                '
                If DPerformaceUp.Text <> "IMPFINISH" Then
                    '
                    If LimitedItemOther("MFGAP") = True Then
                        DStatus.Value = "NG"
                        DCustomer.Text = ""
                        DBuyer.Text = ""
                        DCustomerCode.Text = ""
                        DBuyerCode.Text = ""
                        D1.Text = ""
                        D2.Text = ""
                        '
                        'DISABLE Customer & Buyer BUTTON
                        BCustomer.Visible = False
                        BBuyer.Visible = False
                        BLoadCustomerBuyer.Visible = False
                        '---
                        'DStatus.Value = "OK"
                    Else
                        DStatus.Value = "OK"
                    End If
                    '
                Else
                    DStatus.Value = "OK"
                End If
                '
            End If
            '
            'ADD-START 20250527
            If DStatus.Value = "OK" Then
                DItemInf.Text = UCase(DSize.Text.Trim) & "!" & _
                                UCase(DChain.Text.Trim) & "!" & _
                                UCase(DClass.Text.Trim) & "!" & _
                                UCase(DTape.Text.Trim) & "!" & _
                                UCase(DSlider1.Text.Trim) & "!" & _
                                UCase(DFinish1.Text.Trim) & "!" & _
                                UCase(DSlider2.Text.Trim) & "!" & _
                                UCase(DFinish2.Text.Trim) & "!" & _
                                UCase(DFamily.Text.Trim) & "!" & _
                                UCase(DSRequest1.Text.Trim) & "!" & _
                                UCase(DSRequest2.Text.Trim) & "!" & _
                                UCase(DSRequest3.Text.Trim) & "!" & _
                                UCase(DSRequest4.Text.Trim) & "!" & _
                                UCase(DSRequest5.Text.Trim) & "!" & _
                                UCase(DSRequest6.Text.Trim) & "!"
                DItemInf.Text = Replace(DItemInf.Text, "XXX!", "")
            End If
            'ADD-END

        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(EDXNotFound)
    '**     檢查EDX
    '**
    '*****************************************************************
    Public Function EDXNotFound(ByVal pCat As String) As Boolean
        'MsgBox("EDXNotFound-IN")
        Dim i, j, x, y As Integer
        Dim sql, str, str1, xItemName, xSpc, xSlider, xPuller, xColor, xBody, xBodyGr, xEDXCheck, xEDXVendor, xSubno As String
        Dim xSource As String = ""
        Dim uError As Boolean = False
        Dim iNew As String()
        Dim iMember As String()
        Dim iSlider As String()
        Dim iFinish As String()
        '
        Dim cn As New OleDbConnection
        Dim ds, ds1 As New DataSet
        Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
        '
        LEDXLink1.Visible = False
        LEDXRegister.Visible = False
        DEDXComment1.Visible = False
        DEDXComment2.Visible = False
        '
        xItemName = DItemName1.Text & " " & DItemName2.Text & " " & DItemName3.Text
        '
        xSpc = DSRequest1.Text & "!" & _
               DSRequest2.Text & "!" & _
               DSRequest3.Text & "!" & _
               DSRequest4.Text & "!" & _
               DSRequest5.Text & "!" & _
               DSRequest6.Text & "!"
        '
        str = DSize.Text & "!" & _
                DChain.Text & "!" & _
                DClass.Text & "!" & _
                DTape.Text & "!" & _
                DSlider1.Text & "!" & _
                DFinish1.Text & "!" & _
                DSlider2.Text & "!" & _
                DFinish2.Text & "!" & _
                DSRequest1.Text & "!" & _
                DFamily.Text & "!" & _
                DST1.Text & "!" & _
                DST2.Text & "!" & _
                DST3.Text & "!" & _
                DST4.Text & "!" & _
                DST5.Text & "!" & _
                DST6.Text & "!" & _
                DST7.Text & "!"
        '
        iNew = str.ToString.Split("!")
        '
        'Slider
        str = ""
        If iNew(4) <> "" Then str = iNew(4) & "/"
        If iNew(6) <> "" Then str = str & iNew(6) & "/"
        '
        'Finish
        str1 = ""
        If iNew(5) <> "" Then str1 = iNew(5) & "/"
        If iNew(7) <> "" Then str1 = str1 & iNew(7) & "/"
        '
        If str = "" Then
            DSPDNOUrl.Text = ""
        Else
            If InStr(DSPDNOUrl.Text, Chr(13)) > 0 Then
                DSPDNOUrl.Text = Mid(DSPDNOUrl.Text, 1, InStr(DSPDNOUrl.Text, Chr(13)) - 1) & Chr(13) & pCat
            Else
                DSPDNOUrl.Text = DSPDNOUrl.Text & Chr(13) & pCat
            End If
        End If
        '------------------------------------------
        If str <> "" Then
            iSlider = str.ToString.Split("/")
            iFinish = str1.ToString.Split("/")
            xSubno = "ZZZZZZ"
            'DOrderDate.Text = ""
            '
            For i = 0 To UBound(iSlider) - 1
                xSource = ""
                xSlider = ""
                xPuller = ""
                xBody = ""
                ds.Clear()
                ds1.Clear()
                '
                'A.3Y SALES & FA000 (M_WINGSEDX)
                sql = "SELECT Top 1 Slider, OrderNo, Item "
                sql = sql & "From M_WINGSEDX "
                sql = sql & "Where Slider = '" & iSlider(i) & "' "
                '特別要求PL- 05 CN DA7BA C5 PARTS KENSIN N-ANTI PL-Y749
                Select Case i + 1
                    Case 1
                        If InStr(xSpc, "!PL-") > 0 Or InStr(xSpc, "!T-PL-") > 0 Or Mid(xSpc, 1, 3) = "PL-" Or Mid(xSpc, 1, 5) = "T-PL-" Then
                            If Mid(DSRequest1.Text, 1, 3) = "PL-" Or Mid(DSRequest1.Text, 1, 5) = "T-PL-" Then sql = sql & "AND ItemName Like '%" & " " & DSRequest1.Text & "%' "
                            If Mid(DSRequest2.Text, 1, 3) = "PL-" Or Mid(DSRequest2.Text, 1, 5) = "T-PL-" Then sql = sql & "AND ItemName Like '%" & " " & DSRequest2.Text & "%' "
                            If Mid(DSRequest3.Text, 1, 3) = "PL-" Or Mid(DSRequest3.Text, 1, 5) = "T-PL-" Then sql = sql & "AND ItemName Like '%" & " " & DSRequest3.Text & "%' "
                            If Mid(DSRequest4.Text, 1, 3) = "PL-" Or Mid(DSRequest4.Text, 1, 5) = "T-PL-" Then sql = sql & "AND ItemName Like '%" & " " & DSRequest4.Text & "%' "
                            If Mid(DSRequest5.Text, 1, 3) = "PL-" Or Mid(DSRequest5.Text, 1, 5) = "T-PL-" Then sql = sql & "AND ItemName Like '%" & " " & DSRequest5.Text & "%' "
                            If Mid(DSRequest6.Text, 1, 3) = "PL-" Or Mid(DSRequest6.Text, 1, 5) = "T-PL-" Then sql = sql & "AND ItemName Like '%" & " " & DSRequest6.Text & "%' "
                        End If
                        '
                    Case Else
                        If InStr(xSpc, "!PL-") > 0 Or InStr(xSpc, "!B-PL-") > 0 Or Mid(xSpc, 1, 3) = "PL-" Or Mid(xSpc, 1, 5) = "B-PL-" Then
                            If Mid(DSRequest1.Text, 1, 3) = "PL-" Or Mid(DSRequest1.Text, 1, 5) = "B-PL-" Then sql = sql & "AND ItemName Like '%" & " " & DSRequest1.Text & "%' "
                            If Mid(DSRequest2.Text, 1, 3) = "PL-" Or Mid(DSRequest2.Text, 1, 5) = "B-PL-" Then sql = sql & "AND ItemName Like '%" & " " & DSRequest2.Text & "%' "
                            If Mid(DSRequest3.Text, 1, 3) = "PL-" Or Mid(DSRequest3.Text, 1, 5) = "B-PL-" Then sql = sql & "AND ItemName Like '%" & " " & DSRequest3.Text & "%' "
                            If Mid(DSRequest4.Text, 1, 3) = "PL-" Or Mid(DSRequest4.Text, 1, 5) = "B-PL-" Then sql = sql & "AND ItemName Like '%" & " " & DSRequest4.Text & "%' "
                            If Mid(DSRequest5.Text, 1, 3) = "PL-" Or Mid(DSRequest5.Text, 1, 5) = "B-PL-" Then sql = sql & "AND ItemName Like '%" & " " & DSRequest5.Text & "%' "
                            If Mid(DSRequest6.Text, 1, 3) = "PL-" Or Mid(DSRequest6.Text, 1, 5) = "B-PL-" Then sql = sql & "AND ItemName Like '%" & " " & DSRequest6.Text & "%' "
                        End If
                        '
                End Select
                sql = sql & "Order By OrderNo desc, Item desc "
                Dim dtSales1 As DataTable = uDataBase.GetDataTable(sql)
                If dtSales1.Rows.Count > 0 Then
                    '
                    ' GET WINGS SUPPLIER
                    str = ""
                    str1 = ""
                    cn.ConnectionString = ConnectString
                    sql = "SELECT C.FL1I39||'/'||C.CLNC39||'/'||A.SIZCA0||A.FMLCA0||'/'||D.LN1CA1||'-'||D.LN2CA1||'/'||A.ITMCA0||'/' AS FL1I39 "
                    sql = sql & "FROM WAVEDLIB.FA000 A, WAVEDLIB.FL700 B, WAVEDLIB.S3900 C, WAVEDLIB.FZ030 D "
                    sql = sql & "WHERE A.ITMCA0 = B.ITMCL7 And A.ITMCA0 = D.ITMCA1 And B.CLNCL7 = C.CLNC39 "
                    sql = sql & "AND TRIM(A.SIZCA0) = '" & DSize.Text & "' "
                    sql = sql & "AND TRIM(A.FMLCA0) = '" & ReplaceFamily(DSize.Text, DFamily.Text) & "' "
                    sql = sql & "AND TRIM(A.SLDCA0) = '" & iSlider(i) & "' "
                    sql = sql & "AND TRIM(A.SFNCA0) = '" & iFinish(i) & "' "
                    sql = sql & "AND A.NDPCA0 = '" & "" & "' "
                    sql = sql & "AND (D.LN1CA1 LIKE '59%' OR D.LN1CA1='364') "
                    sql = sql & "ORDER BY B.RADUL7 DESC, B.RADTL7 DESC "
                    '
                    Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                    DBAdapter1.SelectCommand.CommandTimeout = 180
                    DBAdapter1.Fill(ds, "OUTSIDE")
                    If ds.Tables("OUTSIDE").Rows.Count > 0 Then
                        str = Replace(ds.Tables("OUTSIDE").Rows(0).Item("FL1I39"), " ", "")
                    Else
                        sql = "SELECT C.FL1I39||'/'||C.CLNC39||'/'||A.SIZCA0||A.FMLCA0||'/'||D.LN1CA1||'-'||D.LN2CA1||'/'||A.ITMCA0||'/' AS FL1I39 "
                        sql = sql & "FROM WAVEDLIB.FA000 A, WAVEDLIB.FL700 B, WAVEDLIB.S3900 C, WAVEDLIB.FZ030 D "
                        sql = sql & "WHERE A.ITMCA0 = B.ITMCL7 And A.ITMCA0 = D.ITMCA1 And B.CLNCL7 = C.CLNC39 "
                        sql = sql & "AND TRIM(A.SLDCA0) = '" & iSlider(i) & "' "
                        sql = sql & "AND TRIM(A.SFNCA0) = '" & iFinish(i) & "' "
                        sql = sql & "AND A.NDPCA0 = '" & "" & "' "
                        sql = sql & "AND (D.LN1CA1 LIKE '59%' OR D.LN1CA1='364') "
                        sql = sql & "ORDER BY B.RADUL7 DESC, B.RADTL7 DESC "
                        '
                        Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
                        DBAdapter2.SelectCommand.CommandTimeout = 180
                        DBAdapter2.Fill(ds1, "OUTSIDE1")
                        If ds1.Tables("OUTSIDE1").Rows.Count > 0 Then
                            str1 = Replace(ds1.Tables("OUTSIDE1").Rows(0).Item("FL1I39"), " ", "")
                        End If
                    End If
                    cn.Close()
                    '
                    xSource = "WINGS"
                    If str <> "" Or str1 <> "" Then
                        DSPDNOUrl.Text = DSPDNOUrl.Text & Chr(13) & "(W)SLD-" & CStr(i + 1) & "[" & dtSales1.Rows(0)("OrderNo") & "/" & dtSales1.Rows(0)("Item") & "]"
                        If str <> "" Then DSPDNOUrl.Text = DSPDNOUrl.Text & Chr(13) & DSize.Text & "-" & ReplaceFamily(DSize.Text, DFamily.Text) & "-" & iSlider(i) & ":" & "w_Supplier:|" & str & "|"
                        If str1 <> "" Then DSPDNOUrl.Text = DSPDNOUrl.Text & Chr(13) & dtSales1.Rows(0)("Item") & ":" & "w_Supplier:|" & str1 & "|"
                    Else
                        DSPDNOUrl.Text = DSPDNOUrl.Text & Chr(13) & "(W)SLD-" & CStr(i + 1) & "[" & dtSales1.Rows(0)("OrderNo") & "/" & dtSales1.Rows(0)("Item") & "]"
                    End If
                End If
                '
                If xSource = "" Or xSource <> "" Then
                    '
                    '3Y SALES = 無
                    'C.EDX DB (M_EDX)
                    '特別要求PL- 05 CN DA7BA C5 PARTS KENSIN N-ANTI PL-Y749
                    xPuller = ""
                    Select Case i + 1
                        Case 1
                            If InStr(xSpc, "!PL-") > 0 Or InStr(xSpc, "!T-PL-") > 0 Or Mid(xSpc, 1, 3) = "PL-" Or Mid(xSpc, 1, 5) = "T-PL-" Then
                                If Mid(DSRequest1.Text, 1, 3) = "PL-" Or Mid(DSRequest1.Text, 1, 5) = "T-PL-" Then xPuller = DSRequest1.Text
                                If Mid(DSRequest2.Text, 1, 3) = "PL-" Or Mid(DSRequest2.Text, 1, 5) = "T-PL-" Then xPuller = DSRequest2.Text
                                If Mid(DSRequest3.Text, 1, 3) = "PL-" Or Mid(DSRequest3.Text, 1, 5) = "T-PL-" Then xPuller = DSRequest3.Text
                                If Mid(DSRequest4.Text, 1, 3) = "PL-" Or Mid(DSRequest4.Text, 1, 5) = "T-PL-" Then xPuller = DSRequest4.Text
                                If Mid(DSRequest5.Text, 1, 3) = "PL-" Or Mid(DSRequest5.Text, 1, 5) = "T-PL-" Then xPuller = DSRequest5.Text
                                If Mid(DSRequest6.Text, 1, 3) = "PL-" Or Mid(DSRequest6.Text, 1, 5) = "T-PL-" Then xPuller = DSRequest6.Text
                            End If
                            '
                        Case Else
                            If InStr(xSpc, "!PL-") > 0 Or InStr(xSpc, "!B-PL-") > 0 Or Mid(xSpc, 1, 3) = "PL-" Or Mid(xSpc, 1, 5) = "B-PL-" Then
                                If Mid(DSRequest1.Text, 1, 3) = "PL-" Or Mid(DSRequest1.Text, 1, 5) = "B-PL-" Then xPuller = DSRequest1.Text
                                If Mid(DSRequest2.Text, 1, 3) = "PL-" Or Mid(DSRequest2.Text, 1, 5) = "B-PL-" Then xPuller = DSRequest2.Text
                                If Mid(DSRequest3.Text, 1, 3) = "PL-" Or Mid(DSRequest3.Text, 1, 5) = "B-PL-" Then xPuller = DSRequest3.Text
                                If Mid(DSRequest4.Text, 1, 3) = "PL-" Or Mid(DSRequest4.Text, 1, 5) = "B-PL-" Then xPuller = DSRequest4.Text
                                If Mid(DSRequest5.Text, 1, 3) = "PL-" Or Mid(DSRequest5.Text, 1, 5) = "B-PL-" Then xPuller = DSRequest5.Text
                                If Mid(DSRequest6.Text, 1, 3) = "PL-" Or Mid(DSRequest6.Text, 1, 5) = "B-PL-" Then xPuller = DSRequest6.Text
                            End If
                            '
                    End Select
                    If xPuller <> "" Then
                        xPuller = Replace(xPuller, "T-PL", "PL")
                        xPuller = Replace(xPuller, "B-PL", "PL")
                    End If
                    '
                    'SLIDER & BODY
                    xSlider = fpObj.ReplaceSliderString(iSlider(i))    '去除字尾SK及-B關鍵字
                    xBody = ""
                    xBodyGr = ""
                    sql = "SELECT BODY, BodyGroup From V_EDXBodyList "
                    sql = sql & "where '" & xSlider & "' like body+'%' "
                    sql = sql & "ORDER BY LEN([BODY]) DESC, BODY "
                    Dim dtBODY1 As DataTable = uDataBase.GetDataTable(sql)
                    If dtBODY1.Rows.Count > 0 Then
                        xBody = dtBODY1.Rows(0)("BODY")
                        xBodyGr = dtBODY1.Rows(0)("BodyGroup")
                        If xBodyGr = "" Then xBodyGr = xSlider
                    End If
                    'MsgBox(CStr(i) & xSlider)
                    '
                    '決定 EDX CHECK 條件(SIZE,FAMILY,BODY) JOO
                    xEDXCheck = ""
                    xEDXVendor = ""
                    sql = "SELECT "
                    sql = sql & "[Size], [Size]+'/'+[SizeA_A]+'/'+[SizeA_C]+'/'+[SizeB_A]+'/'+[SizeB_C]+'/'+[SizeC_A]+'/'+[SizeC_C]+'/', "            '0
                    sql = sql & "[Chain], [Chain]+'/'+[ChainA_A]+'/'+[ChainA_C]+'/'+[ChainB_A]+'/'+[ChainB_C]+'/'+[ChainC_A]+'/'+[ChainC_C]+'/', "     '2
                    sql = sql & "[Class], [Class]+'/'+[ClassA_A]+'/'+[ClassA_C]+'/'+[ClassB_A]+'/'+[ClassB_C]+'/'+[ClassC_A]+'/'+[ClassC_C]+'/', "     '4
                    sql = sql & "[Tape], [Tape]+'/'+[TapeA_A]+'/'+[TapeA_C]+'/'+[TapeB_A]+'/'+[TapeB_C]+'/'+[TapeC_A]+'/'+[TapeC_C]+'/', "            '6
                    sql = sql & "[Slider1], [Slider1]+'/'+[SLIDER1A_A]+'/'+[SLIDER1A_C]+'/'+[SLIDER1B_A]+'/'+[SLIDER1B_C]+'/'+[SLIDER1C_A]+'/'+[SLIDER1C_C]+'/', "       '8
                    sql = sql & "[Finish1], [Finish1]+'/'+[Finish1A_A]+'/'+[Finish1A_C]+'/'+[Finish1B_A]+'/'+[Finish1B_C]+'/'+[Finish1C_A]+'/'+[Finish1C_C]+'/', "       '10
                    sql = sql & "[Slider2], [Slider2]+'/'+[Slider2A_A]+'/'+[Slider2A_C]+'/'+[Slider2B_A]+'/'+[Slider2B_C]+'/'+[Slider2C_A]+'/'+[Slider2C_C]+'/', "       '12
                    sql = sql & "[Finish2], [Finish2]+'/'+[Finish2A_A]+'/'+[Finish2A_C]+'/'+[Finish2B_A]+'/'+[Finish2B_C]+'/'+[Finish2C_A]+'/'+[Finish2C_C]+'/', "       '14
                    sql = sql & "[SRequestList], [SRequestList]+'/'+[SRequestListA_A]+'/'+[SRequestListA_C]+'/'+[SRequestListB_A]+'/'+[SRequestListB_C]+'/'+[SRequestListC_A]+'/'+[SRequestListC_C]+'/', "        '16
                    sql = sql & "[Family], [Family]+'/'+[FamilyA_A]+'/'+[FamilyA_C]+'/'+[FamilyB_A]+'/'+[FamilyB_C]+'/'+[FamilyC_A]+'/'+[FamilyC_C]+'/', "      '18
                    sql = sql & "[ST1], [ST1]+'/'+[ST1A_A]+'/'+[ST1A_C]+'/'+[ST1B_A]+'/'+[ST1B_C]+'/'+[ST1C_A]+'/'+[ST1C_C]+'/', "       '20
                    sql = sql & "[ST2], [ST2]+'/'+[ST2A_A]+'/'+[ST2A_C]+'/'+[ST2B_A]+'/'+[ST2B_C]+'/'+[ST2C_A]+'/'+[ST2C_C]+'/', "       '22
                    sql = sql & "[ST3], [ST3]+'/'+[ST3A_A]+'/'+[ST3A_C]+'/'+[ST3B_A]+'/'+[ST3B_C]+'/'+[ST3C_A]+'/'+[ST3C_C]+'/', "       '24
                    sql = sql & "[ST4], [ST4]+'/'+[ST4A_A]+'/'+[ST4A_C]+'/'+[ST4B_A]+'/'+[ST4B_C]+'/'+[ST4C_A]+'/'+[ST4C_C]+'/', "       '26
                    sql = sql & "[ST5], [ST5]+'/'+[ST5A_A]+'/'+[ST5A_C]+'/'+[ST5B_A]+'/'+[ST5B_C]+'/'+[ST5C_A]+'/'+[ST5C_C]+'/', "       '28
                    sql = sql & "[ST6], [ST6]+'/'+[ST6A_A]+'/'+[ST6A_C]+'/'+[ST6B_A]+'/'+[ST6B_C]+'/'+[ST6C_A]+'/'+[ST6C_C]+'/', "       '30
                    sql = sql & "[ST7], [ST7]+'/'+[ST7A_A]+'/'+[ST7A_C]+'/'+[ST7B_A]+'/'+[ST7B_C]+'/'+[ST7C_A]+'/'+[ST7C_C]+'/', "       '32
                    '
                    sql = sql & "[Active],[Cat],[Rno],[SubNo], "                '34
                    sql = sql & "[Action], [ACTION_A], [ACTION_C], [msg] "      '38
                    '
                    '
                    sql = sql & "From W_BUYERLIMITED "
                    sql = sql & "where Active = 1 "
                    sql = sql & "and CAT = '9.EDX' "
                    '
                    If i = 0 Then
                        sql = sql & "and Slider1 <> '*' "
                    Else
                        sql = sql & "and Slider2 <> '*' "
                    End If
                    If xSubno <> "ZZZZZZ" Then
                        sql = sql & "and Subno <> '" & xSubno & "' "
                    End If
                    '
                    sql = sql & "order by Active,Cat,Rno,SubNo "
                    '
                    Dim dtRule As DataTable = uDataBase.GetDataTable(sql)
                    If dtRule.Rows.Count > 0 Then

                        For x = 0 To dtRule.Rows.Count - 1
                            y = 0
                            If (dtRule.Rows(x)(y) = "*" Or LimitedItem(iNew(y), dtRule.Rows(x)(y + 1)) = True) And _
                               (dtRule.Rows(x)(y + 2) = "*" Or LimitedItem(iNew(y + 1), dtRule.Rows(x)(y + 3)) = True) And _
                               (dtRule.Rows(x)(y + 4) = "*" Or LimitedItem(iNew(y + 2), dtRule.Rows(x)(y + 5)) = True) And _
                               (dtRule.Rows(x)(y + 6) = "*" Or LimitedItem(iNew(y + 3), dtRule.Rows(x)(y + 7)) = True) And _
                               (dtRule.Rows(x)(y + 8) = "*" Or LimitedItem(iNew(y + 4), dtRule.Rows(x)(y + 9)) = True) And _
                               (dtRule.Rows(x)(y + 10) = "*" Or LimitedItem(iNew(y + 5), dtRule.Rows(x)(y + 11)) = True) And _
                               (dtRule.Rows(x)(y + 12) = "*" Or LimitedItem(iNew(y + 6), dtRule.Rows(x)(y + 13)) = True) And _
                               (dtRule.Rows(x)(y + 14) = "*" Or LimitedItem(iNew(y + 7), dtRule.Rows(x)(y + 15)) = True) And _
                               (dtRule.Rows(x)(y + 16) = "*" Or LimitedItem(xBodyGr, dtRule.Rows(x)(y + 17)) = True) And _
                               (dtRule.Rows(x)(y + 18) = "*" Or LimitedItem(iNew(y + 9), dtRule.Rows(x)(y + 19)) = True) And _
                               (dtRule.Rows(x)(y + 20) = "*" Or LimitedItem(iNew(y + 10), dtRule.Rows(x)(y + 21)) = True) And _
                               (dtRule.Rows(x)(y + 22) = "*" Or LimitedItem(iNew(y + 11), dtRule.Rows(x)(y + 23)) = True) And _
                               (dtRule.Rows(x)(y + 24) = "*" Or LimitedItem(iNew(y + 12), dtRule.Rows(x)(y + 25)) = True) And _
                               (dtRule.Rows(x)(y + 26) = "*" Or LimitedItem(iNew(y + 13), dtRule.Rows(x)(y + 27)) = True) And _
                               (dtRule.Rows(x)(y + 28) = "*" Or LimitedItem(iNew(y + 14), dtRule.Rows(x)(y + 29)) = True) And _
                               (dtRule.Rows(x)(y + 30) = "*" Or LimitedItem(iNew(y + 15), dtRule.Rows(x)(y + 31)) = True) And _
                               (dtRule.Rows(x)(y + 32) = "*" Or LimitedItem(iNew(y + 16), dtRule.Rows(x)(y + 33)) = True) Then
                                '
                                xEDXCheck = dtRule.Rows(x)(40)
                                xEDXVendor = dtRule.Rows(x)(41)
                                xSubno = dtRule.Rows(x)(37)
                                'DOrderDate.Text = DOrderDate.Text & "/" & xSubno
                                Exit For
                                '
                            End If
                            '
                        Next
                    End If
                    '
                    'MsgBox(Replace(xSlider, xBody, "") & "--" & DOrderDate.Text)
                    '
                    'SEARCH EDX
                    sql = "SELECT Top 1 Size, Family, Body, Puller, Color, Finish, Cat, Supplier "
                    sql = sql & "From M_EDX "
                    '
                    '一般SLIDER CODE (SPD,QC)
                    Select Case xEDXCheck
                        Case "SIZE"
                            'SIZE + PULLER + COLOR
                            sql = sql & "Where ( "
                            sql = sql & "      Size + Puller + Color = '" & DSize.Text & Replace(xSlider, xBody, "") & "' "
                            sql = sql & ") "
                        Case "FAMILY"
                            'FAMILY + PULLER + COLOR
                            sql = sql & "Where ( "
                            sql = sql & "      Family + Puller + Color = '" & DFamily.Text & Replace(xSlider, xBody, "") & "' "
                            sql = sql & ") "
                        Case "BODY"
                            'BODY + PULLER + COLOR
                            sql = sql & "Where ( "
                            sql = sql & "      Body + Puller + Color = '" & xSlider & "' "
                            sql = sql & "   OR (Body LIKE '" & xBodyGr & "%' "
                            sql = sql & "        AND Puller + Color = '" & Replace(xSlider, xBody, "") & "' "
                            sql = sql & "      ) "
                            sql = sql & ") "
                        Case "SIZE/FAMILY"
                            'SIZE + FAMILY + PULLER + COLOR
                            sql = sql & "Where ( "
                            sql = sql & "      Size + Family + Puller + Color = '" & DSize.Text & DFamily.Text & Replace(xSlider, xBody, "") & "' "
                            sql = sql & ") "
                        Case "SIZE/FAMILY/BODY"
                            'SIZE + FAMILY + BODY + PULLER + COLOR
                            sql = sql & "Where ( "
                            sql = sql & "      Size + Family + Body + Puller + Color = '" & DSize.Text & DFamily.Text & xSlider & "' "
                            sql = sql & "   OR (Size + Family + Body LIKE '" & DSize.Text & DFamily.Text & xBodyGr & "%' "
                            sql = sql & "        AND Puller + Color = '" & Replace(xSlider, xBody, "") & "' "
                            sql = sql & "      ) "
                            sql = sql & ") "
                        Case "VENDOR"
                            'VENDOR
                            sql = sql & "Where ( "
                            sql = sql & "       Puller + Color = '" & Replace(xSlider, xBody, "") & "' "
                            sql = sql & "   AND Ltrim(Rtrim(Supplier)) = '" & Trim(xEDXVendor) & "' "
                            sql = sql & ") "
                        Case "CHG-BODY"
                            'BODY=DSBL-UU --> DSB-LUU
                            sql = sql & "Where ( "
                            sql = sql & "       Puller + Color = '" & Replace(xSlider, xEDXVendor, "") & "' "
                            sql = sql & ") "
                        Case Else
                            'PULLER + COLOR
                            sql = sql & "Where ( "
                            sql = sql & "       Common = '1' "
                            sql = sql & "       and Puller + Color = '" & Replace(xSlider, xBody, "") & "' "
                            sql = sql & ") "
                    End Select
                    '
                    '一般SLIDER CODE (PRICE)
                    sql = sql & "Or ( "
                    sql = sql & "     Puller = '" & iSlider(i) & "' "
                    sql = sql & ") "
                    'If xEDXCheck = "" Then
                    'End If
                    '
                    'SLIDER CODE縮寫(PL-)
                    If xPuller <> "" And _
                      (Mid(xSlider, 1, 4) = "DA7B" Or Mid(xSlider, 1, 5) = "DAG7B" Or Mid(xSlider, 1, 4) = "DS15" Or Mid(xSlider, 1, 5) = "DAG24") Then
                        'QC EDX
                        'IRW: DA7BA PL-BS054
                        'EDX : PL-BS054 A
                        sql = sql & "Or ( "
                        sql = sql & "      Color + Puller = '" & Replace(xSlider + xPuller, xBody, "") & "' "
                        sql = sql & "   OR Color + Puller = '" & Replace(xSlider + Replace(xPuller, "PL-", ""), xBody, "") & "' "
                        sql = sql & ") "
                        '(特殊) PULLER CODE分離 (PL-LLM)
                        'IRW: DA7B74A  PL-LLM
                        'EDX : PL-LLM74 A
                        sql = sql & "Or ( "
                        sql = sql & "      Puller + Color = '" & xPuller + Replace(xSlider, xBody, "") & "' "
                        sql = sql & "   OR Puller + Color = '" & Replace(xPuller, "PL-", "") + Replace(xSlider, xBody, "") & "' "
                        sql = sql & ") "
                        'PRICE
                        sql = sql & "Or ( "
                        sql = sql & "      Puller = '" & iSlider(i) + xPuller & "' "
                        sql = sql & ") "
                    End If
                    '
                    sql = sql & "Order By SeqNo, len(Puller) desc, Color Desc, CreateTime Desc "
                    '
                    'MsgBox(sql)
                    '
                    Dim dtEDXDB1 As DataTable = uDataBase.GetDataTable(sql)
                    If dtEDXDB1.Rows.Count > 0 Then
                        xSource = "EDX(" & dtEDXDB1.Rows(0)("CAT") & ")"
                        '
                        DSPDNOUrl.Text = DSPDNOUrl.Text & Chr(13) & "(E)SLD-" & CStr(i + 1) & "[" & dtEDXDB1.Rows(0)("CAT") & "-" & dtEDXDB1.Rows(0)("PULLER") & "-" & dtEDXDB1.Rows(0)("COLOR") & "-" & xBody & "]"
                        If dtEDXDB1.Rows(0)("Supplier") <> "" Then
                            DSPDNOUrl.Text = DSPDNOUrl.Text & "e_Supplier:|" & dtEDXDB1.Rows(0)("Supplier") & "|"
                        End If
                    End If
                    '
                    If xEDXCheck <> "" And xSource = "WINGS" Then
                        xSource = ""
                    End If
                    '
                    'MsgBox("[" & xSource & "]")
                    '
                    '3Y SALES & EDX DB = 無 
                    '決定SLIDER CODE (10碼縮寫對應) 
                    If (InStr(xSource, "EDX") <= 0 Or InStr(xSource, "EDX(PRICE)") > 0 Or xSource = "") And DBuyerCode.Text <> "" Then
                        'DSBLM59AAN
                        ' 分離 PULLER / COLOR --> DSBLM59 / AAN 
                        str = fpObj.ReplaceSliderString(iSlider(i))    '去除字尾SK及-B關鍵字
                        xSlider = ""
                        xColor = ""
                        j = Len(str)
                        Do Until j = 0
                            If Mid(str, j, 1) >= "0" And Mid(str, j, 1) <= "9" Then
                                Exit Do
                            Else
                                xSlider = Mid(str, 1, j - 1)
                            End If
                            j = j - 1
                        Loop
                        If xSlider = "" Then xSlider = Trim(str)
                        xColor = Trim(Replace(str, xSlider, ""))
                        '
                        'MsgBox("[" & xSlider & "]" & xColor & "--" & xBody)
                        '
                        If xSlider <> xBody Then
                            sql = "SELECT Buyer, Puller, Short "
                            sql = sql & "From V_PullerShort "
                            sql = sql & "Where ( Buyer like '%" & DBuyerCode.Text & "/%' OR Buyer = 'ALL' ) "
                            sql = sql & "And ( "
                            sql = sql & "Short like '%" & "ZZZZZZZ" & "%' "
                            '
                            '2025/1/10 & 2025/3/15 MODIFY-START
                            'sql = sql & "Or Short like '%" & Replace(xSlider, xBody, "") & "/%' "
                            If xEDXCheck = "CHG-BODY" Then
                                sql = sql & "Or Short like '%/" & Replace(xSlider, xEDXVendor, "") & "/%' "
                            Else
                                sql = sql & "Or Short like '%/" & Replace(xSlider, xBody, "") & "/%' "
                            End If
                            'END
                            '
                            sql = sql & ") "
                            sql = sql & "Order By Buyer, Puller, Short "
                            '
                            Dim dtPuller1 As DataTable = uDataBase.GetDataTable(sql)
                            If dtPuller1.Rows.Count > 0 Then

                                DSPDNOUrl.Text = DSPDNOUrl.Text & Chr(13) & "(E-M)(B)SLD-" & CStr(i + 1) & "[" & dtPuller1.Rows(0)("Buyer") & "]" & Chr(13) & "[" & dtPuller1.Rows(0)("Short") & "]"

                                str = dtPuller1.Rows(0)("Short")
                                iMember = str.ToString.Split("/")

                                For j = 0 To UBound(iMember) - 1
                                    '
                                    'ADD-START 250315
                                    If iMember(j) = "" Then iMember(j) = "ZZZZZZ"
                                    'ADD-END
                                    '
                                    sql = "SELECT Top 1 Size, Family, Body, Puller, Color, Finish, Cat, Supplier "
                                    sql = sql & "From M_EDX "
                                    '
                                    '一般SLIDER CODE (SPD,QC)
                                    Select Case xEDXCheck
                                        Case "SIZE"
                                            'SIZE + PULLER + COLOR
                                            sql = sql & "Where ( "
                                            sql = sql & "      Size + Puller + Color = '" & DSize.Text & iMember(j) + xColor & "' "
                                            sql = sql & ") "
                                        Case "FAMILY"
                                            'FAMILY + PULLER + COLOR
                                            sql = sql & "Where ( "
                                            sql = sql & "      Family + Puller + Color = '" & DFamily.Text & iMember(j) + xColor & "' "
                                            sql = sql & ") "
                                        Case "BODY"
                                            'BODY + PULLER + COLOR
                                            sql = sql & "Where ( "
                                            sql = sql & "      Body + Puller + Color = '" & xBody & iMember(j) + xColor & "' "
                                            sql = sql & "   OR Body + Puller + Color = '" & xBodyGr & iMember(j) + xColor & "' "
                                            sql = sql & ") "
                                        Case "SIZE/FAMILY"
                                            'SIZE + FAMILY + PULLER + COLOR
                                            sql = sql & "Where ( "
                                            sql = sql & "      Size + Family + Puller + Color = '" & DSize.Text & DFamily.Text & iMember(j) + xColor & "' "
                                            sql = sql & ") "
                                        Case "SIZE/FAMILY/BODY"
                                            'SIZE + FAMILY + BODY + PULLER + COLOR
                                            sql = sql & "Where ( "
                                            sql = sql & "      Size + Family + Body + Puller + Color = '" & DSize.Text & DFamily.Text & xBody & iMember(j) + xColor & "' "
                                            sql = sql & ") "
                                        Case "VENDOR"
                                            'VENDOR
                                            sql = sql & "Where ( "
                                            sql = sql & "       Puller + Color = '" & iMember(j) + xColor & "' "
                                            sql = sql & "   AND Ltrim(Rtrim(Supplier)) = '" & Trim(xEDXVendor) & "' "
                                            sql = sql & ") "
                                        Case "CHG-BODY"
                                            'BODY=DSBL-UU --> DSB-LUU
                                            sql = sql & "Where ( "
                                            sql = sql & "       Puller + Color = '" & iMember(j) + xColor & "' "
                                            sql = sql & ") "
                                        Case Else
                                            'PULLER + COLOR
                                            sql = sql & "Where ( "
                                            sql = sql & "      Common = '1' "
                                            sql = sql & "  and Puller + Color = '" & iMember(j) + xColor & "' "
                                            sql = sql & ") "
                                    End Select
                                    '
                                    sql = sql & "Order By SeqNo, len(Puller) desc, Color Desc, CreateTime Desc "
                                    Dim dtMember1 As DataTable = uDataBase.GetDataTable(sql)
                                    If dtMember1.Rows.Count > 0 Then
                                        If xSource = "" Then
                                            xSource = xSource & "-" & "SHORT"
                                        Else
                                            xSource = "SHORT"
                                        End If
                                        DSPDNOUrl.Text = DSPDNOUrl.Text & Chr(13) & "(E-M)SLD-" & CStr(i + 1) & "[" & dtMember1.Rows(0)("CAT") & "-" & dtMember1.Rows(0)("PULLER") & "-" & dtMember1.Rows(0)("COLOR") & "-" & xBody & "]"
                                        If dtMember1.Rows(0)("Supplier") <> "" Then
                                            DSPDNOUrl.Text = DSPDNOUrl.Text & "e_Supplier:|" & dtMember1.Rows(0)("Supplier") & "|"
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
                '
                If xSource = "" Then
                    LEDXLink1.NavigateUrl = "http://10.245.1.6/ISOSQC/QCListinqCommission.aspx" & _
                                            "?pIRW=1" & _
                                            "&pSize=" & DSize.Text & _
                                            "&pSlider=" & xSlider & _
                                            "&pPuller=" & xPuller
                    LEDXLink1.Visible = True
                    LEDXRegister.NavigateUrl = "https://forms.office.com/r/MDyLjRE2V2"
                    LEDXRegister.Visible = True

                    DEDXComment1.Visible = True
                    DEDXComment2.Visible = True
                    '
                    DSPDNOUrl.Text = DSPDNOUrl.Text & Chr(13) & "SLD-" & CStr(i + 1) & "[" & iSlider(i) & ":EDX[NOT FOUND:EDX DB]"
                    DSPDNOUrl.Text = DSPDNOUrl.Text & Chr(13) & "Body=[" & xBody & "] / Puller=[" & Replace(iSlider(i), xBody, "") & "]"
                    uError = True
                    Exit For
                End If
            Next
        End If
        '
        '未上線
        '--------------------------------------------------------------------------------------------
        If uError Then
            'ADD EDX DB
            sql = "INSERT INTO M_SPDIRWEDX ( " & _
                  "[Sts], [Cat], [RecDate], [SchComDate], [AcceptedNo], " & _
                  "[RDNo], [Size], [Family], [Body], [Puller], [Color], [Finish], [EdxMacNo], [Partner], [Area], " & _
                  "[Result1], [ResultDate1], [TestWait], [Result2], [ResultDate2], " & _
                  "[LastResult], [Remark], [SpecGroup], [FormNo], [FormSno], " & _
                  "[YOBI1], [YOBI2], [YOBI3], [YOBI9], " & _
                  "[CreateUser], [CreateTime], [ModifyUser], [ModifyTime]) " & _
            "VALUES("
            '
            'MOD-END
            sql &= "0 ,"
            sql &= "'HAND' ,"
            sql &= "'" & NowDateTime & "' ,"
            sql &= "'" & NowDateTime & "' ,"
            sql &= "'" & Now.ToString("yyyyMMddHHmmss") & "' ,"
            '
            sql &= "'A-LEARNING' ,"
            sql &= "'" & DSize.Text & "' ,"
            sql &= "'" & DFamily.Text & "' ,"
            sql &= "'" & "" & "' ,"
            If InStr(DSPDNOUrl.Text, "SLD-1") > 0 Then
                sql &= "'" & DSlider1.Text & "' ,"
                sql &= "'" & "" & "' ,"
                sql &= "'" & DFinish1.Text & "' ,"
            Else
                sql &= "'" & DSlider2.Text & "' ,"
                sql &= "'" & "" & "' ,"
                sql &= "'" & DFinish2.Text & "' ,"
            End If
            sql &= "'' ,"
            sql &= "'' ,"
            sql &= "'' ,"
            '
            sql &= "'OK' ,"
            sql &= "'" & NowDateTime & "' ,"
            sql &= "'' ,"
            sql &= "'' ,"
            sql &= "NULL ,"
            '
            sql &= "'OK' ,"
            sql &= "'自動學習' ,"
            sql &= "'' ,"
            sql &= "'' ,"
            sql &= "0 ,"
            '
            sql &= "'' ,"
            sql &= "'' ,"
            sql &= "'' ,"
            sql &= "'" & DSPDNOUrl.Text & "' ,"
            '
            sql &= "'" & Request.QueryString("pUserID") & "' ,"
            sql &= "'" & NowDateTime & "' ,"
            sql &= "'" & "" & "' ,"
            sql &= "NULL ) "
            '
            uDataBase.ExecuteNonQuery(sql)
        End If
        '
        'MsgBox("EDXNotFound-OUT")
        Return uError
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ReplaceFamily)
    '**     
    '**
    '*****************************************************************
    Public Function ReplaceFamily(ByVal pSize As String, ByVal pFamily As String) As String
        'ReplaceFamily(DSize.Text & "-" & DFamily.Text)
        Dim NewFamily As String = ""
        Dim SizeFamily As String = pSize & "-" & pFamily
        '
        If InStr(SizeFamily, "05-CI") > 0 Then NewFamily = Replace(SizeFamily, "05-CI", "CN")
        If InStr(SizeFamily, "05-CZ") > 0 Then NewFamily = Replace(SizeFamily, "05-CZ", "CN")
        If InStr(SizeFamily, "05-CS") > 0 Then NewFamily = Replace(SizeFamily, "05-CS", "CN")
        If InStr(SizeFamily, "05-Y") > 0 Then NewFamily = Replace(SizeFamily, "05-Y", "M")
        '
        If InStr(SizeFamily, "03-CZ") > 0 Then NewFamily = Replace(SizeFamily, "03-CZ", "C")
        If InStr(SizeFamily, "03-CS") > 0 Then NewFamily = Replace(SizeFamily, "03-CS", "C")
        '
        If NewFamily = "" Then NewFamily = pFamily
        '
        Return NewFamily
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(LimitedItemError)
    '**     檢查BUYER或ITEM限定 pCat='BUYER' OR 'ITEM'
    '**
    '*****************************************************************
    Public Function LimitedItemError(ByVal pCat As String) As Boolean
        Dim i, j As Integer
        Dim x, y As Integer
        Dim msgNo As Integer
        Dim sql, msg, xItemName, xSpc, sqlISIP As String
        Dim str, str1, str2, ISIPSlider1, ISIPSlider2 As String
        Dim uError As Boolean = False
        Dim uRun As Boolean = False
        Dim iNew As String()
        sqlISIP = ""
        msgNo = 1
        sWarnMsg = ""
        '
        xItemName = DItemName1.Text & " " & DItemName2.Text & " " & DItemName3.Text
        '
        xSpc = GetFullName(DSRequest1.Text) & "!" & _
               GetFullName(DSRequest2.Text) & "!" & _
               GetFullName(DSRequest3.Text) & "!" & _
               GetFullName(DSRequest4.Text) & "!" & _
               GetFullName(DSRequest5.Text) & "!" & _
               GetFullName(DSRequest6.Text) & "!"
        '
        str = DSize.Text & "!" & _
                DChain.Text & "!" & _
                DClass.Text & "!" & _
                DTape.Text & "!" & _
                DSlider1.Text & "!" & _
                DFinish1.Text & "!" & _
                DSlider2.Text & "!" & _
                DFinish2.Text & "!" & _
                DSRequest1.Text & "!" & _
                DFamily.Text & "!" & _
                DST1.Text & "!" & _
                DST2.Text & "!" & _
                DST3.Text & "!" & _
                DST4.Text & "!" & _
                DST5.Text & "!" & _
                DST6.Text & "!" & _
                DST7.Text & "!"
        '
        iNew = str.ToString.Split("!")
        '
        ' SLIDER
        '去除字尾SK及-B關鍵字
        '
        'MODIFY-START 241014
        '  限定:[非ISIP]=不去除字尾SK及-B關鍵字
        '  限定:[ISIP]=去除字尾SK及-B關鍵字
        '
        'iNew(4) = fpObj.ReplaceSliderString(iNew(4))
        'iNew(6) = fpObj.ReplaceSliderString(iNew(6))
        ISIPSlider1 = fpObj.ReplaceSliderString(iNew(4))
        ISIPSlider2 = fpObj.ReplaceSliderString(iNew(6))
        '
        'MODIFY-END
        '
        If Not uError Then
            sql = "SELECT "
            sql = sql & "[Size], [Size]+'/'+[SizeA_A]+'/'+[SizeA_C]+'/'+[SizeB_A]+'/'+[SizeB_C]+'/'+[SizeC_A]+'/'+[SizeC_C]+'/', "            '0
            sql = sql & "[Chain], [Chain]+'/'+[ChainA_A]+'/'+[ChainA_C]+'/'+[ChainB_A]+'/'+[ChainB_C]+'/'+[ChainC_A]+'/'+[ChainC_C]+'/', "     '2
            sql = sql & "[Class], [Class]+'/'+[ClassA_A]+'/'+[ClassA_C]+'/'+[ClassB_A]+'/'+[ClassB_C]+'/'+[ClassC_A]+'/'+[ClassC_C]+'/', "     '4
            sql = sql & "[Tape], [Tape]+'/'+[TapeA_A]+'/'+[TapeA_C]+'/'+[TapeB_A]+'/'+[TapeB_C]+'/'+[TapeC_A]+'/'+[TapeC_C]+'/', "            '6
            sql = sql & "[Slider1], [Slider1]+'/'+[SLIDER1A_A]+'/'+[SLIDER1A_C]+'/'+[SLIDER1B_A]+'/'+[SLIDER1B_C]+'/'+[SLIDER1C_A]+'/'+[SLIDER1C_C]+'/', "       '8
            sql = sql & "[Finish1], [Finish1]+'/'+[Finish1A_A]+'/'+[Finish1A_C]+'/'+[Finish1B_A]+'/'+[Finish1B_C]+'/'+[Finish1C_A]+'/'+[Finish1C_C]+'/', "       '10
            sql = sql & "[Slider2], [Slider2]+'/'+[Slider2A_A]+'/'+[Slider2A_C]+'/'+[Slider2B_A]+'/'+[Slider2B_C]+'/'+[Slider2C_A]+'/'+[Slider2C_C]+'/', "       '12
            sql = sql & "[Finish2], [Finish2]+'/'+[Finish2A_A]+'/'+[Finish2A_C]+'/'+[Finish2B_A]+'/'+[Finish2B_C]+'/'+[Finish2C_A]+'/'+[Finish2C_C]+'/', "       '14
            sql = sql & "[SRequestList], [SRequestList]+'/'+[SRequestListA_A]+'/'+[SRequestListA_C]+'/'+[SRequestListB_A]+'/'+[SRequestListB_C]+'/'+[SRequestListC_A]+'/'+[SRequestListC_C]+'/', "        '16
            sql = sql & "[Family], [Family]+'/'+[FamilyA_A]+'/'+[FamilyA_C]+'/'+[FamilyB_A]+'/'+[FamilyB_C]+'/'+[FamilyC_A]+'/'+[FamilyC_C]+'/', "      '18
            sql = sql & "[ST1], [ST1]+'/'+[ST1A_A]+'/'+[ST1A_C]+'/'+[ST1B_A]+'/'+[ST1B_C]+'/'+[ST1C_A]+'/'+[ST1C_C]+'/', "       '20
            sql = sql & "[ST2], [ST2]+'/'+[ST2A_A]+'/'+[ST2A_C]+'/'+[ST2B_A]+'/'+[ST2B_C]+'/'+[ST2C_A]+'/'+[ST2C_C]+'/', "       '22
            sql = sql & "[ST3], [ST3]+'/'+[ST3A_A]+'/'+[ST3A_C]+'/'+[ST3B_A]+'/'+[ST3B_C]+'/'+[ST3C_A]+'/'+[ST3C_C]+'/', "       '24
            sql = sql & "[ST4], [ST4]+'/'+[ST4A_A]+'/'+[ST4A_C]+'/'+[ST4B_A]+'/'+[ST4B_C]+'/'+[ST4C_A]+'/'+[ST4C_C]+'/', "       '26
            sql = sql & "[ST5], [ST5]+'/'+[ST5A_A]+'/'+[ST5A_C]+'/'+[ST5B_A]+'/'+[ST5B_C]+'/'+[ST5C_A]+'/'+[ST5C_C]+'/', "       '28
            sql = sql & "[ST6], [ST6]+'/'+[ST6A_A]+'/'+[ST6A_C]+'/'+[ST6B_A]+'/'+[ST6B_C]+'/'+[ST6C_A]+'/'+[ST6C_C]+'/', "       '30
            sql = sql & "[ST7], [ST7]+'/'+[ST7A_A]+'/'+[ST7A_C]+'/'+[ST7B_A]+'/'+[ST7B_C]+'/'+[ST7C_A]+'/'+[ST7C_C]+'/', "       '32
            '
            sql = sql & "[Active],[Cat],[Rno],[SubNo], "                '34
            sql = sql & "[Action], [ACTION_A], [ACTION_C], [msg] "      '38
            '
            sql = sql & "From W_BUYERLIMITED "
            sql = sql & "where Active = 1 "
            '
            sqlISIP = sql & "and Cat = '4.ITEM' and SUBSTRING([Rno],1,1)='M' "        '只抓ISIP轉入的ITEM限定資料
            '
            'PERFORMACE-START
            '[ST1] ~ [ST7]
            'sql = sql & "AND ( [ST1] = '*' OR [ST1] = '" & iNew(10) & "' ) "
            'sql = sql & "AND ( [ST2] = '*' OR [ST2] = '" & iNew(11) & "' ) "
            'sql = sql & "AND ( [ST3] = '*' OR [ST3] = '" & iNew(12) & "' ) "
            'sql = sql & "AND ( [ST4] = '*' OR [ST4] = '" & iNew(13) & "' ) "
            'sql = sql & "AND ( [ST5] = '*' OR [ST5] = '" & iNew(14) & "' ) "
            'sql = sql & "AND ( [ST6] = '*' OR [ST6] = '" & iNew(15) & "' ) "
            'sql = sql & "AND ( [ST7] = '*' OR [ST7] = '" & iNew(16) & "' ) "
            '
            '[CUSTOMER] ~ [BUYER]
            If pCat = "ITEM" Then
                sql = sql & "and Cat = '4.ITEM' and SUBSTRING([Rno],1,1)!='M' "         '排除ISIP轉入的資料  
            ElseIf pCat = "WARNING" Then
                sql = sql & "and CAT ='5.WARNING' "
                sqlISIP = ""
            Else
                sql = sql & "and CAT NOT IN ('4.ITEM','5.WARNING','9.EDX') "
                sqlISIP = ""
            End If
            'PERFORMACE-END
            '
            sql = sql & "order by Active,Cat,Rno,SubNo "

            '
            Dim dtRule As DataTable = uDataBase.GetDataTable(sql)
            '
            '取得ISIP轉入ITEM限定的規則
            If sqlISIP <> "" Then
                'Item限定中針對ISIP轉入特定SLIDER(PULLER+COLOR)筆數多的資料只取得Slider相符的資料
                '
                'MODIFY-START 241014
                '  限定:[非ISIP]=不去除字尾SK及-B關鍵字
                '  限定:[ISIP]=去除字尾SK及-B關鍵字
                '
                'sqlISIP = sqlISIP & " and ('" & iNew(4) & "' like Slider1 "
                'If iNew(6) <> "" Then
                '    sqlISIP = sqlISIP & " or '" & iNew(6) & "' like Slider2 "
                'End If
                sqlISIP = sqlISIP & " and ('" & ISIPSlider1 & "' like Slider1 "
                If ISIPSlider2 <> "" Then
                    sqlISIP = sqlISIP & " or '" & ISIPSlider2 & "' like Slider2 "
                End If
                '
                'MODIFY-END
                '
                sqlISIP = sqlISIP & ") order by Active,Cat,Rno,SubNo "
                Dim dtRuleISIP As DataTable = uDataBase.GetDataTable(sqlISIP)
                dtRule.Merge(dtRuleISIP)    '合併限定規則資料
            End If
            '
            If dtRule.Rows.Count > 0 Then

                DLimitedItem.Text = "限定檢測-START"

                For i = 0 To dtRule.Rows.Count - 1
                    uError = False
                    'MsgBox(pCat & "--" & CStr(dtRule.Rows.Count) & dtRule.Rows(i)(37))
                    ' -----------------------------------------------------------------
                    ' 是否限定ITEM
                    ' -----------------------------------------------------------------
                    uRun = False
                    j = 0
                    If (dtRule.Rows(i)(j) = "*" Or LimitedItem(iNew(j), dtRule.Rows(i)(j + 1)) = True) And _
                       (dtRule.Rows(i)(j + 2) = "*" Or LimitedItem(iNew(j + 1), dtRule.Rows(i)(j + 3)) = True) And _
                       (dtRule.Rows(i)(j + 4) = "*" Or LimitedItem(iNew(j + 2), dtRule.Rows(i)(j + 5)) = True) And _
                       (dtRule.Rows(i)(j + 6) = "*" Or LimitedItem(iNew(j + 3), dtRule.Rows(i)(j + 7)) = True) And _
                       ((dtRule.Rows(i)(j + 8) = "*" Or LimitedItem(iNew(j + 4), dtRule.Rows(i)(j + 9)) = True) Or _
                         (dtRule.Rows(i)(j + 8) = "*" Or LimitedItem(ISIPSlider1, dtRule.Rows(i)(j + 9)) = True)) And _
                       (dtRule.Rows(i)(j + 10) = "*" Or LimitedItem(iNew(j + 5), dtRule.Rows(i)(j + 11)) = True) And _
                       ((dtRule.Rows(i)(j + 12) = "*" Or LimitedItem(iNew(j + 6), dtRule.Rows(i)(j + 13)) = True) Or _
                         (dtRule.Rows(i)(j + 12) = "*" Or LimitedItem(ISIPSlider2, dtRule.Rows(i)(j + 13)) = True)) And _
                       (dtRule.Rows(i)(j + 14) = "*" Or LimitedItem(iNew(j + 7), dtRule.Rows(i)(j + 15)) = True) And _
                       (dtRule.Rows(i)(j + 16) = "*" Or LimitedItem(xSpc, dtRule.Rows(i)(j + 17)) = True) And _
                       (dtRule.Rows(i)(j + 18) = "*" Or LimitedItem(iNew(j + 9), dtRule.Rows(i)(j + 19)) = True) And _
                       (dtRule.Rows(i)(j + 20) = "*" Or LimitedItem(iNew(j + 10), dtRule.Rows(i)(j + 21)) = True) And _
                       (dtRule.Rows(i)(j + 22) = "*" Or LimitedItem(iNew(j + 11), dtRule.Rows(i)(j + 23)) = True) And _
                       (dtRule.Rows(i)(j + 24) = "*" Or LimitedItem(iNew(j + 12), dtRule.Rows(i)(j + 25)) = True) And _
                       (dtRule.Rows(i)(j + 26) = "*" Or LimitedItem(iNew(j + 13), dtRule.Rows(i)(j + 27)) = True) And _
                       (dtRule.Rows(i)(j + 28) = "*" Or LimitedItem(iNew(j + 14), dtRule.Rows(i)(j + 29)) = True) And _
                       (dtRule.Rows(i)(j + 30) = "*" Or LimitedItem(iNew(j + 15), dtRule.Rows(i)(j + 31)) = True) And _
                       (dtRule.Rows(i)(j + 32) = "*" Or LimitedItem(iNew(j + 16), dtRule.Rows(i)(j + 33)) = True) Then
                        uRun = True
                    End If
                    '
                    ' -----------------------------------------------------------------
                    ' 檢查 限定BUYER, CUSTOMER, ITEM
                    ' -----------------------------------------------------------------
                    If uRun Then

                        Select Case dtRule.Rows(i)(35)
                            Case "1.BUYER"
                                If DBuyerCode.Text <> "" Then
                                    str = DBuyerCode.Text
                                    Select Case dtRule.Rows(i)(39)
                                        Case "!%"
                                            If InStr(dtRule.Rows(i)(40), str) > 0 Then uError = True
                                        Case "%"
                                            If InStr(dtRule.Rows(i)(40), str) <= 0 Then uError = True
                                        Case "!="
                                            If dtRule.Rows(i)(40) = str Then uError = True
                                        Case "="
                                            If dtRule.Rows(i)(40) <> str Then uError = True
                                        Case Else
                                    End Select
                                End If
                            Case "2.CUSTOMER"
                                If DCustomerCode.Text <> "" Then
                                    str = DCustomerCode.Text
                                    Select Case dtRule.Rows(i)(39)
                                        Case "!%"
                                            If InStr(dtRule.Rows(i)(40), str) > 0 Then uError = True
                                        Case "%"
                                            If InStr(dtRule.Rows(i)(40), str) <= 0 Then uError = True
                                        Case "!="
                                            If dtRule.Rows(i)(40) = str Then uError = True
                                        Case "="
                                            If dtRule.Rows(i)(40) <> str Then uError = True
                                        Case Else
                                    End Select
                                End If
                            Case "3.BUY-CUST"
                                'TW5759-G2600/000151-ALL/ALL-000036/
                                '!% 任一有含ERROR
                                '%  全部不含ERROR
                                '!= 任一等於 ERROR
                                '=  全部不等於 ERRPR
                                If DBuyerCode.Text <> "" And DCustomerCode.Text <> "" Then
                                    str = DBuyerCode.Text & "-" & DCustomerCode.Text
                                    str1 = DBuyerCode.Text & "-" & "ALL"
                                    str2 = "ALL" & "-" & DCustomerCode.Text

                                    Select Case dtRule.Rows(i)(39)
                                        Case "!%"
                                            If InStr(dtRule.Rows(i)(40), str) > 0 Or InStr(dtRule.Rows(i)(40), str1) > 0 Or InStr(dtRule.Rows(i)(40), str2) > 0 Then uError = True
                                        Case "%"
                                            If InStr(dtRule.Rows(i)(40), str) <= 0 And InStr(dtRule.Rows(i)(40), str1) <= 0 And InStr(dtRule.Rows(i)(40), str2) <= 0 Then uError = True
                                        Case "!="
                                            If dtRule.Rows(i)(40) = str Or dtRule.Rows(i)(40) = str1 Or dtRule.Rows(i)(40) = str2 Then uError = True
                                        Case "="
                                            If dtRule.Rows(i)(40) <> str And dtRule.Rows(i)(40) <> str1 And dtRule.Rows(i)(40) <> str2 Then uError = True
                                        Case Else
                                    End Select
                                End If
                            Case "5.WARNING"
                                If dtRule.Rows(i)(39) <> "" And dtRule.Rows(i)(40) <> "" And InStr(dtRule.Rows(i)(40).ToString(), "[") > 0 _
                                    And InStr(dtRule.Rows(i)(40).ToString(), "]") > 0 Then
                                    Dim strCmp() = dtRule.Rows(i)(40).ToString().Trim().Substring(1, Len(dtRule.Rows(i)(40).ToString().Trim()) - 2).Split("|")
                                    Dim strFD1 As String
                                    Dim strFD2 As String
                                    '取得比對欄位的值
                                    strFD1 = fpObj.GetCompareVal(str, strCmp(0))
                                    strFD2 = fpObj.GetCompareVal(str, strCmp(1))
                                    Select Case dtRule.Rows(i)(39)
                                        Case "!="
                                            If strFD1 <> strFD2 Then uError = True
                                        Case "="
                                            If strFD1 = strFD2 Then uError = True
                                        Case Else
                                    End Select
                                Else
                                    uError = True
                                End If
                            Case "4.ITEM"
                                uError = True
                            Case Else
                        End Select
                        '
                        DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & dtRule.Rows(i)(35) & "-" & dtRule.Rows(i)(36) & "-" & dtRule.Rows(i)(37) & "]"
                        If dtRule.Rows(i)(35) <> "4.ITEM" And dtRule.Rows(i)(35) <> "5.WARNING" Then
                            DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & dtRule.Rows(i)(40) & "-" & str & "]"
                        End If
                        '
                        If uError = True Then
                            '
                            DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "限定中" & "]"
                            '
                            ' --------------------------------------------------------------------------------------
                            ' 例外檢測
                            ' --------------------------------------------------------------------------------------
                            '
                            sql = "SELECT "
                            sql = sql & "[Size], [Size]+'/'+[SizeA_A]+'/'+[SizeA_C]+'/'+[SizeB_A]+'/'+[SizeB_C]+'/'+[SizeC_A]+'/'+[SizeC_C]+'/', "            '0
                            sql = sql & "[Chain], [Chain]+'/'+[ChainA_A]+'/'+[ChainA_C]+'/'+[ChainB_A]+'/'+[ChainB_C]+'/'+[ChainC_A]+'/'+[ChainC_C]+'/', "     '2
                            sql = sql & "[Class], [Class]+'/'+[ClassA_A]+'/'+[ClassA_C]+'/'+[ClassB_A]+'/'+[ClassB_C]+'/'+[ClassC_A]+'/'+[ClassC_C]+'/', "     '4
                            sql = sql & "[Tape], [Tape]+'/'+[TapeA_A]+'/'+[TapeA_C]+'/'+[TapeB_A]+'/'+[TapeB_C]+'/'+[TapeC_A]+'/'+[TapeC_C]+'/', "            '6
                            sql = sql & "[Slider1], [Slider1]+'/'+[SLIDER1A_A]+'/'+[SLIDER1A_C]+'/'+[SLIDER1B_A]+'/'+[SLIDER1B_C]+'/'+[SLIDER1C_A]+'/'+[SLIDER1C_C]+'/', "       '8
                            sql = sql & "[Finish1], [Finish1]+'/'+[Finish1A_A]+'/'+[Finish1A_C]+'/'+[Finish1B_A]+'/'+[Finish1B_C]+'/'+[Finish1C_A]+'/'+[Finish1C_C]+'/', "       '10
                            sql = sql & "[Slider2], [Slider2]+'/'+[Slider2A_A]+'/'+[Slider2A_C]+'/'+[Slider2B_A]+'/'+[Slider2B_C]+'/'+[Slider2C_A]+'/'+[Slider2C_C]+'/', "       '12
                            sql = sql & "[Finish2], [Finish2]+'/'+[Finish2A_A]+'/'+[Finish2A_C]+'/'+[Finish2B_A]+'/'+[Finish2B_C]+'/'+[Finish2C_A]+'/'+[Finish2C_C]+'/', "       '14
                            sql = sql & "[SRequestList], [SRequestList]+'/'+[SRequestListA_A]+'/'+[SRequestListA_C]+'/'+[SRequestListB_A]+'/'+[SRequestListB_C]+'/'+[SRequestListC_A]+'/'+[SRequestListC_C]+'/', "        '16
                            sql = sql & "[Family], [Family]+'/'+[FamilyA_A]+'/'+[FamilyA_C]+'/'+[FamilyB_A]+'/'+[FamilyB_C]+'/'+[FamilyC_A]+'/'+[FamilyC_C]+'/', "      '18
                            sql = sql & "[ST1], [ST1]+'/'+[ST1A_A]+'/'+[ST1A_C]+'/'+[ST1B_A]+'/'+[ST1B_C]+'/'+[ST1C_A]+'/'+[ST1C_C]+'/', "       '20
                            sql = sql & "[ST2], [ST2]+'/'+[ST2A_A]+'/'+[ST2A_C]+'/'+[ST2B_A]+'/'+[ST2B_C]+'/'+[ST2C_A]+'/'+[ST2C_C]+'/', "       '22
                            sql = sql & "[ST3], [ST3]+'/'+[ST3A_A]+'/'+[ST3A_C]+'/'+[ST3B_A]+'/'+[ST3B_C]+'/'+[ST3C_A]+'/'+[ST3C_C]+'/', "       '24
                            sql = sql & "[ST4], [ST4]+'/'+[ST4A_A]+'/'+[ST4A_C]+'/'+[ST4B_A]+'/'+[ST4B_C]+'/'+[ST4C_A]+'/'+[ST4C_C]+'/', "       '26
                            sql = sql & "[ST5], [ST5]+'/'+[ST5A_A]+'/'+[ST5A_C]+'/'+[ST5B_A]+'/'+[ST5B_C]+'/'+[ST5C_A]+'/'+[ST5C_C]+'/', "       '28
                            sql = sql & "[ST6], [ST6]+'/'+[ST6A_A]+'/'+[ST6A_C]+'/'+[ST6B_A]+'/'+[ST6B_C]+'/'+[ST6C_A]+'/'+[ST6C_C]+'/', "       '30
                            sql = sql & "[ST7], [ST7]+'/'+[ST7A_A]+'/'+[ST7A_C]+'/'+[ST7B_A]+'/'+[ST7B_C]+'/'+[ST7C_A]+'/'+[ST7C_C]+'/', "       '32
                            sql = sql & "[OTHER], [OTHER]+'/'+[OTHERA_A]+'/'+[OTHERA_C]+'/'+[OTHERB_A]+'/'+[OTHERB_C]+'/'+[OTHERC_A]+'/'+[OTHERC_C]+'/' "       '34
                            '
                            sql = sql & "From W_EXCEPTLIMITED "
                            sql = sql & "where Active = 1 "
                            sql = sql & "and Cat = '" & dtRule.Rows(i)(35) & "' "
                            sql = sql & "and Rno = '" & dtRule.Rows(i)(36) & "' "
                            sql = sql & "and SubNo = '" & dtRule.Rows(i)(37) & "' "
                            '
                            'MODIFY-START 241014
                            '  限定:[非ISIP]=不去除字尾SK及-B關鍵字
                            '  限定:[ISIP]=去除字尾SK及-B關鍵字
                            '
                            '如果SubNo是SLS(PULLER是HH054)，只取得Slider相符的資料
                            'If InStr(dtRule.Rows(i)(37), "SLS-1") > 0 Then
                            '    sql = sql & " and '" & iNew(4) & "' like Slider1 "
                            'ElseIf InStr(dtRule.Rows(i)(37), "SLS-2") > 0 Then
                            '    sql = sql & " and '" & iNew(6) & "' like Slider2 "
                            'End If
                            If InStr(dtRule.Rows(i)(37), "SLS-1") > 0 Then
                                sql = sql & " and '" & ISIPSlider1 & "' like Slider1 "
                            ElseIf InStr(dtRule.Rows(i)(37), "SLS-2") > 0 Then
                                sql = sql & " and '" & ISIPSlider2 & "' like Slider2 "
                            End If
                            If InStr(dtRule.Rows(i)(36), "GRIND") > 0 Then
                                sql = sql & " and '" & ISIPSlider1 & "' like Slider1 "
                            End If
                            '
                            'MODIFY-END
                            '
                            Dim xSpcExcep As String
                            xSpcExcep = xSpc
                            'PULLER是HH054'，要把特殊要求中的顏色做關鍵字Replace，再跟限定規則比對
                            If InStr(dtRule.Rows(i)(37), "SLS") > 0 Then
                                xSpcExcep = fpObj.ReplaceColorString(xSpc)
                            End If
                            '
                            sql = sql & "order by Active,Cat,Rno,SubNo "
                            Dim dtExceptRule As DataTable = uDataBase.GetDataTable(sql)
                            '
                            If dtExceptRule.Rows.Count > 0 Then
                                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & ">>例外檢測-START"
                                '
                                For x = 0 To dtExceptRule.Rows.Count - 1
                                    ' -----------------------------------------------------------------
                                    ' 是否限定ITEM
                                    ' -----------------------------------------------------------------
                                    y = 0
                                    If (dtExceptRule.Rows(x)(y) = "*" Or LimitedItem(iNew(y), dtExceptRule.Rows(x)(y + 1)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 2) = "*" Or LimitedItem(iNew(y + 1), dtExceptRule.Rows(x)(y + 3)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 4) = "*" Or LimitedItem(iNew(y + 2), dtExceptRule.Rows(x)(y + 5)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 6) = "*" Or LimitedItem(iNew(y + 3), dtExceptRule.Rows(x)(y + 7)) = True) And _
                                       ((dtExceptRule.Rows(x)(y + 8) = "*" Or LimitedItem(iNew(y + 4), dtExceptRule.Rows(x)(y + 9)) = True) Or _
                                       (dtExceptRule.Rows(x)(y + 8) = "*" Or LimitedItem(ISIPSlider1, dtExceptRule.Rows(x)(y + 9)) = True)) And _
                                       (dtExceptRule.Rows(x)(y + 10) = "*" Or LimitedItem(iNew(y + 5), dtExceptRule.Rows(x)(y + 11)) = True) And _
                                       ((dtExceptRule.Rows(x)(y + 12) = "*" Or LimitedItem(iNew(y + 6), dtExceptRule.Rows(x)(y + 13)) = True) Or _
                                       (dtExceptRule.Rows(x)(y + 12) = "*" Or LimitedItem(ISIPSlider2, dtExceptRule.Rows(x)(y + 13)) = True)) And _
                                       (dtExceptRule.Rows(x)(y + 14) = "*" Or LimitedItem(iNew(y + 7), dtExceptRule.Rows(x)(y + 15)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 16) = "*" Or LimitedItem(xSpcExcep, dtExceptRule.Rows(x)(y + 17)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 18) = "*" Or LimitedItem(iNew(y + 9), dtExceptRule.Rows(x)(y + 19)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 20) = "*" Or LimitedItem(iNew(y + 10), dtExceptRule.Rows(x)(y + 21)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 22) = "*" Or LimitedItem(iNew(y + 11), dtExceptRule.Rows(x)(y + 23)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 24) = "*" Or LimitedItem(iNew(y + 12), dtExceptRule.Rows(x)(y + 25)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 26) = "*" Or LimitedItem(iNew(y + 13), dtExceptRule.Rows(x)(y + 27)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 28) = "*" Or LimitedItem(iNew(y + 14), dtExceptRule.Rows(x)(y + 29)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 30) = "*" Or LimitedItem(iNew(y + 15), dtExceptRule.Rows(x)(y + 31)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 32) = "*" Or LimitedItem(iNew(y + 16), dtExceptRule.Rows(x)(y + 33)) = True) Then
                                        '
                                        'MsgBox("[" & dtExceptRule.Rows(x)(y + 34) & "]")
                                        If dtExceptRule.Rows(x)(y + 34) <> "*" Then
                                            'PRICE A001/A206/A211/A999/K206/K211
                                            If DA001.Checked = True Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), DA001.Text) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                            If DA206.Checked = True Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), DA206.Text) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                            If DA211.Checked = True Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), DA211.Text) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                            If DA999.Checked = True Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), DA999.Text) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                            If DK206.Checked = True Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), DK206.Text) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                            If DK211.Checked = True Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), DK211.Text) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                            'BUYER/CUSTOMER/BUYER-CUSTOMER
                                            str = ""
                                            If DBuyerCode.Text <> "" And DCustomerCode.Text <> "" Then
                                                str = DBuyerCode.Text & "-" & DCustomerCode.Text
                                            Else
                                                If DBuyerCode.Text <> "" Then
                                                    str = DBuyerCode.Text
                                                Else
                                                    If DCustomerCode.Text <> "" Then
                                                        str = DCustomerCode.Text
                                                    End If
                                                End If
                                            End If
                                            If str <> "" Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), str) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                            '用途DForUse
                                            If DForUse.Text <> "" Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), DForUse.Text) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                            '備註DRemark
                                            If DRemark.Text <> "" Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), DRemark.Text) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                        Else
                                            DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                            uError = False
                                        End If
                                    End If
                                    '只要符合一筆例外就跳出白名單迴圈
                                    If uError = False Then
                                        Exit For
                                    End If
                                Next
                                '
                                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & ">>例外檢測-END"
                            End If
                        Else
                            DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "無限定" & "]"
                        End If
                        '
                        CkLimitedLink.Visible = False
                        If uError = True Then
                            msg = dtRule.Rows(i)(41)
                            If dtRule.Rows(i)(35) = "5.WARNING" Then
                                '不跳出迴圈，記錄所有WARNING的訊息，全部檢測完再顯示
                                sWarnMsg = sWarnMsg & msgNo.ToString() & ". " & msg & "\n"
                                msgNo = msgNo + 1
                                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & msg 'AppendLimitedItem時存入供日後統計用
                            Else
                                AppendLimitedItem(True)         '限定ERROR --> ADD DB
                                '遇到一個錯誤就跳出迴圈，開啟限定檢查頁面
                                Exit For
                            End If
                        End If
                    End If
                Next
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "限定檢測-END"
                '
                'Warning
                If pCat = "WARNING" Then
                    If sWarnMsg <> "" Then
                        DWarnMsg.Value = sWarnMsg
                        AppendLimitedItem(False)        'Warning --> ADD DB但不顯示限定錯誤畫面
                        DCustUnd.Disabled = False
                        'SHOW MSG請申請者勾選客戶了解
                        uJavaScript.PopMsg(Me, sWarnMsg & "\n請勾選[客戶了解]才能申請!")
                    End If
                    uError = True
                End If
                '
            End If
        End If
        '
        '不顯示錯誤訊息，改為直接開啟限定檢查頁面
        'If uError Then uJavaScript.PopMsg(Me, msg)
        '
        Return uError
    End Function
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(LimitedItemOther)
    '**     1.10字縮碼檢查 20240808
    '**     2.生產可行性 20240819
    '**     3.表面處理 20240820
    '**     4.檢針 20241025
    '**     5.P-PACK 20240907
    '**     6.TWIST4 20240907
    '**     7.SCD 20240907
    '**     8.特別要求縮寫 20241009
    '**     9.型別 20241109 
    '**     A.MFGAP 20241109 
    '**
    '*****************************************************************
    Public Function LimitedItemOther(ByVal pCat As String) As Boolean
        Dim i, j As Integer
        Dim xSlider, xBody, xBody1, xPuller, xColor, xSpecial, xALL, xNewSpecial, xChainType As String
        Dim sql, str, xPullerStr, xKensinStr, xRuleStr, xKensinKey, xNantiKey, msg As String
        Dim iSliderList, iFinishList, iKensinList, iRuleList As String()
        Dim uError As Boolean = False
        Dim xKensin As Boolean = False
        '
        msg = ""
        ' --------------------------------
        '**     1.10字縮碼檢查 20240808
        ' Start
        ' MsgBox(pCat)
        If pCat = "SLIDERSHORT" Then
            '
            If DSlider1.Text.Trim & DSlider2.Text.Trim <> "" Then
                ' Slider Information
                str = ""
                If DSlider1.Text <> "" Then str = DSlider1.Text.Trim & "!"
                If DSlider2.Text <> "" Then str = str & DSlider2.Text.Trim & "!"
                iSliderList = str.ToString.Split("!")
                '
                If UBound(iSliderList) > 0 Then
                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "縮碼檢測-Start" & "]"
                End If
                '
                For i = 0 To UBound(iSliderList) - 1
                    xSlider = ""
                    xBody = ""
                    xPuller = ""
                    xColor = ""
                    xSpecial = ""
                    '
                    If Len(iSliderList(i)) = 10 Then
                        ' Get Body
                        sql = "SELECT BODY, BodyGroup From V_EDXBodyList "
                        sql = sql & "where '" & iSliderList(i) & "' like body+'%' "
                        sql = sql & "ORDER BY LEN([BODY]) DESC, BODY "
                        Dim dtBODY1 As DataTable = uDataBase.GetDataTable(sql)
                        If dtBODY1.Rows.Count > 0 Then
                            xBody = dtBODY1.Rows(0)("BODY").ToString.Trim
                        End If
                        ' Get Special
                        sql = "SELECT ShortSpecial From V_SliderShortSpecial "
                        sql = sql & "where '" & iSliderList(i) & "' like '%' + ShortSpecial "
                        sql = sql & "ORDER BY LEN([ShortSpecial]) DESC "
                        Dim dtSpecial As DataTable = uDataBase.GetDataTable(sql)
                        If dtSpecial.Rows.Count > 0 Then
                            xSpecial = dtSpecial.Rows(0)("ShortSpecial").ToString.Trim
                        End If
                        ' Get Puller & Color
                        str = iSliderList(i)
                        If xBody <> "" Then str = Replace(str, xBody, "").Trim
                        If xSpecial <> "" Then str = Replace(str, xSpecial, "").Trim
                        j = Len(str)
                        Do Until j = 0
                            If Mid(str, j, 1) >= "0" And Mid(str, j, 1) <= "9" Then
                                '
                                If xPuller = "" Then
                                    j = 1
                                    Do Until j = Len(str) + 1
                                        If Mid(str, j, 1) >= "0" And Mid(str, j, 1) <= "Z" Then
                                            xPuller = Mid(str, 1, j).Trim
                                        Else
                                            Exit Do
                                        End If
                                        j = j + 1
                                    Loop
                                End If
                                '
                                Exit Do
                            Else
                                xPuller = Mid(str, 1, j - 1).Trim
                            End If
                            j = j - 1
                        Loop
                        '
                        If str <> "" And InStr(str, xPuller) > 0 Then
                            xColor = Replace(str, xPuller, "").Trim
                        End If
                        '
                        'If xPuller = "" Then
                        'End If
                        'msg = "slider=" & iSliderList(i) & "=[" & "BODY=" & xBody & "] [" & _
                        '                                                                        "PULLER=" & xPuller & "] [" & _
                        '                                                                        "COLOR=" & xColor & "] [" & _
                        '                                                                        "SPECIAL=" & xSpecial & "]"
                        'MsgBox(msg)
                        '
                        'A.3Y SALES & FA000 (M_WINGSEDX)
                        If xBody <> "" Or xPuller <> "" Then
                            '
                            str = ""
                            For j = 1 To Len(xColor)
                                str = str & "_"
                            Next
                            xSlider = xBody & xPuller & str & xSpecial
                            '
                            'msg = "slider=" & iSliderList(i) & "=[" & "BODY=" & xBody & "] [" & _
                            '                                              "PULLER=" & xPuller & "] [" & _
                            '                                              "COLOR=" & str & "] [" & _
                            '                                              "SPECIAL=" & xSpecial & "]"
                            'MsgBox(msg)
                            '
                            'xslide
                            '
                            sql = "SELECT Top 1 Slider, OrderNo, Item, ItemName "
                            sql = sql & "From M_WINGSEDX "
                            sql = sql & "Where LTRIM(RTRIM(Slider)) LIKE '" & xSlider & "' "
                            sql = sql & "   Or LTRIM(RTRIM(Slider)) = '" & iSliderList(i) & "' "
                            sql = sql & "Order By OrderNo desc, Item desc "
                            'MsgBox(sql)
                            Dim dtSales1 As DataTable = uDataBase.GetDataTable(sql)
                            If dtSales1.Rows.Count > 0 Then
                                'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                                '                        "[OK]" & "[WINGS:" & dtSales1.Rows(0)("Item").ToString.Trim & "  " & dtSales1.Rows(0)("ItemName").ToString.Trim & "]"
                            Else
                                xPullerStr = ""
                                j = 1
                                Do Until j = Len(xPuller)
                                    If Mid(xPuller, j, 1) >= "A" And Mid(xPuller, j, 1) <= "Z" Then
                                        xPullerStr = Mid(xPuller, 1, j).Trim
                                    Else
                                        Exit Do
                                    End If
                                    j = j + 1
                                Loop
                                '
                                For j = Len(xPullerStr) + 1 To Len(xPuller)
                                    xPullerStr = xPullerStr & "_"
                                Next
                                '
                                'msg = "slider=" & iSliderList(i) & "=[" & "BODY=" & xBody & "] [" & _
                                '                                              "PULLER=" & xPuller & "] [" & _
                                '                                              "PULLER=" & xPullerStr & "] [" & _
                                '                                              "COLOR=" & str & "] [" & _
                                '                                              "SPECIAL=" & xSpecial & "]"
                                'MsgBox(msg)
                                '
                                sql = "SELECT Top 1 ShortSlider, Slider "
                                sql = sql & "From V_SliderShortRule "
                                sql = sql & "Where LTRIM(RTRIM(ShortBody)) = '" & xBody & "' "

                                sql = sql & "  And ( "
                                sql = sql & "        LTRIM(RTRIM(ShortPuller)) = '" & xPuller & "' "
                                sql = sql & "     OR LTRIM(RTRIM(ShortPuller)) LIKE '" & xPullerStr & "' "
                                sql = sql & "  ) "

                                sql = sql & "  And LTRIM(RTRIM(ShortSpecial)) = '" & xSpecial & "' "
                                sql = sql & "  And LTRIM(RTRIM(ShortColor)) = '" & str & "' "
                                sql = sql & "Order By ShortBody, ShortPuller, ShortColor, ShortSpecial "
                                '
                                'MsgBox(sql)
                                '
                                Dim dtShort1 As DataTable = uDataBase.GetDataTable(sql)
                                If dtShort1.Rows.Count > 0 Then
                                    'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                                    '                        "[OK]" & "[ShortList:" & dtShort1.Rows(0)("ShortSlider").ToString.Trim & "  " & dtShort1.Rows(0)("Slider").ToString.Trim & "]"
                                Else
                                    '
                                    '是縮寫 PULLER = ERROR
                                    '
                                    sql = "SELECT Buyer, Puller, Short "
                                    sql = sql & "From V_PullerShort "
                                    sql = sql & "Where Puller = '" & xPuller & "' "
                                    sql = sql & "Order By Buyer, Puller, Short "
                                    Dim dtPuller1 As DataTable = uDataBase.GetDataTable(sql)
                                    If dtPuller1.Rows.Count > 0 Then
                                        uError = True
                                        msg = "slider=" & iSliderList(i) & "=[" & "BODY=" & xBody & "] [" & _
                                                                                    "PULLER=" & xPuller & "] [" & _
                                                                                    "COLOR=" & xColor & "] [" & _
                                                                                    "SPECIAL=" & xSpecial & "]"
                                        DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[NG-01]A1" & msg
                                    End If
                                    '
                                End If
                                '
                            End If
                        Else
                            uError = True
                            msg = "slider=" & iSliderList(i) & "=[" & "BODY=" & xBody & "] [" & _
                                                                         "PULLER=" & xPuller & "] [" & _
                                                                         "COLOR=" & xColor & "] [" & _
                                                                         "SPECIAL=" & xSpecial & "]"
                            DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[NG-01]A2" & msg
                        End If
                    End If
                    '
                    If InStr(DLimitedItem.Text, "[NG-01]") > 0 Then
                        '
                        sql = "Insert into M_SliderShortRule ( "
                        sql = sql & "Active, ShortBody,ShortPuller,ShortColor,ShortSpecial, "
                        sql = sql & "Body,Puller,Color,Special, "
                        sql = sql & "ShortSlider,Slider,Buyer,Person,Remark, "
                        sql = sql & "Yobi1,Yobi2,Yobi3, "
                        sql = sql & "[CreateUser], [CreateTime], [ModifyUser], [ModifyTime] "
                        sql = sql & ") "
                        sql = sql & "VALUES( "
                        sql &= "'0' ,"
                        sql &= "'" & xBody & "' ,"
                        sql &= "'" & xPuller & "' ,"
                        sql &= "'" & xColor & "' ,"
                        sql &= "'" & xSpecial & "' ,"
                        '
                        sql &= "'', '', '', '', "
                        sql &= "'', '', '', '', '', "
                        sql &= "'ERROR', '', '', "
                        '
                        sql &= "'" & Request.QueryString("pUserID") & "' ,"
                        sql &= "'" & NowDateTime & "' ,"
                        sql &= "'" & Request.QueryString("pUserID") & "' ,"
                        sql &= "'" & NowDateTime & "' )"
                        uDataBase.ExecuteNonQuery(sql)
                        '
                    End If
                    '
                Next
                '
                If UBound(iSliderList) > 0 Then
                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "縮碼檢測-End" & "]"
                End If
                '
                If InStr(DLimitedItem.Text, "NG-01") > 0 Then
                    uJavaScript.PopMsg(Me, "[Slider Code]縮寫是否正確？ (無法申請)")
                End If
                '
            End If
            '
        End If
        'End
        ' --------------------------------
        '**     2.生產可行性 20240819 
        ' Start
        If pCat = "PRODREADY" Then
            '
            '--------------------------------------
            '
            '限完成品ZIPPER檢測
            If DST1.Text.Trim & DST4.Text.Trim = "11" Then
                ' Slider Information
                str = ""
                str = DSlider1.Text.Trim & "!" & DSlider2.Text.Trim & "!"
                iSliderList = str.ToString.Split("!")
                '
                If UBound(iSliderList) > 0 Then
                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "生產可行性檢測-Start" & "]"
                End If
                '
                'IRW申請內容，查詢81-41 無申請組合(SIZE+CHAIN+CLASS+TAPE+SLIDER BODY+特殊要求)需附件測試報告或可生產確認信
                ' Get Body
                str = ""
                xBody = ""
                xBody1 = ""
                For i = 0 To UBound(iSliderList) - 1
                    sql = "SELECT BODY, BodyGroup From V_EDXBodyList "
                    sql = sql & "where '" & iSliderList(i) & "' like body+'%' "
                    sql = sql & "ORDER BY LEN([BODY]) DESC, BODY "
                    Dim dtBODY1 As DataTable = uDataBase.GetDataTable(sql)
                    If dtBODY1.Rows.Count > 0 Then
                        str = dtBODY1.Rows(0)("BODY").ToString.Trim
                    Else
                        str = ""
                    End If
                    '
                    Select Case i
                        Case 0
                            xBody = str.Trim & "/"
                            If Left(str.Trim, 1) = "E" Then
                                xBody1 = UCase("D" & Mid(str.Trim, 2, 99) & "/").Trim
                            Else
                                xBody1 = str.Trim & "/"
                            End If
                        Case Else
                            xBody = xBody & str.Trim & "/"
                            If Left(str.Trim, 1) = "E" Then
                                xBody1 = xBody1 & UCase("D" & Mid(str.Trim, 2, 99) & "/").Trim
                            Else
                                xBody1 = xBody1 & str.Trim & "/"
                            End If
                    End Select
                Next
                '
                'MsgBox(xBody1)
                '
                ' Check
                xSpecial = UCase(DSRequest1.Text.Trim) & "/" & UCase(DSRequest2.Text.Trim) & "/" & _
                           UCase(DSRequest3.Text.Trim) & "/" & UCase(DSRequest4.Text.Trim) & "/" & _
                           UCase(DSRequest5.Text.Trim) & "/" & UCase(DSRequest6.Text.Trim) & "/"
                xSpecial = Replace(xSpecial, "XXX", "")
                '
                ' Modify-Start by Joy 20250611
                ' 取得新 Special Feat.
                '    Public Function GetNewxSpecialFeat(ByVal pOption As String, ByVal pData1 As String, ByVal pData2 As String) As String
                xChainType = GetNewxSpecialFeat("MOD-CHAINTYPE", UCase(DChain.Text.Trim), "")

                xNewSpecial = GetNewxSpecialFeat("DEL-SPECIAL-1", "", xSpecial)
                xNewSpecial = GetNewxSpecialFeat("MOD-SPECIAL-1", "", xNewSpecial)
                xNewSpecial = GetNewxSpecialFeat("MOD-SPECIAL-1B", xBody1, xNewSpecial)
                xNewSpecial = GetNewxSpecialFeat("MOD-SPECIAL-1C", "", xNewSpecial)
                For i = 0 To 6
                    xNewSpecial = Replace(xNewSpecial, "//", "/")
                Next
                'MsgBox(xChainType & "--" & xNewSpecial)
                ''
                '' Special = XXX 測試ITEM使用 --> (不列入比較)
                ''         = KENSIN, ND-B 檢針 --> (不列入比較)
                ''         = N-ANTI 低鎳 --> (不列入比較)
                'xNewSpecial = xSpecial
                'xNewSpecial = Replace(xNewSpecial, "KENSIN/", "/")
                'xNewSpecial = Replace(xNewSpecial, "ND-B/", "/")
                'xNewSpecial = Replace(xNewSpecial, "N-ANTI/", "/")
                ''For i = 0 To 6
                ''    xNewSpecial = Replace(xNewSpecial, "//", "/")
                ''Next
                '
                ' Modify-End
                '
                xALL = DSize.Text & "/" & DChain.Text & "/" & _
                       DClass.Text & "/" & DTape.Text & "/" & _
                       DSlider1.Text & "/" & DFinish1.Text & "/" & _
                       DSlider2.Text & "/" & DFinish2.Text & "/" & DFamily.Text & "/" & xSpecial

                msg = UCase(DSize.Text.Trim) & "|" & UCase(DChain.Text.Trim) & "|" & UCase(DClass.Text.Trim) & "|" & _
                      UCase(DTape.Text.Trim) & "|" & UCase(DSlider1.Text.Trim) & "|" & UCase(DSlider2.Text.Trim) & "|" & UCase(xSpecial.Trim) & "|" & Chr(13) & _
                      UCase(xBody.Trim) & "|" & UCase(xSpecial.Trim) & "|"
                'MsgBox(msg)
                'MsgBox(UCase(xSpecial.Trim) & "--" & Chr(13) & UCase(xNewSpecial.Trim))
                '
                sql = "SELECT Top 1 SIZE+'|'+CHAINTYPE+'|'+CLASS+'|'+TAPE1+'|'+BODY1+'|'+BODY2+'|'+SpecialList as PRODREADY "
                sql = sql & "From V_IRWPerformaceUp_COMBI "
                sql = sql & "Where CAT IN ('WCOMBI', 'HCOMBI') "
                sql = sql & "  And Active = 1 "
                sql = sql & "  And ( "
                sql = sql & "        ( "
                sql = sql & "        LTRIM(RTRIM(SIZE)) = '" & UCase(DSize.Text.Trim) & "' "
                '
                sql = sql & "        And ( "
                sql = sql & "          LTRIM(RTRIM(CHAINTYPE)) = '" & UCase(DChain.Text.Trim) & "' "
                sql = sql & "          OR "
                sql = sql & "          LTRIM(RTRIM(CHAINTYPE)) LIKE '" & UCase(xChainType.Trim) & "' "
                sql = sql & "        ) "
                sql = sql & "        And LTRIM(RTRIM(CLASS)) = '" & UCase(DClass.Text.Trim) & "' "
                sql = sql & "        And LTRIM(RTRIM(TAPE1)) = '" & UCase(DTape.Text.Trim) & "' "
                '
                sql = sql & "        And ( "
                sql = sql & "          LTRIM(RTRIM(BODY1)) + '/' + LTRIM(RTRIM(BODY2)) +'/' = '" & UCase(xBody.Trim) & "' "
                sql = sql & "          OR "
                sql = sql & "          LTRIM(RTRIM(BODY1)) + '/' + LTRIM(RTRIM(BODY2)) +'/' = '" & UCase(xBody1.Trim) & "' "
                sql = sql & "        ) "
                '
                sql = sql & "        And ( "
                sql = sql & "              LTRIM(RTRIM(SpecialList)) = '" & UCase(xSpecial.Trim) & "' "
                sql = sql & "              Or "
                sql = sql & "              LTRIM(RTRIM(NewSpecialList)) = '" & UCase(xNewSpecial.Trim) & "' "
                sql = sql & "            ) "
                '
                sql = sql & "        ) "
                sql = sql & "        OR "
                sql = sql & "        ( "
                sql = sql & "        Seqno =  999 "
                sql = sql & "        And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Size))='' THEN '%%' ELSE LTRIM(RTRIM(Size)) END) "
                sql = sql & "        And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(ChainType))='' THEN '%%' ELSE LTRIM(RTRIM(ChainType)) END) "
                sql = sql & "        And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Class))='' THEN '%%' ELSE LTRIM(RTRIM(Class)) END) "
                sql = sql & "        And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Tape1))='' THEN '%%' ELSE LTRIM(RTRIM(Tape1)) END) "
                sql = sql & "        And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider1))='' THEN '%%' ELSE LTRIM(RTRIM(Slider1)) END) "
                sql = sql & "        And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Body1))='' THEN '%%' ELSE LTRIM(RTRIM(Body1)) END) "
                sql = sql & "        And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Finish1))='' THEN '%%' ELSE LTRIM(RTRIM(Finish1)) END) "
                sql = sql & "        And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider2))='' THEN '%%' ELSE LTRIM(RTRIM(Slider2)) END) "
                sql = sql & "        And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Body2))='' THEN '%%' ELSE LTRIM(RTRIM(Body2)) END) "
                sql = sql & "        And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Finish2))='' THEN '%%' ELSE LTRIM(RTRIM(Finish2)) END) "
                sql = sql & "        And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Family))='' THEN '%%' ELSE LTRIM(RTRIM(Family)) END) "
                sql = sql & "        And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(SpecialList))='' THEN '%%' ELSE LTRIM(RTRIM(SpecialList)) END) "
                sql = sql & "        ) "
                sql = sql & "  ) "
                sql = sql & "Order By SIZE, CHAINTYPE, CLASS, TAPE1, BODY1, BODY2, SpecialList "
                '
                'MsgBox(sql)
                '
                Dim dtPerformaceUp1 As DataTable = uDataBase.GetDataTable(sql)
                If dtPerformaceUp1.Rows.Count > 0 Then
                    'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                    '                        "[OK]" & "[WINGS:" & dtPerformaceUp1.Rows(0)("PRODREADY").ToString.Trim & "]"
                Else
                    uError = True
                    msg = UCase(DSize.Text.Trim) & "|" & UCase(DChain.Text.Trim) & "|" & UCase(DClass.Text.Trim) & "|" & _
                          UCase(DTape.Text.Trim) & "|" & UCase(xBody.Trim) & "|" & UCase(xSpecial.Trim) & "|"
                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[NG-02]B1" & msg & Chr(13) & "** 需附件測試報告或可生產確認信 **"
                End If
                '
                If InStr(DLimitedItem.Text, "[NG-02]") > 0 Then
                    ' ERROR
                    sql = "Insert into M_IRWPerformaceUp ( "
                    sql = sql & "Active,Cat,Seqno,Size,ChainType, "
                    sql = sql & "Class,Tape1,Slider1,Body1,Finish1, "
                    sql = sql & "Slider2,Body2,Finish2,Family,SpecialList, "
                    sql = sql & "RDSpec,RDFinish,YOBI1,YOBI2,YOBI3, "
                    sql = sql & "CreateUser, CreateTime, ModifyUser, ModifyTime "
                    sql = sql & ") "
                    sql = sql & "VALUES( "
                    sql &= "'0' ,"
                    sql &= "'" & "HCOMBI" & "' ,"
                    sql &= "1, "
                    sql &= "'" & UCase(DSize.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DChain.Text.Trim) & "' ,"
                    '
                    sql &= "'" & UCase(DClass.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DTape.Text.Trim) & "' ,"
                    sql &= "'', "
                    sql &= "'" & UCase(xBody.Trim) & "' ,"
                    sql &= "'', "
                    '
                    sql &= "'', '', '', '', "
                    sql &= "'" & UCase(xSpecial.Trim) & "' ,"
                    '
                    sql &= "'', '', 'ERROR', '', '', "
                    '
                    sql &= "'" & Request.QueryString("pUserID") & "' ,"
                    sql &= "'" & NowDateTime & "' ,"
                    sql &= "'" & Request.QueryString("pUserID") & "' ,"
                    sql &= "'" & NowDateTime & "' )"
                    uDataBase.ExecuteNonQuery(sql)
                    '
                End If
                '
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "生產可行性-End" & "]"
                '
                If InStr(DLimitedItem.Text, "[NG-02]") > 0 Then
                    uJavaScript.PopMsg(Me, "此新組合產品,需要需附件測試報告或可生產確認信件！ (無附證明將無法申請)")
                End If
                '
            End If
            '
            '--------------------------------------
            '
            '限CH-材料
            If DST1.Text.Trim & DST4.Text.Trim & DST6.Text.Trim = "12M" Then
                '
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "生產可行性檢測-Start" & "]"
                '
                ' Check
                xSpecial = UCase(DSRequest1.Text.Trim) & "/" & UCase(DSRequest2.Text.Trim) & "/" & _
                           UCase(DSRequest3.Text.Trim) & "/" & UCase(DSRequest4.Text.Trim) & "/" & _
                           UCase(DSRequest5.Text.Trim) & "/" & UCase(DSRequest6.Text.Trim) & "/"
                xSpecial = Replace(xSpecial, "XXX", "")
                '
                ' Special = XXX 測試ITEM使用 --> (不列入比較)
                '         = KENSIN, ND-B 檢針 --> (不列入比較)
                '         = N-ANTI 低鎳 --> (不列入比較)
                xNewSpecial = xSpecial
                xNewSpecial = Replace(xNewSpecial, "KENSIN/", "/")
                xNewSpecial = Replace(xNewSpecial, "ND-B/", "/")
                xNewSpecial = Replace(xNewSpecial, "N-ANTI/", "/")
                '
                xALL = DSize.Text & "/" & DChain.Text & "/" & _
                       DClass.Text & "/" & DTape.Text & "/" & _
                       DSlider1.Text & "/" & DFinish1.Text & "/" & _
                       DSlider2.Text & "/" & DFinish2.Text & "/" & DFamily.Text & "/" & xSpecial
                '
                sql = "SELECT Top 1 SIZE+'|'+CHAINTYPE+'|'+CLASS+'|'+TAPE1+'|'+BODY1+'|'+BODY2+'|'+SpecialList as PRODREADY "
                sql = sql & "From V_IRWPerformaceUp_COMBI "
                sql = sql & "Where CAT IN ('HCOMBI') "
                sql = sql & "  And Active = 1 "
                sql = sql & "  And Seqno =  990 "
                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Size))='' THEN '%%' ELSE LTRIM(RTRIM(Size)) END) "
                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(ChainType))='' THEN '%%' ELSE LTRIM(RTRIM(ChainType)) END) "
                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Class))='' THEN '%%' ELSE LTRIM(RTRIM(Class)) END) "
                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Tape1))='' THEN '%%' ELSE LTRIM(RTRIM(Tape1)) END) "
                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider1))='' THEN '%%' ELSE LTRIM(RTRIM(Slider1)) END) "
                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Body1))='' THEN '%%' ELSE LTRIM(RTRIM(Body1)) END) "
                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Finish1))='' THEN '%%' ELSE LTRIM(RTRIM(Finish1)) END) "
                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider2))='' THEN '%%' ELSE LTRIM(RTRIM(Slider2)) END) "
                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Body2))='' THEN '%%' ELSE LTRIM(RTRIM(Body2)) END) "
                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Finish2))='' THEN '%%' ELSE LTRIM(RTRIM(Finish2)) END) "
                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Family))='' THEN '%%' ELSE LTRIM(RTRIM(Family)) END) "
                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(SpecialList))='' THEN '%%' ELSE LTRIM(RTRIM(SpecialList)) END) "
                sql = sql & "Order By SIZE, CHAINTYPE, CLASS, TAPE1, BODY1, BODY2, SpecialList "
                '
                Dim dtPerformaceUp1 As DataTable = uDataBase.GetDataTable(sql)
                If dtPerformaceUp1.Rows.Count > 0 Then
                    uError = True
                    msg = UCase(DST1.Text.Trim) & "|" & UCase(DST4.Text.Trim) & "|" & UCase(DST6.Text.Trim) & "|" & _
                          UCase(DSize.Text.Trim) & "|" & UCase(DChain.Text.Trim) & "|" & UCase(DClass.Text.Trim) & "|" & _
                          UCase(DTape.Text.Trim) & "|" & UCase(xSpecial.Trim) & "|"
                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[NG-02]C1:" & msg & Chr(13) & "** 需附件測試報告或可生產確認信 **"
                End If
                '
                If InStr(DLimitedItem.Text, "[NG-02]") > 0 Then
                    ' ERROR
                    sql = "Insert into M_IRWPerformaceUp ( "
                    sql = sql & "Active,Cat,Seqno,Size,ChainType, "
                    sql = sql & "Class,Tape1,Slider1,Body1,Finish1, "
                    sql = sql & "Slider2,Body2,Finish2,Family,SpecialList, "
                    sql = sql & "RDSpec,RDFinish,YOBI1,YOBI2,YOBI3, "
                    sql = sql & "CreateUser, CreateTime, ModifyUser, ModifyTime "
                    sql = sql & ") "
                    sql = sql & "VALUES( "
                    sql &= "'0' ,"
                    sql &= "'" & "HCOMBI" & "' ,"
                    sql &= "1, "
                    sql &= "'" & UCase(DSize.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DChain.Text.Trim) & "' ,"
                    '
                    sql &= "'" & UCase(DClass.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DTape.Text.Trim) & "' ,"
                    sql &= "'', "
                    sql &= "'" & UCase(DST1.Text.Trim) & "|" & UCase(DST4.Text.Trim) & "|" & UCase(DST6.Text.Trim) & "|" & "' ,"
                    sql &= "'', "
                    '
                    sql &= "'', '', '', '', "
                    sql &= "'" & UCase(xSpecial.Trim) & "' ,"
                    '
                    sql &= "'', '', 'ERROR', '', '', "
                    '
                    sql &= "'" & Request.QueryString("pUserID") & "' ,"
                    sql &= "'" & NowDateTime & "' ,"
                    sql &= "'" & Request.QueryString("pUserID") & "' ,"
                    sql &= "'" & NowDateTime & "' )"
                    uDataBase.ExecuteNonQuery(sql)
                    '
                End If
                '
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "生產可行性-End" & "]"
                '
                If InStr(DLimitedItem.Text, "[NG-02]") > 0 Then
                    uJavaScript.PopMsg(Me, "此新組合產品,需要需附件測試報告或可生產確認信件！ (無附證明將無法申請)")
                End If
                '
            End If
            '
        End If
        ' End
        ' --------------------------------
        '**     3.表面處理 20240820
        ' Start
        If pCat = "FINISH" Then
            '限SLIDER
            If DSlider1.Text.Trim & DSlider2.Text.Trim <> "" Then
                ' Slider Information
                str = ""
                If DSlider1.Text <> "" Then str = DSlider1.Text.Trim & "!"
                If DSlider2.Text <> "" Then str = str & DSlider2.Text.Trim & "!"
                iSliderList = str.ToString.Split("!")
                'Finish Information
                str = ""
                If DFinish1.Text <> "" Then str = DFinish1.Text.Trim & "!"
                If DFinish2.Text <> "" Then str = str & DFinish2.Text.Trim & "!"
                iFinishList = str.ToString.Split("!")

                If UBound(iSliderList) > 0 Then
                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "表面處理-Start" & "]"
                End If
                '
                For i = 0 To UBound(iSliderList) - 1
                    xSlider = ""
                    xBody = ""
                    xPuller = ""
                    xColor = ""
                    xSpecial = ""
                    ' Get Body
                    sql = "SELECT BODY, BodyGroup From V_EDXBodyList "
                    sql = sql & "where '" & iSliderList(i) & "' like body+'%' "
                    sql = sql & "ORDER BY LEN([BODY]) DESC, BODY "
                    Dim dtBODY1 As DataTable = uDataBase.GetDataTable(sql)
                    If dtBODY1.Rows.Count > 0 Then
                        xBody = dtBODY1.Rows(0)("BODY").ToString.Trim
                    End If
                    ' Get Special
                    sql = "SELECT ShortSpecial From V_SliderShortSpecial "
                    sql = sql & "where '" & iSliderList(i) & "' like '%' + ShortSpecial "
                    sql = sql & "ORDER BY LEN([ShortSpecial]) DESC "
                    Dim dtSpecial As DataTable = uDataBase.GetDataTable(sql)
                    If dtSpecial.Rows.Count > 0 Then
                        xSpecial = dtSpecial.Rows(0)("ShortSpecial").ToString.Trim
                    End If
                    ' Get Puller & Color
                    str = iSliderList(i)
                    If xBody <> "" Then str = Replace(str, xBody, "").Trim
                    If xSpecial <> "" Then str = Replace(str, xSpecial, "").Trim
                    j = Len(str)
                    Do Until j = 0
                        '
                        If Mid(str, j, 1) >= "0" And Mid(str, j, 1) <= "9" Then
                            '
                            If xPuller = "" Then
                                j = 1
                                Do Until j = Len(str) + 1
                                    If Mid(str, j, 1) >= "0" And Mid(str, j, 1) <= "Z" Then
                                        xPuller = Mid(str, 1, j).Trim
                                    Else
                                        Exit Do
                                    End If
                                    j = j + 1
                                Loop
                                'Do Until j = Len(str) + 1
                                '    If Mid(str, j, 1) >= "0" And Mid(str, j, 1) <= "Z" Then
                                '        xPuller = Mid(str, 1, j).Trim
                                '    Else
                                '        Exit Do
                                '    End If
                                '    j = j + 1
                                'Loop
                            End If
                            '
                            Exit Do
                        Else
                            xPuller = Mid(str, 1, j - 1).Trim
                        End If
                        j = j - 1
                    Loop
                    '
                    If str <> "" And InStr(str, xPuller) > 0 Then
                        xColor = Replace(str, xPuller, "").Trim
                    End If
                    '
                    If xPuller = "" And xColor <> "" Then
                        xPuller = xColor
                        xColor = ""
                    End If
                    '
                    '是縮寫 PULLER = 還原原 PULLER
                    '
                    If xPuller <> "" Then
                        sql = "SELECT TOP 1 Buyer, Puller, Short "
                        sql = sql & "From V_PullerShort "
                        sql = sql & "Where Short like '%" & "ZZZZZZZ" & "%' "
                        sql = sql & "Or Short like '%/" & xPuller & "/%' "
                        sql = sql & "Order By Buyer, Puller, Short "
                        Dim dtPuller1 As DataTable = uDataBase.GetDataTable(sql)
                        If dtPuller1.Rows.Count > 0 Then
                            xPuller = dtPuller1.Rows(0)("Puller").ToString.Trim
                        End If
                    End If
                    '
                    'msg = "slider=" & iSliderList(i) & "=[" & "BODY=" & xBody & "] [" & _
                    '                                                                        "PULLER=" & xPuller & "] [" & _
                    '                                                                        "COLOR=" & xColor & "] [" & _
                    '                                                                        "FINISH=" & iFinishList(i) & "]"
                    'MsgBox(msg)
                    ''
                    'WINGS 81-41
                    If xBody <> "" Or xPuller <> "" Then
                        sql = "SELECT Top 1 SIZE+'|'+FAMILY+'|'+SLIDER1+'|'+FINISH1 as FINISHREADY "
                        sql = sql & "From M_IRWPerformaceUp "
                        sql = sql & "Where CAT IN ('WFIN', 'HFIN') "
                        sql = sql & "  And Active = 1 "
                        sql = sql & "  And LTRIM(RTRIM(SIZE)) = '" & UCase(DSize.Text.Trim) & "' "
                        sql = sql & "  And LTRIM(RTRIM(FAMILY)) = '" & UCase(ReplaceFamily(DSize.Text, DFamily.Text)) & "' "
                        sql = sql & "  And LTRIM(RTRIM(SLIDER1)) LIKE '" & UCase(xBody.Trim & xPuller.Trim) & "%' "
                        sql = sql & "  And LTRIM(RTRIM(FINISH1)) = '" & UCase(iFinishList(i).Trim) & "' "
                        sql = sql & "Order By SIZE+'|'+FAMILY+'|'+SLIDER1+'|'+FINISH1 "
                        'MsgBox(sql)
                        Dim dtPerformaceUp1 As DataTable = uDataBase.GetDataTable(sql)
                        If dtPerformaceUp1.Rows.Count > 0 Then
                            'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                            '                        "[OK]" & "[WINGS:" & dtPerformaceUp1.Rows(0)("FINISHREADY").ToString.Trim & "]"
                        Else
                            sql = "SELECT Top 1 CAT+'|'+SIZE+'|'+SLIDER1+'|'+FINISH1 as INFIN "
                            sql = sql & "From M_IRWPerformaceUp "
                            sql = sql & "Where CAT IN ('INFIN') "
                            sql = sql & "  And Active = 1 "
                            sql = sql & "  And '" & UCase(DSize.Text.Trim) & "' Like LTRIM(RTRIM([SIZE])) "
                            sql = sql & "  And '" & UCase(iSliderList(i).Trim) & "' Like LTRIM(RTRIM([Slider1])) "
                            sql = sql & "  And LTRIM(RTRIM(FINISH1)) = '" & UCase(iFinishList(i).Trim) & "' "
                            sql = sql & "Order By CAT+'|'+SIZE+'|'+SLIDER1+'|'+FINISH1 "
                            Dim dtPerformaceUp2 As DataTable = uDataBase.GetDataTable(sql)
                            If dtPerformaceUp2.Rows.Count > 0 Then
                                '內製-FINISH
                                'msg = "INFIN" & "|" & dtPerformaceUp2.Rows(0)("INFIN").ToString.Trim
                                'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[OK]" & msg
                                '
                            Else
                                '外製-FINISH
                                'SIZE-FAMILY-BODY
                                str = ""
                                If Mid(DSize.Text.Trim, 1, 1) = "0" Then
                                    str = Mid(DSize.Text.Trim, 2, 1)
                                Else
                                    str = DSize.Text.Trim
                                End If
                                str = str & "_" & ReplaceFamily(DSize.Text, DFamily.Text) & "_" & xBody.Trim

                                'MsgBox(UCase(xPuller.Trim) & "--" & UCase(str.Trim))
                                '
                                sql = "SELECT Top 1 CAT+'|'+FINISH1 as SPEC "
                                sql = sql & "From M_IRWPerformaceUp "
                                sql = sql & "Where CAT IN ('RDSPEC') "
                                sql = sql & "  And Active = 1 "
                                sql = sql & "  And LTRIM(RTRIM(SLIDER1)) LIKE '%" & UCase(xPuller.Trim) & "%' "
                                sql = sql & "  And LTRIM(RTRIM(RDSPEC)) LIKE '%" & UCase(str.Trim) & "%' "
                                sql = sql & "Order By CAT+'|'+FINISH1 "
                                Dim dtPerformaceUp4 As DataTable = uDataBase.GetDataTable(sql)
                                If dtPerformaceUp4.Rows.Count > 0 Then
                                    '
                                    sql = "SELECT Top 1 CAT+'|'+FINISH1 as FIN "
                                    sql = sql & "From M_IRWPerformaceUp "
                                    sql = sql & "Where CAT IN ('RDFIN') "
                                    sql = sql & "  And Active = 1 "
                                    sql = sql & "  And LTRIM(RTRIM(SLIDER1)) LIKE '%" & UCase(xPuller.Trim) & "%' "
                                    sql = sql & "  And LTRIM(RTRIM(RDSPEC)) LIKE '%" & UCase(str.Trim) & "%' "
                                    sql = sql & "  And LTRIM(RTRIM(RDSPEC)) LIKE '%" & UCase(iFinishList(i).Trim) & "%' "
                                    sql = sql & "Order By CAT+'|'+FINISH1 "
                                    '
                                    'MsgBox(sql)
                                    '
                                    Dim dtPerformaceUp5 As DataTable = uDataBase.GetDataTable(sql)
                                    If dtPerformaceUp5.Rows.Count > 0 Then
                                        'msg = "FINISH" & "|" & UCase(DSize.Text.Trim) & "|" & UCase(DFamily.Text.Trim) & "|" & UCase(xBody.Trim & xPuller.Trim) & "|" & UCase(iFinishList(i).Trim)
                                        'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[OK]" & msg
                                    Else
                                        uError = True
                                        msg = UCase(DSize.Text.Trim) & "|" & UCase(DFamily.Text.Trim) & "|" & "|" & UCase(xBody.Trim & xPuller.Trim) & "|" & UCase(iFinishList(i).Trim)
                                        DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[NG-03]C2" & msg & Chr(13) & "** 外製:不適用此FINISH **"
                                        '
                                        sql = "Insert into M_IRWPerformaceUp ( "
                                        sql = sql & "Active,Cat,Seqno,Size,ChainType, "
                                        sql = sql & "Class,Tape1,Slider1,Body1,Finish1, "
                                        sql = sql & "Slider2,Body2,Finish2,Family,SpecialList, "
                                        sql = sql & "RDSpec,RDFinish,YOBI1,YOBI2,YOBI3, "
                                        sql = sql & "CreateUser, CreateTime, ModifyUser, ModifyTime "
                                        sql = sql & ") "
                                        sql = sql & "VALUES( "
                                        sql &= "'0' ,"
                                        sql &= "'" & "HFIN" & "' ,"
                                        sql &= "1, "
                                        sql &= "'" & UCase(DSize.Text.Trim) & "' ,"
                                        sql &= "'" & "" & "' ,"
                                        '
                                        sql &= "'" & "" & "' ,"
                                        sql &= "'" & "" & "' ,"
                                        sql &= "'" & DSlider1.Text & "' ,"
                                        sql &= "'" & "" & "' ,"
                                        sql &= "'" & DFinish1.Text & "' ,"
                                        '
                                        sql &= "'" & DSlider2.Text & "' ,"
                                        sql &= "'" & "" & "' ,"
                                        sql &= "'" & DFinish2.Text & "' ,"
                                        sql &= "'" & UCase(DFamily.Text.Trim) & "' ,"
                                        sql &= "'" & "** 外製:不適用此FINISH **" & "' ,"
                                        '
                                        sql &= "'', '', 'ERROR', '', '', "
                                        '
                                        sql &= "'" & Request.QueryString("pUserID") & "' ,"
                                        sql &= "'" & NowDateTime & "' ,"
                                        sql &= "'" & Request.QueryString("pUserID") & "' ,"
                                        sql &= "'" & NowDateTime & "' )"
                                        uDataBase.ExecuteNonQuery(sql)
                                        '
                                    End If
                                    '
                                Else
                                    uError = True
                                    msg = UCase(DSize.Text.Trim) & "|" & UCase(DFamily.Text.Trim) & "|" & "|" & UCase(xBody.Trim & xPuller.Trim) & "|" & UCase(iFinishList(i).Trim)
                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[NG-03]C3" & msg & Chr(13) & "** 外製:不適用此型別 **"
                                    '
                                    sql = "Insert into M_IRWPerformaceUp ( "
                                    sql = sql & "Active,Cat,Seqno,Size,ChainType, "
                                    sql = sql & "Class,Tape1,Slider1,Body1,Finish1, "
                                    sql = sql & "Slider2,Body2,Finish2,Family,SpecialList, "
                                    sql = sql & "RDSpec,RDFinish,YOBI1,YOBI2,YOBI3, "
                                    sql = sql & "CreateUser, CreateTime, ModifyUser, ModifyTime "
                                    sql = sql & ") "
                                    sql = sql & "VALUES( "
                                    sql &= "'0' ,"
                                    sql &= "'" & "HFIN" & "' ,"
                                    sql &= "1, "
                                    sql &= "'" & UCase(DSize.Text.Trim) & "' ,"
                                    sql &= "'" & "" & "' ,"
                                    '
                                    sql &= "'" & "" & "' ,"
                                    sql &= "'" & "" & "' ,"
                                    sql &= "'" & DSlider1.Text & "' ,"
                                    sql &= "'" & "" & "' ,"
                                    sql &= "'" & DFinish1.Text & "' ,"
                                    '
                                    sql &= "'" & DSlider2.Text & "' ,"
                                    sql &= "'" & "" & "' ,"
                                    sql &= "'" & DFinish2.Text & "' ,"
                                    sql &= "'" & UCase(DFamily.Text.Trim) & "' ,"
                                    sql &= "'" & "** 外製:不適用此型別 **" & "' ,"
                                    '
                                    sql &= "'', '', 'ERROR', '', '', "
                                    '
                                    sql &= "'" & Request.QueryString("pUserID") & "' ,"
                                    sql &= "'" & NowDateTime & "' ,"
                                    sql &= "'" & Request.QueryString("pUserID") & "' ,"
                                    sql &= "'" & NowDateTime & "' )"
                                    uDataBase.ExecuteNonQuery(sql)
                                    '
                                End If
                            End If
                            '
                        End If
                        '
                    Else
                        'xBody = "" AND xPuller = "" 
                        uError = True
                        msg = "slider=" & iSliderList(i) & "=[" & "BODY=" & xBody & "] [" & _
                                                                  "PULLER=" & xPuller & "] [" & _
                                                                  "COLOR=" & xColor & "] [" & _
                                                                  "FINISHL=" & iFinishList(i) & "]"
                        DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[NG-03]C4" & msg
                        '
                        sql = "Insert into M_IRWPerformaceUp ( "
                        sql = sql & "Active,Cat,Seqno,Size,ChainType, "
                        sql = sql & "Class,Tape1,Slider1,Body1,Finish1, "
                        sql = sql & "Slider2,Body2,Finish2,Family,SpecialList, "
                        sql = sql & "RDSpec,RDFinish,YOBI1,YOBI2,YOBI3, "
                        sql = sql & "CreateUser, CreateTime, ModifyUser, ModifyTime "
                        sql = sql & ") "
                        sql = sql & "VALUES( "
                        sql &= "'0' ,"
                        sql &= "'" & "HFIN" & "' ,"
                        sql &= "1, "
                        sql &= "'" & UCase(DSize.Text.Trim) & "' ,"
                        sql &= "'" & "" & "' ,"
                        '
                        sql &= "'" & "" & "' ,"
                        sql &= "'" & "" & "' ,"
                        sql &= "'" & DSlider1.Text & "' ,"
                        sql &= "'" & "" & "' ,"
                        sql &= "'" & DFinish1.Text & "' ,"
                        '
                        sql &= "'" & DSlider2.Text & "' ,"
                        sql &= "'" & "" & "' ,"
                        sql &= "'" & DFinish2.Text & "' ,"
                        sql &= "'" & UCase(DFamily.Text.Trim) & "' ,"
                        sql &= "'" & "" & "' ,"
                        '
                        sql &= "'', '', 'ERROR', '', '', "
                        '
                        sql &= "'" & Request.QueryString("pUserID") & "' ,"
                        sql &= "'" & NowDateTime & "' ,"
                        sql &= "'" & Request.QueryString("pUserID") & "' ,"
                        sql &= "'" & NowDateTime & "' )"
                        uDataBase.ExecuteNonQuery(sql)
                        '
                    End If
                    '
                Next
                '
                If UBound(iSliderList) > 0 Then
                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "表面處理-End" & "]"
                End If
                '
                If InStr(DLimitedItem.Text, "[NG-03]") > 0 Then
                    uJavaScript.PopMsg(Me, "FINISH不適用此型別胴體！ (無法申請)")
                End If
                '
            End If
            '
        End If
        ' End
        ' --------------------------------
        '**     4.檢針-241025
        ' Start
        If pCat = "SLIDERKENSIN" Then
            '
            If UCase(DST1.Text.Trim & DST4.Text.Trim) = "11" Or UCase(DST1.Text.Trim & DST4.Text.Trim) = "13" Then
                '
                ' Slider Information
                str = ""
                If DSlider1.Text <> "" Then str = DSlider1.Text.Trim & "!"
                If DSlider2.Text <> "" Then str = str & DSlider2.Text.Trim & "!"
                '無Sldier
                If str = "" Then str = "!!"
                '
                iSliderList = str.ToString.Split("!")
                'Finish Information
                str = ""
                If DFinish1.Text <> "" Then str = DFinish1.Text.Trim & "!"
                If DFinish2.Text <> "" Then str = str & DFinish2.Text.Trim & "!"
                '無Finish
                If str = "" Then str = "!!"
                '
                iFinishList = str.ToString.Split("!")
                '
                If UBound(iSliderList) > 0 Then
                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "檢針檢測-Start" & "]"
                End If
                '
                xKensinStr = ""
                xPullerStr = ""
                xALL = ""
                xSpecial = ""
                xRuleStr = ""
                For i = 0 To UBound(iSliderList) - 1
                    xSlider = ""
                    xBody = ""
                    xPuller = ""
                    xColor = ""
                    xSpecial = ""
                    '
                    ' Get Body
                    sql = "SELECT BODY, BodyGroup From V_EDXBodyList "
                    sql = sql & "where '" & iSliderList(i) & "' like body+'%' "
                    sql = sql & "ORDER BY LEN([BODY]) DESC, BODY "
                    Dim dtBODY1 As DataTable = uDataBase.GetDataTable(sql)
                    If dtBODY1.Rows.Count > 0 Then
                        xBody = dtBODY1.Rows(0)("BODY").ToString.Trim
                    End If
                    ' Get Special
                    sql = "SELECT ShortSpecial From V_SliderShortSpecial "
                    sql = sql & "where '" & iSliderList(i) & "' like '%' + ShortSpecial "
                    sql = sql & "ORDER BY LEN([ShortSpecial]) DESC "
                    Dim dtSpecial As DataTable = uDataBase.GetDataTable(sql)
                    If dtSpecial.Rows.Count > 0 Then
                        xSpecial = dtSpecial.Rows(0)("ShortSpecial").ToString.Trim
                    End If
                    ' Get Puller & Color
                    str = iSliderList(i)
                    If xBody <> "" Then str = Replace(str, xBody, "").Trim
                    If xSpecial <> "" Then str = Replace(str, xSpecial, "").Trim
                    j = Len(str)
                    Do Until j = 0
                        '
                        If Mid(str, j, 1) >= "0" And Mid(str, j, 1) <= "9" Then
                            '
                            If xPuller = "" Then
                                j = 1
                                Do Until j = Len(str) + 1
                                    If Mid(str, j, 1) >= "0" And Mid(str, j, 1) <= "Z" Then
                                        xPuller = Mid(str, 1, j).Trim
                                    Else
                                        Exit Do
                                    End If
                                    j = j + 1
                                Loop
                            End If
                            '
                            Exit Do
                        Else
                            xPuller = Mid(str, 1, j - 1).Trim
                        End If
                        j = j - 1
                    Loop
                    '
                    If str <> "" And InStr(str, xPuller) > 0 Then
                        xColor = Replace(str, xPuller, "").Trim
                    End If
                    '
                    If xPuller = "" And xColor <> "" Then
                        xPuller = xColor
                        xColor = ""
                    End If
                    '
                    '是縮寫 PULLER = 還原原 PULLER
                    '
                    If xPuller <> "" Then
                        sql = "SELECT TOP 1 Buyer, Puller, Short "
                        sql = sql & "From V_PullerShort "
                        sql = sql & "Where Short like '%" & "ZZZZZZZ" & "%' "
                        sql = sql & "Or Short like '%/" & xPuller & "/%' "
                        sql = sql & "Order By Buyer, Puller, Short "
                        Dim dtPuller1 As DataTable = uDataBase.GetDataTable(sql)
                        If dtPuller1.Rows.Count > 0 Then
                            xPuller = dtPuller1.Rows(0)("Puller").ToString.Trim
                        End If
                    End If
                    '
                    xSpecial = UCase(DSRequest1.Text.Trim) & "/" & UCase(DSRequest2.Text.Trim) & "/" & _
                               UCase(DSRequest3.Text.Trim) & "/" & UCase(DSRequest4.Text.Trim) & "/" & _
                               UCase(DSRequest5.Text.Trim) & "/" & UCase(DSRequest6.Text.Trim) & "/"
                    xSpecial = Replace(xSpecial, "XXX", "")
                    '
                    'MsgBox(xALL)
                    '
                    ' CHECK KENSIN MASTER
                    sql = "SELECT *, RDSpec As KENSIN, RDFinish As NANTI, "
                    sql = sql & "LTRIM(RTRIM(STR(Unique_ID)))+'/'+Size+'/'+ChainType+'/'+Slider1+'/'+Finish1+'/'+Slider2+'/'+Finish2+'/'+Family+'/'+SpecialLIST AS KENSINKEY "
                    sql = sql & "From M_IRWPerformaceUp "
                    sql = sql & "Where CAT IN ('KENSIN') "
                    sql = sql & "  And Active = 1 "
                    sql = sql & "  And '" & UCase(DSize.Text.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Size))='' THEN '%%' ELSE LTRIM(RTRIM(Size)) END) "
                    sql = sql & "  And '" & UCase(DChain.Text.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(ChainType))='' THEN '%%' ELSE LTRIM(RTRIM(ChainType)) END) "
                    '
                    '2024/10/29 MODIFY
                    sql = sql & "  And '" & UCase(iSliderList(i).Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider1))='' THEN '%%' ELSE LTRIM(RTRIM(Slider1)) END) "
                    'sql = sql & "  And '" & UCase(xBody.Trim & xPuller.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider1))='' THEN '%%' ELSE LTRIM(RTRIM(Slider1)) END) "
                    '
                    sql = sql & "  And '" & UCase(iFinishList(i).Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Finish1))='' THEN '%%' ELSE LTRIM(RTRIM(Finish1)) END) "
                    sql = sql & "  And '" & UCase(ReplaceFamily(DSize.Text, DFamily.Text)) & "' Like (CASE WHEN LTRIM(RTRIM(Family))='' THEN '%%' ELSE LTRIM(RTRIM(Family)) END) "
                    sql = sql & "  And '" & UCase(xSpecial.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(SpecialList))='' THEN '%%' ELSE LTRIM(RTRIM(SpecialList)) END) "
                    sql = sql & "Order By LEN(LTRIM(RTRIM(RDSpec)) + LTRIM(RTRIM(RDFinish)) ) "
                    '
                    Dim dtPerformaceUp01 As DataTable = uDataBase.GetDataTable(sql)
                    '
                    If dtPerformaceUp01.Rows.Count > 0 Then
                        '
                        xRuleStr = ""
                        For j = 0 To dtPerformaceUp01.Rows.Count - 1
                            '
                            xKensin = True
                            '
                            'Class
                            If dtPerformaceUp01.Rows(j)("Class").ToString.Trim() <> "" Then
                                '
                                If Left(dtPerformaceUp01.Rows(j)("Class").ToString.Trim(), 1) <> "!" Then
                                    If DClass.Text.Trim <> dtPerformaceUp01.Rows(j)("Class").ToString.Trim() Then xKensin = False
                                End If
                                If Left(dtPerformaceUp01.Rows(j)("Class").ToString.Trim(), 1) = "!" Then
                                    If DClass.Text.Trim = Mid(dtPerformaceUp01.Rows(j)("Class").ToString.Trim(), 2).Trim Then xKensin = False
                                End If
                                '
                            End If
                            '
                            'Slider2
                            If dtPerformaceUp01.Rows(j)("Slider2").ToString.Trim() <> "" Then
                                '
                                If dtPerformaceUp01.Rows(j)("Slider2").ToString.Trim() = "[SP]" Then
                                    If DSlider2.Text.Trim <> "" Then xKensin = False

                                Else
                                    '
                                    If dtPerformaceUp01.Rows(j)("Slider2").ToString.Trim() = "![SP]" Then
                                        If DSlider2.Text.Trim = "" Then xKensin = False
                                    Else
                                        '
                                        'CHECK Slider2
                                        If InStr(dtPerformaceUp01.Rows(j)("Slider2").ToString.Trim(), "%") > 0 Then
                                            '
                                            Select Case True
                                                Case Not DSlider2.Text.Trim Like Replace(dtPerformaceUp01.Rows(j)("Slider2").ToString, "%", "*")
                                                    xKensin = False
                                            End Select
                                        Else
                                            If DSlider2.Text.Trim <> dtPerformaceUp01.Rows(j)("Slider2").ToString.Trim Then
                                                xKensin = False
                                            End If
                                        End If
                                        '
                                    End If
                                    '
                                End If
                            End If
                            '
                            If xKensin = True Then
                                Select Case UCase(xKensinStr.Trim)
                                    Case ""
                                        xKensinStr = dtPerformaceUp01.Rows(j)("KENSIN").ToString.Trim() & "/"
                                    Case Else
                                        xKensinStr = xKensinStr & dtPerformaceUp01.Rows(j)("KENSIN").ToString.Trim() & "/"
                                End Select
                                '
                                Select Case UCase(xPullerStr.Trim)
                                    Case ""
                                        xPullerStr = dtPerformaceUp01.Rows(j)("NANTI").ToString.Trim() & "/"
                                    Case Else
                                        xPullerStr = xPullerStr & dtPerformaceUp01.Rows(j)("NANTI").ToString.Trim() & "/"
                                End Select
                                '
                                ' Error Unique_ID
                                If xRuleStr = "" Then
                                    xRuleStr = dtPerformaceUp01.Rows(j)("Unique_ID").ToString.Trim() & "/"
                                Else
                                    xRuleStr = xRuleStr & dtPerformaceUp01.Rows(j)("Unique_ID").ToString.Trim() & "/"
                                End If
                                '
                                ' TEST
                                xALL = xALL & dtPerformaceUp01.Rows(j)("KENSINKEY").ToString.Trim() & "*" & Chr(13)
                            End If
                        Next
                    End If
                    '
                    'MsgBox(xKensinStr & "--" & xPullerStr)
                Next
                '
                'MsgBox(xALL & Chr(13) & xKensinStr & " -- " & xPullerStr)
                '
                xKensinKey = ""
                xNantiKey = ""
                If xKensinStr <> "" Then
                    str = ""
                    iKensinList = xKensinStr.ToString.Split("/")
                    For i = 0 To UBound(iKensinList) - 1
                        If i = 0 Then str = iKensinList(i)
                        '
                        If iKensinList(i) < str Then
                            str = iKensinList(i)
                        Else
                            If Left(iKensinList(i), 2) = Left(str, 2) Then
                                If Len(iKensinList(i)) > Len(str) Then
                                    str = iKensinList(i)
                                End If
                            End If
                        End If
                    Next
                    xKensinKey = str
                    '
                    str = ""
                    iFinishList = xPullerStr.ToString.Split("/")
                    For i = 0 To UBound(iFinishList) - 1
                        If i = 0 Then str = iFinishList(i)
                        '
                        If iFinishList(i) < str Then
                            str = iFinishList(i)
                        Else
                            If Left(iFinishList(i), 2) = Left(str, 2) Then
                                If Len(iFinishList(i)) > Len(str) Then
                                    str = iFinishList(i)
                                End If
                            End If
                        End If
                    Next
                    xNantiKey = str
                    '
                    'MsgBox(xKensinKey & "---" & xNantiKey)
                    '
                    xKensin = True
                    Select Case UCase(xKensinKey.Trim)
                        Case "1.NONE"
                            If InStr(UCase(xSpecial.Trim), "KENSIN/") > 0 Or InStr(UCase(xSpecial.Trim), "ND-B/") > 0 Then xKensin = False
                        Case "2.ND-B"
                            If InStr(UCase(xSpecial.Trim), "ND-B/") <= 0 Then xKensin = False
                        Case "3.KENSIN"
                            If InStr(UCase(xSpecial.Trim), "KENSIN/") <= 0 Then xKensin = False
                        Case Else
                            str = xKensinKey
                            For i = 1 To 9
                                str = Replace(str, CStr(i).Trim & ".", "")
                            Next
                            If InStr(UCase(xSpecial.Trim), UCase(str.Trim) & "/") <= 0 Then xKensin = False
                    End Select
                    '
                    Select Case UCase(xNantiKey.Trim)
                        Case "1.NONE"
                            If InStr(UCase(xSpecial.Trim), "N-ANTI/") > 0 Then xKensin = False
                        Case "2.N-ANTI"
                            If InStr(UCase(xSpecial.Trim), "N-ANTI/") <= 0 Then xKensin = False
                        Case Else
                            str = xNantiKey
                            For i = 1 To 9
                                str = Replace(str, CStr(i).Trim & ".", "")
                            Next
                            If InStr(UCase(xSpecial.Trim), UCase(str.Trim) & "/") <= 0 Then xKensin = False
                    End Select
                    '
                    ' REPLACE SHOWMESSAGE xKensinKey, xNantiKey
                    sql = "SELECT Top 1 RDSpec As KENSIN, RDFinish As NANTI "
                    sql = sql & "From M_IRWPerformaceUp "
                    sql = sql & "Where CAT IN ('KENSIN') "
                    sql = sql & "  And Active = 1 "
                    sql = sql & "  And Seqno =  990 "
                    sql = sql & "  And '" & UCase(xKensinKey.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(SpecialList))='' THEN '%%' ELSE '%'+LTRIM(RTRIM(SpecialList))+'%' END) "
                    sql = sql & "Order By SIZE+'|'+ChainType+'|'+Tape1+'|'+SpecialList "
                    Dim dtPerformaceUpKensin3 As DataTable = uDataBase.GetDataTable(sql)
                    If dtPerformaceUpKensin3.Rows.Count > 0 Then
                        xKensinKey = dtPerformaceUpKensin3.Rows(0)("KENSIN").ToString.Trim
                        xNantiKey = dtPerformaceUpKensin3.Rows(0)("NANTI").ToString.Trim
                    End If
                    '
                    '**申請=有檢針 / WINGS=有檢針
                    If xKensin = True Then
                        'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                        '                    "[OK-04]A1" & "[WINGS:" & "檢針[" & UCase(xKensinKey.Trim) & "]" & "低鎳[" & UCase(xNantiKey.Trim) & "]"
                    End If
                Else
                    xKensin = False
                End If
                '
                'Error & HKENSIN
                If xKensin = False Then
                    '
                    sql = "SELECT Top 1 SIZE+'|'+SLIDER1+'|'+FINISH1+'|'+SpecialList as KENSINREADY "
                    sql = sql & "From M_IRWPerformaceUp "
                    sql = sql & "Where CAT IN ('HKENSIN') "
                    sql = sql & "  And Active = 1 "
                    sql = sql & "  And '" & UCase(DSize.Text.Trim) & "' Like ( case when LTRIM(RTRIM([SIZE]))='' then '" & UCase(DSize.Text.Trim) & "' else LTRIM(RTRIM([SIZE])) end ) "
                    sql = sql & "  And '" & UCase(DChain.Text.Trim) & "' Like ( case when LTRIM(RTRIM([ChainType]))='' then '" & UCase(DChain.Text.Trim) & "' else LTRIM(RTRIM([ChainType])) end ) "
                    sql = sql & "  And '" & UCase(DClass.Text.Trim) & "' Like ( case when LTRIM(RTRIM([Class]))='' then '" & UCase(DClass.Text.Trim) & "' else LTRIM(RTRIM([Class])) end ) "
                    sql = sql & "  And '" & UCase(DTape.Text.Trim) & "' Like ( case when LTRIM(RTRIM([Tape1]))='' then '" & UCase(DTape.Text.Trim) & "' else LTRIM(RTRIM([Tape1])) end ) "
                    sql = sql & "  And '" & UCase(DSlider1.Text.Trim) & "' Like ( case when LTRIM(RTRIM([Slider1]))='' then '" & UCase(DSlider1.Text.Trim) & "' else LTRIM(RTRIM([Slider1])) end ) "
                    sql = sql & "  And '" & UCase(DFinish1.Text.Trim) & "' Like ( case when LTRIM(RTRIM([FINISH1]))='' then '" & UCase(DFinish1.Text.Trim) & "' else LTRIM(RTRIM([FINISH1])) end ) "
                    sql = sql & "  And '" & UCase(DSlider2.Text.Trim) & "' Like ( case when LTRIM(RTRIM([Slider1]))='' then '" & UCase(DSlider2.Text.Trim) & "' else LTRIM(RTRIM([Slider2])) end ) "
                    sql = sql & "  And '" & UCase(DFinish2.Text.Trim) & "' Like ( case when LTRIM(RTRIM([FINISH1]))='' then '" & UCase(DFinish2.Text.Trim) & "' else LTRIM(RTRIM([FINISH2])) end ) "
                    sql = sql & "  And '" & UCase(DFamily.Text.Trim) & "' Like ( case when LTRIM(RTRIM([Family]))='' then '" & UCase(DFamily.Text.Trim) & "' else LTRIM(RTRIM([Family])) end ) "
                    sql = sql & "  And '" & UCase(xSpecial.Trim) & "' Like ( case when LTRIM(RTRIM([SpecialList]))='' then '" & UCase(xSpecial.Trim) & "' else LTRIM(RTRIM([SpecialList])) end ) "
                    sql = sql & "  And Seqno <> 999 "
                    sql = sql & "Order By SIZE+'|'+SLIDER1+'|'+FINISH1+'|'+SpecialList "
                    Dim dtPerformaceUpKensin1 As DataTable = uDataBase.GetDataTable(sql)
                    If dtPerformaceUpKensin1.Rows.Count > 0 Then
                        '**申請=有檢針 / WINGS=有檢針
                        'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                        '                    "[OK-04]A2" & "[WINGS:" & dtPerformaceUpKensin1.Rows(0)("KENSINREADY").ToString.Trim & "]"
                    Else
                        '**KENSIN-999-START
                        xALL = DSize.Text & "/" & DChain.Text & "/" & _
                               DClass.Text & "/" & DTape.Text & "/" & _
                               DSlider1.Text & "/" & DFinish1.Text & "/" & _
                               DSlider2.Text & "/" & DFinish2.Text & "/" & DFamily.Text & "/" & xSpecial
                        '
                        'MsgBox(xALL)
                        '
                        sql = "SELECT Top 1 SIZE+'|'+ChainType+'|'+Tape1+'|'+SpecialList as KENSINREADY "
                        sql = sql & "From M_IRWPerformaceUp "
                        sql = sql & "Where CAT IN ('HKENSIN') "
                        sql = sql & "  And Active = 1 "
                        sql = sql & "  And Seqno =  999 "
                        sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Size))='' THEN '%%' ELSE LTRIM(RTRIM(Size)) END) "
                        sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(ChainType))='' THEN '%%' ELSE LTRIM(RTRIM(ChainType)) END) "
                        sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Class))='' THEN '%%' ELSE LTRIM(RTRIM(Class)) END) "
                        sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Tape1))='' THEN '%%' ELSE LTRIM(RTRIM(Tape1)) END) "
                        sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider1))='' THEN '%%' ELSE LTRIM(RTRIM(Slider1)) END) "
                        sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Body1))='' THEN '%%' ELSE LTRIM(RTRIM(Body1)) END) "
                        sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Finish1))='' THEN '%%' ELSE LTRIM(RTRIM(Finish1)) END) "
                        sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider2))='' THEN '%%' ELSE LTRIM(RTRIM(Slider2)) END) "
                        sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Body2))='' THEN '%%' ELSE LTRIM(RTRIM(Body2)) END) "
                        sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Finish2))='' THEN '%%' ELSE LTRIM(RTRIM(Finish2)) END) "
                        sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Family))='' THEN '%%' ELSE LTRIM(RTRIM(Family)) END) "
                        sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(SpecialList))='' THEN '%%' ELSE LTRIM(RTRIM(SpecialList)) END) "
                        '
                        sql = sql & "Order By SIZE+'|'+ChainType+'|'+Tape1+'|'+SpecialList "
                        Dim dtPerformaceUpKensin2 As DataTable = uDataBase.GetDataTable(sql)
                        If dtPerformaceUpKensin2.Rows.Count > 0 Then
                            'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                            '                        "[OK-04]A3" & "[WINGS:" & dtPerformaceUpKensin2.Rows(0)("KENSINREADY").ToString.Trim & "]"
                        Else
                            '
                            uError = True
                            msg = "檢針[" & UCase(xKensinKey.Trim) & "]" & "低鎳[" & UCase(xNantiKey.Trim) & "][" & _
                                  "SIZE=" & UCase(DSize.Text.Trim) & "] [" & "CHAINTYPE=" & UCase(DChain.Text.Trim) & "] [" & _
                                  "SLIDER1=" & UCase(DSlider1.Text.Trim) & "] [" & "FINISH1=" & UCase(DFinish1.Text.Trim) & "] [" & _
                                  "SLIDER2=" & UCase(DSlider2.Text.Trim) & "] [" & "FINISH2=" & UCase(DFinish2.Text.Trim) & "] [" & "SPECIAL=" & xSpecial & "]"
                            DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[NG-04]A1" & msg
                        End If
                        'KENSIN-999-END
                    End If
                    '
                End If
                '
                'ERROR & HKENSIN DATA
                If InStr(DLimitedItem.Text, "[NG-04]") > 0 Then
                    '
                    ' Error Unique_ID
                    ' ERROR & HKENSIN DATA(seqno=910) --> CHECK
                    If xRuleStr <> "" Then
                        '**KENSIN-999-START
                        xALL = DSize.Text & "/" & DChain.Text & "/" & _
                               DClass.Text & "/" & DTape.Text & "/" & _
                               DSlider1.Text & "/" & DFinish1.Text & "/" & _
                               DSlider2.Text & "/" & DFinish2.Text & "/" & DFamily.Text & "/" & xSpecial
                        '
                        iRuleList = xRuleStr.ToString.Split("/")
                        For j = 0 To UBound(iRuleList) - 1
                            '
                            sql = "Insert into M_IRWPerformaceUp "
                            sql = sql & "Select "
                            sql = sql & "0, 'HKENSIN', 910, Size, ChainType, "
                            sql = sql & "Class,Tape1,Slider1,Body1,Finish1, "
                            sql = sql & "Slider2,Body2,Finish2,Family,SpecialList, "
                            sql = sql & "RDSpec,RDFinish, "
                            sql = sql & "'" & xALL.Trim() & "' ,"
                            sql = sql & "YOBI2,YOBI3, "
                            sql = sql & "'" & Request.QueryString("pUserID") & "' ,"
                            sql = sql & "'" & NowDateTime & "' ,"
                            sql = sql & "'" & Request.QueryString("pUserID") & "' ,"
                            sql = sql & "'" & NowDateTime & "' "
                            sql = sql & "From M_IRWPerformaceUp "
                            sql = sql & "Where Unique_ID = " & iRuleList(j).Trim() & " "
                            '
                            uDataBase.ExecuteNonQuery(sql)
                            '
                        Next
                    End If
                    '
                    sql = "Insert into M_IRWPerformaceUp ( "
                    sql = sql & "Active,Cat,Seqno,Size,ChainType, "
                    sql = sql & "Class,Tape1,Slider1,Body1,Finish1, "
                    sql = sql & "Slider2,Body2,Finish2,Family,SpecialList, "
                    sql = sql & "RDSpec,RDFinish,YOBI1,YOBI2,YOBI3, "
                    sql = sql & "CreateUser, CreateTime, ModifyUser, ModifyTime "
                    sql = sql & ") "
                    sql = sql & "VALUES( "
                    sql &= "'0' ,"
                    sql &= "'" & "HKENSIN" & "' ,"
                    sql &= "1, "
                    sql &= "'" & UCase(DSize.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DChain.Text.Trim) & "' ,"
                    '
                    sql &= "'" & UCase(DClass.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DTape.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DSlider1.Text.Trim) & "' ,"
                    sql &= "'" & "" & "' ,"
                    sql &= "'" & UCase(DFinish1.Text.Trim) & "' ,"
                    '
                    sql &= "'" & UCase(DSlider2.Text.Trim) & "' ,"
                    sql &= "'" & "" & "' ,"
                    sql &= "'" & UCase(DFinish2.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DFamily.Text.Trim) & "' ,"
                    sql &= "'" & xSpecial & "' ,"
                    '
                    sql &= "'', '', 'ERROR', '', '', "
                    '
                    sql &= "'" & Request.QueryString("pUserID") & "' ,"
                    sql &= "'" & NowDateTime & "' ,"
                    sql &= "'" & Request.QueryString("pUserID") & "' ,"
                    sql &= "'" & NowDateTime & "' )"
                    uDataBase.ExecuteNonQuery(sql)
                    '
                End If
                '
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "檢針檢測-End" & "]"
                '
                If InStr(DLimitedItem.Text, "NG-04") > 0 Then
                    uJavaScript.PopMsg(Me, "[Slider Code]檢針是否正確？")
                End If
                '
            End If
            '
        End If
        'End
        ' --------------------------------
        '**     5.P-PACK
        ' Start
        If pCat = "PPACK" Then
            '
            If DSlider1.Text.Trim & DSlider2.Text.Trim <> "" Then
                ' Slider Information
                str = ""
                If DSlider1.Text <> "" Then str = DSlider1.Text.Trim & "!"
                If DSlider2.Text <> "" Then str = str & DSlider2.Text.Trim & "!"
                iSliderList = str.ToString.Split("!")
                'Finish Information
                str = ""
                If DFinish1.Text <> "" Then str = DFinish1.Text.Trim & "!"
                If DFinish2.Text <> "" Then str = str & DFinish2.Text.Trim & "!"
                iFinishList = str.ToString.Split("!")
                '
                If UBound(iSliderList) > 0 Then
                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "P-PACK檢測-Start" & "]"
                End If
                '
                For i = 0 To UBound(iSliderList) - 1
                    xSlider = ""
                    xBody = ""
                    xPuller = ""
                    xColor = ""
                    xSpecial = ""
                    ' Get Body
                    sql = "SELECT BODY, BodyGroup From V_EDXBodyList "
                    sql = sql & "where '" & iSliderList(i) & "' like body+'%' "
                    sql = sql & "ORDER BY LEN([BODY]) DESC, BODY "
                    Dim dtBODY1 As DataTable = uDataBase.GetDataTable(sql)
                    If dtBODY1.Rows.Count > 0 Then
                        xBody = dtBODY1.Rows(0)("BODY").ToString.Trim
                    End If
                    ' Get Special
                    sql = "SELECT ShortSpecial From V_SliderShortSpecial "
                    sql = sql & "where '" & iSliderList(i) & "' like '%' + ShortSpecial "
                    sql = sql & "ORDER BY LEN([ShortSpecial]) DESC "
                    Dim dtSpecial As DataTable = uDataBase.GetDataTable(sql)
                    If dtSpecial.Rows.Count > 0 Then
                        xSpecial = dtSpecial.Rows(0)("ShortSpecial").ToString.Trim
                    End If
                    ' Get Puller & Color
                    str = iSliderList(i)
                    If xBody <> "" Then str = Replace(str, xBody, "").Trim
                    If xSpecial <> "" Then str = Replace(str, xSpecial, "").Trim
                    j = Len(str)
                    Do Until j = 0
                        If Mid(str, j, 1) >= "0" And Mid(str, j, 1) <= "9" Then
                            '
                            If xPuller = "" Then
                                j = 1
                                Do Until j = Len(str) + 1
                                    If Mid(str, j, 1) >= "0" And Mid(str, j, 1) <= "Z" Then
                                        xPuller = Mid(str, 1, j).Trim
                                    Else
                                        Exit Do
                                    End If
                                    j = j + 1
                                Loop
                            End If
                            '
                            Exit Do
                        Else
                            xPuller = Mid(str, 1, j - 1).Trim
                        End If
                        j = j - 1
                    Loop
                    '
                    If str <> "" And InStr(str, xPuller) > 0 Then
                        xColor = Replace(str, xPuller, "").Trim
                    End If
                    xSlider = xBody & xPuller
                    '
                    'MsgBox(xSlider & "--" & xBody & "--" & xPuller & "--" & xColor)
                    '
                    xSpecial = UCase(DSRequest1.Text.Trim) & "/" & UCase(DSRequest2.Text.Trim) & "/" & _
                               UCase(DSRequest3.Text.Trim) & "/" & UCase(DSRequest4.Text.Trim) & "/" & _
                               UCase(DSRequest5.Text.Trim) & "/" & UCase(DSRequest6.Text.Trim) & "/"
                    xSpecial = Replace(xSpecial, "XXX", "")
                    '
                    '81-41
                    sql = "SELECT Top 1 SLIDER1+'|'+FINISH1+'|'+SpecialList as PPACK, SpecialList As PASS "
                    sql = sql & "From M_IRWPerformaceUp "
                    sql = sql & "Where CAT IN ('WPPACK', 'HPPACK') "
                    sql = sql & "  And Active = 1 "
                    sql = sql & "  And LTRIM(RTRIM(SLIDER1)) = '" & UCase(iSliderList(i).Trim) & "' "
                    sql = sql & "  And LTRIM(RTRIM(FINISH1)) = '" & UCase(iFinishList(i).Trim) & "' "
                    sql = sql & "Order By CAT, SLIDER1+'|'+FINISH1+'|'+SpecialList "
                    Dim dtPerformaceUpPPack As DataTable = uDataBase.GetDataTable(sql)
                    If dtPerformaceUpPPack.Rows.Count > 0 Then
                        '
                        '**P-PACK
                        If InStr(xSpecial, "P-PACK/") > 0 Or InStr(dtPerformaceUpPPack.Rows(0)("PASS").ToString.Trim, "OK") > 0 Then
                            'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                            '                        "[OK-05]" & "[WINGS:" & dtPerformaceUpPPack.Rows(0)("PPACK").ToString.Trim & "]"
                        Else
                            uError = True
                            msg = "slider=" & iSliderList(i) & "=[" & "SLIDER=" & UCase(iSliderList(i).Trim) & "] [" & _
                                                                      "FINISH=" & UCase(iFinishList(i).Trim) & "] [" & _
                                                                      "SPECIAL=" & xSpecial & "]"
                            DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[NG-05]" & msg
                        End If
                    End If
                    '
                    If InStr(DLimitedItem.Text, "[NG-05]") > 0 Then
                        '
                        sql = "Insert into M_IRWPerformaceUp ( "
                        sql = sql & "Active,Cat,Seqno,Size,ChainType, "
                        sql = sql & "Class,Tape1,Slider1,Body1,Finish1, "
                        sql = sql & "Slider2,Body2,Finish2,Family,SpecialList, "
                        sql = sql & "RDSpec,RDFinish,YOBI1,YOBI2,YOBI3, "
                        sql = sql & "CreateUser, CreateTime, ModifyUser, ModifyTime "
                        sql = sql & ") "
                        sql = sql & "VALUES( "
                        sql &= "'0' ,"
                        sql &= "'" & "HPPACK" & "' ,"
                        sql &= "1, "
                        sql &= "'" & UCase(DSize.Text.Trim) & "' ,"
                        sql &= "'" & UCase(DChain.Text.Trim) & "' ,"
                        '
                        sql &= "'" & UCase(DClass.Text.Trim) & "' ,"
                        sql &= "'" & UCase(DTape.Text.Trim) & "' ,"
                        sql &= "'" & UCase(iSliderList(i).Trim) & "' ,"
                        sql &= "'" & "" & "' ,"
                        sql &= "'" & UCase(iFinishList(i).Trim) & "' ,"
                        '
                        sql &= "'" & "" & "' ,"
                        sql &= "'" & "" & "' ,"
                        sql &= "'" & "" & "' ,"
                        sql &= "'" & UCase(DFamily.Text.Trim) & "' ,"
                        sql &= "'" & xSpecial & "' ,"
                        '
                        sql &= "'', '', 'ERROR', '', '', "
                        '
                        sql &= "'" & Request.QueryString("pUserID") & "' ,"
                        sql &= "'" & NowDateTime & "' ,"
                        sql &= "'" & Request.QueryString("pUserID") & "' ,"
                        sql &= "'" & NowDateTime & "' )"
                        uDataBase.ExecuteNonQuery(sql)
                        '
                    End If
                    '
                Next
                '
                If UBound(iSliderList) > 0 Then
                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "P-PACK檢測-End" & "]"
                End If
                '
                If InStr(DLimitedItem.Text, "NG-05") > 0 Then
                    uJavaScript.PopMsg(Me, "[Special]特別要求(P-PACK)是否正確？ (無法申請)")
                End If
                '
            End If
            '
        End If
        'End
        ' --------------------------------
        '**     6.TWIST4
        ' Start
        If pCat = "TWIST4" Then
            '
            If DSlider1.Text.Trim & DSlider2.Text.Trim <> "" Then
                '
                xSlider = ""
                xBody = ""
                xPuller = ""
                xColor = ""
                xSpecial = ""
                '
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "TWIST4檢測-Start" & "]"
                '
                xSpecial = UCase(DSRequest1.Text.Trim) & "/" & UCase(DSRequest2.Text.Trim) & "/" & _
                           UCase(DSRequest3.Text.Trim) & "/" & UCase(DSRequest4.Text.Trim) & "/" & _
                           UCase(DSRequest5.Text.Trim) & "/" & UCase(DSRequest6.Text.Trim) & "/"
                xSpecial = Replace(xSpecial, "XXX", "")
                '
                '81-41
                sql = "SELECT Top 1 SIZE+'|'+SLIDER1+'|'+SpecialList as TWIST4, SpecialList As PASS "
                sql = sql & "From M_IRWPerformaceUp "
                sql = sql & "Where CAT IN ('WTWIST4', 'HTWIST4') "
                sql = sql & "  And Active = 1 "
                sql = sql & "  And LTRIM(RTRIM(SIZE)) = '" & UCase(DSize.Text.Trim) & "' "
                sql = sql & "  And LTRIM(RTRIM(SLIDER1)) = '" & UCase(DSlider1.Text.Trim) & "' "
                sql = sql & "  And LTRIM(RTRIM(SLIDER2)) = '" & UCase(DSlider2.Text.Trim) & "' "
                sql = sql & "Order By CAT, SIZE+'|'+SLIDER1+'|'+SpecialList "
                Dim dtPerformaceUpTWIST4 As DataTable = uDataBase.GetDataTable(sql)
                If dtPerformaceUpTWIST4.Rows.Count > 0 Then
                    '
                    'MsgBox(dtPerformaceUpTWIST4.Rows(0)("TWIST4").ToString.Trim)
                    '**TWIST4
                    If InStr(xSpecial, "TWIST4/") > 0 Or InStr(dtPerformaceUpTWIST4.Rows(0)("PASS").ToString.Trim, "OK") > 0 Then
                        'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                        '                        "[OK-06]F1" & "[WINGS:" & dtPerformaceUpTWIST4.Rows(0)("TWIST4").ToString.Trim & "]"
                    Else
                        uError = True
                        msg = "[" & "SIZE=" & UCase(DSize.Text.Trim) & "] [" & _
                                    "SLIDER1=" & UCase(DSlider1.Text.Trim) & "] [" & _
                                    "SLIDER2=" & UCase(DSlider2.Text.Trim) & "] [" & _
                                    "SPECIAL=" & xSpecial & "]"
                        DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[NG-06]" & msg

                        'sql = "SELECT Top 1 SIZE+'|'+SLIDER1+'|'+SpecialList as TWIST4 "
                        'sql = sql & "From M_IRWPerformaceUp "
                        'sql = sql & "Where CAT IN ('RDTWIST4') "
                        'sql = sql & "  And Active = 1 "
                        'sql = sql & "  And LTRIM(RTRIM(SIZE)) = '" & UCase(DSize.Text.Trim) & "' "
                        'sql = sql & "  And '" & UCase(iSliderList(i).Trim) & "' Like '%'+LTRIM(RTRIM([Slider1]))+'%' "
                        'sql = sql & "Order By SIZE+'|'+SLIDER1+'|'+SpecialList "
                        'Dim dtPerformaceUpRDTWIST4 As DataTable = uDataBase.GetDataTable(sql)
                        'If dtPerformaceUpRDTWIST4.Rows.Count > 0 Then
                        '    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                        '                            "[OK-06]F2" & "[RD:" & dtPerformaceUpRDTWIST4.Rows(0)("TWIST4").ToString.Trim & "]"
                        'Else
                        '    msg = "slider=" & iSliderList(i) & "=[" & "SIZE=" & UCase(DSize.Text.Trim) & "] [" & _
                        '                                              "SLIDER=" & UCase(iSliderList(i).Trim) & "] [" & _
                        '                                              "FINISH=" & UCase(iFinishList(i).Trim) & "] [" & _
                        '                                              "SPECIAL=" & xSpecial & "]"
                        '    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[NG-06]" & msg
                        'End If
                    End If
                End If
                '
                If InStr(DLimitedItem.Text, "[NG-06]") > 0 Then
                    '
                    sql = "Insert into M_IRWPerformaceUp ( "
                    sql = sql & "Active,Cat,Seqno,Size,ChainType, "
                    sql = sql & "Class,Tape1,Slider1,Body1,Finish1, "
                    sql = sql & "Slider2,Body2,Finish2,Family,SpecialList, "
                    sql = sql & "RDSpec,RDFinish,YOBI1,YOBI2,YOBI3, "
                    sql = sql & "CreateUser, CreateTime, ModifyUser, ModifyTime "
                    sql = sql & ") "
                    sql = sql & "VALUES( "
                    sql &= "'0' ,"
                    sql &= "'" & "HTWIST4" & "' ,"
                    sql &= "1, "
                    sql &= "'" & UCase(DSize.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DChain.Text.Trim) & "' ,"
                    '
                    sql &= "'" & UCase(DClass.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DTape.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DSlider1.Text.Trim) & "' ,"
                    sql &= "'" & "" & "' ,"
                    sql &= "'" & UCase(DFinish1.Text.Trim) & "' ,"
                    '
                    sql &= "'" & UCase(DSlider2.Text.Trim) & "' ,"
                    sql &= "'" & "" & "' ,"
                    sql &= "'" & UCase(DFinish2.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DFamily.Text.Trim) & "' ,"
                    sql &= "'" & xSpecial & "' ,"
                    '
                    sql &= "'', '', 'ERROR', '', '', "
                    '
                    sql &= "'" & Request.QueryString("pUserID") & "' ,"
                    sql &= "'" & NowDateTime & "' ,"
                    sql &= "'" & Request.QueryString("pUserID") & "' ,"
                    sql &= "'" & NowDateTime & "' )"
                    uDataBase.ExecuteNonQuery(sql)
                    '
                End If
                '
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "TWIST4檢測-End" & "]"
                '
                If InStr(DLimitedItem.Text, "NG-06") > 0 Then
                    uJavaScript.PopMsg(Me, "[Special]特別要求(TWIST4)是否正確？ (無法申請)")
                End If
                '
            End If
            '
        End If
        'End
        ' --------------------------------
        '**     7.SCD
        ' Start
        If pCat = "SCD" Then
            '
            '	ST1	ST2	ST4
            '	1	1	2
            '	1	2	1
            '	1	2	2
            '	1	3	1
            '	1	3	2
            If UCase(DST1.Text.Trim) & UCase(DST2.Text.Trim) & UCase(DST4.Text.Trim) = "112" Or _
               UCase(DST1.Text.Trim) & UCase(DST2.Text.Trim) & UCase(DST4.Text.Trim) = "121" Or _
               UCase(DST1.Text.Trim) & UCase(DST2.Text.Trim) & UCase(DST4.Text.Trim) = "122" Or _
               UCase(DST1.Text.Trim) & UCase(DST2.Text.Trim) & UCase(DST4.Text.Trim) = "131" Or _
               UCase(DST1.Text.Trim) & UCase(DST2.Text.Trim) & UCase(DST4.Text.Trim) = "132" Then
                '
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "特殊CH檢測-Start" & "]"
                '
                ' 特殊CH找出
                xSpecial = UCase(DSRequest1.Text.Trim) & "/" & UCase(DSRequest2.Text.Trim) & "/" & _
                           UCase(DSRequest3.Text.Trim) & "/" & UCase(DSRequest4.Text.Trim) & "/" & _
                           UCase(DSRequest5.Text.Trim) & "/" & UCase(DSRequest6.Text.Trim) & "/"
                xSpecial = Replace(xSpecial, "XXX", "")
                '
                xSlider = ""
                'sql = "SELECT TOP 1 SpecChain1+'/'+SpecChain2+'/'+SpecChain3+'/'+SpecChain4+'/'+SpecChain5+'/'+SpecChain6+'/' AS SCD "
                sql = "SELECT TOP 1 SpecChain1, SpecChain2, SpecChain3, SpecChain4, SpecChain5, SpecChain6, "
                sql = sql & "SpecChain1+'/'+SpecChain2+'/'+SpecChain3+'/'+SpecChain4+'/'+SpecChain5+'/'+SpecChain6+'/' AS SCD "
                sql = sql & "From V_IRWPerformaceUp_SCDList_01 "
                sql = sql & "Where '" & xSpecial.Trim & "' LIKE '%' + SpecChain1 + '%' "
                sql = sql & "  And '" & xSpecial.Trim & "' LIKE '%' + SpecChain2 + '%' "
                sql = sql & "  And '" & xSpecial.Trim & "' LIKE '%' + SpecChain3 + '%' "
                sql = sql & "  And '" & xSpecial.Trim & "' LIKE '%' + SpecChain4 + '%' "
                sql = sql & "  And '" & xSpecial.Trim & "' LIKE '%' + SpecChain5 + '%' "
                sql = sql & "  And '" & xSpecial.Trim & "' LIKE '%' + SpecChain6 + '%' "
                sql = sql & "ORDER BY LEN(SpecChain1+'/'+SpecChain2+'/'+SpecChain3+'/'+SpecChain4+'/'+SpecChain5+'/'+SpecChain6+'/') DESC "
                'MsgBox(sql)
                Dim dtSCDList As DataTable = uDataBase.GetDataTable(sql)
                If dtSCDList.Rows.Count > 0 Then
                    '
                    'MsgBox(dtSCDList.Rows(0)("SCD").ToString.Trim)
                    '
                    If InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain1").ToString.Trim) > 0 And dtSCDList.Rows(0)("SpecChain1").ToString.Trim <> "" Then
                        '
                        str = Mid(xSpecial.Trim, InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain1").ToString.Trim), 99)
                        str = Mid(str, 1, InStr(str, "/") - 1)
                        '
                        If xSlider = "" Then
                            xSlider = str & "/"
                        Else
                            xSlider = xSlider & str & "/"
                        End If
                    End If
                    '
                    If InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain2").ToString.Trim) > 0 And dtSCDList.Rows(0)("SpecChain2").ToString.Trim <> "" Then
                        '
                        str = Mid(xSpecial.Trim, InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain2").ToString.Trim), 99)
                        str = Mid(str, 1, InStr(str, "/") - 1)
                        '
                        If xSlider = "" Then
                            xSlider = str & "/"
                        Else
                            xSlider = xSlider & str & "/"
                        End If
                    End If
                    '
                    If InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain3").ToString.Trim) > 0 And dtSCDList.Rows(0)("SpecChain3").ToString.Trim <> "" Then
                        '
                        str = Mid(xSpecial.Trim, InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain3").ToString.Trim), 99)
                        str = Mid(str, 1, InStr(str, "/") - 1)
                        '
                        If xSlider = "" Then
                            xSlider = str & "/"
                        Else
                            xSlider = xSlider & str & "/"
                        End If
                    End If
                    '
                    If InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain4").ToString.Trim) > 0 And dtSCDList.Rows(0)("SpecChain4").ToString.Trim <> "" Then
                        '
                        str = Mid(xSpecial.Trim, InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain4").ToString.Trim), 99)
                        str = Mid(str, 1, InStr(str, "/") - 1)
                        '
                        If xSlider = "" Then
                            xSlider = str & "/"
                        Else
                            xSlider = xSlider & str & "/"
                        End If
                    End If
                    '
                    If InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain5").ToString.Trim) > 0 And dtSCDList.Rows(0)("SpecChain5").ToString.Trim <> "" Then
                        '
                        str = Mid(xSpecial.Trim, InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain5").ToString.Trim), 99)
                        str = Mid(str, 1, InStr(str, "/") - 1)
                        '
                        If xSlider = "" Then
                            xSlider = str & "/"
                        Else
                            xSlider = xSlider & str & "/"
                        End If
                    End If
                    '
                    If InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain6").ToString.Trim) > 0 And dtSCDList.Rows(0)("SpecChain6").ToString.Trim <> "" Then
                        '
                        str = Mid(xSpecial.Trim, InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain6").ToString.Trim), 99)
                        str = Mid(str, 1, InStr(str, "/") - 1)
                        '
                        If xSlider = "" Then
                            xSlider = str & "/"
                        Else
                            xSlider = xSlider & str & "/"
                        End If
                    End If
                    '
                End If
                '
                'MsgBox(xSlider)
                '
                If xSlider <> "" Then
                    str = ""
                    iSliderList = xSlider.ToString.Split("/")
                    For j = 0 To UBound(iSliderList) - 1
                        If iSliderList(j).Trim <> "" Then
                            If str = "" Then
                                str = UCase(iSliderList(j).Trim) & "/"
                            Else
                                str = str & UCase(iSliderList(j).Trim) & "/"
                            End If
                        End If
                    Next
                    '
                    'MsgBox(str)
                    '
                    iSliderList = str.ToString.Split("/")
                    '
                    '81-41
                    sql = "SELECT Top 1 SIZE+'|'+ChainType+'|'+Tape1+'|'+SpecialList as SCD "
                    sql = sql & "From M_IRWPerformaceUp "
                    sql = sql & "Where CAT IN ('WSCD') "
                    sql = sql & "  And Active = 1 "
                    sql = sql & "  And LTRIM(RTRIM(SIZE)) = '" & UCase(DSize.Text.Trim) & "' "
                    sql = sql & "  And LTRIM(RTRIM(ChainType)) = '" & UCase(DChain.Text.Trim) & "' "
                    sql = sql & "  And LTRIM(RTRIM(Tape1)) = '" & UCase(DTape.Text.Trim) & "' "
                    '
                    For j = 0 To UBound(iSliderList) - 1
                        sql = sql & "  And LTRIM(RTRIM(SpecialList)) Like '%" & iSliderList(j).Trim & "%' "
                    Next
                    sql = sql & "Order By CAT, SIZE+'|'+ChainType+'|'+Tape1+'|'+SpecialList "
                    '
                    'MsgBox(sql)
                    '
                    Dim dtPerformaceUpSCD As DataTable = uDataBase.GetDataTable(sql)
                    If dtPerformaceUpSCD.Rows.Count > 0 Then
                        'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                        '                        "[OK-07]G1" & "[WINGS:" & dtPerformaceUpSCD.Rows(0)("SCD").ToString.Trim & "]"
                    Else
                        '
                        '**SCD
                        sql = "SELECT Top 1 SIZE+'|'+ChainType+'|'+Tape1+'|'+SpecialList as SCD "
                        sql = sql & "From M_IRWPerformaceUp "
                        sql = sql & "Where CAT IN ('HSCD') "
                        sql = sql & "  And Active = 1 "
                        sql = sql & "  And LTRIM(RTRIM(SIZE)) = '" & UCase(DSize.Text.Trim) & "' "
                        sql = sql & "  And LTRIM(RTRIM(ChainType)) = '" & UCase(DChain.Text.Trim) & "' "
                        sql = sql & "  And LTRIM(RTRIM(Tape1)) = '" & UCase(DTape.Text.Trim) & "' "
                        '
                        sql = sql & "  And ( "
                        sql = sql & "        ( "
                        For j = 0 To UBound(iSliderList) - 1
                            If j = 0 Then
                                sql = sql & "      LTRIM(RTRIM(SpecialList)) Like '%" & iSliderList(j).Trim & "%' "
                            Else
                                sql = sql & "  And LTRIM(RTRIM(SpecialList)) Like '%" & iSliderList(j).Trim & "%' "
                            End If
                        Next
                        sql = sql & "        ) "
                        sql = sql & "        Or "
                        sql = sql & "        ( LTRIM(RTRIM(SpecialList)) = 'OK/') "
                        sql = sql & "  ) "
                        '
                        sql = sql & "Order By SIZE+'|'+ChainType+'|'+Tape1+'|'+SpecialList "
                        Dim dtPerformaceUpSCDA As DataTable = uDataBase.GetDataTable(sql)
                        If dtPerformaceUpSCDA.Rows.Count > 0 Then
                            'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                            '                        "[OK-07]G2" & "[WINGS:" & dtPerformaceUpSCDA.Rows(0)("SCD").ToString.Trim & "]"
                        Else
                            If DST6.Text.Trim = "M" Then
                                'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                                '                        "[OK-07]G2" & "SIZE=" & UCase(DSize.Text.Trim) & "] [" & _
                                '                                      "CHAIN=" & UCase(DChain.Text.Trim) & "] [" & _
                                '                                      "TAPE=" & UCase(DTape.Text.Trim) & "] [" & _
                                '                                      "SPECIAL=" & xSpecial & "]" & _
                                '                                      "ST6=" & DST6.Text.Trim & "]"
                            Else
                                '**SCD-2 SEQNO=999
                                xALL = DSize.Text & "/" & DChain.Text & "/" & _
                                       DClass.Text & "/" & DTape.Text & "/" & _
                                       DSlider1.Text & "/" & DFinish1.Text & "/" & _
                                       DSlider2.Text & "/" & DFinish2.Text & "/" & DFamily.Text & "/" & xSpecial
                                '
                                'MsgBox(xALL)
                                '
                                sql = "SELECT Top 1 SIZE+'|'+ChainType+'|'+Tape1+'|'+SpecialList as SCD "
                                sql = sql & "From M_IRWPerformaceUp "
                                sql = sql & "Where CAT IN ('HSCD') "
                                sql = sql & "  And Active = 1 "
                                sql = sql & "  And Seqno =  999 "
                                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Size))='' THEN '%%' ELSE LTRIM(RTRIM(Size)) END) "
                                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(ChainType))='' THEN '%%' ELSE LTRIM(RTRIM(ChainType)) END) "
                                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Class))='' THEN '%%' ELSE LTRIM(RTRIM(Class)) END) "
                                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Tape1))='' THEN '%%' ELSE LTRIM(RTRIM(Tape1)) END) "
                                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider1))='' THEN '%%' ELSE LTRIM(RTRIM(Slider1)) END) "
                                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Body1))='' THEN '%%' ELSE LTRIM(RTRIM(Body1)) END) "
                                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Finish1))='' THEN '%%' ELSE LTRIM(RTRIM(Finish1)) END) "
                                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider2))='' THEN '%%' ELSE LTRIM(RTRIM(Slider2)) END) "
                                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Body2))='' THEN '%%' ELSE LTRIM(RTRIM(Body2)) END) "
                                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Finish2))='' THEN '%%' ELSE LTRIM(RTRIM(Finish2)) END) "
                                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Family))='' THEN '%%' ELSE LTRIM(RTRIM(Family)) END) "
                                sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(SpecialList))='' THEN '%%' ELSE LTRIM(RTRIM(SpecialList)) END) "
                                '
                                'MsgBox(sql)
                                sql = sql & "Order By SIZE+'|'+ChainType+'|'+Tape1+'|'+SpecialList "
                                Dim dtPerformaceUpSCDB As DataTable = uDataBase.GetDataTable(sql)
                                If dtPerformaceUpSCDB.Rows.Count > 0 Then
                                    'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                                    '                        "[OK-07]G3" & "[WINGS:" & dtPerformaceUpSCDB.Rows(0)("SCD").ToString.Trim & "]"
                                Else
                                    uError = True
                                    msg = "slider=" & iSliderList(i) & "=[" & "SIZE=" & UCase(DSize.Text.Trim) & "] [" & _
                                                                              "CHAIN=" & UCase(DChain.Text.Trim) & "] [" & _
                                                                              "TAPE=" & UCase(DTape.Text.Trim) & "] [" & _
                                                                              "SPECIAL=" & xSpecial & "]"
                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[NG-07]" & msg
                                End If
                            End If
                            '
                        End If
                    End If
                End If
                '
                If InStr(DLimitedItem.Text, "[NG-07]") > 0 Then
                    '
                    sql = "Insert into M_IRWPerformaceUp ( "
                    sql = sql & "Active,Cat,Seqno,Size,ChainType, "
                    sql = sql & "Class,Tape1,Slider1,Body1,Finish1, "
                    sql = sql & "Slider2,Body2,Finish2,Family,SpecialList, "
                    sql = sql & "RDSpec,RDFinish,YOBI1,YOBI2,YOBI3, "
                    sql = sql & "CreateUser, CreateTime, ModifyUser, ModifyTime "
                    sql = sql & ") "
                    sql = sql & "VALUES( "
                    sql &= "'0' ,"
                    sql &= "'" & "HSCD" & "' ,"
                    sql &= "1, "
                    sql &= "'" & UCase(DSize.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DChain.Text.Trim) & "' ,"
                    '
                    sql &= "'" & UCase(DClass.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DTape.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DSlider1.Text.Trim) & "' ,"
                    sql &= "'" & "" & "' ,"
                    sql &= "'" & UCase(DFinish1.Text.Trim) & "' ,"
                    '
                    sql &= "'" & "" & "' ,"
                    sql &= "'" & "" & "' ,"
                    sql &= "'" & "" & "' ,"
                    sql &= "'" & UCase(DFamily.Text.Trim) & "' ,"
                    sql &= "'" & xSpecial & "' ,"
                    '
                    sql &= "'', '', 'ERROR', '', '', "
                    '
                    sql &= "'" & Request.QueryString("pUserID") & "' ,"
                    sql &= "'" & NowDateTime & "' ,"
                    sql &= "'" & Request.QueryString("pUserID") & "' ,"
                    sql &= "'" & NowDateTime & "' )"
                    uDataBase.ExecuteNonQuery(sql)
                    '
                End If
                '
                If InStr(DLimitedItem.Text, "NG-07") > 0 Then
                    uJavaScript.PopMsg(Me, "[Special]特別要求(特殊CH)是否正確？ (無法申請)")
                End If
                '
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "特殊CH檢測-End" & "]"
                '
            End If
            '
        End If
        'End
        ' --------------------------------
        '**     8.特別要求縮寫 20241009
        ' Start
        If pCat = "SPECIALSHORT" Then
            '
            xSpecial = UCase(DSRequest1.Text.Trim) & "/" & UCase(DSRequest2.Text.Trim) & "/" & _
                       UCase(DSRequest3.Text.Trim) & "/" & UCase(DSRequest4.Text.Trim) & "/" & _
                       UCase(DSRequest5.Text.Trim) & "/" & UCase(DSRequest6.Text.Trim) & "/"
            '
            If InStr(UCase(xSpecial.Trim), "//") > 0 Then
                '
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "特別要求縮寫-Start" & "]"
                '
                'SHORT MASTER
                sql = "SELECT Top 1 Special+'|'+Special1+'|'+Special2+'|'+Special3+'|'+Special4+'|'+Special5+'|'+Special6 As SHORT "
                sql = sql & "From M_SpecialShort "
                sql = sql & "Where Active = 1 "
                sql = sql & "  And '" & UCase(xSpecial.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Special))='' THEN '%%' ELSE '%'+LTRIM(RTRIM(Special))+'/%' END) "
                sql = sql & "Order By SortNo, Special "
                Dim dtSpecialShort As DataTable = uDataBase.GetDataTable(sql)
                If dtSpecialShort.Rows.Count <= 0 Then
                    'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                    '                        "[OK-08]H1" & "[WINGS:" & UCase(xSpecial.Trim) & "]"
                    '
                Else
                    uError = True
                    msg = "[SpecialShort" & "=" & dtSpecialShort.Rows(0)("SHORT").ToString.Trim() & "]"
                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[NG-08]" & msg
                    '
                End If
                '
                If InStr(DLimitedItem.Text, "[NG-08]") > 0 Then
                    '
                End If
                '
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "特別要求縮寫-End" & "]"
                '
                If InStr(DLimitedItem.Text, "NG-08") > 0 Then
                    uJavaScript.PopMsg(Me, "[SpecialShort]特別要求縮寫是否正確？ (無法申請)")
                End If
                '
            End If
            '
        End If
        'End
        ' --------------------------------
        '**     9.型別 20241109
        ' Start
        If pCat = "SLIDERSPEC" Then
            '
            If UCase(DST1.Text.Trim & DST4.Text.Trim) = "11" Or UCase(DST1.Text.Trim & DST4.Text.Trim) = "13" Then
                '
                ' Slider Information
                str = ""
                If DSlider1.Text <> "" Then str = DSlider1.Text.Trim & "!"
                If DSlider2.Text <> "" Then str = str & DSlider2.Text.Trim & "!"
                iSliderList = str.ToString.Split("!")
                '
                If UBound(iSliderList) > 0 Then
                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "型別檢測-Start" & "]"
                End If
                '
                For i = 0 To UBound(iSliderList) - 1
                    xSpecial = ""
                    xSpecial = UCase(DSRequest1.Text.Trim) & "/" & UCase(DSRequest2.Text.Trim) & "/" & _
                               UCase(DSRequest3.Text.Trim) & "/" & UCase(DSRequest4.Text.Trim) & "/" & _
                               UCase(DSRequest5.Text.Trim) & "/" & UCase(DSRequest6.Text.Trim) & "/"
                    xSpecial = Replace(xSpecial, "XXX", "")
                    '
                    '81-41
                    sql = "SELECT SIZE + '|' + FAMILY + '|' + SLIDER1 As SLIDERINF "
                    sql = sql & "From M_IRWPerformaceUp "
                    sql = sql & "Where CAT IN ('KENSIN') "
                    sql = sql & "  And Active = 1 "
                    sql = sql & "  And Seqno in (4,920) "
                    sql = sql & "  And '" & UCase(DSize.Text.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Size))='' THEN '%%' ELSE LTRIM(RTRIM(Size)) END) "
                    sql = sql & "  And '" & UCase(ReplaceFamily(DSize.Text, DFamily.Text)) & "' Like (CASE WHEN LTRIM(RTRIM(Family))='' THEN '%%' ELSE LTRIM(RTRIM(Family)) END) "
                    sql = sql & "  And '" & UCase(iSliderList(i).Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider1))='' THEN '%%' ELSE LTRIM(RTRIM(Slider1)) END) "
                    sql = sql & "Order By SIZE + '|' + FAMILY + '|' + SLIDER1 "
                    Dim dtPerformaceUpSpec As DataTable = uDataBase.GetDataTable(sql)
                    If dtPerformaceUpSpec.Rows.Count > 0 Then
                        '[OK-09]A1
                        'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                        '                    "[OK-09]A1" & "[WINGS:" & dtPerformaceUpSpec.Rows(0)("SLIDERINF").ToString.Trim & "]"
                    Else
                        '
                        ' HKENSIN
                        sql = "SELECT Top 1 SIZE + '|' + FAMILY + '|' + SLIDER1 As SLIDERINF "
                        sql = sql & "From M_IRWPerformaceUp "
                        sql = sql & "Where CAT IN ('HKENSIN') "
                        sql = sql & "  And Active = 1 "
                        sql = sql & "  And '" & UCase(DSize.Text.Trim) & "' Like ( case when LTRIM(RTRIM([SIZE]))='' then '" & UCase(DSize.Text.Trim) & "' else LTRIM(RTRIM([SIZE])) end ) "
                        sql = sql & "  And '" & UCase(DChain.Text.Trim) & "' Like ( case when LTRIM(RTRIM([ChainType]))='' then '" & UCase(DChain.Text.Trim) & "' else LTRIM(RTRIM([ChainType])) end ) "
                        sql = sql & "  And '" & UCase(DClass.Text.Trim) & "' Like ( case when LTRIM(RTRIM([Class]))='' then '" & UCase(DClass.Text.Trim) & "' else LTRIM(RTRIM([Class])) end ) "
                        sql = sql & "  And '" & UCase(DTape.Text.Trim) & "' Like ( case when LTRIM(RTRIM([Tape1]))='' then '" & UCase(DTape.Text.Trim) & "' else LTRIM(RTRIM([Tape1])) end ) "
                        sql = sql & "  And '" & UCase(DSlider1.Text.Trim) & "' Like ( case when LTRIM(RTRIM([Slider1]))='' then '" & UCase(DSlider1.Text.Trim) & "' else LTRIM(RTRIM([Slider1])) end ) "
                        sql = sql & "  And '" & UCase(DFinish1.Text.Trim) & "' Like ( case when LTRIM(RTRIM([FINISH1]))='' then '" & UCase(DFinish1.Text.Trim) & "' else LTRIM(RTRIM([FINISH1])) end ) "
                        sql = sql & "  And '" & UCase(DSlider2.Text.Trim) & "' Like ( case when LTRIM(RTRIM([Slider1]))='' then '" & UCase(DSlider2.Text.Trim) & "' else LTRIM(RTRIM([Slider2])) end ) "
                        sql = sql & "  And '" & UCase(DFinish2.Text.Trim) & "' Like ( case when LTRIM(RTRIM([FINISH1]))='' then '" & UCase(DFinish2.Text.Trim) & "' else LTRIM(RTRIM([FINISH2])) end ) "
                        sql = sql & "  And '" & UCase(DFamily.Text.Trim) & "' Like ( case when LTRIM(RTRIM([Family]))='' then '" & UCase(DFamily.Text.Trim) & "' else LTRIM(RTRIM([Family])) end ) "
                        sql = sql & "  And '" & UCase(xSpecial.Trim) & "' Like ( case when LTRIM(RTRIM([SpecialList]))='' then '" & UCase(xSpecial.Trim) & "' else LTRIM(RTRIM([SpecialList])) end ) "
                        sql = sql & "  And Seqno <> 999 "
                        sql = sql & "Order By SIZE + '|' + FAMILY + '|' + SLIDER1 "
                        Dim dtPerformaceUpSpec1 As DataTable = uDataBase.GetDataTable(sql)
                        If dtPerformaceUpSpec1.Rows.Count > 0 Then
                            '[OK-09]A2
                            'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                            '                    "[OK-09]A2" & "[WINGS:" & dtPerformaceUpSpec1.Rows(0)("SLIDERINF").ToString.Trim & "]"
                        Else
                            '**-999-START
                            xALL = DSize.Text & "/" & DChain.Text & "/" & _
                                   DClass.Text & "/" & DTape.Text & "/" & _
                                   DSlider1.Text & "/" & DFinish1.Text & "/" & _
                                   DSlider2.Text & "/" & DFinish2.Text & "/" & DFamily.Text & "/" & xSpecial
                            '
                            'MsgBox(xALL)
                            '
                            sql = "SELECT Top 1 SIZE + '|' + FAMILY + '|' + SLIDER1 As SLIDERINF "
                            sql = sql & "From M_IRWPerformaceUp "
                            sql = sql & "Where CAT IN ('HKENSIN') "
                            sql = sql & "  And Active = 1 "
                            sql = sql & "  And Seqno =  999 "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Size))='' THEN '%%' ELSE LTRIM(RTRIM(Size)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(ChainType))='' THEN '%%' ELSE LTRIM(RTRIM(ChainType)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Class))='' THEN '%%' ELSE LTRIM(RTRIM(Class)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Tape1))='' THEN '%%' ELSE LTRIM(RTRIM(Tape1)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider1))='' THEN '%%' ELSE LTRIM(RTRIM(Slider1)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Body1))='' THEN '%%' ELSE LTRIM(RTRIM(Body1)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Finish1))='' THEN '%%' ELSE LTRIM(RTRIM(Finish1)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider2))='' THEN '%%' ELSE LTRIM(RTRIM(Slider2)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Body2))='' THEN '%%' ELSE LTRIM(RTRIM(Body2)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Finish2))='' THEN '%%' ELSE LTRIM(RTRIM(Finish2)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Family))='' THEN '%%' ELSE LTRIM(RTRIM(Family)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(SpecialList))='' THEN '%%' ELSE LTRIM(RTRIM(SpecialList)) END) "
                            '
                            sql = sql & "Order By SIZE + '|' + FAMILY + '|' + SLIDER1 "
                            Dim dtPerformaceUpSpec2 As DataTable = uDataBase.GetDataTable(sql)
                            If dtPerformaceUpSpec2.Rows.Count > 0 Then
                                '[OK-09]A3
                                'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                                '                    "[OK-09]A3" & "[WINGS:" & dtPerformaceUpSpec2.Rows(0)("SLIDERINF").ToString.Trim & "]"
                            Else
                                '
                                uError = True
                                msg = "[" & "SIZE=" & UCase(DSize.Text.Trim) & "] [" & _
                                            "FAMILY=" & UCase(DFamily.Text.Trim) & "] [" & _
                                            "SLIDER=" & iSliderList(i) & "]"
                                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[NG-09]" & msg
                            End If
                            '-999-END
                        End If
                    End If
                    '
                    If InStr(DLimitedItem.Text, "[NG-09]") > 0 Then
                        '
                        sql = "Insert into M_IRWPerformaceUp ( "
                        sql = sql & "Active,Cat,Seqno,Size,ChainType, "
                        sql = sql & "Class,Tape1,Slider1,Body1,Finish1, "
                        sql = sql & "Slider2,Body2,Finish2,Family,SpecialList, "
                        sql = sql & "RDSpec,RDFinish,YOBI1,YOBI2,YOBI3, "
                        sql = sql & "CreateUser, CreateTime, ModifyUser, ModifyTime "
                        sql = sql & ") "
                        sql = sql & "VALUES( "
                        sql &= "'0' ,"
                        sql &= "'" & "HKENSIN" & "' ,"
                        sql &= "4, "
                        sql &= "'" & UCase(DSize.Text.Trim) & "' ,"
                        sql &= "'" & UCase(DChain.Text.Trim) & "' ,"
                        '
                        sql &= "'" & UCase(DClass.Text.Trim) & "' ,"
                        sql &= "'" & UCase(DTape.Text.Trim) & "' ,"
                        sql &= "'" & UCase(DSlider1.Text.Trim) & "' ,"
                        sql &= "'" & "" & "' ,"
                        sql &= "'" & UCase(DFinish1.Text.Trim) & "' ,"
                        '
                        sql &= "'" & UCase(DSlider2.Text.Trim) & "' ,"
                        sql &= "'" & "" & "' ,"
                        sql &= "'" & UCase(DFinish2.Text.Trim) & "' ,"
                        sql &= "'" & UCase(DFamily.Text.Trim) & "' ,"
                        sql &= "'" & xSpecial & "' ,"
                        '
                        sql &= "'', '', 'ERROR', 'SPEC', '', "
                        '
                        sql &= "'" & Request.QueryString("pUserID") & "' ,"
                        sql &= "'" & NowDateTime & "' ,"
                        sql &= "'" & Request.QueryString("pUserID") & "' ,"
                        sql &= "'" & NowDateTime & "' )"
                        uDataBase.ExecuteNonQuery(sql)
                        '
                    End If
                    '
                Next
                '
                If UBound(iSliderList) > 0 Then
                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "型別檢測-End" & "]"
                End If
                '
                If InStr(DLimitedItem.Text, "NG-09") > 0 Then
                    uJavaScript.PopMsg(Me, "[SPEC]型別(SIZE-FAMILY-SLIDER)是否正確？ (無法申請)")
                End If
                '
            End If
            '
        End If
        'End
        ' --------------------------------
        '**     A.MFGAP 20241109
        ' Start
        If pCat = "MFGAP" Then
            '
            '
            If UCase(DST1.Text.Trim) & UCase(DST2.Text.Trim) & UCase(DST4.Text.Trim) = "111" Then
                '
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "MF-GAP檢測-Start" & "]"
                '
                ' 特殊CH找出
                xSpecial = UCase(DSRequest1.Text.Trim) & "/" & UCase(DSRequest2.Text.Trim) & "/" & _
                           UCase(DSRequest3.Text.Trim) & "/" & UCase(DSRequest4.Text.Trim) & "/" & _
                           UCase(DSRequest5.Text.Trim) & "/" & UCase(DSRequest6.Text.Trim) & "/"
                xSpecial = Replace(xSpecial, "XXX", "")
                '
                ' MFGAP DEFAULT = GAP + 特殊SCDList
                xSlider = "GAP" & "/"
                sql = "SELECT TOP 1 SpecChain1, SpecChain2, SpecChain3, SpecChain4, SpecChain5, SpecChain6, "
                sql = sql & "SpecChain1+'/'+SpecChain2+'/'+SpecChain3+'/'+SpecChain4+'/'+SpecChain5+'/'+SpecChain6+'/' AS SCD "
                sql = sql & "From V_IRWPerformaceUp_SCDList_01 "
                sql = sql & "Where '" & xSpecial.Trim & "' LIKE '%' + SpecChain1 + '%' "
                sql = sql & "  And '" & xSpecial.Trim & "' LIKE '%' + SpecChain2 + '%' "
                sql = sql & "  And '" & xSpecial.Trim & "' LIKE '%' + SpecChain3 + '%' "
                sql = sql & "  And '" & xSpecial.Trim & "' LIKE '%' + SpecChain4 + '%' "
                sql = sql & "  And '" & xSpecial.Trim & "' LIKE '%' + SpecChain5 + '%' "
                sql = sql & "  And '" & xSpecial.Trim & "' LIKE '%' + SpecChain6 + '%' "
                sql = sql & "ORDER BY LEN(SpecChain1+'/'+SpecChain2+'/'+SpecChain3+'/'+SpecChain4+'/'+SpecChain5+'/'+SpecChain6+'/') DESC "
                Dim dtSCDList As DataTable = uDataBase.GetDataTable(sql)
                If dtSCDList.Rows.Count > 0 Then
                    '
                    'MsgBox(dtSCDList.Rows(0)("SCD").ToString.Trim)
                    '
                    If InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain1").ToString.Trim) > 0 And dtSCDList.Rows(0)("SpecChain1").ToString.Trim <> "" Then
                        '
                        str = Mid(xSpecial.Trim, InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain1").ToString.Trim), 99)
                        str = Mid(str, 1, InStr(str, "/") - 1)
                        '
                        If xSlider = "" Then
                            xSlider = str & "/"
                        Else
                            xSlider = xSlider & str & "/"
                        End If
                    End If
                    '
                    If InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain2").ToString.Trim) > 0 And dtSCDList.Rows(0)("SpecChain2").ToString.Trim <> "" Then
                        '
                        str = Mid(xSpecial.Trim, InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain2").ToString.Trim), 99)
                        str = Mid(str, 1, InStr(str, "/") - 1)
                        '
                        If xSlider = "" Then
                            xSlider = str & "/"
                        Else
                            xSlider = xSlider & str & "/"
                        End If
                    End If
                    '
                    If InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain3").ToString.Trim) > 0 And dtSCDList.Rows(0)("SpecChain3").ToString.Trim <> "" Then
                        '
                        str = Mid(xSpecial.Trim, InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain3").ToString.Trim), 99)
                        str = Mid(str, 1, InStr(str, "/") - 1)
                        '
                        If xSlider = "" Then
                            xSlider = str & "/"
                        Else
                            xSlider = xSlider & str & "/"
                        End If
                    End If
                    '
                    If InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain4").ToString.Trim) > 0 And dtSCDList.Rows(0)("SpecChain4").ToString.Trim <> "" Then
                        '
                        str = Mid(xSpecial.Trim, InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain4").ToString.Trim), 99)
                        str = Mid(str, 1, InStr(str, "/") - 1)
                        '
                        If xSlider = "" Then
                            xSlider = str & "/"
                        Else
                            xSlider = xSlider & str & "/"
                        End If
                    End If
                    '
                    If InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain5").ToString.Trim) > 0 And dtSCDList.Rows(0)("SpecChain5").ToString.Trim <> "" Then
                        '
                        str = Mid(xSpecial.Trim, InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain5").ToString.Trim), 99)
                        str = Mid(str, 1, InStr(str, "/") - 1)
                        '
                        If xSlider = "" Then
                            xSlider = str & "/"
                        Else
                            xSlider = xSlider & str & "/"
                        End If
                    End If
                    '
                    If InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain6").ToString.Trim) > 0 And dtSCDList.Rows(0)("SpecChain6").ToString.Trim <> "" Then
                        '
                        str = Mid(xSpecial.Trim, InStr(xSpecial.Trim, dtSCDList.Rows(0)("SpecChain6").ToString.Trim), 99)
                        str = Mid(str, 1, InStr(str, "/") - 1)
                        '
                        If xSlider = "" Then
                            xSlider = str & "/"
                        Else
                            xSlider = xSlider & str & "/"
                        End If
                    End If
                    '
                End If
                '
                'MsgBox(xSlider)
                '
                If xSlider <> "" Then
                    str = ""
                    iSliderList = xSlider.ToString.Split("/")
                    For j = 0 To UBound(iSliderList) - 1
                        If iSliderList(j).Trim <> "" Then
                            If str = "" Then
                                str = UCase(iSliderList(j).Trim) & "/"
                            Else
                                str = str & UCase(iSliderList(j).Trim) & "/"
                            End If
                        End If
                    Next
                    '
                    'MsgBox(str)
                    '
                    iSliderList = str.ToString.Split("/")
                    '
                    '81-41
                    sql = "SELECT Top 1 SIZE+'|'+ChainType+'|'+Class+'|'+Tape1+'|'+SpecialList as MFGAP "
                    sql = sql & "From M_IRWPerformaceUp "
                    sql = sql & "Where CAT IN ('WMFGAP') "
                    sql = sql & "  And Active = 1 "
                    sql = sql & "  And LTRIM(RTRIM(SIZE)) = '" & UCase(DSize.Text.Trim) & "' "
                    sql = sql & "  And LTRIM(RTRIM(ChainType)) = '" & GetGapChain(UCase(DChain.Text.Trim)) & "' "
                    sql = sql & "  And LTRIM(RTRIM(Class)) = '" & UCase(DClass.Text.Trim) & "' "
                    sql = sql & "  And LTRIM(RTRIM(Tape1)) = '" & UCase(DTape.Text.Trim) & "' "
                    '
                    For j = 0 To UBound(iSliderList) - 1
                        sql = sql & "  And LTRIM(RTRIM(SpecialList)) Like '%" & iSliderList(j).Trim & "%' "
                    Next
                    sql = sql & "Order By SIZE+'|'+ChainType+'|'+Class+'|'+Tape1+'|'+SpecialList "
                    '
                    'MsgBox(sql)
                    '
                    Dim dtPerformaceUpMFGAP As DataTable = uDataBase.GetDataTable(sql)
                    If dtPerformaceUpMFGAP.Rows.Count > 0 Then
                        'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                        '                    "[OK-A0]G1" & "[WINGS:" & dtPerformaceUpMFGAP.Rows(0)("MFGAP").ToString.Trim & "]"
                    Else
                        '
                        '**GAP
                        sql = "SELECT Top 1 SIZE+'|'+ChainType+'|'+Class+'|'+Tape1+'|'+SpecialList as MFGAP "
                        sql = sql & "From M_IRWPerformaceUp "
                        sql = sql & "Where CAT IN ('HMFGAP') "
                        sql = sql & "  And Active = 1 "
                        sql = sql & "  And LTRIM(RTRIM(SIZE)) = '" & UCase(DSize.Text.Trim) & "' "
                        sql = sql & "  And LTRIM(RTRIM(ChainType)) = '" & GetGapChain(UCase(DChain.Text.Trim)) & "' "
                        sql = sql & "  And LTRIM(RTRIM(Class)) = '" & UCase(DClass.Text.Trim) & "' "
                        sql = sql & "  And LTRIM(RTRIM(Tape1)) = '" & UCase(DTape.Text.Trim) & "' "
                        '
                        sql = sql & "  And ( "
                        sql = sql & "        ( "
                        For j = 0 To UBound(iSliderList) - 1
                            If j = 0 Then
                                sql = sql & "      LTRIM(RTRIM(SpecialList)) Like '%" & iSliderList(j).Trim & "%' "
                            Else
                                sql = sql & "  And LTRIM(RTRIM(SpecialList)) Like '%" & iSliderList(j).Trim & "%' "
                            End If
                        Next
                        sql = sql & "        ) "
                        sql = sql & "        Or "
                        sql = sql & "        ( LTRIM(RTRIM(SpecialList)) = 'OK/') "
                        sql = sql & "  ) "
                        '
                        sql = sql & "Order By SIZE+'|'+ChainType+'|'+Class+'|'+Tape1+'|'+SpecialList "
                        '
                        'MsgBox(sql)
                        '
                        Dim dtPerformaceUpMFGAPA As DataTable = uDataBase.GetDataTable(sql)
                        If dtPerformaceUpMFGAPA.Rows.Count > 0 Then
                            'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                            '                    "[OK-A0]G2" & "[WINGS:" & dtPerformaceUpMFGAPA.Rows(0)("MFGAP").ToString.Trim & "]"
                        Else
                            '**MFGAP-2 SEQNO=999
                            xALL = DSize.Text & "/" & DChain.Text & "/" & _
                                   DClass.Text & "/" & DTape.Text & "/" & _
                                   DSlider1.Text & "/" & DFinish1.Text & "/" & _
                                   DSlider2.Text & "/" & DFinish2.Text & "/" & DFamily.Text & "/" & xSpecial
                            '
                            'MsgBox(xALL)
                            '
                            sql = "SELECT Top 1 SIZE+'|'+ChainType+'|'+Class+'|'+Tape1+'|'+SpecialList as MFGAP "
                            sql = sql & "From M_IRWPerformaceUp "
                            sql = sql & "Where CAT IN ('HMFGAP') "
                            sql = sql & "  And Active = 1 "
                            sql = sql & "  And Seqno =  999 "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Size))='' THEN '%%' ELSE LTRIM(RTRIM(Size)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(ChainType))='' THEN '%%' ELSE LTRIM(RTRIM(ChainType)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Class))='' THEN '%%' ELSE LTRIM(RTRIM(Class)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Tape1))='' THEN '%%' ELSE LTRIM(RTRIM(Tape1)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider1))='' THEN '%%' ELSE LTRIM(RTRIM(Slider1)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Body1))='' THEN '%%' ELSE LTRIM(RTRIM(Body1)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Finish1))='' THEN '%%' ELSE LTRIM(RTRIM(Finish1)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider2))='' THEN '%%' ELSE LTRIM(RTRIM(Slider2)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Body2))='' THEN '%%' ELSE LTRIM(RTRIM(Body2)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Finish2))='' THEN '%%' ELSE LTRIM(RTRIM(Finish2)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Family))='' THEN '%%' ELSE LTRIM(RTRIM(Family)) END) "
                            sql = sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(SpecialList))='' THEN '%%' ELSE LTRIM(RTRIM(SpecialList)) END) "
                            '
                            'MsgBox(sql)
                            sql = sql & "Order By SIZE+'|'+ChainType+'|'+Class+'|'+Tape1+'|'+SpecialList "
                            Dim dtPerformaceUpMFGAPB As DataTable = uDataBase.GetDataTable(sql)
                            If dtPerformaceUpMFGAPB.Rows.Count > 0 Then
                                'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & _
                                '                    "[OK-A0]G3" & "[WINGS:" & dtPerformaceUpMFGAPB.Rows(0)("MFGAP").ToString.Trim & "]"
                            Else
                                uError = True
                                msg = "[" & "SIZE=" & UCase(DSize.Text.Trim) & "] [" & _
                                            "CHAIN=" & UCase(DChain.Text.Trim) & "] [" & _
                                            "CLASS=" & UCase(DClass.Text.Trim) & "] [" & _
                                            "TAPE=" & UCase(DTape.Text.Trim) & "] [" & _
                                            "SPECIAL=" & xSpecial & "]"
                                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[NG-A0]" & msg
                            End If
                            '
                        End If
                    End If
                End If
                '
                If InStr(DLimitedItem.Text, "[NG-A0]") > 0 Then
                    '
                    sql = "Insert into M_IRWPerformaceUp ( "
                    sql = sql & "Active,Cat,Seqno,Size,ChainType, "
                    sql = sql & "Class,Tape1,Slider1,Body1,Finish1, "
                    sql = sql & "Slider2,Body2,Finish2,Family,SpecialList, "
                    sql = sql & "RDSpec,RDFinish,YOBI1,YOBI2,YOBI3, "
                    sql = sql & "CreateUser, CreateTime, ModifyUser, ModifyTime "
                    sql = sql & ") "
                    sql = sql & "VALUES( "
                    sql &= "'0' ,"
                    sql &= "'" & "HMFGAP" & "' ,"
                    sql &= "1, "
                    sql &= "'" & UCase(DSize.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DChain.Text.Trim) & "' ,"
                    '
                    sql &= "'" & UCase(DClass.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DTape.Text.Trim) & "' ,"
                    sql &= "'" & UCase(DSlider1.Text.Trim) & "' ,"
                    sql &= "'" & "" & "' ,"
                    sql &= "'" & UCase(DFinish1.Text.Trim) & "' ,"
                    '
                    sql &= "'" & "" & "' ,"
                    sql &= "'" & "" & "' ,"
                    sql &= "'" & "" & "' ,"
                    sql &= "'" & UCase(DFamily.Text.Trim) & "' ,"
                    sql &= "'" & xSpecial & "' ,"
                    '
                    sql &= "'', '', 'ERROR', '', '', "
                    '
                    sql &= "'" & Request.QueryString("pUserID") & "' ,"
                    sql &= "'" & NowDateTime & "' ,"
                    sql &= "'" & Request.QueryString("pUserID") & "' ,"
                    sql &= "'" & NowDateTime & "' )"
                    uDataBase.ExecuteNonQuery(sql)
                    '
                End If
                '
                If InStr(DLimitedItem.Text, "NG-A0") > 0 Then
                    uJavaScript.PopMsg(Me, "[MF-GAP]金屬拉鏈(GAP)是否正確？ (無法申請)")
                End If
                '
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "MF-GAP檢測-End" & "]"
            End If
            '
        End If
        'End
        '
        Return uError
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetNewxSpecialFeat)
    '**     
    '** Add Function by Joy 20250611
    '*****************************************************************
    Public Function GetNewxSpecialFeat(ByVal pOption As String, ByVal pData As String, ByVal pSpecial As String) As String
        Dim i As Integer
        Dim Sql As String
        Dim SpecialList As String()
        Dim xData As String
        '
        xData = ""
        Select Case pOption
            Case "MOD-CHAINTYPE"
                ' ChainType.(MODIFY)
                xData = pData
                '
                Sql = "Select RDSpec From M_IRWPerformaceUp "
                Sql &= "Where Active = 1 "
                Sql &= "And CAT = 'HCOMBI' "
                Sql &= "And SEQNO = 920 "
                Sql &= "And RDFINISH = 'MOD' "
                Sql &= "And YOBI1 = 'CHAINTYPE' "
                '
                Sql &= "And '" & pData & "' LIKE ChainType "
                '
                Dim dt_ChainType As DataTable = uDataBase.GetDataTable(Sql)
                If dt_ChainType.Rows.Count > 0 Then
                    xData = dt_ChainType.Rows(0).Item("RDSpec")
                End If
                '
            Case "DEL-SPECIAL-1"
                ' SPECIAL(DELETE)
                SpecialList = pSpecial.ToString.Split("/")
                For i = 0 To UBound(SpecialList) - 1
                    Sql = "Select RDSpec From M_IRWPerformaceUp "
                    Sql &= "Where Active = 1 "
                    Sql &= "And CAT = 'HCOMBI' "
                    Sql &= "And SEQNO = 920 "
                    Sql &= "And RDFINISH = 'DEL' "
                    Sql &= "And YOBI1 = 'SPECIAL-1' "
                    '
                    Sql &= "And '" & SpecialList(i) & "' LIKE Slider1 "
                    '
                    Dim dt_Special As DataTable = uDataBase.GetDataTable(Sql)
                    If dt_Special.Rows.Count > 0 Then
                        SpecialList(i) = ""
                    End If
                Next
                '
                For i = 0 To UBound(SpecialList) - 1
                    If xData = "" Then
                        xData = SpecialList(i) & "/"
                    Else
                        xData = xData & SpecialList(i) & "/"
                    End If
                Next
                '
            Case "MOD-SPECIAL-1"
                ' SPECIAL(MODIFY)
                SpecialList = pSpecial.ToString.Split("/")
                For i = 0 To UBound(SpecialList) - 1
                    Sql = "Select RDSpec From M_IRWPerformaceUp "
                    Sql &= "Where Active = 1 "
                    Sql &= "And CAT = 'HCOMBI' "
                    Sql &= "And SEQNO = 920 "
                    Sql &= "And RDFINISH = 'MOD' "
                    Sql &= "And YOBI1 = 'SPECIAL-1' "
                    '
                    Sql &= "And '" & SpecialList(i) & "' LIKE Slider1 "
                    '
                    Dim dt_Special As DataTable = uDataBase.GetDataTable(Sql)
                    If dt_Special.Rows.Count > 0 Then
                        SpecialList(i) = dt_Special.Rows(0).Item("RDSpec")
                    End If
                Next
                '
                For i = 0 To UBound(SpecialList) - 1
                    If xData = "" Then
                        xData = SpecialList(i) & "/"
                    Else
                        xData = xData & SpecialList(i) & "/"
                    End If
                Next
                '
            Case "MOD-SPECIAL-1B"
                ' SPECIAL(MODIFY)
                SpecialList = pSpecial.ToString.Split("/")
                For i = 0 To UBound(SpecialList) - 1
                    Sql = "Select RDSpec From M_IRWPerformaceUp "
                    Sql &= "Where Active = 1 "
                    Sql &= "And CAT = 'HCOMBI' "
                    Sql &= "And SEQNO = 920 "
                    Sql &= "And RDFINISH = 'MOD' "
                    Sql &= "And YOBI1 = 'SPECIAL-1B' "
                    '
                    Sql &= "And '" & pData & "' LIKE '%'+BODY1 "
                    Sql &= "And '" & SpecialList(i) & "' LIKE Slider1 "
                    '
                    Dim dt_Special As DataTable = uDataBase.GetDataTable(Sql)
                    If dt_Special.Rows.Count > 0 Then
                        SpecialList(i) = dt_Special.Rows(0).Item("RDSpec")
                    End If
                Next
                '
                For i = 0 To UBound(SpecialList) - 1
                    If xData = "" Then
                        xData = SpecialList(i) & "/"
                    Else
                        xData = xData & SpecialList(i) & "/"
                    End If
                Next
                '
            Case "MOD-SPECIAL-1C"
                ' SPECIAL(MODIFY)
                SpecialList = pSpecial.ToString.Split("/")
                For i = 0 To UBound(SpecialList) - 1
                    Sql = "Select RDSpec From M_IRWPerformaceUp "
                    Sql &= "Where Active = 1 "
                    Sql &= "And CAT = 'HCOMBI' "
                    Sql &= "And SEQNO = 920 "
                    Sql &= "And RDFINISH = 'MOD' "
                    Sql &= "And YOBI1 = 'SPECIAL-1C' "
                    '
                    Sql &= "And '" & pSpecial & "' LIKE '%'+Body1 "
                    Sql &= "And '" & pSpecial & "' LIKE '%'+Slider1 "
                    Sql &= "And '" & SpecialList(i) & "' LIKE RDSpec "
                    '
                    Dim dt_Special As DataTable = uDataBase.GetDataTable(Sql)
                    If dt_Special.Rows.Count > 0 Then
                        SpecialList(i) = dt_Special.Rows(0).Item("RDSpec")
                    End If
                Next
                '
                For i = 0 To UBound(SpecialList) - 1
                    If xData = "" Then
                        xData = SpecialList(i) & "/"
                    Else
                        xData = xData & SpecialList(i) & "/"
                    End If
                Next
                '
            Case Else
        End Select
        '
        If xData = "" Then xData = pSpecial
        '
        Return xData
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetFullName)
    '**     
    '**
    '*****************************************************************
    Public Function GetFullName(ByVal pShortName As String) As String
        Dim Sql As String
        Dim xFullName As String
        '
        xFullName = ""
        Sql = "Select FullName From M_LimitedReviationList "
        Sql &= "Where Active = 1 "
        Sql &= "And ShortName = '" & pShortName & "' "
        Dim dt_ReviationList As DataTable = uDataBase.GetDataTable(Sql)
        If dt_ReviationList.Rows.Count > 0 Then
            xFullName = dt_ReviationList.Rows(0).Item("FullName")
        Else
            xFullName = ""
        End If
        '
        If xFullName = "" Then xFullName = pShortName
        '
        Return xFullName
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetGapChain)
    '**     
    '**
    '*****************************************************************
    Public Function GetGapChain(ByVal pChain As String) As String
        Dim Sql As String
        Dim xGapChain As String
        '
        xGapChain = ""
        Sql = "SELECT Top 1 Slider1 as NewChain "
        Sql = Sql & "From M_IRWPerformaceUp "
        Sql = Sql & "Where CAT IN ('GAP') "
        Sql = Sql & "  And Active = 1 "
        Sql = Sql & "  And LTRIM(RTRIM(ChainType)) = '" & UCase(pChain.Trim) & "' "
        Sql = Sql & "Order By Len(ChainType) Desc "
        Dim dt_GapChainList As DataTable = uDataBase.GetDataTable(Sql)
        If dt_GapChainList.Rows.Count > 0 Then
            xGapChain = dt_GapChainList.Rows(0).Item("NewChain")
        Else
            xGapChain = ""
        End If
        '
        If xGapChain = "" Then xGapChain = pChain
        '
        Return UCase(xGapChain.Trim)
        '
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(PerformaceUp) joyjoy
    '**     
    '**
    '*****************************************************************
    Public Function PerformaceUp(ByVal pCat As String) As Boolean
        Dim i As Integer
        Dim Sql, xSpecial, xALL As String
        Dim xPerformaceUp As Boolean
        '
        xPerformaceUp = False

        xSpecial = UCase(DSRequest1.Text.Trim) & "/" & UCase(DSRequest2.Text.Trim) & "/" & _
                   UCase(DSRequest3.Text.Trim) & "/" & UCase(DSRequest4.Text.Trim) & "/" & _
                   UCase(DSRequest5.Text.Trim) & "/" & UCase(DSRequest6.Text.Trim) & "/"
        xSpecial = Replace(xSpecial, "XXX", "")
        For i = 0 To 6
            xSpecial = Replace(xSpecial, "//", "/")
        Next
        '
        xALL = DSize.Text & "/" & DChain.Text & "/" & _
               DClass.Text & "/" & DTape.Text & "/" & _
               DSlider1.Text & "/" & DFinish1.Text & "/" & _
               DSlider2.Text & "/" & DFinish2.Text & "/" & DFamily.Text & "/" & xSpecial
        '
        Sql = "SELECT Top 1 * "
        Sql = Sql & "From M_IRWPerformaceUp "
        Sql = Sql & "Where CAT IN ('" & pCat & "') "
        Sql = Sql & "  And Active = 1 "
        Sql = Sql & "  And Seqno = 1 "
        Sql = Sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Size))='' THEN '%%' ELSE LTRIM(RTRIM(Size)) END) "
        Sql = Sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(ChainType))='' THEN '%%' ELSE LTRIM(RTRIM(ChainType)) END) "
        Sql = Sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Class))='' THEN '%%' ELSE LTRIM(RTRIM(Class)) END) "
        Sql = Sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Tape1))='' THEN '%%' ELSE LTRIM(RTRIM(Tape1)) END) "
        Sql = Sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider1))='' THEN '%%' ELSE LTRIM(RTRIM(Slider1)) END) "
        Sql = Sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Body1))='' THEN '%%' ELSE LTRIM(RTRIM(Body1)) END) "
        Sql = Sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Finish1))='' THEN '%%' ELSE LTRIM(RTRIM(Finish1)) END) "
        Sql = Sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Slider2))='' THEN '%%' ELSE LTRIM(RTRIM(Slider2)) END) "
        Sql = Sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Body2))='' THEN '%%' ELSE LTRIM(RTRIM(Body2)) END) "
        Sql = Sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Finish2))='' THEN '%%' ELSE LTRIM(RTRIM(Finish2)) END) "
        Sql = Sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(Family))='' THEN '%%' ELSE LTRIM(RTRIM(Family)) END) "
        Sql = Sql & "  And '" & UCase(xALL.Trim) & "' Like (CASE WHEN LTRIM(RTRIM(SpecialList))='' THEN '%%' ELSE LTRIM(RTRIM(SpecialList)) END) "
        Sql = Sql & "Order By SIZE, CHAINTYPE, CLASS, TAPE1, Slider1, Slider2, SpecialList "
        '
        Dim dt_IRWPerformaceUp As DataTable = uDataBase.GetDataTable(Sql)
        If dt_IRWPerformaceUp.Rows.Count > 0 Then
            xPerformaceUp = True
        End If
        '
        Return xPerformaceUp
        '
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ItemError)
    '**     檢查限定是否限定ITEM
    '**
    '*****************************************************************
    Public Function LimitedItem(ByVal pApply As String, ByVal pDataStr As String) As Boolean
        Dim uLimitedItem As Boolean = True
        Dim p0, p1A, p1C, p2A, p2C, p3A, p3C, sql As String
        Dim xDataStr As String()
        Dim i As Integer
        '
        xDataStr = pDataStr.Split("/")
        For i = 0 To xDataStr.Length - 1
            Select Case i
                Case 0
                    p0 = xDataStr(i)
                Case 1
                    p1A = xDataStr(i)
                Case 2
                    p1C = xDataStr(i)
                Case 3
                    p2A = xDataStr(i)
                Case 4
                    p2C = xDataStr(i)
                Case 5
                    p3A = xDataStr(i)
                Case 6
                    p3C = xDataStr(i)
                Case Else
            End Select
        Next
        '
        If p0 <> "*" Then
            'CONDITION-1
            If uLimitedItem = True Then
                Select Case p1A
                    Case "!%"
                        sql = "select "
                        sql = sql & "ISNULL( (SELECT TOP 1 1 FROM M_Referp WHERE '" & pApply & "' LIKE '" & ReplaceString(p1C) & "'), 0) AS WK "
                        Dim dtWK As DataTable = uDataBase.GetDataTable(sql)
                        If dtWK.Rows(0)("WK") = 1 Then uLimitedItem = False
                    Case "%"
                        sql = "select "
                        sql = sql & "ISNULL( (SELECT TOP 1 1 FROM M_Referp WHERE '" & pApply & "' LIKE '" & ReplaceString(p1C) & "'), 0) AS WK "
                        Dim dtWK As DataTable = uDataBase.GetDataTable(sql)
                        If dtWK.Rows(0)("WK") = 0 Then uLimitedItem = False
                    Case "!="
                        If pApply = ReplaceString(p1C) Then uLimitedItem = False
                    Case "="
                        If pApply <> ReplaceString(p1C) Then uLimitedItem = False
                    Case Else
                End Select
            End If
            'CONDITION-2
            If uLimitedItem = True Then
                Select Case p2A
                    Case "!%"
                        sql = "select "
                        sql = sql & "ISNULL( (SELECT TOP 1 1 FROM M_Referp WHERE '" & pApply & "' LIKE '" & ReplaceString(p2C) & "'), 0) AS WK "
                        Dim dtWK As DataTable = uDataBase.GetDataTable(sql)
                        If dtWK.Rows(0)("WK") = 1 Then uLimitedItem = False
                    Case "%"
                        sql = "select "
                        sql = sql & "ISNULL( (SELECT TOP 1 1 FROM M_Referp WHERE '" & pApply & "' LIKE '" & ReplaceString(p2C) & "'), 0) AS WK "
                        Dim dtWK As DataTable = uDataBase.GetDataTable(sql)
                        If dtWK.Rows(0)("WK") = 0 Then uLimitedItem = False
                    Case "!="
                        If pApply = ReplaceString(p2C) Then uLimitedItem = False
                    Case "="
                        If pApply <> ReplaceString(p2C) Then uLimitedItem = False
                    Case Else
                End Select
            End If
            'CONDITION-3
            If uLimitedItem = True Then
                Select Case p3A
                    Case "!%"
                        sql = "select "
                        sql = sql & "ISNULL( (SELECT TOP 1 1 FROM M_Referp WHERE '" & pApply & "' LIKE '" & ReplaceString(p3C) & "'), 0) AS WK "
                        Dim dtWK As DataTable = uDataBase.GetDataTable(sql)
                        If dtWK.Rows(0)("WK") = 1 Then uLimitedItem = False
                    Case "%"
                        sql = "select "
                        sql = sql & "ISNULL( (SELECT TOP 1 1 FROM M_Referp WHERE '" & pApply & "' LIKE '" & ReplaceString(p3C) & "'), 0) AS WK "
                        Dim dtWK As DataTable = uDataBase.GetDataTable(sql)
                        If dtWK.Rows(0)("WK") = 0 Then uLimitedItem = False
                    Case "!="
                        If pApply = ReplaceString(p3C) Then uLimitedItem = False
                    Case "="
                        If pApply <> ReplaceString(p3C) Then uLimitedItem = False
                    Case Else
                End Select
            End If
        End If
        '
        Return uLimitedItem
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(LimitedItemData)
    '**     新增LimitedItem資料
    '**
    '*****************************************************************
    Sub AppendLimitedItem(ByVal pShowLmtPg As Boolean)
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim sql As String = ""
        sql = "INSERT INTO W_LimitedItemRegister ( " & _
              "Sts, CompletedTime, FormNo, FormSno, " & _
              "No, Date, Name, JobTitle, Division, " & _
              "RCode, RItemName1, RItemName2, RItemName3, RSize, RChain, RClass, RTape, RSlider1, RFinish1, RSlider2, RFinish2, " & _
              "RSRequest1, RSRequest2, RSRequest3, RSRequest4, RSRequest5, RSRequest6, RFamily, " & _
              "RST1, RST2, RST3, RST4, RST5, RST6, RST7, RNoDisplay, " & _
              "Code, ItemName1, ItemName2, ItemName3, Size, Chain, Class, Tape, Slider1, Finish1, Slider2, Finish2, " & _
              "SRequest1, SRequest2, SRequest3, SRequest4, SRequest5, SRequest6, Family, " & _
              "ST1, ST2, ST3, ST4, ST5, ST6, ST7, NoDisplay, " & _
              "SLDPrice, Remark, Attachfile1, Attachfile2, A001, A206, A211, A999, K206, K211,Buyer,BuyerCode,Customer,Customercode,ForUse, OrderDate, SPDNo, SPDUrl, " & _
              "CreateUser, CreateTime, ModifyUser, ModifyTime) " & _
        "VALUES("
        '
        'MOD-END
        sql &= "'0' ,"
        sql &= "'" & NowDateTime & "' ,"
        sql &= "'" & wFormNo & "' ,"
        sql &= "'" & CStr(0) & "' ,"
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
        sql = sql + " '" + "NOFILE" + "', "
        sql = sql + " '" + "NOFILE" + "', "
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

        sql &= "'" & YKK.ReplaceString(DBuyer.Text) & "' ,"
        sql &= "'" & DBuyerCode.Text & "' ,"
        sql &= "'" & YKK.ReplaceString(DCustomer.Text) & "' ,"
        sql &= "'" & DCustomerCode.Text & "' ,"
        sql &= "'" & DForUse.SelectedValue & "' ,"
        sql &= "'" & DOrderDate.Text & "' ,"
        '
        sql &= "'" & DSPDNO.Text & "' ,"
        sql &= "'" & DLimitedItem.Text & "' ,"
        '
        sql &= "'" & Request.QueryString("pUserID") & "' ,"
        sql &= "'" & NowDateTime & "' ,"
        sql &= "'" & Request.QueryString("pUserID") & "' ,"
        sql &= "'" & NowDateTime & "' )"

        'uDataBase.ExecuteNonQuery(sql)
        'Insert後取得ID，供CkLimitedPage使用
        LimitedId = uDataBase.InsertAndReturnID(sql)
        If pShowLmtPg Then
            '設定限定檢查的連結，開啟限定檢查頁面
            setCkLimitedLink(LimitedId)
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(LimitedItem資料匯入)
    '**     
    '**
    '*****************************************************************
    Protected Sub BSPDNO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSPDNO.Click
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("ItemRegisterFilePath")
        Dim SQL, str As String
        '
        If DRemark.Text = "ITEMNAME" Then
            DNewItemInf.Text = UCase(DSize.Text.Trim) & "!" & _
                               UCase(DChain.Text.Trim) & "!" & _
                               UCase(DClass.Text.Trim) & "!" & _
                               UCase(DTape.Text.Trim) & "!" & _
                               UCase(DSlider1.Text.Trim) & "!" & _
                               UCase(DFinish1.Text.Trim) & "!" & _
                               UCase(DSlider2.Text.Trim) & "!" & _
                               UCase(DFinish2.Text.Trim) & "!" & _
                               UCase(DFamily.Text.Trim) & "!" & _
                               UCase(DSRequest1.Text.Trim) & "!" & _
                               UCase(DSRequest2.Text.Trim) & "!" & _
                               UCase(DSRequest3.Text.Trim) & "!" & _
                               UCase(DSRequest4.Text.Trim) & "!" & _
                               UCase(DSRequest5.Text.Trim) & "!" & _
                               UCase(DSRequest6.Text.Trim) & "!"
            DNewItemInf.Text = Replace(DNewItemInf.Text, "XXX!", "")
            '
            str = UCase(DItemName1.Text.Trim) & UCase(DItemName2.Text.Trim) & UCase(DItemName3.Text.Trim)
            If DNewItemInf.Text <> DItemInf.Text Or _
               (UCase(DTape.Text.Trim) <> "" And InStr(str, UCase(DTape.Text.Trim)) <= 0) Or _
               (UCase(DSlider1.Text.Trim) <> "" And InStr(str, UCase(DSlider1.Text.Trim)) <= 0) Or _
               (UCase(DFinish1.Text.Trim) <> "" And InStr(str, UCase(DFinish1.Text.Trim)) <= 0) Or _
               (UCase(DSlider2.Text.Trim) <> "" And InStr(str, UCase(DSlider2.Text.Trim)) <= 0) Or _
               (UCase(DFinish2.Text.Trim) <> "" And InStr(str, UCase(DFinish2.Text.Trim)) <= 0) Or _
               (UCase(DSRequest1.Text.Trim) <> "" And InStr(str, UCase(DSRequest1.Text.Trim)) <= 0) Or _
               (UCase(DSRequest2.Text.Trim) <> "" And InStr(str, UCase(DSRequest2.Text.Trim)) <= 0) Or _
               (UCase(DSRequest3.Text.Trim) <> "" And InStr(str, UCase(DSRequest3.Text.Trim)) <= 0) Or _
               (UCase(DSRequest4.Text.Trim) <> "" And InStr(str, UCase(DSRequest4.Text.Trim)) <= 0) Or _
               (UCase(DSRequest5.Text.Trim) <> "" And InStr(str, UCase(DSRequest5.Text.Trim)) <= 0) Or _
               (UCase(DSRequest6.Text.Trim) <> "" And InStr(str, UCase(DSRequest6.Text.Trim)) <= 0) Then
                uJavaScript.PopMsg(Me, "[ITEMNAME]: ERROR ! ")
            Else
                uJavaScript.PopMsg(Me, "[ITEMNAME]: OK ! ")
            End If
        End If
        '
        If DRemark.Text = "EDX" Then
            '
            'ADD-START JOY 20250409
            If PerformaceUp("IMPFINISH") = True Then
                DPerformaceUp.Text = "IMPFINISH"
                DPerformaceUp.Style("left") = -500 & "px"
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[進口品檢測-Start]"
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "ITEM = 進口品"
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[進口品-End]"
            End If
            'ADD-END
            '
            'MODIFY-START JOY 20250409
            If DPerformaceUp.Text <> "IMPFINISH" Then
                '
                'EDX CHECK
                If EDXNotFound("BUYER") = True Then
                    uJavaScript.PopMsg(Me, "[EDX1]:EDX ERROR ! ")
                Else
                    uJavaScript.PopMsg(Me, "[EDX1]:EDX OK ! ")
                End If
                '
            Else
                uJavaScript.PopMsg(Me, "[EDX1]:EDX OK ! ")
            End If
            'MODIFY-END
        ElseIf Trim(DSLDPrice.Text) = "NEW" Then
            '將Item資料存入table並執行限定檢測
            AppendLimitedItem(True)
        Else
            If IsNumeric(DRemark.Text) = True Then
                '
                SQL = "Select * From W_LimitedItemRegister "
                SQL &= " Where Unique_id =  " & DRemark.Text & " "
                Dim dtItemRegisterSheet As DataTable = uDataBase.GetDataTable(SQL)
                If dtItemRegisterSheet.Rows.Count > 0 Then
                    '表單資料
                    DNo.Text = dtItemRegisterSheet.Rows(0).Item("No")                           ' No
                    DDate.Text = dtItemRegisterSheet.Rows(0).Item("Date")                       ' 申請日
                    DName.Text = dtItemRegisterSheet.Rows(0).Item("Name")                       ' 申請人姓名
                    DJobTitle.Text = dtItemRegisterSheet.Rows(0).Item("JobTitle")               ' 職稱
                    DDivision.Text = dtItemRegisterSheet.Rows(0).Item("Division")               ' 部門
                    DSPDNO.Text = dtItemRegisterSheet.Rows(0).Item("SPDNO")                     ' SPDNO


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
                    DPriceDescr.Text = ""
                    DPriceNo.Text = ""
                    DSLDPrice.Text = dtItemRegisterSheet.Rows(0).Item("SLDPrice")               ' SLD Price
                    DRemark.Text = dtItemRegisterSheet.Rows(0).Item("Remark")                   ' Remark
                    If dtItemRegisterSheet.Rows(0).Item("Attachfile1") <> "" Then
                        LAttachfile1.NavigateUrl = Path & dtItemRegisterSheet.Rows(0).Item("Attachfile1")   '附件
                        LAttachfile1.Visible = True
                    Else
                        LAttachfile1.Visible = False
                    End If

                    If dtItemRegisterSheet.Rows(0).Item("Attachfile2") <> "" Then
                        LAttachfile2.NavigateUrl = Path & dtItemRegisterSheet.Rows(0).Item("Attachfile2")   '憑證
                        LAttachfile2.Visible = True
                    Else
                        LAttachfile2.Visible = False
                    End If


                    DBuyer.Text = dtItemRegisterSheet.Rows(0).Item("Buyer")                         ' BUYER
                    DBuyerCode.Text = dtItemRegisterSheet.Rows(0).Item("BuyerCode")                         ' BUYER
                    DCustomer.Text = dtItemRegisterSheet.Rows(0).Item("Customer")                         ' Customer
                    DCustomerCode.Text = dtItemRegisterSheet.Rows(0).Item("CustomerCode")                         ' Customer
                    SetFieldData("ForUse", dtItemRegisterSheet.Rows(0).Item("ForUse"))    '用途
                    '
                    'ADD-START ISOS-2308 PJ
                    'Order Date
                    DOrderDate.Text = dtItemRegisterSheet.Rows(0).Item("OrderDate")     ' Order Date
                    DSPDNO.Text = dtItemRegisterSheet.Rows(0).Item("SPDNo")             ' SPD No
                    DSPDNOUrl.Text = dtItemRegisterSheet.Rows(0).Item("SPDUrl")         ' SPD URL
                    '
                    'ADD-END
                    '
                End If
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ReplaceString)
    '**     置換限定ITEM關鍵字
    '**
    '*****************************************************************
    Public Function ReplaceString(ByVal pData As String) As String
        Dim i As Integer
        Dim str As String = pData
        Dim xFrom As String = "[SP]/"
        Dim xTo As String = "/"
        Dim xFromWord As String()
        Dim xToWord As String()
        '
        xFromWord = xFrom.Split("/")
        xToWord = xTo.Split("/")
        For i = 0 To xFromWord.Length - 1
            str = Replace(str, xFromWord(i), xToWord(i))
        Next
        '
        Return str
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ItemNotProd)
    '**     檢查生產NG
    '**
    '*****************************************************************
    'Public Function ItemNotProd(ByVal pCat As String) As Boolean
    '    Dim i, j As Integer
    '    Dim sql, str, msg, xItemName, xSpc As String
    '    Dim uError As Boolean = False
    '    Dim uRun As Boolean = False
    '    Dim iNew As String()
    '    '
    '    xItemName = DItemName1.Text & " " & DItemName2.Text & " " & DItemName3.Text
    '    '
    '    xSpc = DSRequest1.Text & "!" & _
    '           DSRequest2.Text & "!" & _
    '           DSRequest3.Text & "!" & _
    '           DSRequest4.Text & "!" & _
    '           DSRequest5.Text & "!" & _
    '           DSRequest6.Text & "!"
    '    '
    '    str = DSize.Text & "!" & _
    '            DChain.Text & "!" & _
    '            DClass.Text & "!" & _
    '            DTape.Text & "!" & _
    '            DSlider1.Text & "!" & _
    '            DFinish1.Text & "!" & _
    '            DSlider2.Text & "!" & _
    '            DFinish2.Text & "!" & _
    '            DSRequest1.Text & "!" & _
    '            DSRequest2.Text & "!" & _
    '            DSRequest3.Text & "!" & _
    '            DSRequest4.Text & "!" & _
    '            DSRequest5.Text & "!" & _
    '            DSRequest6.Text & "!" & _
    '            DFamily.Text & "!" & _
    '            DST1.Text & "!" & _
    '            DST2.Text & "!" & _
    '            DST3.Text & "!" & _
    '            DST4.Text & "!" & _
    '            DST5.Text & "!" & _
    '            DST6.Text & "!" & _
    '            DST7.Text & "!"
    '    '
    '    iNew = str.ToString.Split("!")
    '    '
    '    If Not uError Then
    '        sql = "SELECT "
    '        sql = sql & "Size,Chain,Class,Tape,Slider1,Finish1,Slider2,Finish2,SRequest1,SRequest2,SRequest3,SRequest4,SRequest5,SRequest6,Family, "
    '        sql = sql & "ST1,ST2,ST3,ST4,ST5,ST6,ST7, CONDITION,RESULT,MSG,Cat,Rno,SubNo "
    '        sql = sql & "From M_ITEMRULE "
    '        sql = sql & "where Active = 1 "
    '        '
    '        If pCat = "BUYER" Then
    '            sql = sql & "and Cat <> 'ITEM' "
    '        Else
    '            sql = sql & "and Cat = '" & pCat & "' "
    '        End If
    '        '
    '        sql = sql & "order by Active,Cat,Rno,SubNo "
    '        '
    '        Dim dtRule As DataTable = uDataBase.GetDataTable(sql)
    '        If dtRule.Rows.Count > 0 Then
    '            For i = 0 To dtRule.Rows.Count - 1
    '                'MsgBox(CStr(i))
    '                uRun = False
    '                j = 0
    '                If (dtRule.Rows(i)(j) = "*" Or ErrorRule(iNew(j), dtRule.Rows(i)(j)) = False) And _
    '                   (dtRule.Rows(i)(j + 1) = "*" Or ErrorRule(iNew(j + 1), dtRule.Rows(i)(j + 1)) = False) And _
    '                   (dtRule.Rows(i)(j + 2) = "*" Or ErrorRule(iNew(j + 2), dtRule.Rows(i)(j + 2)) = False) And _
    '                   (dtRule.Rows(i)(j + 3) = "*" Or ErrorRule(iNew(j + 3), dtRule.Rows(i)(j + 3)) = False) And _
    '                   (dtRule.Rows(i)(j + 4) = "*" Or ErrorRule(iNew(j + 4), dtRule.Rows(i)(j + 4)) = False) And _
    '                   (dtRule.Rows(i)(j + 5) = "*" Or ErrorRule(iNew(j + 5), dtRule.Rows(i)(j + 5)) = False) And _
    '                   (dtRule.Rows(i)(j + 6) = "*" Or ErrorRule(iNew(j + 6), dtRule.Rows(i)(j + 6)) = False) And _
    '                   (dtRule.Rows(i)(j + 7) = "*" Or ErrorRule(iNew(j + 7), dtRule.Rows(i)(j + 7)) = False) And _
    '                   (dtRule.Rows(i)(j + 8) = "*" Or ErrorRule(xSpc, dtRule.Rows(i)(j + 8)) = False) And _
    '                   (dtRule.Rows(i)(j + 9) = "*" Or ErrorRule(xSpc, dtRule.Rows(i)(j + 9)) = False) And _
    '                   (dtRule.Rows(i)(j + 10) = "*" Or ErrorRule(xSpc, dtRule.Rows(i)(j + 10)) = False) And _
    '                   (dtRule.Rows(i)(j + 11) = "*" Or ErrorRule(xSpc, dtRule.Rows(i)(j + 11)) = False) And _
    '                   (dtRule.Rows(i)(j + 12) = "*" Or ErrorRule(xSpc, dtRule.Rows(i)(j + 12)) = False) And _
    '                   (dtRule.Rows(i)(j + 13) = "*" Or ErrorRule(xSpc, dtRule.Rows(i)(j + 13)) = False) And _
    '                   (dtRule.Rows(i)(j + 14) = "*" Or ErrorRule(iNew(j + 14), dtRule.Rows(i)(j + 14)) = False) And _
    '                   (dtRule.Rows(i)(j + 15) = "*" Or ErrorRule(iNew(j + 15), dtRule.Rows(i)(j + 15)) = False) And _
    '                   (dtRule.Rows(i)(j + 16) = "*" Or ErrorRule(iNew(j + 16), dtRule.Rows(i)(j + 16)) = False) And _
    '                   (dtRule.Rows(i)(j + 17) = "*" Or ErrorRule(iNew(j + 17), dtRule.Rows(i)(j + 17)) = False) And _
    '                   (dtRule.Rows(i)(j + 18) = "*" Or ErrorRule(iNew(j + 18), dtRule.Rows(i)(j + 18)) = False) And _
    '                   (dtRule.Rows(i)(j + 19) = "*" Or ErrorRule(iNew(j + 19), dtRule.Rows(i)(j + 19)) = False) And _
    '                   (dtRule.Rows(i)(j + 20) = "*" Or ErrorRule(iNew(j + 20), dtRule.Rows(i)(j + 20)) = False) And _
    '                   (dtRule.Rows(i)(j + 21) = "*" Or ErrorRule(iNew(j + 21), dtRule.Rows(i)(j + 21)) = False) Then
    '                    uRun = True
    '                Else
    '                    If dtRule.Rows(i)(j) = "*ALL" Or dtRule.Rows(i)(j + 1) = "*ALL" Or _
    '                       dtRule.Rows(i)(j + 2) = "*ALL" Or dtRule.Rows(i)(j + 3) = "*ALL" Or _
    '                       dtRule.Rows(i)(j + 4) = "*ALL" Or dtRule.Rows(i)(j + 5) = "*ALL" Or _
    '                       dtRule.Rows(i)(j + 6) = "*ALL" Or dtRule.Rows(i)(j + 7) = "*ALL" Or _
    '                       dtRule.Rows(i)(j + 8) = "*ALL" Or dtRule.Rows(i)(j + 9) = "*ALL" Or _
    '                       dtRule.Rows(i)(j + 10) = "*ALL" Or dtRule.Rows(i)(j + 11) = "*ALL" Or _
    '                       dtRule.Rows(i)(j + 12) = "*ALL" Or dtRule.Rows(i)(j + 13) = "*ALL" Or _
    '                       dtRule.Rows(i)(j + 14) = "*ALL" Or dtRule.Rows(i)(j + 15) = "*ALL" Or _
    '                       dtRule.Rows(i)(j + 16) = "*ALL" Or dtRule.Rows(i)(j + 17) = "*ALL" Or _
    '                       dtRule.Rows(i)(j + 18) = "*ALL" Or dtRule.Rows(i)(j + 19) = "*ALL" Or _
    '                       dtRule.Rows(i)(j + 20) = "*ALL" Or dtRule.Rows(i)(j + 21) = "*ALL" Then
    '                        uRun = True
    '                    End If
    '                End If
    '                '
    '                If uRun Then
    '                    'MsgBox("IN-ITEM-BUYER")
    '                    str = dtRule.Rows(i)("CONDITION")
    '                    str = Replace(str, "%USERID%", Request.QueryString("pUserID"))
    '                    str = Replace(str, "%SLIDER1%", DSlider1.Text)
    '                    str = Replace(str, "%SLIDER2%", DSlider2.Text)
    '                    str = Replace(str, "%BUYER%", DBuyerCode.Text)
    '                    str = Replace(str, "%CUSTOMER%", DCustomerCode.Text)
    '                    str = Replace(str, "%ITEMNAME%", xItemName)
    '                    str = Replace(str, "^", "'")
    '                    sql = str
    '                    'MsgBox(sql)
    '                    Dim dtData As DataTable = uDataBase.GetDataTable(sql)
    '                    If dtData.Rows(0)("result") = 0 Then
    '                        msg = dtRule.Rows(i)("subno") & ":" & dtRule.Rows(i)("msg")
    '                        uError = True
    '                        Exit For
    '                    Else
    '                        msg = dtRule.Rows(i)("subno") & ":" & "ok"
    '                    End If
    '                End If
    '            Next
    '        End If
    '    End If
    '    '
    '    If uError Then uJavaScript.PopMsg(Me, msg)
    '    '
    '    Return uError
    'End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(NotEDXItem)
    '**     檢查不是EDX ITEM
    '**
    '*****************************************************************
    Public Function NotEDXItem(ByVal pSlider As String, ByVal pBuyer As String) As Boolean
        Dim i As Integer
        Dim xPuller, xOPuller, xCut As String
        Dim sql, str, msg As String
        Dim xSource, xTarget As String()
        Dim uError As Boolean = False
        Dim uRun As Boolean = False
        Dim uRunDB As Boolean = False
        Dim uPullerGroup As Boolean = False
        '
        xPuller = ""
        xOPuller = ""
        str = ""
        msg = ""
        'MsgBox(pSlider)
        '
        'Slider 
        If pSlider <> "" Then
            uRun = True
            uRunDB = True
            '
            'SLIDER重新編成 (DA7B% & DAG7B)
            If Mid(pSlider, 1, 4) = "DA7B" Or Mid(pSlider, 1, 5) = "DAG7B" Or Mid(pSlider, 1, 4) = "DS15" Or Mid(pSlider, 1, 5) = "DAG24" Then
                'PULLER編成
                str = DSRequest1.Text & "/" & DSRequest2.Text & "/" & DSRequest3.Text & "/" & DSRequest4.Text & "/" & DSRequest5.Text & "/" & DSRequest6.Text & "/"
                xSource = Split(str, "/")
                For i = 0 To UBound(xSource) - 1
                    If Mid(xSource(i), 1, 3) = "PL-" Or Mid(xSource(i), 1, 5) = "T-PL-" Or Mid(xSource(i), 1, 5) = "B-PL-" Then
                        xPuller = Mid(xSource(i), InStr(xSource(i), "PL-") + 3)
                    End If
                Next
                If xPuller <> "" Then
                    '橡膠色編成
                    xOPuller = pSlider
                    If InStr(pSlider, "-") > 0 Then
                        xOPuller = Mid(pSlider, 1, InStr(pSlider, "-") - 1)
                    End If
                    If Mid(pSlider, 1, 4) = "DA7B" Then xOPuller = Replace(xOPuller, "DA7B", "")
                    If Mid(pSlider, 1, 5) = "DAG7B" Then xOPuller = Replace(xOPuller, "DAG7B", "")
                    If Mid(pSlider, 1, 4) = "DS15" Then xOPuller = Replace(xOPuller, "DS15", "")
                    If Mid(pSlider, 1, 5) = "DAG24" Then xOPuller = Replace(xOPuller, "DAG24", "")
                    'SLIDER編成
                    str = ""
                    If Mid(pSlider, 1, 4) = "DA7B" Then
                        If xPuller <> "" Then str = "DA7B" & xPuller
                        If xOPuller <> "" Then str = str & xOPuller
                    End If
                    If Mid(pSlider, 1, 5) = "DAG7B" Then
                        If xPuller <> "" Then str = "DAG7B" & xPuller
                        If xOPuller <> "" Then str = str & xOPuller
                    End If
                    If Mid(pSlider, 1, 4) = "DS15" Then
                        If xPuller <> "" Then str = "DS15" & xPuller
                        If xOPuller <> "" Then str = str & xOPuller
                    End If
                    If Mid(pSlider, 1, 5) = "DAG24" Then
                        If xPuller <> "" Then str = "DAG24" & xPuller
                        If xOPuller <> "" Then str = str & xOPuller
                    End If
                    If str <> "" Then pSlider = str
                End If
            End If
            'MsgBox("SLIDER:" & pSlider)
            '
            xPuller = ""
            xOPuller = ""
            str = ""
            msg = ""
            '
            'REPLACE BODY 
            sql = "SELECT Top 1 "
            sql = sql & "Body, Source, Target "
            sql = sql & "From M_EDXBodyList "
            sql = sql & "where Active = 1 "
            sql = sql & "and '" & pSlider & "' like '%' + body + '%' "
            sql = sql & "order by len(body) desc, body desc "
            Dim dtBody As DataTable = uDataBase.GetDataTable(sql)
            If dtBody.Rows.Count > 0 Then
                str = Replace(pSlider, dtBody.Rows(0)("Body"), "")
                '
                If dtBody.Rows(0)("Source") <> "" And dtBody.Rows(0)("Target") <> "" Then
                    xSource = Split(dtBody.Rows(0)("Source"), "/")
                    xTarget = Split(dtBody.Rows(0)("Target"), "/")
                    If UBound(xSource) = UBound(xTarget) Then
                        '
                        For i = 0 To UBound(xSource) - 1
                            If Mid(str, 1, Len(xSource(i))) = xSource(i) Then
                                str = Replace(str, xSource(i), xTarget(i))
                            End If
                        Next
                    End If
                End If
                '
                ' 只有 BODY
                If str = "" Then
                    str = dtBody.Rows(0)("Body")
                    uRun = False
                End If
                '
                xPuller = str
            Else
                msg = msg & "[B]"
            End If
            'MsgBox("BODY:" & str)
            '
            ' Run
            'If uRun Then
            '    '
            '    'SPD PULLER / BUYER 有?
            '    '  (亞鉛)
            '    xPuller = str
            '    '
            '    sql = "SELECT Top 1 "
            '    sql = sql & "PullerCode, Buyer "
            '    sql = sql & "From M_ManufOutForEDX "
            '    '
            '    sql = sql & "Where PullerCode = '" & xPuller & "' "
            '    sql = sql & "And (Buyer = 'ALL' or Buyer like '%" & pBuyer & "%') "
            '    sql = sql & "order by SYS "
            '    '
            '    'MsgBox(sql)
            '    Dim dtSPDOut1 As DataTable = uDataBase.GetDataTable(sql)
            '    If dtSPDOut1.Rows.Count > 0 Then
            '        'MsgBox(pSlider & ":" & xPuller)
            '        If InStr(pSlider, xPuller) > 0 Then
            '            xPuller = Mid(pSlider, InStr(pSlider, xPuller))
            '            '
            '            ' Run CHECK WINGS & EDX DB
            '            If uRun Then
            '                str = xPuller & msg
            '                If InStr(str, "[") > 0 Then
            '                    str = Mid(str, 1, InStr(str, "[") - 1)
            '                End If
            '                sql = "SELECT Top 1 "
            '                sql = sql & "Slider "
            '                sql = sql & "From M_EDXSlider "
            '                sql = sql & "Where Slider LIKE '%" & str & "%' "
            '                sql = sql & "order by Slider "
            '                'MsgBox(sql)
            '                Dim dtWings As DataTable = uDataBase.GetDataTable(sql)
            '                If dtWings.Rows.Count > 0 Then
            '                    msg = msg & "[WY]"
            '                Else
            '                    msg = msg & "[WN]"
            '                End If
            '            End If
            '        End If
            '        'RUN FLAG
            '        uRun = False
            '    End If
            '    'SPD PULLER / BUYER 有?
            'End If
            ' Run
            '
            ' Run (橡膠)
            If uRun Then
                '
                'PULLER GROUP
                xPuller = str
                i = Len(str)
                Do Until i = 0
                    If Mid(str, i, 1) >= "0" And Mid(str, i, 1) <= "9" Then
                        Exit Do
                    Else
                        xPuller = Mid(str, 1, i - 1)
                    End If
                    i = i - 1
                Loop
                If xPuller = "" Then xPuller = str
                'MsgBox("PULLER GR-1:" & xPuller & "/" & str)
                '
                xOPuller = xPuller
                sql = "SELECT Top 1 "
                sql = sql & "PullerGroup "
                sql = sql & "From M_EDXPullerGroup "
                sql = sql & "where Active = 1 "
                sql = sql & "and PullerList like '%" & "/" & xPuller & "/" & "%' "
                sql = sql & "order by PullerGroup "
                Dim dtPullerGr As DataTable = uDataBase.GetDataTable(sql)
                If dtPullerGr.Rows.Count > 0 Then
                    xPuller = dtPullerGr.Rows(0)("PullerGroup")
                    uPullerGroup = True
                Else
                    msg = msg & "[G]"
                End If
                'MsgBox("PULLER GR-2:" & xPuller)
                '
                'SPD PULLER / BUYER 有?
                sql = "SELECT Top 1 "
                sql = sql & "PullerCode, Buyer "
                sql = sql & "From M_ManufOutForEDX "
                '
                If Mid(xPuller, 1, 2) = "YG" Then
                    sql = sql & "Where PullerCode like 'YG%' "
                Else
                    sql = sql & "Where PullerCode like '" & xPuller & "%' "
                End If
                sql = sql & "And (Buyer = 'ALL' or Buyer like '%" & pBuyer & "%') "
                sql = sql & "order by SYS "
                '
                'MsgBox(sql)
                Dim dtSPDOut2 As DataTable = uDataBase.GetDataTable(sql)
                If dtSPDOut2.Rows.Count > 0 Then
                    '
                    'GET SEARCH KEY
                    If uPullerGroup Then
                        '有PullerGroup
                        If InStr(pSlider, xOPuller) > 0 Then
                            xPuller = Mid(pSlider, InStr(pSlider, xOPuller))
                        Else
                            msg = msg & "[K1]"
                        End If
                    Else
                        '無PullerGroup
                        'MsgBox(Mid(dtSPDOut2.Rows(0)("PullerCode"), 1, 2))
                        If Mid(dtSPDOut2.Rows(0)("PullerCode"), 1, 2) = "YG" Then
                            'YG%
                            'MsgBox(pSlider & ":" & Mid(dtSPDOut2.Rows(0)("PullerCode"), 1, 2))
                            If InStr(pSlider, Mid(dtSPDOut2.Rows(0)("PullerCode"), 1, 2)) > 0 Then
                                xPuller = Mid(pSlider, InStr(pSlider, Mid(dtSPDOut2.Rows(0)("PullerCode"), 1, 2)))
                            Else
                                msg = msg & "[K2]"
                            End If
                        Else
                            '一般
                            xCut = ""
                            str = xPuller
                            i = Len(str)
                            Do Until i = 0
                                If Mid(str, i, 1) >= "0" And Mid(str, i, 1) <= "9" Then
                                    If xCut = "" Then xCut = str
                                    Exit Do
                                Else
                                    xCut = Mid(str, 1, i - 1)
                                End If
                                i = i - 1
                            Loop
                            'MsgBox("CUT:" & xCut)
                            '
                            If xCut <> "" Then
                                If InStr(pSlider, xCut) > 0 Then
                                    xPuller = Mid(pSlider, InStr(pSlider, xCut))
                                Else
                                    msg = msg & "[K3]"
                                End If
                            Else
                                xPuller = str
                            End If
                            'MsgBox("CUTOVER:" & xPuller)
                        End If
                    End If
                    'GET SEARCH KEY
                Else
                    msg = msg & "[S]"
                End If
                'SPD PULLER / BUYER 有?
                '
                'RUN FLAG
                str = Replace(xPuller & msg, "[G]", "")
                'MsgBox("RUN:" & str & "/" & xPuller)
                If InStr(str, "[") > 0 And InStr(str, "]") > 0 Then uRunDB = False
            End If
            ' Run
            '
            ' Run CHECK WINGS & EDX DB
            If uRunDB Then
                str = xPuller & msg
                If InStr(str, "[") > 0 Then
                    str = Mid(str, 1, InStr(str, "[") - 1)
                End If
                sql = "SELECT Top 1 "
                sql = sql & "Slider "
                sql = sql & "From M_EDXSlider "
                sql = sql & "Where Slider LIKE '%" & str & "%' "
                sql = sql & "order by Slider "
                'MsgBox(sql)
                Dim dtWings As DataTable = uDataBase.GetDataTable(sql)
                If dtWings.Rows.Count > 0 Then
                    msg = msg & "[WY]"
                Else
                    msg = msg & "[WN]"
                End If
            End If
        End If
        '
        'Slider 
        If xPuller <> "" Then DSPDNO.Text = Mid(xPuller & msg, 1, 18)
        If xPuller <> "" Then DSPDNOUrl.Text = xPuller
        '--------------------------
        '
        If uError Then uJavaScript.PopMsg(Me, msg)
        '
        Return uError
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetPuller)
    '** 
    '**
    '*****************************************************************
    Public Function GetPuller(ByVal pSlider As String) As String
        Dim i As Integer
        Dim xPuller, xOPuller, xCut As String
        Dim sql, str As String
        Dim xSource, xTarget As String()
        Dim uError As Boolean = False
        Dim uRun As Boolean = False
        Dim uPullerGroup As Boolean = False
        '
        xPuller = ""
        xOPuller = ""
        str = ""
        'MsgBox(pSlider)
        '
        'Slider 
        If pSlider <> "" Then
            uRun = True
            '
            'SLIDER重新編成 (DA7B% & DAG7B)
            If Mid(pSlider, 1, 4) = "DA7B" Or Mid(pSlider, 1, 5) = "DAG7B" Or Mid(pSlider, 1, 4) = "DS15" Or Mid(pSlider, 1, 5) = "DAG24" Then
                'PULLER編成
                str = DSRequest1.Text & "/" & DSRequest2.Text & "/" & DSRequest3.Text & "/" & DSRequest4.Text & "/" & DSRequest5.Text & "/" & DSRequest6.Text & "/"
                xSource = Split(str, "/")
                For i = 0 To UBound(xSource) - 1
                    If Mid(xSource(i), 1, 3) = "PL-" Or Mid(xSource(i), 1, 5) = "T-PL-" Or Mid(xSource(i), 1, 5) = "B-PL-" Then
                        xPuller = Mid(xSource(i), InStr(xSource(i), "PL-") + 3)
                    End If
                Next
                '橡膠色編成
                xOPuller = pSlider
                If InStr(pSlider, "-") > 0 Then
                    xOPuller = Mid(pSlider, 1, InStr(pSlider, "-") - 1)
                End If
                If Mid(pSlider, 1, 4) = "DA7B" Then xOPuller = Replace(xOPuller, "DA7B", "")
                If Mid(pSlider, 1, 5) = "DAG7B" Then xOPuller = Replace(xOPuller, "DAG7B", "")
                If Mid(pSlider, 1, 4) = "DS15" Then xOPuller = Replace(xOPuller, "DS15", "")
                If Mid(pSlider, 1, 5) = "DAG24" Then xOPuller = Replace(xOPuller, "DAG24", "")
                'SLIDER編成
                str = ""
                If Mid(pSlider, 1, 4) = "DA7B" Then
                    If xPuller <> "" Then str = "DA7B" & xPuller
                    If xOPuller <> "" Then str = str & xOPuller
                End If
                If Mid(pSlider, 1, 5) = "DAG7B" Then
                    If xPuller <> "" Then str = "DAG7B" & xPuller
                    If xOPuller <> "" Then str = str & xOPuller
                End If
                If Mid(pSlider, 1, 4) = "DS15" Then
                    If xPuller <> "" Then str = "DS15" & xPuller
                    If xOPuller <> "" Then str = str & xOPuller
                End If
                If Mid(pSlider, 1, 5) = "DAG24" Then
                    If xPuller <> "" Then str = "DAG24" & xPuller
                    If xOPuller <> "" Then str = str & xOPuller
                End If
                If str <> "" Then pSlider = str
            End If
            'MsgBox("SLIDER:" & pSlider)
            '
            xPuller = ""
            xOPuller = ""
            str = ""
            '
            'REPLACE BODY 
            sql = "SELECT Top 1 "
            sql = sql & "Body, Source, Target "
            sql = sql & "From M_EDXBodyList "
            sql = sql & "where Active = 1 "
            sql = sql & "and '" & pSlider & "' like '' + body + '%' "
            sql = sql & "order by len(body) desc, body desc "
            Dim dtBody As DataTable = uDataBase.GetDataTable(sql)
            If dtBody.Rows.Count > 0 Then
                '
                str = Mid(pSlider, InStr(pSlider, dtBody.Rows(0)("Body")) + Len(dtBody.Rows(0)("Body")))
                '
                If dtBody.Rows(0)("Source") <> "" And dtBody.Rows(0)("Target") <> "" Then
                    xSource = Split(dtBody.Rows(0)("Source"), "/")
                    xTarget = Split(dtBody.Rows(0)("Target"), "/")
                    If UBound(xSource) = UBound(xTarget) Then
                        '
                        For i = 0 To UBound(xSource) - 1
                            If Mid(str, 1, Len(xSource(i))) = xSource(i) Then
                                str = Replace(str, xSource(i), xTarget(i))
                            End If
                        Next
                    End If
                End If
                '
                ' 只有 BODY
                If str = "" Then
                    str = dtBody.Rows(0)("Body")
                    uRun = False
                End If
                '
                xPuller = str
            Else
                uRun = False
            End If
            'MsgBox("BODY:" & str)
            '
            ' Run (橡膠)
            If uRun Then
                '
                'PULLER GROUP
                xPuller = str
                i = Len(str)
                Do Until i = 0
                    If Mid(str, i, 1) >= "0" And Mid(str, i, 1) <= "9" Then
                        Exit Do
                    Else
                        xPuller = Mid(str, 1, i - 1)
                    End If
                    i = i - 1
                Loop
                If xPuller = "" Then xPuller = str
                'MsgBox("PULLER GR-1:" & xPuller & "/" & str)
                '
                xOPuller = xPuller
                sql = "SELECT Top 1 "
                sql = sql & "PullerGroup "
                sql = sql & "From M_EDXPullerGroup "
                sql = sql & "where Active = 1 "
                sql = sql & "and PullerList like '%" & "/" & xPuller & "/" & "%' "
                sql = sql & "order by PullerGroup "
                Dim dtPullerGr As DataTable = uDataBase.GetDataTable(sql)
                If dtPullerGr.Rows.Count > 0 Then
                    xPuller = dtPullerGr.Rows(0)("PullerGroup")
                    uPullerGroup = True
                End If
                'MsgBox("PULLER GR-2:" & xPuller)
            End If
            ' Run
        End If
        '
        'Slider 
        If xPuller <> "" Then DSPDNOUrl.Text = xPuller
        '
        Return xPuller
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ErrorRule) 
    '**     檢測規則
    '**
    '*****************************************************************
    Public Function ErrorRule(ByVal pApply As String, ByVal pRule As String) As Boolean
        Dim uError As Boolean = True
        Dim uRuleError As Boolean
        Dim CheckMethod As String = ""
        Dim xSRequest, xRuleList As String()
        Dim xstr As String
        '
        'MsgBox(pApply)
        '
        If pRule = "*" Or pRule = "" Then uError = False
        '
        If uError = True Then
            '
            If InStr(pApply, "!") > 0 Then
                'M特別仕樣 (6 特別仕樣中有即可)
                If InStr(pRule, "!") > 0 Then CheckMethod = "NOTEQ"
                If InStr(pRule, "%") > 0 Then CheckMethod = "LIKE"
                '
                xSRequest = pApply.Split("!")
                For i As Integer = 0 To xSRequest.Length - 1
                    If xSRequest(i) <> "" Then
                        'MsgBox(xSRequest(i) & "--" & pRule)
                        Select Case CheckMethod
                            Case "NOTEQ"
                                If xSRequest(i) <> Replace(pRule, "!", "") Then uError = False
                                '
                            Case "LIKE"
                                xstr = pRule
                                xstr = Replace(xstr, "%", "*")
                                Select Case True
                                    Case xSRequest(i) Like xstr
                                        uError = False
                                    Case Else
                                End Select
                                '
                            Case Else   'EQ
                                If xSRequest(i) = pRule Then uError = False
                                '
                        End Select
                        '
                        If uError = False Then Exit For
                    End If
                Next
            Else
                '一般 (支援多個[AND]條件)
                uRuleError = False
                xRuleList = Split(pRule & "/", "/")
                For i As Integer = 0 To xRuleList.Length - 1
                    If xRuleList(i) <> "" Then
                        If InStr(xRuleList(i), "!") > 0 Then CheckMethod = "NOTEQ"
                        If InStr(xRuleList(i), "%") > 0 Then CheckMethod = "LIKE"
                        '
                        Select Case CheckMethod
                            Case "NOTEQ"
                                If pApply = Replace(xRuleList(i), "!", "") Then uRuleError = True
                                '
                            Case "LIKE"
                                xstr = xRuleList(i)
                                xstr = Replace(xstr, "%", "*")
                                Select Case True
                                    Case Not pApply Like xstr
                                        uRuleError = True
                                    Case Else
                                End Select
                                '
                            Case Else   'EQ
                                If pApply <> xRuleList(i) Then uRuleError = True
                                '
                        End Select
                        '
                        If uRuleError = True Then Exit For
                    End If
                Next
                uError = uRuleError
            End If
        End If
        '
        '-------------------------------------------------------
        'MODIFY-START BY 20230128
        '
        'Dim uError As Boolean = True
        'Dim CheckMethod As String = ""
        'Dim xSRequest As String()
        'Dim str As String
        ''
        ''MsgBox(pApply)
        ''
        'If pRule = "*" Then uError = False
        ''
        'If uError = True Then
        '    If InStr(pRule, "!") > 0 Then CheckMethod = "NOTEQ"
        '    If InStr(pRule, "%") > 0 Then CheckMethod = "LIKE"
        '    '
        '    If InStr(pApply, "!") > 0 Then
        '        '特別仕樣
        '        xSRequest = pApply.Split("!")
        '        For i As Integer = 0 To xSRequest.Length - 1
        '            If xSRequest(i) <> "" Then
        '                'MsgBox(xSRequest(i) & "--" & pRule)
        '                Select Case CheckMethod
        '                    Case "NOTEQ"
        '                        If xSRequest(i) <> Replace(pRule, "!", "") Then uError = False
        '                        '
        '                    Case "LIKE"
        '                        str = pRule
        '                        str = Replace(str, "%", "*")
        '                        Select Case True
        '                            Case xSRequest(i) Like str
        '                                uError = False
        '                            Case Else
        '                        End Select
        '                        '
        '                    Case Else   'EQ
        '                        If xSRequest(i) = pRule Then uError = False
        '                        '
        '                End Select
        '                '
        '                If uError = False Then Exit For
        '            End If
        '        Next
        '    Else
        '        '一般
        '        Select Case CheckMethod
        '            Case "NOTEQ"
        '                If pApply <> Replace(pRule, "!", "") Then uError = False
        '                '
        '            Case "LIKE"
        '                str = pRule
        '                str = Replace(str, "%", "*")
        '                Select Case True
        '                    Case pApply Like str
        '                        uError = False
        '                    Case Else
        '                End Select
        '                '
        '            Case Else   'EQ
        '                If pApply = pRule Then uError = False
        '                '
        '        End Select
        '    End If
        'End If
        '
        'MODIFY-END
        '-------------------------------------------------------
        'If uError = False Then
        'MsgBox("FASE:[" & CheckMethod & ":" & pApply & "=" & pRule & "]")
        'Else
        'MsgBox("TRUE:[" & CheckMethod & ":" & pApply & "=" & pRule & "]")
        'End If
        '
        Return uError
    End Function
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
        ''2024/07/08 DOrderDate改為客戶了解使用 (Emily)
        'ADD-START ISOS-2308 PJ
        '販賣實績-中止
        '
        'If Not uError Then
        '    'MsgBox("[" & DOrderDate.Text & "]:[" & HOrderDate.Text & "]")
        '    If IsDate(DOrderDate.Text) = True Then
        '        If DOrderDate.Text > HOrderDate.Text Or _
        '           DOrderDate.Text <= CDate(NowDateTime).ToString("yyyy/MM/dd") Then uError = True
        '    Else
        '        uError = True
        '    End If
        '    If uError Then uJavaScript.PopMsg(Me, "[" + DOrderDate.Text + "]:受注預定日 ?,請確認 ! ")
        'End If
        '
        'ADD-END
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

                'MODIFY-START ISOS-2308 PJ
                'If wFormSno = 0 And wStep < 3 Then    '判斷是否起單
                If wFormSno = 0 And wStep <= 3 Then    '判斷是否起單
                    'MODIFY-END

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

                'ADD-START ISOS-2308 PJ
                '違反者採用流程3 - 指定簽核者User ID
                If wStep = 3 Or wStep = 503 Then
                    ' xxMK057 林依儒 --> MK010 劉明德
                    ' sl057 李宜樺 --> MK014 吳松惠
                    ' tn006 李偉群 --> tn002 史文雄
                    ' td003 梁菀筑 --> MK010 劉明德
                    ' xx sl056 張家綱 --> MK014 吳松惠
                    ' xx mk038 游佳蓁 --> MK010 劉明德 --> MK003 蔡逸星
                    '
                    'tc006	蘇宗毅 --> tn002 史文雄
                    'mk058	莊若楨 --> MK010 劉明德 
                    '250208
                    'tc006	蘇宗毅	2	mk003	蔡逸星
                    'tc023	廖彥凱	1	tn002	史文雄
                    'mk006	周宇婕	1	mk010	劉明德
                    'sl044	張凱鈞	1	mk013	劉玉慧
                    '
                    Select Case Request.QueryString("pUserID")
                        Case "tc006"
                            wAllocateID = "mk003"
                            'Case "tc023"
                            '    wAllocateID = "tn002"
                            'Case "mk006"
                            '    wAllocateID = "mk010"
                            'Case "sl044"
                            '    wAllocateID = "mk013"
                        Case Else
                    End Select
                    '------
                    'Select Case Request.QueryString("pUserID")
                    '    Case "sl057"
                    '        wAllocateID = "mk014"
                    '    Case Else
                    'End Select
                    'Select Case Request.QueryString("pUserID")
                    '    Case "sl020"
                    '        wAllocateID = "sl015"
                    '    Case Else
                    'End Select
                End If
                'ADD-END

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
                'MODIFY-START ISOS-2308 PJ
                'If wFormSno = 0 And wStep < 3 Then    '判斷是否起單
                If wFormSno = 0 And wStep <= 3 Then    '判斷是否起單
                    'MODIFY-END

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
            ' 限定抓漏-START 2024/11/05
            DStatus.Value = "NG"
            ' 限定抓漏-END
            '
            If wbFormSno = 0 And wbStep < 3 Then    '判斷是否起單
                If DReRegister.Checked = True Then
                    uJavaScript.PopMsg(Me, "已成功送出您所填寫之申請單，請繼續申請！")
                    EnabledButton()                             '起動Button運作
                    DStatus.Value = "NG"
                    'ADD-START ISOS-2308 PJ
                    DSPDNOUrl.Text = ""
                    'ADD-END
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
        'MOD-START ISOS-2308 PJ
        'Dim sql As String = ""
        'sql = "INSERT INTO F_ItemRegisterSheet ( " & _
        '      "Sts, CompletedTime, FormNo, FormSno, " & _
        '      "No, Date, Name, JobTitle, Division, " & _
        '      "RCode, RItemName1, RItemName2, RItemName3, RSize, RChain, RClass, RTape, RSlider1, RFinish1, RSlider2, RFinish2, " & _
        '      "RSRequest1, RSRequest2, RSRequest3, RSRequest4, RSRequest5, RSRequest6, RFamily, " & _
        '      "RST1, RST2, RST3, RST4, RST5, RST6, RST7, RNoDisplay, " & _
        '      "Code, ItemName1, ItemName2, ItemName3, Size, Chain, Class, Tape, Slider1, Finish1, Slider2, Finish2, " & _
        '      "SRequest1, SRequest2, SRequest3, SRequest4, SRequest5, SRequest6, Family, " & _
        '      "ST1, ST2, ST3, ST4, ST5, ST6, ST7, NoDisplay, " & _
        '      "SLDPrice, Remark, Attachfile1, Attachfile2, A001, A206, A211, A999, K206, K211,Buyer,BuyerCode,Customer,Customercode,ForUse, SPDNo, SPDUrl, " & _
        '      "CreateUser, CreateTime, ModifyUser, ModifyTime) " & _
        '"VALUES("
        '追加 OrderDate
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
              "SLDPrice, Remark, Attachfile1, Attachfile2, A001, A206, A211, A999, K206, K211,Buyer,BuyerCode,Customer,Customercode,ForUse, OrderDate, SPDNo, SPDUrl, " & _
              "CreateUser, CreateTime, ModifyUser, ModifyTime) " & _
        "VALUES("
        '
        'MOD-END
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
                'FileName = CStr(NewFormSno) & "-" & "Attachfile1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DAttachfile1.PostedFile.FileName)
                FileName = CStr(NewFormSno) & "-" & "Attachfile1" & "-" & UploadDateTime & "-" & wApplyID & IO.Path.GetExtension(DAttachfile1.PostedFile.FileName) '避免錯誤檔名造成無法開啟，檔名改為年月日時分秒+UserID
                '*** IE8對應-End
                Try    '上傳圖檔
                    DAttachfile1.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "

        FileName = ""
        If DAttachfile2.Visible Then
            If DAttachfile2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Attachfile2" & "-" & UploadDateTime & "-" & Right(DAttachfile2.PostedFile.FileName, InStr(StrReverse(DAttachfile2.PostedFile.FileName), "\") - 1)
                'FileName = CStr(NewFormSno) & "-" & "Attachfile2" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DAttachfile2.PostedFile.FileName)
                FileName = CStr(NewFormSno) & "-" & "Attachfile2" & "-" & UploadDateTime & "-" & wApplyID & IO.Path.GetExtension(DAttachfile2.PostedFile.FileName)  '避免錯誤檔名造成無法開啟，檔名改為年月日時分秒+UserID
                '*** IE8對應-End
                Try    '上傳圖檔
                    DAttachfile2.PostedFile.SaveAs(Path & FileName)
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
        'ADD-START ISOS-2308 PJ
        'Order Date
        sql &= "'" & DOrderDate.Text & "' ,"
        '
        'ADD-END
        '
        sql &= "'" & DSPDNO.Text & "' ,"
        sql &= "'" & DSPDNOUrl.Text & "' ,"
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
                'FileName = CStr(wFormSno) & "-" & "Attachfile1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DAttachfile1.PostedFile.FileName)
                FileName = CStr(wFormSno) & "-" & "Attachfile1" & "-" & UploadDateTime & "-" & wApplyID & IO.Path.GetExtension(DAttachfile1.PostedFile.FileName) '避免錯誤檔名造成無法開啟，檔名改為年月日時分秒+UserID
                '*** IE8對應-End
                Try    '上傳圖檔
                    DAttachfile1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql &= " AttachFile1 = '" & FileName & "',"
            End If
        End If

        If DAttachfile2.Visible Then
            If DAttachfile2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Attachfile2" & "-" & UploadDateTime & "-" & Right(DAttachfile2.PostedFile.FileName, InStr(StrReverse(DAttachfile2.PostedFile.FileName), "\") - 1)
                'FileName = CStr(wFormSno) & "-" & "Attachfile2" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DAttachfile2.PostedFile.FileName)
                FileName = CStr(wFormSno) & "-" & "Attachfile2" & "-" & UploadDateTime & "-" & wApplyID & IO.Path.GetExtension(DAttachfile2.PostedFile.FileName) '避免錯誤檔名造成無法開啟，檔名改為年月日時分秒+UserID
                '*** IE8對應-End
                Try    '上傳圖檔
                    DAttachfile2.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql &= " Attachfile2 = '" & FileName & "',"
            End If
        End If


        'Jessia 20160128 Modify
        sql &= " Buyer = N'" & YKK.ReplaceString(DBuyer.Text) & "',"
        sql &= " BuyerCode = N'" & DBuyerCode.Text & "',"
        sql &= " Customer = N'" & YKK.ReplaceString(DCustomer.Text) & "',"
        sql &= " CustomerCode = N'" & DCustomerCode.Text & "',"
        sql &= " ForUse = N'" & DForUse.SelectedValue & "',"
        '
        'ADD-START ISOS-2308 PJ
        'Order Date
        sql &= " OrderDate = N'" & DOrderDate.Text & "',"
        '
        'ADD-END
        '
        sql &= " SPDNO = N'" & DSPDNO.Text & "',"
        sql &= " SPDURL = N'" & DSPDNOUrl.Text & "',"
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
            If wStep = 500 Then
            Else
                ErrCode = 9010
            End If
        End If

        ' 申請時才檢查販賣實蹟 2022/08/02
        If ErrCode = 0 Then

            'MODIFY-START ISOS-2308 PJ
            'If (wStep = 1 Or wStep = 500) And pAction = 0 Then
            If (wStep <= 3 Or wStep = 500 Or wStep = 503) And pAction = 0 Then
                'MODIFY-END

                '2024/07/08 DOrderDate改為客戶了解使用 (Emily)
                '檢查 預定受注日
                'If IsDate(DOrderDate.Text) = True Then
                '    If DOrderDate.Text > HOrderDate.Text Or DOrderDate.Text <= CDate(NowDateTime).ToString("yyyy/MM/dd") Then ErrCode = 9075
                'Else
                '    ErrCode = 9075
                'End If

                '客戶了解勾選檢查
                If DCustUnd.Checked Then
                    DOrderDate.Text = "確認客戶了解"
                Else
                    DOrderDate.Text = ""
                End If
                If DWarnMsg.Value <> "" And Trim(DOrderDate.Text) = "" Then
                    ErrCode = 8010
                End If

                If DSUITABLECHECK.Text = "" Then
                    ErrCode = 9073
                End If

                If ErrCode = 0 Then
                    '開發NO
                    If DSPDNO.Text = "" Then
                        ErrCode = 9074
                    End If
                End If
                If ErrCode = 0 Then
                    'MODIFY-START JOY 20250409
                    If DPerformaceUp.Text <> "IMPFINISH" Then
                        '
                        'EDX CHECK
                        If EDXNotFound("BUYER") = True Then
                            ErrCode = 9076
                        End If
                        '
                    End If
                    'MODIFY-END
                End If
                '
                'ADD-START 20250527
                '檢測ITEM時DESCRIPTION & 確定申請時DESCRIPTION 是否相同
                If ErrCode = 0 Then
                    DNewItemInf.Text = UCase(DSize.Text.Trim) & "!" & _
                                       UCase(DChain.Text.Trim) & "!" & _
                                       UCase(DClass.Text.Trim) & "!" & _
                                       UCase(DTape.Text.Trim) & "!" & _
                                       UCase(DSlider1.Text.Trim) & "!" & _
                                       UCase(DFinish1.Text.Trim) & "!" & _
                                       UCase(DSlider2.Text.Trim) & "!" & _
                                       UCase(DFinish2.Text.Trim) & "!" & _
                                       UCase(DFamily.Text.Trim) & "!" & _
                                       UCase(DSRequest1.Text.Trim) & "!" & _
                                       UCase(DSRequest2.Text.Trim) & "!" & _
                                       UCase(DSRequest3.Text.Trim) & "!" & _
                                       UCase(DSRequest4.Text.Trim) & "!" & _
                                       UCase(DSRequest5.Text.Trim) & "!" & _
                                       UCase(DSRequest6.Text.Trim) & "!"
                    DNewItemInf.Text = Replace(DNewItemInf.Text, "XXX!", "")
                    '
                    Dim str1 As String = UCase(DItemName1.Text.Trim) & UCase(DItemName2.Text.Trim) & UCase(DItemName3.Text.Trim)
                    If DNewItemInf.Text <> DItemInf.Text Or _
                       (UCase(DTape.Text.Trim) <> "" And InStr(str1, UCase(DTape.Text.Trim)) <= 0) Or _
                       (UCase(DSlider1.Text.Trim) <> "" And InStr(str1, UCase(DSlider1.Text.Trim)) <= 0) Or _
                       (UCase(DFinish1.Text.Trim) <> "" And InStr(str1, UCase(DFinish1.Text.Trim)) <= 0) Or _
                       (UCase(DSlider2.Text.Trim) <> "" And InStr(str1, UCase(DSlider2.Text.Trim)) <= 0) Or _
                       (UCase(DFinish2.Text.Trim) <> "" And InStr(str1, UCase(DFinish2.Text.Trim)) <= 0) Or _
                       (UCase(DSRequest1.Text.Trim) <> "" And InStr(str1, UCase(DSRequest1.Text.Trim)) <= 0) Or _
                       (UCase(DSRequest2.Text.Trim) <> "" And InStr(str1, UCase(DSRequest2.Text.Trim)) <= 0) Or _
                       (UCase(DSRequest3.Text.Trim) <> "" And InStr(str1, UCase(DSRequest3.Text.Trim)) <= 0) Or _
                       (UCase(DSRequest4.Text.Trim) <> "" And InStr(str1, UCase(DSRequest4.Text.Trim)) <= 0) Or _
                       (UCase(DSRequest5.Text.Trim) <> "" And InStr(str1, UCase(DSRequest5.Text.Trim)) <= 0) Or _
                       (UCase(DSRequest6.Text.Trim) <> "" And InStr(str1, UCase(DSRequest6.Text.Trim)) <= 0) Then
                        '
                        ErrCode = 9077
                    End If
                    '
                End If
                'ADD-END
                '
            End If
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
                Else
                    If DProdReady.Text = "NG" Then
                        ErrCode = 9035
                    End If
                End If
            End If
        End If
        '
        'TEST-START
        'MsgBox("BEFORE=[" & CStr(ErrCode) & "]")
        'ErrCode = 9099
        'MsgBox("AFTER-9099=[" & CStr(ErrCode) & "]")
        'TEST-END
        '
        If ErrCode = 0 Then
            If DAttachfile2.Visible Then
                If DAttachfile2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DAttachfile2)
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

        '檢查是否重複 JOO
        If wFormSno > 0 And wStep > 3 Then    '判斷是否[簽核]
        Else
            Dim SQL As String = "Select No From F_ItemRegisterSheet " & _
                                "Where (Sts='0' or Sts='1') " & _
                                "  And ItemName1 = '" & DItemName1.Text.Trim & "' " & _
                                "  And ItemName2 = '" & DItemName2.Text.Trim & "' " & _
                                "  And ItemName3 = '" & DItemName3.Text.Trim & "' "
            Dim dtItemRegister As DataTable = uDataBase.GetDataTable(SQL)
            If dtItemRegister.Rows.Count > 0 Then
                ErrCode = 9099
            End If
        End If
        '
        If ErrCode <> 0 Then
            If ErrCode = 8010 Then Message = "未勾選客戶了解，請確認!"
            If ErrCode = 9010 Then Message = "Item資料有誤,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案格式有誤,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案Size有誤,請確認!"
            If ErrCode = 9035 Then Message = "需上傳測試報告或可生產確認信,請確認!"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9050 Then Message = "委託書No.重覆,請確認委託書No.!"
            If ErrCode = 9060 Then Message = "修改後未重新檢測資料,請確認!"
            If ErrCode = 9070 Then Message = "發現有空白重覆,請重新檢測資料,請確認!"
            If ErrCode = 9071 Then Message = "Item Name(2)字串過長(>34),請重新檢測資料,請確認!"
            If ErrCode = 9072 Then Message = "發現特殊要求中有未排序資料,請重新檢測資料,請確認!"
            If ErrCode = 9073 Then Message = "需搜尋販賣實績!"
            If ErrCode = 9074 Then Message = "需附上開發委託書NO!"
            If ErrCode = 9075 Then Message = "受注預定日非有效日期或超過 [" & HOrderDate.Text & "],請確認 ! "
            If ErrCode = 9076 Then Message = "搜尋不到[EDX]，請確認是否已完成 ! "
            If ErrCode = 9077 Then Message = "Item細項內容與[Item Name]不符合 ! "
            If ErrCode = 9099 Then Message = "[ItemRegister-1101]:此Item資料已經申請,請確認 ! "
            If ErrCode = 9999 Then Message = "IT TEST ! "
            uJavaScript.PopMsg(Me, Message)
        Else
            isOK = True
        End If
        '
        '測試使用
        'isOK = False
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
        Dim allowedExtensions As String() = {".png", ".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".ai", ".xls", ".doc", ".ppt", ".xlsx"}   '定義允許的檔案格式
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
            If UPFile.PostedFile.ContentLength <= 5000 * 1024 Then
                UPFileIsNormal = 0
            Else
                UPFileIsNormal = 9030
            End If
        End If
    End Function

    Sub CheckItem()
        Dim str As String
        Dim j As String
        'DSPDNO.Text = ""
        DSPDNO.Text = "NOT-WINGSITEM"
        DSPDNOUrl.Text = ""
        '
        '生產NG
        Dim ColorStr, KeepStr, xSize, p1, p2 As String

        If CDbl(DSize.Text) > 9 Then
            xSize = DSize.Text
        Else
            xSize = Replace(DSize.Text, "0", "")
        End If
        '
        If DSlider1.Text <> "" Then
            KeepStr = ""
            ColorStr = ""
            str = DSlider1.Text
            j = Len(DSlider1.Text)
            '**
            Do Until j < 0 Or ColorStr = DSlider1.Text
                '** 

                If Mid(str, j, 1) >= "0" And Mid(str, j, 1) <= "9" Then
                    KeepStr = Mid(str, 1, j)
                    Exit Do
                Else
                    ColorStr = Mid(str, j, 1) & ColorStr
                End If
                j = j - 1
            Loop
            '**
            If ColorStr = DSlider1.Text Then KeepStr = ColorStr
            '**
            p1 = xSize & "-" & DFamily.Text & "-" & KeepStr
            p2 = xSize & "-" & DFamily.Text & "-" & DSlider1.Text
            '
            '2022/8/5 ADD-START
            p1 = Replace(p1, "5-CI", "5-CN")
            p1 = Replace(p1, "5-CZ", "5-CN")
            p1 = Replace(p1, "5-CS", "5-CN")
            p1 = Replace(p1, "5-Y", "5-M")
            '
            p1 = Replace(p1, "3-CZ", "3-C")
            p1 = Replace(p1, "3-CS", "3-C")
            'ADD-END
            '
            '
            'MODIFY-START LIMITEDITEM
            'BSPDNO.Attributes("onclick") = "window.open('" & _
            '   "FindSPDPage.aspx?pUserID=" & wApplyID & "&pSLDKey1=" & p2 & "&pSLDKey2=" & DSlider1.Text & "&pFINKey1=" & p1 & "&pFINKey2=" & KeepStr & _
            '   "','FindSPDPage','resizable=yes,scrollbars=yes');"
            'MODIFY-END
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**     設定限定檢查的連結
    '**
    '*****************************************************************
    Sub setCkLimitedLink(ByVal id As String)
        CkLimitedLink.NavigateUrl = "/IRW/CkLimitedPage.aspx?id=" & id & "&msg="
        If Trim(DSLDPrice.Text) = "NEW" Then
            CkLimitedLink.NavigateUrl = CkLimitedLink.NavigateUrl & "NEW"
        Else
            CkLimitedLink.Visible = True
        End If
        '開啟限定檢查的頁面
        Response.Write("<script>window.open('" & CkLimitedLink.NavigateUrl & "','','height=650,width=440,menubar=no,location=no,resizable=yes,scrollbars=yes');</script>")
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**     判斷若特殊要求有開發中產品keyword，則取消勾選單價A001/A999
    '**     
    '*****************************************************************
    Sub SetDevPrice()
        Dim xSpc As String = GetFullName(DSRequest1.Text) & "!" & _
                       GetFullName(DSRequest2.Text) & "!" & _
                       GetFullName(DSRequest3.Text) & "!" & _
                       GetFullName(DSRequest4.Text) & "!" & _
                       GetFullName(DSRequest5.Text) & "!" & _
                       GetFullName(DSRequest6.Text) & "!"
        '若為開發中產品(SLS-ANY、CH-ANY、TAPE-ANY、ILYTEMP、I-YTEMP、ILTP***、I-TP***)不可勾選A001或A999
        If InStr(xSpc, "SLS-ANY") > 0 Or InStr(xSpc, "CH-ANY") > 0 Or InStr(xSpc, "TAPE-ANY") > 0 Or InStr(xSpc, "ILYTEMP") > 0 _
            Or InStr(xSpc, "I-YTEMP") > 0 Or InStr(xSpc, "ILTP") > 0 Or InStr(xSpc, "I-TP") > 0 Then
            DA001.Checked = False
            DA999.Checked = False
            DA001.Enabled = False
            DA999.Enabled = False
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '** 限定-恢復原客戶及BUYER
    '**     
    '*****************************************************************
    Protected Sub BLoadCustomerBuyer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BLoadCustomerBuyer.Click

        If HCustomerCode.Text <> "" Or HBuyerCode.Text <> "" Then
            '--
            '-- 限定
            DCustomer.Text = HCustomer.Text
            DBuyer.Text = HBuyer.Text
            DCustomerCode.Text = HCustomerCode.Text
            DBuyerCode.Text = HBuyerCode.Text
            '--
            '-- BUYER CHECK
            Dim SQL As String
            DMemo.Text = ""
            DA001.Enabled = True
            DA999.Enabled = True
            If DBuyer.Text <> "" Then
                If DST4.Text = "1" Then
                    SQL = "SELECT  * FROM m_REFERP"
                    SQL = SQL & " WHERE CAT = '1151'"
                    SQL = SQL & " AND DKEY  in ('A999','A001')"
                    SQL = SQL & " AND DATA ='" + DBuyerCode.Text + "'"
                    Dim dtPrice As DataTable = uDataBase.GetDataTable(SQL)
                    If dtPrice.Rows.Count > 0 Then
                        If dtPrice.Rows(0)("DKEY").ToString = "A999" Then
                            DA999.Checked = True
                            DA001.Checked = False
                        ElseIf dtPrice.Rows(0)("DKEY").ToString = "A001" Then
                            DA001.Checked = True
                            DA999.Checked = False
                        End If

                        DMemo.Text = "請確認若為開發中產品(SLS-ANY、CH-ANY、TAPE-ANY、ILYTEMP、I-YTEMP、ILTP***、I-TP***)不可勾選A001或A999。"
                        '判斷若特殊要求有開發中產品keyword，則取消勾選單價A001/A999
                        SetDevPrice()
                    Else
                        DA999.Checked = False
                        DA001.Checked = False
                    End If
                Else
                    DA999.Checked = False
                    DA001.Checked = False
                    '統計區分4<>1，不可勾選單價A001/A999
                    DA999.Enabled = False
                    DA001.Enabled = False

                    DMemo.Text = ""
                End If
            End If
            '
            'ADD-START 2311-BUYER&ITEM限定
            If LimitedItemError("BUYER") = True Then
                DStatus.Value = "NG"
                DCustomer.Text = ""
                DBuyer.Text = ""
                DCustomerCode.Text = ""
                DBuyerCode.Text = ""
                D1.Text = ""
                D2.Text = ""
            Else
                DStatus.Value = "OK"
            End If
            'ADD-END
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '** 特別要求合併及縮寫
    '**     
    '*****************************************************************
    Protected Sub BSpecialShort_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSpecialShort.Click
        Dim sql, xID As String
        Dim Cmd As String
        '
        xID = DateTime.Now.ToString("yyyyMMddHHmmss")
        '
        ' CLEAR之前處理過或已過期DATA
        sql = "update M_IRWSpecialShort "
        sql = sql & "set Active = 0 "
        sql = sql & "where CreateUser = '" & Request.QueryString("pUserID") & "' "
        uDataBase.ExecuteNonQuery(sql)
        '
        ' 新需求寫入I/F TABLE
        sql = "Insert into M_IRWSpecialShort ( "
        sql = sql & "ID, Active, Special1,Special2,Special3,Special4,Special5,Special6, "
        sql = sql & "CreateUser, CreateTime, ModifyUser, ModifyTime "
        sql = sql & ") "
        sql = sql & "VALUES( "
        sql = sql & "'" & xID & "', "
        sql = sql & " " & "1" & ",  "
        sql = sql & "'" & DSRequest1.Text & "', "
        sql = sql & "'" & DSRequest2.Text & "', "
        sql = sql & "'" & DSRequest3.Text & "', "
        sql = sql & "'" & DSRequest4.Text & "', "
        sql = sql & "'" & DSRequest5.Text & "', "
        sql = sql & "'" & DSRequest6.Text & "', "
        '
        sql &= "'" & Request.QueryString("pUserID") & "' ,"
        sql &= "'" & NowDateTime & "' ,"
        sql &= "'" & Request.QueryString("pUserID") & "' ,"
        sql &= "'" & NowDateTime & "' )"
        uDataBase.ExecuteNonQuery(sql)
        '
        ' CALL EXCEL
        Cmd = "<script>" + _
                    "window.open('http://10.245.0.205/EDI/SPHttp2File.aspx?pFun=SRSHORT&pUserID=" & Trim(UCase(Request.QueryString("pUserID"))) & "','SpecialShort','');" + _
              "</script>"
        Response.Write(Cmd)
        '
    End Sub

End Class

