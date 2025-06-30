Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DTMW_NewColorSheet01_UAKIPLING
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
    Dim wbFormSno As Integer        '連續起單-表單流水號
    Dim wbStep As Integer           '連續起單-工程代碼
    Dim wStep As Integer            '工程代碼
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID
    Dim wUserID As String          '使用者ID
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
    Dim oWaves As New WAVES.CommonService
    Dim aAgentID As String
    Dim Makemap1, Makemap2, Makemap3, Makemap4, Makemap5, Makemap6 As Integer
    Dim AQty, DevNo, Manufout As String
    Dim isOK As Boolean = True
    Dim Message As String = ""
    Dim LastStep As String






    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置
        SetControlPosition()    ' 設定控制項位置

        If Not Me.IsPostBack Then   '不是PostBack
            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()     '表單功能按鈕顯示

            If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
                ShowFormData()      '顯示表單資料
                UpdateTranFile()    '更新交易資料
                TopPosition()
                SetControlPosition()    ' 設定控制項位置

            End If
            SetPopupFunction()      '設定彈出視窗事件



        Else
            ShowSheetFunction()     '表單功能按鈕顯示
            ' fpObj.GetDeafultNextGate(DNextGate, DDivision.Text,wUserID) ' 設定預設的簽核者
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
        wbFormSno = Request.QueryString("pFormSno")  '連續起單用流水號
        wbStep = Request.QueryString("pStep")        '連續起單用工程代碼

        wSeqNo = Request.QueryString("pSeqNo")      '序號
        wApplyID = Request.QueryString("pApplyID")  '申請者ID

        wAgentID = Request.QueryString("pAgentID")  '被代理人ID
        'Add-Start by Joy 2009/11/20(2010行事曆對應)
        '申請者-群組行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        If wStep = 10 Or wStep = 510 Or wStep = 20 Or wStep = 520 Or wStep = 520 Or wStep = 530 Then '指定徐慧芬
            Dim SQL As String

            SQL = "Select * From T_waitHandle "
            SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & " and step ='" + CStr(wStep) + "'"
            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
            If DBAdapter1.Rows.Count > 0 Then
                wUserID = DBAdapter1.Rows(0).Item("Workid")
            End If

        Else
            wUserID = Request.QueryString("pUserID")
        End If

        'wUserID = Request.QueryString("UserID")
        wDecideCalendar = oCommon.GetCalendarGroup(wUserID)

        Response.Cookies("PGM").Value = "DTMW_NewColorSheet01_03CFP12.aspx"
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
        'oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDepo,wUserID)
        '
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, wUserID)
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

        SQL = "Select * From F_NewColorUAKIPLING "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then


            If DBAdapter1.Rows(0).Item("CustomerCheck") = 1 Then '客戶自核
                DCustomerCheck.Checked = True
                If (wStep = 10) Or (wStep = 510) Then
                    BOK.Visible = True   ' 跳到20
                    BNG2.Visible = True
                End If
            End If

            If DBAdapter1.Rows(0).Item("FactoryCheck") = 1 Then '工廠自核
                DFactoryCheck.Checked = True
                If (wStep = 10) Or (wStep = 510) Then
                    BOK.Visible = False
                    BNG2.Visible = True   ' 跳到40
                End If
            End If

            If DBAdapter1.Rows(0).Item("VCACheck") = 1 Then '工廠自核
                DVCACheck.Checked = True
                If (wStep = 10) Or (wStep = 510) Then
                    BOK.Visible = False
                    BNG2.Visible = True   ' 跳到40
                End If
            End If


            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            '   DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
            DDate.Text = DBAdapter1.Rows(0).Item("Date")                         'No 
            DDepoName.Text = DBAdapter1.Rows(0).Item("DepoName")                         'No
            DName.Text = DBAdapter1.Rows(0).Item("Name")                         'No
            DCustomer.Text = DBAdapter1.Rows(0).Item("Customer")                         'No
            DCustomerCode.Text = DBAdapter1.Rows(0).Item("CustomerCode")                         'No
            DBuyer.Text = DBAdapter1.Rows(0).Item("Buyer")                         'No
            DBuyerCode.Text = DBAdapter1.Rows(0).Item("BuyerCode")
            DCustomerColor.Text = DBAdapter1.Rows(0).Item("CustomerColor")                         'No3
            DCustomerColorCode.Text = DBAdapter1.Rows(0).Item("CustomerColorCode")
            DOverseaYKKCode.Text = DBAdapter1.Rows(0).Item("OverseaYKKCode")
            DPANTONECode.Text = DBAdapter1.Rows(0).Item("PANTONECode")
            SetFieldData("DevYear", DBAdapter1.Rows(0).Item("DevYear"))    '年
            SetFieldData("DevSeason", DBAdapter1.Rows(0).Item("DevSeason"))    '季

            If Mid(DBAdapter1.Rows(0).Item("ReceiveDate").ToString, 1, 4) = "1900" Then
                DReceiveDate.Text = ""
            Else
                DReceiveDate.Text = DBAdapter1.Rows(0).Item("ReceiveDate")               '下單時間
            End If

            SetFieldData("ColorLight1", DBAdapter1.Rows(0).Item("ColorLight1"))    '類別1
            SetFieldData("ColorLight2", DBAdapter1.Rows(0).Item("ColorLight2"))    '類別2
            SetFieldData("ColorLight3", DBAdapter1.Rows(0).Item("ColorLight3"))    '類別2

            SetFieldData("NewOldColor", DBAdapter1.Rows(0).Item("NewOldColor"))    '類別2

            If Mid(DBAdapter1.Rows(0).Item("DeliveryDate").ToString, 1, 4) = "1900" Then
                DDeliveryDate.Value = ""
            Else
                DDeliveryDate.Value = DBAdapter1.Rows(0).Item("DeliveryDate")               '下單時間
            End If

            If wStep = 500 Then  '(NG再送出希望交期需重新選擇)
                DDeliveryDate.Value = ""
            End If


            DDuplicateNo.Text = DBAdapter1.Rows(0).Item("DuplicateNo")
            If DBAdapter1.Rows(0).Item("NOCCS") = 1 Then
                DNOCCS.Checked = True
            End If
            DNOCCSReason.Text = DBAdapter1.Rows(0).Item("NOCCSReason")
            DCustomerNGColor.Text = DBAdapter1.Rows(0).Item("CustomerNGColor")
            DRemark.Text = DBAdapter1.Rows(0).Item("Remark")
            SetFieldData("CustomerSample", DBAdapter1.Rows(0).Item("CustomerSample"))    '類別2

            DColorSystem.Text = DBAdapter1.Rows(0).Item("ColorSystem")

            DCheckNo.Text = DBAdapter1.Rows(0).Item("CheckNo")
            DPFBWire.Value = DBAdapter1.Rows(0).Item("PFBWire")
            DPFOpenParts.Value = DBAdapter1.Rows(0).Item("PFOpenParts")

            SetFieldData("YKKColorType", DBAdapter1.Rows(0).Item("YKKColorType"))    '類別2

            DYKKColorCode.Value = DBAdapter1.Rows(0).Item("YKKColorCode")

            DYKKColorCodeVF.Value = DBAdapter1.Rows(0).Item("YKKColorCodeVF")
            DYKKColorCodeSLD.Value = DBAdapter1.Rows(0).Item("YKKColorCodeSLD")

            If wStep = 70 Then
                If DBAdapter1.Rows(0).Item("DTNO") = 0 Then

                    DYKKColorCodeVF.Value = DBAdapter1.Rows(0).Item("VFColor")
                    DYKKColorCodeSLD.Value = DBAdapter1.Rows(0).Item("SLDColor")
                End If

                If DYKKColorCodeVF.Value = "" Then
             
                    DYKKColorCodeVF.Value = DBAdapter1.Rows(0).Item("VFColor")
                End If

                If DYKKColorCodeSLD.Value = "" Then
                
                    DYKKColorCodeSLD.Value = DBAdapter1.Rows(0).Item("SLDColor")
                End If




            End If



            Dim Version As Integer
            Dim Formname As String

            If wStep = 20 Or wStep = 520 Then
                Formname = "3CF DTM"
            Else
                Formname = "5VS  DTM"
            End If

            '計算是第幾版         
            SQL = " select  count(*)cun  from f_newcolorcomplete "
            SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And FormName =  '" & Formname & "'"

            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
            If DBAdapter3.Rows.Count > 0 Then
                Version = DBAdapter3.Rows(0).Item("cun")
            Else
                Version = 1
            End If

            If wStep = 520 Or wStep = 530 Then
                Version = Version + 1
            End If



            DVersion.Text = Version
            DSLDColor.Text = DBAdapter1.Rows(0).Item("SLDColor")
            SetFieldData("ColorType", DBAdapter1.Rows(0).Item("again"))    '濃淡色
            If DBAdapter1.Rows(0).Item("again") = 1 Then
                DAgain.SelectedValue = "淡色"
            ElseIf DBAdapter1.Rows(0).Item("again") = 2 Then
                DAgain.SelectedValue = "濃色"
            End If

            DVFColor.Text = DBAdapter1.Rows(0).Item("VFColor")
            DVFColor1.Text = DBAdapter1.Rows(0).Item("VFColor1")


            If DBAdapter1.Rows(0).Item("Complete") = 1 Then       '有新色依賴完成表
                LComplete.NavigateUrl = "NewColoCompleteList.aspx?pNo=" & DNo.Text
                LComplete.Visible = True
            Else
                LComplete.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("DTNO") = 1 Then       '有追加核可單
                LDTSheet.NavigateUrl = "NewColorDTSheetList.aspx?pNo=" & DNo.Text
                LDTSheet.Visible = True
            Else
                LDTSheet.Visible = False
            End If

        End If



        If wStep = 80 Then  '強迫變成空白最後輸入色塊
            DSLDColor.Text = ""
            DAgain.SelectedValue = ""
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




        '核定履歷資料
        SQL = "SELECT "
        SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
        SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
        SQL = SQL + "AEndTimeDesc As Description, "
        SQL = SQL + "URL "
        SQL = SQL + "FROM V_WaitHandle_01 "
        SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
        SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
        ' SQL = SQL + "Order by Unique_ID Desc "
        SQL = SQL + " Order by CreateTime Desc, Step Desc, SeqNo Desc "
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
                    Top = 1200
                Else
                    DReasonCode.Visible = False     '延遲理由代碼
                    DReason.Visible = False         '延遲理由
                    DReasonDesc.Visible = False     '延遲其他說明
                    If DBAdapter2.Rows(0).Item("Flowtype") = "1" Then
                        Top = 1120
                    Else
                        Top = 1200
                    End If

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
            Top = 1100
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
        BDate.Attributes.Add("onClick", "calendarPicker('DDeliveryDate')")
        BCustomer.Attributes.Add("onClick", "GetCustomer()") '找客戶
        BBuyer.Attributes.Add("onClick", "GetBuyer()") '找buyer
        BCopySheet.Attributes.Add("onClick", "CopyNewColor('" + wFormNo + "')") '找buyer

        BYKKColor.Attributes.Add("onClick", "GetYKKColor()") '找buyer
        BYKKColorCode.Attributes.Add("onClick", "GetYKKColorCode('DYKKColorCode')") '找buyer
        BYKKColorCodeVF.Attributes.Add("onClick", "GetYKKColorCode('DYKKColorCodeVF')") '找buyer
        BYKKColorCodeSLD.Attributes.Add("onClick", "GetYKKColorCode('DYKKColorCodeSLD')") '找buyer

        Me.DColorSystem.Attributes.CssStyle.Add("text-transform", "uppercase")

        BSLDColor.Attributes.Add("onClick", "CheckColor('" + DYKKColorCode.Value + "')") '找buyer


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
                    Top = 1200
                Else
                    If DDelay.Visible = True Then
                        Top = 1200
                    Else
                        Top = 1200
                    End If
                End If
            End If
        Else
            Top = 1600

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



        If wFormNo = "005011" Then
            LSheetName.Text = "拉鏈ZIPPER CHAIN (UA＆KIPLING)" '表單名稱
        End If



        Dim SQL As String
        Dim InputData1 As String
        Dim InputData2 As String


        InputData1 = D1.Text
        InputData2 = D2.Text


        '帶入CUSTOMER
        If InputData1 <> "" Then

            SQL = " Select  * from  MST_Custmer where 1=1 "
            If InputData1 <> "" Then
                SQL = SQL + " and ( custmer like '%" + InputData1 + "%' or name_c like '%" + InputData1 + "%')"
            End If
            SQL = SQL + " order by custmer,name_c "

            Dim DBData As DataTable = uDataBase.GetDataTable(SQL)

            DCustomer.Text = DBData.Rows(0).Item("Name_C")
            DCustomerCode.Text = DBData.Rows(0).Item("Custmer")
        End If

        '帶入BUYER
        If InputData2 <> "" Then


            SQL = " Select  * from  MST_Buyer where 1=1 "
            If InputData2 <> "" Then
                SQL = SQL + " and ( buyer like '%" + InputData2 + "%' or buyer_name like '%" + InputData2 + "%')"
            End If

            SQL = SQL + " order by buyer_name,buyer "
            Dim DBData As DataTable = uDataBase.GetDataTable(SQL)

            DBuyer.Text = DBData.Rows(0).Item("Buyer_Name")
            DBuyerCode.Text = DBData.Rows(0).Item("Buyer")

        End If

        'jessica 20221202 DTMW 色號檢查
        '帶入SLDCOLOR
        If D3.Text <> "" Then
            DSLDColor.Text = D3.Text
            DYKKColorCodeSLD.Value = D3.Text
        End If


        'Modify 20150417
        '帶入舊色
        If wStep = 70 Then
            If DNewOldColor.SelectedValue = "舊色" Then

                SQL = " Select  ltrim(CPS1XX)CPS1XX,ltrim(CPT1XX)CPT1XX from   MST_color_structure where  paccxx ='" + DYKKColorCode.Value + "'"
                SQL = SQL + " and ctbnxx ='1'"


                Dim DBData As DataTable = uDataBase.GetDataTable(SQL)
                If DBData.Rows.Count > 0 Then
                    DSLDColor.Text = DBData.Rows(0).Item("CPS1XX")
                    DVFColor.Text = DBData.Rows(0).Item("CPT1XX")
                End If
            Else
                DSLDColor.Text = ""
                DVFColor.Text = ""
            End If
        End If




        '客戶自核
        Select Case FindFieldInf("CustomerCheck")
            Case 0  '顯示
                ' DCustomerCheck.BackColor = Color.LightGray
                DCustomerCheck.Enabled = False

            Case 1  '修改+檢查
                DCustomerCheck.BackColor = Color.GreenYellow
                '  ShowRequiredFieldValidator("DCustomerCheckRqd", "DCustomerCheck", "異常：需輸入客戶自核")

            Case 2  '修改
                DCustomerCheck.BackColor = Color.Yellow

            Case Else   '隱藏
                DCustomerCheck.Visible = False
        End Select
        If pPost = "New" Then DCustomerCheck.Checked = True

        '工廠自核
        Select Case FindFieldInf("FactoryCheck")
            Case 0  '顯示
                '    DCustomerCheck.BackColor = Color.LightGray
                DFactoryCheck.Enabled = False

            Case 1  '修改+檢查
                DFactoryCheck.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFactoryCheckRqd", "DFactoryCheck", "異常：需輸入工廠自核")

            Case 2  '修改
                DFactoryCheck.BackColor = Color.Yellow

            Case Else   '隱藏
                DFactoryCheck.Visible = False
        End Select
        If pPost = "New" Then DFactoryCheck.Checked = False


        'VCA核
        Select Case FindFieldInf("VCACheck")
            Case 0  '顯示
                ' DCustomerCheck.BackColor = Color.LightGray
                DVCACheck.Enabled = False

            Case 1  '修改+檢查
                DVCACheck.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVCACheckRqd", "VCACheck", "異常：需輸入VCA")

            Case 2  '修改
                DVCACheck.BackColor = Color.Yellow

            Case Else   '隱藏
                DVCACheck.Visible = False
        End Select
        If pPost = "New" Then DVCACheck.Checked = False



        'VCA核
        Select Case FindFieldInf("NOCCS")
            Case 0  '顯示
                ' DCustomerCheck.BackColor = Color.LightGray
                DNOCCS.Enabled = False

            Case 1  '修改+檢查
                DNOCCS.BackColor = Color.GreenYellow
                '      ShowRequiredFieldValidator("DNOCCSRqd", "NOCCS", "異常：需輸入無法CCS")

            Case 2  '修改
                DNOCCS.BackColor = Color.Yellow

            Case Else   '隱藏
                DNOCCS.Visible = False
        End Select
        If pPost = "New" Then DNOCCS.Checked = False


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
        Select Case FindFieldInf("Date")
            Case 0  '顯示
                DDate.BackColor = Color.LightGray
                DDate.ReadOnly = True
                DDate.Visible = True

            Case 1  '修改+檢查
                DDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDateRqd", "DDate", "異常：需輸入日期")
                DDate.Visible = True

            Case 2  '修改
                DDate.BackColor = Color.Yellow
                DDate.Visible = True

            Case Else   '隱藏
                DDate.Visible = False

        End Select
        If pPost = "New" Then DDate.Text = Now.ToString("yyyy/MM/dd") '現在日時

        'DepoName
        Select Case FindFieldInf("DepoName")
            Case 0  '顯示
                DDepoName.BackColor = Color.LightGray
                DDepoName.ReadOnly = True
                DDepoName.Visible = True
            Case 1  '修改+檢查
                DDepoName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDepoNameRqd", "DDepoName", "異常：需輸入部門")
                DDepoName.Visible = True
            Case 2  '修改
                DDepoName.BackColor = Color.Yellow
                DDepoName.Visible = True
            Case Else   '隱藏
                DDepoName.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DepoName", "ZZZZZZ")


        'Name
        Select Case FindFieldInf("Name")
            Case 0  '顯示
                DName.BackColor = Color.LightGray
                DName.ReadOnly = True
                DName.Visible = True
            Case 1  '修改+檢查
                DName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNameRqd", "DName", "異常：需輸入姓名")
                DName.Visible = True
            Case 2  '修改
                DName.BackColor = Color.Yellow
                DName.Visible = True
            Case Else   '隱藏
                DName.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Name", "ZZZZZZ")


        'Customer
        Select Case FindFieldInf("Customer")
            Case 0  '顯示
                DCustomer.BackColor = Color.LightGray
                DCustomer.ReadOnly = True
                DCustomer.Visible = True
                BCustomer.Visible = False
                BCopySheet.Visible = False
            Case 1  '修改+檢查
                DCustomer.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerRqd", "DCustomer", "異常：需輸入客戶名稱")
                DCustomer.Visible = True
                DCustomer.ReadOnly = True
                BCustomer.Visible = True
                BCopySheet.Visible = True
            Case 2  '修改
                DCustomer.BackColor = Color.Yellow
                DCustomer.Visible = True
                BCustomer.Visible = True
                BCopySheet.Visible = True
            Case Else   '隱藏
                DCustomer.Visible = False
                BCustomer.Visible = False
                BCopySheet.Visible = False
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




        'CustomerColor
        Select Case FindFieldInf("CustomerColor")
            Case 0  '顯示
                DCustomerColor.BackColor = Color.LightGray
                DCustomerColor.ReadOnly = True
                DCustomerColor.Visible = True
            Case 1  '修改+檢查
                DCustomerColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerColorRqd", "DCustomerColor", "異常：需輸入客戶色名")
                DCustomerColor.Visible = True
                DCustomerColor.ReadOnly = True
            Case 2  '修改
                DCustomerColor.BackColor = Color.Yellow
                DCustomerColor.Visible = True
            Case Else   '隱藏
                DCustomerColor.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CustomerColor", "ZZZZZZ")

        'CustomerColorCode
        Select Case FindFieldInf("CustomerColorCode")
            Case 0  '顯示
                DCustomerColorCode.BackColor = Color.LightGray
                DCustomerColorCode.ReadOnly = True
                DCustomerColorCode.Visible = True
            Case 1  '修改+檢查
                DCustomerColorCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerColorCodeRqd", "DCustomerColorCode", "異常：需輸入客戶色號")
                DCustomerColorCode.Visible = True
                DCustomerColorCode.ReadOnly = True
            Case 2  '修改
                DCustomerColorCode.BackColor = Color.Yellow
                DCustomerColorCode.Visible = True
            Case Else   '隱藏
                DCustomerColorCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CustomerColorCode", "ZZZZZZ")


        'OverseaYKKCode
        Select Case FindFieldInf("OverseaYKKCode")
            Case 0  '顯示
                DOverseaYKKCode.BackColor = Color.LightGray
                DOverseaYKKCode.ReadOnly = True
                DOverseaYKKCode.Visible = True
            Case 1  '修改+檢查
                DOverseaYKKCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOverseaYKKCodeRqd", "DOverseaYKKCode", "異常：需輸入海外YKK色號")
                DOverseaYKKCode.Visible = True
                DOverseaYKKCode.ReadOnly = True
            Case 2  '修改
                DOverseaYKKCode.BackColor = Color.Yellow
                DOverseaYKKCode.Visible = True
            Case Else   '隱藏
                DOverseaYKKCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("OverseaYKKCode", "ZZZZZZ")


        'PANTONECode
        Select Case FindFieldInf("PANTONECode")
            Case 0  '顯示
                DPANTONECode.BackColor = Color.LightGray
                DPANTONECode.ReadOnly = True
                DPANTONECode.Visible = True
            Case 1  '修改+檢查
                DPANTONECode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPANTONECodeRqd", "DPANTONECode", "異常：需輸入PANTONE色號")
                DPANTONECode.Visible = True
                DPANTONECode.ReadOnly = True
            Case 2  '修改
                DPANTONECode.BackColor = Color.Yellow
                DPANTONECode.Visible = True
            Case Else   '隱藏
                DPANTONECode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("PANTONECode", "ZZZZZZ")




        '年
        Select Case FindFieldInf("DevYear")
            Case 0  '顯示
                DDevYear.BackColor = Color.LightGray
                DDevYear.Visible = True

            Case 1  '修改+檢查
                DDevYear.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDevYearRqd", "DDevYear", "異常：需輸入年")
                DDevYear.Visible = True
            Case 2  '修改
                DDevYear.BackColor = Color.Yellow
                DDevYear.Visible = True
            Case Else   '隱藏
                DDevYear.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DevYear", "ZZZZZZ")


        '年
        Select Case FindFieldInf("DevSeason")
            Case 0  '顯示
                DDevSeason.BackColor = Color.LightGray
                DDevSeason.Visible = True

            Case 1  '修改+檢查
                DDevSeason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDevSeasonRqd", "DevSeason", "異常：需輸入季")
                DDevSeason.Visible = True
            Case 2  '修改
                DDevSeason.BackColor = Color.Yellow
                DDevSeason.Visible = True
            Case Else   '隱藏
                DDevSeason.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("DevSeason", "ZZZZZZ")

        'DuplicateNo
        Select Case FindFieldInf("DuplicateNo")
            Case 0  '顯示
                DDuplicateNo.BackColor = Color.LightGray
                DDuplicateNo.ReadOnly = True
                DDuplicateNo.Visible = True
            Case 1  '修改+檢查
                DDuplicateNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDuplicateNoRqd", "DDuplicateNo", "異常：需輸入重覆依賴的編號")
                DDuplicateNo.Visible = True
                DDuplicateNo.ReadOnly = True
            Case 2  '修改
                DDuplicateNo.BackColor = Color.Yellow
                DDuplicateNo.Visible = True
            Case Else   '隱藏
                DDuplicateNo.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DuplicateNo", "ZZZZZZ")




        'NOCCSReason
        Select Case FindFieldInf("NOCCSReason")
            Case 0  '顯示
                DNOCCSReason.BackColor = Color.LightGray
                DNOCCSReason.ReadOnly = True
                DNOCCSReason.Visible = True
            Case 1  '修改+檢查
                DNOCCSReason.BackColor = Color.GreenYellow
                '  ShowRequiredFieldValidator("DNOCCSReasonRqd", "DNOCCSReason", "異常：需輸入無法CCS原因")
                DNOCCSReason.Visible = True
            Case 2  '修改
                DNOCCSReason.BackColor = Color.Yellow
                DNOCCSReason.Visible = True
            Case Else   '隱藏
                DNOCCSReason.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("NOCCSReason", "ZZZZZZ")

        'CustomerNGColor
        Select Case FindFieldInf("CustomerNGColor")
            Case 0  '顯示
                DCustomerNGColor.BackColor = Color.LightGray
                DCustomerNGColor.ReadOnly = True
                DCustomerNGColor.Visible = True
            Case 1  '修改+檢查
                DCustomerNGColor.BackColor = Color.GreenYellow
                '  ShowRequiredFieldValidator("DCustomerNGColorRqd", "DCustomerNGColor", "異常：需輸入客戶否決之YKK色號")
                DCustomerNGColor.Visible = True

            Case 2  '修改       

                DCustomerNGColor.BackColor = Color.Yellow
                DCustomerNGColor.Visible = True

            Case Else   '隱藏
                DCustomerNGColor.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CustomerNGColor", "ZZZZZZ")



        'Remark
        Select Case FindFieldInf("Remark")
            Case 0  '顯示
                DRemark.BackColor = Color.LightGray
                DRemark.ReadOnly = True
                DRemark.Visible = True
            Case 1  '修改+檢查
                DRemark.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DRemarkRqd", "DRemark", "異常：需輸入備註")
                DRemark.Visible = True
            Case 2  '修改
                DRemark.BackColor = Color.Yellow
                DRemark.Visible = True
            Case Else   '隱藏
                DRemark.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Remark", "ZZZZZZ")


        '客戶色樣
        Select Case FindFieldInf("CustomerSample")
            Case 0  '顯示
                DCustomerSample.BackColor = Color.LightGray
                DCustomerSample.Visible = True

            Case 1  '修改+檢查
                DCustomerSample.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerSampleRqd", "DCustomerSample", "異常：需輸客戶色樣")
                DCustomerSample.Visible = True
            Case 2  '修改
                DCustomerSample.BackColor = Color.Yellow
                DCustomerSample.Visible = True
            Case Else   '隱藏
                DCustomerSample.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CustomerSample", "ZZZZZZ")





        'Version
        Select Case FindFieldInf("Version")
            Case 0  '顯示
                DVersion.BackColor = Color.LightGray
                DVersion.ReadOnly = True
                DVersion.Visible = True
            Case 1  '修改+檢查
                DVersion.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVersionRqd", "DVersion", "異常：需輸入版本")
                DVersion.Visible = True
                DVersion.ReadOnly = True
            Case 2  '修改
                DVersion.BackColor = Color.Yellow
                DVersion.Visible = True
            Case Else   '隱藏
                DVersion.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Version", "ZZZZZZ")


        'CheckNo
        Select Case FindFieldInf("CheckNo")
            Case 0  '顯示
                DCheckNo.BackColor = Color.LightGray
                DCheckNo.ReadOnly = True
                DCheckNo.Visible = True
            Case 1  '修改+檢查
                DCheckNo.BackColor = Color.GreenYellow
                '    ShowRequiredFieldValidator("DCheckNoRqd", "DCheckNo", "異常：需輸入確認核可代號")
                DCheckNo.Visible = True

            Case 2  '修改
                DCheckNo.BackColor = Color.Yellow
                DCheckNo.Visible = True
            Case Else   '隱藏
                DCheckNo.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CheckNo", "ZZZZZZ")



        'ReceiveDate
        Select Case FindFieldInf("ReceiveDate")
            Case 0  '顯示
                DReceiveDate.BackColor = Color.LightGray
                DReceiveDate.ReadOnly = True
                DReceiveDate.Visible = True
            Case 1  '修改+檢查
                DReceiveDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DReceiveDateRqd", "DReceiveDate", "異常：需輸入接受")
                DReceiveDate.Visible = True
                DReceiveDate.ReadOnly = True
            Case 2  '修改
                DReceiveDate.BackColor = Color.Yellow
                DReceiveDate.Visible = True
            Case Else   '隱藏
                DReceiveDate.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ReceiveDate", "ZZZZZZ")



        'SLDColor
        Select Case FindFieldInf("SLDColor")
            Case 0  '顯示
                DSLDColor.BackColor = Color.LightGray
                DSLDColor.ReadOnly = True
                DSLDColor.Visible = True
                BYKKColor.Visible = False
            Case 1  '修改+檢查
                DSLDColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSLDColorRqd", "DSLDColor", "異常：需輸入確認拉頭兼用色")
                DSLDColor.Visible = True

            Case 2  '修改
                DSLDColor.BackColor = Color.Yellow
                DSLDColor.Visible = True
                DSLDColor.ReadOnly = False
            Case Else   '隱藏
                DSLDColor.Visible = False
                BYKKColor.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SLDColor", "ZZZZZZ")

        'VFColor1 
        Select Case FindFieldInf("VFColor")
            Case 0  '顯示
                DVFColor.BackColor = Color.LightGray
                DVFColor.ReadOnly = True
                DVFColor.Visible = True
            Case 1  '修改+檢查
                DVFColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVFColorRqd", "DVFColor", "異常：需輸入確認VF霧齒")
                DVFColor.Visible = True

            Case 2  '修改
                DVFColor.BackColor = Color.Yellow
                DVFColor.Visible = True
            Case Else   '隱藏
                DVFColor.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("VFColor", "ZZZZZZ")

        'VFColor11 
        Select Case FindFieldInf("VFColor1")
            Case 0  '顯示
                DVFColor1.BackColor = Color.LightGray
                DVFColor1.ReadOnly = True
                DVFColor1.Visible = True
            Case 1  '修改+檢查
                DVFColor1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVFColor1Rqd", "DVFColor1", "異常：需輸入確認VF兼用色")
                DVFColor1.Visible = True

            Case 2  '修改
                DVFColor1.BackColor = Color.Yellow
                DVFColor1.Visible = True
            Case Else   '隱藏
                DVFColor1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("VFColor1", "ZZZZZZ")


        '光源色
        Select Case FindFieldInf("ColorLight1")
            Case 0  '顯示
                DColorLight1.BackColor = Color.LightGray
                DColorLight1.Visible = True

            Case 1  '修改+檢查
                DColorLight1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DColorLight1Rqd", "DColorLight1", "異常：需輸第一對色光源")
                DColorLight1.Visible = True
            Case 2  '修改
                DColorLight1.BackColor = Color.Yellow
                DColorLight1.Visible = True
            Case Else   '隱藏
                DColorLight1.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("ColorLight1", "ZZZZZZ")
        '光源色
        Select Case FindFieldInf("ColorLight2")
            Case 0  '顯示
                DColorLight2.BackColor = Color.LightGray
                DColorLight2.Visible = True

            Case 1  '修改+檢查
                DColorLight2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DColorLight2Rqd", "DColorLight2", "異常：需輸第二對色光源")
                DColorLight2.Visible = True
            Case 2  '修改
                DColorLight2.BackColor = Color.Yellow
                DColorLight2.Visible = True
            Case Else   '隱藏
                DColorLight1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ColorLight2", "ZZZZZZ")


        '光源色
        Select Case FindFieldInf("ColorLight3")
            Case 0  '顯示
                DColorLight3.BackColor = Color.LightGray
                DColorLight3.Visible = True

            Case 1  '修改+檢查
                DColorLight3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("Dcolorlight3Rqd", "Dcolorlight3", "異常：需輸第三對色光源")
                DColorLight3.Visible = True
            Case 2  '修改
                DColorLight3.BackColor = Color.Yellow
                DColorLight3.Visible = True
            Case Else   '隱藏
                DColorLight3.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ColorLight3", "ZZZZZZ")



        'DeliveryDate
        Select Case FindFieldInf("DeliveryDate")
            Case 0  '顯示
                DDeliveryDate.Visible = True
                DDeliveryDate.Style.Add("background-color", "lightgrey")
                DDeliveryDate.Attributes.Add("readonly", "true")
                BDate.Visible = False

            Case 1  '修改+檢查
                DDeliveryDate.Visible = True
                DDeliveryDate.Style.Add("background-color", "greenyellow")
                DDeliveryDate.Attributes.Add("readonly", "true")
                ' ShowRequiredFieldValidator("DDeliveryDateRqd", "DDeliveryDate", "異常：需輸入希望交期")
                BDate.Visible = True

            Case 2  '修改
                DDeliveryDate.Visible = True
                DDeliveryDate.Style.Add("background-color", "yellow")
                DDeliveryDate.Attributes.Add("readonly", "true")
                BDate.Visible = True
            Case Else   '隱藏
                DDeliveryDate.Visible = False
                BDate.Visible = False

        End Select
        If pPost = "New" Then DDeliveryDate.Value = ""





        'YKK色別
        Select Case FindFieldInf("YKKColorType")
            Case 0  '顯示
                DYKKColorType.BackColor = Color.LightGray
                DYKKColorType.Visible = True

            Case 1  '修改+檢查
                DYKKColorType.BackColor = Color.GreenYellow
                '      ShowRequiredFieldValidator("DYKKColorTypeRqd", "DYKKColorType", "異常：需輸YKK色別")
                DYKKColorType.Visible = True
            Case 2  '修改
                DYKKColorType.BackColor = Color.Yellow
                DYKKColorType.Visible = True
            Case Else   '隱藏
                DYKKColorType.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("YKKColorType", "ZZZZZZ")



        '濃淡色
        Select Case FindFieldInf("SLDColor")
            Case 0  '顯示
                DAgain.BackColor = Color.LightGray
                DAgain.Visible = True

            Case 1  '修改+檢查
                DAgain.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAgainRqd", "DAgain", "異常：需輸入濃淡色")
                DAgain.Visible = True
            Case 2  '修改
                DAgain.BackColor = Color.Yellow
                DAgain.Visible = True
            Case Else   '隱藏
                DAgain.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("ColorType", "ZZZZZZ")





        'YKKColorCode
        Select Case FindFieldInf("YKKColorCode")
            Case 0  '顯示
                DYKKColorCode.Visible = True
                DYKKColorCode.Style.Add("background-color", "lightgrey")
                DYKKColorCode.Attributes.Add("readonly", "true")
                BYKKColorCode.Visible = False

            Case 1  '修改+檢查
                DYKKColorCode.Visible = True
                DYKKColorCode.Style.Add("background-color", "greenyellow")
                '  DYKKColorCode.Attributes.Add("readonly", "true")
                '  ShowRequiredFieldValidator("DYKKColorCodeRqd", "DYKKColorCode", "異常：需輸入YKK色號")
                BYKKColorCode.Visible = True

            Case 2  '修改
                DYKKColorCode.Visible = True
                DYKKColorCode.Style.Add("background-color", "yellow")
                DYKKColorCode.Attributes.Add("readonly", "true")
                BYKKColorCode.Visible = True
            Case Else   '隱藏
                DYKKColorCode.Visible = False
                BYKKColorCode.Visible = False

        End Select
        If pPost = "New" Then DYKKColorCode.Value = ""


        'YKKColorcodeVF
        Select Case FindFieldInf("YKKColorCodeVF")
            Case 0  '顯示
                DYKKColorCodeVF.Visible = True
                DYKKColorCodeVF.Style.Add("background-color", "lightgrey")
                DYKKColorCodeVF.Attributes.Add("readonly", "true")
                BYKKColorCodeVF.Visible = False

            Case 1  '修改+檢查
                DYKKColorCodeVF.Visible = True
                DYKKColorCodeVF.Style.Add("background-color", "greenyellow")
                '  DYKKColorcodeVF.Attributes.Add("readonly", "true")
                '  ShowRequiredFieldValidator("DYKKColorcodeVFRqd", "DYKKColorcodeVF", "異常：需輸入YKK色號")
                BYKKColorCodeVF.Visible = True

            Case 2  '修改
                DYKKColorCodeVF.Visible = True
                DYKKColorCodeVF.Style.Add("background-color", "yellow")
                'DYKKColorCodeVF.Attributes.Add("readonly", "true")
                BYKKColorCodeVF.Visible = True
            Case Else   '隱藏
                DYKKColorCodeVF.Visible = False
                BYKKColorCodeVF.Visible = False

        End Select
        If pPost = "New" Then DYKKColorCodeVF.Value = ""


        'YKKColorcodeSLD
        Select Case FindFieldInf("YKKColorCodeSLD")
            Case 0  '顯示
                DYKKColorCodeSLD.Visible = True
                DYKKColorCodeSLD.Style.Add("background-color", "lightgrey")
                DYKKColorCodeSLD.Attributes.Add("readonly", "true")
                BYKKColorCodeSLD.Visible = False

            Case 1  '修改+檢查
                DYKKColorCodeSLD.Visible = True
                DYKKColorCodeSLD.Style.Add("background-color", "greenyellow")
                '  DYKKColorcodeSLD.Attributes.Add("readonly", "true")
                '  ShowRequiredFieldValidator("DYKKColorcodeSLDRqd", "DYKKColorcodeSLD", "異常：需輸入YKK色號")
                BYKKColorCodeSLD.Visible = True

            Case 2  '修改
                DYKKColorCodeSLD.Visible = True
                DYKKColorCodeSLD.Style.Add("background-color", "yellow")
                'DYKKColorCodeSLD.Attributes.Add("readonly", "true")
                BYKKColorCodeSLD.Visible = True
            Case Else   '隱藏
                DYKKColorCodeSLD.Visible = False
                BYKKColorCodeSLD.Visible = False

        End Select
        If pPost = "New" Then DYKKColorCodeSLD.Value = ""




        'PFBWire
        Select Case FindFieldInf("PFBWire")
            Case 0  '顯示
                DPFBWire.Visible = True
                DPFBWire.Style.Add("background-color", "lightgrey")
                DPFBWire.Attributes.Add("readonly", "true")
                BYKKColor.Visible = False

            Case 1  '修改+檢查
                DPFBWire.Visible = True
                DPFBWire.Style.Add("background-color", "greenyellow")
                DPFBWire.Attributes.Add("readonly", "true")
                '  ShowRequiredFieldValidator("DPFBWireRqd", "DPFBWire", "異常：需輸入PF噴塗下止兼用色")
                BYKKColor.Visible = True

            Case 2  '修改
                DPFBWire.Visible = True
                DPFBWire.Style.Add("background-color", "yellow")
                DPFBWire.Attributes.Add("readonly", "true")
                BYKKColor.Visible = True
            Case Else   '隱藏
                DPFBWire.Visible = False
                BYKKColor.Visible = False

        End Select
        If pPost = "New" Then DPFBWire.Value = ""


        'DPFOpenParts
        Select Case FindFieldInf("PFOpenParts")
            Case 0  '顯示
                DPFOpenParts.Visible = True
                DPFOpenParts.Style.Add("background-color", "lightgrey")
                DPFOpenParts.Attributes.Add("readonly", "true")


            Case 1  '修改+檢查
                DPFOpenParts.Visible = True
                DPFOpenParts.Style.Add("background-color", "greenyellow")
                DPFOpenParts.Attributes.Add("readonly", "true")
                '  ShowRequiredFieldValidator("DPFOpenPartsRqd", "DPFOpenParts", "異常：需輸入PF開具色")


            Case 2  '修改
                DPFOpenParts.Visible = True
                DPFOpenParts.Style.Add("background-color", "yellow")
                DPFOpenParts.Attributes.Add("readonly", "true")

            Case Else   '隱藏
                DPFOpenParts.Visible = False


        End Select
        If pPost = "New" Then DPFOpenParts.Value = ""




        'YKK色別
        Select Case FindFieldInf("ColorSystem")
            Case 0  '顯示
                DColorSystem.BackColor = Color.LightGray
                DColorSystem.Visible = True

            Case 1  '修改+檢查
                DColorSystem.BackColor = Color.GreenYellow
                '     ShowRequiredFieldValidator("DColorSystemRqd", "DColorSystem", "異常：需輸色系")
                DColorSystem.Visible = True
            Case 2  '修改
                DColorSystem.BackColor = Color.Yellow
                DColorSystem.Visible = True
            Case Else   '隱藏
                DColorSystem.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("ColorSystem", "ZZZZZZ")




        '新舊色
        Select Case FindFieldInf("NewOldColor")
            Case 0  '顯示
                DNewOldColor.BackColor = Color.LightGray
                DNewOldColor.Visible = True

            Case 1  '修改+檢查
                DNewOldColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNewOldColorRqd", "DNewOldColor", "異常：需輸新舊色")
                DNewOldColor.Visible = True
            Case 2  '修改
                DNewOldColor.BackColor = Color.Yellow
                DNewOldColor.Visible = True
            Case Else   '隱藏
                DNewOldColor.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("NewOldColor", "ZZZZZZ")


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
                DReRegister.Checked = False
        End Select


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
        DDepoName.Text = DBUser.Rows(0).Item("Divname")
        DName.Text = DBUser.Rows(0).Item("Username")

        '光源
        If pFieldName = "ColorLight1" Then
            DColorLight1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DColorLight1.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'ColorLight'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DColorLight1.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DColorLight1.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        '光源
        If pFieldName = "ColorLight2" Then
            DColorLight2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DColorLight2.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  data from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'ColorLight2' "
                SQL = SQL & " order by createtime desc "



                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DColorLight2.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DColorLight2.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        '光源
        If pFieldName = "ColorLight3" Then
            DColorLight3.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DColorLight3.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  data from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'colorlight2' "
                SQL = SQL & " order by createtime desc "



                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DColorLight3.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DColorLight3.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'YKK色別
        If pFieldName = "YKKColorType" Then
            DYKKColorType.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DYKKColorType.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'YKKColorType'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DYKKColorType.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DYKKColorType.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'YKK色別
        If pFieldName = "ColorType" Then
            DAgain.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAgain.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'ColorType'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DAgain.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAgain.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'YKK色別
        If pFieldName = "CustomerSample" Then

            DCustomerSample.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCustomerSample.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'CustomerSample'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DCustomerSample.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCustomerSample.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        '開發年
        If pFieldName = "DevYear" Then

            DDevYear.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDevYear.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'DevYear' order by data"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DDevYear.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDevYear.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        '開發季
        If pFieldName = "DevSeason" Then

            DDevSeason.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDevSeason.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'DevSeason'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DDevSeason.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDevSeason.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        '新色舊色
        If pFieldName = "NewOldColor" Then
            DNewOldColor.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DNewOldColor.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'NewOldColor'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DNewOldColor.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DNewOldColor.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
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
        If Request.Cookies("RunBSAVE").Value = True And OK() = True Then

            DisabledButton()   '停止Button運作
            FlowControl("OK", 3, "1")

        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BOK_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
        If Request.Cookies("RunBOK").Value = True And OK() = True Then

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

            If wStep = 10 Then
                If OK() = True Then
                    DisabledButton()   '停止Button運作
                    FlowControl("NG2", 2, "3")

                End If
            Else
                DisabledButton()   '停止Button運作
                FlowControl("NG2", 2, "3")

            End If
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
                ErrCode = oCommon.CommissionNo("005011", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
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
                            'oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDepo,wUserID, wApplyID)
                            '
                            oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, wUserID, wApplyID)
                            '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者
                            'Modify-End
                            '設定委託No
                            DNo.Text = SetNo(NewFormSno)

                        End If
                        pRunNextStep = 1
                    Else
                        If RepeatRun = False Then   '不是通知的重覆執行
                            '更新交易資料
                            'jessica 2021/7/21 step 65 退DYE
                            If pAction = 3 Then
                                ModifyTranData(pFun, "4")
                            Else
                                ModifyTranData(pFun, pSts)
                            End If



                            '流程資料結束
                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDepo,wUserID, pRunNextStep)
                            '
                            RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, wUserID, pRunNextStep)
                            '表單號碼,表單流水號,工程關卡號碼,行事曆,簽核者, 流程結束否(會簽)
                            'Modify-End

                            'If wStep = 541 Or wStep = 546 Then
                            'pRunNextStep = 1
                            'End If

                            '檢查是否都完成才送至60
                            Dim SQL1 As String
                            If wStep = 51 Then
                                SQL1 = " Select step From T_WaitHandle Where "
                                SQL1 = SQL1 + "    FormNo  = '005011'   And FormSno = '" + CStr(wFormSno) + "'   And Step = '52'   and  modifyuser <> 'Operation_02'  "
                                Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL1)
                                If dtFlow.Rows.Count > 0 Then
                                    pRunNextStep = 1
                                Else
                                    pRunNextStep = 0
                                End If
                            End If

                            If wStep = 52 Then
                                SQL1 = " Select step From T_WaitHandle Where  "
                                SQL1 = SQL1 + "    FormNo  = '005011'   And FormSno = '" + CStr(wFormSno) + "'   And Step = '51'  and  modifyuser <> 'Operation_02'  "
                                Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL1)
                                If dtFlow.Rows.Count > 0 Then
                                    pRunNextStep = 1
                                Else
                                    pRunNextStep = 0
                                End If
                            End If

                            If RtnCode <> 0 Then ErrCode = 9120
                        Else
                            pRunNextStep = 1    '是通知的重覆執行
                    End If
                    End If
                End If

                '取得下一關
                If ErrCode = 0 And pRunNextStep = 1 Then

                    Dim wAllocateID As String = ""
                    '只有螢光的才CC給怡蓉，若不是就直接給起單者
                    'If wStep = 40 And pAction = 0 And DYKKColorType.SelectedValue <> "螢光" Then
                    ' pAction = 3
                    'End If
                    ' 20150408 修改40工程如果不是舊色就跳到70工程

                    If (wStep = 30 Or wStep = 530) And pFun = "OK" Then
                        If DBuyerCode.Text = "TW0371" Then 'UA
                            pAction = 0
                        ElseIf DBuyerCode.Text = "TW0011" Then 'KIPLING
                            pAction = 2
                        ElseIf DBuyerCode.Text = "000021" Then 'KIPLING
                            pAction = 0
                        End If
                    End If


                    If wStep = 60 And pFun = "NG1" Then
                        If DBuyerCode.Text = "TW0371" Then 'UA
                            pAction = 1
                        ElseIf DBuyerCode.Text = "TW0011" Then 'KIPLING
                            pAction = 3
                        ElseIf DBuyerCode.Text = "000021" Then 'KIPLING
                            pAction = 1
                        End If
                    End If


                    If (wStep = 48 Or wStep = 548) And pFun = "OK" Then
                        If DBuyerCode.Text = "TW0371" Then 'UA
                            pAction = 0
                        ElseIf DBuyerCode.Text = "TW0011" Then 'KIPLING
                            pAction = 2
                        ElseIf DBuyerCode.Text = "000021" Then 'KIPLING
                            pAction = 0
                        End If
                    End If


                    If wStep = 70 Then
                        Dim SQL1 As String
                        Dim DTNO As Integer
                        SQL1 = " select * from F_NewcolorUAKIPLINGDT"
                        SQL1 = SQL1 + " where dtsheet <> '拉頭' and OFORMSNO = '" + CStr(wFormSno) + "'  AND STS <>'2' "
                        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL1)
                        If dtFlow.Rows.Count = 0 Then
                            DTNO = 0  '判斷是否有核可單
                        Else
                            DTNO = 1
                        End If

                        If DNewOldColor.SelectedValue = "新色" And DTNO = 0 Then
                            pAction = 0
                        Else
                            If DYKKColorType.SelectedValue = "螢光" Then
                                pAction = 1  '151
                            Else
                                pAction = 2
                            End If
                        End If

                    End If

                    If wStep = 75 And pFun = "OK" Then
                        If DYKKColorType.SelectedValue = "螢光" Then
                            pAction = 0
                        Else
                            pAction = 1
                        End If
                    End If


                    If wStep = 155 Then
                        If DBuyerCode.Text = "TW0371" Then 'UA
                            pAction = 0
                        ElseIf DBuyerCode.Text = "TW0011" Then 'KIPLING
                            pAction = 1
                        ElseIf DBuyerCode.Text = "000021" Then 'TNF
                            pAction = 1
                        End If
                    End If

                    RtnCode = oCommon.GetNextGate(wFormNo, wStep, wUserID, wApplyID, wAgentID, wAllocateID, MultiJob, _
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
                            'RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wDepo,wUserID, pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
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
                            RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, wUserID, pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '表單號碼,表單流水號,工程關卡號碼,序號,申請者ID,行事曆,建置者, 簽核者, 被代理者, 申請者, 預定開始日, 預定完成日, 重要性
                            'Modify-End

                            If RtnCode <> 0 Then
                                ErrCode = 9150
                                Exit For
                            End If
                        Next i
                    Else
                        'Modify-Start by 2009/11/20(2010行事曆對應)
                        'RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDepo,wUserID, wApplyID)
                        '
                        RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDecideCalendar, wUserID, wApplyID)
                        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者
                        'Modify-End

                        If RtnCode <> 0 Then ErrCode = 9160
                    End If
                End If
                '當工程日程調整
                If ErrCode = 0 Then
                    If RepeatRun = False Then
                        'Modify-Start by 2009/11/20(2010行事曆對應)
                        'RtnCode = oSchedule.AdjustSchedule(wUserID, wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDepo)
                        '
                        RtnCode = oSchedule.AdjustSchedule(wUserID, wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDecideCalendar)
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
                            oCommon.Send(wUserID, pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "FLOW")
                            '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                        Next i
                    Else
                        oCommon.Send(wUserID, wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "END")
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


                    LastStep = wStep  '記錄上一個工程
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

                Dim URL As String

                ' MODIFY-START BY JOY 2015/4/7 (徐慧芬)
                ' If (LastStep = 20 Or LastStep = 520 Or LastStep = 25) And pFun = "OK" Then '新色確認完成表按OK才進來
                '增加連續起單 Modify 20150419
                If wbFormSno = 0 And wbStep < 3 And DReRegister.Checked = True Then    '判斷是否起單

                    uJavaScript.PopMsg(Me, "已成功送出您所填寫之申請單，請繼續申請！")
                    EnabledButton()                             '起動Button運作


                    DNo.Text = ""
                    wFormSno = wbFormSno
                    wStep = wbStep


                Else


                    If ((wbStep = 20 Or wbStep = 520) And pFun = "OK") Or ((wbStep = 30 Or wbStep = 530) And pFun = "OK") Then '新色確認完成表按OK才進來
                        Dim FormName As String
                        If LastStep = 20 Or LastStep = 520 Then
                            FormName = "3CF DTM"
                        Else
                            FormName = "5VS  DTM"
                        End If

                        'MODIFY-END
                        Dim SQL As String  '檢查是否有相同版本的新色依賴完成書
                        SQL = " select  *  from F_NewColorUAKIPLING a, F_NewColorComplete b"
                        SQL = SQL & " where a.formno = b.formno And a.formsno = b.formsno"
                        SQL = SQL & " and a.version = b.version "
                        SQL = SQL & " and formname ='" & FormName & "'"
                        SQL = SQL & " and a.No = '" & DNo.Text + "'"
                        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)
                        If dtFlow.Rows.Count = 0 Then '如果沒有才可以新增
                            URL = uCommon.GetAppSetting("NewColorCompleteUrlUA") & "?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno & "&pStep=" & LastStep & "&pNextGate=" & wNextGate & _
                                                                                                       "&pUserID=" & wUserID & "&pApplyID=" & wApplyID
                        Else
                            URL = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                                                  "&pUserID=" & wUserID & "&pApplyID=" & wApplyID
                        End If


                    Else

                        URL = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                                                   "&pUserID=" & wUserID & "&pApplyID=" & wApplyID
                    End If
                    Response.Redirect(URL)
                End If '

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

        Dim SQl As String
        SQl = "Insert into F_NewColorUAKIPLING "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno,"
        SQl = SQl + "No, DepoName,Date,Name,CustomerCheck,"  '1~5                
        SQl = SQl + "FactoryCheck,VCACheck,Customer,CustomerCode,Buyer,"   '6~10              
        SQl = SQl + "BuyerCode,CustomerColor,CustomerColorCode,OverseaYKKCode,PANTONECode,"  '11~15
        SQl = SQl + "DevYear,DevSeason,ReceiveDate,ColorLight1,ColorLight2,ColorLight3,NewOldColor," ' 16~20
        SQl = SQl + "DeliveryDate,DuplicateNo,NOCCS,NOCCSReason,CustomerNGColor," '21~25
        SQl = SQl + "Remark,CustomerSample,ColorSystem,PFBWire,PFOpenParts," '26~30
        SQl = SQl + "YKKColorType,YKKColorCode,YKKColorCodeVF,YKKColorCodeSLD,Version,CheckNo,SLDColor,Again," '31~35
        SQl = SQl + "VFCOLOR,VFCOLOR1," '36
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime) "
        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        SQl = SQl + " '" + NowDateTime + "', "        '結案日
        SQl = SQl + " '005011', "                     '表單代號
        SQl = SQl + " '" + CStr(NewFormSno) + "', "   '表單流水號
        SQl = SQl + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO  1
        SQl = SQl + " N'" + DDepoName.Text + "', "                '部門2
        SQl = SQl + " N'" + DDate.Text + "', "                '日期3
        SQl = SQl + " N'" + DName.Text + "', "                '姓名4
        If DCustomerCheck.Checked = True Then
            SQl = SQl + " '1', "  '5
        Else
            SQl = SQl + " '0', "  '5:
        End If

        If DFactoryCheck.Checked = True Then
            SQl = SQl + " '1', "  '6
        Else
            SQl = SQl + " '0', "  '6
        End If


        If DVCACheck.Checked = True Then
            SQl = SQl + " '1', "  '7
        Else
            SQl = SQl + " '0', "  '7
        End If


        SQl = SQl + " N'" + DCustomer.Text.ToUpper + "', "                '客戶8
        SQl = SQl + " N'" + DCustomerCode.Text.ToUpper + "', "                '客戶9
        SQl = SQl + " N'" + DBuyer.Text.ToUpper + "', "                'BUYER10
        SQl = SQl + " N'" + DBuyerCode.Text.ToUpper + "', "                'BUYER11
        SQl = SQl + " N'" + YKK.ReplaceString(DCustomerColor.Text.ToUpper) + "', "                ' 客戶色名12
        SQl = SQl + " N'" + YKK.ReplaceString(DCustomerColorCode.Text.ToUpper) + "', "                ' 客戶色號13
        SQl = SQl + " N'" + YKK.ReplaceString(DOverseaYKKCode.Text.ToUpper) + "', "                ' 海外YKK色號14
        SQl = SQl + " N'" + YKK.ReplaceString(DPANTONECode.Text.ToUpper) + "', "                ' PANTONECode色號15
        SQl = SQl + " N'" + DDevYear.SelectedValue + "', "                ' 開發年16
        SQl = SQl + " N'" + DDevSeason.SelectedValue + "', "                ' 開發季17
        SQl = SQl + " N'" + DReceiveDate.Text + "', "                ' 收受日期18
        SQl = SQl + " N'" + DColorLight1.SelectedValue.ToUpper + "', "                ' 第一對色光源19
        SQl = SQl + " N'" + DColorLight2.SelectedValue.ToUpper + "', "                ' 第二對色光源20
        SQl = SQl + " N'" + DColorLight3.SelectedValue.ToUpper + "', "                ' 第二對色光源20
        SQl = SQl + " N'" + DNewOldColor.SelectedValue.ToUpper + "', "                ' 第二對色光源20
        SQl = SQl + " N'" + DDeliveryDate.Value + "', "                ' 重覆依賴的編號22
        SQl = SQl + " N'" + DDuplicateNo.Text.ToUpper + "', "                ' 收受日期21
        If DNOCCS.Checked = True Then
            SQl = SQl + " '1', " '23
        Else
            SQl = SQl + " '0', " '23
        End If
        SQl = SQl + " N'" + YKK.ReplaceString(DNOCCSReason.Text) + "', "                ' 無法CCS原因 24
        SQl = SQl + " N'" + YKK.ReplaceString(DCustomerNGColor.Text.ToUpper) + "', "                ' 客戶否決之色號25
        SQl = SQl + " N'" + YKK.ReplaceString(DRemark.Text) + "', "                ' 備註26
        SQl = SQl + " N'" + DCustomerSample.SelectedValue.ToUpper + "', "                ' 客戶樣品27
        SQl = SQl + " N'" + DColorSystem.Text.ToUpper + "', "                ' 色系29
        SQl = SQl + " N'" + DPFBWire.Value.ToUpper + "', "                ' BWire29
        SQl = SQl + " N'" + DPFOpenParts.Value.ToUpper + "', "                ' 開具色30
        SQl = SQl + " N'" + DYKKColorType.SelectedValue.ToUpper + "', "                ' YKK色別31
        SQl = SQl + " N'" + DYKKColorCode.Value.ToUpper + "', "                ' YKK色號32
        SQl = SQl + " N'" + DYKKColorCodeVF.Value.ToUpper + "', "                ' YKK色號32
        SQl = SQl + " N'" + DYKKColorCodeSLD.Value.ToUpper + "', "                ' YKK色號32

        If DVersion.Text = "" Then
            SQl = SQl + " 0, "                ' 版本33
        Else
            SQl = SQl + " N'" + DVersion.Text + "', "                ' 版本33
        End If

        SQl = SQl + " N'" + DCheckNo.Text.ToUpper + "', "                ' 確認核可代號34
        SQl = SQl + " N'" + DSLDColor.Text.ToUpper + "', "                ' 拉頭部色號35

        If DAgain.SelectedValue = "淡色" Then
            SQl = SQl + "1, "                ' 淡色
        ElseIf DAgain.SelectedValue = "濃色" Then
            SQl = SQl + "2, "                ' 濃色
        Else
            SQl = SQl + "0, "
        End If

        SQl = SQl + " N'" + DVFColor.Text.ToUpper + "', "                ' VF色號36
        SQl = SQl + " N'" + DVFColor1.Text.ToUpper + "', "                ' VF色號36
        SQl = SQl + " '" + wUserID + "', "     '作成者
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
        Dim SQL As String = ""

        SQL = " Update F_NewColorUAKIPLING"
        SQL = SQL + " Set "
        If pFun <> "SAVE" Then
            SQL = SQL + " Sts = '" & pSts & "',"
            SQL = SQL + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQL = SQL + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"
        SQL = SQL + " Date = N'" & DDate.Text & "',"
        SQL = SQL + " DepoName = N'" & DDepoName.Text & "',"
        SQL = SQL + " Name = N'" & DName.Text & "',"
        If DCustomerCheck.Checked = True Then
            SQL = SQL + " CustomerCheck =1,"
        Else
            SQL = SQL + " CustomerCheck =0,"
        End If
        If DFactoryCheck.Checked = True Then
            SQL = SQL + " FactoryCheck =1,"
        Else
            SQL = SQL + " FactoryCheck =0,"
        End If
        If DVCACheck.Checked = True Then
            SQL = SQL + " VCACheck =1,"
        Else
            SQL = SQL + " VCACheck =0,"
        End If
        SQL = SQL + " Customer = N'" & DCustomer.Text.ToUpper & "',"
        SQL = SQL + " CustomerCode = N'" & DCustomerCode.Text.ToUpper & "',"
        SQL = SQL + " Buyer = N'" & DBuyer.Text & "',"
        SQL = SQL + " BuyerCode = N'" & DBuyerCode.Text & "',"
        SQL = SQL + " CustomerColor = N'" & YKK.ReplaceString(DCustomerColor.Text.ToUpper) & "',"
        SQL = SQL + " CustomerColorCode = N'" & YKK.ReplaceString(DCustomerColorCode.Text.ToUpper) & "',"
        SQL = SQL + " OverseaYKKCode = N'" & YKK.ReplaceString(DOverseaYKKCode.Text.ToUpper) & "',"
        SQL = SQL + " PANTONECode = N'" & YKK.ReplaceString(DPANTONECode.Text.ToUpper) & "',"
        SQL = SQL + " DevYear = N'" & DDevYear.SelectedValue & "',"
        SQL = SQL + " DevSeason = N'" & DDevSeason.SelectedValue & "',"
        If wStep = 10 Then
            SQL = SQL + " ReceiveDate = N'" & CDate(NowDateTime).ToString("yyyy/MM/dd") & "',"
        Else
            SQL = SQL + " ReceiveDate = N'" & DReceiveDate.Text & "',"
        End If

        SQL = SQL + " ColorLight1 = N'" & DColorLight1.SelectedValue.ToUpper & "',"
        SQL = SQL + " ColorLight2 = N'" & DColorLight2.SelectedValue.ToUpper & "',"
        SQL = SQL + " ColorLight3 = N'" & DColorLight3.SelectedValue.ToUpper & "',"
        SQL = SQL + " NewOldColor = N'" & DNewOldColor.SelectedValue.ToUpper & "',"
        SQL = SQL + " DeliveryDate = N'" & DDeliveryDate.Value & "',"
        SQL = SQL + " DuplicateNo = N'" & DDuplicateNo.Text.ToUpper & "',"

        If DNOCCS.Checked = True Then
            SQL = SQL + " NOCCS =1,"
        Else
            SQL = SQL + " NOCCS =0,"
        End If


        SQL = SQL + " NOCCSReason = N'" & YKK.ReplaceString(DNOCCSReason.Text) & "',"
        SQL = SQL + " CustomerNGColor = N'" & YKK.ReplaceString(DCustomerNGColor.Text.ToUpper) & "',"
        SQL = SQL + " Remark = N'" & YKK.ReplaceString(DRemark.Text) & "',"
        SQL = SQL + " CustomerSample = N'" & DCustomerSample.SelectedValue.ToUpper & "',"
        SQL = SQL + " ColorSystem = N'" & DColorSystem.Text.ToUpper & "',"
        SQL = SQL + " PFBWire = N'" & DPFBWire.Value.ToUpper & "',"
        SQL = SQL + " PFOpenParts = N'" & DPFOpenParts.Value.ToUpper & "',"
        SQL = SQL + " YKKColorType = N'" & DYKKColorType.SelectedValue.ToUpper & "',"
        SQL = SQL + " YKKColorCode = N'" & DYKKColorCode.Value.ToUpper & "',"
        SQL = SQL + " YKKColorCodeVF = N'" & DYKKColorCodeVF.Value.ToUpper & "',"
        SQL = SQL + " YKKColorCodeSLD = N'" & DYKKColorCodeSLD.Value.ToUpper & "',"

        ' If wStep = 20 Then
        ' SQL = SQL + " Version =1,"
        ' Else
        SQL = SQL + " Version ='" & DVersion.Text.ToUpper & "',"
        ' End If
        SQL = SQL + " CheckNo = N'" & DCheckNo.Text.ToUpper & "',"

        If DSLDColor.Text.ToUpper <> "" Then
            SQL = SQL + " SLDColor = N'" & DSLDColor.Text.ToUpper & "',"
        End If


        If DVFColor.Text.ToUpper <> "" Then
            SQL = SQL + " VFColor = N'" & DVFColor.Text.ToUpper & "',"
        End If


        If DAgain.SelectedValue <> "" Then
            If DAgain.SelectedValue = "淡色" Then
                SQL = SQL + " again =1,"
            ElseIf DAgain.SelectedValue = "濃色" Then
                SQL = SQL + " again =2,"
            End If

        End If


      

        If DVFColor1.Text.ToUpper <> "" Then
            SQL = SQL + " VFColor1 = N'" & DVFColor1.Text.ToUpper & "',"
        End If


        SQL = SQL + " ModifyUser = '" & wUserID & "',"
        SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
        SQL = SQL + " Where Formsno ='" + Str(wFormSno) + "'"
        uDataBase.ExecuteNonQuery(SQL)
        Dim Code As String = ""

        If wStep = 90 And DNewOldColor.SelectedValue = "新色" Then 'Master 建檔後轉資料到WINS COLOR STRUCTURE

            If DSLDColor.Text <> "" And DAgain.SelectedValue <> "" Then
                ' INSERT 濃淡色
                Dim YB1CP2 As String = ""
                If DAgain.SelectedValue = "淡色" Then
                    YB1CP2 = "0"
                ElseIf DAgain.SelectedValue = "濃色" Then
                    YB1CP2 = "1"
                End If

                '先確定colorcode 是否存在, 不存在才INSERT 
                Dim SLDColor As String

                If Len(DSLDColor.Text) = 3 Then  '三碼加2格空白
                    SLDColor = "  " + DSLDColor.Text
                Else
                    SLDColor = DSLDColor.Text
                End If

                SQL = "Select * From R_NCTP200WK  "
                SQL = SQL & " Where CLRCP2 = '" + SLDColor + "'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                If DBAdapter1.Rows.Count = 0 Then
                    SQL = " INSERT INTO R_NCTP200WK (NO,STS,CLRCP2,YB1CP2,UIDCP2,PRGCP2,DEVCP2,RADUP2,RADTP2,RUPUP2,RUPTP2,RCMCP2)"
                    SQL = SQL + "VALUES('" + DNo.Text + "','0','" + SLDColor + "','" + YB1CP2 + "','RPA','UPLOAD','WINGS',convert(char(10),getdate(),112) ,"
                    SQL = SQL + " substring(convert(char(10),getdate(),108),1,2)+substring(convert(char(10),getdate(),108),4,2)+substring(convert(char(10),getdate(),108),7,2) ,"
                    SQL = SQL + " '0','0','000034')"
                    uDataBase.ExecuteNonQuery(SQL)

                    Code = oWaves.NewColorType2Wings(DNo.Text)
                    If Code <> "0" Then
                        uJavaScript.PopMsg(Me, "匯入WINS異常，請洽電腦室")

                    End If
                End If


            End If



            SQL = "insert into R_NCFA300WK (no,sts,PDPCW3,CTBNW3,PACCW3,ST1CW3,ST6CW3,SCSVW3,CCLCW3,UIDCW3,PRGCW3,DEVCW3,RADUW3,RADTW3,RUPUW3,RUPTW3,RCMCW3, PRBFW3, CHKFW3,WKDTW3)"
            SQL = SQL + " select no,STS,PDPCW3,CTBNW3,PACCW3,ST1CW3,ST6CW3,SCSVW3,"
            SQL = SQL + "  CASE WHEN CCLCW3 ='V7' THEN '  V7+' ELSE RIGHT(REPLICATE(' ', 5) + CAST(CCLCW3 as NVARCHAR), 5) END AS CCLCW3,UIDCW3,PRGCW3,DEVCW3,RADUW3,RADTW3,RUPUW3,RUPTW3,RCMCW3, PRBFW3, CHKFW3,WKDTW3 from ("
            SQL = SQL + " select NO,'0' AS STS,'01' as PDPCW3 ,CTBNW3,ykkcolorcode as PACCW3,ST1CW3,ST6CW3,SCSVW3,"
            SQL = SQL + " Case when CCLCW3 ='YKKColorCode' then YKKColorCode "
            SQL = SQL + " when CCLCW3 ='VFColor' then VFColor"
            SQL = SQL + " when CCLCW3 ='PFBWire' then PFBWire"
            SQL = SQL + " when CCLCW3 ='PFOpenParts' then PFOpenParts"
            SQL = SQL + " when CCLCW3 ='SLDColor' then SLDColor "
            SQL = SQL + " when CCLCW3 ='YKKColorCodeSLD' then YKKColorCodeSLD"
            SQL = SQL + " when CCLCW3 ='YKKColorCodeVF' then YKKColorCodeVF"
            SQL = SQL + " when CCLCW3 ='TSCC' then CCLCW3"
            SQL = SQL + " when CCLCW3 ='C5+' then CCLCW3"
            SQL = SQL + " end as CCLCW3 ,'' as UIDCW3,'' as PRGCW3,'' as DEVCW3,convert(char(10),getdate(),112)as  RADUW3,"
            SQL = SQL + " substring(convert(char(10),getdate(),108),1,2)+substring(convert(char(10),getdate(),108),4,2)+substring(convert(char(10),getdate(),108),7,2) as RADTW3,"
            SQL = SQL + " '' as RUPUW3,'' as RUPTW3,'' as RCMCW3,'' as PRBFW3,'' as CHKFW3,'' as WKDTW3"
            SQL = SQL + " from v_newcolor a,R_NewColorST_01 b"
            SQL = SQL + " where sts in(1,0) and no = '" + DNo.Text + "' AND TYPE ='N2'"
            SQL = SQL + " )a"
            uDataBase.ExecuteNonQuery(SQL)


            Code = oWaves.NewColor2Wings(DNo.Text)

            If Code <> "0" Then
                uJavaScript.PopMsg(Me, "匯入WINS異常，請洽電腦室")
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
            'jessica 2021/7/21 step 65 退DYE
            If pSts = "4" Then SQl = SQl + " StsDes = '" & BSAVE.Value & "',"
            SQl = SQl + " AEndTime = '" & NowDateTime & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
            If DDelay.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & wUserID & "',"
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
            SQl = SQl + " ModifyUser = '" & wUserID & "',"
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
                SQl = SQl + " '" + wUserID + "', "
                SQl = SQl + " '" + NowDateTime + "' "
                SQl = SQl + " ) "
                uDataBase.ExecuteNonQuery(SQl)
            End If
        Else
            If DNo.Text <> "" Then
                If DNo.Text <> DBAdapter1.Rows(0).Item("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNo.Text & "',"
                    SQl = SQl + " CreateUser = '" & wUserID & "',"
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



        If DDescSheet.Visible Then                                      ' 說明
            DDescSheet.Style("top") = Top - 130 & "px"
            DDecideDesc.Style("top") = Top - 130 + 6 & "px"
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


        Top = Top + 50
        If GridView2.Rows.Count > 0 Then                                ' 核定履歷
            DHistoryLabel.Style("top") = Top & "px"
            Top += 20
            GridView2.Style("top") = Top & "px"
        End If
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
        Dim SQL As String = ""

        'Set當日日期
        Str2 = CStr(DateTime.Now.Month)  '月
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        ' Str = Str2
        ' Str2 = CStr(DateTime.Now.Day)    '日
        'For i = Str2.Length To 1
        'Str2 = "0" + Str2
        'Next
        'Str = Str + Str2
        Str = "J" + CStr(DateTime.Now.Year) + Str2
        '年
        'Set單號
        '當月份筆數有幾筆  20150414 Modify by Jessica
        SQL = " select   isnull(max(convert(int,substring(no,8,4))),0) cun from  F_NewColorUAKIPLING"
        SQL = SQL + " where  left(convert(char(10),date ,111),7) = left(convert(char(10), getdate(),111),7)"
        Dim dt1 As DataTable = uDataBase.GetDataTable(SQL)
        If dt1.Rows.Count > 0 Then
            Str1 = CStr(CInt(dt1.Rows(0).Item("cun")) + 1)
        End If

        ' Str1 = CStr(Seq)

        For i = Str1.Length To 4 - 1
            Str1 = "0" + Str1
        Next

        SetNo = Str + Str1
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     輸入檢查
    '**
    '*****************************************************************
    Function OK() As Boolean
        Top = 1200
        SetControlPosition()
        '  ShowFormData()
        Dim SQL As String
        Dim DOUBLENo As Integer = 0

        Dim Errcode As Integer = 0


        If wStep = 1 Then
            If DCustomerCheck.Checked = False And DFactoryCheck.Checked = False And DVCACheck.Checked = False Then
                isOK = False
                Message = "異常：請勾選客戶自核或工廠自核或VCA!"

            End If

            If DVCACheck.Checked = False Then
                If DCustomerNGColor.Text = "" Then
                    DCustomerNGColor.BackColor = Color.GreenYellow
                    DCustomerNGColor.Visible = True
                    isOK = False
                    Message = "異常：需輸入客戶否決之YKK色號!"

                End If
            End If
        End If



        If wStep = 20 Or wStep = 30 Or wStep = 530 Or wStep = 520 Then
            If DColorSystem.Text = "" Then
                isOK = False
                Message = "異常：請輸入色系!"
            End If

            If DPFBWire.Value = "" Then
                isOK = False
                Message = Message + "\n" + "異常：請輸入PF下止噴塗兼用色!"
            End If
            If DPFOpenParts.Value = "" Then
                isOK = False
                Message = Message + "\n" + "異常：請輸入PF開具色!"
            End If

        End If

        '  If wStep = 60 Then
        ' If DCheckNo.Text = "" Then
        'isOK = False
        'Message = "異常：請輸入確認核可代號!"
        'End If

        'End If

        If wStep = 500 Or wStep = 1 Then  '(NG再送出希望交期需重新選擇)
            If DDeliveryDate.Value = "" Then
                isOK = False
                Message = "異常：請輸入希望交期!"
            End If
        End If

        If DBuyerCode.Text <> "TW0371" And DBuyerCode.Text <> "TW0011" And DBuyerCode.Text <> "000021" Then
            isOK = False
            Message = "異常：BUYER 請輸入UA 或 KIPPLING!"
        End If



        If wStep = 1 Then

            '檢查是否有重覆依賴的編號
            ' Dim SQL As String
            SQL = "select no,customer,customercode,buyer,buyercode,customercolor,customercolorcode,overseaykkcode,pantonecode devyear,devseason "
            SQL = SQL + " from  v_newcolor"
            SQL = SQL + " where sts <>2 and customercode='" + DCustomerCode.Text + "'"
            SQL = SQL + " and  buyercode='" + DBuyerCode.Text + "'"
            SQL = SQL + " and customercolor='" + YKK.ReplaceString(DCustomerColor.Text) + "'"
            SQL = SQL + " and  customercolorcode='" + YKK.ReplaceString(DCustomerColorCode.Text) + "'"
            SQL = SQL + " and overseaykkcode='" + YKK.ReplaceString(DOverseaYKKCode.Text) + "'"
            SQL = SQL + " and pantonecode='" + YKK.ReplaceString(DPANTONECode.Text) + "'"
            SQL = SQL + " and  devyear='" + DDevYear.SelectedValue + "'"
            SQL = SQL + " and devseason='" + DDevSeason.SelectedValue + "'"
            SQL = SQL + " and formno = '" + wFormNo + "'"


            Dim NoStr As String = ""
            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
            If DBAdapter1.Rows.Count > 0 Then
                For Each dtr As Data.DataRow In DBAdapter1.Rows
                    If NoStr = "" Then
                        NoStr = dtr("No")
                    Else
                        NoStr = NoStr + "," + dtr("No")
                    End If

                Next
                If NoStr = DDuplicateNo.Text Then  '如果重覆的編號一樣就不用再執行一次
                    DOUBLENo = DOUBLENo + 1
                End If
            Else
                DDuplicateNo.Text = "" '如果沒有重覆就清空

            End If


            '20180612 重覆取消
            DDuplicateNo.Text = NoStr
            If DDuplicateNo.Text <> "" Then

                Message = Message + "\n" + "有重覆依賴的編號請確認後再繼續執行?!"
                isOK = False
            End If

            'If NoStr <> "" And DOUBLENo = 0 Then

            '    DDuplicateNo.Text = NoStr

            '    Message = Message + "\n" + "有重覆依賴的編號請確認後再繼續執行?!"

            '    isOK = False


            'End If

        End If

        '檢查是否有重覆依PANTONE色號 20150521
        ' Dim SQL As String
        If wStep = 1 Or wStep = 500 Then
            If YKK.ReplaceString(DPANTONECode.Text) <> "" Then

                SQL = "select no "
                SQL = SQL + " from V_newColor"
                SQL = SQL + " where sts =0 and  pantonecode='" + YKK.ReplaceString(DPANTONECode.Text) + "'"
                If wStep = 500 Then
                    SQL = SQL + " and no <> '" + DNo.Text + "'"
                End If





                Dim NoStr As String = ""
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                If DBAdapter1.Rows.Count > 0 Then
                    For Each dtr As Data.DataRow In DBAdapter1.Rows
                        If NoStr = "" Then
                            NoStr = dtr("No")
                        Else
                            NoStr = NoStr + "," + dtr("No")
                        End If


                    Next

                Else
                    DDuplicateNo.Text = "" '如果沒有重覆就清空

                End If

                If NoStr <> "" Then

                    DDuplicateNo.Text = NoStr
                    Message = Message + "\n" + "有重覆依賴的PANTONE色號，不可執行?!"
                    isOK = False


                End If
            End If
        End If





        'Dim Q As Integer = 0
        ''jessica 20150326
        'If wStep = 70 Then
        '    '檢查是否有Q 
        '    Q = InStr(1, DYKKColorCode.Value, "Q", 1)
        '    If DYKKColorType.SelectedValue = "螢光" Then


        '        SQL = " select * from   m_referp"
        '        SQL = SQL + " where dkey  = 'LightCode'"
        '        SQL = SQL + " and cat = 5001"
        '        SQL = SQL + " and data  ='" + DYKKColorCode.Value + "'"

        '        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        '        If DBAdapter1.Rows.Count > 0 Then

        '        Else
        '            If Q = 0 Then
        '                Message = Message + "\n" + "只有Q字樣選項才能選螢光!"
        '                DYKKColorCode.Value = ""
        '                isOK = False
        '            End If
        '        End If
        '    Else
        '        If Q > 0 Then
        '            Message = Message + "\n" + "有Q字樣選項只能選螢光!"
        '            DYKKColorCode.Value = ""
        '            isOK = False
        '        End If
        '    End If


        'End If






        If Not isOK Then

            uJavaScript.PopMsg(Me, Message)



        End If




        Return isOK


    End Function

    Sub CheckDuplicateNo()

        '檢查是否有重覆依賴的編號
        Dim SQL As String
        SQL = "select no,customer,customercode,buyer,buyercode,customercolor,customercolorcode,overseaykkcode,pantonecode devyear,devseason "
        SQL = SQL + " from F_NewColorUAKIPLING"
        SQL = SQL + " where  customercode='" + DCustomerCode.Text + "'"
        SQL = SQL + " and  buyercode='" + DBuyerCode.Text + "'"
        SQL = SQL + " and customercolor='" + YKK.ReplaceString(DCustomerColor.Text) + "'"
        SQL = SQL + " and  customercolorcode='" + YKK.ReplaceString(DCustomerColorCode.Text) + "'"
        SQL = SQL + " and overseaykkcode='" + YKK.ReplaceString(DOverseaYKKCode.Text) + "'"
        SQL = SQL + " and pantonecode='" + YKK.ReplaceString(DPANTONECode.Text) + "'"
        SQL = SQL + " and  devyear='" + DDevYear.SelectedValue + "'"
        SQL = SQL + " and devseason='" + DDevSeason.SelectedValue + "'"

        Dim NoStr As String = ""
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter1.Rows.Count > 0 Then
            For Each dtr As Data.DataRow In DBAdapter1.Rows
                If NoStr = "" Then
                    NoStr = dtr("No")
                Else
                    NoStr = NoStr + "," + dtr("No")
                End If

            Next
        End If

        If NoStr <> "" Then
            DDuplicateNo.Text = NoStr
            uJavaScript.PopMsg(Me, "有重覆依賴的編號!")
            ' If MsgBox("有重覆依賴的編號是否繼續執行?", MsgBoxStyle.OkCancel + MsgBoxStyle.MsgBoxSetForeground, "檢查") <> MsgBoxResult.Ok Then
            '不執行
            'Message = "有重覆依賴的編號"
            isOK = False

            'End If
            'MsgBox("有重覆依賴的編號是否繼續執行?")

        End If





    End Sub

    Sub MsgBox(ByVal text As String)
        'Dim scriptstr As String
        'scriptstr = "<script language=javascript>" + Chr(10) _
        '+ "confirm(""" + text + """)" + Chr(10) _
        '+ "</script>"
        'Response.Write(scriptstr)
        'Response.Write("<script language='javascript'>if(confirm(""" + text + """)==true)function1();else function2();</script>")
        'Response.Write("<script language=javascript>if (confirm('你確定要繼續輸出文字？')==false) {return fale;};</script>")
    End Sub


    Protected Sub DCustomerCheck_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DCustomerCheck.CheckedChanged
        If DCustomerCheck.Checked Then
            DFactoryCheck.Checked = False
            DVCACheck.Checked = False
            CheckNGColor()


        End If
    End Sub

    Protected Sub DNOCCS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DNOCCS.CheckedChanged
        CheckDnoCCS()

    End Sub

    Protected Sub DFactoryCheck_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFactoryCheck.CheckedChanged
        If DFactoryCheck.Checked Then
            DCustomerCheck.Checked = False
            DVCACheck.Checked = False
            CheckNGColor()

        End If
    End Sub

    Protected Sub DVCACheck_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DVCACheck.CheckedChanged
        If DVCACheck.Checked Then
            DCustomerCheck.Checked = False
            DFactoryCheck.Checked = False
            CheckNGColor()


        End If
    End Sub

    Sub CheckDnoCCS()
        If DNOCCS.Checked = True Then
            DNOCCSReason.BackColor = Color.GreenYellow
            ShowRequiredFieldValidator("DNOCCSReasonRqd", "DNOCCSReason", "異常：需輸入無法CCS原因")
            DNOCCSReason.ReadOnly = False
        Else
            DNOCCSReason.Text = ""
            DNOCCSReason.BackColor = Color.LightGray
            DNOCCSReason.ReadOnly = True


        End If


    End Sub

    Sub CheckNGColor()
        If DVCACheck.Checked = True Then

            DCustomerNGColor.BackColor = Color.Yellow
        Else


            DCustomerNGColor.BackColor = Color.GreenYellow
            '  ShowRequiredFieldValidator("DCustomerNGColorRqd", "DCustomerNGColor", "異常：需輸入客戶否決之YKK色號")
            DCustomerNGColor.Visible = True


        End If


    End Sub

    Sub CopyNo()
        'COPY 表單
        DFactoryCheck.Checked = False
        DVCACheck.Checked = False
        DCustomerCheck.Checked = False

        Dim NO1, sql As String
        NO1 = DNO1.Text
        If NO1 <> "" Then
            sql = "Select * From V_NewColor "
            sql = sql & " Where no  =  '" & NO1 & "'"

            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

            If DBAdapter1.Rows.Count > 0 Then


                If DBAdapter1.Rows(0).Item("CustomerCheck") = 1 Then '客戶自核
                    DCustomerCheck.Checked = True


                End If
                If DBAdapter1.Rows(0).Item("FactoryCheck") = 1 Then '工廠自核
                    DFactoryCheck.Checked = True


                End If
                If DBAdapter1.Rows(0).Item("VCACheck") = 1 Then '工廠自核
                    DVCACheck.Checked = True

                End If



                DCustomer.Text = DBAdapter1.Rows(0).Item("Customer")                         'No
                DCustomerCode.Text = DBAdapter1.Rows(0).Item("CustomerCode")                         'No
                DBuyer.Text = DBAdapter1.Rows(0).Item("Buyer")                         'No
                DBuyerCode.Text = DBAdapter1.Rows(0).Item("BuyerCode")
                DCustomerColor.Text = DBAdapter1.Rows(0).Item("CustomerColor")                         'No3
                DCustomerColorCode.Text = DBAdapter1.Rows(0).Item("CustomerColorCode")
                DOverseaYKKCode.Text = DBAdapter1.Rows(0).Item("OverseaYKKCode")
                DPANTONECode.Text = DBAdapter1.Rows(0).Item("PANTONECode")
                SetFieldData("DevYear", DBAdapter1.Rows(0).Item("DevYear"))    '年
                SetFieldData("DevSeason", DBAdapter1.Rows(0).Item("DevSeason"))    '季

                If Mid(DBAdapter1.Rows(0).Item("ReceiveDate").ToString, 1, 4) = "1900" Then
                    DReceiveDate.Text = ""
                Else
                    DReceiveDate.Text = DBAdapter1.Rows(0).Item("ReceiveDate")               '下單時間
                End If


                SetFieldData("ColorLight1", DBAdapter1.Rows(0).Item("ColorLight1"))    '類別1
                SetFieldData("ColorLight2", DBAdapter1.Rows(0).Item("ColorLight2"))    '類別2
                SetFieldData("ColorLight3", DBAdapter1.Rows(0).Item("ColorLight3"))    '類別2

                If Mid(DBAdapter1.Rows(0).Item("DeliveryDate").ToString, 1, 4) = "1900" Then
                    DDeliveryDate.Value = ""
                Else
                    DDeliveryDate.Value = DBAdapter1.Rows(0).Item("DeliveryDate")               '下單時間
                End If

                DDeliveryDate.Value = ""


                DDuplicateNo.Text = DBAdapter1.Rows(0).Item("DuplicateNo")
                If DBAdapter1.Rows(0).Item("NOCCS") = 1 Then
                    DNOCCS.Checked = True
                End If
                DNOCCSReason.Text = DBAdapter1.Rows(0).Item("NOCCSReason")
                DCustomerNGColor.Text = DBAdapter1.Rows(0).Item("CustomerNGColor")
                DRemark.Text = DBAdapter1.Rows(0).Item("Remark")
                SetFieldData("CustomerSample", DBAdapter1.Rows(0).Item("CustomerSample"))    '類別2
            End If


        End If
    End Sub


    Protected Sub DNO1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DNO1.TextChanged
        CopyNo()
    End Sub



End Class

