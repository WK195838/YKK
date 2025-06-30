Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
Imports System.IO

 


Partial Class DTMW_COMBISheet01
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
    Dim oWaves As New WAVES.CommonService

    Dim uJavaScript As New Utility.JScript
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
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

            '上傳資料檢查及顯示訊息

        End If
        ChangeVisible()
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

        wbFormSno = Request.QueryString("pFormSno")  '連續起單用流水號
        wbStep = Request.QueryString("pStep")        '連續起單用工程代碼
        wApplyID = Request.QueryString("pApplyID")  '申請者ID

        wAgentID = Request.QueryString("pAgentID")  '被代理人ID
        'Add-Start by Joy 2009/11/20(2010行事曆對應)
        '申請者-群組行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
   
        wUserID = Request.QueryString("pUserID")


        'wUserID = Request.QueryString("UserID")
        wDecideCalendar = oCommon.GetCalendarGroup(wUserID)

        Response.Cookies("PGM").Value = "DTMW_COMBISheet01.aspx"
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

        SQL = "Select * From F_COMBISheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then


           



            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            '   DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
            DDate.Text = DBAdapter1.Rows(0).Item("Date")                         'No 
            DDepoName.Text = DBAdapter1.Rows(0).Item("DepoName")                         'No
            DName.Text = DBAdapter1.Rows(0).Item("Name")                         'No
            DCustomer.Text = DBAdapter1.Rows(0).Item("Customer")                         'No
            DCustomerCode.Text = DBAdapter1.Rows(0).Item("CustomerCode")                         'No
            DBuyer.Text = DBAdapter1.Rows(0).Item("Buyer")                         'No
            DBuyerCode.Text = DBAdapter1.Rows(0).Item("BuyerCode")
      
   


            DDuplicateNo.Text = DBAdapter1.Rows(0).Item("DuplicateNo")
          

         

            SetFieldData("YKKColorType", DBAdapter1.Rows(0).Item("YKKColorType"))

            DYKKColorCode.Value = DBAdapter1.Rows(0).Item("YKKColorCode")

            SetFieldData("COMBIItem", DBAdapter1.Rows(0).Item("COMBIITem"))

            If DCOMBIItem.SelectedValue = "塑鋼 VS (一般)" Or DCOMBIItem.SelectedValue = "塑鋼VT810 / VT108" Then
                DITEMNAME1.Text = DCOMBIItem.SelectedValue
            End If
            'VF 
            DVFLTape.Text = DBAdapter1.Rows(0).Item("VFLTape")
            DVFLChain.Text = DBAdapter1.Rows(0).Item("VFLChain")
            DVFRChain.Text = DBAdapter1.Rows(0).Item("VFRChain")
            DVFRTape.Text = DBAdapter1.Rows(0).Item("VFRTape")

            'VF & MF
            DVFMLTape.Text = DBAdapter1.Rows(0).Item("VFMLTape")
            DVFMLChain.Text = DBAdapter1.Rows(0).Item("VFMLChain")
            DVFMRChain.Text = DBAdapter1.Rows(0).Item("VFMRChain")
            DVFMRTape.Text = DBAdapter1.Rows(0).Item("VFMRTape")

            'VF & PF
            DPFMFLTape.Text = DBAdapter1.Rows(0).Item("PFMFLTape")
            DPFMFRTape.Text = DBAdapter1.Rows(0).Item("PFMFRTape")
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
            '
            '遲納管理
            If dtFlow.Rows(0)("Delay") = 1 Then
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
                    Top = 750
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
            Top = 620
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

        BCustomer.Attributes.Add("onClick", "GetCustomer()") '找客戶
        BBuyer.Attributes.Add("onClick", "GetBuyer()") '找buyer

        BYKKColorCode.Attributes.Add("onClick", "GetYKKColorCode()") '找buyer
        BCopySheet.Attributes.Add("onClick", "CopyCombiSheet()") '找buyer

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
                    Top = 800
                Else
                    If DDelay.Visible = True Then
                        Top = 800
                    Else
                        Top = 800
                    End If
                End If
            End If
        Else
            Top = 800

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


        Dim SQL As String
        Dim InputData1 As String
        Dim InputData2 As String


        InputData1 = D1.Text
        InputData2 = D2.Text



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



        '項目
        Select Case FindFieldInf("COMBIItem")
            Case 0  '顯示
                DCOMBIItem.BackColor = Color.LightGray
                DCOMBIItem.Visible = True

            Case 1  '修改+檢查
                DCOMBIItem.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCOMBIItemRqd", "DCOMBIItem", "異常：需輸入項目")
                DCOMBIItem.Visible = True
            Case 2  '修改
                DCOMBIItem.BackColor = Color.Yellow
                DCOMBIItem.Visible = True
            Case Else   '隱藏
                DCOMBIItem.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("COMBIItem", "ZZZZZZ")


        'VFLTape
        Select Case FindFieldInf("VFLTape")
            Case 0  '顯示
                DVFLTape.BackColor = Color.LightGray
                DVFLTape.ReadOnly = True
                DVFLTape.Visible = True
            Case 1  '修改+檢查
                DVFLTape.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVFLTapeRqd", "DVFLTape", "異常：需輸入" + DCOMBIItem.SelectedValue + "左布帶")
                DVFLTape.Visible = True

            Case 2  '修改
                DVFLTape.BackColor = Color.Yellow
                DVFLTape.Visible = True
            Case Else   '隱藏
                DVFLTape.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("VFLTape", "ZZZZZZ")


        'VFLChain
        Select Case FindFieldInf("VFLChain")
            Case 0  '顯示
                DVFLChain.BackColor = Color.LightGray
                DVFLChain.ReadOnly = True
                DVFLChain.Visible = True
            Case 1  '修改+檢查
                DVFLChain.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVFLChainRqd", "DVFLChain", "異常：需輸入" + DCOMBIItem.SelectedValue + "左齒")
                DVFLChain.Visible = True

            Case 2  '修改
                DVFLChain.BackColor = Color.Yellow
                DVFLChain.Visible = True
            Case Else   '隱藏
                DVFLChain.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("VFLChain", "ZZZZZZ")


        'VFRChain
        Select Case FindFieldInf("VFRChain")
            Case 0  '顯示
                DVFRChain.BackColor = Color.LightGray
                DVFRChain.ReadOnly = True
                DVFRChain.Visible = True
            Case 1  '修改+檢查
                DVFRChain.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVFRChainRqd", "DVFRChain", "異常：需輸入" + DCOMBIItem.SelectedValue + "右齒")
                DVFRChain.Visible = True

            Case 2  '修改
                DVFRChain.BackColor = Color.Yellow
                DVFRChain.Visible = True
            Case Else   '隱藏
                DVFRChain.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("VFRChain", "ZZZZZZ")


        'VFRTape
        Select Case FindFieldInf("VFRTape")
            Case 0  '顯示
                DVFRTape.BackColor = Color.LightGray
                DVFRTape.ReadOnly = True
                DVFRTape.Visible = True
            Case 1  '修改+檢查
                DVFRTape.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVFRTapeRqd", "DVFRTape", "異常：需輸入" + DCOMBIItem.SelectedValue + "右布帶")
                DVFRTape.Visible = True

            Case 2  '修改
                DVFRTape.BackColor = Color.Yellow
                DVFRTape.Visible = True
            Case Else   '隱藏
                DVFRTape.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("VFRTape", "ZZZZZZ")


        'VFMRTape
        Select Case FindFieldInf("VFMLTape")
            Case 0  '顯示
                DVFMLTape.BackColor = Color.LightGray
                DVFMLTape.ReadOnly = True
                DVFMLTape.Visible = True
            Case 1  '修改+檢查
                DVFMLTape.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVFMLTapeRqd", "DVFMLTape", "異常：需輸入塑鋼 VS (金屬齒色)左布帶")
                DVFMLTape.Visible = True

            Case 2  '修改
                DVFMLTape.BackColor = Color.Yellow
                DVFMLTape.Visible = True
            Case Else   '隱藏
                DVFMLTape.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("VFMLTape", "ZZZZZZ")



        'VFMLChain
        Select Case FindFieldInf("VFMLChain")
            Case 0  '顯示
                DVFMLChain.BackColor = Color.LightGray
                DVFMLChain.ReadOnly = True
                DVFMLChain.Visible = True
            Case 1  '修改+檢查
                DVFMLChain.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVFMLChainRqd", "DVFMLChain", "異常：需輸入塑鋼 VS (金屬齒色)左齒")
                DVFMLChain.Visible = True

            Case 2  '修改
                DVFMLChain.BackColor = Color.Yellow
                DVFMLChain.Visible = True
            Case Else   '隱藏
                DVFMLChain.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("VFMLChain", "ZZZZZZ")



        'VFMRChain
        Select Case FindFieldInf("VFMRChain")
            Case 0  '顯示
                DVFMRChain.BackColor = Color.LightGray
                DVFMRChain.ReadOnly = True
                DVFMRChain.Visible = True
            Case 1  '修改+檢查
                DVFMRChain.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVFMRChainRqd", "DVFMRChain", "異常：需輸入塑鋼 VS (金屬齒色)右齒")
                DVFMRChain.Visible = True

            Case 2  '修改
                DVFMRChain.BackColor = Color.Yellow
                DVFMRChain.Visible = True
            Case Else   '隱藏
                DVFMRChain.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("VFMRChain", "ZZZZZZ")



        'VFMRTape
        Select Case FindFieldInf("VFMRTape")
            Case 0  '顯示
                DVFMRTape.BackColor = Color.LightGray
                DVFMRTape.ReadOnly = True
                DVFMRTape.Visible = True
            Case 1  '修改+檢查
                DVFMRTape.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVFMRTapeRqd", "DVFMRTape", "異常：需輸入塑鋼 VS (金屬齒色)右布帶")
                DVFMRTape.Visible = True

            Case 2  '修改
                DVFMRTape.BackColor = Color.Yellow
                DVFMRTape.Visible = True
            Case Else   '隱藏
                DVFMRTape.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("VFMRTape", "ZZZZZZ")





        'PFMFLTape
        Select Case FindFieldInf("PFMFLTape")
            Case 0  '顯示
                DPFMFLTape.BackColor = Color.LightGray
                DPFMFLTape.ReadOnly = True
                DPFMFLTape.Visible = True
            Case 1  '修改+檢查
                DPFMFLTape.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPFMFLTapeRqd", "DPFMFLTape", "異常：需輸入尼龍 PF/金屬MF 左布帶")
                DPFMFLTape.Visible = True

            Case 2  '修改
                DPFMFLTape.BackColor = Color.Yellow
                DPFMFLTape.Visible = True
            Case Else   '隱藏
                DPFMFLTape.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("PFMFLTape", "ZZZZZZ")



        'PFMFRTape
        Select Case FindFieldInf("PFMFRTape")
            Case 0  '顯示
                DPFMFRTape.BackColor = Color.LightGray
                DPFMFRTape.ReadOnly = True
                DPFMFRTape.Visible = True
            Case 1  '修改+檢查
                DPFMFRTape.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPFMFRTapeRqd", "DPFMFRTape", "異常：需輸入尼龍 PF/金屬MF 右布帶")
                DPFMFRTape.Visible = True

            Case 2  '修改
                DPFMFRTape.BackColor = Color.Yellow
                DPFMFRTape.Visible = True
            Case Else   '隱藏
                DPFMFRTape.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("PFMFRTape", "ZZZZZZ")

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

        '項目
        If pFieldName = "COMBIItem" Then
            DCOMBIItem.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCOMBIItem.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'COMBIItem'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DCOMBIItem.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True

                    DCOMBIItem.Items.Add(ListItem1)
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
    Protected Sub BSAVE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSAVE.Click

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.Click
        If OK() = True Then

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
    Protected Sub BNG1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNG1.Click
        'If Request.Cookies("RunBNG1").Value = True Then
        DisabledButton()   '停止Button運作
        FlowControl("NG1", 1, "2")


        ' End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG2按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BNG2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNG2.Click
        ' If Request.Cookies("RunBNG2").Value = True Then

        If wStep = 10 Then
            If OK() = True Then
                DisabledButton()   '停止Button運作
                FlowControl("NG2", 2, "3")

            End If
        Else
            DisabledButton()   '停止Button運作
            FlowControl("NG2", 2, "3")

        End If
        ' End If
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
                ErrCode = oCommon.CommissionNo("005007", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
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
                            ModifyTranData(pFun, pSts)

                            '流程資料結束
                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDepo,wUserID, pRunNextStep)
                            '
                            RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, wUserID, pRunNextStep)
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
                    '只有螢光的才CC給怡蓉，若不是就直接給起單者
                    If wStep = 10 And pAction = 0 And DYKKColorType.SelectedValue <> "螢光" Then
                        pAction = 3
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

                If wbFormSno = 0 And wbStep < 3 And DReRegister.Checked = True Then    '判斷是否起單

                    uJavaScript.PopMsg(Me, "已成功送出您所填寫之申請單，請繼續申請！")
                    EnabledButton()                             '起動Button運作

                    DNo.Text = ""
                    wFormSno = wbFormSno
                    wStep = wbStep


                Else
                  
                 
                    URL = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                                               "&pUserID=" & wUserID & "&pApplyID=" & wApplyID

                    Response.Redirect(URL)
                End If

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
        SQl = "Insert into F_COMBISheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno,"
        SQl = SQl + "No, DepoName,Date,Name,"  '1~5                
        SQl = SQl + "Customer,CustomerCode,Buyer,BuyerCode,"   '6~10              
        SQl = SQl + "DuplicateNo,COMBIItem,YKKColorType,YKKColorCode,"  '11~15 
        SQl = SQl + "VFLTape,VFLChain,VFRChain,VFRTape," '36
        SQl = SQl + "VFMLTape,VFMLChain,VFMRChain,VFMRTape," '36
        SQl = SQl + "PFMFLTape,PFMFRTape," '36
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime) "
        SQl = SQl + "VALUES( "

        SQl = SQl + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        SQl = SQl + " '" + NowDateTime + "', "        '結案日
        SQl = SQl + " '005007', "                     '表單代號
        SQl = SQl + " '" + CStr(NewFormSno) + "', "   '表單流水號

        SQl = SQl + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO  1
        SQl = SQl + " '" + DDepoName.Text + "', "                '部門2
        SQl = SQl + " '" + DDate.Text + "', "                '日期3
        SQl = SQl + " '" + DName.Text + "', "                '姓名4
      
        SQl = SQl + " '" + YKK.ReplaceString(DCustomer.Text.ToUpper) + "', "                '客戶8
        SQl = SQl + " '" + DCustomerCode.Text.ToUpper + "', "                '客戶9
        SQl = SQl + " '" + YKK.ReplaceString(DBuyer.Text.ToUpper) + "', "                'BUYER10
        SQl = SQl + " '" + DBuyerCode.Text.ToUpper + "', "                'BUYER11
     
        SQl = SQl + " '" + DDuplicateNo.Text.ToUpper + "', "                ' 收受日期21
        SQl = SQl + " '" + DCOMBIItem.SelectedValue.ToUpper + "', "                ' 收受日期21   
        SQl = SQl + " '" + DYKKColorType.SelectedValue.ToUpper + "', "                ' YKK色別31
        SQl = SQl + " '" + DYKKColorCode.Value.ToUpper + "', "                ' YKK色號32

        SQl = SQl + " '" + DVFLTape.Text.ToUpper + "', "                ' YKK色別31

        SQl = SQl + " '" + DVFLChain.Text.ToUpper + "', "                ' YKK色別31

        '20151231  可輸入左齒，自動帶出右齒=左齒色號
        If DCOMBIItem.SelectedValue = "塑鋼VT810 / VT108" Then
            SQl = SQl + " '" + DVFLChain.Text.ToUpper + "', "                ' YKK色別31
        Else
            SQl = SQl + " '" + DVFRChain.Text.ToUpper + "', "                ' YKK色別31
        End If




        SQl = SQl + " '" + DVFRTape.Text.ToUpper + "', "                ' YKK色別31

        SQl = SQl + " '" + DVFMLTape.Text.ToUpper + "', "                ' YKK色別31
        SQl = SQl + " '" + DVFMLChain.Text.ToUpper + "', "                ' YKK色別31
        SQl = SQl + " '" + DVFMRChain.Text.ToUpper + "', "                ' YKK色別31
        SQl = SQl + " '" + DVFMRTape.Text.ToUpper + "', "                ' YKK色別31

        SQl = SQl + " '" + DPFMFLTape.Text.ToUpper + "', "                ' YKK色別31
        SQl = SQl + " '" + DPFMFRTape.Text.ToUpper + "', "                ' YKK色別31

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

        SQL = " Update F_COMBISheet"
        SQL = SQL + " Set "
        If pFun <> "SAVE" Then
            SQL = SQL + " Sts = '" & pSts & "',"
            SQL = SQL + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQL = SQL + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"
        SQL = SQL + " Date = N'" & DDate.Text & "',"
        SQL = SQL + " DepoName = N'" & DDepoName.Text & "',"
        SQL = SQL + " Name = N'" & DName.Text & "',"
    
        SQL = SQL + " Customer = N'" & YKK.ReplaceString(DCustomer.Text.ToUpper) & "',"
        SQL = SQL + " CustomerCode = N'" & DCustomerCode.Text.ToUpper & "',"
        SQL = SQL + " Buyer = N'" & YKK.ReplaceString(DBuyer.Text) & "',"
        SQL = SQL + " BuyerCode = N'" & DBuyerCode.Text & "',"

        SQL = SQL + " DuplicateNo = N'" & DDuplicateNo.Text.ToUpper & "',"
        SQL = SQL + " COMBIItem = N'" & DCOMBIItem.SelectedValue.ToUpper & "',"
        SQL = SQL + " YKKColorType = N'" & DYKKColorType.SelectedValue.ToUpper & "',"
        SQL = SQL + " YKKColorCode = N'" & DYKKColorCode.Value.ToUpper & "',"

        SQL = SQL + " VFLTape = N'" & DVFLTape.Text.ToUpper & "',"
        SQL = SQL + " VFLChain = N'" & DVFLChain.Text.ToUpper & "',"

        '20151231  可輸入左齒，自動帶出右齒=左齒色號
        If DCOMBIItem.SelectedValue = "塑鋼VT810 / VT108" Then
            SQL = SQL + " VFRChain = N'" & DVFLChain.Text.ToUpper & "',"
        Else
            SQL = SQL + " VFRChain = N'" & DVFRChain.Text.ToUpper & "',"
        End If

        SQL = SQL + " VFRTape = N'" & DVFRTape.Text.ToUpper & "',"

        SQL = SQL + " VFMLTape = N'" & DVFMLTape.Text.ToUpper & "',"
        SQL = SQL + " VFMLChain = N'" & DVFMLChain.Text.ToUpper & "',"
        SQL = SQL + " VFMRChain = N'" & DVFMRChain.Text.ToUpper & "',"
        SQL = SQL + " VFMRTape = N'" & DVFMRTape.Text.ToUpper & "',"

        SQL = SQL + " PFMFLTape = N'" & DPFMFLTape.Text.ToUpper & "',"
        SQL = SQL + " PFMFRTape = N'" & DPFMFRTape.Text.ToUpper & "',"


        SQL = SQL + " ModifyUser = '" & wUserID & "',"
        SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
        SQL = SQL + " Where Formsno ='" + Str(wFormSno) + "'"
        uDataBase.ExecuteNonQuery(SQL)

        If wStep = 10 And pFun = "OK" Then
            If (DCOMBIItem.SelectedValue = "塑鋼 VS (一般)") Or (DCOMBIItem.SelectedValue = "尼龍 PF / 金屬MF") Then

                'Master 建檔後轉資料到WINS COLOR STRUCTURE
                SQL = "insert into R_NCFA300WK (no,sts,PDPCW3,CTBNW3,PACCW3,ST1CW3,ST6CW3,SCSVW3,CCLCW3,UIDCW3,PRGCW3,DEVCW3,RADUW3,RADTW3,RUPUW3,RUPTW3,RCMCW3, PRBFW3, CHKFW3,WKDTW3)"
                SQL = SQL + " select NO,STS, PDPCW3,CTBNW3,PACCW3,ST1CW3,ST6CW3,SCSVW3,"
                SQL = SQL + " isnull(RIGHT(REPLICATE(' ', 5) + CAST(CCLCW3 as NVARCHAR), 5),'') as CCLCW3,"
                SQL = SQL + " UIDCW3,PRGCW3,DEVCW3,RADUW3,RADTW3,RUPUW3,RUPTW3,RCMCW3, PRBFW3, CHKFW3,WKDTW3 from ("
                SQL = SQL + " select NO,'0' AS STS,'01' as PDPCW3 ,CTBNW3,ykkcolorcode as PACCW3,ST1CW3,ST6CW3,SCSVW3,"
                SQL = SQL + " Case when CCLCW3 ='YKKColorCode' then YKKColorCode "
                SQL = SQL + " when CCLCW3 ='TSCC' then CCLCW3"
                SQL = SQL + " when CCLCW3 ='PFMFLTape' then PFMFLTape"
                SQL = SQL + " when CCLCW3 ='PFMFRTape' then PFMFRTape"
                SQL = SQL + " when CCLCW3 ='C5+' then CCLCW3"
                SQL = SQL + " when CCLCW3 ='VFLChain' then VFLChain"
                SQL = SQL + " when CCLCW3 ='VFLTape' then VFLTape"
                SQL = SQL + " when CCLCW3 ='VFRChain' then VFRChain"
                SQL = SQL + " when CCLCW3 ='VFRTape' then VFRTape"
                SQL = SQL + " end as CCLCW3 ,'' as UIDCW3,'' as PRGCW3,'' as DEVCW3,convert(char(10),getdate(),112)as  RADUW3,"
                SQL = SQL + " substring(convert(char(10),getdate(),108),1,2)+substring(convert(char(10),getdate(),108),4,2)+substring(convert(char(10),getdate(),108),7,2) as RADTW3,"
                SQL = SQL + " '' as RUPUW3,'' as RUPTW3,'' as RCMCW3,'' as PRBFW3,'' as CHKFW3,'' as WKDTW3"
                SQL = SQL + " from F_COMBISHEET a,R_NewColorST_01 b"
                SQL = SQL + " where no = '" + DNo.Text + "' AND TYPE ='" + DCOMBIItem.SelectedValue + "'"
                SQL = SQL + " )a"
                uDataBase.ExecuteNonQuery(SQL)



                If DCOMBIItem.SelectedValue = "塑鋼 VS (一般)" Then
                    Dim strL = DVFLChain.Text
                    Dim strR = DVFRChain.Text
                    '從WINS 傳回資料
                    SQL = " update b"
                    SQL = SQL + " set b.CCLCW3 =ISNULL(A.CCLCW3,'')"
                    SQL = SQL + " from ("
                    SQL = SQL + " select ctbnw3,st1cw3,st6cw3,scsvw3,"
                    SQL = SQL + " case when cclcw3 ='CCP1XX' then CCP1XX"
                    SQL = SQL + " when cclcw3 ='CPO1XX' then CPO1XX"
                    SQL = SQL + " when cclcw3 ='CPO2XX' then CPO2XX"
                    SQL = SQL + " when cclcw3 ='CPO3XX' then CPO3XX"
                    SQL = SQL + " when cclcw3 ='CPS1XX' then CPS1XX"
                    SQL = SQL + " when cclcw3 ='CPS2XX' then CPS2XX"
                    SQL = SQL + " when cclcw3 ='CPS3XX' then CPS3XX"
                    SQL = SQL + " when cclcw3 ='CPS4XX' then CPS4XX"
                    SQL = SQL + " when cclcw3 ='CPB1XX' then CPB1XX"
                    SQL = SQL + " when cclcw3 ='CPB2XX' then CPB2XX"
                    SQL = SQL + " when cclcw3 ='CPT1XX' then CPT1XX"
                    SQL = SQL + " when cclcw3 ='CPT2XX' then CPT2XX"
                    SQL = SQL + " when cclcw3 ='CSP1XX' then CSP1XX"
                    SQL = SQL + " when cclcw3 ='CSP2XX' then CSP2XX"
                    SQL = SQL + " when cclcw3 ='CSP3XX' then CSP3XX"
                    SQL = SQL + " when cclcw3 ='CSP4XX' then CSP4XX end as CCLCw3"
                    SQL = SQL + " from MST_COLOR_STRUCTURE A,R_NewColorST_01 b"
                    SQL = SQL + " where paccxx ='" + strL.Padleft(5, " ") + "'"
                    SQL = SQL + " and a.ctbnxx=b.ctbnw3"
                    SQL = SQL + " AND TYPE ='" + DCOMBIItem.SelectedValue + "'"
                    SQL = SQL + " and left(cclcw3,1)='c'"
                    SQL = SQL + " and len(cclcw3)=6"
                    SQL = SQL + " )a,R_NCFA300WK b where "
                    SQL = SQL + " a.st1cw3 = b.st1cw3 And a.st6cw3 = b.st6cw3 And a.scsvw3 = b.scsvw3"
                    SQL = SQL + " and no =  '" + DNo.Text + "' "
                    SQL = SQL + " and a.ctbnw3 =b.ctbnw3"
                    uDataBase.ExecuteNonQuery(SQL)

                    SQL = " update b "
                    SQL = SQL + " set b.CCLCW3 =ISNULL(A.CCLCW3,'')"
                    SQL = SQL + " from ("
                    SQL = SQL + " select ctbnw3,st1cw3,st6cw3,scsvw3,"
                    SQL = SQL + " case when cclcw3 ='CCP2XX' then CCP2XX  end as CCLCw3"
                    SQL = SQL + " from MST_COLOR_STRUCTURE A,R_NewColorST_01 b"
                    SQL = SQL + " where paccxx ='" + strR.Padleft(5, " ") + "'"
                    SQL = SQL + " and a.ctbnxx=b.ctbnw3"
                    SQL = SQL + " AND TYPE ='" + DCOMBIItem.SelectedValue + "'"
                    SQL = SQL + " and left(cclcw3,1)='c'"
                    SQL = SQL + " and len(cclcw3)=6"
                    SQL = SQL + " and st1cw3 =5 and st6cw3 ='F' and scsvw3 =2"
                    SQL = SQL + " )a,R_NCFA300WK b where "
                    SQL = SQL + " a.st1cw3 = b.st1cw3 And a.st6cw3 = b.st6cw3 And a.scsvw3 = b.scsvw3"
                    SQL = SQL + " and no =  '" + DNo.Text + "' "
                    SQL = SQL + " and a.ctbnw3 =b.ctbnw3"
                    uDataBase.ExecuteNonQuery(SQL)
                End If
               

                Dim Code As String = ""
                Code = oWaves.NewColor2Wings(DNo.Text) 'CALL 轉檔程式
                Me.dataTableToCsv(DNo.Text) '轉存excel

                If Code <> "0" Then
                    uJavaScript.PopMsg(Me, "匯入WINS異常，請洽電腦室")
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
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
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
        Str = "G" + CStr(DateTime.Now.Year) + Str2
        '年
        'Set單號
        '當月份筆數有幾筆  20150414 Modify by Jessica
        Dim sql As String
        sql = " select isnull(max(convert(int,substring(no,8,4))),0)cun  from  F_COMBISheet  "
        sql = sql + " where  left(convert(char(10),date ,111),7) = left(convert(char(10), getdate(),111),7)"
        Dim dt1 As DataTable = uDataBase.GetDataTable(sql)
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
        If wStep = 1 Then
            Top = 620
        Else
            Top = 800
        End If

        SetControlPosition()
        '  ShowFormData()

        Dim SQL As String
        Dim DOUBLENo As Integer = 0
        Dim Errcode As Integer = 0



        '檢查是否有重覆依賴的編號
        If wStep = 1 Or wStep = 500 Then

            If DCOMBIItem.SelectedValue = "塑鋼VT810 / VT108" Then
                SQL = " select distinct  CombiItem,colorcode,LTape,LChain,RChain,RTape  from ( "
                SQL = SQL + " select  CombiItem,colorTable,colorcode,LTape,LChain,RChain,RTape  from ("
                '   SQL = SQL + " select '' as CombiItem, convert(char,ctbnxx) as colorTable,paccxx as colorcode,  ltrim(cta1xx) as  LTape  , ltrim(ccp1xx) as LChain ,ltrim(ccp2xx)  as RChain , ltrim(cta2xx) as RTape     from MST_color_structure "
                '  SQL = SQL + " union all "
                SQL = SQL + " select CombiItem,'' as colorTable, no as colorcode ,vfltape as LTape,vflchain as LChain,vfrchain as RChain,vfrtape as  RTape  from F_COMBISheet "
                SQL = SQL + " where sts not in (2,3)"
                SQL = SQL + " union all"
                SQL = SQL + " select CombiItem,'' as colorTable, no  as colorcode,vfltape as LTape,vflchain as LChain,vfrchain as RChain,vfrtape as  RTape  from F_COMBISheet "
                SQL = SQL + " where sts not in (2,3)"
                SQL = SQL + " union all"
                SQL = SQL + " select  CombiItem, '' as colorTable, no  as colorcode, pfmfltape as LTape,'' as LChain ,'' as RChain,pfmfrtape as RTape from  F_COMBISheet "
                SQL = SQL + " where sts not in (2,3)"
                SQL = SQL + " )a where 1=1 "
            Else
                SQL = " select distinct  CombiItem,colorcode,LTape,LChain,RChain,RTape  from ( "
                SQL = SQL + " select  CombiItem,colorTable,colorcode,LTape,LChain,RChain,RTape  from ("
                SQL = SQL + " select '' as CombiItem, convert(char,ctbnxx) as colorTable,paccxx as colorcode,  ltrim(cta1xx) as  LTape  , ltrim(ccp1xx) as LChain ,ltrim(ccp2xx)  as RChain , ltrim(cta2xx) as RTape     from MST_color_structure where  right(paccxx,1) <> 'v' "
                SQL = SQL + " union all "
                SQL = SQL + " select CombiItem,'' as colorTable, no ,vfltape,vflchain,vfrchain,vfrtape  from F_COMBISheet "
                SQL = SQL + " where sts not in (2,3)"
                SQL = SQL + " union all"
                SQL = SQL + " select CombiItem,'' as colorTable, no,vfmltape,vfmlchain,vfmrchain,vfmrtape  from F_COMBISheet "
                SQL = SQL + " where sts not in (2,3)"
                SQL = SQL + " union all"
                SQL = SQL + " select  CombiItem, '' as colorTable, no, pfmfltape,'' ,'',pfmfrtape from  F_COMBISheet "
                SQL = SQL + " where sts not in (2,3)"
                SQL = SQL + " )a where 1=1 "
            End If


            If DCOMBIItem.SelectedValue = "塑鋼 VS (一般)" Or DCOMBIItem.SelectedValue = "塑鋼VT810 / VT108" Then
                '塑鋼 VS (一般)
                If DVFLTape.Text <> "" Then
                    SQL = SQL + " and LTape = '" + DVFLTape.Text + "'"
                End If
                If DVFLChain.Text <> "" Then
                    SQL = SQL + " and LChain = '" + DVFLChain.Text + "'"
                End If

                If DCOMBIItem.SelectedValue = "塑鋼VT810 / VT108" Then
                    If DVFRChain.Text <> "" Then
                        SQL = SQL + " and RChain = '" + DVFLChain.Text + "'"
                    End If
                Else
                    If DVFRChain.Text <> "" Then
                        SQL = SQL + " and RChain = '" + DVFRChain.Text + "'"
                    End If
                End If


                If DVFRTape.Text <> "" Then
                    SQL = SQL + " and RTape = '" + DVFRTape.Text + "'"
                End If

                SQL = SQL + " and colorTable in ('001','006','')  "
                'and combiitem= '" + DCOMBIItem.SelectedValue + "'"


            ElseIf DCOMBIItem.SelectedValue = "塑鋼 VS (金屬齒色)" Then
                '塑鋼 VS (金屬齒色)-
                If DVFMLTape.Text <> "" Then
                    SQL = SQL + " and LTape = '" + DVFLTape.Text + "'"
                End If
                If DVFMLChain.Text <> "" Then
                    SQL = SQL + " and LChain = '" + DVFMLChain.Text + "'"
                End If
                If DVFMRChain.Text <> "" Then
                    SQL = SQL + " and RChain = '" + DVFMRChain.Text + "'"
                End If
                If DVFMRTape.Text <> "" Then
                    SQL = SQL + " and RTape = '" + DVFMRTape.Text + "'"
                End If
                SQL = SQL + " and colorTable in ('001','006','') "
            Else
                '尼龍PF VS 金屬MF
                If DPFMFLTape.Text <> "" Then
                    SQL = SQL + " and LTape = '" + DPFMFLTape.Text + "'"
                End If
                If DPFMFRTape.Text <> "" Then
                    SQL = SQL + " and RTape = '" + DPFMFRTape.Text + "'"
                End If

                SQL = SQL + " and colorTable in('002','') "
            End If
            SQL = SQL + " and  CombiItem in ('','" + DCOMBIItem.SelectedValue + "')"
            SQL = SQL + " )a"

            If wStep = 500 Then
                SQL = SQL + " where colorcode <> '" + DNo.Text + "'"
            End If

            Dim NoStr As String = ""
            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
            If DBAdapter1.Rows.Count > 0 Then
                For Each dtr As Data.DataRow In DBAdapter1.Rows
                    If NoStr = "" Then
                        NoStr = dtr("colorcode")
                    Else
                        NoStr = NoStr + "," + dtr("colorcode")
                    End If
                Next
                If NoStr = DDuplicateNo.Text Then  '如果重覆的編號一樣就不用再執行一次
                    DOUBLENo = DOUBLENo + 1
                End If
            Else
                DDuplicateNo.Text = "" '如果沒有重覆就清空

            End If

            If NoStr <> "" Then
                DDuplicateNo.Text = NoStr
                'uJavaScript.PopMsg(Me, "有重覆依賴的編號!")
                Message = Message + "\n" + "有重覆依賴的編號請確認!"
                ' If MsgBox("有重覆依賴的編號是否繼續執行?", MsgBoxStyle.OkCancel + MsgBoxStyle.MsgBoxSetForeground, "檢查") <> MsgBoxResult.Ok Then
                '不執行
                'Message = "有重覆依賴的編號"
                isOK = False

                'End If
                'MsgBox("有重覆依賴的編號是否繼續執行?")

            End If


        End If

        'JESSICA 20220728 
        If DVFLTape.Text = "T054R" Or DVFLTape.Text = "T184R" Or DVFLTape.Text = "T355" Or DVFLTape.Text = "T356R" _
        Or DVFLTape.Text = "T502R" Or DVFLTape.Text = "T503R" Or DVFLTape.Text = "T803R" Or DVFLTape.Text = "T841R" Or DVFLTape.Text = "T872R" Then
            Message = DVFLTape.Text & "[不可輸入，請查閱公佈欄-遠東PBR COLOR一覽表]"
            isOK = False
        End If

        If DVFRTape.Text = "T054R" Or DVFRTape.Text = "T184R" Or DVFRTape.Text = "T355" Or DVFRTape.Text = "T356R" _
        Or DVFRTape.Text = "T502R" Or DVFRTape.Text = "T503R" Or DVFRTape.Text = "T803R" Or DVFRTape.Text = "T841R" Or DVFRTape.Text = "T872R" Then
            Message = DVFRTape.Text & "[不可輸入，請查閱公佈欄-遠東PBR COLOR一覽表]"
            isOK = False
        End If

        If DVFMLTape.Text = "T054R" Or DVFMLTape.Text = "T184R" Or DVFMLTape.Text = "T355" Or DVFMLTape.Text = "T356R" _
        Or DVFMLTape.Text = "T502R" Or DVFMLTape.Text = "T503R" Or DVFMLTape.Text = "T803R" Or DVFMLTape.Text = "T841R" Or DVFMLTape.Text = "T872R" Then
            Message = DVFMLTape.Text & "[不可輸入，請查閱公佈欄-遠東PBR COLOR一覽表]"
            isOK = False
        End If

        If DVFMRTape.Text = "T054R" Or DVFMRTape.Text = "T184R" Or DVFMRTape.Text = "T355" Or DVFMRTape.Text = "T356R" _
               Or DVFMRTape.Text = "T502R" Or DVFMRTape.Text = "T503R" Or DVFMRTape.Text = "T803R" Or DVFMRTape.Text = "T841R" Or DVFMRTape.Text = "T872R" Then
            Message = DVFMRTape.Text & "[不可輸入，請查閱公佈欄-遠東PBR COLOR一覽表]"
            isOK = False
        End If


        If DPFMFLTape.Text = "T054R" Or DPFMFLTape.Text = "T184R" Or DPFMFLTape.Text = "T355" Or DPFMFLTape.Text = "T356R" _
               Or DPFMFLTape.Text = "T502R" Or DPFMFLTape.Text = "T503R" Or DPFMFLTape.Text = "T803R" Or DPFMFLTape.Text = "T841R" Or DPFMFLTape.Text = "T872R" Then
            Message = DPFMFLTape.Text & "[不可輸入，請查閱公佈欄-遠東PBR COLOR一覽表]"
            isOK = False
        End If


        If DPFMFRTape.Text = "T054R" Or DPFMFRTape.Text = "T184R" Or DPFMFRTape.Text = "T355" Or DPFMFRTape.Text = "T356R" _
               Or DPFMFRTape.Text = "T502R" Or DPFMFRTape.Text = "T503R" Or DPFMFRTape.Text = "T803R" Or DPFMFRTape.Text = "T841R" Or DPFMFRTape.Text = "T872R" Then
            Message = DPFMFRTape.Text & "[不可輸入，請查閱公佈欄-遠東PBR COLOR一覽表]"
            isOK = False
        End If



        If wStep = 10 Then
            If DYKKColorType.SelectedValue = "" Then
                isOK = False
                Message = Message + "\n" + "異常：請輸入YKK色別!"
            End If
            If DYKKColorCode.Value = "" Then
                isOK = False
                Message = Message + "\n" + "異常：請輸入YKK色號!"
            End If
        End If




        Dim Q As Integer = 0


        If wStep = 10 Then
            '檢查是否有Q 
            If DYKKColorType.SelectedValue = "螢光" Then


                SQL = " select * from   m_referp"
                SQL = SQL + " where dkey  = 'LightCode'"
                SQL = SQL + " and cat = 5001"
                SQL = SQL + " and data  ='" + DYKKColorCode.Value + "'"

                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                If DBAdapter1.Rows.Count > 0 Then

                Else

                    Q = InStr(1, DVFLTape.Text, "Q", 1)
                    Q = Q + InStr(1, DVFLChain.Text, "Q", 1)
                    Q = Q + InStr(1, DVFRChain.Text, "Q", 1)
                    Q = Q + InStr(1, DVFRTape.Text, "Q", 1)

                    Q = Q + InStr(1, DVFMLTape.Text, "Q", 1)
                    Q = Q + InStr(1, DVFMLChain.Text, "Q", 1)
                    Q = Q + InStr(1, DVFMRChain.Text, "Q", 1)
                    Q = Q + InStr(1, DVFMRTape.Text, "Q", 1)

                    Q = Q + InStr(1, DPFMFLTape.Text, "Q", 1)
                    Q = Q + InStr(1, DPFMFRTape.Text, "Q", 1)

                    If Q = 0 Then
                        Message = Message + "\n" + "只有Q字樣選項才能選螢光!"
                        DYKKColorCode.Value = ""
                        isOK = False
                    End If
                End If
            Else

                Q = InStr(1, DVFLTape.Text, "Q", 1)
                Q = Q + InStr(1, DVFLChain.Text, "Q", 1)
                Q = Q + InStr(1, DVFRChain.Text, "Q", 1)
                Q = Q + InStr(1, DVFRTape.Text, "Q", 1)

                Q = Q + InStr(1, DVFMLTape.Text, "Q", 1)
                Q = Q + InStr(1, DVFMLChain.Text, "Q", 1)
                Q = Q + InStr(1, DVFMRChain.Text, "Q", 1)
                Q = Q + InStr(1, DVFMRTape.Text, "Q", 1)

                Q = Q + InStr(1, DPFMFLTape.Text, "Q", 1)
                Q = Q + InStr(1, DPFMFRTape.Text, "Q", 1)

                If Q > 0 Then
                    Message = Message + "\n" + "有Q字樣選項只能選螢光!"
                    DYKKColorCode.Value = ""
                    isOK = False
                End If



            End If



        End If


        '檢查是否有輸入廢止色
        Dim Code(10) As String
        Dim CodeStr(10) As String

        Code(1) = DVFLTape.Text
        Code(2) = DVFLChain.Text
        Code(3) = DVFRChain.Text
        Code(4) = DVFRTape.Text
        Code(5) = DVFMLTape.Text
        Code(6) = DVFMLChain.Text
        Code(7) = DVFMRChain.Text
        Code(8) = DVFMRTape.Text
        Code(9) = DPFMFLTape.Text
        Code(10) = DPFMFRTape.Text

        CodeStr(1) = "塑鋼 VS (一般)-左布帶輸入的資料為廢止色"
        CodeStr(2) = "塑鋼 VS (一般)-左齒輸入的資料為廢止色"
        CodeStr(3) = "塑鋼 VS (一般)-右齒輸入的資料為廢止色"
        CodeStr(4) = "塑鋼 VS (一般)-右布帶輸入的資料為廢止色"
        CodeStr(5) = "塑鋼 VS (金屬齒色)-左布帶輸入的資料為廢止色"
        CodeStr(6) = "塑鋼 VS (金屬齒色)-左齒輸入的資料為廢止色"
        CodeStr(7) = "塑鋼 VS (金屬齒色)-右齒輸入的資料為廢止色"
        CodeStr(8) = "塑鋼 VS (金屬齒色)-右布帶輸入的資料為廢止色"
        CodeStr(9) = "尼龍PF VS 金屬MF-左布帶輸入的資料為廢止色"
        CodeStr(10) = "尼龍PF VS 金屬MF-右布帶輸入的資料為廢止色"

        Dim i, j, k As Integer

        If DCOMBIItem.SelectedValue = "塑鋼 VS (一般)" Or DCOMBIItem.SelectedValue = "塑鋼VT810 / VT108" Then
            k = 1
            j = 4
        ElseIf DCOMBIItem.SelectedValue = "塑鋼 VS (金屬齒色)" Then
            k = 5
            j = 8
        Else
            k = 9
            j = 10
        End If


        For i = 1 To 10
            SQL = " select * from F_combirepeal"
            SQL = SQL + " where  1=1 and colorcode ='" + Code(i) + "'"
            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
            If DBAdapter3.Rows.Count > 0 Then
                For Each dtr As Data.DataRow In DBAdapter3.Rows
                    Message = Message + "\n" + CodeStr(i)
                    isOK = False

                Next

            End If

        Next



        If Not isOK Then
            uJavaScript.PopMsg(Me, Message)
        End If



        Return isOK


    End Function

    Sub CheckDuplicateNo()

        '檢查是否有重覆依賴的編號
        Dim SQL As String = ""
        Dim LTape As String = ""
        Dim LChain As String = ""
        Dim RChain As String = ""
        Dim RTape As String = ""


        SQL = " select colorcode,LTape,LChain,RChain,RTape  from ("
        SQL = SQL + " select paccxx as colorcode, ltrim(cta1xx) as  LTape  , ltrim(ccp1xx) as LChain ,ltrim(ccp2xx)  as RChain , ltrim(cta2xx) as RTape    from MST_color_structure "
        SQL = SQL + " union all "
        SQL = SQL + " select no ,vfltape,vflchain,vfrchain,vfrtape  from F_COMBISheet "
        SQL = SQL + " where sts not in (2,3)"
        SQL = SQL + " union all"
        SQL = SQL + " select  no,vfmltape,vfmlchain,vfmrchain,vfmrtape  from F_COMBISheet "
        SQL = SQL + " where sts not in (2,3)"
        SQL = SQL + " union all"
        SQL = SQL + " select no, pfmfltape,'' ,'',pfmfrtape from  F_COMBISheet "
        SQL = SQL + " where sts not in (2,3)"
        SQL = SQL + " )a where 1=1 "

        If DCOMBIItem.SelectedValue = "塑鋼 VS (一般)" Or DCOMBIItem.SelectedValue = "塑鋼VT810 / VT108" Then
            '塑鋼 VS (一般)
            If DVFLTape.Text <> "" Then
                SQL = SQL + " and LTape = '" + DVFLTape.TextMode + "'"
            End If
            If DVFLChain.Text <> "" Then
                SQL = SQL + " and LChain = '" + DVFLChain.TextMode + "'"
            End If
            If DVFRChain.Text <> "" Then
                SQL = SQL + " and RChain = '" + DVFRChain.TextMode + "'"
            End If
            If DVFRTape.Text <> "" Then
                SQL = SQL + " and RTape = '" + DVFRTape.TextMode + "'"
            End If
        ElseIf DCOMBIItem.SelectedValue = "塑鋼 VS (金屬齒色)" Then
            '塑鋼 VS (金屬齒色)-
            If DVFMLTape.Text <> "" Then
                SQL = SQL + " and LTape = '" + DVFLTape.TextMode + "'"
            End If
            If DVFMLChain.Text <> "" Then
                SQL = SQL + " and LChain = '" + DVFMLChain.TextMode + "'"
            End If
            If DVFMRChain.Text <> "" Then
                SQL = SQL + " and RChain = '" + DVFMRChain.TextMode + "'"
            End If
            If DVFMRTape.Text <> "" Then
                SQL = SQL + " and RTape = '" + DVFMRTape.TextMode + "'"
            End If
        Else
            '尼龍PF VS 金屬MF
            If DPFMFLTape.Text <> "" Then
                SQL = SQL + " and LTape = '" + DPFMFLTape.TextMode + "'"
            End If
            If DVFMRTape.Text <> "" Then
                SQL = SQL + " and RTape = '" + DVFMRTape.TextMode + "'"
            End If

        End If




        Dim NoStr As String = ""
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter1.Rows.Count > 0 Then
            For Each dtr As Data.DataRow In DBAdapter1.Rows
                If NoStr = "" Then
                    NoStr = dtr("colorcode")
                Else
                    NoStr = NoStr + "," + dtr("colorcode")
                End If

            Next
        End If

        If NoStr <> "" Then
            DDuplicateNo.Text = NoStr
            ' uJavaScript.PopMsg(Me, "有重覆依賴的編號!")
            If MsgBox("有重覆依賴的編號是否繼續執行?", MsgBoxStyle.OkCancel + MsgBoxStyle.MsgBoxSetForeground, "檢查") <> MsgBoxResult.Ok Then
                '不執行
                Message = "有重覆依賴的編號或WAVE'S已有相同組合"
                isOK = False

            End If


        End If


        If DYKKColorType.SelectedValue <> "" Or DYKKColorCode.Value <> "" Then
            '檢查色別色號是否重覆
            SQL = "select * "
            SQL = SQL + " from F_COMBISheet "
            SQL = SQL + " where  YKKColorType='" + DYKKColorType.SelectedValue + "'"
            SQL = SQL + " and  YKKColorCode='" + DYKKColorCode.Value + "'"

            Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)
            If DBAdapter2.Rows.Count > 0 Then
                Message = Message + "\n" + "有重覆色別色號"
                isOK = False

            End If

        End If

        '檢查是否有輸入廢止色
        Dim Code(6) As String
        Dim CodeStr(6) As String

        Code(1) = DVFLTape.Text
        '   Code(2) = DVFLChain.Text
        '  Code(3) = DVFRChain.Text
        Code(2) = DVFRTape.Text
        Code(3) = DVFMLTape.Text
        ' Code(6) = DVFMLChain.Text
        'Code(7) = DVFMRChain.Text
        Code(4) = DVFMRTape.Text
        Code(5) = DPFMFLTape.Text
        Code(6) = DPFMFRTape.Text

        CodeStr(1) = DITEMNAME1.Text + "-左布帶輸入的資料為廢止色"
        ' CodeStr(2) = "塑鋼 VS (一般)-左齒輸入的資料為廢止色"
        ' CodeStr(3) = "塑鋼 VS (一般)-右齒輸入的資料為廢止色"
        CodeStr(2) = DITEMNAME1.Text + "-右布帶輸入的資料為廢止色"
        CodeStr(3) = "塑鋼 VS (金屬齒色)-左布帶輸入的資料為廢止色"
        'CodeStr(6) = "塑鋼 VS (金屬齒色)-左齒輸入的資料為廢止色"
        'CodeStr(7) = "塑鋼 VS (金屬齒色)-右齒輸入的資料為廢止色"
        CodeStr(4) = "塑鋼 VS (金屬齒色)-右布帶輸入的資料為廢止色"
        CodeStr(5) = "尼龍PF VS 金屬MF-左布帶輸入的資料為廢止色"
        CodeStr(6) = "尼龍PF VS 金屬MF-右布帶輸入的資料為廢止色"

        Dim i, j, k As Integer

        If DCOMBIItem.SelectedValue = "塑鋼 VS (一般)" Or DCOMBIItem.SelectedValue = "塑鋼VT810 / VT108" Then
            k = 1
            j = 4
        ElseIf DCOMBIItem.SelectedValue = "塑鋼 VS (金屬齒色)" Then
            k = 5
            j = 8
        Else
            k = 9
            j = 10
        End If


        For i = 1 To 6
            SQL = " select * from F_combirepeal"
            SQL = SQL + " where  1=1 and colorcode ='" + Code(i) + "'"
            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
            If DBAdapter3.Rows.Count > 0 Then
                For Each dtr As Data.DataRow In DBAdapter3.Rows
                    Message = Message + "\n" + CodeStr(i)
                    isOK = False

                Next

            End If

        Next



    End Sub


    Sub ChangeVisible()
        If DCOMBIItem.SelectedValue = "塑鋼 VS (一般)" Or DCOMBIItem.SelectedValue = "塑鋼VT810 / VT108" Then
            DITEMNAME1.Text = DCOMBIItem.SelectedValue
            DVFMLTape.Visible = False
            DVFMLChain.Visible = False
            DVFMRChain.Visible = False
            DVFMRTape.Visible = False
            DPFMFLTape.Visible = False
            DPFMFRTape.Visible = False

            DVFLTape.Visible = True
            DVFLChain.Visible = True
            '20151231  可輸入左齒，自動帶出右齒=左齒色號
            If DCOMBIItem.SelectedValue = "塑鋼VT810 / VT108" Then
                DVFRChain.Visible = False
            Else
                DVFRChain.Visible = True
            End If

            DVFRTape.Visible = True

        ElseIf DCOMBIItem.SelectedValue = "塑鋼 VS (金屬齒色)" Then
            DVFLTape.Visible = False
            DVFLChain.Visible = False
            DVFRChain.Visible = False
            DVFRTape.Visible = False
            DPFMFLTape.Visible = False
            DPFMFRTape.Visible = False

            DVFMLTape.Visible = True
            DVFMLChain.Visible = True
            DVFMRChain.Visible = True
            DVFMRTape.Visible = True

        ElseIf DCOMBIItem.SelectedValue = "尼龍 PF / 金屬MF" Then
            DVFMLTape.Visible = False
            DVFMLChain.Visible = False
            DVFMRChain.Visible = False
            DVFMRTape.Visible = False
            DVFLTape.Visible = False
            DVFLChain.Visible = False
            DVFRChain.Visible = False
            DVFRTape.Visible = False

            DPFMFLTape.Visible = True
            DPFMFRTape.Visible = True
        Else
            DVFMLTape.Visible = False
            DVFMLChain.Visible = False
            DVFMRChain.Visible = False
            DVFMRTape.Visible = False
            DPFMFLTape.Visible = False
            DPFMFRTape.Visible = False

            DVFLTape.Visible = False
            DVFLChain.Visible = False
            DVFRChain.Visible = False
            DVFRTape.Visible = False

        End If
    End Sub

    Protected Sub DCOMBIItem_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DCOMBIItem.SelectedIndexChanged
        ChangeVisible()
    End Sub

    Sub copyNO()
        Dim sql, NO1 As String
        NO1 = DNO1.Text


        If DNO1.Text <> "" Then
            Sql = "Select * From  F_Combisheet "
            Sql = Sql & " Where no = '" + NO1 + "'"
            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(Sql)
            DCustomer.Text = DBAdapter1.Rows(0).Item("Customer")
            DCustomerCode.Text = DBAdapter1.Rows(0).Item("CustomerCode")                         'No
            DBuyer.Text = DBAdapter1.Rows(0).Item("Buyer")                         'No
            DBuyerCode.Text = DBAdapter1.Rows(0).Item("BuyerCode")

            ' DDuplicateNo.Text = DBAdapter1.Rows(0).Item("DuplicateNo")
            ' SetFieldData("YKKColorType", DBAdapter1.Rows(0).Item("YKKColorType"))
            ' DYKKColorCode.Value = DBAdapter1.Rows(0).Item("YKKColorCode")
            SetFieldData("COMBIItem", DBAdapter1.Rows(0).Item("COMBIITem"))
            'VF 
            DVFLTape.Text = DBAdapter1.Rows(0).Item("VFLTape")
            DVFLChain.Text = DBAdapter1.Rows(0).Item("VFLChain")
            DVFRChain.Text = DBAdapter1.Rows(0).Item("VFRChain")
            DVFRTape.Text = DBAdapter1.Rows(0).Item("VFRTape")

            'VF & MF
            DVFMLTape.Text = DBAdapter1.Rows(0).Item("VFMLTape")
            DVFMLChain.Text = DBAdapter1.Rows(0).Item("VFMLChain")
            DVFMRChain.Text = DBAdapter1.Rows(0).Item("VFMRChain")
            DVFMRTape.Text = DBAdapter1.Rows(0).Item("VFMLTape")

            'VF & PF
            DPFMFLTape.Text = DBAdapter1.Rows(0).Item("PFMFLTape")
            DPFMFRTape.Text = DBAdapter1.Rows(0).Item("PFMFLTape")


        End If

    End Sub
  
    Protected Sub DNO1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DNO1.TextChanged
        copyNO()
        ChangeVisible()
    End Sub

    Protected Sub DVFRTape_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DVFRTape.TextChanged
        OK()
    End Sub
    Private Sub dataTableToCsv(ByVal No As String)

        Dim SavePath As String = "D:\db\EXCEL\"
        Dim FileName As String = No + "-" + DateTime.Now.ToString("yyyy-MM-dd HHmmss") + ".csv"
        Dim FilePath As String = SavePath + FileName

        Dim SQL As String

        SQL = " select  ''''+pdpcw3 as pdpcw3,''''+ctbnw3 as ctbnw3,paccw3,st1cw3,st6cw3,scsvw3,''''+cclcw3 as cclcw3 from  R_NCFA300WK  where no = '" + No + "' "

        Dim MyDataTable As DataTable = uDataBase.GetDataTable(SQL)

        Dim sw As New System.IO.StreamWriter(FilePath, False, System.Text.Encoding.Default)
        '寫入欄位名稱()
        If MyDataTable.Columns.Count > 0 Then
            sw.Write(MyDataTable.Columns.Item(0).ColumnName.ToString)
        End If
        For i As Integer = 1 To MyDataTable.Columns.Count - 1
            sw.Write("," + MyDataTable.Columns.Item(i).ColumnName.ToString)
        Next
        sw.Write(sw.NewLine)

        Dim astr As String

        '寫入各欄位資料
        For i As Integer = 0 To MyDataTable.Rows.Count - 1
            For j As Integer = 0 To MyDataTable.Columns.Count - 1
                If j = 0 Then
                    astr = String.Format("{0:00}", MyDataTable.Rows(i)(j).ToString)
                    sw.Write(astr)
                    ' sw.Write(MyDataTable.Rows(i)(j).ToString)

                Else
                    sw.Write("," + MyDataTable.Rows(i)(j).ToString)
                End If
            Next
            sw.Write(sw.NewLine)
        Next

        sw.Close()


    End Sub

End Class

