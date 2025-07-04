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





Partial Class QASheet_01
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
    Dim oWaves As New WAVES.CommonService
    Dim OpenDir1 As String = ""
    Dim OpenDir2 As String = ""
    Dim AttachFileName As String
    Dim sourceDir, backupDir As String
    Dim CheckItemStr, CheckTypeStr, RPReportStr, SampleStr, MaterialStr, LocationStr As String
    Dim chktempFolderPath As String
    Dim ONo, OSeqno, OPuller, OColor As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置
        SetControlPosition()    ' 設定控制項位置
        GetData()



        If Not Me.IsPostBack Then   '不是PostBack
            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()     '表單功能按鈕顯示


            CheckItem1()

            NewAttachFilePath()
            If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]

                ShowFormData()      '顯示表單資料
                UpdateTranFile()    '更新交易資料
                TopPosition()
                SetControlPosition()    ' 設定控制項位置

                NewAttachFilePath1()
                NewAttachFilePath2()
            End If
            SetPopupFunction()      '設定彈出視窗事件


        Else
          

            ShowSheetFunction()     '表單功能按鈕顯示
            ' fpObj.GetDeafultNextGate(DNextGate, DDivision.Text,Request.QueryString("pUserID")) ' 設定預設的簽核者
            ShowSheetField("Post")  '表單欄位顯示及欄位輸入檢查

            ShowMessage()           '上傳資料檢查及顯示訊息

            '上傳資料檢查及顯示訊息

            '   ShowFormData()      '顯示表單資料

            If wFormSno > 0 And (wStep > 2 And wStep <> 500) Then    '判斷是否[簽核]
                CheckItem1()
                ShowFormData()      '顯示表單資料
                UpdateTranFile()    '更新交易資料
                TopPosition()
                SetControlPosition()    ' 設定控制項位置

                NewAttachFilePath1()
                NewAttachFilePath2()
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
            Top = 950
            Dim SQL As String = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(SQL)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0)("BEndTime").ToString < NowDateTime Then
                    If DDelay.Visible = True Then
                        Top = 950
                    End If
                End If
            End If
        Else
            Top = 950
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

            DDescSheet.Style("top") = Top & "px"
            DDecideDesc.Style("top") = Top + 6 & "px"
            Top = Top + 50
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
        '
        Response.Cookies("PGM").Value = "QASheet_01.aspx"
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
                Top = 950
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
            Top = 950
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

   

        BSAVE.Style("top") = Top + 70 & "px"
        BNG1.Style("top") = Top + 70 & "px"
        BNG2.Style("top") = Top + 70 & "px"
        BOK.Style("top") = Top + 70 & "px"
        DHistoryLabel.Style("top") = Top + 100 & "px"
        GridView2.Style("top") = Top + 100 + 16 & "px"
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



        Dim InputData1 As String
        Dim InputData2 As String


        Dim SQL As String

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



        sql = " Select a.* From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a ,m_users b"
        sql = sql & "  where a.empname =b.username"
        sql = sql & " and UserID = '" & wApplyID & "'"
        sql = sql & "   And Active = '1' "
        Dim DBUser As DataTable = uDataBase.GetDataTable(sql)

        If DBUser.Rows(0).Item("EmpName") = "陳怡君" Then
            sql = " Select a.* From m_wfsemp a,m_users b"
            sql = sql & "  where a.empname =b.username and depid = hrwdivid "
            sql = sql & " and UserID = '" & wApplyID & "'"
            sql = sql & "   And Active = '1' "
            Dim DBUser1 As DataTable = uDataBase.GetDataTable(sql)
            DDepName.Text = DBUser1.Rows(0).Item("DepName")
            DName.Text = DBUser1.Rows(0).Item("EmpName")

        Else
            DDepName.Text = DBUser.Rows(0).Item("DepName")
            DName.Text = DBUser.Rows(0).Item("EmpName")
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
        If pPost = "New" Then DDate.Text = Now.ToString("yyyy/MM/dd") '現在日時


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





        'Select Case FindFieldInf("Division")
        Select Case FindFieldInf("DepName")
            Case 0  '顯示
                DDepName.BackColor = Color.LightGray
                DDepName.Visible = True
                DDepName.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DDepName.Visible = True
                DDepName.BackColor = Color.GreenYellow
                DDepName.ReadOnly = False
                ShowRequiredFieldValidator("DDepNameRqd", "DDepName", "異常：需輸入部門")
            Case 2  '修改
                DDepName.Visible = True
                DDepName.BackColor = Color.Yellow
                DDepName.ReadOnly = False
            Case Else   '隱藏
                DDepName.Visible = False
        End Select


        '目的
        Select Case FindFieldInf("Subject")
            Case 0  '顯示
                DSubject.BackColor = Color.LightGray
                DSubject.Visible = True

            Case 1  '修改+檢查
                DSubject.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSubjectRqd", "DSubject", "異常：需輸入目的")
                DSubject.Visible = True
            Case 2  '修改
                DSubject.BackColor = Color.Yellow
                DSubject.Visible = True
            Case Else   '隱藏
                DSubject.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Subject", "ZZZZZZ")




        '品名
        Select Case FindFieldInf("ItemName")
            Case 0  '顯示
                DItemName.BackColor = Color.LightGray
                DItemName.Visible = True

            Case 1  '修改+檢查
                DItemName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DItemNameRqd", "DItemName", "異常：需輸入品名")
                DItemName.Visible = True
            Case 2  '修改
                DItemName.BackColor = Color.Yellow
                DItemName.Visible = True
            Case Else   '隱藏
                DItemName.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ItemName", "ZZZZZZ")


        '背景
        Select Case FindFieldInf("BackGround")
            Case 0  '顯示
                DBackGround.BackColor = Color.LightGray
                DBackGround.Visible = True

            Case 1  '修改+檢查
                DBackGround.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBackGroundRqd", "DBackGround", "異常：需輸入實務確認者")
                DBackGround.Visible = True
            Case 2  '修改
                DBackGround.BackColor = Color.Yellow
                DBackGround.Visible = True
            Case Else   '隱藏
                DBackGround.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BackGround", "ZZZZZZ")

        'ADate
        Select Case FindFieldInf("ADate")
            Case 0  '顯示
                DADate.BackColor = Color.LightGray
                DADate.Visible = True
                DADate.Attributes.Add("readonly", "true")
                BADate.Visible = False
            Case 1  '修改+檢查
                DADate.Visible = True
                DADate.BackColor = Color.GreenYellow
                DADate.ReadOnly = False
                BADate.Visible = True
                ShowRequiredFieldValidator("DADateRqd", "DADate", "異常：需輸入分析依賴日")
            Case 2  '修改
                DADate.Visible = True
                DADate.BackColor = Color.Yellow
                DADate.ReadOnly = False
                BADate.Visible = True
            Case Else   '隱藏
                DADate.Visible = False
                BADate.Visible = False
        End Select
        If pPost = "New" Then DADate.Text = Now.ToString("yyyy/MM/dd") '現在日時




        'Customer
        Select Case FindFieldInf("Customer")
            Case 0  '顯示
                DCustomer.BackColor = Color.LightGray
                DCustomer.ReadOnly = True
                DCustomer.Visible = True
                DCustomerCode.BackColor = Color.LightGray
                DCustomerCode.ReadOnly = True
                DCustomerCode.Visible = True
                BCustomer.Visible = False

            Case 1  '修改+檢查
                DCustomer.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerRqd", "DCustomer", "異常：需輸入客戶名稱")
                DCustomer.Visible = True
                DCustomer.ReadOnly = True
                DCustomerCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerCodeRqd", "DCustomerCode", "異常：需輸入客戶名稱")
                DCustomerCode.Visible = True
                DCustomerCode.ReadOnly = True
                BCustomer.Visible = True

            Case 2  '修改
                DCustomer.BackColor = Color.Yellow
                DCustomer.Visible = True
                DCustomer.ReadOnly = True
                DCustomerCode.BackColor = Color.Yellow
                DCustomerCode.Visible = True
                DCustomerCode.ReadOnly = True
                BCustomer.Visible = True

            Case Else   '隱藏
                DCustomer.Visible = False
                DCustomerCode.Visible = False
                BCustomer.Visible = False

        End Select
        ' If pPost = "New" Then SetFieldData("CUSTOMER", "ZZZZZZ")


        'BuyerCode
        Select Case FindFieldInf("Buyer")
            Case 0  '顯示
                DBuyerCode.BackColor = Color.LightGray
                DBuyerCode.ReadOnly = True
                DBuyerCode.Visible = True

                DBuyer.BackColor = Color.LightGray
                DBuyer.ReadOnly = True
                DBuyer.Visible = True

                BBuyer.Visible = False

            Case 1  '修改+檢查
                DBuyerCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerCodeRqd", "DBuyerCode", "異常：需輸入Buyer")
                DBuyerCode.Visible = True
                DBuyerCode.ReadOnly = True
                DBuyer.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerRqd", "DBuyer", "異常：需輸入Buyer")
                DBuyer.Visible = True
                DBuyer.ReadOnly = True

                BBuyer.Visible = True
            Case 2  '修改
                DBuyerCode.BackColor = Color.Yellow
                DBuyerCode.Visible = True
                BBuyer.Visible = True
                DBuyer.BackColor = Color.Yellow
                DBuyer.Visible = True

                BBuyer.Visible = True
            Case Else   '隱藏
                DBuyerCode.Visible = False
                BBuyer.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BuyerCode", "ZZZZZZ")



        'FinishDate
        Select Case FindFieldInf("FinishDate")
            Case 0  '顯示
                DFinishDate.BackColor = Color.LightGray
                DFinishDate.Visible = True
                DFinishDate.Attributes.Add("readonly", "true")
                BFinishDate.Visible = False
            Case 1  '修改+檢查
                DFinishDate.Visible = True
                DFinishDate.BackColor = Color.GreenYellow
                DFinishDate.ReadOnly = False
                BFinishDate.Visible = True
                ShowRequiredFieldValidator("DFinishDateRqd", "DFinishDate", "異常：需輸入希望完了日")
            Case 2  '修改
                DFinishDate.Visible = True
                DFinishDate.BackColor = Color.Yellow
                DFinishDate.ReadOnly = False
                BFinishDate.Visible = True
            Case Else   '隱藏
                DFinishDate.Visible = False
                BFinishDate.Visible = False
        End Select
        If pPost = "New" Then DFinishDate.Text = Now.AddDays(1).ToString("yyyy/MM/dd") '現在日時


        'CheckItem
        Select Case FindFieldInf("CheckItem")
            Case 0  '顯示
                'DCheckItem.BackColor = Color.LightGray
                DCheckItem.Visible = False
                DCheckItem.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCheckItem.Visible = True
                '   DCheckItem.BackColor = Color.GreenYellow

                ' ShowRequiredFieldValidator("DCheckItemRqd", "DCheckItem", "異常：需輸入營業依賴")
            Case 2  '修改
                DCheckItem.Visible = True
                '  DCheckItem.BackColor = Color.Yellow

            Case Else   '隱藏
                DCheckItem.Visible = False
        End Select



        'CheckType
        Select Case FindFieldInf("CheckType")
            Case 0  '顯示
                ' DCheckType.BackColor = Color.LightGray
                DCheckType.Visible = False
                DCheckType.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCheckType.Visible = True
                '  DCheckType.BackColor = Color.GreenYellow

                ' ShowRequiredFieldValidator("DCheckTypeRqd", "DCheckType", "異常：需輸入營業依賴")
            Case 2  '修改
                DCheckType.Visible = True
                ' DCheckType.BackColor = Color.Yellow

            Case Else   '隱藏
                DCheckType.Visible = False
        End Select




        'RPReport
        Select Case FindFieldInf("RPReport")
            Case 0  '顯示
                '  DRPReport.BackColor = Color.LightGray
                DRPReport.Visible = False
                DRPReport.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRPReport.Visible = True
                ' DRPReport.BackColor = Color.GreenYellow

                '  ShowRequiredFieldValidator("DRPReportRqd", "DRPReport", "異常：需輸入回覆報告書")
            Case 2  '修改
                DRPReport.Visible = True
                '  DRPReport.BackColor = Color.Yellow

            Case Else   '隱藏
                DRPReport.Visible = False
        End Select

        'Sample
        Select Case FindFieldInf("Sample")
            Case 0  '顯示
                '     DSample.BackColor = Color.LightGray
                DSample.Visible = False
                DSample.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSample.Visible = True
                '  DSample.BackColor = Color.GreenYellow

                '  ShowRequiredFieldValidator("DSampleRqd", "DSample", "異常：需輸入樣品")
            Case 2  '修改
                DSample.Visible = True
                '  DSample.BackColor = Color.Yellow

            Case Else   '隱藏
                DSample.Visible = False
        End Select

        'Material
        Select Case FindFieldInf("Material")
            Case 0  '顯示
                '     DMaterial.BackColor = Color.LightGray
                DMaterial.Visible = False
                DMaterial.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DMaterial.Visible = True
                '  DMaterial.BackColor = Color.GreenYellow

                '  ShowRequiredFieldValidator("DMaterialRqd", "DMaterial", "異常：需輸入殘試料")
            Case 2  '修改
                DMaterial.Visible = True
                '   DMaterial.BackColor = Color.Yellow

            Case Else   '隱藏
                DMaterial.Visible = False
        End Select

        'Location
        Select Case FindFieldInf("Location")
            Case 0  '顯示
                '     DLocation.BackColor = Color.LightGray
                DLocation.Visible = False
                DLocation.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DLocation.Visible = True
                '  DLocation.BackColor = Color.GreenYellow

                '  ShowRequiredFieldValidator("DLocationRqd", "DLocation", "異常：需輸入分析資料配布先")
            Case 2  '修改
                DLocation.Visible = True
                '  DLocation.BackColor = Color.Yellow

            Case Else   '隱藏
                DLocation.Visible = False
        End Select







        '
        Select Case FindFieldInf("AttachFile1")
            Case 0  '顯示            

                DAttachFile1.Visible = True

            Case 1  '修改+檢查
                DAttachFile1.Visible = True

            Case 2  '修改
                DAttachFile1.Visible = True



            Case Else   '隱藏
                DAttachFile1.Visible = False


        End Select

        Select Case FindFieldInf("AttachFile2")
            Case 0  '顯示            

                DAttachFile2.Visible = True

            Case 1  '修改+檢查
                DAttachFile2.Visible = True

            Case 2  '修改
                DAttachFile2.Visible = True



            Case Else   '隱藏
                DAttachFile2.Visible = False


        End Select


        'QCNo
        Select Case FindFieldInf("QCNo")
            Case 0  '顯示
                DQCNo.BackColor = Color.LightGray
                DQCNo.Visible = True
                DQCNo.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DQCNo.Visible = True
                DQCNo.BackColor = Color.GreenYellow

                ShowRequiredFieldValidator("DQCNoRqd", "DQCNo", "異常：需輸入品管受理NO")
            Case 2  '修改
                DQCNo.Visible = True
                DQCNo.BackColor = Color.Yellow

            Case Else   '隱藏
                DQCNo.Visible = False
        End Select



        'QCDate
        Select Case FindFieldInf("QCDate")
            Case 0  '顯示
                DQCDate.BackColor = Color.LightGray
                DQCDate.Visible = True
                DQCDate.Attributes.Add("readonly", "true")
                BQCDate.Visible = False
            Case 1  '修改+檢查
                DQCDate.Visible = True
                DQCDate.BackColor = Color.GreenYellow
                BQCDate.Visible = True
                ShowRequiredFieldValidator("DQCDateRqd", "DQCDate", "異常：需輸入分析完了日")
            Case 2  '修改
                DQCDate.Visible = True
                DQCDate.BackColor = Color.Yellow
                BQCDate.Visible = True

            Case Else   '隱藏
                DQCDate.Visible = False
                BQCDate.Visible = False
        End Select
        'If pPost = "New" Then DQCDate.Text = Now.ToString("yyyy/MM/dd") '現在日時





        'ADD
        Select Case FindFieldInf("Add")
            Case 0  '顯示

                BAdd.Visible = True
            Case 1  '修改+檢查
                BAdd.Visible = True


            Case 2  '修改
                BAdd.Visible = True


            Case Else   '隱藏
                BAdd.Visible = False
        End Select
















    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()

 
        Dim SQL As String


        SQL = "Select * From F_QASheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'" '
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then

            DNo.Text = dtData.Rows(0).Item("No")                         'No
            DDate.Text = dtData.Rows(0).Item("Date")
            DDepName.Text = dtData.Rows(0).Item("DepName")
            DName.Text = dtData.Rows(0).Item("Name")
            DADate.Text = dtData.Rows(0).Item("ADate")
            DSubject.Text = dtData.Rows(0).Item("Subject")
            DItemName.Text = dtData.Rows(0).Item("ItemName")

            SetFieldData("BackGround", dtData.Rows(0).Item("BackGround"))


            DBuyerCode.Text = dtData.Rows(0).Item("BuyerCode")
            DBuyer.Text = dtData.Rows(0).Item("Buyer")


            DCustomerCode.Text = dtData.Rows(0).Item("CustomerCode")
            DCustomer.Text = dtData.Rows(0).Item("Customer")

            If dtData.Rows(0).Item("CheckItem") <> "" Then
                GetCheckItem(dtData.Rows(0).Item("CheckItem"))
            End If


            If dtData.Rows(0).Item("CheckType") <> "" Then
                GetCheckType(dtData.Rows(0).Item("CheckType"))
            End If
            

            If dtData.Rows(0).Item("RPReport") <> "" Then
                GetRPReport(dtData.Rows(0).Item("RPReport"))
            End If

            If dtData.Rows(0).Item("Sample") <> "" Then
                GetSample(dtData.Rows(0).Item("Sample"))
            End If

            If dtData.Rows(0).Item("Material") <> "" Then
                GetMaterial(dtData.Rows(0).Item("Material"))
            End If


            If dtData.Rows(0).Item("Location") <> "" Then
                GetLocation(dtData.Rows(0).Item("Location"))
            End If


        End If

        DFinishDate.Text = dtData.Rows(0).Item("FinishDate")

        DQCNo.Text = dtData.Rows(0).Item("QCNO")

        If dtData.Rows(0).Item("QCDate") = "1900-01-01 00:00:00.000" Then
            DQCDate.Text = ""
        Else

            DQCDate.Text = dtData.Rows(0).Item("QCDate")
        End If





        SQL = " select Unique_id,no,SeqNo,Type,Supplier,Code,Size,Family,Body,Puller,COLOR,Finish,ApproveNo,OldToNewNo,OldToNew,Remark,Element,Result1,Result2,QCRemark"
        SQL = SQL + " from  F_QASheetDT "
        SQL = SQL + "  where  No='" + DNo.Text + "' "
        SQL = SQL + " order by SeqNo "


        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()
        DCount.Text = "共" + CStr(GridView1.Rows.Count) + "筆"
        If GridView1.Rows.Count > 0 Then

            DHaveData.Text = "1"
        Else
            DHaveData.Text = "0"
        End If



        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '8002'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
        Dim OpenDir As String = ""
        If DBAdapter3.Rows.Count > 0 Then
            OpenDir = "file://" + DBAdapter3.Rows(0).Item("Data") + DNo.Text + "/核可卡"
        End If
        Dim OpenDir2 As String = ""
        OpenDir2 = "file://" + DBAdapter3.Rows(0).Item("Data") + DNo.Text + "/核可卡"
        DAttachFile1.Attributes.Add("onclick", "window.open('" + OpenDir + "','_blank');return false;")




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
        SQL = SQL + "Order by  CreateTime Desc, Step Desc, SeqNo Desc "
        GridView2.DataSource = uDataBase.GetDataTable(SQL)
        GridView2.DataBind()

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
        '擔當者及部門 



        'BACKGROUND
        If pFieldName = "BackGround" Then
            DBackGround.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBackGround.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select   substring(data,1,CHARINDEX('-', data)-1)appuser,substring(data,CHARINDEX('-', data)+1,len(data)-1)data  from M_referp"
                sql = sql & " where  cat = '8002'"
                sql = sql & " and dkey = 'CheckUser'"
                sql = sql & " and substring(data,1,CHARINDEX('-', data)-1)='" & DName.Text & "'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DBackGround.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBackGround.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
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
        rqdVal.Style.Add("Top", Top - 20 & "px")
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
   


        If wFormSno <> 0 Then
            DTNo.Text = DNo.Text
        Else
            DTNo.Text = Now.ToString("yyyyMMddHHmmss") '虛擬單號
        End If


        BCustomer.Attributes.Add("onClick", "GetCustomer()") '找客戶
        BBuyer.Attributes.Add("onClick", "GetBuyer()") '找buyer

        BADate.Attributes("onclick") = "calendarPicker('Form1.DADate');"
        BFinishDate.Attributes("onclick") = "calendarPicker('Form1.DFinishDate');"
        BQCDate.Attributes("onclick") = "calendarPicker('Form1.DQCDate');"

        ' BAdd.Attributes("onclick") = "window.open('AddItemListBC.aspx?pUserID=" & Request.QueryString("pUserID") & "&pDTNo=" & DTNo.Text & "','','height=600,width=950,menubar=no,location=no');"
        BAdd.Attributes("onclick") = "window.open('AddItemList.aspx?pwFormNo=008002&pUserID=" & Request.QueryString("pUserID") & "&pDTNo=" & DTNo.Text & "&pwStep=" & wStep & "','newwindow','height=600,width=1500,menubar=no,location=no');"

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

                If (wStep = 1 Or wStep = 500) And pFun = "OK" Then

                    wAllocateID = GetRelated(DBackGround.SelectedValue)  '實務確認者
                End If





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
            Dim URL As String = ""


            URL = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                           "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID

            Response.Redirect(URL)


        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     新增表單資料
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)

        '執行複選項字串相接
        CheckResult()

    

        Dim SQL As String
        SQL = " Insert into F_QASheet (Sts, CompletedTime, FormNo, FormSno,"
        SQL = SQL + " NO,DATE,DepName,Name,ADate, "  '11
        SQL = SQL + " Subject,ItemName,BackGround,CheckItem,CheckType,"
        SQL = SQL + " CustomerCode,Customer,BuyerCode,Buyer,FinishDate,"
        SQL = SQL + " RPReport,Sample,Material,Location,"
        SQL = SQL + " QCNO,QCDate,"
        SQL = SQL + " CreateUser, CreateTime, ModifyUser, ModifyTime) "
        SQL = SQL + " VALUES( "
        SQL = SQL + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        SQL = SQL + " '" + NowDateTime + "', "        '結案日
        SQL = SQL + " '008002', "                     '表單代號
        SQL = SQL + " '" + CStr(NewFormSno) + "', "   '表單流水號
        SQL = SQL + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO  1
        SQL = SQL + " N'" + YKK.ReplaceString(DDate.Text) + "', "
        SQL = SQL + " N'" + YKK.ReplaceString(DDepName.Text) + "', "
        SQL = SQL + " N'" + YKK.ReplaceString(DName.Text) + "', "
        SQL = SQL + " N'" + YKK.ReplaceString(DADate.Text) + "', "
        SQL = SQL + " N'" + YKK.ReplaceString(DSubject.Text) + "', "
        SQL = SQL + " N'" + YKK.ReplaceString(DItemName.Text) + "', "
        SQL = SQL + " N'" + YKK.ReplaceString(DBackGround.SelectedValue) + "', "
        SQL = SQL + " '" + CheckItemStr + "', "
        SQL = SQL + " '" + CheckTypeStr + "', "
        SQL = SQL + " N'" + YKK.ReplaceString(DCustomerCode.Text) + "', "
        SQL = SQL + " N'" + YKK.ReplaceString(DCustomer.Text) + "', "
        SQL = SQL + " N'" + YKK.ReplaceString(DBuyerCode.Text) + "', "
        SQL = SQL + " N'" + YKK.ReplaceString(DBuyer.Text) + "', "
        SQL = SQL + " N'" + YKK.ReplaceString(DFinishDate.Text) + "', "
        SQL = SQL + " '" + RPReportStr + "', "
        SQL = SQL + " '" + SampleStr + "', "
        SQL = SQL + " '" + MaterialStr + "', "
        SQL = SQL + " '" + LocationStr + "', "
        SQL = SQL + " N'" + YKK.ReplaceString(DQCNo.Text) + "', "
        SQL = SQL + " N'" + YKK.ReplaceString(DQCDate.Text) + "', "
        SQL = SQL + "'" & Request.QueryString("pUserID") & "' ,"
        SQL = SQL + "'" & NowDateTime & "' ,"
        SQL = SQL + "'" & Request.QueryString("pUserID") & "' ,"
        SQL = SQL + "'" & NowDateTime & "' )"
        uDataBase.ExecuteNonQuery(SQL)


        If wStep = 1 Then '先刪除同單號細項
            SQL = " delete from  F_QASheetDT"
            SQL = SQL + " where no ='" + YKK.ReplaceString(DNo.Text) + "'"
            uDataBase.ExecuteNonQuery(SQL)
        End If

        '細項新增


        SQL = " Insert into F_QASheetDT (No,SeqNo,Type,Supplier,Code,Size,Family,Body,Puller,COLOR,Finish,ApproveNo,OldToNewNo,OldToNew,Remark,Element,Result1,Result2,QCRemark,Createuser,CreateTime,ModifyUser,ModifyTime) "
        SQL = SQL + " select N'" + YKK.ReplaceString(DNo.Text) + "',SeqNo,Type,Supplier,Code,Size,Family,Body,Puller,COLOR,Finish,ApproveNo,OldToNewNo,OldToNew,Remark,Element,case when type in ('NEW','REF') THEN 'OK' ELSE Result1 END AS RESULT1,Result2,QCRemark,"
        SQL = SQL + " '" + Request.QueryString("pUserID") + "',"
        SQL = SQL + " getdate(),'" + Request.QueryString("pUserID") + "',getdate() "
        SQL = SQL + " from  F_QASheetDTTemp "
        SQL = SQL + "  where  No='" + DTNo.Text + "' "
        SQL = SQL + " order by SeqNo "
        uDataBase.ExecuteNonQuery(SQL)




        If wStep = 1 Then '刪掉暫存單號
            SQL = " delete from F_QASheetDT where No='" + DTNo.Text + "'"
            uDataBase.ExecuteNonQuery(SQL)

            SQL = " delete from F_QASheetDTTemp where No='" + DTNo.Text + "'"
            uDataBase.ExecuteNonQuery(SQL)

        End If

        Dim i As Integer

        '回寫舊廢除，先檢查有核可號或舊廢除嗎?
        SQL = " select OldtonewNo  as Data from F_QASheetDT "
        SQL = SQL + " where   OldToNewNo <> '' "
        SQL = SQL + " and  No ='" + YKK.ReplaceString(DNo.Text) + "'"
        Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter3.Rows.Count > 0 Then
 
            For i = 0 To DBAdapter3.Rows.Count - 1
         
                GetOLDNo(DBAdapter3.Rows(i).Item("Data"))
                '更新舊換新  廢除FLAG   0:無廢除  /  1:廢除
                SQL = " update F_QASheetdt "
                SQL = SQL + " set oldcancel = 1,"
                SQL = SQL + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
                SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
                SQL = SQL + " where No = '" + ONo + "' and seqno ='" + OSeqno + "' and puller = '" + OPuller + "' and color = '" + OColor + "'"
                uDataBase.ExecuteNonQuery(SQL)

                SQL = " update  M_SPDIRWEDX "
                SQL = SQL + " set formsno = 1,"
                SQL = SQL + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
                SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
                SQL = SQL + " where STS=1 AND CAT='QC'  and Acceptedno = '" + ONo + "-" + OSeqno + "' and puller = '" + OPuller + "' and color = '" + OColor + "'"
                uDataBase.ExecuteNonQuery(SQL)

            Next


        End If



        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '8002'"
        SQL = SQL + " and dkey ='AttachfilePath' "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            sourceDir = "\\" + DBAdapter1.Rows(0).Item("Data1") + D3.Text   '來源
        End If

        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '8002'"
        SQL = SQL + " and dkey ='AttachfilePath1' "
        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter2.Rows.Count > 0 Then
            backupDir = "\\" + DBAdapter2.Rows(0).Item("Data1") + DNo.Text      '目的  
        End If

        '複製前先確認是否有原本的資料夾再刪除
        If Directory.Exists(backupDir) Then
            Directory.Delete(backupDir, True)
        End If
        '從暫存複製到正式
        CopyDir(sourceDir, backupDir)
        '只複製*.jpg
        CopyJPG(backupDir + "\核可卡", "\\10.245.1.6\program$\CutImage\In")

        NewAttachFilePath()




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



    '複製*.JPG
    Public Shared Sub CopyJPG(ByVal sourceDir As String, ByVal backupDir As String)

        'Dim sourceDir As String = "\\10.245.1.6\wfs$\ISOSQC\008002\2407040798\核可卡"
        'Dim backupDir As String = "\\10.245.1.6\program$\CutImage\In"

        Try
            Dim picList() As String = Directory.GetFiles(sourceDir, "*.jpg")


            ' Copy picture files.
            Dim f As String
            For Each f In picList
                ' Remove path from the file name.
                Dim fName As String = f.Substring(sourceDir.Length + 1)

                ' Use the Path.Combine method to safely append the file name to the path.
                ' Will overwrite if the destination file already exists.
                File.Copy(Path.Combine(sourceDir, fName), Path.Combine(backupDir, fName), True)
            Next


            '' Delete source files that were copied.

            'For Each f In txtList
            '    File.Delete(f)
            'Next
            'Dim f As String
            'For Each f In picList
            '    File.Delete(f)
            'Next
        Catch dirNotFound As DirectoryNotFoundException
            Console.WriteLine("無法複製!")
        End Try
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)

        '執行複選項字串相接
        CheckResult()



        Dim sql As String

        sql = " Update F_QASheet"
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
        sql = sql + " DepName = N'" & YKK.ReplaceString(Trim(DDepName.Text)) & "',"
        sql = sql + " Name = N'" & YKK.ReplaceString(Trim(DName.Text)) & "',"
        sql = sql + " ADate = N'" & YKK.ReplaceString(Trim(DADate.Text)) & "',"
        sql = sql + " Subject = N'" & YKK.ReplaceString(Trim(DSubject.Text)) & "',"
        sql = sql + " ItemName = N'" & YKK.ReplaceString(DItemName.Text) & "',"
        sql = sql + " BackGround = N'" & YKK.ReplaceString(DBackGround.SelectedValue) & "',"
        sql = sql + " CustomerCode = N'" & YKK.ReplaceString(DCustomerCode.Text) & "',"
        sql = sql + " Customer = N'" & YKK.ReplaceString(DCustomer.Text) & "',"
        sql = sql + " BuyerCode = N'" & YKK.ReplaceString(DBuyerCode.Text) & "',"
        sql = sql + " Buyer = N'" & YKK.ReplaceString(DBuyer.Text) & "',"
        sql = sql + " FinishDate = N'" & YKK.ReplaceString(DFinishDate.Text) & "',"
        sql = sql + " checkItem ='" & CheckItemStr + "',"
        sql = sql + " checkType ='" & CheckTypeStr + "',"
        sql = sql + " RPReport ='" & RPReportStr + "',"
        sql = sql + " Sample ='" & SampleStr + "',"
        sql = sql + " Material ='" & MaterialStr + "',"
        sql = sql + " Location ='" & LocationStr + "',"
        sql = sql + " QCNO = N'" & YKK.ReplaceString(DQCNo.Text) & "',"
        sql = sql + " QCDate = N'" & YKK.ReplaceString(DQCDate.Text) & "',"
        sql &= " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        sql &= " ModifyTime = '" & NowDateTime & "' "
        sql &= " Where FormNo  =  '" & wFormNo & "'"
        sql &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        uDataBase.ExecuteNonQuery(sql)
        Dim i As Integer
        If wStep = 500 And pFun = "NG1" Then  '如果取消就將註記取消
            '回寫舊廢除，先檢查有核可號或舊廢除嗎?
            sql = " select  OldtonewNo  as Data from F_QASheetDT "
            sql = sql + " where   OldToNewNo <> ''  "
            sql = sql + " and  No ='" + YKK.ReplaceString(DNo.Text) + "'"
            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(sql)
            If DBAdapter3.Rows.Count > 0 Then

                For i = 0 To DBAdapter3.Rows.Count - 1

                    GetOLDNo(DBAdapter3.Rows(i).Item("Data"))
                    '更新舊換新  廢除FLAG   0:無廢除  /  1:廢除
                    sql = " update F_QASheetdt "
                    sql = sql + " set oldcancel = 0,"
                    sql = sql + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " where No = '" + ONo + "' and seqno ='" + OSeqno + "' and puller = '" + OPuller + "' and color = '" + OColor + "'"
                    uDataBase.ExecuteNonQuery(sql)

                    sql = " update  M_SPDIRWEDX "
                    sql = sql + " set formsno = 0,"
                    sql = sql + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " where STS=1 AND CAT='QC'  and Acceptedno = '" + ONo + "-" + OSeqno + "' and puller = '" + OPuller + "' and color = '" + OColor + "'"
                    uDataBase.ExecuteNonQuery(sql)

                Next

            End If
        End If

        '只複製*.jpg
        '  CopyJPG(backupDir + "\核可卡", "\\10.245.1.6\program$\CutImage\In")


        '主檔資料
        sql = " select  data,replace(data,'/','\')data1  from M_referp"
        sql = sql + " where cat = '8002'"
        sql = sql + " and dkey ='AttachfilePath1' "
        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(sql)

        If DBAdapter2.Rows.Count > 0 Then
            backupDir = "\\" + DBAdapter2.Rows(0).Item("Data1") + DNo.Text      '目的  
        End If


  
        '只複製*.jpg
        CopyJPG(backupDir + "\核可卡", "\\10.245.1.6\program$\CutImage\In")






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
        If InputDataOK(0) Then
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
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".ai", ".xls", ".doc", ".ppt", ".xlsx", ".docx", ".pptx"}   '定義允許的檔案格式
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
            If UPFile.PostedFile.ContentLength <= 950 * 1024 Then
                UPFileIsNormal = 0
            Else
                UPFileIsNormal = 9030
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



        '檢查資料夾是否有檔案

        If wStep = 10 And pAction = "1" Then


            If ErrCode = 0 Then
                Dim dirInfo As New System.IO.DirectoryInfo(chktemp.Text)
                Dim FileDir As Integer  '資料夾
                FileDir = dirInfo.GetDirectories("*").Length
                Dim FileCount As Integer '檔案
                FileCount = dirInfo.GetFiles("*.*").Length
                If FileCount > 0 Or FileDir > 0 Then

                Else

                    ErrCode = 9010
                End If
            End If

        End If


        If wStep = 1 Then
            '明細
            SQL = " Select  *  from  F_QASheetDTTemp where 1=1 "
            SQL = SQL + " and NO = '" + DTNo.Text + "'"

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
                ErrCode = 9020
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
                ErrCode = oCommon.CommissionNo("008002", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9050
                End If
            End If
        End If



        If ErrCode <> 0 Then
            If ErrCode = 9010 Then Message = "請開啟資料夾放入核可卡附件!!"
            If ErrCode = 9020 Then Message = "細項沒有資料不能送出!"
            If ErrCode = 9030 Then Message = "上傳檔案Size有誤,請確認!"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9050 Then Message = "委託書No.重覆,請確認委託書No.!"

            uJavaScript.PopMsg(Me, Message)

            If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]

                CheckItem1()
                ShowFormData()      '顯示表單資料
            End If

        Else
            isOK = True
        End If
        '
        Return isOK
    End Function


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**CheckItem1()
    '**      '加入選項 checkboxlist radioboxlist 
    '**
    '*****************************************************************


    Sub CheckItem1()
        '展開工程選項 
        'Engine-A 單獨選取
        Dim i As Integer

        DCheckItem.RepeatColumns = 1
        DCheckItem.RepeatDirection = RepeatDirection.Horizontal
        DCheckItem.Visible = True
        DCheckItem.Items.Clear()

        Dim SQL As String

        SQL = "select * from M_referp "
        SQL = SQL + "where cat = '8002'"
        SQL = SQL + "and dkey ='CheckItem'"
        SQL = SQL + " Order by Unique_ID"

        Dim dt As Data.DataTable = uDataBase.GetDataTable(SQL)
        CheckItemCount.Text = dt.Rows.Count
        For Each dtr As Data.DataRow In dt.Rows

            Dim CheckItem As New ListItem(dtr("Data"), dtr("Data"))
            'New ListItem(dtr("RUserName"), dtr("RUserID"))
            DCheckItem.Items.Add(CheckItem)
        Next

        DCheckItem.Items(2).Selected = True


        'Dim SQL As String

        'SQL = "select * from M_referp "
        'SQL = SQL + "where cat = '4001'"
        'SQL = SQL + "and dkey ='EngineSelect-單獨選取'"
        'SQL = SQL + " Order by Unique_ID"

        'Dim dt As Data.DataTable = uDataBase.GetDataTable(SQL)

        'For Each dtr As Data.DataRow In dt.Rows
        '    Dim EngineA As New ListItem(dtr("Data"), dtr("Data"))
        '    'New ListItem(dtr("RUserName"), dtr("RUserID"))
        '    CheckBoxList1.Items.Add(EngineA)
        'Next


        '營業依賴 2 CheckType


        DCheckType.RepeatColumns = 2
        DCheckType.RepeatDirection = RepeatDirection.Horizontal
        DCheckType.Visible = True
        DCheckType.Items.Clear()




        SQL = "select * from M_referp "
        SQL = SQL + "where cat = '8002'"
        SQL = SQL + "and dkey ='CheckType'"
        SQL = SQL + " Order by Unique_ID"

        Dim dt1 As Data.DataTable = uDataBase.GetDataTable(SQL)

        CheckTypeCount.Text = dt1.Rows.Count

        For Each dtr As Data.DataRow In dt1.Rows
            Dim CheckType As New ListItem(dtr("Data"), dtr("Data"))
            'New ListItem(dtr("RUserName"), dtr("RUserID"))
            DCheckType.Items.Add(CheckType)
        Next


        DCheckType.Items(10).Selected = True





        '回覆報告書 RPReport


        DRPReport.RepeatColumns = 3
        DRPReport.RepeatDirection = RepeatDirection.Horizontal
        DRPReport.Visible = True
        DRPReport.Items.Clear()




        SQL = "select * from M_referp "
        SQL = SQL + "where cat = '8002'"
        SQL = SQL + "and dkey ='RPReport'"
        SQL = SQL + " Order by Unique_ID"

        Dim dt2 As Data.DataTable = uDataBase.GetDataTable(SQL)
        RPReportCount.Text = dt2.Rows.Count
        For Each dtr As Data.DataRow In dt2.Rows
            Dim RPReport As New ListItem(dtr("Data"), dtr("Data"))
            'New ListItem(dtr("RUserName"), dtr("RUserID"))
            DRPReport.Items.Add(RPReport)
        Next


        DRPReport.Items(1).Selected = True




        '樣品 Sample


        DSample.RepeatColumns = 3
        DSample.RepeatDirection = RepeatDirection.Horizontal
        DSample.Visible = True
        DSample.Items.Clear()




        SQL = "select * from M_referp "
        SQL = SQL + "where cat = '8002'"
        SQL = SQL + "and dkey ='Sample'"
        SQL = SQL + " Order by Unique_ID"

        Dim dt3 As Data.DataTable = uDataBase.GetDataTable(SQL)
        SampleCount.Text = dt3.Rows.Count
        For Each dtr As Data.DataRow In dt3.Rows
            Dim Sample As New ListItem(dtr("Data"), dtr("Data"))
            'New ListItem(dtr("RUserName"), dtr("RUserID"))
            DSample.Items.Add(Sample)
        Next


        DSample.Items(1).Selected = True


        '殘試料 Material


        DMaterial.RepeatColumns = 3
        DMaterial.RepeatDirection = RepeatDirection.Horizontal
        DMaterial.Visible = True
        DMaterial.Items.Clear()





        SQL = "select * from M_referp "
        SQL = SQL + "where cat = '8002'"
        SQL = SQL + "and dkey ='Material'"
        SQL = SQL + " Order by Unique_ID"

        Dim dt4 As Data.DataTable = uDataBase.GetDataTable(SQL)
        MaterialCount.Text = dt4.Rows.Count
        For Each dtr As Data.DataRow In dt4.Rows
            Dim Material As New ListItem(dtr("Data"), dtr("Data"))
            'New ListItem(dtr("RUserName"), dtr("RUserID"))
            DMaterial.Items.Add(Material)
        Next


        DMaterial.Items(1).Selected = True




        '配布先 Location


        DLocation.RepeatColumns = 3
        DLocation.RepeatDirection = RepeatDirection.Horizontal
        DLocation.Visible = True
        DLocation.Items.Clear()




        SQL = "select * from M_referp "
        SQL = SQL + "where cat = '8002'"
        SQL = SQL + "and dkey ='Location'"
        SQL = SQL + " Order by Unique_ID"

        Dim dt5 As Data.DataTable = uDataBase.GetDataTable(SQL)
        LocationCount.Text = dt5.Rows.Count
        For Each dtr As Data.DataRow In dt5.Rows
            Dim Location As New ListItem(dtr("Data"), dtr("Data"))
            'New ListItem(dtr("RUserName"), dtr("RUserID"))
            DLocation.Items.Add(Location)
        Next

        DLocation.Items(0).Selected = True





    End Sub




    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**CheckResult()
    '**      '組合選項字串 checkboxlist radioboxlist 
    '**
    '*****************************************************************


    Sub CheckResult()
        Dim i, j As Integer
        Dim CheckStr As String
        CheckStr = ""





        j = 1
        For i = 0 To (DCheckItem.Items.Count - 1)
            If (DCheckItem.Items(i).Selected) Then
                If CheckStr = "" Then
                    CheckStr = Mid(Trim(DCheckItem.Items(i).Text), 1, 2)
                Else
                    CheckStr = CheckStr + "," + Mid(Trim(DCheckItem.Items(i).Text), 1, 2)
                End If



                'Dim DText As TextBox = mDirectCast(Page.FindControl("TextBox1"), TextBox)

                j = j + 1
            End If
            CheckItemStr = CheckStr

        Next

        CheckStr = ""

        j = 1
        For i = 0 To (DCheckType.Items.Count - 1)
            If (DCheckType.Items(i).Selected) Then
                If CheckStr = "" Then
                    CheckStr = Mid(Trim(DCheckType.Items(i).Text), 1, 2)
                Else
                    CheckStr = CheckStr + "," + Mid(Trim(DCheckType.Items(i).Text), 1, 2)
                End If



                'Dim DText As TextBox = mDirectCast(Page.FindControl("TextBox1"), TextBox)

                j = j + 1
            End If
            CheckTypeStr = CheckStr

        Next

        CheckStr = ""
        j = 1
        For i = 0 To (DRPReport.Items.Count - 1)
            If (DRPReport.Items(i).Selected) Then
                If CheckStr = "" Then
                    CheckStr = Mid(Trim(DRPReport.Items(i).Text), 1, 2)
                Else
                    CheckStr = CheckStr + "," + Mid(Trim(DRPReport.Items(i).Text), 1, 2)
                End If



                'Dim DText As TextBox = mDirectCast(Page.FindControl("TextBox1"), TextBox)

                j = j + 1
            End If
            RPReportStr = CheckStr

        Next

        CheckStr = ""

        j = 1
        For i = 0 To (DSample.Items.Count - 1)
            If (DSample.Items(i).Selected) Then
                If CheckStr = "" Then
                    CheckStr = Mid(Trim(DSample.Items(i).Text), 1, 2)
                Else
                    CheckStr = CheckStr + "," + Mid(Trim(DSample.Items(i).Text), 1, 2)
                End If



                'Dim DText As TextBox = mDirectCast(Page.FindControl("TextBox1"), TextBox)

                j = j + 1
            End If
            SampleStr = CheckStr

        Next


        CheckStr = ""

        j = 1
        For i = 0 To (DMaterial.Items.Count - 1)
            If (DMaterial.Items(i).Selected) Then
                If CheckStr = "" Then
                    CheckStr = Mid(Trim(DMaterial.Items(i).Text), 1, 2)
                Else
                    CheckStr = CheckStr + "," + Mid(Trim(DMaterial.Items(i).Text), 1, 2)
                End If



                'Dim DText As TextBox = mDirectCast(Page.FindControl("TextBox1"), TextBox)

                j = j + 1
            End If
            MaterialStr = CheckStr

        Next




        CheckStr = ""
        j = 1
        For i = 0 To (DLocation.Items.Count - 1)
            If (DLocation.Items(i).Selected) Then
                If CheckStr = "" Then
                    CheckStr = Mid(Trim(DLocation.Items(i).Text), 1, 2)
                Else
                    CheckStr = CheckStr + "," + Mid(Trim(DLocation.Items(i).Text), 1, 2)
                End If



                'Dim DText As TextBox = mDirectCast(Page.FindControl("TextBox1"), TextBox)

                j = j + 1
            End If
            LocationStr = CheckStr

        Next




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
        SQL = SQL + " where cat = '8002'"
        SQL = SQL + " and dkey ='AttachfilePath'"


        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If


        If D3.Text = "" Then
            D3.Text = Now.ToString("yyyyMMddHHmmss")
        End If



        OpenDir1 = OpenDir1 + D3.Text + "/核可卡"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + D3.Text + "\核可卡"
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If

        Dim dirInfo As New System.IO.DirectoryInfo(tempFolderPath)


        chktemp.Text = tempFolderPath
        '開啟附檔資料夾路徑
        DAttachFile1.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定附檔路徑1
    '**
    '*****************************************************************
    Sub NewAttachFilePath1()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '8002'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If





        OpenDir1 = OpenDir1 + DNo.Text + "/核可卡"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNo.Text + "\核可卡"
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If

        Dim dirInfo As New System.IO.DirectoryInfo(tempFolderPath)
        Dim FileDir As Integer  '資料夾
        FileDir = dirInfo.GetDirectories("*").Length
        Dim FileCount As Integer '檔案
        FileCount = dirInfo.GetFiles("*.*").Length


        chktemp.Text = tempFolderPath

        '開啟附檔資料夾路徑
        DAttachFile1.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")

    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定附檔路徑 2
    '**
    '*****************************************************************
    Sub NewAttachFilePath2()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '8002'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If





        OpenDir1 = OpenDir1 + "EDX/" + DQCNo.Text    '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + "EDX\" + DQCNo.Text
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If

        Dim dirInfo As New System.IO.DirectoryInfo(tempFolderPath)
        Dim FileDir As Integer  '資料夾
        FileDir = dirInfo.GetDirectories("*").Length
        Dim FileCount As Integer '檔案
        FileCount = dirInfo.GetFiles("*.*").Length

        '開啟附檔資料夾路徑
        DAttachFile2.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**    showform 時將 CheckType 字串資料，加入Checked 
    '**
    '*****************************************************************


    Function GetCheckType(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {","c})
        Dim Result As String = ""
        Dim j As Integer
        For Each i As String In sArray
            Result = i.ToString()

            For j = 0 To CInt(CheckTypeCount.Text) - 1
                If Mid(DCheckType.Items(j).Text, 1, 2) = Result Then
                    DCheckType.Items(j).Selected = True
                End If
                If wStep <> 1 And wStep <> 500 Then
                    DCheckType.Items(j).Enabled = False
                End If

            Next


        Next
        Return Result
    End Function


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**    showform 時將 CheckItem 字串資料，加入Checked 
    '**
    '*****************************************************************

    Function GetCheckItem(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {","c})
        Dim Result As String = ""
        Dim j As Integer
        For Each i As String In sArray
            Result = i.ToString()

            For j = 0 To CInt(CheckItemCount.Text) - 1

                If Mid(DCheckItem.Items(j).Text, 1, 2) = Result Then
                    DCheckItem.Items(j).Selected = True

                End If
                If wStep <> 1 And wStep <> 500 Then
                    DCheckItem.Items(j).Enabled = False
                End If


            Next

        Next
        Return Result
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**    showform 時將 RPReport 字串資料，加入Checked 
    '**
    '*****************************************************************


    Function GetRPReport(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {","c})
        Dim Result As String = ""
        Dim j As Integer
        For Each i As String In sArray
            Result = i.ToString()

            For j = 0 To CInt(RPReportCount.Text) - 1

                If Mid(DRPReport.Items(j).Text, 1, 2) = Result Then
                    DRPReport.Items(j).Selected = True
                End If
                If wStep <> 1 And wStep <> 500 Then
                    DRPReport.Items(j).Enabled = False
                End If

            Next


        Next
        Return Result
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**    showform 時將 Sample 字串資料，加入Checked 
    '**
    '*****************************************************************


    Function GetSample(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {","c})
        Dim Result As String = ""
        Dim j As Integer
        For Each i As String In sArray
            Result = i.ToString()

            For j = 0 To CInt(SampleCount.Text) - 1

                DSample.Items(j).Selected = False

                If Mid(DSample.Items(j).Text, 1, 2) = Result Then
                    DSample.Items(j).Selected = True
                End If
                If wStep <> 1 And wStep <> 500 Then
                    DSample.Items(j).Enabled = False
                End If

            Next


        Next
        Return Result
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**    showform 時將 Material 字串資料，加入Checked 
    '**
    '*****************************************************************

    Function GetMaterial(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {","c})
        Dim Result As String = ""
        Dim j As Integer
        For Each i As String In sArray
            Result = i.ToString()

            For j = 0 To CInt(MaterialCount.Text) - 1

                DMaterial.Items(j).Selected = False

                If Mid(DMaterial.Items(j).Text, 1, 2) = Result Then
                    DMaterial.Items(j).Selected = True
                End If
                If wStep <> 1 And wStep <> 500 Then
                    DMaterial.Items(j).Enabled = False
                End If

            Next


        Next
        Return Result
    End Function

    '*****************************************************************
    '**
    '**    showform 時將Location 字串資料，加入Checked 
    '**
    '*****************************************************************
    Function GetLocation(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {","c})
        Dim Result As String = ""
        Dim j As Integer
        For Each i As String In sArray
            Result = i.ToString()

            For j = 0 To CInt(LocationCount.Text) - 1

                DLocation.Items(j).Selected = False

                If Mid(DLocation.Items(j).Text, 1, 2) = Result Then
                    DLocation.Items(j).Selected = True
                End If
                If wStep <> 1 And wStep <> 500 Then
                    DLocation.Items(j).Enabled = False
                End If

            Next


        Next
        Return Result
    End Function





    Sub GetData()
        Dim SQL As String


        SQL = " select Unique_id,no,SeqNo,Type,Supplier,Code,Size,Family,Body,Puller,COLOR,Finish,ApproveNo,OldToNewNo,OldToNew,Remark,Element,Result1,Result2,QCRemark"
        SQL = SQL + " from  F_QASheetDTTemp "
        SQL = SQL + "  where 1=1 and No='" + DTNo.Text + "' "
        SQL = SQL + " order by SeqNo "
        'uDataBase.ExecuteNonQuery(SQL)
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            DCount.Text = "共" + CStr(dtData.Rows.Count) + "筆"
            GridView1.DataSource = uDataBase.GetDataTable(SQL)
            GridView1.DataBind()

        End If


    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        e.Row.Cells(0).Visible = False
    End Sub


    '取得實務確認者
    Function GetRelated(ByVal UserName As String) As String

        Dim sql As String = "select  Userid  from M_Users where username=N'" & UserName & "'"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            NextGate = dt.Rows(0)("UserID")
        End If
        Return NextGate
    End Function

    '*****************************************************************
    '**
    '**    showform 時將Location 字串資料，加入Checked 
    '**
    '*****************************************************************
    Function GetOLDNo(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {"-"c})
        Dim Result As String = ""
        Dim j As Integer = 0

        For Each i As String In sArray
            Result = i.ToString()

            If j = 0 Then
                ONo = Result
            ElseIf j = 1 Then
                OSeqno = Result
            ElseIf j = 2 Then
                OPuller = Result
            ElseIf j = 3 Then
                OColor = Result
            End If

            j = j + 1
        Next
        Return Result
    End Function


End Class
