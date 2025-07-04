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


Partial Class ComplaintInSheet_01
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
    Dim OpenDir1 As String = ""
    Dim OpenDir2 As String = ""
    Dim AttachFileName As String
    Dim sourceDir, backupDir As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'BASDate.Attributes("onclick") = "calendarPicker('DCheckDate');"
        ' BQCIDATE1.Attributes("onclick") = "calendarPicker('Form1.DAppDate');"
        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置
        SetControlPosition()    ' 設定控制項位置

        If Not Me.IsPostBack Then   '不是PostBack
            NewAttachFilePath()

            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()     '表單功能按鈕顯示

            If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
                ShowFormData()      '顯示表單資料
                NewAttachFilePath2()
                NewAttachFilePath3()

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
            Top = 1500
            Dim SQL As String = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(SQL)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0)("BEndTime").ToString < NowDateTime Then
                    If DDelay.Visible = True Then
                        Top = 1200
                    End If
                End If
            End If
        Else
            Top = 1500

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


        Top = 1500
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
        If wStep > 0 Then
            Top = 1450
        End If


        BSAVE.Style("top") = Top & "px"
        BNG1.Style("top") = Top & "px"
        BNG2.Style("top") = Top & "px"
        BOK.Style("top") = Top & "px"


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
        ' Response.Cookies("PGM").Value = "FinalcheckSheet_01.aspx"
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
                Top = 1000
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

        If wStep = 1 Then
            Top = 1350
        Else
            Top = 1500
        End If


        BSAVE.Style("top") = Top & "px"
        BNG1.Style("top") = Top & "px"
        BNG2.Style("top") = Top & "px"
        BOK.Style("top") = Top & "px"
        DHistoryLabel.Style("top") = Top + 32 & "px"
        GridView2.Style("top") = Top + 32 + 16 & "px"
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
        Dim InputData3 As String

        Dim SQL As String

        InputData1 = D1.Text
        InputData2 = D2.Text
        InputData3 = DCODE.Text



        '帶入CUSTOMER
        If InputData1 <> "" Then

            Sql = " Select  * from  MST_Custmer where 1=1 "
            If InputData1 <> "" Then
                Sql = Sql + " and ( custmer like '%" + InputData1 + "%' or name_c like '%" + InputData1 + "%')"
            End If
            Sql = Sql + " order by custmer,name_c "

            Dim DBData As DataTable = uDataBase.GetDataTable(Sql)

            DCUSTOMER.Text = DBData.Rows(0).Item("Name_C")
            DCUSTOMERCODE.Text = DBData.Rows(0).Item("Custmer")
        End If

        '帶入BUYER
        If InputData2 <> "" Then


            Sql = " Select  * from  MST_Buyer where 1=1 "
            If InputData2 <> "" Then
                Sql = Sql + " and ( buyer like '%" + InputData2 + "%' or buyer_name like '%" + InputData2 + "%')"
            End If

            Sql = Sql + " order by buyer_name,buyer "
            Dim DBData As DataTable = uDataBase.GetDataTable(Sql)

            DBUYER.Text = DBData.Rows(0).Item("Buyer_Name")
            DBuyerCode.Text = DBData.Rows(0).Item("Buyer")

        End If


        '帶入ITEM  '20210913
        If InputData3 <> "" Then


            Sql = " Select item,rtrim(item_name1)+' '+rtrim(item_name2)+' '+ltrim(item_name3) as itemname from Mst_item where 1=1 "
            If InputData3 <> "" Then
                Sql = Sql + " and ( item ='" + InputData3 + "' )"
            End If

            Sql = Sql + " order by item "
            Dim DBData As DataTable = uDataBase.GetDataTable(Sql)

            DSPECNAME.Text = DBData.Rows(0).Item("itemname")
            DSPEC.Text = DBData.Rows(0).Item("item")

        End If

        'No
        Select Case FindFieldInf("NO")
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



        'COMDATE
        Select Case FindFieldInf("COMDATE")
            Case 0  '顯示
                DCOMDATE.BackColor = Color.LightGray
                DCOMDATE.Visible = True
                DCOMDATE.Attributes.Add("readonly", "true")
                BCOMDATE.Visible = False
            Case 1  '修改+檢查
                DCOMDATE.Visible = True
                DCOMDATE.BackColor = Color.GreenYellow
                DCOMDATE.ReadOnly = False
                BCOMDATE.Visible = True
                ShowRequiredFieldValidator("DCOMDATERqd", "DCOMDATE", "異常：需輸入投訴日期")
            Case 2  '修改
                DCOMDATE.Visible = True
                DCOMDATE.BackColor = Color.Yellow
                DCOMDATE.ReadOnly = False
                BCOMDATE.Visible = True
            Case Else   '隱藏
                DCOMDATE.Visible = False
                BCOMDATE.Visible = False
        End Select


        '投訴部門
        Select Case FindFieldInf("COMDEPNAME")
            Case 0  '顯示
                DCOMDEPNAME.BackColor = Color.LightGray
                DCOMDEPNAME.Visible = True

            Case 1  '修改+檢查
                DCOMDEPNAME.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCOMDEPNAMERqd", "DCOMDEPNAME", "異常：需輸入投訴部門")
                DCOMDEPNAME.Visible = True
            Case 2  '修改
                DCOMDEPNAME.BackColor = Color.Yellow
                DCOMDEPNAME.Visible = True
            Case Else   '隱藏
                DCOMDEPNAME.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("COMDEPNAME", "ZZZZZZ")



        '投訴部門擔當
        Select Case FindFieldInf("COMNAME")
            Case 0  '顯示
                DCOMNAME.BackColor = Color.LightGray
                DCOMNAME.Visible = True
                DCOMNAME.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCOMNAME.Visible = True
                DCOMNAME.BackColor = Color.GreenYellow
                DCOMNAME.ReadOnly = False
                ShowRequiredFieldValidator("DCOMNAMERqd", "DCOMNAME", "異常：需輸入投訴部門擔當")
            Case 2  '修改
                DCOMNAME.Visible = True
                DCOMNAME.BackColor = Color.Yellow
                DCOMNAME.ReadOnly = False
            Case Else   '隱藏
                DCOMNAME.Visible = False
        End Select


        '投訴部門擔當
        Select Case FindFieldInf("COMADMIN")
            Case 0  '顯示
                DCOMADMIN.BackColor = Color.LightGray
                DCOMADMIN.Visible = True
                DCOMADMIN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCOMADMIN.Visible = True
                DCOMADMIN.BackColor = Color.GreenYellow
                DCOMADMIN.ReadOnly = False
                ShowRequiredFieldValidator("DCOMADMINRqd", "DCOMADMIN", "異常：需輸入投訴部門主管")
            Case 2  '修改
                DCOMADMIN.Visible = True
                DCOMADMIN.BackColor = Color.Yellow
                DCOMADMIN.ReadOnly = False
            Case Else   '隱藏
                DCOMADMIN.Visible = False
        End Select



        'Customer
        Select Case FindFieldInf("CUSTOMER")
            Case 0  '顯示
                DCUSTOMER.BackColor = Color.LightGray
                DCUSTOMER.ReadOnly = True
                DCUSTOMER.Visible = True
                DCUSTOMERCODE.BackColor = Color.LightGray
                DCUSTOMERCODE.ReadOnly = True
                DCUSTOMERCODE.Visible = True
                BCustomer.Visible = False

            Case 1  '修改+檢查
                DCUSTOMER.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerRqd", "DCustomer", "異常：需輸入客戶名稱")
                DCUSTOMER.Visible = True
                DCUSTOMER.ReadOnly = True
                DCUSTOMERCODE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerCodeRqd", "DCustomerCode", "異常：需輸入客戶名稱")
                DCUSTOMERCODE.Visible = True
                DCUSTOMERCODE.ReadOnly = True
                BCustomer.Visible = True

            Case 2  '修改
                DCUSTOMER.BackColor = Color.Yellow
                DCUSTOMER.Visible = True
                DCUSTOMER.ReadOnly = True
                DCUSTOMERCODE.BackColor = Color.Yellow
                DCUSTOMERCODE.Visible = True
                DCUSTOMERCODE.ReadOnly = True
                BCustomer.Visible = True

            Case Else   '隱藏
                DCUSTOMER.Visible = False
                DCUSTOMERCODE.Visible = False
                BCustomer.Visible = False

        End Select
        ' If pPost = "New" Then SetFieldData("CUSTOMER", "ZZZZZZ")


        'SPEC
        Select Case FindFieldInf("SPEC")
            Case 0  '顯示
                DSPEC.BackColor = Color.LightGray
                DSPEC.ReadOnly = True
                DSPEC.Visible = True
                BSPEC.Visible = False
            Case 1  '修改+檢查
                DSPEC.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSPECRqd", "DSPEC", "異常：需輸入型別")
                DSPEC.Visible = True
                DSPEC.ReadOnly = True
                BSPEC.Visible = True
            Case 2  '修改
                DSPEC.BackColor = Color.Yellow
                DSPEC.Visible = True
                BSPEC.Visible = True
                DSPEC.ReadOnly = True
            Case Else   '隱藏
                DSPEC.Visible = False
                BSPEC.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SPEC", "ZZZZZZ")



        'SPECNAME
        Select Case FindFieldInf("SPEC")
            Case 0  '顯示
                DSPECNAME.BackColor = Color.LightGray
                DSPECNAME.ReadOnly = True
                DSPECNAME.Visible = True
                BSPEC.Visible = False
            Case 1  '修改+檢查
                DSPECNAME.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSPECNAMERqd", "DSPECNAME", "異常：需輸入型別")
                DSPECNAME.Visible = True
                DSPECNAME.ReadOnly = True
                BSPEC.Visible = True
            Case 2  '修改
                DSPECNAME.BackColor = Color.Yellow
                DSPECNAME.Visible = True
                BBuyer.Visible = True
                DSPECNAME.ReadOnly = True
            Case Else   '隱藏
                DSPECNAME.Visible = False
                BSPEC.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SPECNAME", "ZZZZZZ")


        'Buyer
        Select Case FindFieldInf("BUYER")
            Case 0  '顯示
                DBUYER.BackColor = Color.LightGray
                DBUYER.ReadOnly = True
                DBUYER.Visible = True
                BBuyer.Visible = False
            Case 1  '修改+檢查
                DBUYER.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerRqd", "DBuyer", "異常：需輸入Buyer")
                DBUYER.Visible = True
                DBUYER.ReadOnly = True
                BBuyer.Visible = True
            Case 2  '修改
                DBUYER.BackColor = Color.Yellow
                DBUYER.Visible = True
                BBuyer.Visible = True
                DBUYER.ReadOnly = True

            Case Else   '隱藏
                DBUYER.Visible = False
                BBuyer.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BUYER", "ZZZZZZ")


        'BuyerCode
        Select Case FindFieldInf("BUYER")
            Case 0  '顯示
                DBuyerCode.BackColor = Color.LightGray
                DBuyerCode.ReadOnly = True
                DBuyerCode.Visible = True
                BBuyer.Visible = False
            Case 1  '修改+檢查
                DBuyerCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerCodeRqd", "DBuyerCode", "異常：需輸入Buyer")
                DBuyerCode.Visible = True
                DBuyerCode.ReadOnly = True
                BBuyer.Visible = True
            Case 2  '修改
                DBuyerCode.BackColor = Color.Yellow
                DBuyerCode.Visible = True
                BBuyer.Visible = True
                DBuyerCode.ReadOnly = True
            Case Else   '隱藏
                DBuyerCode.Visible = False
                BBuyer.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BuyerCode", "ZZZZZZ")




        '訂單號碼
        Select Case FindFieldInf("ORNO")
            Case 0  '顯示
                DORNO.BackColor = Color.LightGray
                DORNO.Visible = True
                DORNO.Attributes.Add("readonly", "true")
                BOrder.Visible = False
            Case 1  '修改+檢查
                DORNO.Visible = True
                DORNO.BackColor = Color.GreenYellow
                DORNO.ReadOnly = False
                ShowRequiredFieldValidator("DORNORqd", "DORNO", "異常：需輸入ORNO）")
                BOrder.Visible = True
            Case 2  '修改
                DORNO.Visible = True
                DORNO.BackColor = Color.Yellow
                DORNO.ReadOnly = False
                BOrder.Visible = True
            Case Else   '隱藏
                DORNO.Visible = False
                BOrder.Visible = False
        End Select



        '訂單號碼
        Select Case FindFieldInf("ColorCode")
            Case 0  '顯示
                DColorCode.BackColor = Color.LightGray
                DColorCode.Visible = True
                DColorCode.Attributes.Add("readonly", "true")
                BOrder.Visible = False
            Case 1  '修改+檢查
                DColorCode.Visible = True
                DColorCode.BackColor = Color.GreenYellow
                DColorCode.ReadOnly = False
                ShowRequiredFieldValidator("DColorCodeRqd", "DColorCode", "異常：需輸入ColorCode）")
                BOrder.Visible = True
            Case 2  '修改
                DColorCode.Visible = True
                DColorCode.BackColor = Color.Yellow
                DColorcode.ReadOnly = False
                BOrder.Visible = True
            Case Else   '隱藏
                DColorcode.Visible = False
                BOrder.Visible = False
        End Select



        'PR號碼
        Select Case FindFieldInf("Remark3")
            Case 0  '顯示
                DREMARK3.BackColor = Color.LightGray
                DREMARK3.Visible = True
                DREMARK3.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DREMARK3.Visible = True
                DREMARK3.BackColor = Color.GreenYellow
                DREMARK3.ReadOnly = False
                ShowRequiredFieldValidator("DREMARK3Rqd", "DREMARK3", "異常：需輸入REMARK3）")
            Case 2  '修改
                DREMARK3.Visible = True
                DREMARK3.BackColor = Color.Yellow
                DREMARK3.ReadOnly = False
            Case Else   '隱藏
                DREMARK3.Visible = False
        End Select



        '訂單數量
        Select Case FindFieldInf("ORQTY")
            Case 0  '顯示
                DORQTY.BackColor = Color.LightGray
                DORQTY.Visible = True
                DORQTY.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DORQTY.Visible = True
                DORQTY.BackColor = Color.GreenYellow
                DORQTY.ReadOnly = False
                ShowRequiredFieldValidator("DORQTYRqd", "DORQTY", "異常：需輸入訂單數量")
            Case 2  '修改
                DORQTY.Visible = True
                DORQTY.BackColor = Color.Yellow
                DORQTY.ReadOnly = False
            Case Else   '隱藏
                DORQTY.Visible = False
        End Select

        '數量單位
        Select Case FindFieldInf("UNIT1")
            Case 0  '顯示
                DUNIT1.BackColor = Color.LightGray
                DUNIT1.Visible = True

            Case 1  '修改+檢查
                DUNIT1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DUNIT1Rqd", "DUNIT1", "異常：需輸入數量單位")
                DUNIT1.Visible = True
            Case 2  '修改
                DUNIT1.BackColor = Color.Yellow
                DUNIT1.Visible = True
            Case Else   '隱藏
                DUNIT1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("UNIT1", "ZZZZZZ")




        '前工程材料號碼
        Select Case FindFieldInf("MATERIALNO")
            Case 0  '顯示
                DMATERIALNO.BackColor = Color.LightGray
                DMATERIALNO.Visible = True
                DMATERIALNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DMATERIALNO.Visible = True
                DMATERIALNO.BackColor = Color.GreenYellow
                DMATERIALNO.ReadOnly = False
                ShowRequiredFieldValidator("DMATERIALNORqd", "DMATERIALNO", "異常：需輸入前工程材料號碼")
            Case 2  '修改
                DMATERIALNO.Visible = True
                DMATERIALNO.BackColor = Color.Yellow
                DMATERIALNO.ReadOnly = False
            Case Else   '隱藏
                DMATERIALNO.Visible = False
        End Select




        '前工程材料號碼箱號
        Select Case FindFieldInf("Remark4")
            Case 0  '顯示
                DREMARK4.BackColor = Color.LightGray
                DREMARK4.Visible = True
                DREMARK4.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DREMARK4.Visible = True
                DREMARK4.BackColor = Color.GreenYellow
                DREMARK4.ReadOnly = False
                ShowRequiredFieldValidator("DREMARK4Rqd", "DREMARK4", "異常：需輸入前工程材料箱號")
            Case 2  '修改
                DREMARK4.Visible = True
                DREMARK4.BackColor = Color.Yellow
                DREMARK4.ReadOnly = False
            Case Else   '隱藏
                DREMARK4.Visible = False
        End Select



        '在庫確認
        Select Case FindFieldInf("STOCKCHECK")
            Case 0  '顯示
                DSTOCKCHECK.BackColor = Color.LightGray
                DSTOCKCHECK.Visible = True

            Case 1  '修改+檢查
                DSTOCKCHECK.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSTOCKCHECKRqd", "DSTOCKCHECK", "異常：需輸入在庫確認")
                DSTOCKCHECK.Visible = True
            Case 2  '修改
                DSTOCKCHECK.BackColor = Color.Yellow
                DSTOCKCHECK.Visible = True
            Case Else   '隱藏
                DSTOCKCHECK.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("STOCKCHECK", "ZZZZZZ")




        '在庫狀況
        Select Case FindFieldInf("STOCKSTS")
            Case 0  '顯示
                DSTOCKSTS.BackColor = Color.LightGray
                DSTOCKSTS.Visible = True
                DSTOCKSTS.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSTOCKSTS.Visible = True
                DSTOCKSTS.BackColor = Color.GreenYellow
                DSTOCKSTS.ReadOnly = False
                ShowRequiredFieldValidator("DSTOCKSTSRqd", "DSTOCKSTS", "異常：需輸入在庫狀況")
            Case 2  '修改
                DSTOCKSTS.Visible = True
                DSTOCKSTS.BackColor = Color.Yellow
                DSTOCKSTS.ReadOnly = False
            Case Else   '隱藏
                DSTOCKSTS.Visible = False
        End Select


        '內部客訴分類
        Select Case FindFieldInf("COMINTYPE")
            Case 0  '顯示
                DCOMINTYPE.BackColor = Color.LightGray
                DCOMINTYPE.Visible = True

            Case 1  '修改+檢查
                DCOMINTYPE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCOMINTYPERqd", "DCOMINTYPE", "異常：需輸入內部客訴分類")
                DCOMINTYPE.Visible = True
            Case 2  '修改
                DCOMINTYPE.BackColor = Color.Yellow
                DCOMINTYPE.Visible = True
            Case Else   '隱藏
                DCOMINTYPE.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("COMINTYPE", "ZZZZZZ")


        '不良狀況
        Select Case FindFieldInf("STATUS")
            Case 0  '顯示
                DSTATUS.BackColor = Color.LightGray
                DSTATUS.Visible = True
                DSTATUS.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSTATUS.Visible = True
                DSTATUS.BackColor = Color.GreenYellow
                DSTATUS.ReadOnly = False
                ShowRequiredFieldValidator("DSTATUSRqd", "DSTATUS", "異常：需輸入不良狀況")
            Case 2  '修改
                DSTATUS.Visible = True
                DSTATUS.BackColor = Color.Yellow
                DSTATUS.ReadOnly = False
            Case Else   '隱藏
                DSTATUS.Visible = False
        End Select


        '納期
        Select Case FindFieldInf("DELIVERYDATE")
            Case 0  '顯示
                DDELIVERYDATE.BackColor = Color.LightGray
                DDELIVERYDATE.Visible = True
                DDELIVERYDATE.Attributes.Add("readonly", "true")
                BDELIVERYDATE.Visible = False
            Case 1  '修改+檢查
                DDELIVERYDATE.Visible = True
                DDELIVERYDATE.BackColor = Color.GreenYellow
                DDELIVERYDATE.ReadOnly = False
                ShowRequiredFieldValidator("DDELIVERYDATERqd", "DDELIVERYDATE", "異常：需輸入納期")
                BDELIVERYDATE.Visible = True
            Case 2  '修改
                DDELIVERYDATE.Visible = True
                DDELIVERYDATE.BackColor = Color.Yellow
                DDELIVERYDATE.ReadOnly = False
                BDELIVERYDATE.Visible = True
            Case Else   '隱藏
                DDELIVERYDATE.Visible = False
                BDELIVERYDATE.Visible = False
        End Select


        '發生地
        Select Case FindFieldInf("LOCATION")
            Case 0  '顯示
                DLOCATION.BackColor = Color.LightGray
                DLOCATION.Visible = True

            Case 1  '修改+檢查
                DLOCATION.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DLOCATIONRqd", "DLOCATION", "異常：需輸入發生地")
                DLOCATION.Visible = True
            Case 2  '修改
                DLOCATION.BackColor = Color.Yellow
                DLOCATION.Visible = True
            Case Else   '隱藏
                DLOCATION.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("LOCATION", "ZZZZZZ")



        '最終再檢驗擔當
        Select Case FindFieldInf("LASTQC")
            Case 0  '顯示
                DLASTQC.BackColor = Color.LightGray
                DLASTQC.Visible = True

            Case 1  '修改+檢查
                DLASTQC.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DLASTQCRqd", "DLASTQC", "異常：需輸入最終再檢驗擔當")
                DLASTQC.Visible = True
            Case 2  '修改
                DLASTQC.BackColor = Color.Yellow
                DLASTQC.Visible = True
            Case Else   '隱藏
                DLASTQC.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("LASTQC", "ZZZZZZ")


        '投訴狀況
        Select Case FindFieldInf("StstusList")
            Case 0  '顯示
                DStstusList.BackColor = Color.LightGray
                DStstusList.Visible = True

            Case 1  '修改+檢查
                DStstusList.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DStstusListRqd", "DStstusList", "異常：需輸入投訴狀況")
                DStstusList.Visible = True
            Case 2  '修改
                DStstusList.BackColor = Color.Yellow
                DStstusList.Visible = True
            Case Else   '隱藏
                DStstusList.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("StstusList", "ZZZZZZ")





        '********************************************************************************************************************************************************
        ' 品管
        '********************************************************************************************************************************************************




        '品管受理NO
        Select Case FindFieldInf("QCNO")
            Case 0  '顯示
                DQCNO.BackColor = Color.LightGray
                DQCNO.Visible = True
                DQCNO.Attributes.Add("readonly", "true")
               
            Case 1  '修改+檢查
                DQCNO.Visible = True
                DQCNO.BackColor = Color.GreenYellow
                DQCNO.ReadOnly = False
                ShowRequiredFieldValidator("DQCNORqd", "DQCNO", "異常：需輸入品管受理NO")

            
            Case 2  '修改
                DQCNO.Visible = True
                DQCNO.BackColor = Color.Yellow
                DQCNO.ReadOnly = False

            Case Else   '隱藏
                DQCNO.Visible = False
                DQCDate.Visible = False

        End Select


        '被訴部門
        Select Case FindFieldInf("ACCDEPNAME")
            Case 0  '顯示
                DACCDEPNAME.BackColor = Color.LightGray
                DACCDEPNAME.Visible = True
                DACCEMPNAME.BackColor = Color.LightGray
                DACCEMPNAME.Visible = True
                DQCDate.BackColor = Color.LightGray
                DQCDate.Visible = True
                DQCDate.Attributes.Add("readonly", "true")
                BQCDate.Visible = False
            Case 1  '修改+檢查
                DACCDEPNAME.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DACCDEPNAMERqd", "DACCDEPNAME", "異常：需輸入被訴部門")
                DACCDEPNAME.Visible = True
                DACCEMPNAME.BackColor = Color.LightGray
                DACCEMPNAME.Visible = True
                DQCDate.Visible = True
                DQCDate.BackColor = Color.GreenYellow
                DQCDate.ReadOnly = False
                ShowRequiredFieldValidator("DQCDATERqd", "DQCDATE", "異常：需輸入品管受理日期")
                BQCDate.Visible = True
            Case 2  '修改
                DACCDEPNAME.BackColor = Color.Yellow
                DACCDEPNAME.Visible = True
                DACCEMPNAME.BackColor = Color.LightGray
                DACCEMPNAME.Visible = True

                DQCDate.Visible = True
                DQCDate.BackColor = Color.Yellow
                DQCDate.ReadOnly = False
                BQCDate.Visible = True
            Case Else   '隱藏
                DACCDEPNAME.Visible = False
                DQCDate.Visible = False

        End Select
        If pPost = "New" Then SetFieldData("ACCDEPNAME", "ZZZZZZ")



        '被訴部門擔當
        Select Case FindFieldInf("ACCEMPNAME")
            Case 0  '顯示
                DACCEMPNAME.BackColor = Color.LightGray
                DACCEMPNAME.Visible = True
                DACCEMPNAME.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DACCEMPNAME.Visible = True
                DACCEMPNAME.BackColor = Color.GreenYellow
                DACCEMPNAME.ReadOnly = False
                ShowRequiredFieldValidator("DACCEMPNAMERqd", "DACCEMPNAME", "異常：需輸入被訴部門擔當")
            Case 2  '修改
                DACCEMPNAME.Visible = True
                DACCEMPNAME.BackColor = Color.Yellow
                DACCEMPNAME.ReadOnly = False
            Case Else   '隱藏
                DACCEMPNAME.Visible = False
        End Select



        '備註
        Select Case FindFieldInf("REMARK1")
            Case 0  '顯示
                DREMARK1.BackColor = Color.LightGray
                DREMARK1.Visible = True
                DREMARK1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DREMARK1.Visible = True
                DREMARK1.BackColor = Color.GreenYellow
                DREMARK1.ReadOnly = False
                ShowRequiredFieldValidator("DREMARK1Rqd", "DREMARK1", "異常：需輸入備註")
            Case 2  '修改
                DREMARK1.Visible = True
                DREMARK1.BackColor = Color.Yellow
                DREMARK1.ReadOnly = False
            Case Else   '隱藏
                DREMARK1.Visible = False
        End Select



        '簡圖
        Select Case FindFieldInf("MapFile")
            Case 0  '顯示
                DMapFile.Visible = False
                DMapFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                'ShowRequiredFieldValidator("DMapFileRqd", "DMapFile", "異常：需上傳不良圖片")
                DMapFile.Visible = True
                DMapFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DMapFile.Visible = True
                DMapFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DMapFile.Visible = False
        End Select
        LMapFile.Visible = False
        LMapFile1.Visible = False


        '********************************************************************************************************************************************************
        ' 責任部門
        '********************************************************************************************************************************************************



        '投訴部門及波及性確認       
        Select Case FindFieldInf("FCONFIRM")
            Case 0  '顯示
                DFCONFIRM.BackColor = Color.LightGray
                DFCONFIRM.Visible = True

            Case 1  '修改+檢查
                DFCONFIRM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFCONFIRMRqd", "DFCONFIRM", "異常：需輸入波及性確認")
                DFCONFIRM.Visible = True
            Case 2  '修改
                DFCONFIRM.BackColor = Color.Yellow
                DFCONFIRM.Visible = True
            Case Else   '隱藏
                DFCONFIRM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("FCONFIRM", "ZZZZZZ")




        '內容別
        Select Case FindFieldInf("FCONTENT")
            Case 0  '顯示
                DFCONTENT.BackColor = Color.LightGray
                DFCONTENT.Visible = True
                DMach1.Visible = False
                DMach2.Visible = False
                DMACHNO.Visible = False
            Case 1  '修改+檢查
                DFCONTENT.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFCONTENTRqd", "DFCONTENT", "異常：需輸入內容別")
                DFCONTENT.Visible = True

                DMach1.Visible = True
                DMach2.Visible = True
                DMACHNO.Visible = True

            Case 2  '修改
                DFCONTENT.BackColor = Color.Yellow
                DFCONTENT.Visible = True

                DMach1.Visible = True
                DMach2.Visible = True
                DMACHNO.Visible = True
            Case Else   '隱藏
                DFCONTENT.Visible = False
                DMach1.Visible = False
                DMach2.Visible = False
                DMACHNO.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("FCONTENT", "ZZZZZZ")

        '原因
        Select Case FindFieldInf("FREASON")
            Case 0  '顯示
                DFREASON.BackColor = Color.LightGray
                DFREASON.Visible = True

            Case 1  '修改+檢查
                DFREASON.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFREASONRqd", "DFREASON", "異常：需輸入補貨")
                DFREASON.Visible = True
            Case 2  '修改
                DFREASON.BackColor = Color.Yellow
                DFREASON.Visible = True
            Case Else   '隱藏
                DFREASON.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("FREASON", "ZZZZZZ")






        '處理情形
        Select Case FindFieldInf("SITUATION")
            Case 0  '顯示
                DSITUATION.BackColor = Color.LightGray
                DSITUATION.Visible = True

            Case 1  '修改+檢查
                DSITUATION.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSITUATIONRqd", "DSITUATION", "異常：需輸入處理情形")
                DSITUATION.Visible = True
            Case 2  '修改
                DSITUATION.BackColor = Color.Yellow
                DSITUATION.Visible = True
            Case Else   '隱藏
                DSITUATION.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SITUATION", "ZZZZZZ")

        '原因矯正對策
        Select Case FindFieldInf("ANSWER")
            Case 0  '顯示
                DANSWER.BackColor = Color.LightGray
                DANSWER.Visible = True

            Case 1  '修改+檢查
                DANSWER.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DANSWERRqd", "DANSWER", "異常：需輸入原因矯正對策")
                DANSWER.Visible = True
            Case 2  '修改
                DANSWER.BackColor = Color.Yellow
                DANSWER.Visible = True
            Case Else   '隱藏
                DANSWER.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ANSWER", "ZZZZZZ")




        '全檢後不良數量
        Select Case FindFieldInf("ERRORQTY")
            Case 0  '顯示
                DERRORQTY.BackColor = Color.LightGray
                DERRORQTY.Visible = True
                DERRORQTY.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DERRORQTY.Visible = True
                DERRORQTY.BackColor = Color.GreenYellow
                DERRORQTY.ReadOnly = False
                ShowRequiredFieldValidator("DERRORQTYRqd", "DERRORQTY", "異常：需輸入全檢後不良數量")
            Case 2  '修改
                DERRORQTY.Visible = True
                DERRORQTY.BackColor = Color.Yellow
                DERRORQTY.ReadOnly = False
            Case Else   '隱藏
                DERRORQTY.Visible = False
        End Select



        '全檢後不良%
        Select Case FindFieldInf("ERRORP")
            Case 0  '顯示
                DERRORP.BackColor = Color.LightGray
                DERRORP.Visible = True
                DERRORP.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DERRORP.Visible = True
                DERRORP.BackColor = Color.GreenYellow
                DERRORP.ReadOnly = False
                ShowRequiredFieldValidator("DERRORPRqd", "DERRORP", "異常：需輸入全檢後不良%")
            Case 2  '修改
                DERRORP.Visible = True
                DERRORP.BackColor = Color.Yellow
                DERRORP.ReadOnly = False
            Case Else   '隱藏
                DERRORP.Visible = False
        End Select



        'SHIP
        Select Case FindFieldInf("SHIP")
            Case 0  '顯示
                DSHIP.BackColor = Color.LightGray
                DSHIP.Visible = True

            Case 1  '修改+檢查
                DSHIP.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSHIPRqd", "DSHIP", "異常：需輸入是否有單品出貨")
                DSHIP.Visible = True
            Case 2  '修改
                DSHIP.BackColor = Color.Yellow
                DSHIP.Visible = True
            Case Else   '隱藏
                DSHIP.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SHIP", "ZZZZZZ")


        'SLD工程別             
        Select Case FindFieldInf("SLDDIVISION")
            Case 0  '顯示
                DSLDDIVISION.BackColor = Color.LightGray
                DSLDDIVISION.Visible = True

            Case 1  '修改+檢查
                DSLDDIVISION.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSLDDIVISIONRqd", "DSLDDIVISION", "異常：需輸入SLD工程別")
                DSLDDIVISION.Visible = True
            Case 2  '修改
                DSLDDIVISION.BackColor = Color.Yellow
                DSLDDIVISION.Visible = True
            Case Else   '隱藏
                DSLDDIVISION.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SLDDIVISION", "ZZZZZZ")


        '單位
        Select Case FindFieldInf("UNIT2")
            Case 0  '顯示
                DUNIT2.BackColor = Color.LightGray
                DUNIT2.Visible = True
                DUNIT2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DUNIT2.Visible = True
                DUNIT2.BackColor = Color.GreenYellow

                ShowRequiredFieldValidator("DUNIT2Rqd", "DUNIT2", "異常：需輸入單位")
            Case 2  '修改
                DUNIT2.Visible = True
                DUNIT2.BackColor = Color.Yellow

            Case Else   '隱藏
                DUNIT2.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("UNIT2", "ZZZZZZ")



        '備註
        Select Case FindFieldInf("REMARK2")
            Case 0  '顯示
                DREMARK2.BackColor = Color.LightGray
                DREMARK2.Visible = True
                DREMARK2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DREMARK2.Visible = True
                DREMARK2.BackColor = Color.GreenYellow
                DREMARK2.ReadOnly = False
                ShowRequiredFieldValidator("DREMARK2Rqd", "DREMARK2", "異常：需輸入備註")
            Case 2  '修改
                DREMARK2.Visible = True
                DREMARK2.BackColor = Color.Yellow
                DREMARK2.ReadOnly = False
            Case Else   '隱藏
                DREMARK2.Visible = False
        End Select

        '補退數量
        Select Case FindFieldInf("BUQTY")
            Case 0  '顯示
                DBUQTY.BackColor = Color.LightGray
                DBUQTY.Visible = True
                DBUQTY.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DBUQTY.Visible = True
                DBUQTY.BackColor = Color.GreenYellow
                DBUQTY.ReadOnly = False
                ShowRequiredFieldValidator("DBUQTYRqd", "DBUQTY", "異常：需輸入補退數量")
            Case 2  '修改
                DBUQTY.Visible = True
                DBUQTY.BackColor = Color.Yellow
                DBUQTY.ReadOnly = False
            Case Else   '隱藏
                DBUQTY.Visible = False
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



        Select Case FindFieldInf("ATTACHFILE2")
            Case 0  '顯示            

                DAttachfile2.Visible = True



            Case 1  '修改+檢查
                DAttachfile2.Visible = True

            Case 2  '修改
                DAttachfile2.Visible = True


            Case Else   '隱藏
                DAttachfile2.Visible = False

        End Select


        Select Case FindFieldInf("ATTACHFILE3")
            Case 0  '顯示            

                DAttachfile3.Visible = True



            Case 1  '修改+檢查
                DAttachfile3.Visible = True

            Case 2  '修改
                DAttachfile3.Visible = True


            Case Else   '隱藏
                DAttachfile3.Visible = False

        End Select





    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()

        Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                    CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("Http") & _
                    System.Configuration.ConfigurationManager.AppSettings("ComplaintInPath")
        Dim FileName As String = ""

        Dim SQL As String
        SQL = "Select * From F_ComplaintInSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            '表單資料
            DNO.Text = dtData.Rows(0).Item("No")                         'No




            DCOMDATE.Text = dtData.Rows(0).Item("COMDATE")
            SetFieldData("COMDEPNAME", dtData.Rows(0).Item("COMDEPNAME"))
            DCOMNAME.Text = dtData.Rows(0).Item("COMNAME")
            DCOMADMIN.Text = dtData.Rows(0).Item("COMADMIN")

            DCUSTOMERCODE.Text = dtData.Rows(0).Item("CUSTOMERCODE")
            DCUSTOMER.Text = dtData.Rows(0).Item("CUSTOMER")
            DBuyerCode.Text = dtData.Rows(0).Item("BUYERCODE")
            DBUYER.Text = dtData.Rows(0).Item("BUYER")
            DBuyerCode.Text = dtData.Rows(0).Item("ORNO")
            DColorcode.Text = dtData.Rows(0).Item("Colorcode")
            DREMARK3.Text = dtData.Rows(0).Item("REMARK3")
            DSPEC.Text = dtData.Rows(0).Item("SPEC")
            DSPECNAME.Text = dtData.Rows(0).Item("SPECNAME")
            DORQTY.Text = dtData.Rows(0).Item("ORQTY")
            SetFieldData("UNIT1", dtData.Rows(0).Item("UNIT1"))
            DMATERIALNO.Text = dtData.Rows(0).Item("MATERIALNO")
            DREMARK4.Text = dtData.Rows(0).Item("REMARK4")
            DSTOCKSTS.Text = dtData.Rows(0).Item("STOCKSTS")
            SetFieldData("STOCKCHECK", dtData.Rows(0).Item("STOCKCHECK"))
            DCOMINTYPE.Text = dtData.Rows(0).Item("COMINTYPE")
            DSTATUS.Text = dtData.Rows(0).Item("STATUS")

            If Mid(dtData.Rows(0).Item("DELIVERYDATE").ToString, 1, 4) = "1900" Then
                DDELIVERYDATE.Text = ""
            Else
                DDELIVERYDATE.Text = dtData.Rows(0).Item("DELIVERYDATE")
            End If
            SetFieldData("LOCATION", dtData.Rows(0).Item("LOCATION"))
            SetFieldData("LASTQC", dtData.Rows(0).Item("LASTQC"))
            SetFieldData("StstusList", dtData.Rows(0).Item("StstusList"))

            DQCNO.Text = dtData.Rows(0).Item("QCNO")
            SetFieldData("ACCDEPNAME", dtData.Rows(0).Item("ACCDEPNAME"))
            DACCEMPNAME.Text = dtData.Rows(0).Item("ACCEMPNAME")
            DREMARK1.Text = dtData.Rows(0).Item("REMARK1")

            DMappath.Text = Path & dtData.Rows(0).Item("MapFile")

            '不良品樣品圖
            If Trim(dtData.Rows(0).Item("MapFile")) <> "" Then

                LMapFile.ImageUrl = DMappath.Text
                LMapFile1.NavigateUrl = DMappath.Text
                LMapFile.Visible = True
                LMapFile1.Visible = True
            Else
                LMapFile.Visible = False
                LMapFile1.Visible = False
            End If

            SetFieldData("FCONFIRM", dtData.Rows(0).Item("FCONFIRM"))
            SetFieldData("FCONTENT", dtData.Rows(0).Item("FCONTENT"))
            SetFieldData("FREASON", dtData.Rows(0).Item("FREASON"))
            SetFieldData("SITUATION", dtData.Rows(0).Item("SITUATION"))
            SetFieldData("ANSWER", dtData.Rows(0).Item("ANSWER"))
            DERRORQTY.Text = dtData.Rows(0).Item("ERRORQTY")
            DERRORP.Text = dtData.Rows(0).Item("ERRORP")

            SetFieldData("SHIP", dtData.Rows(0).Item("SHIP"))
            SetFieldData("SLDDIVISION", dtData.Rows(0).Item("SLDDIVISION"))
            SetFieldData("UNIT2", dtData.Rows(0).Item("UNIT2"))
            DREMARK2.Text = dtData.Rows(0).Item("REMARK2")
            DBUQTY.Text = dtData.Rows(0).Item("BUQTY")


            If dtData.Rows(0).Item("MACH") = "0" Then
                DMach1.Checked = True
                DMach2.Checked = False
            ElseIf dtData.Rows(0).Item("MACH") = "1" Then
                DMach2.Checked = True
                DMach1.Checked = False
            Else
                DMach1.Checked = False
                DMach2.Checked = False
            End If

            If dtData.Rows(0).Item("QCDate") = "1900-01-01 00:00:00.000" Then
                DQCDate.Text =""
            Else
                DQCDate.Text = dtData.Rows(0).Item("QCDate")
            End If


            '主檔資料
            SQL = " select  data,replace(data,'/','\')data1  from M_referp"
            SQL = SQL + " where cat = '3108'"
            SQL = SQL + " and dkey ='AttachfilePath1'"
            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
            Dim OpenDir As String = ""
            If DBAdapter3.Rows.Count > 0 Then
                OpenDir = "file://" + DBAdapter3.Rows(0).Item("Data") + DNO.Text + "/投訴部門"
            End If
            Dim OpenDir2 As String = ""
            OpenDir2 = "\\" + DBAdapter3.Rows(0).Item("Data1") + DNO.Text + "\投訴部門"

            Dim dirInfo As New System.IO.DirectoryInfo(OpenDir2)
            chktemp.Text = OpenDir2

            Dim FileDir As Integer  '資料夾
            FileDir = dirInfo.GetDirectories("*").Length
            Dim FileCount As Integer '檔案
            FileCount = dirInfo.GetFiles("*.*").Length


            If FileCount > 0 Or FileDir > 0 Then
                DChkData1.Checked = True

            Else
                DChkData1.Checked = False

            End If


            DAttachfile1.Attributes.Add("onclick", "window.open('" + OpenDir + "','_blank');return false;")
            '  DAttachfile2.Attributes.Add("onclick", "window.open('" + OpenDir2 + "','_blank');return false;")



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
        '擔當者及部門 

        sql = "Select Divname,Username From M_Users "
        sql = sql & " Where UserID = '" & wApplyID & "'"
        sql = sql & "   And Active = '1' "
        Dim DBUser As DataTable = uDataBase.GetDataTable(sql)

        '被訴部門
        If pFieldName = "ACCDEPNAME" Then
            DACCDEPNAME.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DACCDEPNAME.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  data  from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'COMDEPNAME'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DACCDEPNAME.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DACCDEPNAME.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        '投訴部門
        If pFieldName = "COMDEPNAME" Then
            DCOMDEPNAME.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCOMDEPNAME.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  data  from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'COMDEPNAME'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DCOMDEPNAME.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCOMDEPNAME.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        '投訴部門
        If pFieldName = "LASTQC" Then
            DLASTQC.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLASTQC.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  data  from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'LASTQC'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DLASTQC.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLASTQC.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If




        '投訴狀況
        If pFieldName = "StstusList" Then
            DStstusList.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DStstusList.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  data  from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'statuslist'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DStstusList.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DStstusList.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If





        'STOCKCHECK
        If pFieldName = "STOCKCHECK" Then
            DSTOCKCHECK.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSTOCKCHECK.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'STOCKCHECK'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DSTOCKCHECK.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSTOCKCHECK.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        'COMINTYPE
        If pFieldName = "COMINTYPE" Then
            DCOMINTYPE.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCOMINTYPE.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'COMINTYPE'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DCOMINTYPE.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCOMINTYPE.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If





        'LOCATION
        If pFieldName = "LOCATION" Then
            DLOCATION.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLOCATION.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'LOCATION'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DLOCATION.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLOCATION.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If




        'SLD工程別        
        If pFieldName = "SLDDIVISION" Then
            DSLDDIVISION.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSLDDIVISION.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  data  from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'COMDEPNAME'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DSLDDIVISION.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSLDDIVISION.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If




        'FREASON
        If pFieldName = "FREASON" Then
            DFREASON.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DFREASON.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'FREASON'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DFREASON.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DFREASON.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        'SHIP
        If pFieldName = "SHIP" Then
            DSHIP.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSHIP.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'SHIP'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DSHIP.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSHIP.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If




        'FCONTENT
        If pFieldName = "FCONTENT" Then
            DFCONTENT.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DFCONTENT.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'FCONTENT'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DFCONTENT.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DFCONTENT.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If




        'SITUATION
        If pFieldName = "SITUATION" Then
            DSITUATION.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSITUATION.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'SITUATION'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DSITUATION.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSITUATION.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If





        'unit1
        If pFieldName = "UNIT1" Then
            DUNIT1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DUNIT1.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'CHECKQTY'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DUNIT1.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DUNIT1.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If




        'UNIT2
        If pFieldName = "UNIT2" Then
            DUNIT2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DUNIT2.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'CHECKQTY'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DUNIT2.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DUNIT2.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If





        'CONFIRM2
        If pFieldName = "FCONFIRM" Then
            DFCONFIRM.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DFCONFIRM.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'CONFIRM2'"
                sql = sql & " order by DATA"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DFCONFIRM.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DFCONFIRM.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        'QCANSWER
        If pFieldName = "ANSWER" Then
            DANSWER.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DANSWER.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'QCANSWER'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DANSWER.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DANSWER.Items.Add(ListItem1)
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

        BCOMDATE.Attributes("onclick") = "calendarPicker('Form1.DCOMDATE');"

        BQCDate.Attributes("onclick") = "calendarPicker('Form1.DQCDate');"

        BDELIVERYDATE.Attributes("onclick") = "calendarPicker('Form1.DDELIVERYDATE');"

        BCustomer.Attributes.Add("onClick", "GetCustomer()") '找客戶
        BBuyer.Attributes.Add("onClick", "GetBuyer()") '找buyer
        BSPEC.Attributes.Add("onClick", "GetSPEC()") '找buyer
        BOrder.Attributes.Add("onClick", "GetOrder()") '找buyer

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
                        DNO.Text = SetNo(NewFormSno)
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


                'CC品管確認後給責任部門台日籍主管
                If wStep = 10 And pFun = "OK" Then

                    Dim SQL1 As String
                    Dim stri As Integer
                    stri = InStr(1, DACCDEPNAME.SelectedValue, "-")

                    SQL1 = " SELECT  DATA"
                    SQL1 = SQL1 + "  from M_referp where cat ='3108' and left(dkey,3) = 'DEP' and SUBSTRING(dkey,CHARINDEX('-',dkey)+1,len(dkey)) ='" + Mid(DACCDEPNAME.SelectedValue, 1, stri - 1) + "'"

                    Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)
                    pCount = dtReferp.Rows.Count
                    For i = 0 To dtReferp.Rows.Count - 1
                        pNextGate(i + 1) = GetUerID(dtReferp.Rows(i).Item("DATA"))
                    Next
                    dtReferp.Clear()
                    RtnCode = 0
                    pRunNextStep = 1
                End If

            End If

            '再送給工程長簽核
            If wStep = 15 Then

                Dim SQL1 As String

                SQL1 = " SELECT * from M_referp where cat ='3108' and  dkey ='10" + DACCDEPNAME.SelectedValue + "'"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)
                pCount = dtReferp.Rows.Count
                For i = 0 To dtReferp.Rows.Count - 1
                    pNextGate(i + 1) = GetUerID(dtReferp.Rows(i).Item("DATA"))
                Next
                dtReferp.Clear()
                RtnCode = 0
                pRunNextStep = 1
                'End If

            End If

            ''CC投訴部門工程長
            'If wStep = 30 Then

            '    Dim SQL1 As String

            '    SQL1 = " SELECT * from M_referp where cat ='3108' and  dkey ='10" + DCOMDEPNAME.SelectedValue + "'"

            '    Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)
            '    pCount = dtReferp.Rows.Count
            '    For i = 0 To dtReferp.Rows.Count - 1
            '        pNextGate(i + 1) = GetUerID(dtReferp.Rows(i).Item("DATA"))
            '    Next
            '    dtReferp.Clear()
            '    RtnCode = 0
            '    pRunNextStep = 1
            '    'End If

            'End If



            ''CC責任部門工程長
            'If wStep = 35 Then

            '    Dim SQL1 As String

            '    SQL1 = " SELECT * from M_referp where cat ='3108' and  dkey ='10" + DACCDEPNAME.SelectedValue + "'"

            '    Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)
            '    pCount = dtReferp.Rows.Count
            '    For i = 0 To dtReferp.Rows.Count - 1
            '        pNextGate(i + 1) = GetUerID(dtReferp.Rows(i).Item("DATA"))
            '    Next
            '    dtReferp.Clear()
            '    RtnCode = 0
            '    pRunNextStep = 1
            '    'End If

            'End If



            If RtnCode <> 0 Then ErrCode = 9130
            If pCount = 0 And pNextStep <> 999 Then ErrCode = 9131
            If ErrCode = 0 Then pAction = 0
          
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
            Dim URL As String = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
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

        Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                        CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                    System.Configuration.ConfigurationManager.AppSettings("ComplaintInPath")
        Dim FileName As String
        Dim sql As String = ""
        sql = " Insert into F_ComplaintInSheet (Sts, CompletedTime, FormNo, FormSno,"
        sql = sql + " NO,DATE,COMDATE,COMNAME,COMDEPNAME,COMADMIN,CUSTOMERCODE,CUSTOMER,BUYERCODE,BUYER,"
        sql = sql + " ORNO,ColorCode,REMARK3,SPEC,SPECNAME,ORQTY,UNIT1,MATERIALNO,REMARK4,STOCKSTS,STOCKCHECK,COMINTYPE,"
        sql = sql + " STATUS,DELIVERYDATE,LOCATION,LASTQC,StstusList,"
        sql = sql + " QCNO, ACCDEPNAME, ACCEMPNAME,REMARK1,MAPFILE, "
        sql = sql + " FCONFIRM,FCONTENT,FREASON,SITUATION,ANSWER,"
        sql = sql + " ERRORQTY,ERRORP,SHIP,SLDDIVISION,UNIT2,REMARK2,BUQTY,"

        sql = sql + " CreateUser, CreateTime, ModifyUser, ModifyTime) "
        sql = sql + " VALUES( "
        sql = sql + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        sql = sql + " '" + NowDateTime + "', "        '結案日
        sql = sql + " '003108', "                     '表單代號
        sql = sql + " '" + CStr(NewFormSno) + "', "   '表單流水號

        sql = sql + " N'" + YKK.ReplaceString(DNO.Text) + "', "   'NO  1
        sql = sql + " N'" + NowDateTime + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DCOMDATE.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DCOMNAME.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DCOMDEPNAME.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DCOMADMIN.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DCUSTOMERCODE.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DCUSTOMER.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DBuyerCode.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DBUYER.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DORNO.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DColorcode.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DREMARK3.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DSPEC.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DSPECNAME.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DORQTY.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DUNIT1.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DMATERIALNO.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DREMARK4.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DSTOCKSTS.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DSTOCKCHECK.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DCOMINTYPE.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DSTATUS.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DDELIVERYDATE.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DLOCATION.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DLASTQC.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DStstusList.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DQCNO.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DACCDEPNAME.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DACCEMPNAME.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DREMARK1.Text) + "', "
        FileName = ""
        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "MapFile" & "-" & UploadDateTime & System.IO.Path.GetExtension(DMapFile.PostedFile.FileName)

                DMapFile.PostedFile.SaveAs(Path + FileName)

            Else
                FileName = ""
            End If
        End If
        sql = sql + " '" + FileName + "'," '不良樣品圖

        sql = sql + " N'" + YKK.ReplaceString(DFCONFIRM.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DFCONTENT.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DFREASON.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DSITUATION.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DANSWER.SelectedValue) + "', "

        sql = sql + " N'" + YKK.ReplaceString(DERRORQTY.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DERRORP.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DSHIP.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DSLDDIVISION.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DUNIT2.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DREMARK2.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DBUQTY.Text) + "', "

        sql = sql + "'" & Request.QueryString("pUserID") & "' ,"
        sql = sql + "'" & NowDateTime & "' ,"
        sql = sql + "'" & Request.QueryString("pUserID") & "' ,"
        sql = sql + "'" & NowDateTime & "' )"
        uDataBase.ExecuteNonQuery(sql)


        '主檔資料
        sql = " select  data,replace(data,'/','\')data1  from M_referp"
        sql = sql + " where cat = '3108'"
        sql = sql + " and dkey ='AttachfilePath' "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

        If DBAdapter1.Rows.Count > 0 Then
            sourceDir = "\\" + DBAdapter1.Rows(0).Item("Data1") + D3.Text   '來源
        End If

        '主檔資料
        sql = " select  data,replace(data,'/','\')data1  from M_referp"
        sql = sql + " where cat = '3108'"
        sql = sql + " and dkey ='AttachfilePath1' "
        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(sql)

        If DBAdapter2.Rows.Count > 0 Then
            backupDir = "\\" + DBAdapter2.Rows(0).Item("Data1") + DNO.Text      '目的  
        End If


        'sourceDir = "\\10.245.1.61\wfs$\N2W\003109Temp\" + D3.Text   '來源
        'backupDir = "\\10.245.1.61\wfs$\N2W\003109\" + DNO.Text      '目的     

        '複製前先確認是否有原本的資料夾再刪除
        If Directory.Exists(backupDir) Then
            Directory.Delete(backupDir, True)
        End If
        CopyDir(sourceDir, backupDir)
        NewAttachFilePath()




    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)

        Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                         CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                    System.Configuration.ConfigurationManager.AppSettings("ComplaintInPath")
        Dim sql As String
        Dim FileName As String = ""

        sql = " Update F_ComplaintInSheet"
        sql = sql + " Set "
        If pFun <> "SAVE" Then
            sql = sql + " Sts = '" & pSts & "',"
            sql = sql + " CompletedTime = '" & NowDateTime & "',"
        End If
        sql = sql + " No = N'" & YKK.ReplaceString(DNO.Text) & "',"

        sql = sql + " COMDATE = N'" + YKK.ReplaceString(DCOMDATE.Text) + "', "
        sql = sql + " COMNAME= N'" + YKK.ReplaceString(DCOMNAME.Text) + "', "
        sql = sql + " COMDEPNAME=  N'" + YKK.ReplaceString(DCOMDEPNAME.SelectedValue) + "', "
        sql = sql + " COMADMIN = N'" + YKK.ReplaceString(DCOMADMIN.Text) + "', "
        sql = sql + " CUSTOMERCODE = N'" + YKK.ReplaceString(DCUSTOMERCODE.Text) + "', "
        sql = sql + " CUSTOMER = N'" + YKK.ReplaceString(DCUSTOMER.Text) + "', "
        sql = sql + " BUYERCODE = N'" + YKK.ReplaceString(DBuyerCode.Text) + "', "
        sql = sql + " BUYER = N'" + YKK.ReplaceString(DBUYER.Text) + "', "
        sql = sql + " ORNO = N'" + YKK.ReplaceString(DORNO.Text) + "', "
        sql = sql + " ColorCode = N'" + YKK.ReplaceString(DColorcode.Text) + "', "
        sql = sql + " REMARK3 = N'" + YKK.ReplaceString(DREMARK3.Text) + "', "
        sql = sql + " SPEC = N'" + YKK.ReplaceString(DSPEC.Text) + "', "
        sql = sql + " SPECNAME = N'" + YKK.ReplaceString(DSPECNAME.Text) + "', "
        sql = sql + " ORQTY = N'" + YKK.ReplaceString(DORQTY.Text) + "', "
        sql = sql + " UNIT1 = N'" + YKK.ReplaceString(DUNIT1.SelectedValue) + "', "
        sql = sql + " MATERIALNO = N'" + YKK.ReplaceString(DMATERIALNO.Text) + "', "
        sql = sql + " REMARK4 = N'" + YKK.ReplaceString(DREMARK4.Text) + "', "
        sql = sql + " STATUS = N'" + YKK.ReplaceString(DSTATUS.Text) + "', "
        sql = sql + " STOCKCHECK = N'" + YKK.ReplaceString(DSTOCKCHECK.SelectedValue) + "', "
        sql = sql + " COMINTYPE = N'" + YKK.ReplaceString(DCOMINTYPE.SelectedValue) + "', "
        sql = sql + " DELIVERYDATE = N'" + YKK.ReplaceString(DDELIVERYDATE.Text) + "', "
        sql = sql + " LOCATION = N'" + YKK.ReplaceString(DLOCATION.SelectedValue) + "', "
        sql = sql + " LASTQC = N'" + YKK.ReplaceString(DLASTQC.SelectedValue) + "', "
        sql = sql + " StstusList = N'" + YKK.ReplaceString(DStstusList.SelectedValue) + "', "
        sql = sql + " QCNO = N'" + YKK.ReplaceString(DQCNO.Text) + "', "
        sql = sql + " ACCDEPNAME = N'" + YKK.ReplaceString(DACCDEPNAME.SelectedValue) + "', "
        sql = sql + " ACCEMPNAME = N'" + YKK.ReplaceString(DACCEMPNAME.Text) + "', "
        sql = sql + " REMARK1 = N'" + YKK.ReplaceString(DREMARK1.Text) + "', "
        sql = sql + " FCONFIRM = N'" + YKK.ReplaceString(DFCONFIRM.SelectedValue) + "', "
        sql = sql + " FCONTENT = N'" + YKK.ReplaceString(DFCONTENT.SelectedValue) + "', "
        sql = sql + " FREASON = N'" + YKK.ReplaceString(DFREASON.SelectedValue) + "', "
        sql = sql + " SITUATION = N'" + YKK.ReplaceString(DSITUATION.SelectedValue) + "', "
        sql = sql + " ANSWER = N'" + YKK.ReplaceString(DANSWER.SelectedValue) + "', "
        sql = sql + " ERRORQTY = N'" + YKK.ReplaceString(DERRORQTY.Text) + "', "
        sql = sql + " ERRORP = N'" + YKK.ReplaceString(DERRORP.Text) + "', "
        sql = sql + " SLDDIVISION = N'" + YKK.ReplaceString(DSLDDIVISION.SelectedValue) + "', "
        sql = sql + " UNIT2 = N'" + YKK.ReplaceString(DUNIT2.SelectedValue) + "', "
        sql = sql + " REMARK2 = N'" + YKK.ReplaceString(DREMARK2.Text) + "', "
        sql = sql + " BUQTY = N'" + YKK.ReplaceString(DBUQTY.Text) + "', "
        sql = sql + " QCDATE = N'" + YKK.ReplaceString(DQCDate.Text) + "', "

        If DMach1.Checked = True Then
            sql = sql + " MACH='0',"
            sql = sql + " MACHNO= N'" + YKK.ReplaceString(DMACHNO.Text) + "', "
        ElseIf DMach2.Checked = True Then
            sql = sql + " MACH='1',"
            sql = sql + " MACHNO= '', "
        Else
            sql = sql + " MACH='',"
            sql = sql + " MACHNO= '', "
        End If


        FileName = ""


        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then

                'FileName = CStr(wFormSno) & "-" & "MapFile" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "MapFile" & "-" & UploadDateTime & System.IO.Path.GetExtension(DMapFile.PostedFile.FileName)
                DMapFile.PostedFile.SaveAs(Path + FileName)
                sql = sql + " MapFile= N'" + YKK.ReplaceString(FileName) + "',"           '不良樣品圖
            Else
                FileName = ""
            End If
        End If




        sql = sql + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        sql = sql + " ModifyTime = '" & NowDateTime & "' "
        sql = sql + " Where FormNo  =  '" & wFormNo & "'"
        sql = sql + "   And FormSno =  '" & CStr(wFormSno) & "'"
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
            If DNO.Text <> "" Then
                SQl = "Insert into T_CommissionNo "
                SQl = SQl + "( "
                SQl = SQl + "FormNo, FormSno, No, MapNo, CreateUser, CreateTime "    '1~5
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '" + pFormNo + "', "
                SQl = SQl + " '" + CStr(pFormSno) + "', "
                SQl = SQl + " '" + DNO.Text + "', "
                SQl = SQl + " '" + "" + "', "
                SQl = SQl + " '" + Request.QueryString("pUserID") + "', "
                SQl = SQl + " '" + NowDateTime + "' "
                SQl = SQl + " ) "
                uDataBase.ExecuteNonQuery(SQl)
            End If
        Else
            If DNO.Text <> "" Then
                If DNO.Text <> dtCommissionNo.Rows(0)("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNO.Text & "',"
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
            If UPFile.PostedFile.ContentLength <= 1000 * 1024 Then
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
            If DNO.Text <> "" Then
                ErrCode = oCommon.CommissionNo("003108", wFormSno, wStep, DNO.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9050
                End If
            End If
        End If


        '檢查圖片是否有上傳 
        If ErrCode = 0 Then
            If (wStep = 10 Or wStep = 30) And pAction = "0" Then

                If DMapFile.PostedFile.FileName = "" Then
                    ErrCode = 9010
                End If
            End If
        End If


        '機台號碼
        If ErrCode = 0 Then
            If DMach1.Checked = True And DMach2.Checked Then
                ErrCode = 9084
            ElseIf DMach1.Checked = True Then
                If DMACHNO.Text = "" Then
                    ErrCode = 9083
                End If
            End If

        End If

        If wStep = 20 Then
            If ErrCode = 0 Then
                If DMach1.Checked = False And DMach2.Checked = False Then
                    ErrCode = 9085
                End If
            End If
        End If



        If ErrCode <> 0 Then
            If ErrCode = 9010 Then Message = "異常：需上傳不良圖片!"
            If ErrCode = 9020 Then Message = "上傳檔案格式有誤,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案Size有誤,請確認!"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9050 Then Message = "委託書No.重覆,請確認委託書No.!"
            If ErrCode = 9060 Then Message = "修改後未重新檢測資料,請確認!"
            If ErrCode = 9070 Then Message = "發現有空白重覆,請重新檢測資料,請確認!"
            If ErrCode = 9071 Then Message = "Item Name(2)字串過長(>34),請重新檢測資料,請確認!"
            If ErrCode = 9072 Then Message = "發現特殊要求中有未排序資料,請重新檢測資料,請確認!"
            If ErrCode = 9073 Then Message = "金額需大於0，請確認!"
            If ErrCode = 9075 Then Message = "非數字格式,請確認!"

            If ErrCode = 9083 Then Message = "請輸入該工程POP機台號!"
            If ErrCode = 9084 Then Message = "只能勾選機台或非機台其中一項!"
            If ErrCode = 9085 Then Message = "請勾選機台或非機台其中一項!"

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
    '**     設定附檔路徑 1
    '**
    '*****************************************************************
    Sub NewAttachFilePath()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '3108'"
        SQL = SQL + " and dkey ='AttachfilePath'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If



        If D3.Text = "" Then
            D3.Text = Now.ToString("yyyyMMddHHmmss")
        End If



        OpenDir1 = OpenDir1 + D3.Text + "/投訴部門"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + D3.Text + "\投訴部門"
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If

        Dim dirInfo As New System.IO.DirectoryInfo(tempFolderPath)
        chktemp.Text = tempFolderPath

        '開啟附檔資料夾路徑
        DAttachfile1.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定附檔路徑
    '**
    '*****************************************************************
    Sub NewAttachFilePath2()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '3108'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If




        OpenDir1 = OpenDir1 + DNO.Text + "/品管確認"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNO.Text + "\品管確認"
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

        If FileCount > 0 Or FileDir > 0 Then
            DChkData2.Checked = True
            '   DChkData2.Text = Str(FileCount) + "件"
        Else
            DChkData2.Checked = False
            ' DChkData2.Text = ""
        End If
        '開啟附檔資料夾路徑
        DAttachfile2.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")
    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定附檔路徑
    '**
    '*****************************************************************
    Sub NewAttachFilePath3()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '3108'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If




        OpenDir1 = OpenDir1 + DNO.Text + "/被訴部門"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNO.Text + "\被訴部門"
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

        If FileCount > 0 Or FileDir > 0 Then
            DChkData3.Checked = True
            '   DChkData2.Text = Str(FileCount) + "件"
        Else
            DChkData3.Checked = False
            ' DChkData2.Text = ""
        End If
        '開啟附檔資料夾路徑
        DAttachfile3.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")
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

    '取得關係人
    Function GetRelated(ByVal userId As String) As String

        Dim sql As String = "select RUserID,RRUserID,USERID  from M_Related where userid='" & userId & "' and RelatedID='A'"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            If wStep = 20 Or wStep = 40 Then  '台藉
                If dt.Rows(0)("UserID") = dt.Rows(0)("RUserID") Then '如果被訴擔當跟被訴擔當台籍主管一樣就直接跳日籍主管
                    NextGate = dt.Rows(0)("RRUserID")
                Else
                    NextGate = dt.Rows(0)("RUserID")
                End If

            ElseIf wStep = 30 Or wStep = 50 Then  '日籍
                NextGate = dt.Rows(0)("RRUserID")
            End If

        End If
        Return NextGate
    End Function

    '取得關係人
    Function GetUerID(ByVal Username As String) As String

        Dim sql As String = "select UserID from M_users where username=N'" & Username + "'"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            NextGate = dt.Rows(0)("UserID")
        End If
        Return NextGate
    End Function



    Protected Sub DACCDEPNAME_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DACCDEPNAME.SelectedIndexChanged
        Dim sql As String
        sql = "  SELECT * from M_referp where cat ='3108' "
        sql = sql & " and dkey = '10" + DACCDEPNAME.SelectedValue + "'"
        sql = sql & " order by unique_id"
        Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
        DACCEMPNAME.Text = ""
        Dim i As Integer
        For i = 0 To dtReferp.Rows.Count - 1
            If i = 0 Then
                DACCEMPNAME.Text = dtReferp.Rows(i).Item("Data")
            Else
                DACCEMPNAME.Text = DACCEMPNAME.Text + "," + dtReferp.Rows(i).Item("Data")
            End If

        Next
        dtReferp.Clear()

    End Sub

    Protected Sub DSLDDIVISION_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DSLDDIVISION.SelectedIndexChanged
        If wStep = 10 Then
            Dim sql As String
            sql = "  Select substring(data,CHARINDEX('-', data)+1,len(data)-1)data  from M_referp"
            sql = sql & " where  cat = '3106'"
            sql = sql & " and dkey = 'SLDDIVISION'"
            sql = sql & " and  data like '" + DSLDDIVISION.SelectedValue + "%'"
            sql = sql & " order by unique_id"
            Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
            ' DACCEMPNAME.Text = dtReferp.Rows(0).Item("Data")
        End If

    End Sub



    Protected Sub DERRORQTY_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DERRORQTY.TextChanged
        Dim ERRORP As Decimal
        If DERRORQTY.Text <> "" Then
            ERRORP = CInt(DERRORQTY.Text) / CInt(DORQTY.Text) * 100
            DERRORP.Text = ERRORP.ToString("#0.00")
        End If
    End Sub

    Protected Sub DCOMDEPNAME_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DCOMDEPNAME.SelectedIndexChanged
        Dim sql As String
        sql = "  SELECT * from M_referp where cat ='3108' "
        sql = sql & " and dkey = '10" + DCOMDEPNAME.SelectedValue + "'"
        sql = sql & " order by unique_id"
        Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
        DCOMNAME.Text = ""
        Dim i As Integer
        For i = 0 To dtReferp.Rows.Count - 1
            If i = 0 Then
                DCOMNAME.Text = dtReferp.Rows(i).Item("Data")
            Else
                DCOMNAME.Text = DCOMNAME.Text + "," + dtReferp.Rows(i).Item("Data")
            End If
     
        Next
        dtReferp.Clear()
        Dim dep As String = ""
        dep = Mid(DCOMDEPNAME.SelectedValue, 1, InStr(DCOMDEPNAME.SelectedValue, "-") - 1)

        sql = "  SELECT * from M_referp where cat ='3108' "
        sql = sql & " and dkey = 'DEP-" + dep + "'"
        sql = sql & " order by unique_id"
        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(sql)
        DCOMADMIN.Text = ""

        For i = 0 To dtReferp1.Rows.Count - 1
            If i = 0 Then
                DCOMADMIN.Text = dtReferp1.Rows(i).Item("Data")
            Else
                DCOMADMIN.Text = DCOMADMIN.Text + "," + dtReferp1.Rows(i).Item("Data")
            End If

        Next
        dtReferp.Clear()


    End Sub

End Class
