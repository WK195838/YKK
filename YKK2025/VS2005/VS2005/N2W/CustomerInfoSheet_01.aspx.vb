Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Collections.Generic


Partial Class CustomerInfoSheet_01
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


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'BASDate.Attributes("onclick") = "calendarPicker('DCheckDate');"

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
            Top = 880
            Dim SQL As String = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(SQL)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0)("BEndTime").ToString < NowDateTime Then
                    If DDelay.Visible = True Then
                        Top = 880
                    End If
                End If
            End If
        Else
            Top = 880
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
        wUserIP = Request.ServerVariables("REMOTE_ADDR")
        '
        Response.Cookies("PGM").Value = "CustomerInfoSheet_01.aspx"
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
                Top = 880
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
            Top = 880
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

        ' DAddCH.Text = D1.Text

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
        Select Case FindFieldInf("AppName")
            Case 0  '顯示
                DAppName.BackColor = Color.LightGray
                DAppName.Visible = True
                DAppName.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DAppName.Visible = True
                DAppName.BackColor = Color.GreenYellow
                DAppName.ReadOnly = False
                ShowRequiredFieldValidator("DAppNameRqd", "DAppName", "異常：需輸入姓名")
            Case 2  '修改
                DAppName.Visible = True
                DAppName.BackColor = Color.Yellow
                DAppName.ReadOnly = False
            Case Else   '隱藏
                DAppName.Visible = False
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


        'Select Case FindFieldInf("Division")
        Select Case FindFieldInf("Customer")
            Case 0  '顯示
                DCustomerCode.BackColor = Color.LightGray
                DCustomerCode.Visible = True
                DCustomerCode.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCustomerCode.Visible = True
                DCustomerCode.BackColor = Color.GreenYellow
                DCustomerCode.ReadOnly = False
                ShowRequiredFieldValidator("DCustomerCodeRqd", "DCustomerCode", "異常：需輸入顧客")
            Case 2  '修改
                DCustomerCode.Visible = True
                DCustomerCode.BackColor = Color.Yellow
                DCustomerCode.ReadOnly = False
            Case Else   '隱藏
                DCustomerCode.Visible = False
        End Select



        'Select Case FindFieldInf("Division")
        Select Case FindFieldInf("Relationship")
            Case 0  '顯示
                DRelationCode.BackColor = Color.LightGray
                DRelationCode.Visible = True
                DRelationCode.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRelationCode.Visible = True
                DRelationCode.BackColor = Color.GreenYellow
                DRelationCode.ReadOnly = False
                ShowRequiredFieldValidator("DRelationshipCodeRqd", "DRelationCode", "異常：需輸入關聯代號")
            Case 2  '修改
                DRelationCode.Visible = True
                DRelationCode.BackColor = Color.Yellow
                DRelationCode.ReadOnly = False
            Case Else   '隱藏
                DRelationCode.Visible = False
        End Select




        'Select Case FindFieldInf("Division")
        Select Case FindFieldInf("IDNumber")
            Case 0  '顯示
                DIDNumber.BackColor = Color.LightGray
                DIDNumber.Visible = True
                DIDNumber.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DIDNumber.Visible = True
                DIDNumber.BackColor = Color.GreenYellow
                DIDNumber.ReadOnly = False
                ShowRequiredFieldValidator("DIDNumberRqd", "DIDNumber", "異常：需輸入統一編號")
            Case 2  '修改
                DIDNumber.Visible = True
                DIDNumber.BackColor = Color.Yellow
                DIDNumber.ReadOnly = False
            Case Else   '隱藏
                DIDNumber.Visible = False
        End Select


        'Select Case FindFieldInf("Division")
        Select Case FindFieldInf("TEL")
            Case 0  '顯示
                DTEL1.BackColor = Color.LightGray
                DTEL1.Visible = True
                DTEL1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTEL1.Visible = True
                DTEL1.BackColor = Color.GreenYellow
                DTEL1.ReadOnly = False
                ShowRequiredFieldValidator("DTEL1Rqd", "DTEL1", "異常：需輸入電話區碼")
            Case 2  '修改
                DTEL1.Visible = True
                DTEL1.BackColor = Color.Yellow
                DTEL1.ReadOnly = False
            Case Else   '隱藏
                DTEL1.Visible = False
        End Select



        'Select Case FindFieldInf("Division")
        Select Case FindFieldInf("FAX")
            Case 0  '顯示
                DFAX1.BackColor = Color.LightGray
                DFAX1.Visible = True
                DFAX1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DFAX1.Visible = True
                DFAX1.BackColor = Color.GreenYellow
                DFAX1.ReadOnly = False
                ShowRequiredFieldValidator("DFAX1Rqd", "DFAX1", "異常：需輸入FAX區碼")
            Case 2  '修改
                DFAX1.Visible = True
                DFAX1.BackColor = Color.Yellow
                DFAX1.ReadOnly = False
            Case Else   '隱藏
                DFAX1.Visible = False
        End Select





        'Sales
        Select Case FindFieldInf("Sales")
            Case 0  '顯示
                DSales.BackColor = Color.LightGray
                DSales.Visible = True

            Case 1  '修改+檢查
                DSales.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSalesRqd", "DSales", "異常：需輸入業務員")
                DSales.Visible = True
            Case 2  '修改
                DSales.BackColor = Color.Yellow
                DSales.Visible = True
            Case Else   '隱藏
                DSales.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Sales", "ZZZZZZ")





        'Goods
        Select Case FindFieldInf("Goods")
            Case 0  '顯示
                DGoods.BackColor = Color.LightGray
                DGoods.Visible = True

            Case 1  '修改+檢查
                DGoods.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DGoodsRqd", "DGoods", "異常：需輸入業別")
                DGoods.Visible = True
            Case 2  '修改
                DGoods.BackColor = Color.Yellow
                DGoods.Visible = True
            Case Else   '隱藏
                DGoods.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Goods", "ZZZZZZ")





        'Select Case FindFieldInf("Division")
        Select Case FindFieldInf("Customs")
            Case 0  '顯示
                DCustoms.BackColor = Color.LightGray
                DCustoms.Visible = True
                DCustoms.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCustoms.Visible = True
                DCustoms.BackColor = Color.GreenYellow
                DCustoms.ReadOnly = False
                ShowRequiredFieldValidator("DCustomsRqd", "DCustoms", "異常：需輸入海關企編")
            Case 2  '修改
                DCustoms.Visible = True
                DCustoms.BackColor = Color.Yellow
                DCustoms.ReadOnly = False
            Case Else   '隱藏
                DCustoms.Visible = False
        End Select






        'Location
        Select Case FindFieldInf("Location")
            Case 0  '顯示
                DLocation.BackColor = Color.LightGray
                DLocation.Visible = True

            Case 1  '修改+檢查
                DLocation.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DLocationRqd", "DLocation", "異常：需輸入業別")
                DLocation.Visible = True
            Case 2  '修改
                DLocation.BackColor = Color.Yellow
                DLocation.Visible = True
            Case Else   '隱藏
                DLocation.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Location", "ZZZZZZ")





        'Delivery
        Select Case FindFieldInf("Delivery")
            Case 0  '顯示
                DDelivery.BackColor = Color.LightGray
                DDelivery.Visible = True

            Case 1  '修改+檢查
                DDelivery.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDeliveryRqd", "DDelivery", "異常：需輸入送貨代號")
                DDelivery.Visible = True
            Case 2  '修改
                DDelivery.BackColor = Color.Yellow
                DDelivery.Visible = True
            Case Else   '隱藏
                DDelivery.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Delivery", "ZZZZZZ")




        'Select Case FindFieldInf("NameCH")
        Select Case FindFieldInf("NameCH")
            Case 0  '顯示
                DNameCH.BackColor = Color.LightGray
                DNameCH.Visible = True
                DNameCH.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DNameCH.Visible = True
                DNameCH.BackColor = Color.GreenYellow
                DNameCH.ReadOnly = False
                ShowRequiredFieldValidator("DNameCHRqd", "DNameCH", "異常：需輸入中文名稱")
            Case 2  '修改
                DNameCH.Visible = True
                DNameCH.BackColor = Color.Yellow
                DNameCH.ReadOnly = False
            Case Else   '隱藏
                DNameCH.Visible = False
        End Select

        'Select Case FindFieldInf("NameEN")
        Select Case FindFieldInf("NameEN")
            Case 0  '顯示
                DNameEN.BackColor = Color.LightGray
                DNameEN.Visible = True
                DNameEN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DNameEN.Visible = True
                DNameEN.BackColor = Color.GreenYellow
                DNameEN.ReadOnly = False
                ShowRequiredFieldValidator("DNameENRqd", "DNameEN", "異常：需輸入英文名稱")
            Case 2  '修改
                DNameEN.Visible = True
                DNameEN.BackColor = Color.Yellow
                DNameEN.ReadOnly = False
            Case Else   '隱藏
                DNameEN.Visible = False
        End Select



        'Select Case FindFieldInf("InvoiceCH")
        Select Case FindFieldInf("InvoiceCH")
            Case 0  '顯示
                DInvoiceCH.BackColor = Color.LightGray
                DInvoiceCH.Visible = True
                DInvoiceCH.Attributes.Add("readonly", "true")
                BAddress1.Visible = False
            Case 1  '修改+檢查
                DInvoiceCH.Visible = True
                DInvoiceCH.BackColor = Color.GreenYellow
                DInvoiceCH.ReadOnly = False
                ShowRequiredFieldValidator("DInvoiceCHRqd", "DInvoiceCH", "異常：需輸入中文發票地址")
                BAddress1.Visible = True
            Case 2  '修改
                DInvoiceCH.Visible = True
                DInvoiceCH.BackColor = Color.Yellow
                DInvoiceCH.ReadOnly = False
                BAddress1.Visible = True
            Case Else   '隱藏
                DInvoiceCH.Visible = False
                BAddress1.Visible = False
        End Select




        'Select Case FindFieldInf("InvoiceEN")
        Select Case FindFieldInf("InvoiceEN")
            Case 0  '顯示
                DInvoiceEN.BackColor = Color.LightGray
                DInvoiceEN.Visible = True
                DInvoiceEN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DInvoiceEN.Visible = True
                DInvoiceEN.BackColor = Color.GreenYellow
                DInvoiceEN.ReadOnly = False
                ShowRequiredFieldValidator("DInvoiceENRqd", "DInvoiceEN", "異常：需輸入英文發票地址")
            Case 2  '修改
                DInvoiceEN.Visible = True
                DInvoiceEN.BackColor = Color.Yellow
                DInvoiceEN.ReadOnly = False
            Case Else   '隱藏
                DInvoiceEN.Visible = False
        End Select




        'Select Case FindFieldInf("PostCode")
        Select Case FindFieldInf("PostCode")
            Case 0  '顯示
                DPostCode.BackColor = Color.LightGray
                DPostCode.Visible = True
                DPostCode.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DPostCode.Visible = True
                DPostCode.BackColor = Color.GreenYellow
                DPostCode.ReadOnly = False
                ShowRequiredFieldValidator("DPostCodeRqd", "DPostCode", "異常：需輸入郵遞區號")
            Case 2  '修改
                DPostCode.Visible = True
                DPostCode.BackColor = Color.Yellow
                DPostCode.ReadOnly = False
            Case Else   '隱藏
                DPostCode.Visible = False
        End Select




        'Select Case FindFieldInf("PostCode")
        Select Case FindFieldInf("TJCode")
            Case 0  '顯示
                DTJCode.BackColor = Color.LightGray
                DTJCode.Visible = True
                DTJCode.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DTJCode.Visible = True
                DTJCode.BackColor = Color.GreenYellow
                DTJCode.ReadOnly = False
                ShowRequiredFieldValidator("DTJCodeRqd", "DTJCode", "異常：需輸入大榮簡碼")
            Case 2  '修改
                DTJCode.Visible = True
                DTJCode.BackColor = Color.Yellow
                DTJCode.ReadOnly = False
            Case Else   '隱藏
                DTJCode.Visible = False
        End Select




        'Select Case FindFieldInf("AddCH")
        Select Case FindFieldInf("AddCH")
            Case 0  '顯示
                DAddCH.BackColor = Color.LightGray
                DAddCH.Visible = True
                DAddCH.Attributes.Add("readonly", "true")
                BAddress2.Visible = False

            Case 1  '修改+檢查
                DAddCH.Visible = True
                DAddCH.BackColor = Color.GreenYellow
                DAddCH.ReadOnly = False
                ShowRequiredFieldValidator("DAddCHRqd", "DAddCH", "異常：需輸入中文送貨地址")
                BAddress2.Visible = True
            Case 2  '修改
                DAddCH.Visible = True
                DAddCH.BackColor = Color.Yellow
                DAddCH.ReadOnly = False
                BAddress2.Visible = True
            Case Else   '隱藏
                DAddCH.Visible = False
                BAddress2.Visible = False
        End Select





        'Select Case FindFieldInf("AddEN")
        Select Case FindFieldInf("AddEN")
            Case 0  '顯示
                DAddEN.BackColor = Color.LightGray
                DAddEN.Visible = True
                DAddEN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DAddEN.Visible = True
                DAddEN.BackColor = Color.GreenYellow
                DAddEN.ReadOnly = False
                ShowRequiredFieldValidator("DAddENRqd", "DAddEN", "異常：需輸入英文送貨地址")
            Case 2  '修改
                DAddEN.Visible = True
                DAddEN.BackColor = Color.Yellow
                DAddEN.ReadOnly = False
            Case Else   '隱藏
                DAddEN.Visible = False
        End Select



        '
        'Select Case FindFieldInf("Division")
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


        'Paytype
        Select Case FindFieldInf("Paytype")
            Case 0  '顯示
                DPaytype.BackColor = Color.LightGray
                DPaytype.Visible = True

            Case 1  '修改+檢查
                DPaytype.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPaytypeRqd", "DPaytype", "異常：需輸入付款代號")
                DPaytype.Visible = True
            Case 2  '修改
                DPaytype.BackColor = Color.Yellow
                DPaytype.Visible = True
            Case Else   '隱藏
                DPaytype.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Paytype", "ZZZZZZ")


        'Select Case FindFieldInf("PaytypeOld")
        Select Case FindFieldInf("PaytypeOld")
            Case 0  '顯示
                DPaytypeOld.BackColor = Color.LightGray
                DPaytypeOld.Visible = True
                DPaytypeOld.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DPaytypeOld.Visible = True
                DPaytypeOld.BackColor = Color.GreenYellow
                DPaytypeOld.ReadOnly = False
                ShowRequiredFieldValidator("DPaytypeOldRqd", "DPaytypeOld", "異常：需輸入討款代號")
            Case 2  '修改
                DPaytypeOld.Visible = True
                DPaytypeOld.BackColor = Color.Yellow
                DPayTypeOld.ReadOnly = False
            Case Else   '隱藏
                DPayTypeOld.Visible = False
        End Select



        '附檔1
        Select Case FindFieldInf("AttachFile")
            Case 0  '顯示
                DAttachfile.Visible = False
                DAttachfile.Style.Add("BACKGROUND-COLOR", "LightGrey")


            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DAttachFileRqd", "DAttachFile", "異常：需附檔")
                DAttachfile.Visible = True
                DAttachfile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                If LAttachfile.NavigateUrl = "" Then
                    LAttachfile.Visible = False
                End If
            Case 2  '修改
                DAttachfile.Visible = True
                DAttachfile.Style.Add("BACKGROUND-COLOR", "Yellow")
                If LAttachfile.NavigateUrl = "" Then
                    LAttachfile.Visible = False
                End If
            Case Else   '隱藏
                DAttachfile.Visible = False
        End Select


    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()

        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("http") & _
                             System.Configuration.ConfigurationManager.AppSettings("CustomerInfoPath")  'WIS-TempPath

        Dim SQL As String
        SQL = "Select * From F_CustomerInfoSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtCust As DataTable = uDataBase.GetDataTable(SQL)
        If dtCust.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dtCust.Rows(0).Item("No")                         'No
            '   DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
            DDate.Text = dtCust.Rows(0).Item("Date")                         'No 
            DDepName.Text = dtCust.Rows(0).Item("DepName")                         'No
            DAppName.Text = dtCust.Rows(0).Item("AppName")                         'No
            DCustomerCode.Text = dtCust.Rows(0).Item("CustomerCode")                         'No
            DRelationCode.Text = dtCust.Rows(0).Item("RelationCode")

            SetFieldData("Location", dtCust.Rows(0).Item("Location"))    ' 

            DIDNumber.Text = dtCust.Rows(0).Item("IDNumber")                         'No
            DTEL1.Text = dtCust.Rows(0).Item("Tel1")                         'No

            DFAX1.Text = dtCust.Rows(0).Item("Fax1")                         'No


            SetFieldData("Sales", dtCust.Rows(0).Item("Sales"))    ' 
            SetFieldData("Goods", dtCust.Rows(0).Item("Goods"))    ' 
            SetFieldData("Delivery", dtCust.Rows(0).Item("Delivery"))    ' 

            DCustoms.Text = dtCust.Rows(0).Item("Customs")                         'No
            DNameCH.Text = dtCust.Rows(0).Item("NameCH")                         'No
            DNameEN.Text = dtCust.Rows(0).Item("NameEN")
            DInvoiceCH.Text = dtCust.Rows(0).Item("InvoiceCH")
            DInvoiceEN.Text = dtCust.Rows(0).Item("InvoiceEN")
            DPostCode.Text = dtCust.Rows(0).Item("PostCode")
            DTJCode.Text = dtCust.Rows(0).Item("TJCode")
            DAddCH.Text = dtCust.Rows(0).Item("AddCH")                         'No
            DAddEN.Text = dtCust.Rows(0).Item("AddEN")
            DRemark.Text = dtCust.Rows(0).Item("Remark")

            SetFieldData("Paytype", dtCust.Rows(0).Item("PayType"))    ' 
            DPayTypeOld.Text = dtCust.Rows(0).Item("Paytypeold")

            If dtCust.Rows(0).Item("AttachFile") <> "" Then
                LAttachfile.NavigateUrl = Path & dtCust.Rows(0).Item("AttachFile") '折扣
                LAttachfile.Visible = True
            Else
                LAttachfile.Visible = False
            End If


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
        Dim i As Integer
        '擔當者及部門 

        sql = "Select Divname,Username From M_Users "
        sql = sql & " Where UserID = '" & wApplyID & "'"
        sql = sql & "   And Active = '1' "
        Dim DBUser As DataTable = uDataBase.GetDataTable(sql)


        DDepName.Text = DBUser.Rows(0).Item("Divname")
        DAppName.Text = DBUser.Rows(0).Item("Username")
        DLocation.SelectedValue = DBUser.Rows(0).Item("Username")


        'Goods
        If pFieldName = "Goods" Then
            DGoods.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DGoods.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3103'"
                sql = sql & " and dkey = 'Goods'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)

                DGoods.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DGoods.Items.Add(ListItem1)



                Next
                dtReferp.Clear()
            End If
        End If


        'Sales
        If pFieldName = "Sales" Then
            DSales.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSales.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3103'"
                sql = sql & " and dkey = 'Sales'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DSales.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSales.Items.Add(ListItem1)



                Next
                dtReferp.Clear()
            End If
        End If



        'Location
        If pFieldName = "Location" Then
            DLocation.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLocation.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3103'"
                sql = sql & " and dkey = 'Location'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLocation.Items.Add(ListItem1)



                Next
                dtReferp.Clear()
            End If
        End If

        'Delivery
        If pFieldName = "Delivery" Then
            DDelivery.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDelivery.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3103'"
                sql = sql & " and dkey = 'Delivery'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDelivery.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If




        'Paytype
        If pFieldName = "Paytype" Then
            DPaytype.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPaytype.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3103'"
                sql = sql & " and dkey = 'Paytype'"
                sql = sql & " order by data "

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DPaytype.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPaytype.Items.Add(ListItem1)
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
        rqdVal.Style.Add("Top", Top - 50 & "px")
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

        'BASDate.Attributes("onclick") = "calendarPicker('DCheckDate');"


        BAddress1.Attributes.Add("onClick", "GetAddres('DInvoiceCH')") '找地址
        BAddress2.Attributes.Add("onClick", "GetAddres('DAddCH')") '找地址

        'BAddress2.Attributes.Add("onClick", "GetAddres(2)") '找地址
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
        Dim wTableName As String = "F_CustomerInfoSheet"
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
        Dim FileName As String
        Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                              CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                    System.Configuration.ConfigurationManager.AppSettings("CustomerInfoPath")

        Dim sql As String = ""
        sql = " Insert into F_CustomerInfoSheet (Sts, CompletedTime, FormNo, FormSno,"
        sql &= " NO,DepName,Date,AppName,CustomerCode,RelationCode,Location,IDNumber,TEL1,FAX1,Sales,Goods,Delivery,Customs,NameCH,NameEN,InvoiceCH,InvoiceEN,PostCode,TJCode,AddCH,AddEN,Remark,Paytype,PayTypeOld,Attachfile"
        sql = sql + ",CreateUser, CreateTime, ModifyUser, ModifyTime) "
        sql = sql + "VALUES( "
        sql = sql + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        sql = sql + " '" + NowDateTime + "', "        '結案日
        sql = sql + " '003103', "                     '表單代號
        sql = sql + " '" + CStr(NewFormSno) + "', "   '表單流水號
        sql = sql + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO  1
        sql = sql + " '" + DDepName.Text + "', "                '部門2
        sql = sql + " '" + DDate.Text + "', "                '日期3
        sql = sql + " '" + DAppName.Text + "', "                '姓名4

        sql = sql + " '" + DCustomerCode.Text + "', "                '姓名4
        sql = sql + " '" + DRelationCode.Text + "', "                '姓名4
        sql = sql + " '" + DLocation.SelectedValue + "', "                '姓名4
        sql = sql + " '" + DIDNumber.Text + "', "                '姓名4
        sql = sql + " '" + DTEL1.Text + "', "                '姓名4
        sql = sql + " '" + DFAX1.Text + "', "                '姓名4
        sql = sql + " '" + DSales.Text + "', "                '姓名4
        sql = sql + " '" + DGoods.Text + "', "                '姓名4
        sql = sql + " '" + DDelivery.SelectedValue + "', "                '姓名4
        sql = sql + " '" + DCustoms.Text + "', "                '姓名4
        sql = sql + " '" + uCommon.ReplaceString(DNameCH.Text) + "', "                '姓名4
        sql = sql + " '" + uCommon.ReplaceString(DNameEN.Text) + "', "                '姓名4
        sql = sql + " '" + uCommon.ReplaceString(DInvoiceCH.Text) + "', "                '姓名4
        sql = sql + " '" + uCommon.ReplaceString(DInvoiceEN.Text) + "', "                '姓名4
        sql = sql + " '" + DPostCode.Text + "', "                '姓名4
        sql = sql + " '" + DTJCode.Text + "', "
        sql = sql + " '" + uCommon.ReplaceString(DAddCH.Text) + "', "                '姓名4
        sql = sql + " '" + uCommon.ReplaceString(DAddEN.Text) + "', "                '姓名4
        sql = sql + " '" + uCommon.ReplaceString(DRemark.Text) + "', "                '備註
        sql = sql + " '" + uCommon.ReplaceString(DPaytype.SelectedValue) + "', "                '備註
        sql = sql + " '" + uCommon.ReplaceString(DPayTypeOld.Text) + "', "                '³Æ

        FileName = ""  '附檔1
        If DAttachfile.PostedFile.FileName <> "" Then

           
            '20170912 將檔名修改成不含原始檔名
            Dim fileExtension As String  '副檔名
            fileExtension = IO.Path.GetExtension(DAttachfile.PostedFile.FileName).ToLower   '取得檔案格式

            FileName = CStr(NewFormSno) + "-Attachfile-" + UploadDateTime + fileExtension

            DAttachfile.PostedFile.SaveAs(Path + FileName)

        Else
            FileName = ""
        End If

        sql = sql + "'" + FileName + "'," ' 附檔1


        sql = sql + "'" & Request.QueryString("pUserID") & "' ,"
        sql = sql + "'" & NowDateTime & "' ,"
        sql = sql + "'" & Request.QueryString("pUserID") & "' ,"
        sql = sql + "'" & NowDateTime & "' )"
        uDataBase.ExecuteNonQuery(sql)
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
                    System.Configuration.ConfigurationManager.AppSettings("CustomerInfoPath")
        Dim sql As String
        Dim FileName As String

        sql = " Update F_CustomerInfoSheet"
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
        sql = sql + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"
        sql = sql + " Date = N'" & DDate.Text & "',"
        sql = sql + " DepName = N'" & DDepName.Text & "',"
        sql = sql + " AppName = N'" & DAppName.Text & "',"

        sql = sql + " CustomerCode = N'" & DCustomerCode.Text & "',"
        sql = sql + " RelationCode = N'" & DRelationCode.Text & "',"
        sql = sql + " Location = N'" & DLocation.SelectedValue & "',"
        sql = sql + " IDNumber = N'" & DIDNumber.Text & "',"
        sql = sql + " TEL1 = N'" & DTEL1.Text & "',"
        sql = sql + " FAX1 = N'" & DFAX1.Text & "',"
        sql = sql + " Sales = N'" & DSales.SelectedValue & "',"
        sql = sql + " Goods = N'" & DGoods.SelectedValue & "',"
        sql = sql + " Delivery = N'" & DDelivery.SelectedValue & "',"
        sql = sql + " Customs = N'" & DCustoms.Text & "',"
        sql = sql + " NameCH = N'" & uCommon.ReplaceString(DNameCH.Text) & "',"
        sql = sql + " NameEN= N'" & uCommon.ReplaceString(DNameEN.Text) & "',"
        sql = sql + " InvoiceCH = N'" & uCommon.ReplaceString(DInvoiceCH.Text) & "',"
        sql = sql + " InvoiceEN = N'" & uCommon.ReplaceString(DInvoiceEN.Text) & "',"
        sql = sql + " AddCH = N'" & uCommon.ReplaceString(DAddCH.Text) & "',"
        sql = sql + " AddEN = N'" & uCommon.ReplaceString(DAddEN.Text) & "',"
        sql = sql + " PostCode = N'" & DPostCode.Text & "',"
        sql = sql + " TJCode = N'" & DTJCode.Text & "',"
        sql = sql + " Remark= N'" & uCommon.ReplaceString(DRemark.Text) & "',"        '
        sql = sql + " Paytype= N'" & uCommon.ReplaceString(DPaytype.Text) & "',"        '
        FileName = ""
        If DAttachfile.Visible = True Then
            '   And LCertifcateFile.NavigateUrl = "" Then
            If DAttachfile.PostedFile.FileName <> "" And LAttachfile.NavigateUrl <> "" Then

                '20170912 將檔名修改成不含原始檔名
                Dim fileExtension As String  '副檔名
                fileExtension = IO.Path.GetExtension(DAttachfile.PostedFile.FileName).ToLower   '取得檔案格式

                FileName = CStr(wFormSno) + "-Attachfile1-" + UploadDateTime + fileExtension

                If DAttachfile.PostedFile.FileName = "" Then
                    '  FileName = Right(LDISPOSALFILE.NavigateUrl, InStr(StrReverse(LDISPOSALFILE.NavigateUrl), "\") - 1)
                    DAttachfile.PostedFile.SaveAs(Path + FileName)
                Else
                    ' FileName = CStr(wFormSno) + "-DPFILE-" + UploadDateTime & "-" & Right(DAttachFile1.PostedFile.FileName, InStr(StrReverse(DAttachFile1.PostedFile.FileName), "\") - 1)
                    ' FileName = CStr(wFormSno) + "-DPFILE-" + UploadDateTime + fileExtension
                    DAttachfile.PostedFile.SaveAs(Path + FileName)
                End If

            Else
                FileName = ""
            End If
            If FileName <> "" Then
                sql = sql + " Attachfile='" + FileName + "'," '附檔1
            End If

        End If



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
            '##-5
            '##AgentApprov-Start
            If oCommon.UseAgentApprov(wFormNo, wStep, "OK") = 0 Then
                AgentApprovProc("OK", 0, "1")
            Else
                FlowControl("OK", 0, "1")
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

        If ErrCode = 0 Then  ' 判斷檔案上傳
            If DAttachfile.Visible Then
                If DAttachfile.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DAttachfile)

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
                ErrCode = oCommon.CommissionNo("003103", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9050
                End If
            End If
        End If

        Try
          
        Catch ex As Exception
            ErrCode = 9075
        End Try

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
            If ErrCode = 9073 Then Message = "溫度不正常，請確認!"
            If ErrCode = 9074 Then Message = "溼度不正常，請確認!"
            If ErrCode = 9075 Then Message = "非數字格式,請確認!"
            uJavaScript.PopMsg(Me, Message)
        Else
            isOK = True
        End If
        '
        Return isOK
    End Function


End Class
