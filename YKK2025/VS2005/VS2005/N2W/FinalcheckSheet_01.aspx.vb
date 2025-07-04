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


Partial Class FinalcheckSheet_01
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
            Top = 1100
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
            Top = 1100

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


        Top = 1100
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
            Top = 1110
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
        Response.Cookies("PGM").Value = "FinalcheckSheet_01.aspx"
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
            Top = 910
        Else
            Top = 1100
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
        Select Case FindFieldInf("DATE")
            Case 0  '顯示
                DDATE.BackColor = Color.LightGray
                DDATE.Visible = True
                DDATE.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DDATE.Visible = True
                DDATE.BackColor = Color.GreenYellow
                DDATE.ReadOnly = False
                ShowRequiredFieldValidator("DDateRqd", "DDate", "異常：需輸入申請日期")
            Case 2  '修改
                DDATE.Visible = True
                DDATE.BackColor = Color.Yellow
                DDATE.ReadOnly = False
            Case Else   '隱藏
                DDATE.Visible = False
        End Select
        If pPost = "New" Then DDATE.Text = Now.ToString("yyyy/MM/dd") '現在日時





        'CUSTOMER
        Select Case FindFieldInf("CUSTOMER")
            Case 0  '顯示
                DCUSTOMER.BackColor = Color.White
                DCUSTOMER.Visible = True
                DCUSTOMER.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCUSTOMER.Visible = True
                DCUSTOMER.BackColor = Color.GreenYellow
                DCUSTOMER.ReadOnly = False
                ShowRequiredFieldValidator("DCUSTOMERRqd", "DCUSTOMER", "異常：需輸入顧客名稱")
            Case 2  '修改
                DCUSTOMER.Visible = True
                DCUSTOMER.BackColor = Color.Yellow
                DCUSTOMER.ReadOnly = False
            Case Else   '隱藏
                DCUSTOMER.Visible = False
        End Select



        '被訴部門
        Select Case FindFieldInf("ACCDEPNAME")
            Case 0  '顯示
                DACCDEPNAME.BackColor = Color.LightGray
                DACCDEPNAME.Visible = True

            Case 1  '修改+檢查
                DACCDEPNAME.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DACCDEPNAMERqd", "DACCDEPNAME", "異常：需輸入被訴部門")
                DACCDEPNAME.Visible = True
            Case 2  '修改
                DACCDEPNAME.BackColor = Color.Yellow
                DACCDEPNAME.Visible = True
            Case Else   '隱藏
                DACCDEPNAME.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ACCDEPNAME", "ZZZZZZ")




        ''SLD工程別             
        'Select Case FindFieldInf("SLDDIVISION")
        '    Case 0  '顯示
        '        DSLDDIVISION.BackColor = Color.LightGray
        '        DSLDDIVISION.Visible = True

        '    Case 1  '修改+檢查
        '        DSLDDIVISION.BackColor = Color.GreenYellow
        '        ShowRequiredFieldValidator("DSLDDIVISIONRqd", "DSLDDIVISION", "異常：需輸入SLD工程別")
        '        DSLDDIVISION.Visible = True
        '    Case 2  '修改
        '        DSLDDIVISION.BackColor = Color.Yellow
        '        DSLDDIVISION.Visible = True
        '    Case Else   '隱藏
        '        DSLDDIVISION.Visible = False
        'End Select
        'If pPost = "New" Then SetFieldData("SLDDIVISION", "ZZZZZZ")


        '相關部門
        Select Case FindFieldInf("RDIVISION")
            Case 0  '顯示
                DRDIVISION.BackColor = Color.LightGray
                DRDIVISION.Visible = True

            Case 1  '修改+檢查
                DRDIVISION.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DRDIVISIONRqd", "DRDIVISION", "異常：需輸入相關部門")
                DRDIVISION.Visible = True
            Case 2  '修改
                DRDIVISION.BackColor = Color.Yellow
                DRDIVISION.Visible = True
            Case Else   '隱藏
                DRDIVISION.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("RDIVISION", "ZZZZZZ")



        'type
        Select Case FindFieldInf("TYPE")
            Case 0  '顯示
                DTYPE.BackColor = Color.LightGray
                DTYPE.Visible = True
                DTYPE.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTYPE.Visible = True
                DTYPE.BackColor = Color.GreenYellow
                DTYPE.ReadOnly = False
                ShowRequiredFieldValidator("DTYPERqd", "DTYPE", "異常：需輸入類別（型別）")
            Case 2  '修改
                DTYPE.Visible = True
                DTYPE.BackColor = Color.Yellow
                DTYPE.ReadOnly = False
            Case Else   '隱藏
                DTYPE.Visible = False
        End Select


        '檢查日期
        Select Case FindFieldInf("CHECKDATE")
            Case 0  '顯示
                DCHECKDATE.Visible = True
                DCHECKDATE.Style.Add("background-color", "lightgrey")
                DCHECKDATE.Attributes.Add("readonly", "true")
                BCHECKDATE.Visible = False

            Case 1  '修改+檢查
                DCHECKDATE.Visible = True
                DCHECKDATE.Style.Add("background-color", "greenyellow")
                DCHECKDATE.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DCHECKDATERqd", "DCHECKDATE", "異常：需輸入檢查日期")
                BCHECKDATE.Visible = True

            Case 2  '修改
                DCHECKDATE.Visible = True
                DCHECKDATE.Style.Add("background-color", "yellow")
                DCHECKDATE.Attributes.Add("readonly", "true")
                BCHECKDATE.Visible = True
            Case Else   '隱藏
                DCHECKDATE.Visible = False
                BCHECKDATE.Visible = False

        End Select
        If pPost = "New" Then DCHECKDATE.Text = Now.ToString("yyyy/MM/dd") '現在日時


        '訂單號碼
        Select Case FindFieldInf("ORNO")
            Case 0  '顯示
                DORNO.BackColor = Color.LightGray
                DORNO.Visible = True
                DORNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DORNO.Visible = True
                DORNO.BackColor = Color.GreenYellow
                DORNO.ReadOnly = False
                ShowRequiredFieldValidator("DORNORqd", "DORNO", "異常：需輸入訂單號碼")
            Case 2  '修改
                DORNO.Visible = True
                DORNO.BackColor = Color.Yellow
                DORNO.ReadOnly = False
            Case Else   '隱藏
                DORNO.Visible = False
        End Select




        '色番項目
        Select Case FindFieldInf("COLORITEM")
            Case 0  '顯示
                DCOLORITEM.BackColor = Color.LightGray
                DCOLORITEM.Visible = True
                DCOLORITEM.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCOLORITEM.Visible = True
                DCOLORITEM.BackColor = Color.GreenYellow
                DCOLORITEM.ReadOnly = False
                ShowRequiredFieldValidator("DCOLORITEMRqd", "DCOLORITEM", "異常：需輸入色番項目")
            Case 2  '修改
                DCOLORITEM.Visible = True
                DCOLORITEM.BackColor = Color.Yellow
                DCOLORITEM.ReadOnly = False
            Case Else   '隱藏
                DCOLORITEM.Visible = False
        End Select



        '數量
        Select Case FindFieldInf("QTY")
            Case 0  '顯示
                DQTY.BackColor = Color.LightGray
                DQTY.Visible = True
                DQTY.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DQTY.Visible = True
                DQTY.BackColor = Color.GreenYellow
                DQTY.ReadOnly = False
                ShowRequiredFieldValidator("DQTYRqd", "DQTY", "異常：需輸入數量")
            Case 2  '修改
                DQTY.Visible = True
                DQTY.BackColor = Color.Yellow
                DQTY.ReadOnly = False
            Case Else   '隱藏
                DQTY.Visible = False
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



        '抽查量
        Select Case FindFieldInf("CHECKQTY")
            Case 0  '顯示
                DCHECKQTY.BackColor = Color.LightGray
                DCHECKQTY.Visible = True
                DCHECKQTY.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCHECKQTY.Visible = True
                DCHECKQTY.BackColor = Color.GreenYellow
                DCHECKQTY.ReadOnly = False
                ShowRequiredFieldValidator("DCHECKQTYRqd", "DCHECKQTY", "異常：需輸入抽查量")
            Case 2  '修改
                DCHECKQTY.Visible = True
                DCHECKQTY.BackColor = Color.Yellow
                DCHECKQTY.ReadOnly = False
            Case Else   '隱藏
                DCHECKQTY.Visible = False
        End Select



        '抽查量單位
        Select Case FindFieldInf("UNIT2")
            Case 0  '顯示
                DUNIT2.BackColor = Color.LightGray
                DUNIT2.Visible = True

            Case 1  '修改+檢查
                DUNIT2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DUNIT2Rqd", "DUNIT2", "異常：需輸入數量單位")
                DUNIT2.Visible = True
            Case 2  '修改
                DUNIT2.BackColor = Color.Yellow
                DUNIT2.Visible = True
            Case Else   '隱藏
                DUNIT2.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("UNIT2", "ZZZZZZ")



        '不良量
        Select Case FindFieldInf("ERRORQTY")
            Case 0  '顯示
                DERRORQTY.BackColor = Color.LightGray
                DERRORQTY.Visible = True
                DERRORQTY.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DERRORQTY.Visible = True
                DERRORQTY.BackColor = Color.GreenYellow
                DERRORQTY.ReadOnly = False
                ShowRequiredFieldValidator("DERRORQTYRqd", "DERRORQTY", "異常：需輸入不良量")
            Case 2  '修改
                DERRORQTY.Visible = True
                DERRORQTY.BackColor = Color.Yellow
                DERRORQTY.ReadOnly = False
            Case Else   '隱藏
                DERRORQTY.Visible = False
        End Select



        '抽查量單位
        Select Case FindFieldInf("UNIT3")
            Case 0  '顯示
                DUNIT3.BackColor = Color.LightGray
                DUNIT3.Visible = True

            Case 1  '修改+檢查
                DUNIT3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DUNIT3Rqd", "DUNIT3", "異常：需輸入數量單位")
                DUNIT3.Visible = True
            Case 2  '修改
                DUNIT3.BackColor = Color.Yellow
                DUNIT3.Visible = True
            Case Else   '隱藏
                DUNIT3.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("UNIT3", "ZZZZZZ")


        '品管擔當
        Select Case FindFieldInf("APPNAME")
            Case 0  '顯示
                DAPPNAME.BackColor = Color.LightGray
                DAPPNAME.Visible = True
                DAPPNAME.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DAPPNAME.Visible = True
                DAPPNAME.BackColor = Color.GreenYellow
                DAPPNAME.ReadOnly = False
                ShowRequiredFieldValidator("DAPPNAMERqd", "DAPPNAME", "異常：需輸入品管擔當")
            Case 2  '修改
                DAPPNAME.Visible = True
                DAPPNAME.BackColor = Color.Yellow
                DAPPNAME.ReadOnly = False
            Case Else   '隱藏
                DAPPNAME.Visible = False
        End Select




        'SLD不良項目
        Select Case FindFieldInf("ERROR1")
            Case 0  '顯示
                DERROR1.BackColor = Color.LightGray
                DERROR1.Visible = True
                DERROR1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DERROR1.Visible = True
                DERROR1.BackColor = Color.GreenYellow
                DERROR1.ReadOnly = False
                ShowRequiredFieldValidator("DERROR1Rqd", "DERROR1", "異常：需輸入SLD不良項目")
            Case 2  '修改
                DERROR1.Visible = True
                DERROR1.BackColor = Color.Yellow
                DERROR1.ReadOnly = False
            Case Else   '隱藏
                DERROR1.Visible = False
        End Select


        'PF/VF/MF 不良項目
        Select Case FindFieldInf("ERROR2")
            Case 0  '顯示
                DERROR2.BackColor = Color.LightGray
                DERROR2.Visible = True
                DERROR2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DERROR2.Visible = True
                DERROR2.BackColor = Color.GreenYellow
                DERROR2.ReadOnly = False
                ShowRequiredFieldValidator("DERROR2Rqd", "DERROR2", "異常：需輸入PF/VF/MF 不良項目")
            Case 2  '修改
                DERROR2.Visible = True
                DERROR2.BackColor = Color.Yellow
                DERROR2.ReadOnly = False
            Case Else   '隱藏
                DERROR2.Visible = False
        End Select

        'PN 不良項目
        Select Case FindFieldInf("ERROR3")
            Case 0  '顯示
                DERROR3.BackColor = Color.LightGray
                DERROR3.Visible = True
                DERROR3.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DERROR3.Visible = True
                DERROR3.BackColor = Color.GreenYellow
                DERROR3.ReadOnly = False
                ShowRequiredFieldValidator("DERROR3Rqd", "DERROR3", "異常：需輸入PN 不良項目")
            Case 2  '修改
                DERROR3.Visible = True
                DERROR3.BackColor = Color.Yellow
                DERROR3.ReadOnly = False
            Case Else   '隱藏
                DERROR3.Visible = False
        End Select

        'QF 不良項目
        Select Case FindFieldInf("ERROR4")
            Case 0  '顯示
                DERROR4.BackColor = Color.LightGray
                DERROR4.Visible = True
                DERROR4.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DERROR4.Visible = True
                DERROR4.BackColor = Color.GreenYellow
                DERROR4.ReadOnly = False
                ShowRequiredFieldValidator("DERROR4Rqd", "DERROR4", "異常：需輸入QF 不良項目")
            Case 2  '修改
                DERROR4.Visible = True
                DERROR4.BackColor = Color.Yellow
                DERROR4.ReadOnly = False
            Case Else   '隱藏
                DERROR4.Visible = False
        End Select


        'NT 不良項目
        Select Case FindFieldInf("ERROR5")
            Case 0  '顯示
                DERROR5.BackColor = Color.LightGray
                DERROR5.Visible = True
                DERROR5.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DERROR5.Visible = True
                DERROR5.BackColor = Color.GreenYellow
                DERROR5.ReadOnly = False
                ShowRequiredFieldValidator("DERROR5Rqd", "DERROR5", "異常：需輸入NT 不良項目")
            Case 2  '修改
                DERROR5.Visible = True
                DERROR5.BackColor = Color.Yellow
                DERROR5.ReadOnly = False
            Case Else   '隱藏
                DERROR5.Visible = False
        End Select

        '不良狀況
        Select Case FindFieldInf("ERRORSTS")
            Case 0  '顯示
                DERRORSTS.BackColor = Color.LightGray
                DERRORSTS.Visible = True
                DERRORSTS.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DERRORSTS.Visible = True
                DERRORSTS.BackColor = Color.GreenYellow
                DERRORSTS.ReadOnly = False
                ShowRequiredFieldValidator("DERRORSTSRqd", "DERRORSTS", "異常：需輸入不良狀況")
            Case 2  '修改
                DERRORSTS.Visible = True
                DERRORSTS.BackColor = Color.Yellow
                DERRORSTS.ReadOnly = False
            Case Else   '隱藏
                DERRORSTS.Visible = False
        End Select


        '不良內容
        Select Case FindFieldInf("ECONTENT")
            Case 0  '顯示
                DECONTENT.BackColor = Color.LightGray
                DECONTENT.Visible = True

            Case 1  '修改+檢查
                DECONTENT.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DECONTENTRqd", "DECONTENT", "異常：需輸入不良內容")
                DECONTENT.Visible = True
            Case 2  '修改
                DECONTENT.BackColor = Color.Yellow
                DECONTENT.Visible = True
            Case Else   '隱藏
                DECONTENT.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ECONTENT", "ZZZZZZ")


        '其它不良內容
        Select Case FindFieldInf("ECONTENT1")
            Case 0  '顯示
                DECONTENT1.BackColor = Color.LightGray
                DECONTENT1.Visible = True
                DECONTENT1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DECONTENT1.Visible = True
                DECONTENT1.BackColor = Color.GreenYellow
                DECONTENT1.ReadOnly = False
                ShowRequiredFieldValidator("DECONTENT1Rqd", "DECONTENT1", "異常：需輸入不良狀況")
            Case 2  '修改
                DECONTENT1.Visible = True
                DECONTENT1.BackColor = Color.Yellow
                DECONTENT1.ReadOnly = False
            Case Else   '隱藏
                DECONTENT1.Visible = False
        End Select

        '不良原因
        Select Case FindFieldInf("EREASON")
            Case 0  '顯示
                DEREASON.BackColor = Color.LightGray
                DEREASON.Visible = True

            Case 1  '修改+檢查
                DEREASON.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEREASONRqd", "DEREASON", "異常：需輸入不良原因")
                DEREASON.Visible = True
            Case 2  '修改
                DEREASON.BackColor = Color.Yellow
                DEREASON.Visible = True
            Case Else   '隱藏
                DEREASON.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("EREASON", "ZZZZZZ")

        '其它不良原因
        Select Case FindFieldInf("EREASON1")
            Case 0  '顯示
                DEREASON1.BackColor = Color.LightGray
                DEREASON1.Visible = True
                DEREASON1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DEREASON1.Visible = True
                DEREASON1.BackColor = Color.GreenYellow
                DEREASON1.ReadOnly = False
                ShowRequiredFieldValidator("DEREASON1Rqd", "DEREASON1", "異常：需輸入不良狀況")
            Case 2  '修改
                DEREASON1.Visible = True
                DEREASON1.BackColor = Color.Yellow
                DEREASON1.ReadOnly = False
            Case Else   '隱藏
                DEREASON1.Visible = False
        End Select

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


        '其它處理情形
        Select Case FindFieldInf("SITUATION1")
            Case 0  '顯示
                DSITUATION1.BackColor = Color.LightGray
                DSITUATION1.Visible = True
                DSITUATION1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSITUATION1.Visible = True
                DSITUATION1.BackColor = Color.GreenYellow
                DSITUATION1.ReadOnly = False
                ShowRequiredFieldValidator("DSITUATION1Rqd", "DSITUATION1", "異常：需輸入其它處理情形")
            Case 2  '修改
                DSITUATION1.Visible = True
                DSITUATION1.BackColor = Color.Yellow
                DSITUATION1.ReadOnly = False
            Case Else   '隱藏
                DSITUATION1.Visible = False
        End Select


        '矯正及預防措施
        Select Case FindFieldInf("QCANSWER")
            Case 0  '顯示
                DQCANSWER.BackColor = Color.LightGray
                DQCANSWER.Visible = True

            Case 1  '修改+檢查
                DQCANSWER.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCANSWERRqd", "DQCANSWER", "異常：需輸入矯正及預防措施")
                DQCANSWER.Visible = True
            Case 2  '修改
                DQCANSWER.BackColor = Color.Yellow
                DQCANSWER.Visible = True
            Case Else   '隱藏
                DQCANSWER.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCANSWER", "ZZZZZZ")



        '其它矯正及預防措施
        Select Case FindFieldInf("QCANSWER1")
            Case 0  '顯示
                DQCANSWER1.BackColor = Color.LightGray
                DQCANSWER1.Visible = True
                DQCANSWER1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DQCANSWER1.Visible = True
                DQCANSWER1.BackColor = Color.GreenYellow
                DQCANSWER1.ReadOnly = False
                ShowRequiredFieldValidator("DQCANSWER1Rqd", "DQCANSWER1", "異常：需輸入其它矯正及預防措施")
            Case 2  '修改
                DQCANSWER1.Visible = True
                DQCANSWER1.BackColor = Color.Yellow
                DQCANSWER1.ReadOnly = False
            Case Else   '隱藏
                DQCANSWER1.Visible = False
        End Select




        '發生主要原因
        Select Case FindFieldInf("QCI1")
            Case 0  '顯示
                DQCI1.BackColor = Color.LightGray
                DQCI1.Visible = True
                DQCI1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DQCI1.Visible = True
                DQCI1.BackColor = Color.GreenYellow
                DQCI1.ReadOnly = False
                ShowRequiredFieldValidator("DQCI1Rqd", "DQCI1", "異常：需輸入發生主要原因")
            Case 2  '修改
                DQCI1.Visible = True
                DQCI1.BackColor = Color.Yellow
                DQCI1.ReadOnly = False
            Case Else   '隱藏
                DQCI1.Visible = False
        End Select



        '暫定對應
        Select Case FindFieldInf("QCI2")
            Case 0  '顯示
                DQCI2.BackColor = Color.LightGray
                DQCI2.Visible = True
                DQCI2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DQCI2.Visible = True
                DQCI2.BackColor = Color.GreenYellow
                DQCI2.ReadOnly = False
                ShowRequiredFieldValidator("DQCI2Rqd", "DQCI2", "異常：需輸入發生暫定對應")
            Case 2  '修改
                DQCI2.Visible = True
                DQCI2.BackColor = Color.Yellow
                DQCI2.ReadOnly = False
            Case Else   '隱藏
                DQCI2.Visible = False
        End Select


        '發生開始期間
        Select Case FindFieldInf("QCIDATE1")
            Case 0  '顯示
                DQCIDATE1.Visible = True
                DQCIDATE1.Style.Add("background-color", "lightgrey")
                DQCIDATE1.Attributes.Add("readonly", "true")
                BQCIDATE1.Visible = False

            Case 1  '修改+檢查
                DQCIDATE1.Visible = True
                DQCIDATE1.Style.Add("background-color", "greenyellow")
                DQCIDATE1.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DQCIDATE1Rqd", "DQCIDATE1", "異常：需輸入發生開始期間")
                BQCIDATE1.Visible = True

            Case 2  '修改
                DQCIDATE1.Visible = True
                DQCIDATE1.Style.Add("background-color", "yellow")
                DQCIDATE1.Attributes.Add("readonly", "true")
                BQCIDATE1.Visible = True
            Case Else   '隱藏
                DQCIDATE1.Visible = False
                BQCIDATE1.Visible = False

        End Select
        If pPost = "New" Then DQCIDATE1.Text = Now.ToString("yyyy/MM/dd") '現在日時


        '發生結束期間
        Select Case FindFieldInf("QCIDATE2")
            Case 0  '顯示
                DQCIDATE2.Visible = True
                DQCIDATE2.Style.Add("background-color", "lightgrey")
                DQCIDATE2.Attributes.Add("readonly", "true")
                BQCIDATE2.Visible = False

            Case 1  '修改+檢查
                DQCIDATE2.Visible = True
                DQCIDATE2.Style.Add("background-color", "greenyellow")
                DQCIDATE2.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DQCIDATE2Rqd", "DQCIDATE2", "異常：需輸入發生開始期間")
                BQCIDATE2.Visible = True

            Case 2  '修改
                DQCIDATE2.Visible = True
                DQCIDATE2.Style.Add("background-color", "yellow")
                DQCIDATE2.Attributes.Add("readonly", "true")
                BQCIDATE2.Visible = True
            Case Else   '隱藏
                DQCIDATE2.Visible = False
                BQCIDATE2.Visible = False

        End Select
        If pPost = "New" Then DQCIDATE2.Text = Now.ToString("yyyy/MM/dd") '現在日時


        '恆久對應
        Select Case FindFieldInf("QCI3")
            Case 0  '顯示
                DQCI3.BackColor = Color.LightGray
                DQCI3.Visible = True
                DQCI3.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DQCI3.Visible = True
                DQCI3.BackColor = Color.GreenYellow
                DQCI3.ReadOnly = False
                ShowRequiredFieldValidator("DQCI3Rqd", "DQCI3", "異常：需輸入發生恆久對應")
            Case 2  '修改
                DQCI3.Visible = True
                DQCI3.BackColor = Color.Yellow
                DQCI3.ReadOnly = False
            Case Else   '隱藏
                DQCI3.Visible = False
        End Select


        '發生開始實施期間
        Select Case FindFieldInf("QCIDATE3")
            Case 0  '顯示
                DQCIDATE3.Visible = True
                DQCIDATE3.Style.Add("background-color", "lightgrey")
                DQCIDATE3.Attributes.Add("readonly", "true")
                BQCIDATE3.Visible = False

            Case 1  '修改+檢查
                DQCIDATE3.Visible = True
                DQCIDATE3.Style.Add("background-color", "greenyellow")
                DQCIDATE3.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DQCIDATE3Rqd", "DQCIDATE3", "異常：需輸入發生開始實施期間")
                BQCIDATE3.Visible = True

            Case 2  '修改
                DQCIDATE3.Visible = True
                DQCIDATE3.Style.Add("background-color", "yellow")
                DQCIDATE3.Attributes.Add("readonly", "true")
                BQCIDATE3.Visible = True
            Case Else   '隱藏
                DQCIDATE3.Visible = False
                BQCIDATE3.Visible = False

        End Select
        If pPost = "New" Then DQCIDATE3.Text = Now.ToString("yyyy/MM/dd") '現在日時


        '發生結束實施期間
        Select Case FindFieldInf("QCIDATE4")
            Case 0  '顯示
                DQCIDATE4.Visible = True
                DQCIDATE4.Style.Add("background-color", "lightgrey")
                DQCIDATE4.Attributes.Add("readonly", "true")
                BQCIDATE4.Visible = False

            Case 1  '修改+檢查
                DQCIDATE4.Visible = True
                DQCIDATE4.Style.Add("background-color", "greenyellow")
                DQCIDATE4.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DQCIDATE4Rqd", "DQCIDATE4", "異常：需輸入發生結束實施期間")
                BQCIDATE4.Visible = True

            Case 2  '修改
                DQCIDATE4.Visible = True
                DQCIDATE4.Style.Add("background-color", "yellow")
                DQCIDATE4.Attributes.Add("readonly", "true")
                BQCIDATE4.Visible = True
            Case Else   '隱藏
                DQCIDATE4.Visible = False
                BQCIDATE4.Visible = False

        End Select
        If pPost = "New" Then DQCIDATE4.Text = Now.ToString("yyyy/MM/dd") '現在日時




        '流出主要原因
        Select Case FindFieldInf("QCO1")
            Case 0  '顯示
                DQCO1.BackColor = Color.LightGray
                DQCO1.Visible = True
                DQCO1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DQCO1.Visible = True
                DQCO1.BackColor = Color.GreenYellow
                DQCO1.ReadOnly = False
                ShowRequiredFieldValidator("DQCO1Rqd", "DQCO1", "異常：需輸入流出主要原因")
            Case 2  '修改
                DQCO1.Visible = True
                DQCO1.BackColor = Color.Yellow
                DQCO1.ReadOnly = False
            Case Else   '隱藏
                DQCO1.Visible = False
        End Select



        '流出暫定對應
        Select Case FindFieldInf("QCO2")
            Case 0  '顯示
                DQCO2.BackColor = Color.LightGray
                DQCO2.Visible = True
                DQCO2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DQCO2.Visible = True
                DQCO2.BackColor = Color.GreenYellow
                DQCO2.ReadOnly = False
                ShowRequiredFieldValidator("DQCO2Rqd", "DQCO2", "異常：需輸入流出暫定對應")
            Case 2  '修改
                DQCO2.Visible = True
                DQCO2.BackColor = Color.Yellow
                DQCO2.ReadOnly = False
            Case Else   '隱藏
                DQCO2.Visible = False
        End Select


        '流出開始期間
        Select Case FindFieldInf("QCODATE1")
            Case 0  '顯示
                DQCODATE1.Visible = True
                DQCODATE1.Style.Add("background-color", "lightgrey")
                DQCODATE1.Attributes.Add("readonly", "true")
                BQCODATE1.Visible = False

            Case 1  '修改+檢查
                DQCODATE1.Visible = True
                DQCODATE1.Style.Add("background-color", "greenyellow")
                DQCODATE1.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DQCODATE1Rqd", "DQCODATE1", "異常：需輸入流出開始期間")
                BQCODATE1.Visible = True

            Case 2  '修改
                DQCODATE1.Visible = True
                DQCODATE1.Style.Add("background-color", "yellow")
                DQCODATE1.Attributes.Add("readonly", "true")
                BQCODATE1.Visible = True
            Case Else   '隱藏
                DQCODATE1.Visible = False
                BQCODATE1.Visible = False

        End Select
        If pPost = "New" Then DQCODATE1.Text = Now.ToString("yyyy/MM/dd") '現在日時


        '流出結束期間
        Select Case FindFieldInf("QCODATE2")
            Case 0  '顯示
                DQCODATE2.Visible = True
                DQCODATE2.Style.Add("background-color", "lightgrey")
                DQCODATE2.Attributes.Add("readonly", "true")
                BQCODATE2.Visible = False

            Case 1  '修改+檢查
                DQCODATE2.Visible = True
                DQCODATE2.Style.Add("background-color", "greenyellow")
                DQCODATE2.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DQCODATE2Rqd", "DQCODATE2", "異常：需輸入流出結束期間")
                BQCODATE2.Visible = True

            Case 2  '修改
                DQCODATE2.Visible = True
                DQCODATE2.Style.Add("background-color", "yellow")
                DQCODATE2.Attributes.Add("readonly", "true")
                BQCODATE2.Visible = True
            Case Else   '隱藏
                DQCODATE2.Visible = False
                BQCODATE2.Visible = False

        End Select
        If pPost = "New" Then DQCODATE2.Text = Now.ToString("yyyy/MM/dd") '現在日時


        '流出恆久對應
        Select Case FindFieldInf("QCO3")
            Case 0  '顯示
                DQCO3.BackColor = Color.LightGray
                DQCO3.Visible = True
                DQCO3.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DQCO3.Visible = True
                DQCO3.BackColor = Color.GreenYellow
                DQCO3.ReadOnly = False
                ShowRequiredFieldValidator("DQCO3Rqd", "DQCO3", "異常：需輸入發生恆久對應")
            Case 2  '修改
                DQCO3.Visible = True
                DQCO3.BackColor = Color.Yellow
                DQCO3.ReadOnly = False
            Case Else   '隱藏
                DQCO3.Visible = False
        End Select


        '流出開始實施期間
        Select Case FindFieldInf("QCODATE3")
            Case 0  '顯示
                DQCODATE3.Visible = True
                DQCODATE3.Style.Add("background-color", "lightgrey")
                DQCODATE3.Attributes.Add("readonly", "true")
                BQCODATE3.Visible = False

            Case 1  '修改+檢查
                DQCODATE3.Visible = True
                DQCODATE3.Style.Add("background-color", "greenyellow")
                DQCODATE3.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DQCODATE3Rqd", "DQCODATE3", "異常：需輸入流出開始實施期間")
                BQCODATE3.Visible = True

            Case 2  '修改
                DQCODATE3.Visible = True
                DQCODATE3.Style.Add("background-color", "yellow")
                DQCODATE3.Attributes.Add("readonly", "true")
                BQCODATE3.Visible = True
            Case Else   '隱藏
                DQCODATE3.Visible = False
                BQCODATE3.Visible = False

        End Select
        If pPost = "New" Then DQCODATE3.Text = Now.ToString("yyyy/MM/dd") '現在日時


        '流出結束實施期間
        Select Case FindFieldInf("QCODATE4")
            Case 0  '顯示
                DQCODATE4.Visible = True
                DQCODATE4.Style.Add("background-color", "lightgrey")
                DQCODATE4.Attributes.Add("readonly", "true")
                BQCODATE4.Visible = False

            Case 1  '修改+檢查
                DQCODATE4.Visible = True
                DQCODATE4.Style.Add("background-color", "greenyellow")
                DQCODATE4.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DQCODATE4Rqd", "DQCODATE4", "異常：需輸入流出結束實施期間")
                BQCODATE4.Visible = True

            Case 2  '修改
                DQCODATE4.Visible = True
                DQCODATE4.Style.Add("background-color", "yellow")
                DQCODATE4.Attributes.Add("readonly", "true")
                BQCODATE4.Visible = True
            Case Else   '隱藏
                DQCODATE4.Visible = False
                BQCODATE4.Visible = False

        End Select
        If pPost = "New" Then DQCODATE4.Text = Now.ToString("yyyy/MM/dd") '現在日時



        '預防措施(OLD)
        Select Case FindFieldInf("QCANSWER2")
            Case 0  '顯示
                DQCANSWER2.BackColor = Color.LightGray
                DQCANSWER2.Visible = True
                DQCANSWER2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DQCANSWER2.Visible = True
                DQCANSWER2.BackColor = Color.GreenYellow
                DQCANSWER2.ReadOnly = False
                ShowRequiredFieldValidator("DQCANSWER2Rqd", "DQCANSWER2", "異常：需輸入預防措施(OLD)")
            Case 2  '修改
                DQCANSWER2.Visible = True
                DQCANSWER2.BackColor = Color.Yellow
                DQCANSWER2.ReadOnly = False
            Case Else   '隱藏
                DQCANSWER2.Visible = False
        End Select


        '何時處理完成（出貨日）
        Select Case FindFieldInf("FINISHDATE")
            Case 0  '顯示
                DFINISHDATE.Visible = True
                DFINISHDATE.Style.Add("background-color", "lightgrey")
                DFINISHDATE.Attributes.Add("readonly", "true")
                BFINISHDATE.Visible = False

            Case 1  '修改+檢查
                DFINISHDATE.Visible = True
                DFINISHDATE.Style.Add("background-color", "greenyellow")
                DFINISHDATE.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DFINISHDATERqd", "DFINISHDATE", "異常：需輸入何時處理完成（出貨日）")
                BFINISHDATE.Visible = True

            Case 2  '修改
                DFINISHDATE.Visible = True
                DFINISHDATE.Style.Add("background-color", "yellow")
                DFINISHDATE.Attributes.Add("readonly", "true")
                BFINISHDATE.Visible = True
            Case Else   '隱藏
                DFINISHDATE.Visible = False
                BFINISHDATE.Visible = False

        End Select
        If pPost = "New" Then DFINISHDATE.Text = Now.ToString("yyyy/MM/dd") '現在日時



        '被訴部門擔當者 
        Select Case FindFieldInf("ACCEMPNAME")
            Case 0  '顯示
                DACCEMPNAME.BackColor = Color.LightGray
                DACCEMPNAME.Visible = True
                DACCEMPNAME.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DACCEMPNAME.Visible = True
                DACCEMPNAME.BackColor = Color.GreenYellow
                DACCEMPNAME.ReadOnly = False
                ShowRequiredFieldValidator("DACCEMPNAMERqd", "DACCEMPNAME", "異常：需輸入被訴部門擔當者")
            Case 2  '修改
                DACCEMPNAME.Visible = True
                DACCEMPNAME.BackColor = Color.Yellow
                DACCEMPNAME.ReadOnly = False
            Case Else   '隱藏
                DACCEMPNAME.Visible = False
        End Select





        '附檔1
        Select Case FindFieldInf("ATTACHFILE1")
            Case 0  '顯示
                DAttachfile1.Visible = False
            Case 1  '修改+檢查

                DAttachfile1.Visible = True

            Case 2  '修改
                DAttachfile1.Visible = True

            Case Else   '隱藏
                DAttachfile1.Visible = False
        End Select


        '附檔2
        Select Case FindFieldInf("ATTACHFILE2")
            Case 0  '顯示
                DATTACHFILE2.Visible = False
            Case 1  '修改+檢查

                DATTACHFILE2.Visible = True

            Case 2  '修改
                DAttachfile2.Visible = True

            Case Else   '隱藏
                DAttachfile2.Visible = False
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
                    System.Configuration.ConfigurationManager.AppSettings("ExpensePath")
        Dim FileName As String = ""

        Dim SQL As String
        SQL = "Select * From F_FinalcheckSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dtData.Rows(0).Item("No")                         'No
            DCUSTOMER.Text = dtData.Rows(0).Item("CUSTOMER")
            SetFieldData("ACCDEPNAME", dtData.Rows(0).Item("ACCDEPNAME"))    '顧客
            SetFieldData("SLDDIVISION", dtData.Rows(0).Item("SLDDIVISION"))
            SetFieldData("RDIVISION", dtData.Rows(0).Item("RDIVISION"))
            DCHECKDATE.Text = dtData.Rows(0).Item("CHECKDATE")
            DORNO.Text = dtData.Rows(0).Item("ORNO")
            DCOLORITEM.Text = dtData.Rows(0).Item("COLORITEM")
            DQTY.Text = dtData.Rows(0).Item("QTY")
            SetFieldData("UNIT1", dtData.Rows(0).Item("UNIT1"))
            DCHECKQTY.Text = dtData.Rows(0).Item("CHECKQTY")
            SetFieldData("UNIT2", dtData.Rows(0).Item("UNIT2"))
            DERRORQTY.Text = dtData.Rows(0).Item("ERRORQTY")
            SetFieldData("UNIT3", dtData.Rows(0).Item("UNIT3"))
            DAPPNAME.Text = dtData.Rows(0).Item("APPNAME")
            DDATE.Text = dtData.Rows(0).Item("DATE")

            DERROR1.Text = dtData.Rows(0).Item("ERROR1")
            DERROR2.Text = dtData.Rows(0).Item("ERROR2")
            DERROR3.Text = dtData.Rows(0).Item("ERROR3")
            DERROR4.Text = dtData.Rows(0).Item("ERROR4")
            DERROR5.Text = dtData.Rows(0).Item("ERROR5")
            DERRORSTS.Text = dtData.Rows(0).Item("ERRORSTS")


            SetFieldData("ECONTENT", dtData.Rows(0).Item("ECONTENT"))
            DECONTENT1.Text = dtData.Rows(0).Item("ECONTENT1")
            SetFieldData("EREASON", dtData.Rows(0).Item("EREASON"))
            DEREASON1.Text = dtData.Rows(0).Item("EREASON1")
            SetFieldData("SITUATION", dtData.Rows(0).Item("SITUATION"))
            DSITUATION1.Text = dtData.Rows(0).Item("SITUATION1")
            SetFieldData("QCANSWER", dtData.Rows(0).Item("QCANSWER"))
            DQCANSWER1.Text = dtData.Rows(0).Item("QCANSWER1")

            DQCI1.Text = dtData.Rows(0).Item("QCI1")
            DQCI2.Text = dtData.Rows(0).Item("QCI2")
            DQCIDATE1.Text = dtData.Rows(0).Item("QCIDATE1")
            DQCIDATE2.Text = dtData.Rows(0).Item("QCIDATE2")
            DQCI3.Text = dtData.Rows(0).Item("QCI3")
            DQCIDATE3.Text = dtData.Rows(0).Item("QCIDATE3")
            DQCIDATE4.Text = dtData.Rows(0).Item("QCIDATE4")

            DQCO1.Text = dtData.Rows(0).Item("QCO1")
            DQCO2.Text = dtData.Rows(0).Item("QCO2")
            DQCODATE1.Text = dtData.Rows(0).Item("QCODATE1")
            DQCODATE2.Text = dtData.Rows(0).Item("QCODATE2")
            DQCO3.Text = dtData.Rows(0).Item("QCO3")
            DQCODATE3.Text = dtData.Rows(0).Item("QCODATE3")
            DQCODATE4.Text = dtData.Rows(0).Item("QCODATE4")


            DFINISHDATE.Text = dtData.Rows(0).Item("FINISHDATE")
            DACCEMPNAME.Text = dtData.Rows(0).Item("ACCEMPNAME")


            If dtData.Rows(0).Item("Mod") = 1 Then       '有新色依賴完成表
                LComplete.NavigateUrl = "FinalcheckModList.aspx?pformno=" & dtData.Rows(0).Item("formNo") & "&pformsno=" & dtData.Rows(0).Item("formsNo")
                LComplete.Visible = True
            Else
                LComplete.Visible = False
            End If



            '主檔資料
            SQL = " select  data,replace(data,'/','\')data1  from M_referp"
            SQL = SQL + " where cat = '3106'"
            SQL = SQL + " and dkey ='AttachfilePath1'"
            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
            Dim OpenDir As String = ""
            If DBAdapter3.Rows.Count > 0 Then
                OpenDir = "file://" + DBAdapter3.Rows(0).Item("Data") + DNO.Text + "/QC"
            End If

            'Dim OpenDir As String
            'OpenDir = "file://10.245.1.18/MIS/DASW/" + DNo.Text

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



        DAPPNAME.Text = DBUser.Rows(0).Item("Username")

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
                sql = "  Select substring(data,1,CHARINDEX('-', data)-1)data  from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'RDIVISION'"
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
                sql = "  Select  substring(data,1,CHARINDEX('-', data)-1)data from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'SLDDIVISION'"
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


        '相關部門
        If pFieldName = "RDIVISION" Then
            DRDIVISION.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DRDIVISION.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  substring(data,1,CHARINDEX('-', data)-1)data from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'RDIVISION'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DRDIVISION.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DRDIVISION.Items.Add(ListItem1)
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

        'UNIT3
        If pFieldName = "UNIT3" Then
            DUNIT3.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DUNIT3.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'CHECKQTY'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DUNIT3.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DUNIT3.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'ECONTENT
        If pFieldName = "ECONTENT" Then
            DECONTENT.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DECONTENT.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'ECONTENT'"
                sql = sql & " order by DATA"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DECONTENT.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DECONTENT.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'EREASON
        If pFieldName = "EREASON" Then
            DEREASON.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DEREASON.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'EREASON'"
                sql = sql & " order by DATA"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DEREASON.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DEREASON.Items.Add(ListItem1)
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
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'SITUATION'"
                sql = sql & " order by DATA"
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


        'QCANSWER
        If pFieldName = "QCANSWER" Then
            DQCANSWER.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCANSWER.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'QCANSWER'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DQCANSWER.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCANSWER.Items.Add(ListItem1)
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

        BCHECKDATE.Attributes("onclick") = "calendarPicker('Form1.DCHECKDATE');"
        BQCIDATE1.Attributes("onclick") = "calendarPicker('Form1.DQCIDATE1');"
        BQCIDATE2.Attributes("onclick") = "calendarPicker('Form1.DQCIDATE2');"
        BQCIDATE3.Attributes("onclick") = "calendarPicker('Form1.DQCIDATE3');"
        BQCIDATE4.Attributes("onclick") = "calendarPicker('Form1.DQCIDATE4');"

        BQCODATE1.Attributes("onclick") = "calendarPicker('Form1.DQCODATE1');"
        BQCODATE2.Attributes("onclick") = "calendarPicker('Form1.DQCODATE2');"
        BQCODATE3.Attributes("onclick") = "calendarPicker('Form1.DQCODATE3');"
        BQCODATE4.Attributes("onclick") = "calendarPicker('Form1.DQCODATE4');"


        BERROR1.Attributes.Add("onClick", "GetError('ERROR1')") '找buyer
        BERROR2.Attributes.Add("onClick", "GetError('ERROR2')") '找buyer
        BERROR3.Attributes.Add("onClick", "GetError('ERROR3')") '找buyer
        BERROR4.Attributes.Add("onClick", "GetError('ERROR4')") '找buyer
        BERROR5.Attributes.Add("onClick", "GetError('ERROR5')") '找buyer

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
                        'DNo.Text = SetNo(NewFormSno)
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

                If wStep = 1 Or wStep = 500 Then
                    wAllocateID = GetUerID(Trim(DACCEMPNAME.Text))  '被訴擔當
                End If
                If (wStep = 10 Or wStep = 20 Or wStep = 15) And pFun = "OK" Then

                    If DACCDEPNAME.SelectedValue = "SLD" And wStep = 10 Then
                        wAllocateID = GetUerID(Trim(DACCEMPNAME.Text))
                        pAction = 2
                    ElseIf wStep = 10 And wStep = 15 Then
                        If GetRelated(GetUerID(Trim(DACCEMPNAME.Text))) = GetUerID(Trim(DACCEMPNAME.Text)) Then  '如果被訴擔當跟台籍主管一樣直接跳日籍
                            If wStep = 10 Then
                                pAction = 3
                            Else
                                pAction = 1
                            End If

                        End If
                        wAllocateID = GetRelated(GetUerID(Trim(DACCEMPNAME.Text)))

                    Else

                        wAllocateID = GetRelated(GetUerID(Trim(DACCEMPNAME.Text)))
                    End If

                End If

                If (wStep = 30 Or wStep = 40) And pFun = "OK" Then
                    wAllocateID = GetRelated(wApplyID)
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
                        '   oCommon.Send(Request.QueryString("pUserID"), pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "FLOW")
                        '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                    Next i
                Else
                    '  oCommon.Send(Request.QueryString("pUserID"), wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "END")
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
                    System.Configuration.ConfigurationManager.AppSettings("ExpensePath")

        Dim sql As String = ""
        sql = " Insert into F_FinalcheckSheet (Sts, CompletedTime, FormNo, FormSno,"
        sql = sql + " NO,CUSTOMER,ACCDEPNAME,SLDDIVISION,RDIVISION,TYPE,CHECKDATE,ORNO,COLORITEM,"
        sql = sql + " QTY,UNIT1,CHECKQTY,UNIT2,ERRORQTY,UNIT3,APPNAME,DATE,"
        sql = sql + " ERROR1,ERROR2,ERROR3,ERROR4,ERROR5,ERRORSTS,"
        sql = sql + " ECONTENT,ECONTENT1,EREASON,EREASON1,SITUATION,SITUATION1,QCANSWER,QCANSWER1,"
        sql = sql + " QCI1,QCI2,QCIDATE1,QCIDATE2,QCI3,QCIDATE3,QCIDATE4,"
        sql = sql + " QCO1,QCO2,QCODATE1,QCODATE2,QCO3,QCODATE3,QCODATE4,"
        sql = sql + " QCANSWER2,FINISHDATE,ACCEMPNAME,"
        sql = sql + " CreateUser, CreateTime, ModifyUser, ModifyTime) "
        sql = sql + " VALUES( "
        sql = sql + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        sql = sql + " '" + NowDateTime + "', "        '結案日
        sql = sql + " '003106', "                     '表單代號
        sql = sql + " '" + CStr(NewFormSno) + "', "   '表單流水號

        sql = sql + " N'" + YKK.ReplaceString(DNO.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DCUSTOMER.Text) + "', "   'NO  1
        sql = sql + " N'" + DACCDEPNAME.SelectedValue + "', "
        sql = sql + " N'" + DSLDDIVISION.SelectedValue + "', "
        sql = sql + " N'" + DRDIVISION.SelectedValue + "', "
        sql = sql + " N'" + DTYPE.Text + "', "
        sql = sql + "'" + DCHECKDATE.Text + "', "
        sql = sql + " N'" + DORNO.Text + "', "
        sql = sql + " N'" + DCOLORITEM.Text + "', "

        sql = sql + " N'" + DQTY.Text + "', "
        sql = sql + " N'" + DUNIT1.SelectedValue + "', "
        sql = sql + " N'" + DCHECKQTY.Text + "', "
        sql = sql + " N'" + DUNIT2.SelectedValue + "', "
        sql = sql + " N'" + DERRORQTY.Text + "', "
        sql = sql + " N'" + DUNIT3.SelectedValue + "', "

        sql = sql + " N'" + DAPPNAME.Text + "', "
        sql = sql + "'" + DDATE.Text + "', "

        sql = sql + " N'" + YKK.ReplaceString(DERROR1.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DERROR2.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DERROR3.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DERROR4.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DERROR5.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DERRORSTS.Text) + "', "

        sql = sql + " N'" + YKK.ReplaceString(DECONTENT.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DECONTENT1.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DEREASON.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DEREASON1.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DSITUATION.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DSITUATION1.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DQCANSWER.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DQCANSWER1.Text) + "', "


        sql = sql + " N'" + YKK.ReplaceString(DQCI1.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DQCI2.Text) + "', "
        sql = sql + "'" + DQCIDATE1.Text + "', "
        sql = sql + "'" + DQCIDATE2.Text + "', "
        sql = sql + " N'" + YKK.ReplaceString(DQCI3.Text) + "', "
        sql = sql + "'" + DQCIDATE3.Text + "', "
        sql = sql + "'" + DQCIDATE4.Text + "', "


        sql = sql + " N'" + YKK.ReplaceString(DQCO1.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DQCO2.Text) + "', "
        sql = sql + "'" + DQCODATE1.Text + "', "
        sql = sql + "'" + DQCODATE2.Text + "', "
        sql = sql + " N'" + YKK.ReplaceString(DQCO3.Text) + "', "
        sql = sql + "'" + DQCODATE3.Text + "', "
        sql = sql + "'" + DQCODATE4.Text + "', "

        sql = sql + " N'" + YKK.ReplaceString(DQCANSWER2.Text) + "', "
        sql = sql + "'" + DFINISHDATE.Text + "', "
        sql = sql + " N'" + YKK.ReplaceString(DACCEMPNAME.Text) + "', "


        sql = sql + "'" & Request.QueryString("pUserID") & "' ,"
        sql = sql + "'" & NowDateTime & "' ,"
        sql = sql + "'" & Request.QueryString("pUserID") & "' ,"
        sql = sql + "'" & NowDateTime & "' )"
        uDataBase.ExecuteNonQuery(sql)

        sourceDir = "\\10.245.1.61\wfs$\N2W\003106Temp\" + D3.Text   '來源
        backupDir = "\\10.245.1.61\wfs$\N2W\003106\" + DNO.Text      '目的     

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
                    System.Configuration.ConfigurationManager.AppSettings("ExpensePath")
        Dim sql As String


        sql = " Update F_FinalcheckSheet"
        sql = sql + " Set "
        If pFun <> "SAVE" Then
            sql = sql + " Sts = '" & pSts & "',"
            sql = sql + " CompletedTime = '" & NowDateTime & "',"
        End If
        sql = sql + " No = N'" & YKK.ReplaceString(DNO.Text) & "',"

        sql = sql + " CUSTOMER = N'" + YKK.ReplaceString(DCUSTOMER.Text) + "', "   'NO  1
        sql = sql + " ACCDEPNAME = N'" + DACCDEPNAME.SelectedValue + "', "
        sql = sql + " SLDDIVISION = N'" + DSLDDIVISION.SelectedValue + "', "
        sql = sql + " RDIVISION = N'" + DRDIVISION.SelectedValue + "', "
        sql = sql + " CHECKDATE = '" + DCHECKDATE.Text + "', "
        sql = sql + " ORNO = N'" + DORNO.Text + "', "
        sql = sql + " COLORITEM = N'" + DCOLORITEM.Text + "', "
        sql = sql + " QTY = N'" + DQTY.Text + "', "
        sql = sql + " UNIT1 = N'" + DUNIT1.SelectedValue + "', "
        sql = sql + " CHECKQTY = N'" + DCHECKQTY.Text + "', "
        sql = sql + " UNIT2 = N'" + DUNIT2.SelectedValue + "', "
        sql = sql + " ERRORQTY = N'" + DERRORQTY.Text + "', "
        sql = sql + " UNIT3 = N'" + DUNIT3.SelectedValue + "', "

        sql = sql + " APPNAME = N'" + DAPPNAME.Text + "', "
        sql = sql + " DATE = '" + DDATE.Text + "', "

        sql = sql + " ERROR1 = N'" + YKK.ReplaceString(DERROR1.Text) + "', "
        sql = sql + " ERROR2 = N'" + YKK.ReplaceString(DERROR2.Text) + "', "
        sql = sql + " ERROR3 = N'" + YKK.ReplaceString(DERROR3.Text) + "', "
        sql = sql + " ERROR4 = N'" + YKK.ReplaceString(DERROR4.Text) + "', "
        sql = sql + " ERROR5 = N'" + YKK.ReplaceString(DERROR5.Text) + "', "
        sql = sql + " ERRORSTS = N'" + YKK.ReplaceString(DERRORSTS.Text) + "', "

        sql = sql + " ECONTENT = N'" + YKK.ReplaceString(DECONTENT.Text) + "', "
        sql = sql + " ECONTENT1 = N'" + YKK.ReplaceString(DECONTENT1.Text) + "', "
        sql = sql + " EREASON = N'" + YKK.ReplaceString(DEREASON.Text) + "', "
        sql = sql + " EREASON1 = N'" + YKK.ReplaceString(DEREASON1.Text) + "', "
        sql = sql + " SITUATION = N'" + YKK.ReplaceString(DSITUATION.Text) + "', "
        sql = sql + " SITUATION1 = N'" + YKK.ReplaceString(DSITUATION1.Text) + "', "
        sql = sql + " QCANSWER = N'" + YKK.ReplaceString(DQCANSWER.Text) + "', "
        sql = sql + " QCANSWER1 = N'" + YKK.ReplaceString(DQCANSWER1.Text) + "', "


        sql = sql + " QCI1 = N'" + YKK.ReplaceString(DQCI1.Text) + "', "
        sql = sql + " QCI2 = N'" + YKK.ReplaceString(DQCI2.Text) + "', "
        sql = sql + " QCIDATE1 = '" + DQCIDATE1.Text + "', "
        sql = sql + " QCIDATE2 = '" + DQCIDATE2.Text + "', "
        sql = sql + " QCI3 = N'" + YKK.ReplaceString(DQCI3.Text) + "', "
        sql = sql + " QCIDATE3 = '" + DQCIDATE3.Text + "', "
        sql = sql + " QCIDATE4 = '" + DQCIDATE4.Text + "', "


        sql = sql + " QCO1 = N'" + YKK.ReplaceString(DQCO1.Text) + "', "
        sql = sql + " QCO2 = N'" + YKK.ReplaceString(DQCO2.Text) + "', "
        sql = sql + " QCODATE1 = '" + DQCODATE1.Text + "', "
        sql = sql + " QCODATE2 ='" + DQCODATE2.Text + "', "
        sql = sql + " QCO3 = N'" + YKK.ReplaceString(DQCO3.Text) + "', "
        sql = sql + " QCODATE3 = '" + DQCODATE3.Text + "', "
        sql = sql + " QCODATE4 = '" + DQCODATE4.Text + "', "

        sql = sql + " QCANSWER2 = N'" + YKK.ReplaceString(DQCANSWER2.Text) + "', "
        sql = sql + " FINISHDATE= '" + DFINISHDATE.Text + "', "
        sql = sql + " ACCEMPNAME = N'" + YKK.ReplaceString(DACCEMPNAME.Text) + "', "

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
                ErrCode = oCommon.CommissionNo("003106", wFormSno, wStep, DNO.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9050
                End If
            End If
        End If

        Try
            If ErrCode = 0 Then
                If DQTY.Text <> "" Then
                    Dim QTY As Integer
                    QTY = DQTY.Text
                    If QTY <= 0 Then
                        ErrCode = 9073
                    End If
                End If


            End If

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
            If ErrCode = 9073 Then Message = "金額需大於0，請確認!"

            If ErrCode = 9075 Then Message = "非數字格式,請確認!"
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
    '**     設定附檔路徑
    '**
    '*****************************************************************
    Sub NewAttachFilePath()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '3106'"
        SQL = SQL + " and dkey ='AttachfilePath'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If


        If D3.Text = "" Then
            D3.Text = Now.ToString("yyyyMMddHHmmss")
        End If



        OpenDir1 = OpenDir1 + D3.Text + "/QC"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + D3.Text + "\QC"
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If



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
        SQL = SQL + " where cat = '3106'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If


        If D3.Text = "" Then
            D3.Text = Now.ToString("yyyyMMddHHmmss")
        End If



        OpenDir1 = OpenDir1 + DNO.Text + "/Other"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNO.Text + "\Other"
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If



        '開啟附檔資料夾路徑
        DAttachfile2.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")
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
            If wStep = 10 Or wStep = 30 Or wStep = 15 Then
                If dt.Rows(0)("UserID") = dt.Rows(0)("RUserID") Then '如果被訴擔當跟被訴擔當台籍主管一樣就直接跳日籍主管
                    NextGate = dt.Rows(0)("RRUserID")
                Else
                    NextGate = dt.Rows(0)("RUserID")
                End If

            ElseIf wStep = 20 Or wStep = 40 Then
                NextGate = dt.Rows(0)("RRUserID")
            End If

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



    Protected Sub DACCDEPNAME_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DACCDEPNAME.SelectedIndexChanged
        Dim sql As String
        sql = "  Select substring(data,CHARINDEX('-', data)+1,len(data)-1)data  from M_referp"
        Sql = Sql & " where  cat = '3106'"
        sql = sql & " and dkey = 'RDIVISION'"
        sql = sql & "and  data like '" + DACCDEPNAME.SelectedValue + "%'"
        sql = sql & " order by unique_id"
        Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
        DACCEMPNAME.Text = dtReferp.Rows(0).Item("Data")
        If DACCDEPNAME.SelectedValue = "SLD" Then
            DSLDDIVISION.BackColor = Color.GreenYellow
            ShowRequiredFieldValidator("DSLDDIVISIONRqd", "DSLDDIVISION", "異常：需輸入SLD工程別")
            DSLDDIVISION.Visible = True
            SetFieldData("SLDDIVISION", "ZZZZZZ")
        Else
            DSLDDIVISION.BackColor = Color.LightGray
            DSLDDIVISION.Visible = False
            DSLDDIVISION.SelectedValue = ""
        End If

    End Sub

    Protected Sub DSLDDIVISION_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DSLDDIVISION.SelectedIndexChanged
        If wStep = 10 Then
            Dim sql As String
            sql = "  Select substring(data,CHARINDEX('-', data)+1,len(data)-1)data  from M_referp"
            sql = sql & " where  cat = '3106'"
            sql = sql & " and dkey = 'SLDDIVISION'"
            sql = sql & "and  data like '" + DSLDDIVISION.SelectedValue + "%'"
            sql = sql & " order by unique_id"
            Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
            DACCEMPNAME.Text = dtReferp.Rows(0).Item("Data")
        End If

    End Sub
End Class
