Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Collections.Generic

Partial Class CommissionSheet_01
    Inherits System.Web.UI.Page
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
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    Dim wUserName As String = ""            '姓名代理用
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        SetPopupFunction()                          '設定彈出視窗事件
        ShowSheetFunction()                         '表單功能按鈕/欄位顯示
        '
        If Not IsPostBack Then                      'PostBack
            ShowSheetField("New")                   '表單欄位顯示及欄位輸入檢查
            If wFormSno > 0 And wStep > 3 Then      '起單/簽核
                UpdateTranFile()                    '更新交易資料
                ShowFormData()                      '顯示表單資料
            End If
        Else
            ShowSheetField("IsPostBack")              '表單欄位顯示及欄位輸入檢查
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '-----------------------------------------------------------------
        '-- 系統參數
        '-----------------------------------------------------------------
        Server.ScriptTimeout = 900                                                                  '設定逾時時間
        Response.Cookies("PGM").Value = "CommissionSheet_01.aspx"                        '程式名
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")                      '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")                            '工程代碼
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")                                           '現在日時
        wFormNo = Request.QueryString("pFormNo")                                                    '表單號碼
        wFormSno = Request.QueryString("pFormSno")                                                  '表單流水號
        wStep = Request.QueryString("pStep")                                                        '工程代碼
        wSeqNo = Request.QueryString("pSeqNo")                                                      '序號
        wApplyID = Request.QueryString("pApplyID")                                                  '申請者ID
        wAgentID = Request.QueryString("pAgentID")                                                  '被代理人ID
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)                                         '申請者行事曆類別
        wDecideCalendar = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))                 '簽核者行事曆類別
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        BEXPDEL.Attributes("onclick") = "OpenDatePicker('" + wDecideCalendar + "','DEXPDEL');"      '希望交期-選擇輸入
        BREFNO.Attributes("onclick") = "window.open('OpenNoPicker.aspx','NoPicker','status=0,toolbar=0,width=620,height=650,resizable=yes,scrollbars=yes');"    '選取參考開發資料
        DReadHistory.Attributes.Add("onclick", "ReadHistory('01')")                                 '閱讀核定履歷
        DOP40.Attributes.Add("onclick", "GoBackOP('40')")                                           '回試作工程CheckBox
        DOP50.Attributes.Add("onclick", "GoBackOP('50')")                                           '回試作工程CheckBox
        DOP60.Attributes.Add("onclick", "GoBackOP('60')")                                           '回試作工程CheckBox
        DOP70.Attributes.Add("onclick", "GoBackOP('70')")                                           '回試作工程CheckBox
        DOP80.Attributes.Add("onclick", "GoBackOP('80')")                                           '回試作工程CheckBox
        DOP90.Attributes.Add("onclick", "GoBackOP('90')")                                           '回試作工程CheckBox
        DOP100.Attributes.Add("onclick", "GoBackOP('100')")                                         '回試作工程CheckBox
        DOP110.Attributes.Add("onclick", "GoBackOP('110')")                                         '回試作工程CheckBox
        DNeedSample.Attributes.Add("onclick", "NeedSample('SAMPLE')")                               '已開發-需樣品CheckBox
        DNeedItemRegister.Attributes.Add("onclick", "NeedSample('ITEM')")                           '已開發-需登錄CheckBox
        DECOL.Attributes.Add("onclick", "SetInputDataAttr('ECOL')")                                 '鏈齒屬行設定
        DCCOL.Attributes.Add("onclick", "SetInputDataAttr('CCOL')")                                 '丸扭屬行設定
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
    '**     表單欄位顯示及欄位輸入檢查
    '**
    '*****************************************************************
    Sub ShowSheetField(ByVal pPost As String)
        oCommon.GetFieldAttribute(wFormNo, wStep, FieldName, Attribute)    '取得欄位及屬性陣列(表單號碼,工程關卡號碼,欄位名陣列,欄位屬性陣列)
        '2408
        SetFieldAttribute(pPost)                                           '表單各欄位屬性及欄位輸入檢查等設定
    End Sub
    '*****************************************************************
    '**(SetFieldAttribute)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '-----------------------------------------------------------------
        '-- 系統控制
        '-----------------------------------------------------------------
        If pPost = "New" Then
            '自動設定閱讀履歷(適用 = 廠長)  Modify-2011/10/5
            If UCase(Request.QueryString("pUserID")) = "FADM02" Then
                DReadHistory.Checked = True
            Else
                DReadHistory.Checked = False
            End If
        End If
        '-----------------------------------------------------------------
        '-- 開發委託
        '-----------------------------------------------------------------
        '----基本欄位設定-------------------------------------------------    
        'No
        Select Case FindFieldInf("1-NO")
            Case 0  '顯示
                DNO.BackColor = Color.LightGray
                DNO.Visible = True
                DNO.Attributes.Add("readonly", "true")
                '
                LREFNO.Visible = True
                BREFNO.Visible = False
                DREFNO.Style("top") = -100 & "px"
                DREFNO.BackColor = Color.LightGray
                DREFNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DNO.Visible = True
                DNO.BackColor = Color.GreenYellow
                DNO.ReadOnly = False
                ShowRequiredFieldValidator("DNORqd", "DNO", "異常：需輸入ＮＯ")
                '
                LREFNO.Visible = False
                BREFNO.Visible = True
                DREFNO.Visible = True
                DREFNO.BackColor = Color.LightGray
                DREFNO.Attributes.Add("readonly", "true")
            Case 2  '修改
                DNO.Visible = True
                DNO.BackColor = Color.Yellow
                DNO.ReadOnly = False
                '
                LREFNO.Visible = False
                BREFNO.Visible = True
                DREFNO.Visible = True
                DREFNO.BackColor = Color.LightGray
                DREFNO.Attributes.Add("readonly", "true")
            Case Else   '隱藏
                DNO.Visible = False
                '
                LREFNO.Visible = False
                BREFNO.Visible = False
                DREFNO.Style("top") = -100 & "px"
                DREFNO.BackColor = Color.LightGray
                DREFNO.Attributes.Add("readonly", "true")
        End Select
        If pPost = "New" Then DNO.Text = ""
        'BUYER
        Select Case FindFieldInf("1-BASE")
            Case 0  '顯示
                DAPPBUYER.BackColor = Color.LightGray
                DAPPBUYER.Visible = True
            Case 1  '修改+檢查
                DAPPBUYER.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAPPBUYERRqd", "DAPPBUYER", "異常：需輸入Buyer")
                DAPPBUYER.Visible = True
            Case 2  '修改
                DAPPBUYER.BackColor = Color.Yellow
                DAPPBUYER.Visible = True
            Case Else   '隱藏
                DAPPBUYER.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-BASE"), "APPBUYER", "ZZZZZZ")
        '委託廠商
        Select Case FindFieldInf("1-BASE")
            Case 0  '顯示
                DSellVendor.BackColor = Color.LightGray
                DSellVendor.Visible = True
                DSellVendor.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSellVendor.Visible = True
                DSellVendor.BackColor = Color.GreenYellow
                DSellVendor.ReadOnly = False
                ShowRequiredFieldValidator("DSELLVENDORRqd", "DSELLVENDOR", "異常：需輸入委託廠商")
            Case 2  '修改
                DSellVendor.Visible = True
                DSellVendor.BackColor = Color.Yellow
                DSellVendor.ReadOnly = False
            Case Else   '隱藏
                DSellVendor.Visible = False
        End Select
        If pPost = "New" Then DSellVendor.Text = ""
        '預估量
        Select Case FindFieldInf("1-BASE")
            Case 0  '顯示
                DESYQTY.BackColor = Color.LightGray
                DESYQTY.Visible = True
                DESYQTY.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DESYQTY.Visible = True
                DESYQTY.BackColor = Color.GreenYellow
                DESYQTY.ReadOnly = False
                ShowRequiredFieldValidator("DESYQTYRqd", "DESYQTY", "異常：需輸入預估量")
            Case 2  '修改
                DESYQTY.Visible = True
                DESYQTY.BackColor = Color.Yellow
                DESYQTY.ReadOnly = False
            Case Else   '隱藏
                DESYQTY.Visible = False
        End Select
        If pPost = "New" Then DESYQTY.Text = ""
        '希望交期
        Select Case FindFieldInf("1-EXPDEL")
            Case 0  '顯示
                DEXPDEL.BackColor = Color.LightGray
                DEXPDEL.ReadOnly = True
                DEXPDEL.Visible = True
                BEXPDEL.Visible = False
            Case 1  '修改+檢查
                DEXPDEL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEXPDELRqd", "DEXPDEL", "異常：需輸入希望交期")
                DEXPDEL.Visible = True
                BEXPDEL.Visible = True
            Case 2  '修改
                DEXPDEL.BackColor = Color.Yellow
                DEXPDEL.Visible = True
                BEXPDEL.Visible = True
            Case Else   '隱藏
                DEXPDEL.Visible = False
                BEXPDEL.Visible = False
        End Select
        If pPost = "New" Then DEXPDEL.Text = ""
        '客戶ITEM
        Select Case FindFieldInf("1-BASE")
            Case 0  '顯示
                DCUSTITEM.BackColor = Color.LightGray
                DCUSTITEM.Visible = True
                DCUSTITEM.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCUSTITEM.Visible = True
                DCUSTITEM.BackColor = Color.GreenYellow
                DCUSTITEM.ReadOnly = False
                ShowRequiredFieldValidator("DCUSTITEMRqd", "DCUSTITEM", "異常：需輸入客戶ITEM")
            Case 2  '修改
                DCUSTITEM.Visible = True
                DCUSTITEM.BackColor = Color.Yellow
                DCUSTITEM.ReadOnly = False
            Case Else   '隱藏
                DCUSTITEM.Visible = False
        End Select
        If pPost = "New" Then DCUSTITEM.Text = ""
        '用途
        Select Case FindFieldInf("1-BASE")
            Case 0  '顯示
                DUSAGE.BackColor = Color.LightGray
                DUSAGE.Visible = True
            Case 1  '修改+檢查
                DUSAGE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DUSAGERqd", "DUSAGE", "異常：需輸入用途")
                DUSAGE.Visible = True
            Case 2  '修改
                DUSAGE.BackColor = Color.Yellow
                DUSAGE.Visible = True
            Case Else   '隱藏
                DUSAGE.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-BASE"), "USAGE", "ZZZZZZ")
        'OR-NO
        Select Case FindFieldInf("1-BASE")
            Case 0  '顯示
                DORNO.BackColor = Color.LightGray
                DORNO.Visible = True
                DORNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DORNO.Visible = True
                DORNO.BackColor = Color.GreenYellow
                DORNO.ReadOnly = False
                ShowRequiredFieldValidator("DORNORqd", "DORNO", "異常：需輸入OR-NO")
            Case 2  '修改
                DORNO.Visible = True
                DORNO.BackColor = Color.Yellow
                DORNO.ReadOnly = False
            Case Else   '隱藏
                DORNO.Visible = False
        End Select
        If pPost = "New" Then DORNO.Text = ""
        '需圖面
        Select Case FindFieldInf("1-BASE")
            Case 0  '顯示
                DNEEDMAP.BackColor = Color.LightGray
                DNEEDMAP.Visible = True
            Case 1  '修改+檢查
                DNEEDMAP.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNEEDMAPRqd", "DNEEDMAP", "異常：需輸入需圖面")
                DNEEDMAP.Visible = True
            Case 2  '修改
                DNEEDMAP.BackColor = Color.Yellow
                DNEEDMAP.Visible = True
            Case Else   '隱藏
                DNEEDMAP.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-BASE"), "NEEDMAP", "ZZZZZZ")
        '草圖
        Select Case FindFieldInf("1-BASE")
            Case 0  '顯示
                DMAPREFFILE.Visible = False
                DMAPREFFILE.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DMAPREFFILERqd", "DMAPREFFILE", "異常：需輸入草圖")
                DMAPREFFILE.Visible = True
                DMAPREFFILE.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DMAPREFFILE.Visible = True
                DMAPREFFILE.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DMAPREFFILE.Visible = False
        End Select
        If pPost = "New" Then LMAPREFFILE.Visible = False
        '----樣品欄位設定-------------------------------------------------    
        '需樣品/需登錄
        Select Case FindFieldInf("1-SAMPLE")
            Case 0  '顯示
                DNeedSample.Visible = True
                DNeedItemRegister.Visible = True
                DModify.Checked = False     '需樣品/需登錄-修改可否
            Case 1  '修改+檢查
                DNeedSample.Visible = True
                DNeedSample.Enabled = True
                DNeedItemRegister.Visible = True
                DNeedItemRegister.Enabled = True
                DModify.Checked = True     '需樣品/需登錄-修改可否
            Case 2  '修改
                DNeedSample.Visible = True
                DNeedSample.Enabled = True
                DNeedItemRegister.Visible = True
                DNeedItemRegister.Enabled = True
                DModify.Checked = True     '需樣品/需登錄-修改可否
            Case Else   '隱藏
                DNeedSample.Visible = False
                DNeedItemRegister.Visible = False
                DModify.Checked = False     '需樣品/需登錄-修改可否
        End Select
        '製品區分
        Select Case FindFieldInf("1-SAMPLE")
            Case 0  '顯示
                DPRO.BackColor = Color.LightGray
                DPRO.Visible = True
            Case 1  '修改+檢查
                DPRO.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPRORqd", "DPRO", "異常：需輸入製品區分")
                DPRO.Visible = True
            Case 2  '修改
                DPRO.BackColor = Color.Yellow
                DPRO.Visible = True
            Case Else   '隱藏
                DPRO.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-SAMPLE"), "PRO", "ZZZZZZ")
        '開具(色)
        Select Case FindFieldInf("1-SAMPLE")
            Case 0  '顯示
                DOPPART.BackColor = Color.LightGray
                DOPPART.Visible = True
                DOPPART.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DOPPART.Visible = True
                DOPPART.BackColor = Color.GreenYellow
                DOPPART.ReadOnly = False
                ShowRequiredFieldValidator("DOPPARTRqd", "DOPPART", "異常：需輸入開具(色)")
            Case 2  '修改
                DOPPART.Visible = True
                DOPPART.BackColor = Color.Yellow
                DOPPART.ReadOnly = False
            Case Else   '隱藏
                DOPPART.Visible = False
        End Select
        If pPost = "New" Then DOPPART.Text = ""
        '長度(企)
        Select Case FindFieldInf("1-SAMPLE-P")
            Case 0  '顯示
                DPLEN.BackColor = Color.LightGray
                DPLEN.Visible = True
                DPLEN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DPLEN.Visible = True
                DPLEN.BackColor = Color.GreenYellow
                DPLEN.ReadOnly = False
                ShowRequiredFieldValidator("DPLENRqd", "DPLEN", "異常：需輸入長度(企)")
            Case 2  '修改
                DPLEN.Visible = True
                DPLEN.BackColor = Color.Yellow
                DPLEN.ReadOnly = False
            Case Else   '隱藏
                DPLEN.Visible = False
        End Select
        If pPost = "New" Then DPLEN.Text = "0"
        '長度單位(企)
        Select Case FindFieldInf("1-SAMPLE-P")
            Case 0  '顯示
                DPLENUN.BackColor = Color.LightGray
                DPLENUN.Visible = True
            Case 1  '修改+檢查
                DPLENUN.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPLENUNRqd", "DPLENUN", "異常：需輸入長度單位(企)")
                DPLENUN.Visible = True
            Case 2  '修改
                DPLENUN.BackColor = Color.Yellow
                DPLENUN.Visible = True
            Case Else   '隱藏
                DPLENUN.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-SAMPLE"), "PLENUN", "ZZZZZZ")
        '數量(企)
        Select Case FindFieldInf("1-SAMPLE-P")
            Case 0  '顯示
                DPQTY.BackColor = Color.LightGray
                DPQTY.Visible = True
                DPQTY.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DPQTY.Visible = True
                DPQTY.BackColor = Color.GreenYellow
                DPQTY.ReadOnly = False
                ShowRequiredFieldValidator("DPQTYRqd", "DPQTY", "異常：需輸入數量(企)")
            Case 2  '修改
                DPQTY.Visible = True
                DPQTY.BackColor = Color.Yellow
                DPQTY.ReadOnly = False
            Case Else   '隱藏
                DPQTY.Visible = False
        End Select
        If pPost = "New" Then DPQTY.Text = "0"
        '數量單位(企)
        Select Case FindFieldInf("1-SAMPLE-P")
            Case 0  '顯示
                DPQTYUN.BackColor = Color.LightGray
                DPQTYUN.Visible = True
            Case 1  '修改+檢查
                DPQTYUN.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPQTYUNRqd", "DPQTYUN", "異常：需輸入數量單位(企)")
                DPQTYUN.Visible = True
            Case 2  '修改
                DPQTYUN.BackColor = Color.Yellow
                DPQTYUN.Visible = True
            Case Else   '隱藏
                DPQTYUN.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-SAMPLE"), "PQTYUN", "ZZZZZZ")
        '長度(EA)
        Select Case FindFieldInf("1-SAMPLE-EA")
            Case 0  '顯示
                DEALEN.BackColor = Color.LightGray
                DEALEN.Visible = True
                DEALEN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DEALEN.Visible = True
                DEALEN.BackColor = Color.GreenYellow
                DEALEN.ReadOnly = False
                ShowRequiredFieldValidator("DEALENRqd", "DEALEN", "異常：需輸入長度(EA)")
            Case 2  '修改
                DEALEN.Visible = True
                DEALEN.BackColor = Color.Yellow
                DEALEN.ReadOnly = False
            Case Else   '隱藏
                DEALEN.Visible = False
        End Select
        If pPost = "New" Then DEALEN.Text = "0"
        '長度單位(EA)
        Select Case FindFieldInf("1-SAMPLE-EA")
            Case 0  '顯示
                DEALENUN.BackColor = Color.LightGray
                DEALENUN.Visible = True
            Case 1  '修改+檢查
                DEALENUN.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEALENUNRqd", "DEALENUN", "異常：需輸入長度單位(EA)")
                DEALENUN.Visible = True
            Case 2  '修改
                DEALENUN.BackColor = Color.Yellow
                DEALENUN.Visible = True
            Case Else   '隱藏
                DEALENUN.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-SAMPLE-EA"), "EALENUN", "ZZZZZZ")
        '數量(EA)
        Select Case FindFieldInf("1-SAMPLE-EA")
            Case 0  '顯示
                DEAQTY.BackColor = Color.LightGray
                DEAQTY.Visible = True
                DEAQTY.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DEAQTY.Visible = True
                DEAQTY.BackColor = Color.GreenYellow
                DEAQTY.ReadOnly = False
                ShowRequiredFieldValidator("DEAQTYRqd", "DEAQTY", "異常：需輸入數量(EA)")
            Case 2  '修改
                DEAQTY.Visible = True
                DEAQTY.BackColor = Color.Yellow
                DEAQTY.ReadOnly = False
            Case Else   '隱藏
                DEAQTY.Visible = False
        End Select
        If pPost = "New" Then DEAQTY.Text = "0"
        '數量單位(EA)
        Select Case FindFieldInf("1-SAMPLE-EA")
            Case 0  '顯示
                DEAQTYUN.BackColor = Color.LightGray
                DEAQTYUN.Visible = True
            Case 1  '修改+檢查
                DEAQTYUN.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEAQTYUNRqd", "DEAQTYUN", "異常：需輸入數量單位(EA)")
                DEAQTYUN.Visible = True
            Case 2  '修改
                DEAQTYUN.BackColor = Color.Yellow
                DEAQTYUN.Visible = True
            Case Else   '隱藏
                DEAQTYUN.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-SAMPLE-EA"), "EAQTYUN", "ZZZZZZ")
        '拉頭(上)
        Select Case FindFieldInf("1-SAMPLE")
            Case 0  '顯示
                DUPSLI.BackColor = Color.LightGray
                DUPSLI.Visible = True
                DUPSLI.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DUPSLI.Visible = True
                DUPSLI.BackColor = Color.GreenYellow
                DUPSLI.ReadOnly = False
                ShowRequiredFieldValidator("DUPSLIRqd", "DUPSLI", "異常：需輸入拉頭(上)")
            Case 2  '修改
                DUPSLI.Visible = True
                DUPSLI.BackColor = Color.Yellow
                DUPSLI.ReadOnly = False
            Case Else   '隱藏
                DUPSLI.Visible = False
        End Select
        If pPost = "New" Then DUPSLI.Text = ""
        '拉頭(下)
        Select Case FindFieldInf("1-SAMPLE")
            Case 0  '顯示
                DLOSLI.BackColor = Color.LightGray
                DLOSLI.Visible = True
                DLOSLI.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DLOSLI.Visible = True
                DLOSLI.BackColor = Color.GreenYellow
                DLOSLI.ReadOnly = False
                ShowRequiredFieldValidator("DLOSLIRqd", "DLOSLI", "異常：需輸入拉頭(下)")
            Case 2  '修改
                DLOSLI.Visible = True
                DLOSLI.BackColor = Color.Yellow
                DLOSLI.ReadOnly = False
            Case Else   '隱藏
                DLOSLI.Visible = False
        End Select
        If pPost = "New" Then DLOSLI.Text = ""
        '表面處理(上)
        Select Case FindFieldInf("1-SAMPLE")
            Case 0  '顯示
                DUPFIN.BackColor = Color.LightGray
                DUPFIN.Visible = True
                DUPFIN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DUPFIN.Visible = True
                DUPFIN.BackColor = Color.GreenYellow
                DUPFIN.ReadOnly = False
                ShowRequiredFieldValidator("DUPFINRqd", "DUPFIN", "異常：需輸入表面處理(上)")
            Case 2  '修改
                DUPFIN.Visible = True
                DUPFIN.BackColor = Color.Yellow
                DUPFIN.ReadOnly = False
            Case Else   '隱藏
                DUPFIN.Visible = False
        End Select
        If pPost = "New" Then DUPFIN.Text = ""
        '表面處理(下)
        Select Case FindFieldInf("1-SAMPLE")
            Case 0  '顯示
                DLOFIN.BackColor = Color.LightGray
                DLOFIN.Visible = True
                DLOFIN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DLOFIN.Visible = True
                DLOFIN.BackColor = Color.GreenYellow
                DLOFIN.ReadOnly = False
                ShowRequiredFieldValidator("DLOFINRqd", "DLOFIN", "異常：需輸入表面處理(下)")
            Case 2  '修改
                DLOFIN.Visible = True
                DLOFIN.BackColor = Color.Yellow
                DLOFIN.ReadOnly = False
            Case Else   '隱藏
                DLOFIN.Visible = False
        End Select
        If pPost = "New" Then DLOFIN.Text = ""
        '上止種類
        Select Case FindFieldInf("1-SAMPLE")
            Case 0  '顯示
                DUPSTK.BackColor = Color.LightGray
                DUPSTK.Visible = True
                DUPSTK.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DUPSTK.Visible = True
                DUPSTK.BackColor = Color.GreenYellow
                DUPSTK.ReadOnly = False
                ShowRequiredFieldValidator("DUPSTKRqd", "DUPSTK", "異常：需輸入上止種類")
            Case 2  '修改
                DUPSTK.Visible = True
                DUPSTK.BackColor = Color.Yellow
                DUPSTK.ReadOnly = False
            Case Else   '隱藏
                DUPSTK.Visible = False
        End Select
        If pPost = "New" Then DUPSTK.Text = ""
        '下止種類
        Select Case FindFieldInf("1-SAMPLE")
            Case 0  '顯示
                DLOSTK.BackColor = Color.LightGray
                DLOSTK.Visible = True
                DLOSTK.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DLOSTK.Visible = True
                DLOSTK.BackColor = Color.GreenYellow
                DLOSTK.ReadOnly = False
                ShowRequiredFieldValidator("DLOSTKRqd", "DLOSTK", "異常：需輸入下止種類")
            Case 2  '修改
                DLOSTK.Visible = True
                DLOSTK.BackColor = Color.Yellow
                DLOSTK.ReadOnly = False
            Case Else   '隱藏
                DLOSTK.Visible = False
        End Select
        If pPost = "New" Then DLOSTK.Text = ""
        '特殊規格
        Select Case FindFieldInf("1-SAMPLE")
            Case 0  '顯示
                DSPSPEC.BackColor = Color.LightGray
                DSPSPEC.Visible = True
                DSPSPEC.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSPSPEC.Visible = True
                DSPSPEC.BackColor = Color.GreenYellow
                DSPSPEC.ReadOnly = False
                ShowRequiredFieldValidator("DSPSPECRqd", "DSPSPEC", "異常：需輸入特殊規格")
            Case 2  '修改
                DSPSPEC.Visible = True
                DSPSPEC.BackColor = Color.Yellow
                DSPSPEC.ReadOnly = False
            Case Else   '隱藏
                DSPSPEC.Visible = False
        End Select
        If pPost = "New" Then DSPSPEC.Text = ""
        '----開發規格欄位設定-------------------------------------------------    
        '型別
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DSIZENO.BackColor = Color.LightGray
                DSIZENO.Visible = True
                BGetYColor.Visible = False
            Case 1  '修改+檢查
                DSIZENO.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSIZENORqd", "DSIZENO", "異常：需輸入型別")
                DSIZENO.Visible = True
                BGetYColor.Visible = True
            Case 2  '修改
                DSIZENO.BackColor = Color.Yellow
                DSIZENO.Visible = True
                BGetYColor.Visible = True
            Case Else   '隱藏
                DSIZENO.Visible = False
                BGetYColor.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "SIZENO", "ZZZZZZ")
        '鏈條型式
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DITEM.BackColor = Color.LightGray
                DITEM.Visible = True
            Case 1  '修改+檢查
                DITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DITEMRqd", "DITEM", "異常：需輸入鏈條型式")
                DITEM.Visible = True
            Case 2  '修改
                DITEM.BackColor = Color.Yellow
                DITEM.Visible = True
            Case Else   '隱藏
                DITEM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "ITEM", "ZZZZZZ")
        '布帶
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTATYPE.BackColor = Color.LightGray
                DTATYPE.Visible = True
            Case 1  '修改+檢查
                DTATYPE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTATYPERqd", "DTATYPE", "異常：需輸入布帶")
                DTATYPE.Visible = True
            Case 2  '修改
                DTATYPE.BackColor = Color.Yellow
                DTATYPE.Visible = True
            Case Else   '隱藏
                DTATYPE.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "TATYPE", "ZZZZZZ")
        '布帶寬度
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTAWIDTH.BackColor = Color.LightGray
                DTAWIDTH.Visible = True
                DTAWIDTH.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTAWIDTH.Visible = True
                DTAWIDTH.BackColor = Color.GreenYellow
                DTAWIDTH.ReadOnly = False
                ShowRequiredFieldValidator("DTAWIDTHRqd", "DTAWIDTH", "異常：需輸入布帶寬度")
            Case 2  '修改
                DTAWIDTH.Visible = True
                DTAWIDTH.BackColor = Color.Yellow
                DTAWIDTH.ReadOnly = False
            Case Else   '隱藏
                DTAWIDTH.Visible = False
        End Select
        If pPost = "New" Then DTAWIDTH.Text = "0"
        '鏈齒顏色-SEL
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DECOLSEL.BackColor = Color.LightGray
                DECOLSEL.Visible = True
            Case 1  '修改+檢查
                DECOLSEL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DECOLSELRqd", "DECOLSEL", "異常：需輸入鏈齒顏色")
                DECOLSEL.Visible = True
            Case 2  '修改
                DECOLSEL.BackColor = Color.Yellow
                DECOLSEL.Visible = True
            Case Else   '隱藏
                DECOLSEL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "ECOLSEL", "ZZZZZZ")
        '鏈齒顏色
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DECOL.BackColor = Color.LightGray
                DECOL.Visible = True
                DECOL.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DECOL.Visible = True
                DECOL.BackColor = Color.GreenYellow
                DECOL.ReadOnly = False
                ShowRequiredFieldValidator("DECOLRqd", "DECOL", "異常：需輸入鏈齒顏色")
            Case 2  '修改
                DECOL.Visible = True
                DECOL.BackColor = Color.Yellow
                DECOL.ReadOnly = False
            Case Else   '隱藏
                DECOL.Visible = False
        End Select
        If pPost = "New" Then DECOL.Text = ""
        '丸紐-SEL
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DCCOLSEL.BackColor = Color.LightGray
                DCCOLSEL.Visible = True
            Case 1  '修改+檢查
                DCCOLSEL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCCOLSELRqd", "DCCOLSEL", "異常：需輸入丸紐")
                DCCOLSEL.Visible = True
            Case 2  '修改
                DCCOLSEL.BackColor = Color.Yellow
                DCCOLSEL.Visible = True
            Case Else   '隱藏
                DCCOLSEL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "CCOLSEL", "ZZZZZZ")
        '丸紐
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DCCOL.BackColor = Color.LightGray
                DCCOL.Visible = True
                DCCOL.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCCOL.Visible = True
                DCCOL.BackColor = Color.GreenYellow
                DCCOL.ReadOnly = False
                ShowRequiredFieldValidator("DCCOLRqd", "DCCOL", "異常：需輸入丸紐")
            Case 2  '修改
                DCCOL.Visible = True
                DCCOL.BackColor = Color.Yellow
                DCCOL.ReadOnly = False
            Case Else   '隱藏
                DCCOL.Visible = False
        End Select
        If pPost = "New" Then DCCOL.Text = ""
        '布帶-色番(同)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTACOL.BackColor = Color.LightGray
                DTACOL.Visible = True
            Case 1  '修改+檢查
                DTACOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTACOLRqd", "DTACOL", "異常：需輸入布帶-色番(同)")
                DTACOL.Visible = True
            Case 2  '修改
                DTACOL.BackColor = Color.Yellow
                DTACOL.Visible = True
            Case Else   '隱藏
                DTACOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "TACOL", "ZZZZZZ")
        '布帶-色番(同)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTACOLNO.BackColor = Color.LightGray
                DTACOLNO.Visible = True
                DTACOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTACOLNO.Visible = True
                DTACOLNO.BackColor = Color.GreenYellow
                DTACOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTACOLNORqd", "DTACOLNO", "異常：需輸入布帶-色番(同)")
            Case 2  '修改
                DTACOLNO.Visible = True
                DTACOLNO.BackColor = Color.Yellow
                DTACOLNO.ReadOnly = False
            Case Else   '隱藏
                DTACOLNO.Visible = False
        End Select
        If pPost = "New" Then DTACOLNO.Text = ""
        '布帶-YKK(同)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTAYCOLNO.BackColor = Color.LightGray
                DTAYCOLNO.Visible = True
                DTAYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTAYCOLNO.Visible = True
                DTAYCOLNO.BackColor = Color.GreenYellow
                DTAYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTAYCOLNORqd", "DTAYCOLNO", "異常：需輸入布帶-YKK(同)")
            Case 2  '修改
                DTAYCOLNO.Visible = True
                DTAYCOLNO.BackColor = Color.Yellow
                DTAYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTAYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTAYCOLNO.Text = ""
        '布帶-色番(左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTALCOL.BackColor = Color.LightGray
                DTALCOL.Visible = True
            Case 1  '修改+檢查
                DTALCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTALCOLRqd", "DTALCOL", "異常：需輸入布帶-色番(左)")
                DTALCOL.Visible = True
            Case 2  '修改
                DTALCOL.BackColor = Color.Yellow
                DTALCOL.Visible = True
            Case Else   '隱藏
                DTALCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "TALCOL", "ZZZZZZ")
        '布帶-色番(左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTALCOLNO.BackColor = Color.LightGray
                DTALCOLNO.Visible = True
                DTALCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTALCOLNO.Visible = True
                DTALCOLNO.BackColor = Color.GreenYellow
                DTALCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTALCOLNORqd", "DTALCOLNO", "異常：需輸入布帶-色番(左)")
            Case 2  '修改
                DTALCOLNO.Visible = True
                DTALCOLNO.BackColor = Color.Yellow
                DTALCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTALCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTALCOLNO.Text = ""
        '布帶-YKK(左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTALYCOLNO.BackColor = Color.LightGray
                DTALYCOLNO.Visible = True
                DTALYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTALYCOLNO.Visible = True
                DTALYCOLNO.BackColor = Color.GreenYellow
                DTALYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTALYCOLNORqd", "DTALYCOLNO", "異常：需輸入布帶-YKK(左)")
            Case 2  '修改
                DTALYCOLNO.Visible = True
                DTALYCOLNO.BackColor = Color.Yellow
                DTALYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTALYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTALYCOLNO.Text = ""
        '布帶-色番(右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTARCOL.BackColor = Color.LightGray
                DTARCOL.Visible = True
            Case 1  '修改+檢查
                DTARCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTARCOLRqd", "DTARCOL", "異常：需輸入布帶-色番(右)")
                DTARCOL.Visible = True
            Case 2  '修改
                DTARCOL.BackColor = Color.Yellow
                DTARCOL.Visible = True
            Case Else   '隱藏
                DTARCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "TARCOL", "ZZZZZZ")
        '布帶-色番(右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTARCOLNO.BackColor = Color.LightGray
                DTARCOLNO.Visible = True
                DTARCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTARCOLNO.Visible = True
                DTARCOLNO.BackColor = Color.GreenYellow
                DTARCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTARCOLNORqd", "DTARCOLNO", "異常：需輸入布帶-色番(右)")
            Case 2  '修改
                DTARCOLNO.Visible = True
                DTARCOLNO.BackColor = Color.Yellow
                DTARCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTARCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTARCOLNO.Text = ""
        '布帶-YKK(右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTARYCOLNO.BackColor = Color.LightGray
                DTARYCOLNO.Visible = True
                DTARYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTARYCOLNO.Visible = True
                DTARYCOLNO.BackColor = Color.GreenYellow
                DTARYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTARYCOLNORqd", "DTARYCOLNO", "異常：需輸入布帶-YKK(右)")
            Case 2  '修改
                DTARYCOLNO.Visible = True
                DTARYCOLNO.BackColor = Color.Yellow
                DTARYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTARYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTARYCOLNO.Text = ""
        '縫上-色番(同)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHUPCOL.BackColor = Color.LightGray
                DTHUPCOL.Visible = True
            Case 1  '修改+檢查
                DTHUPCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTHUPCOLRqd", "DTHUPCOL", "異常：需輸入縫上-色番(同)")
                DTHUPCOL.Visible = True
            Case 2  '修改
                DTHUPCOL.BackColor = Color.Yellow
                DTHUPCOL.Visible = True
            Case Else   '隱藏
                DTHUPCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "THUPCOL", "ZZZZZZ")
        '縫上-色番(同)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHUPCOLNO.BackColor = Color.LightGray
                DTHUPCOLNO.Visible = True
                DTHUPCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHUPCOLNO.Visible = True
                DTHUPCOLNO.BackColor = Color.GreenYellow
                DTHUPCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHUPCOLNORqd", "DTHUPCOLNO", "異常：需輸入縫上-色番(同)")
            Case 2  '修改
                DTHUPCOLNO.Visible = True
                DTHUPCOLNO.BackColor = Color.Yellow
                DTHUPCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHUPCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHUPCOLNO.Text = ""
        '縫上-YKK(同)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHUPYCOLNO.BackColor = Color.LightGray
                DTHUPYCOLNO.Visible = True
                DTHUPYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHUPYCOLNO.Visible = True
                DTHUPYCOLNO.BackColor = Color.GreenYellow
                DTHUPYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHUPYCOLNORqd", "DTHUPYCOLNO", "異常：需輸入縫上-YKK(同)")
            Case 2  '修改
                DTHUPYCOLNO.Visible = True
                DTHUPYCOLNO.BackColor = Color.Yellow
                DTHUPYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHUPYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHUPYCOLNO.Text = ""
        '縫上-色番(左左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLUPCOL.BackColor = Color.LightGray
                DTHLUPCOL.Visible = True
            Case 1  '修改+檢查
                DTHLUPCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTHLUPCOLRqd", "DTHLUPCOL", "異常：需輸入縫上-色番(左左)")
                DTHLUPCOL.Visible = True
            Case 2  '修改
                DTHLUPCOL.BackColor = Color.Yellow
                DTHLUPCOL.Visible = True
            Case Else   '隱藏
                DTHLUPCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "THLUPCOL", "ZZZZZZ")
        '縫上-色番(左左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLUPCOLNO.BackColor = Color.LightGray
                DTHLUPCOLNO.Visible = True
                DTHLUPCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHLUPCOLNO.Visible = True
                DTHLUPCOLNO.BackColor = Color.GreenYellow
                DTHLUPCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHLUPCOLNORqd", "DTHLUPCOLNO", "異常：需輸入縫上-色番(左左)")
            Case 2  '修改
                DTHLUPCOLNO.Visible = True
                DTHLUPCOLNO.BackColor = Color.Yellow
                DTHLUPCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHLUPCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHLUPCOLNO.Text = ""
        '縫上-YKK(左左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLUPYCOLNO.BackColor = Color.LightGray
                DTHLUPYCOLNO.Visible = True
                DTHLUPYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHLUPYCOLNO.Visible = True
                DTHLUPYCOLNO.BackColor = Color.GreenYellow
                DTHLUPYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHLUPYCOLNORqd", "DTHLUPYCOLNO", "異常：需輸入縫上-YKK(左左)")
            Case 2  '修改
                DTHLUPYCOLNO.Visible = True
                DTHLUPYCOLNO.BackColor = Color.Yellow
                DTHLUPYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHLUPYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHLUPYCOLNO.Text = ""

        '縫上-色番(左右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLRUPCOL.BackColor = Color.LightGray
                DTHLRUPCOL.Visible = True
            Case 1  '修改+檢查
                DTHLRUPCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTHLRUPCOLRqd", "DTHLRUPCOL", "異常：需輸入縫上-色番(左右)")
                DTHLUPCOL.Visible = True
            Case 2  '修改
                DTHLRUPCOL.BackColor = Color.Yellow
                DTHLRUPCOL.Visible = True
            Case Else   '隱藏
                DTHLUPCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "THLRUPCOL", "ZZZZZZ")
        '縫上-色番(左右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLRUPCOLNO.BackColor = Color.LightGray
                DTHLRUPCOLNO.Visible = True
                DTHLRUPCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHLRUPCOLNO.Visible = True
                DTHLRUPCOLNO.BackColor = Color.GreenYellow
                DTHLRUPCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHLRUPCOLNORqd", "DTHLRUPCOLNO", "異常：需輸入縫上-色番(左右)")
            Case 2  '修改
                DTHLRUPCOLNO.Visible = True
                DTHLRUPCOLNO.BackColor = Color.Yellow
                DTHLRUPCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHLRUPCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHLRUPCOLNO.Text = ""
        '縫上-YKK(SR 左右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLRUPYCOLNO.BackColor = Color.LightGray
                DTHLRUPYCOLNO.Visible = True
                DTHLRUPYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHLRUPYCOLNO.Visible = True
                DTHLRUPYCOLNO.BackColor = Color.GreenYellow
                DTHLRUPYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHLRUPYCOLNORqd", "DTHLRUPYCOLNO", "異常：需輸入縫上-YKK(左右)")
            Case 2  '修改
                DTHLRUPYCOLNO.Visible = True
                DTHLRUPYCOLNO.BackColor = Color.Yellow
                DTHLRUPYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHLRUPYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHLUPYCOLNO.Text = ""

        '縫上-色番(右左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRUPCOL.BackColor = Color.LightGray
                DTHRUPCOL.Visible = True
            Case 1  '修改+檢查
                DTHRUPCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTHRUPCOLRqd", "DTHRUPCOL", "異常：需輸入縫上-色番(右左)")
                DTHRUPCOL.Visible = True
            Case 2  '修改
                DTHRUPCOL.BackColor = Color.Yellow
                DTHRUPCOL.Visible = True
            Case Else   '隱藏
                DTHRUPCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "THRUPCOL", "ZZZZZZ")
        '縫上-色番(右左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRUPCOLNO.BackColor = Color.LightGray
                DTHRUPCOLNO.Visible = True
                DTHRUPCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHRUPCOLNO.Visible = True
                DTHRUPCOLNO.BackColor = Color.GreenYellow
                DTHRUPCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHRUPCOLNORqd", "DTHRUPCOLNO", "異常：需輸入縫上-色番(右左)")
            Case 2  '修改
                DTHRUPCOLNO.Visible = True
                DTHRUPCOLNO.BackColor = Color.Yellow
                DTHRUPCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHRUPCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHRUPCOLNO.Text = ""
        '縫上-YKK(右左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRUPYCOLNO.BackColor = Color.LightGray
                DTHRUPYCOLNO.Visible = True
                DTHRUPYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHRUPYCOLNO.Visible = True
                DTHRUPYCOLNO.BackColor = Color.GreenYellow
                DTHRUPYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHRUPYCOLNORqd", "DTHRUPYCOLNO", "異常：需輸入縫上-YKK(右左)")
            Case 2  '修改
                DTHRUPYCOLNO.Visible = True
                DTHRUPYCOLNO.BackColor = Color.Yellow
                DTHRUPYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHRUPYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHRUPYCOLNO.Text = ""

        '縫上-色番(右右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRRUPCOL.BackColor = Color.LightGray
                DTHRRUPCOL.Visible = True
            Case 1  '修改+檢查
                DTHRRUPCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTHRRUPCOLRqd", "DTHRRUPCOL", "異常：需輸入縫上-色番(右右)")
                DTHRRUPCOL.Visible = True
            Case 2  '修改
                DTHRRUPCOL.BackColor = Color.Yellow
                DTHRRUPCOL.Visible = True
            Case Else   '隱藏
                DTHRRUPCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "THRRUPCOL", "ZZZZZZ")
        '縫上-色番(右右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRRUPCOLNO.BackColor = Color.LightGray
                DTHRRUPCOLNO.Visible = True
                DTHRRUPCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHRRUPCOLNO.Visible = True
                DTHRRUPCOLNO.BackColor = Color.GreenYellow
                DTHRRUPCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHRRUPCOLNORqd", "DTHRRUPCOLNO", "異常：需輸入縫上-色番(右右)")
            Case 2  '修改
                DTHRRUPCOLNO.Visible = True
                DTHRRUPCOLNO.BackColor = Color.Yellow
                DTHRRUPCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHRRUPCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHRRUPCOLNO.Text = ""
        '縫上-YKK(右右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRRUPYCOLNO.BackColor = Color.LightGray
                DTHRRUPYCOLNO.Visible = True
                DTHRRUPYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHRRUPYCOLNO.Visible = True
                DTHRRUPYCOLNO.BackColor = Color.GreenYellow
                DTHRRUPYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHRRUPYCOLNORqd", "DTHRRUPYCOLNO", "異常：需輸入縫上-YKK(右右)")
            Case 2  '修改
                DTHRRUPYCOLNO.Visible = True
                DTHRRUPYCOLNO.BackColor = Color.Yellow
                DTHRRUPYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHRRUPYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHRRUPYCOLNO.Text = ""


        '縫下-色番(同)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLOCOL.BackColor = Color.LightGray
                DTHLOCOL.Visible = True
            Case 1  '修改+檢查
                DTHLOCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTHLOCOLRqd", "DTHLOCOL", "異常：需輸入縫下-色番(同)")
                DTHLOCOL.Visible = True
            Case 2  '修改
                DTHLOCOL.BackColor = Color.Yellow
                DTHLOCOL.Visible = True
            Case Else   '隱藏
                DTHLOCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "THLOCOL", "ZZZZZZ")
        '縫下-色番(同)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLOCOLNO.BackColor = Color.LightGray
                DTHLOCOLNO.Visible = True
                DTHLOCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHLOCOLNO.Visible = True
                DTHLOCOLNO.BackColor = Color.GreenYellow
                DTHLOCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHLOCOLNORqd", "DTHLOCOLNO", "異常：需輸入縫下-色番(同)")
            Case 2  '修改
                DTHLOCOLNO.Visible = True
                DTHLOCOLNO.BackColor = Color.Yellow
                DTHLOCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHLOCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHLOCOLNO.Text = ""
        '縫下-YKK(同)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLOYCOLNO.BackColor = Color.LightGray
                DTHLOYCOLNO.Visible = True
                DTHLOYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHLOYCOLNO.Visible = True
                DTHLOYCOLNO.BackColor = Color.GreenYellow
                DTHLOYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHLOYCOLNORqd", "DTHLOYCOLNO", "異常：需輸入縫下-YKK(同)")
            Case 2  '修改
                DTHLOYCOLNO.Visible = True
                DTHLOYCOLNO.BackColor = Color.Yellow
                DTHLOYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHLOYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHLOYCOLNO.Text = ""

        '縫下-色番(左左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLLOCOL.BackColor = Color.LightGray
                DTHLLOCOL.Visible = True
            Case 1  '修改+檢查
                DTHLLOCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTHLLOCOLRqd", "DTHLLOCOL", "異常：需輸入縫下-色番(左左)")
                DTHLLOCOL.Visible = True
            Case 2  '修改
                DTHLLOCOL.BackColor = Color.Yellow
                DTHLLOCOL.Visible = True
            Case Else   '隱藏
                DTHLLOCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "THLLOCOL", "ZZZZZZ")
        '縫下-色番(左左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLLOCOLNO.BackColor = Color.LightGray
                DTHLLOCOLNO.Visible = True
                DTHLLOCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHLLOCOLNO.Visible = True
                DTHLLOCOLNO.BackColor = Color.GreenYellow
                DTHLLOCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHLLOCOLNORqd", "DTHLLOCOLNO", "異常：需輸入縫下-色番(左左)")
            Case 2  '修改
                DTHLLOCOLNO.Visible = True
                DTHLLOCOLNO.BackColor = Color.Yellow
                DTHLLOCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHLLOCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHLLOCOLNO.Text = ""
        '縫下-YKK(左左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLLOYCOLNO.BackColor = Color.LightGray
                DTHLLOYCOLNO.Visible = True
                DTHLLOYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHLLOYCOLNO.Visible = True
                DTHLLOYCOLNO.BackColor = Color.GreenYellow
                DTHLLOYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHLLOYCOLNORqd", "DTHLLOYCOLNO", "異常：需輸入縫下-YKK(左左)")
            Case 2  '修改
                DTHLLOYCOLNO.Visible = True
                DTHLLOYCOLNO.BackColor = Color.Yellow
                DTHLLOYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHLLOYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHLLOYCOLNO.Text = ""

        '縫下-色番(左右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLRLOCOL.BackColor = Color.LightGray
                DTHLRLOCOL.Visible = True
            Case 1  '修改+檢查
                DTHLRLOCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTHLRLOCOLRqd", "DTHLRLOCOL", "異常：需輸入縫下-色番(左右)")
                DTHLRLOCOL.Visible = True
            Case 2  '修改
                DTHLRLOCOL.BackColor = Color.Yellow
                DTHLRLOCOL.Visible = True
            Case Else   '隱藏
                DTHLRLOCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "THLRLOCOL", "ZZZZZZ")
        '縫下-色番(左左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLRLOCOLNO.BackColor = Color.LightGray
                DTHLRLOCOLNO.Visible = True
                DTHLRLOCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHLRLOCOLNO.Visible = True
                DTHLRLOCOLNO.BackColor = Color.GreenYellow
                DTHLRLOCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHLRLOCOLNORqd", "DTHLRLOCOLNO", "異常：需輸入縫下-色番(左左)")
            Case 2  '修改
                DTHLRLOCOLNO.Visible = True
                DTHLRLOCOLNO.BackColor = Color.Yellow
                DTHLRLOCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHLRLOCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHLRLOCOLNO.Text = ""
        '縫下-YKK(左左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLRLOYCOLNO.BackColor = Color.LightGray
                DTHLRLOYCOLNO.Visible = True
                DTHLRLOYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHLRLOYCOLNO.Visible = True
                DTHLRLOYCOLNO.BackColor = Color.GreenYellow
                DTHLRLOYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHLRLOYCOLNORqd", "DTHLRLOYCOLNO", "異常：需輸入縫下-YKK(左左)")
            Case 2  '修改
                DTHLRLOYCOLNO.Visible = True
                DTHLRLOYCOLNO.BackColor = Color.Yellow
                DTHLRLOYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHLRLOYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHLRLOYCOLNO.Text = ""


        '縫下-色番(右左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRLOCOL.BackColor = Color.LightGray
                DTHRLOCOL.Visible = True
            Case 1  '修改+檢查
                DTHRLOCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTHRLOCOLRqd", "DTHRLOCOL", "異常：需輸入縫下-色番(右左)")
                DTHRLOCOL.Visible = True
            Case 2  '修改
                DTHRLOCOL.BackColor = Color.Yellow
                DTHRLOCOL.Visible = True
            Case Else   '隱藏
                DTHRLOCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "THRLOCOL", "ZZZZZZ")
        '縫下-色番(右左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRLOCOLNO.BackColor = Color.LightGray
                DTHRLOCOLNO.Visible = True
                DTHRLOCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHRLOCOLNO.Visible = True
                DTHRLOCOLNO.BackColor = Color.GreenYellow
                DTHRLOCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHRLOCOLNORqd", "DTHRLOCOLNO", "異常：需輸入縫下-色番(右左)")
            Case 2  '修改
                DTHRLOCOLNO.Visible = True
                DTHRLOCOLNO.BackColor = Color.Yellow
                DTHRLOCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHRLOCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHRLOCOLNO.Text = ""
        '縫下-YKK(右左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRLOYCOLNO.BackColor = Color.LightGray
                DTHRLOYCOLNO.Visible = True
                DTHRLOYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHRLOYCOLNO.Visible = True
                DTHRLOYCOLNO.BackColor = Color.GreenYellow
                DTHRLOYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHRLOYCOLNORqd", "DTHRLOYCOLNO", "異常：需輸入縫下-YKK(右左)")
            Case 2  '修改
                DTHRLOYCOLNO.Visible = True
                DTHRLOYCOLNO.BackColor = Color.Yellow
                DTHRLOYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHRLOYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHRLOYCOLNO.Text = ""


        '縫下-色番(右右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRRLOCOL.BackColor = Color.LightGray
                DTHRRLOCOL.Visible = True
            Case 1  '修改+檢查
                DTHRRLOCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTHRRLOCOLRqd", "DTHRRLOCOL", "異常：需輸入縫下-色番(右右)")
                DTHRRLOCOL.Visible = True
            Case 2  '修改
                DTHRRLOCOL.BackColor = Color.Yellow
                DTHRRLOCOL.Visible = True
            Case Else   '隱藏
                DTHRRLOCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "THRRLOCOL", "ZZZZZZ")
        '縫下-色番(右右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRRLOCOLNO.BackColor = Color.LightGray
                DTHRRLOCOLNO.Visible = True
                DTHRRLOCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHRRLOCOLNO.Visible = True
                DTHRRLOCOLNO.BackColor = Color.GreenYellow
                DTHRRLOCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHRRLOCOLNORqd", "DTHRRLOCOLNO", "異常：需輸入縫下-色番(右右)")
            Case 2  '修改
                DTHRRLOCOLNO.Visible = True
                DTHRRLOCOLNO.BackColor = Color.Yellow
                DTHRRLOCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHRRLOCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHRRLOCOLNO.Text = ""
        '縫下-YKK(右右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRRLOYCOLNO.BackColor = Color.LightGray
                DTHRRLOYCOLNO.Visible = True
                DTHRRLOYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHRRLOYCOLNO.Visible = True
                DTHRRLOYCOLNO.BackColor = Color.GreenYellow
                DTHRRLOYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHRRLOYCOLNORqd", "DTHRRLOYCOLNO", "異常：需輸入縫下-YKK(右右)")
            Case 2  '修改
                DTHRRLOYCOLNO.Visible = True
                DTHRRLOYCOLNO.BackColor = Color.Yellow
                DTHRRLOYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHRRLOYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHRRLOYCOLNO.Text = ""

        'X-尺寸
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DXMLEN.BackColor = Color.LightGray
                DXMLEN.Visible = True
                DXMLEN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DXMLEN.Visible = True
                DXMLEN.BackColor = Color.GreenYellow
                DXMLEN.ReadOnly = False
                ShowRequiredFieldValidator("DXMLENRqd", "DXMLEN", "異常：需輸入X-尺寸")
            Case 2  '修改
                DXMLEN.Visible = True
                DXMLEN.BackColor = Color.Yellow
                DXMLEN.ReadOnly = False
            Case Else   '隱藏
                DXMLEN.Visible = False
        End Select
        If pPost = "New" Then DXMLEN.Text = "0"
        'X-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DXMCOL.BackColor = Color.LightGray
                DXMCOL.Visible = True
            Case 1  '修改+檢查
                DXMCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DXMCOLRqd", "DXMCOL", "異常：需輸入X-色番")
                DXMCOL.Visible = True
            Case 2  '修改
                DXMCOL.BackColor = Color.Yellow
                DXMCOL.Visible = True
            Case Else   '隱藏
                DXMCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "XMCOL", "ZZZZZZ")
        'X-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DXMCOLNO.BackColor = Color.LightGray
                DXMCOLNO.Visible = True
                DXMCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DXMCOLNO.Visible = True
                DXMCOLNO.BackColor = Color.GreenYellow
                DXMCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DXMCOLNORqd", "DXMCOLNO", "異常：需輸入X-色番")
            Case 2  '修改
                DXMCOLNO.Visible = True
                DXMCOLNO.BackColor = Color.Yellow
                DXMCOLNO.ReadOnly = False
            Case Else   '隱藏
                DXMCOLNO.Visible = False
        End Select
        If pPost = "New" Then DXMCOLNO.Text = ""
        'X-YKK
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DXMYCOLNO.BackColor = Color.LightGray
                DXMYCOLNO.Visible = True
                DXMYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DXMYCOLNO.Visible = True
                DXMYCOLNO.BackColor = Color.GreenYellow
                DXMYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DXMYCOLNORqd", "DXMYCOLNO", "異常：需輸入X-YKK")
            Case 2  '修改
                DXMYCOLNO.Visible = True
                DXMYCOLNO.BackColor = Color.Yellow
                DXMYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DXMYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DXMYCOLNO.Text = ""
        'A-尺寸
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DAMLEN.BackColor = Color.LightGray
                DAMLEN.Visible = True
                DAMLEN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DAMLEN.Visible = True
                DAMLEN.BackColor = Color.GreenYellow
                DAMLEN.ReadOnly = False
                ShowRequiredFieldValidator("DAMLENRqd", "DAMLEN", "異常：需輸入A-尺寸")
            Case 2  '修改
                DAMLEN.Visible = True
                DAMLEN.BackColor = Color.Yellow
                DAMLEN.ReadOnly = False
            Case Else   '隱藏
                DAMLEN.Visible = False
        End Select
        If pPost = "New" Then DAMLEN.Text = "0"
        'A-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DAMCOL.BackColor = Color.LightGray
                DAMCOL.Visible = True
            Case 1  '修改+檢查
                DAMCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAMCOLRqd", "DAMCOL", "異常：需輸入A-色番")
                DAMCOL.Visible = True
            Case 2  '修改
                DAMCOL.BackColor = Color.Yellow
                DAMCOL.Visible = True
            Case Else   '隱藏
                DAMCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "AMCOL", "ZZZZZZ")
        'A-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DAMCOLNO.BackColor = Color.LightGray
                DAMCOLNO.Visible = True
                DAMCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DAMCOLNO.Visible = True
                DAMCOLNO.BackColor = Color.GreenYellow
                DAMCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DAMCOLNORqd", "DAMCOLNO", "異常：需輸入A-色番")
            Case 2  '修改
                DAMCOLNO.Visible = True
                DAMCOLNO.BackColor = Color.Yellow
                DAMCOLNO.ReadOnly = False
            Case Else   '隱藏
                DAMCOLNO.Visible = False
        End Select
        If pPost = "New" Then DAMCOLNO.Text = ""
        'A-YKK
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DAMYCOLNO.BackColor = Color.LightGray
                DAMYCOLNO.Visible = True
                DAMYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DAMYCOLNO.Visible = True
                DAMYCOLNO.BackColor = Color.GreenYellow
                DAMYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DAMYCOLNORqd", "DAMYCOLNO", "異常：需輸入A-YKK")
            Case 2  '修改
                DAMYCOLNO.Visible = True
                DAMYCOLNO.BackColor = Color.Yellow
                DAMYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DAMYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DAMYCOLNO.Text = ""
        'B-尺寸
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DBMLEN.BackColor = Color.LightGray
                DBMLEN.Visible = True
                DBMLEN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DBMLEN.Visible = True
                DBMLEN.BackColor = Color.GreenYellow
                DBMLEN.ReadOnly = False
                ShowRequiredFieldValidator("DBMLENRqd", "DBMLEN", "異常：需輸入B-尺寸")
            Case 2  '修改
                DBMLEN.Visible = True
                DBMLEN.BackColor = Color.Yellow
                DBMLEN.ReadOnly = False
            Case Else   '隱藏
                DBMLEN.Visible = False
        End Select
        If pPost = "New" Then DBMLEN.Text = "0"
        'B-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DBMCOL.BackColor = Color.LightGray
                DBMCOL.Visible = True
            Case 1  '修改+檢查
                DBMCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBMCOLRqd", "DBMCOL", "異常：需輸入B-色番")
                DBMCOL.Visible = True
            Case 2  '修改
                DBMCOL.BackColor = Color.Yellow
                DBMCOL.Visible = True
            Case Else   '隱藏
                DBMCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "BMCOL", "ZZZZZZ")
        'B-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DBMCOLNO.BackColor = Color.LightGray
                DBMCOLNO.Visible = True
                DBMCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DBMCOLNO.Visible = True
                DBMCOLNO.BackColor = Color.GreenYellow
                DBMCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DBMCOLNORqd", "DBMCOLNO", "異常：需輸入B-色番")
            Case 2  '修改
                DBMCOLNO.Visible = True
                DBMCOLNO.BackColor = Color.Yellow
                DBMCOLNO.ReadOnly = False
            Case Else   '隱藏
                DBMCOLNO.Visible = False
        End Select
        If pPost = "New" Then DBMCOLNO.Text = ""
        'B-YKK
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DBMYCOLNO.BackColor = Color.LightGray
                DBMYCOLNO.Visible = True
                DBMYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DBMYCOLNO.Visible = True
                DBMYCOLNO.BackColor = Color.GreenYellow
                DBMYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DBMYCOLNORqd", "DBMYCOLNO", "異常：需輸入B-YKK")
            Case 2  '修改
                DBMYCOLNO.Visible = True
                DBMYCOLNO.BackColor = Color.Yellow
                DBMYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DBMYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DBMYCOLNO.Text = ""
        'C-尺寸
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DCMLEN.BackColor = Color.LightGray
                DCMLEN.Visible = True
                DCMLEN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCMLEN.Visible = True
                DCMLEN.BackColor = Color.GreenYellow
                DCMLEN.ReadOnly = False
                ShowRequiredFieldValidator("DCMLENRqd", "DCMLEN", "異常：需輸入C-尺寸")
            Case 2  '修改
                DCMLEN.Visible = True
                DCMLEN.BackColor = Color.Yellow
                DCMLEN.ReadOnly = False
            Case Else   '隱藏
                DCMLEN.Visible = False
        End Select
        If pPost = "New" Then DCMLEN.Text = "0"
        'C-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DCMCOL.BackColor = Color.LightGray
                DCMCOL.Visible = True
            Case 1  '修改+檢查
                DCMCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCMCOLRqd", "DCMCOL", "異常：需輸入C-色番")
                DCMCOL.Visible = True
            Case 2  '修改
                DCMCOL.BackColor = Color.Yellow
                DCMCOL.Visible = True
            Case Else   '隱藏
                DCMCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "CMCOL", "ZZZZZZ")
        'C-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DCMCOLNO.BackColor = Color.LightGray
                DCMCOLNO.Visible = True
                DCMCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCMCOLNO.Visible = True
                DCMCOLNO.BackColor = Color.GreenYellow
                DCMCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DCMCOLNORqd", "DCMCOLNO", "異常：需輸入C-色番")
            Case 2  '修改
                DCMCOLNO.Visible = True
                DCMCOLNO.BackColor = Color.Yellow
                DCMCOLNO.ReadOnly = False
            Case Else   '隱藏
                DCMCOLNO.Visible = False
        End Select
        If pPost = "New" Then DCMCOLNO.Text = ""
        'C-YKK
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DCMYCOLNO.BackColor = Color.LightGray
                DCMYCOLNO.Visible = True
                DCMYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCMYCOLNO.Visible = True
                DCMYCOLNO.BackColor = Color.GreenYellow
                DCMYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DCMYCOLNORqd", "DCMYCOLNO", "異常：需輸入C-YKK")
            Case 2  '修改
                DCMYCOLNO.Visible = True
                DCMYCOLNO.BackColor = Color.Yellow
                DCMYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DCMYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DCMYCOLNO.Text = ""
        'D-尺寸
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DDMLEN.BackColor = Color.LightGray
                DDMLEN.Visible = True
                DDMLEN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DDMLEN.Visible = True
                DDMLEN.BackColor = Color.GreenYellow
                DDMLEN.ReadOnly = False
                ShowRequiredFieldValidator("DDMLENRqd", "DDMLEN", "異常：需輸入D-尺寸")
            Case 2  '修改
                DDMLEN.Visible = True
                DDMLEN.BackColor = Color.Yellow
                DDMLEN.ReadOnly = False
            Case Else   '隱藏
                DDMLEN.Visible = False
        End Select
        If pPost = "New" Then DDMLEN.Text = "0"
        'D-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DDMCOL.BackColor = Color.LightGray
                DDMCOL.Visible = True
            Case 1  '修改+檢查
                DDMCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDMCOLRqd", "DDMCOL", "異常：需輸入D-色番")
                DDMCOL.Visible = True
            Case 2  '修改
                DDMCOL.BackColor = Color.Yellow
                DDMCOL.Visible = True
            Case Else   '隱藏
                DDMCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "DMCOL", "ZZZZZZ")
        'D-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DDMCOLNO.BackColor = Color.LightGray
                DDMCOLNO.Visible = True
                DDMCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DDMCOLNO.Visible = True
                DDMCOLNO.BackColor = Color.GreenYellow
                DDMCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DDMCOLNORqd", "DDMCOLNO", "異常：需輸入D-色番")
            Case 2  '修改
                DDMCOLNO.Visible = True
                DDMCOLNO.BackColor = Color.Yellow
                DDMCOLNO.ReadOnly = False
            Case Else   '隱藏
                DDMCOLNO.Visible = False
        End Select
        If pPost = "New" Then DDMCOLNO.Text = ""
        'D-YKK
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DDMYCOLNO.BackColor = Color.LightGray
                DDMYCOLNO.Visible = True
                DDMYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DDMYCOLNO.Visible = True
                DDMYCOLNO.BackColor = Color.GreenYellow
                DDMYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DDMYCOLNORqd", "DDMYCOLNO", "異常：需輸入D-YKK")
            Case 2  '修改
                DDMYCOLNO.Visible = True
                DDMYCOLNO.BackColor = Color.Yellow
                DDMYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DDMYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DDMYCOLNO.Text = ""
        'E-尺寸
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DEMLEN.BackColor = Color.LightGray
                DEMLEN.Visible = True
                DEMLEN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DEMLEN.Visible = True
                DEMLEN.BackColor = Color.GreenYellow
                DEMLEN.ReadOnly = False
                ShowRequiredFieldValidator("DEMLENRqd", "DEMLEN", "異常：需輸入E-尺寸")
            Case 2  '修改
                DEMLEN.Visible = True
                DEMLEN.BackColor = Color.Yellow
                DEMLEN.ReadOnly = False
            Case Else   '隱藏
                DEMLEN.Visible = False
        End Select
        If pPost = "New" Then DEMLEN.Text = "0"
        'E-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DEMCOL.BackColor = Color.LightGray
                DEMCOL.Visible = True
            Case 1  '修改+檢查
                DEMCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEMCOLRqd", "DEMCOL", "異常：需輸入E-色番")
                DEMCOL.Visible = True
            Case 2  '修改
                DEMCOL.BackColor = Color.Yellow
                DEMCOL.Visible = True
            Case Else   '隱藏
                DEMCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "EMCOL", "ZZZZZZ")
        'E-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DEMCOLNO.BackColor = Color.LightGray
                DEMCOLNO.Visible = True
                DEMCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DEMCOLNO.Visible = True
                DEMCOLNO.BackColor = Color.GreenYellow
                DEMCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DEMCOLNORqd", "DEMCOLNO", "異常：需輸入E-色番")
            Case 2  '修改
                DEMCOLNO.Visible = True
                DEMCOLNO.BackColor = Color.Yellow
                DEMCOLNO.ReadOnly = False
            Case Else   '隱藏
                DEMCOLNO.Visible = False
        End Select
        If pPost = "New" Then DEMCOLNO.Text = ""
        'E-YKK
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DEMYCOLNO.BackColor = Color.LightGray
                DEMYCOLNO.Visible = True
                DEMYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DEMYCOLNO.Visible = True
                DEMYCOLNO.BackColor = Color.GreenYellow
                DEMYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DEMYCOLNORqd", "DEMYCOLNO", "異常：需輸入E-YKK")
            Case 2  '修改
                DEMYCOLNO.Visible = True
                DEMYCOLNO.BackColor = Color.Yellow
                DEMYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DEMYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DEMYCOLNO.Text = ""
        'F-尺寸
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DFMLEN.BackColor = Color.LightGray
                DFMLEN.Visible = True
                DFMLEN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DFMLEN.Visible = True
                DFMLEN.BackColor = Color.GreenYellow
                DFMLEN.ReadOnly = False
                ShowRequiredFieldValidator("DFMLENRqd", "DFMLEN", "異常：需輸入F-尺寸")
            Case 2  '修改
                DFMLEN.Visible = True
                DFMLEN.BackColor = Color.Yellow
                DFMLEN.ReadOnly = False
            Case Else   '隱藏
                DFMLEN.Visible = False
        End Select
        If pPost = "New" Then DFMLEN.Text = "0"
        'F-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DFMCOL.BackColor = Color.LightGray
                DFMCOL.Visible = True
            Case 1  '修改+檢查
                DFMCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFMCOLRqd", "DFMCOL", "異常：需輸入F-色番")
                DFMCOL.Visible = True
            Case 2  '修改
                DFMCOL.BackColor = Color.Yellow
                DFMCOL.Visible = True
            Case Else   '隱藏
                DFMCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "FMCOL", "ZZZZZZ")
        'F-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DFMCOLNO.BackColor = Color.LightGray
                DFMCOLNO.Visible = True
                DFMCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DFMCOLNO.Visible = True
                DFMCOLNO.BackColor = Color.GreenYellow
                DFMCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DFMCOLNORqd", "DFMCOLNO", "異常：需輸入F-色番")
            Case 2  '修改
                DFMCOLNO.Visible = True
                DFMCOLNO.BackColor = Color.Yellow
                DFMCOLNO.ReadOnly = False
            Case Else   '隱藏
                DFMCOLNO.Visible = False
        End Select
        If pPost = "New" Then DFMCOLNO.Text = ""
        'F-YKK
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DFMYCOLNO.BackColor = Color.LightGray
                DFMYCOLNO.Visible = True
                DFMYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DFMYCOLNO.Visible = True
                DFMYCOLNO.BackColor = Color.GreenYellow
                DFMYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DFMYCOLNORqd", "DFMYCOLNO", "異常：需輸入F-YKK")
            Case 2  '修改
                DFMYCOLNO.Visible = True
                DFMYCOLNO.BackColor = Color.Yellow
                DFMYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DFMYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DFMYCOLNO.Text = ""
        'G-尺寸
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DGMLEN.BackColor = Color.LightGray
                DGMLEN.Visible = True
                DGMLEN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DGMLEN.Visible = True
                DGMLEN.BackColor = Color.GreenYellow
                DGMLEN.ReadOnly = False
                ShowRequiredFieldValidator("DGMLENRqd", "DGMLEN", "異常：需輸入G-尺寸")
            Case 2  '修改
                DGMLEN.Visible = True
                DGMLEN.BackColor = Color.Yellow
                DGMLEN.ReadOnly = False
            Case Else   '隱藏
                DGMLEN.Visible = False
        End Select
        If pPost = "New" Then DGMLEN.Text = "0"
        'G-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DGMCOL.BackColor = Color.LightGray
                DGMCOL.Visible = True
            Case 1  '修改+檢查
                DGMCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DGMCOLRqd", "DGMCOL", "異常：需輸入G-色番")
                DGMCOL.Visible = True
            Case 2  '修改
                DGMCOL.BackColor = Color.Yellow
                DGMCOL.Visible = True
            Case Else   '隱藏
                DGMCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "GMCOL", "ZZZZZZ")
        'G-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DGMCOLNO.BackColor = Color.LightGray
                DGMCOLNO.Visible = True
                DGMCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DGMCOLNO.Visible = True
                DGMCOLNO.BackColor = Color.GreenYellow
                DGMCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DGMCOLNORqd", "DGMCOLNO", "異常：需輸入G-色番")
            Case 2  '修改
                DGMCOLNO.Visible = True
                DGMCOLNO.BackColor = Color.Yellow
                DGMCOLNO.ReadOnly = False
            Case Else   '隱藏
                DGMCOLNO.Visible = False
        End Select
        If pPost = "New" Then DGMCOLNO.Text = ""
        'G-YKK
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DGMYCOLNO.BackColor = Color.LightGray
                DGMYCOLNO.Visible = True
                DGMYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DGMYCOLNO.Visible = True
                DGMYCOLNO.BackColor = Color.GreenYellow
                DGMYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DGMYCOLNORqd", "DGMYCOLNO", "異常：需輸入G-YKK")
            Case 2  '修改
                DGMYCOLNO.Visible = True
                DGMYCOLNO.BackColor = Color.Yellow
                DGMYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DGMYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DGMYCOLNO.Text = ""
        'H-尺寸
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DHMLEN.BackColor = Color.LightGray
                DHMLEN.Visible = True
                DHMLEN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DHMLEN.Visible = True
                DHMLEN.BackColor = Color.GreenYellow
                DHMLEN.ReadOnly = False
                ShowRequiredFieldValidator("DHMLENRqd", "DHMLEN", "異常：需輸入H-尺寸")
            Case 2  '修改
                DHMLEN.Visible = True
                DHMLEN.BackColor = Color.Yellow
                DHMLEN.ReadOnly = False
            Case Else   '隱藏
                DHMLEN.Visible = False
        End Select
        If pPost = "New" Then DHMLEN.Text = "0"
        'H-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DHMCOL.BackColor = Color.LightGray
                DHMCOL.Visible = True
            Case 1  '修改+檢查
                DHMCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DHMCOLRqd", "DHMCOL", "異常：需輸入H-色番")
                DHMCOL.Visible = True
            Case 2  '修改
                DHMCOL.BackColor = Color.Yellow
                DHMCOL.Visible = True
            Case Else   '隱藏
                DHMCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "HMCOL", "ZZZZZZ")
        'H-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DHMCOLNO.BackColor = Color.LightGray
                DHMCOLNO.Visible = True
                DHMCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DHMCOLNO.Visible = True
                DHMCOLNO.BackColor = Color.GreenYellow
                DHMCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DHMCOLNORqd", "DHMCOLNO", "異常：需輸入H-色番")
            Case 2  '修改
                DHMCOLNO.Visible = True
                DHMCOLNO.BackColor = Color.Yellow
                DHMCOLNO.ReadOnly = False
            Case Else   '隱藏
                DHMCOLNO.Visible = False
        End Select
        If pPost = "New" Then DHMCOLNO.Text = ""
        'H-YKK
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DHMYCOLNO.BackColor = Color.LightGray
                DHMYCOLNO.Visible = True
                DHMYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DHMYCOLNO.Visible = True
                DHMYCOLNO.BackColor = Color.GreenYellow
                DHMYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DHMYCOLNORqd", "DHMYCOLNO", "異常：需輸入H-YKK")
            Case 2  '修改
                DHMYCOLNO.Visible = True
                DHMYCOLNO.BackColor = Color.Yellow
                DHMYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DHMYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DHMYCOLNO.Text = ""
        '緯紗-尺寸
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DLYLEN.BackColor = Color.LightGray
                DLYLEN.Visible = True
                DLYLEN.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DLYLEN.Visible = True
                DLYLEN.BackColor = Color.GreenYellow
                DLYLEN.ReadOnly = False
                ShowRequiredFieldValidator("DLYLENRqd", "DLYLEN", "異常：需輸入緯紗-尺寸")
            Case 2  '修改
                DLYLEN.Visible = True
                DLYLEN.BackColor = Color.Yellow
                DLYLEN.ReadOnly = False
            Case Else   '隱藏
                DLYLEN.Visible = False
        End Select
        If pPost = "New" Then DLYLEN.Text = "0"
        '緯紗-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DLYCOL.BackColor = Color.LightGray
                DLYCOL.Visible = True
            Case 1  '修改+檢查
                DLYCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DLYCOLRqd", "DLYCOL", "異常：需輸入緯紗-色番")
                DLYCOL.Visible = True
            Case 2  '修改
                DLYCOL.BackColor = Color.Yellow
                DLYCOL.Visible = True
            Case Else   '隱藏
                DLYCOL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-DEVELOP"), "LYCOL", "ZZZZZZ")
        '緯紗-色番
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DLYCOLNO.BackColor = Color.LightGray
                DLYCOLNO.Visible = True
                DLYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DLYCOLNO.Visible = True
                DLYCOLNO.BackColor = Color.GreenYellow
                DLYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DLYCOLNORqd", "DLYCOLNO", "異常：需輸入緯紗-色番")
            Case 2  '修改
                DLYCOLNO.Visible = True
                DLYCOLNO.BackColor = Color.Yellow
                DLYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DLYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DLYCOLNO.Text = ""
        '緯紗-YKK
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DLYYCOLNO.BackColor = Color.LightGray
                DLYYCOLNO.Visible = True
                DLYYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DLYYCOLNO.Visible = True
                DLYYCOLNO.BackColor = Color.GreenYellow
                DLYYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DLYYCOLNORqd", "DLYYCOLNO", "異常：需輸入緯紗-YKK")
            Case 2  '修改
                DLYYCOLNO.Visible = True
                DLYYCOLNO.BackColor = Color.Yellow
                DLYYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DLYYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DLYYCOLNO.Text = ""
        '其他條件
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DOTCON.BackColor = Color.LightGray
                DOTCON.Visible = True
                DOTCON.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DOTCON.Visible = True
                DOTCON.BackColor = Color.GreenYellow
                DOTCON.ReadOnly = False
                ShowRequiredFieldValidator("DOTCONRqd", "DOTCON", "異常：需輸入其他條件")
            Case 2  '修改
                DOTCON.Visible = True
                DOTCON.BackColor = Color.Yellow
                DOTCON.ReadOnly = False
            Case Else   '隱藏
                DOTCON.Visible = False
        End Select
        If pPost = "New" Then DOTCON.Text = ""
        '----圖面欄位設定-------------------------------------------------    

        '製圖者
        Select Case FindFieldInf("MAKEMAP")
            Case 0  '顯示
                DMAKEMAP.BackColor = Color.LightGray
                DMAKEMAP.Visible = True
            Case 1  '修改+檢查
                DMAKEMAP.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMAKEMAPRqd", "DMAKEMAP", "異常：需輸入製圖者")
                DMAKEMAP.Visible = True
            Case 2  '修改
                DMAKEMAP.BackColor = Color.Yellow
                DMAKEMAP.Visible = True
            Case Else   '隱藏
                DMAKEMAP.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("MAKEMAP"), "MAKEMAP", "ZZZZZZ")


        '圖號
        Select Case FindFieldInf("1-MAP")
            Case 0  '顯示
                DMAPNO.BackColor = Color.LightGray
                DMAPNO.Visible = True
                DMAPNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DMAPNO.Visible = True
                DMAPNO.BackColor = Color.GreenYellow
                DMAPNO.ReadOnly = False
                ShowRequiredFieldValidator("DMAPNORqd", "DMAPNO", "異常：需輸入圖號")
            Case 2  '修改
                DMAPNO.Visible = True
                DMAPNO.BackColor = Color.Yellow
                DMAPNO.ReadOnly = False
            Case Else   '隱藏
                DMAPNO.Visible = False
        End Select
        If pPost = "New" Then DMAPNO.Text = ""
      
        '難易度
        Select Case FindFieldInf("1-MAP")
            Case 0  '顯示
                DLEVEL.BackColor = Color.LightGray
                DLEVEL.Visible = True
            Case 1  '修改+檢查
                DLEVEL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DLEVELRqd", "DLEVEL", "異常：需輸入難易度")
                DLEVEL.Visible = True
            Case 2  '修改
                DLEVEL.BackColor = Color.Yellow
                DLEVEL.Visible = True
            Case Else   '隱藏
                DLEVEL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-MAP"), "LEVEL", "ZZZZZZ")
        '圖檔
        Select Case FindFieldInf("1-MAP")
            Case 0  '顯示
                DMAPFILE.Visible = False
                DMAPFILE.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DMAPFILERqd", "DMAPFILE", "異常：需輸入圖檔")
                DMAPFILE.Visible = True
                DMAPFILE.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DMAPFILE.Visible = True
                DMAPFILE.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DMAPFILE.Visible = False
        End Select
        If pPost = "New" Then LMAPFILE.Visible = False
        '----適用型別欄位設定-------------------------------------------------    
        '適用型別檔
        Select Case FindFieldInf("1-FORTYPE")
            Case 0  '顯示
                DFORTYPEFILE.Visible = False
                DFORTYPEFILE.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DFORTYPEFILERqd", "DFORTYPEFILE", "異常：需輸入適用型別檔")
                DFORTYPEFILE.Visible = True
                DFORTYPEFILE.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DFORTYPEFILE.Visible = True
                DFORTYPEFILE.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DFORTYPEFILE.Visible = False
        End Select
        If pPost = "New" Then LFORTYPEFILE.Visible = False
        '----品質檢測欄位設定-------------------------------------------------    
        '品質檢測項目檔
        Select Case FindFieldInf("1-QCCHECKFILE")
            Case 0  '顯示
                DQCCHECKFILE.Visible = False
                DQCCHECKFILE.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQCCHECKFILERqd", "DQCCHECKFILE", "異常：需輸入品質檢測項目檔")
                DQCCHECKFILE.Visible = True
                DQCCHECKFILE.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DQCCHECKFILE.Visible = True
                DQCCHECKFILE.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DQCCHECKFILE.Visible = False
        End Select
        If pPost = "New" Then LQCCHECKFILE.Visible = False
        'QC-Leadtime
        Select Case FindFieldInf("1-QCLEADTIME")
            Case 0  '顯示
                DQCLT.BackColor = Color.LightGray
                DQCLT.Visible = True
                DQCLT.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DQCLT.Visible = True
                DQCLT.BackColor = Color.GreenYellow
                DQCLT.ReadOnly = False
                ShowRequiredFieldValidator("DQCLTRqd", "DQCLT", "異常：需輸入品質測試時間")
            Case 2  '修改
                DQCLT.Visible = True
                DQCLT.BackColor = Color.Yellow
                DQCLT.ReadOnly = False
            Case Else   '隱藏
                'DQCLT.Visible = False
                DQCLT.Style("top") = -100 & "px"
        End Select
        If pPost = "New" Then DQCLT.Text = "0"
        'QC-PEOPLE
        Select Case FindFieldInf("1-QCPEOPLE")
            Case 0  '顯示
                DQCPeople.BackColor = Color.LightGray
                DQCPeople.Visible = True
            Case 1  '修改+檢查
                DQCPeople.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCPeopleRqd", "DQCPeople", "異常：需輸入QC人員")
                DQCPeople.Visible = True
            Case 2  '修改
                DQCPeople.BackColor = Color.Yellow
                DQCPeople.Visible = True
            Case Else   '隱藏
                DQCPeople.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("1-QCPEOPLE"), "QCPEOPLE", "ZZZZZZ")
        'QC檢測文件-1
        Select Case FindFieldInf("1-QCFILE")
            Case 0  '顯示
                DQCFILE1.Visible = False
                DQCFILE1.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DQCFILE2.Visible = False
                DQCFILE2.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DQCFILE3.Visible = False
                DQCFILE3.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DQCFILE4.Visible = False
                DQCFILE4.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DQCFILE5.Visible = False
                DQCFILE5.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DQCFILE6.Visible = False
                DQCFILE6.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DGenFILE1.Visible = False
                DGenFILE1.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQCFILE1Rqd", "DQCFILE1", "異常：需輸入QC檢測文件-1")
                DQCFILE1.Visible = True
                DQCFILE1.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DQCFILE2Rqd", "DQCFILE2", "異常：需輸入QC檢測文件-2")
                DQCFILE2.Visible = True
                DQCFILE2.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DQCFILE3Rqd", "DQCFILE3", "異常：需輸入QC檢測文件-3")
                DQCFILE3.Visible = True
                DQCFILE3.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DQCFILE4Rqd", "DQCFILE4", "異常：需輸入QC檢測文件-4")
                DQCFILE4.Visible = True
                DQCFILE4.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DQCFILE5Rqd", "DQCFILE5", "異常：需輸入QC檢測文件-5")
                DQCFILE5.Visible = True
                DQCFILE5.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DQCFILE6Rqd", "DQCFILE6", "異常：需輸入QC檢測文件-6")
                DQCFILE6.Visible = True
                DQCFILE6.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DGenFILE1Rqd", "DGenFILE1", "異常：需輸入原單位文件")
                DGenFILE1.Visible = True
                DGenFILE1.Style.Add("BACKGROUND-COLOR", "GreenYellow")

            Case 2  '修改
                DQCFILE1.Visible = True
                DQCFILE1.Style.Add("BACKGROUND-COLOR", "Yellow")
                DQCFILE2.Visible = True
                DQCFILE2.Style.Add("BACKGROUND-COLOR", "Yellow")
                DQCFILE3.Visible = True
                DQCFILE3.Style.Add("BACKGROUND-COLOR", "Yellow")
                DQCFILE4.Visible = True
                DQCFILE4.Style.Add("BACKGROUND-COLOR", "Yellow")
                DQCFILE5.Visible = True
                DQCFILE5.Style.Add("BACKGROUND-COLOR", "Yellow")
                DQCFILE6.Visible = True
                DQCFILE6.Style.Add("BACKGROUND-COLOR", "Yellow")
                DGenFILE1.Visible = True
                DGenFILE1.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DQCFILE1.Visible = False
                DQCFILE2.Visible = False
                DQCFILE3.Visible = False
                DQCFILE4.Visible = False
                DQCFILE5.Visible = False
                DQCFILE6.Visible = False
                DGenFILE1.Visible = False
        End Select
        If pPost = "New" Then
            LQCFILE1.Visible = False
            LQCFILE2.Visible = False
            LQCFILE3.Visible = False
            LQCFILE4.Visible = False
            LQCFILE5.Visible = False
            LQCFILE6.Visible = False
            LGenFILE1.Visible = False
        End If
        '客戶切結書
        Select Case FindFieldInf("1-CONTACTFILE")
            Case 0  '顯示
                DCONTACTFILE.Visible = False
                DCONTACTFILE.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DCONTACTFILERqd", "DCONTACTFILE", "異常：需輸入客戶切結書")
                DCONTACTFILE.Visible = True
                DCONTACTFILE.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DCONTACTFILE.Visible = True
                DCONTACTFILE.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DCONTACTFILE.Visible = False
        End Select
        If pPost = "New" Then LCONTACTFILE.Visible = False
        '樣品確認書
        Select Case FindFieldInf("1-SAMPLECONFIRMFILE")
            Case 0  '顯示
                DSAMPLECONFIRMFILE.Visible = False
                DSAMPLECONFIRMFILE.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DSAMPLECONFIRMFILERqd", "DSAMPLECONFIRMFILE", "異常：需輸入樣品確認書")
                DSAMPLECONFIRMFILE.Visible = True
                DSAMPLECONFIRMFILE.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DSAMPLECONFIRMFILE.Visible = True
                DSAMPLECONFIRMFILE.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DSAMPLECONFIRMFILE.Visible = False
        End Select
        If pPost = "New" Then LSAMPLECONFIRMFILE.Visible = False
        '製造授權書
        Select Case FindFieldInf("1-MANUFAUTHORITYFILE")
            Case 0  '顯示
                DMANUFAUTHORITYFILE.Visible = False
                DMANUFAUTHORITYFILE.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DMANUFAUTHORITYFILERqd", "DMANUFAUTHORITYFILE", "異常：需輸入製造授權書")
                DMANUFAUTHORITYFILE.Visible = True
                DMANUFAUTHORITYFILE.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DMANUFAUTHORITYFILE.Visible = True
                DMANUFAUTHORITYFILE.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DMANUFAUTHORITYFILE.Visible = False
        End Select
        If pPost = "New" Then LMANUFAUTHORITYFILE.Visible = False

        '報價單
        Select Case FindFieldInf("1-FORCASTFILE")
            Case 0  '顯示
                DFORCASTFILE.Visible = False
                DFORCASTFILE.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DForCastFileRqd", "DForCastFileFile", "異常：需輸入報價單")
                DFORCASTFILE.Visible = True
                DFORCASTFILE.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DFORCASTFILE.Visible = True
                DFORCASTFILE.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DFORCASTFILE.Visible = False
        End Select
        If pPost = "New" Then LFORCASTFILE.Visible = False


        Select Case FindFieldInf("1-MANUOUTPRICE")
            Case 0  '顯示
                DMANUOUTPRICE.BackColor = Color.LightGray
                DMANUOUTPRICE.Visible = True
                DMANUOUTPRICE.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DMANUOUTPRICE.Visible = True
                DMANUOUTPRICE.BackColor = Color.GreenYellow
                DMANUOUTPRICE.ReadOnly = False
                ShowRequiredFieldValidator("DManuOutPriceRqd", "DManuOutPrice", "異常：外注加工費")
            Case 2  '修改
                DMANUOUTPRICE.Visible = True
                DMANUOUTPRICE.BackColor = Color.Yellow
                DMANUOUTPRICE.ReadOnly = False
            Case Else   '隱藏
                DMANUOUTPRICE.Visible = False
        End Select
        If pPost = "New" Then DMANUOUTPRICE.Text = ""

        '----系統自動顯示---------------------------------------------    
        '日期
        DAPPDATE.BackColor = Color.LightGray
        DAPPDATE.Visible = True
        DAPPDATE.Attributes.Add("readonly", "true")
        If pPost = "New" Then
            DAPPDATE.Text = CDate(NowDateTime).ToString("yyyy/MM/dd")
        End If
        '部門&擔當
        '部門
        DAPPDEPT.BackColor = Color.LightGray
        DAPPDEPT.Visible = True
        DAPPDEPT.Attributes.Add("readonly", "true")
        '擔當
        DAPPPER.BackColor = Color.LightGray
        DAPPPER.Visible = True
        DAPPPER.Attributes.Add("readonly", "true")
        If pPost = "New" Then
            Dim sql As String = "Select e.Com_Code,e.Com_Name,e.Dep_Code,e.Dep_Name,e.[ID],e.[Name],e.Job_Title_Code,e.Job_Title from M_Users u inner join M_Emp e on u.EmpID = e.[ID] and u.DepoID = e.Com_Code where u.UserID='" & wApplyID & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(sql)
            If dt.Rows.Count > 0 Then
                DAPPDEPT.Text = dt.Rows(0)("Dep_Name").ToString.Trim
                DAPPPER.Text = dt.Rows(0)("Name").ToString.Trim
            End If
        End If
        '-----------------------------------------------------------------
        '-- 製造委託
        '-----------------------------------------------------------------
        '開發主題
        DDEVTITLE.BackColor = Color.LightGray
        DDEVTITLE.Visible = True
        DDEVTITLE.Attributes.Add("readonly", "true")
        '開發NO
        DDEVNO.BackColor = Color.LightGray
        DDEVNO.Visible = True
        DDEVNO.Attributes.Add("readonly", "true")
        'CODE-NO
        DCODENO.BackColor = Color.LightGray
        DCODENO.Visible = True
        DCODENO.Attributes.Add("readonly", "true")
        '發行日
        DISSDATE.BackColor = Color.LightGray
        DISSDATE.Visible = True
        DISSDATE.Attributes.Add("readonly", "true")
        '開發擔當
        DDEVPER1.BackColor = Color.LightGray
        DDEVPER1.Visible = True
        DDEVPER1.Attributes.Add("readonly", "true")
      
        '上止
        DUPSTK1.BackColor = Color.LightGray
        DUPSTK1.Visible = True
        DUPSTK1.Attributes.Add("readonly", "true")
        '下止
        DLOSTK1.BackColor = Color.LightGray
        DLOSTK1.Visible = True
        DLOSTK1.Attributes.Add("readonly", "true")
        '開具
        DOPPART1.BackColor = Color.LightGray
        DOPPART1.Visible = True
        DOPPART1.Attributes.Add("readonly", "true")
        '布帶
        DTASPEC.BackColor = Color.LightGray
        DTASPEC.Visible = True
        DTASPEC.Attributes.Add("readonly", "true")
        '鏈齒顏色
        DECOL1.BackColor = Color.LightGray
        DECOL1.Visible = True
        DECOL1.Attributes.Add("readonly", "true")
        '丸紐
        DCCOL1.BackColor = Color.LightGray
        DCCOL1.Visible = True
        DCCOL1.Attributes.Add("readonly", "true")
        '縫工線
        DTHSPEC.BackColor = Color.LightGray
        DTHSPEC.Visible = True
        DTHSPEC.Attributes.Add("readonly", "true")
        '企-長度
        DPLEN1.BackColor = Color.LightGray
        DPLEN1.Visible = True
        DPLEN1.Attributes.Add("readonly", "true")
        '企-數量
        DPQTY1.BackColor = Color.LightGray
        DPQTY1.Visible = True
        DPQTY1.Attributes.Add("readonly", "true")
        'EA-長度
        DEALEN1.BackColor = Color.LightGray
        DEALEN1.Visible = True
        DEALEN1.Attributes.Add("readonly", "true")
        'EA-數量
        DEAQTY1.BackColor = Color.LightGray
        DEAQTY1.Visible = True
        DEAQTY1.Attributes.Add("readonly", "true")
        '內外製
        DMANUFTYPE.BackColor = Color.LightGray
        DMANUFTYPE.Visible = True
        DMANUFTYPE.Attributes.Add("readonly", "true")
        '示意圖檔
        Select Case FindFieldInf("2-HINTFILE")
            Case 0  '顯示
                DHINTFILE.Visible = False
                DHINTFILE.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DHINTFILERqd", "DHINTFILE", "異常：需輸入示意圖檔")
                DHINTFILE.Visible = True
                DHINTFILE.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DHINTFILE.Visible = True
                DHINTFILE.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DHINTFILE.Visible = False
        End Select
        If pPost = "New" Then LHINTFILE.Visible = False
        'OPCAT
        If pPost = "New" Then
            Select Case FindFieldInf("2-OPCAT")
                Case 0  '顯示
                Case 1  '修改+檢查
                Case 2  '修改
                    DOP40.Visible = True
                    DOP50.Visible = True
                    DOP60.Visible = True
                    DOP70.Visible = True
                    DOP80.Visible = True
                    DOP90.Visible = True
                    DOP100.Visible = True
                    DOP110.Visible = True
                Case Else   '隱藏
                    DOP40.Visible = False
                    DOP50.Visible = False
                    DOP60.Visible = False
                    DOP70.Visible = False
                    DOP80.Visible = False
                    DOP90.Visible = False
                    DOP100.Visible = False
                    DOP110.Visible = False
            End Select
        End If
        'OP1
        HOP1.BackColor = Color.LightGray
        HOP1.Visible = True
        HOP1.Attributes.Add("readonly", "true")
        DOP1PER.BackColor = Color.LightGray
        DOP1PER.Visible = True
        DOP1PER.Attributes.Add("readonly", "true")
        DOP1BTIME.BackColor = Color.LightGray
        DOP1BTIME.Visible = True
        DOP1BTIME.Attributes.Add("readonly", "true")
        DOP1BHOURS.BackColor = Color.LightGray
        DOP1BHOURS.Visible = True
        DOP1BHOURS.Attributes.Add("readonly", "true")
        DOP1ATIME.BackColor = Color.LightGray
        DOP1ATIME.Visible = True
        DOP1ATIME.Attributes.Add("readonly", "true")
        DOP1AHOURS.BackColor = Color.LightGray
        DOP1AHOURS.Visible = True
        DOP1AHOURS.Attributes.Add("readonly", "true")
        DOP1CON.BackColor = Color.LightGray
        DOP1CON.Visible = True
        DOP1CON.Attributes.Add("readonly", "true")
        'OP1--遲納原因類別
        Select Case FindFieldInf("2-DELAYCAT1")
            Case 0  '顯示
                DOP1DELAYC1.BackColor = Color.LightGray
                DOP1DELAYC1.Visible = True
                DOP1DELAYC2.BackColor = Color.LightGray
                DOP1DELAYC2.Visible = True
                DOP1REM.BackColor = Color.LightGray
                DOP1REM.Visible = True
                DOP1REM.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DOP1DELAYC1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP1DELAYC1Rqd", "DOP1DELAYC1", "異常：需輸入遲納原因類別-1")
                DOP1DELAYC1.Visible = True
                DOP1DELAYC2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP1DELAYC2Rqd", "DOP1DELAYC2", "異常：需輸入遲納原因類別-2")
                DOP1DELAYC2.Visible = True
                DOP1REM.Visible = True
                DOP1REM.BackColor = Color.GreenYellow
                DOP1REM.ReadOnly = False
                ShowRequiredFieldValidator("DOP1REMRqd", "DOP1REM", "異常：需輸入遲納原因內容")
            Case 2  '修改
                DOP1DELAYC1.BackColor = Color.Yellow
                DOP1DELAYC1.Visible = True
                DOP1DELAYC2.BackColor = Color.Yellow
                DOP1DELAYC2.Visible = True
                DOP1REM.Visible = True
                DOP1REM.BackColor = Color.Yellow
                DOP1REM.ReadOnly = False
            Case Else   '隱藏
                DOP1DELAYC1.Visible = False
                DOP1DELAYC2.Visible = False
                DOP1REM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("2-DELAYCAT1"), "OP1DELAYC1", "ZZZZZZ")
        If pPost = "New" Then SetFieldData(FindFieldInf("2-DELAYCAT1"), "OP1DELAYC2", "ZZZZZZ")
        If pPost = "New" Then DOP1REM.Text = ""
        'OP2
        HOP2.BackColor = Color.LightGray
        HOP2.Visible = True
        HOP2.Attributes.Add("readonly", "true")
        DOP2PER.BackColor = Color.LightGray
        DOP2PER.Visible = True
        DOP2PER.Attributes.Add("readonly", "true")
        DOP2BTIME.BackColor = Color.LightGray
        DOP2BTIME.Visible = True
        DOP2BTIME.Attributes.Add("readonly", "true")
        DOP2BHOURS.BackColor = Color.LightGray
        DOP2BHOURS.Visible = True
        DOP2BHOURS.Attributes.Add("readonly", "true")
        DOP2ATIME.BackColor = Color.LightGray
        DOP2ATIME.Visible = True
        DOP2ATIME.Attributes.Add("readonly", "true")
        DOP2AHOURS.BackColor = Color.LightGray
        DOP2AHOURS.Visible = True
        DOP2AHOURS.Attributes.Add("readonly", "true")
        DOP2CON.BackColor = Color.LightGray
        DOP2CON.Visible = True
        DOP2CON.Attributes.Add("readonly", "true")
        'OP2--遲納原因類別
        Select Case FindFieldInf("2-DELAYCAT2")
            Case 0  '顯示
                DOP2DELAYC1.BackColor = Color.LightGray
                DOP2DELAYC1.Visible = True
                DOP2DELAYC2.BackColor = Color.LightGray
                DOP2DELAYC2.Visible = True
                DOP2REM.BackColor = Color.LightGray
                DOP2REM.Visible = True
                DOP2REM.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DOP2DELAYC1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP2DELAYC1Rqd", "DOP2DELAYC1", "異常：需輸入遲納原因類別-1")
                DOP2DELAYC1.Visible = True
                DOP2DELAYC2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP2DELAYC2Rqd", "DOP2DELAYC2", "異常：需輸入遲納原因類別-2")
                DOP2DELAYC2.Visible = True
                DOP2REM.Visible = True
                DOP2REM.BackColor = Color.GreenYellow
                DOP2REM.ReadOnly = False
                ShowRequiredFieldValidator("DOP2REMRqd", "DOP2REM", "異常：需輸入遲納原因內容")
            Case 2  '修改
                DOP2DELAYC1.BackColor = Color.Yellow
                DOP2DELAYC1.Visible = True
                DOP2DELAYC2.BackColor = Color.Yellow
                DOP2DELAYC2.Visible = True
                DOP2REM.Visible = True
                DOP2REM.BackColor = Color.Yellow
                DOP2REM.ReadOnly = False
            Case Else   '隱藏
                DOP2DELAYC1.Visible = False
                DOP2DELAYC2.Visible = False
                DOP2REM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("2-DELAYCAT2"), "OP2DELAYC1", "ZZZZZZ")
        If pPost = "New" Then SetFieldData(FindFieldInf("2-DELAYCAT2"), "OP2DELAYC2", "ZZZZZZ")
        If pPost = "New" Then DOP2REM.Text = ""
        'OP3
        HOP3.BackColor = Color.LightGray
        HOP3.Visible = True
        HOP3.Attributes.Add("readonly", "true")
        DOP3PER.BackColor = Color.LightGray
        DOP3PER.Visible = True
        DOP3PER.Attributes.Add("readonly", "true")
        DOP3BTIME.BackColor = Color.LightGray
        DOP3BTIME.Visible = True
        DOP3BTIME.Attributes.Add("readonly", "true")
        DOP3BHOURS.BackColor = Color.LightGray
        DOP3BHOURS.Visible = True
        DOP3BHOURS.Attributes.Add("readonly", "true")
        DOP3ATIME.BackColor = Color.LightGray
        DOP3ATIME.Visible = True
        DOP3ATIME.Attributes.Add("readonly", "true")
        DOP3AHOURS.BackColor = Color.LightGray
        DOP3AHOURS.Visible = True
        DOP3AHOURS.Attributes.Add("readonly", "true")
        DOP3CON.BackColor = Color.LightGray
        DOP3CON.Visible = True
        DOP3CON.Attributes.Add("readonly", "true")
        'OP3--遲納原因類別
        Select Case FindFieldInf("2-DELAYCAT3")
            Case 0  '顯示
                DOP3DELAYC1.BackColor = Color.LightGray
                DOP3DELAYC1.Visible = True
                DOP3DELAYC2.BackColor = Color.LightGray
                DOP3DELAYC2.Visible = True
                DOP3REM.BackColor = Color.LightGray
                DOP3REM.Visible = True
                DOP3REM.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DOP3DELAYC1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP3DELAYC1Rqd", "DOP3DELAYC1", "異常：需輸入遲納原因類別-1")
                DOP3DELAYC1.Visible = True
                DOP3DELAYC2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP3DELAYC2Rqd", "DOP3DELAYC2", "異常：需輸入遲納原因類別-2")
                DOP3DELAYC2.Visible = True
                DOP3REM.Visible = True
                DOP3REM.BackColor = Color.GreenYellow
                DOP3REM.ReadOnly = False
                ShowRequiredFieldValidator("DOP3REMRqd", "DOP3REM", "異常：需輸入遲納原因內容")
            Case 2  '修改
                DOP3DELAYC1.BackColor = Color.Yellow
                DOP3DELAYC1.Visible = True
                DOP3DELAYC2.BackColor = Color.Yellow
                DOP3DELAYC2.Visible = True
                DOP3REM.Visible = True
                DOP3REM.BackColor = Color.Yellow
                DOP3REM.ReadOnly = False
            Case Else   '隱藏
                DOP3DELAYC1.Visible = False
                DOP3DELAYC2.Visible = False
                DOP3REM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("2-DELAYCAT3"), "OP3DELAYC1", "ZZZZZZ")
        If pPost = "New" Then SetFieldData(FindFieldInf("2-DELAYCAT3"), "OP3DELAYC2", "ZZZZZZ")
        If pPost = "New" Then DOP3REM.Text = ""
        'OP4
        HOP4.BackColor = Color.LightGray
        HOP4.Visible = True
        HOP4.Attributes.Add("readonly", "true")
        DOP4PER.BackColor = Color.LightGray
        DOP4PER.Visible = True
        DOP4PER.Attributes.Add("readonly", "true")
        DOP4BTIME.BackColor = Color.LightGray
        DOP4BTIME.Visible = True
        DOP4BTIME.Attributes.Add("readonly", "true")
        DOP4BHOURS.BackColor = Color.LightGray
        DOP4BHOURS.Visible = True
        DOP4BHOURS.Attributes.Add("readonly", "true")
        DOP4ATIME.BackColor = Color.LightGray
        DOP4ATIME.Visible = True
        DOP4ATIME.Attributes.Add("readonly", "true")
        DOP4AHOURS.BackColor = Color.LightGray
        DOP4AHOURS.Visible = True
        DOP4AHOURS.Attributes.Add("readonly", "true")
        DOP4CON.BackColor = Color.LightGray
        DOP4CON.Visible = True
        DOP4CON.Attributes.Add("readonly", "true")
        'OP4--遲納原因類別
        Select Case FindFieldInf("2-DELAYCAT4")
            Case 0  '顯示
                DOP4DELAYC1.BackColor = Color.LightGray
                DOP4DELAYC1.Visible = True
                DOP4DELAYC2.BackColor = Color.LightGray
                DOP4DELAYC2.Visible = True
                DOP4REM.BackColor = Color.LightGray
                DOP4REM.Visible = True
                DOP4REM.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DOP4DELAYC1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP4DELAYC1Rqd", "DOP4DELAYC1", "異常：需輸入遲納原因類別-1")
                DOP4DELAYC1.Visible = True
                DOP4DELAYC2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP4DELAYC2Rqd", "DOP4DELAYC2", "異常：需輸入遲納原因類別-2")
                DOP4DELAYC2.Visible = True
                DOP4REM.Visible = True
                DOP4REM.BackColor = Color.GreenYellow
                DOP4REM.ReadOnly = False
                ShowRequiredFieldValidator("DOP4REMRqd", "DOP4REM", "異常：需輸入遲納原因內容")
            Case 2  '修改
                DOP4DELAYC1.BackColor = Color.Yellow
                DOP4DELAYC1.Visible = True
                DOP4DELAYC2.BackColor = Color.Yellow
                DOP4DELAYC2.Visible = True
                DOP4REM.Visible = True
                DOP4REM.BackColor = Color.Yellow
                DOP4REM.ReadOnly = False
            Case Else   '隱藏
                DOP4DELAYC1.Visible = False
                DOP4DELAYC2.Visible = False
                DOP4REM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("2-DELAYCAT4"), "OP4DELAYC1", "ZZZZZZ")
        If pPost = "New" Then SetFieldData(FindFieldInf("2-DELAYCAT4"), "OP4DELAYC2", "ZZZZZZ")
        If pPost = "New" Then DOP4REM.Text = ""
        'OP5
        HOP5.BackColor = Color.LightGray
        HOP5.Visible = True
        HOP5.Attributes.Add("readonly", "true")
        DOP5PER.BackColor = Color.LightGray
        DOP5PER.Visible = True
        DOP5PER.Attributes.Add("readonly", "true")
        DOP5BTIME.BackColor = Color.LightGray
        DOP5BTIME.Visible = True
        DOP5BTIME.Attributes.Add("readonly", "true")
        DOP5BHOURS.BackColor = Color.LightGray
        DOP5BHOURS.Visible = True
        DOP5BHOURS.Attributes.Add("readonly", "true")
        DOP5ATIME.BackColor = Color.LightGray
        DOP5ATIME.Visible = True
        DOP5ATIME.Attributes.Add("readonly", "true")
        DOP5AHOURS.BackColor = Color.LightGray
        DOP5AHOURS.Visible = True
        DOP5AHOURS.Attributes.Add("readonly", "true")
        DOP5CON.BackColor = Color.LightGray
        DOP5CON.Visible = True
        DOP5CON.Attributes.Add("readonly", "true")
        'OP5--遲納原因類別
        Select Case FindFieldInf("2-DELAYCAT5")
            Case 0  '顯示
                DOP5DELAYC1.BackColor = Color.LightGray
                DOP5DELAYC1.Visible = True
                DOP5DELAYC2.BackColor = Color.LightGray
                DOP5DELAYC2.Visible = True
                DOP5REM.BackColor = Color.LightGray
                DOP5REM.Visible = True
                DOP5REM.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DOP5DELAYC1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP5DELAYC1Rqd", "DOP5DELAYC1", "異常：需輸入遲納原因類別-1")
                DOP5DELAYC1.Visible = True
                DOP5DELAYC2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP5DELAYC2Rqd", "DOP5DELAYC2", "異常：需輸入遲納原因類別-2")
                DOP5DELAYC2.Visible = True
                DOP5REM.Visible = True
                DOP5REM.BackColor = Color.GreenYellow
                DOP5REM.ReadOnly = False
                ShowRequiredFieldValidator("DOP5REMRqd", "DOP5REM", "異常：需輸入遲納原因內容")
            Case 2  '修改
                DOP5DELAYC1.BackColor = Color.Yellow
                DOP5DELAYC1.Visible = True
                DOP5DELAYC2.BackColor = Color.Yellow
                DOP5DELAYC2.Visible = True
                DOP5REM.Visible = True
                DOP5REM.BackColor = Color.Yellow
                DOP5REM.ReadOnly = False
            Case Else   '隱藏
                DOP5DELAYC1.Visible = False
                DOP5DELAYC2.Visible = False
                DOP5REM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("2-DELAYCAT5"), "OP5DELAYC1", "ZZZZZZ")
        If pPost = "New" Then SetFieldData(FindFieldInf("2-DELAYCAT5"), "OP5DELAYC2", "ZZZZZZ")
        If pPost = "New" Then DOP5REM.Text = ""
        'OP6
        HOP6.BackColor = Color.LightGray
        HOP6.Visible = True
        HOP6.Attributes.Add("readonly", "true")
        DOP6PER.BackColor = Color.LightGray
        DOP6PER.Visible = True
        DOP6PER.Attributes.Add("readonly", "true")
        DOP6BTIME.BackColor = Color.LightGray
        DOP6BTIME.Visible = True
        DOP6BTIME.Attributes.Add("readonly", "true")
        DOP6BHOURS.BackColor = Color.LightGray
        DOP6BHOURS.Visible = True
        DOP6BHOURS.Attributes.Add("readonly", "true")
        DOP6ATIME.BackColor = Color.LightGray
        DOP6ATIME.Visible = True
        DOP6ATIME.Attributes.Add("readonly", "true")
        DOP6AHOURS.BackColor = Color.LightGray
        DOP6AHOURS.Visible = True
        DOP6AHOURS.Attributes.Add("readonly", "true")
        DOP6CON.BackColor = Color.LightGray
        DOP6CON.Visible = True
        DOP6CON.Attributes.Add("readonly", "true")
        'OP6--遲納原因類別
        Select Case FindFieldInf("2-DELAYCAT6")
            Case 0  '顯示
                DOP6DELAYC1.BackColor = Color.LightGray
                DOP6DELAYC1.Visible = True
                DOP6DELAYC2.BackColor = Color.LightGray
                DOP6DELAYC2.Visible = True
                DOP6REM.BackColor = Color.LightGray
                DOP6REM.Visible = True
                DOP6REM.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DOP6DELAYC1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP6DELAYC1Rqd", "DOP6DELAYC1", "異常：需輸入遲納原因類別-1")
                DOP6DELAYC1.Visible = True
                DOP6DELAYC2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP6DELAYC2Rqd", "DOP6DELAYC2", "異常：需輸入遲納原因類別-2")
                DOP6DELAYC2.Visible = True
                DOP6REM.Visible = True
                DOP6REM.BackColor = Color.GreenYellow
                DOP6REM.ReadOnly = False
                ShowRequiredFieldValidator("DOP6REMRqd", "DOP6REM", "異常：需輸入遲納原因內容")
            Case 2  '修改
                DOP6DELAYC1.BackColor = Color.Yellow
                DOP6DELAYC1.Visible = True
                DOP6DELAYC2.BackColor = Color.Yellow
                DOP6DELAYC2.Visible = True
                DOP6REM.Visible = True
                DOP6REM.BackColor = Color.Yellow
                DOP6REM.ReadOnly = False
            Case Else   '隱藏
                DOP6DELAYC1.Visible = False
                DOP6DELAYC2.Visible = False
                DOP6REM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("2-DELAYCAT6"), "OP6DELAYC1", "ZZZZZZ")
        If pPost = "New" Then SetFieldData(FindFieldInf("2-DELAYCAT6"), "OP6DELAYC2", "ZZZZZZ")
        If pPost = "New" Then DOP6REM.Text = ""
        'OP7
        HOP7.BackColor = Color.LightGray
        HOP7.Visible = True
        HOP7.Attributes.Add("readonly", "true")
        DOP7PER.BackColor = Color.LightGray
        DOP7PER.Visible = True
        DOP7PER.Attributes.Add("readonly", "true")
        DOP7BTIME.BackColor = Color.LightGray
        DOP7BTIME.Visible = True
        DOP7BTIME.Attributes.Add("readonly", "true")
        DOP7BHOURS.BackColor = Color.LightGray
        DOP7BHOURS.Visible = True
        DOP7BHOURS.Attributes.Add("readonly", "true")
        DOP7ATIME.BackColor = Color.LightGray
        DOP7ATIME.Visible = True
        DOP7ATIME.Attributes.Add("readonly", "true")
        DOP7AHOURS.BackColor = Color.LightGray
        DOP7AHOURS.Visible = True
        DOP7AHOURS.Attributes.Add("readonly", "true")
        DOP7CON.BackColor = Color.LightGray
        DOP7CON.Visible = True
        DOP7CON.Attributes.Add("readonly", "true")
        'OP7--遲納原因類別
        Select Case FindFieldInf("2-DELAYCAT7")
            Case 0  '顯示
                DOP7DELAYC1.BackColor = Color.LightGray
                DOP7DELAYC1.Visible = True
                DOP7DELAYC2.BackColor = Color.LightGray
                DOP7DELAYC2.Visible = True
                DOP7REM.BackColor = Color.LightGray
                DOP7REM.Visible = True
                DOP7REM.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DOP7DELAYC1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP7DELAYC1Rqd", "DOP7DELAYC1", "異常：需輸入遲納原因類別-1")
                DOP7DELAYC1.Visible = True
                DOP7DELAYC2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP7DELAYC2Rqd", "DOP7DELAYC2", "異常：需輸入遲納原因類別-2")
                DOP7DELAYC2.Visible = True
                DOP7REM.Visible = True
                DOP7REM.BackColor = Color.GreenYellow
                DOP7REM.ReadOnly = False
                ShowRequiredFieldValidator("DOP7REMRqd", "DOP7REM", "異常：需輸入遲納原因內容")
            Case 2  '修改
                DOP7DELAYC1.BackColor = Color.Yellow
                DOP7DELAYC1.Visible = True
                DOP7DELAYC2.BackColor = Color.Yellow
                DOP7DELAYC2.Visible = True
                DOP7REM.Visible = True
                DOP7REM.BackColor = Color.Yellow
                DOP7REM.ReadOnly = False
            Case Else   '隱藏
                DOP7DELAYC1.Visible = False
                DOP7DELAYC2.Visible = False
                DOP7REM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("2-DELAYCAT7"), "OP7DELAYC1", "ZZZZZZ")
        If pPost = "New" Then SetFieldData(FindFieldInf("2-DELAYCAT7"), "OP7DELAYC2", "ZZZZZZ")
        If pPost = "New" Then DOP7REM.Text = ""
        'OP8
        HOP8.BackColor = Color.LightGray
        HOP8.Visible = True
        HOP8.Attributes.Add("readonly", "true")
        DOP8PER.BackColor = Color.LightGray
        DOP8PER.Visible = True
        DOP8PER.Attributes.Add("readonly", "true")
        DOP8BTIME.BackColor = Color.LightGray
        DOP8BTIME.Visible = True
        DOP8BTIME.Attributes.Add("readonly", "true")
        DOP8BHOURS.BackColor = Color.LightGray
        DOP8BHOURS.Visible = True
        DOP8BHOURS.Attributes.Add("readonly", "true")
        DOP8ATIME.BackColor = Color.LightGray
        DOP8ATIME.Visible = True
        DOP8ATIME.Attributes.Add("readonly", "true")
        DOP8AHOURS.BackColor = Color.LightGray
        DOP8AHOURS.Visible = True
        DOP8AHOURS.Attributes.Add("readonly", "true")
        DOP8CON.BackColor = Color.LightGray
        DOP8CON.Visible = True
        DOP8CON.Attributes.Add("readonly", "true")
        'OP8--遲納原因類別
        Select Case FindFieldInf("2-DELAYCAT8")
            Case 0  '顯示
                DOP8DELAYC1.BackColor = Color.LightGray
                DOP8DELAYC1.Visible = True
                DOP8DELAYC2.BackColor = Color.LightGray
                DOP8DELAYC2.Visible = True
                DOP8REM.BackColor = Color.LightGray
                DOP8REM.Visible = True
                DOP8REM.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DOP8DELAYC1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP8DELAYC1Rqd", "DOP8DELAYC1", "異常：需輸入遲納原因類別-1")
                DOP8DELAYC1.Visible = True
                DOP8DELAYC2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP8DELAYC2Rqd", "DOP8DELAYC2", "異常：需輸入遲納原因類別-2")
                DOP8DELAYC2.Visible = True
                DOP8REM.Visible = True
                DOP8REM.BackColor = Color.GreenYellow
                DOP8REM.ReadOnly = False
                ShowRequiredFieldValidator("DOP8REMRqd", "DOP8REM", "異常：需輸入遲納原因內容")
            Case 2  '修改
                DOP8DELAYC1.BackColor = Color.Yellow
                DOP8DELAYC1.Visible = True
                DOP8DELAYC2.BackColor = Color.Yellow
                DOP8DELAYC2.Visible = True
                DOP8REM.Visible = True
                DOP8REM.BackColor = Color.Yellow
                DOP8REM.ReadOnly = False
            Case Else   '隱藏
                DOP8DELAYC1.Visible = False
                DOP8DELAYC2.Visible = False
                DOP8REM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData(FindFieldInf("2-DELAYCAT8"), "OP8DELAYC1", "ZZZZZZ")
        If pPost = "New" Then SetFieldData(FindFieldInf("2-DELAYCAT8"), "OP8DELAYC2", "ZZZZZZ")
        If pPost = "New" Then DOP8REM.Text = ""
        '延遲--背景色變更
        If DBEndTime.Text < NowDateTime Then
            Select Case wStep
                Case 40
                    HOP1.BackColor = Color.LightPink
                    DOP1PER.BackColor = Color.LightPink
                    DOP1BTIME.BackColor = Color.LightPink
                    DOP1BHOURS.BackColor = Color.LightPink
                    DOP1ATIME.BackColor = Color.LightPink
                    DOP1AHOURS.BackColor = Color.LightPink
                    DOP1CON.BackColor = Color.LightPink
                Case 50
                    HOP2.BackColor = Color.LightPink

                    DOP2PER.BackColor = Color.LightPink
                    DOP2BTIME.BackColor = Color.LightPink
                    DOP2BHOURS.BackColor = Color.LightPink
                    DOP2ATIME.BackColor = Color.LightPink
                    DOP2AHOURS.BackColor = Color.LightPink
                    DOP2CON.BackColor = Color.LightPink
                Case 60
                    HOP3.BackColor = Color.LightPink
                    DOP3PER.BackColor = Color.LightPink
                    DOP3BTIME.BackColor = Color.LightPink
                    DOP3BHOURS.BackColor = Color.LightPink
                    DOP3ATIME.BackColor = Color.LightPink
                    DOP3AHOURS.BackColor = Color.LightPink
                    DOP3CON.BackColor = Color.LightPink
                Case 70
                    HOP4.BackColor = Color.LightPink
                    DOP4PER.BackColor = Color.LightPink
                    DOP4BTIME.BackColor = Color.LightPink
                    DOP4BHOURS.BackColor = Color.LightPink
                    DOP4ATIME.BackColor = Color.LightPink
                    DOP4AHOURS.BackColor = Color.LightPink
                    DOP4CON.BackColor = Color.LightPink
                Case 80
                    HOP5.BackColor = Color.LightPink
                    DOP5PER.BackColor = Color.LightPink
                    DOP5BTIME.BackColor = Color.LightPink
                    DOP5BHOURS.BackColor = Color.LightPink
                    DOP5ATIME.BackColor = Color.LightPink
                    DOP5AHOURS.BackColor = Color.LightPink
                    DOP5CON.BackColor = Color.LightPink
                Case 90
                    HOP6.BackColor = Color.LightPink
                    DOP6PER.BackColor = Color.LightPink
                    DOP6BTIME.BackColor = Color.LightPink
                    DOP6BHOURS.BackColor = Color.LightPink
                    DOP6ATIME.BackColor = Color.LightPink
                    DOP6AHOURS.BackColor = Color.LightPink
                    DOP6CON.BackColor = Color.LightPink
                Case 100
                    HOP7.BackColor = Color.LightPink
                    DOP7PER.BackColor = Color.LightPink
                    DOP7BTIME.BackColor = Color.LightPink
                    DOP7BHOURS.BackColor = Color.LightPink
                    DOP7ATIME.BackColor = Color.LightPink
                    DOP7AHOURS.BackColor = Color.LightPink
                    DOP7CON.BackColor = Color.LightPink
                Case 110
                    HOP8.BackColor = Color.LightPink
                    DOP8PER.BackColor = Color.LightPink
                    DOP8BTIME.BackColor = Color.LightPink
                    DOP8BHOURS.BackColor = Color.LightPink
                    DOP8ATIME.BackColor = Color.LightPink
                    DOP8AHOURS.BackColor = Color.LightPink
                    DOP8CON.BackColor = Color.LightPink
                Case Else
            End Select
        End If
        '-----------------------------------------------------------------
        '-- 開發見本
        '-----------------------------------------------------------------
        '樣品圖檔
        Select Case FindFieldInf("3-SAMPLEFILE")
            Case 0  '顯示
                D3SAMPLEFILE.Visible = False
                D3SAMPLEFILE.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("D3SAMPLEFILERqd", "D3SAMPLEFILE", "異常：需輸入樣品圖檔")
                D3SAMPLEFILE.Visible = True
                D3SAMPLEFILE.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                D3SAMPLEFILE.Visible = True
                D3SAMPLEFILE.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                D3SAMPLEFILE.Visible = False
        End Select
        If pPost = "New" Then LSAMPLEFILE.Visible = False
        '注意事項
        Select Case FindFieldInf("3-OTHER")
            Case 0  '顯示
                D3OTHER.BackColor = Color.LightGray
                D3OTHER.Visible = True
                D3OTHER.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                D3OTHER.Visible = True
                D3OTHER.BackColor = Color.GreenYellow
                D3OTHER.ReadOnly = False
                ShowRequiredFieldValidator("D3OTHERRqd", "D3OTHER", "異常：需輸入注意事項")
            Case 2  '修改
                D3OTHER.Visible = True
                D3OTHER.BackColor = Color.Yellow
                D3OTHER.ReadOnly = False
            Case Else   '隱藏
                D3OTHER.Visible = False
        End Select
        If pPost = "New" Then D3OTHER.Text = ""
        '品管-吋法
        Select Case FindFieldInf("3-QCFILE")
            Case 0  '顯示
                D3QCFILE1.Visible = False
                D3QCFILE1.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("D3QCFILE1Rqd", "D3QCFILE1", "異常：需輸入品管-吋法")
                D3QCFILE1.Visible = True
                D3QCFILE1.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                D3QCFILE1.Visible = True
                D3QCFILE1.Style.Add("BACKGROUND-COLOR", "Yellow")

            Case Else   '隱藏
                D3QCFILE1.Visible = False
        End Select
        If pPost = "New" Then L3QCFILE1.Visible = False
        '品管-強度
        Select Case FindFieldInf("3-QCFILE")
            Case 0  '顯示
                D3QCFILE2.Visible = False
                D3QCFILE2.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("D3QCFILE2Rqd", "D3QCFILE2", "異常：需輸入品管-強度")
                D3QCFILE2.Visible = True
                D3QCFILE2.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                D3QCFILE2.Visible = True
                D3QCFILE2.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                D3QCFILE2.Visible = False
        End Select
        If pPost = "New" Then L3QCFILE2.Visible = False
        '品管-往覆測試
        Select Case FindFieldInf("3-QCFILE")
            Case 0  '顯示
                D3QCFILE3.Visible = False
                D3QCFILE3.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("D3QCFILE3Rqd", "D3QCFILE3", "異常：需輸入品管-往覆測試")
                D3QCFILE3.Visible = True
                D3QCFILE3.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                D3QCFILE3.Visible = True
                D3QCFILE3.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                D3QCFILE3.Visible = False
        End Select
        If pPost = "New" Then L3QCFILE3.Visible = False
        '品管-式樣／組織
        Select Case FindFieldInf("3-QCFILE")
            Case 0  '顯示
                D3QCFILE4.Visible = False
                D3QCFILE4.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("D3QCFILE4Rqd", "D3QCFILE4", "異常：需輸入品管-式樣／組織")
                D3QCFILE4.Visible = True
                D3QCFILE4.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                D3QCFILE4.Visible = True
                D3QCFILE4.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                D3QCFILE4.Visible = False
        End Select
        If pPost = "New" Then L3QCFILE4.Visible = False
        '品管-其他
        Select Case FindFieldInf("3-QCFILE")
            Case 0  '顯示
                D3QCFILE5.Visible = False
                D3QCFILE5.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("D3QCFILE5Rqd", "D3QCFILE5", "異常：需輸入品管-其他")
                D3QCFILE5.Visible = True
                D3QCFILE5.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                D3QCFILE5.Visible = True
                D3QCFILE5.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                D3QCFILE5.Visible = False
        End Select
        If pPost = "New" Then L3QCFILE5.Visible = False
        '客戶
        Select Case FindFieldInf("3-APPBUYER")
            Case 0  '顯示
                D3APPBUYER.BackColor = Color.LightGray
                D3APPBUYER.ReadOnly = True
                D3APPBUYER.Visible = True
                BCREATE.Visible = False
            Case 1  '修改+檢查
                D3APPBUYER.BackColor = Color.LightGray
                D3APPBUYER.ReadOnly = True
                D3APPBUYER.Visible = True
                BCREATE.Visible = True
                ShowRequiredFieldValidator("D3APPBUYERRqd", "D3APPBUYER", "異常：需輸入客戶")
            Case 2  '修改
                D3APPBUYER.BackColor = Color.LightGray
                D3APPBUYER.ReadOnly = True
                D3APPBUYER.Visible = True
                BCREATE.Visible = True
            Case Else   '隱藏
                D3APPBUYER.Visible = False
                BCREATE.Visible = False
        End Select
        If pPost = "New" Then D3APPBUYER.Text = ""
        '發行日
        D3DATE.BackColor = Color.LightGray
        D3DATE.Visible = True
        D3DATE.Attributes.Add("readonly", "true")
        'SIZE
        D3SIZENO.BackColor = Color.LightGray
        D3SIZENO.Visible = True
        D3SIZENO.Attributes.Add("readonly", "true")
        'ITEM
        D3ITEM.BackColor = Color.LightGray
        D3ITEM.Visible = True
        D3ITEM.Attributes.Add("readonly", "true")
        'TAPE(CODENO)
        D3CODENO.BackColor = Color.LightGray
        D3CODENO.Visible = True
        D3CODENO.Attributes.Add("readonly", "true")
        '布帶寬度
        D3TAWIDTH.BackColor = Color.LightGray
        D3TAWIDTH.Visible = True
        D3TAWIDTH.Attributes.Add("readonly", "true")
        '開發NO
        D3DEVNO.BackColor = Color.LightGray
        D3DEVNO.Visible = True
        D3DEVNO.Attributes.Add("readonly", "true")
        '開發期間
        D3DEVPRD.BackColor = Color.LightGray
        D3DEVPRD.Visible = True
        D3DEVPRD.Attributes.Add("readonly", "true")
        '布帶
        D3TACOL.BackColor = Color.LightGray
        D3TACOL.Visible = True
        D3TACOL.Attributes.Add("readonly", "true")
        '條紋線
        D3TALINE.BackColor = Color.LightGray
        D3TALINE.Visible = True
        D3TALINE.Attributes.Add("readonly", "true")
        '務齒
        D3ECOL.BackColor = Color.LightGray
        D3ECOL.Visible = True
        D3ECOL.Attributes.Add("readonly", "true")
        '丸紐
        D3CCOL.BackColor = Color.LightGray
        D3CCOL.Visible = True
        D3CCOL.Attributes.Add("readonly", "true")
        '縫工線
        D3THCOL.BackColor = Color.LightGray
        D3THCOL.Visible = True
        D3THCOL.Attributes.Add("readonly", "true")
        'WAVE'S
        D3TNLITEM.BackColor = Color.LightGray
        D3TNLITEM.Visible = True
        D3TNLITEM.Attributes.Add("readonly", "true")
        D3TNRITEM.BackColor = Color.LightGray
        D3TNRITEM.Visible = True
        D3TNRITEM.Attributes.Add("readonly", "true")
        D3TSLITEM.BackColor = Color.LightGray
        D3TSLITEM.Visible = True
        D3TSLITEM.Attributes.Add("readonly", "true")
        D3TSRITEM.BackColor = Color.LightGray
        D3TSRITEM.Visible = True
        D3TSRITEM.Attributes.Add("readonly", "true")
        D3TDLITEM.BackColor = Color.LightGray
        D3TDLITEM.Visible = True
        D3TDLITEM.Attributes.Add("readonly", "true")
        D3TDRITEM.BackColor = Color.LightGray
        D3TDRITEM.Visible = True
        D3TDRITEM.Attributes.Add("readonly", "true")
        D3CNITEM.BackColor = Color.LightGray
        D3CNITEM.Visible = True
        D3CNITEM.Attributes.Add("readonly", "true")
        D3CSITEM.BackColor = Color.LightGray
        D3CSITEM.Visible = True
        D3CSITEM.Attributes.Add("readonly", "true")
        D3CDITEM.BackColor = Color.LightGray
        D3CDITEM.Visible = True
        D3CDITEM.Attributes.Add("readonly", "true")
        D3CITEM.BackColor = Color.LightGray
        D3CITEM.Visible = True
        D3CITEM.Attributes.Add("readonly", "true")
        'D31Other.BackColor = Color.LightGray
        D31Other.Visible = True
        D31Other.Attributes.Add("readonly", "true")
        'D32Other.BackColor = Color.LightGray
        D32Other.Visible = True
        D32Other.Attributes.Add("readonly", "true")
        D3O1ITEM.BackColor = Color.LightGray
        D3O1ITEM.Visible = True
        D3O1ITEM.Attributes.Add("readonly", "true")
        D3O2ITEM.BackColor = Color.LightGray
        D3O2ITEM.Visible = True
        D3O2ITEM.Attributes.Add("readonly", "true")
        '承認-作成者
        D3WF1.BackColor = Color.Yellow
        D3WF1.Visible = True
        If pPost = "New" Then SetFieldData(FindFieldInf("3-FLOW"), "WF1", "ZZZZZZ")
        '承認-責任者
        D3WF2.BackColor = Color.Yellow
        D3WF2.Visible = True
        If pPost = "New" Then SetFieldData(FindFieldInf("3-FLOW"), "WF2", "ZZZZZZ")
        '承認-製造1
        D3WF3Name.BackColor = Color.Yellow
        D3WF3Name.Visible = True
        If pPost = "New" Then SetFieldData(FindFieldInf("3-FLOW"), "WF3NAME", "ZZZZZZ")
        '
        D3WF3.BackColor = Color.Yellow
        D3WF3.Visible = True
        If pPost = "New" Then SetFieldData(FindFieldInf("3-FLOW"), "WF3", "ZZZZZZ")
        '承認-製造2
        D3WF4Name.BackColor = Color.Yellow
        D3WF4Name.Visible = True
        If pPost = "New" Then SetFieldData(FindFieldInf("3-FLOW"), "WF4NAME", "ZZZZZZ")
        '
        D3WF4.BackColor = Color.Yellow
        D3WF4.Visible = True
        If pPost = "New" Then SetFieldData(FindFieldInf("3-FLOW"), "WF4", "ZZZZZZ")
        '承認-製造3
        D3WF5Name.BackColor = Color.Yellow
        D3WF5Name.Visible = True
        If pPost = "New" Then SetFieldData(FindFieldInf("3-FLOW"), "WF5NAME", "ZZZZZZ")
        '
        D3WF5.BackColor = Color.Yellow
        D3WF5.Visible = True
        If pPost = "New" Then SetFieldData(FindFieldInf("3-FLOW"), "WF5", "ZZZZZZ")
        '承認-製造4
        D3WF6Name.BackColor = Color.Yellow
        D3WF6Name.Visible = True
        If pPost = "New" Then SetFieldData(FindFieldInf("3-FLOW"), "WF6NAME", "ZZZZZZ")
        '
        D3WF6.BackColor = Color.Yellow
        D3WF6.Visible = True
        If pPost = "New" Then SetFieldData(FindFieldInf("3-FLOW"), "WF6", "ZZZZZZ")
        '承認-廠長
        D3WF7Name.BackColor = Color.Yellow
        D3WF7Name.Visible = True
        If pPost = "New" Then SetFieldData(FindFieldInf("3-FLOW"), "WF7NAME", "ZZZZZZ")
        '
        D3WF7.BackColor = Color.Yellow
        D3WF7.Visible = True
        If pPost = "New" Then SetFieldData(FindFieldInf("3-FLOW"), "WF7", "ZZZZZZ")
        '-----------------------------------------------------------------
        '-- 原單純
        '-----------------------------------------------------------------
        '目前不控制
    End Sub
    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pIdx As Integer, ByVal pFieldName As String, ByVal pName As String)
        Dim sql As String = ""
        '-----------------------------------------------------------------
        '-- 開發委託
        '-----------------------------------------------------------------
        '----基本-------------------------------------------------    
        'BUYER
        If pFieldName = "APPBUYER" Then
            DAPPBUYER.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAPPBUYER.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='700' and DKey='BUYER' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAPPBUYER.Items.Add(ListItem1)
                Next
            End If
        End If
        '用途
        If pFieldName = "USAGE" Then
            DUSAGE.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DUSAGE.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='USAGE' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DUSAGE.Items.Add(ListItem1)
                Next
            End If
        End If
        '需圖面
        If pFieldName = "NEEDMAP" Then
            DNEEDMAP.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DNEEDMAP.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='NEEDMAP' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DNEEDMAP.Items.Add(ListItem1)
                Next
            End If
        End If
        '----樣品-------------------------------------------------    
        '製品區分
        If pFieldName = "PRO" Then
            DPRO.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPRO.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='PRO' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPRO.Items.Add(ListItem1)
                Next
            End If
        End If
        '長度單位(企)
        If pFieldName = "PLENUN" Then
            DPLENUN.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPLENUN.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='LENUN' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPLENUN.Items.Add(ListItem1)
                Next
            End If
        End If
        '數量單位(企)
        If pFieldName = "PQTYUN" Then
            DPQTYUN.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPQTYUN.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='QTYUN' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPQTYUN.Items.Add(ListItem1)
                Next
            End If
        End If
        '長度單位(EA)
        If pFieldName = "EALENUN" Then
            DEALENUN.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DEALENUN.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='LENUN' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DEALENUN.Items.Add(ListItem1)
                Next
            End If
        End If
        '數量單位(EA)
        If pFieldName = "EAQTYUN" Then
            DEAQTYUN.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DEAQTYUN.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='QTYUN' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DEAQTYUN.Items.Add(ListItem1)
                Next
            End If
        End If
        '----開發-------------------------------------------------    
        '型別
        If pFieldName = "SIZENO" Then
            DSIZENO.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSIZENO.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='SIZE' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSIZENO.Items.Add(ListItem1)
                Next
            End If
        End If
        '鏈條形式
        If pFieldName = "ITEM" Then
            DITEM.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DITEM.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='CHAIN' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DITEM.Items.Add(ListItem1)
                Next
            End If
        End If
        '布帶
        If pFieldName = "TATYPE" Then
            DTATYPE.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTATYPE.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='TAPE' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTATYPE.Items.Add(ListItem1)
                Next
            End If
        End If
        '鏈齒顏色-SEL
        If pFieldName = "ECOLSEL" Then
            DECOLSEL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DECOLSEL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-A' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DECOLSEL.Items.Add(ListItem1)
                Next
            End If
        End If
        '丸紐-SEL
        If pFieldName = "CCOLSEL" Then
            DCCOLSEL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCCOLSEL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-C' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCCOLSEL.Items.Add(ListItem1)
                Next
            End If
        End If
        '布帶-色番(同)
        If pFieldName = "TACOL" Then
            DTACOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTACOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTACOL.Items.Add(ListItem1)
                Next
            End If
        End If
        '布帶-色番(左)
        If pFieldName = "TALCOL" Then
            DTALCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTALCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTALCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        '布帶-色番(右)
        If pFieldName = "TARCOL" Then
            DTARCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTARCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTARCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        '縫上-色番(同)
        If pFieldName = "THUPCOL" Then
            DTHUPCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHUPCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-D' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHUPCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        '縫上-色番(左左)
        If pFieldName = "THLUPCOL" Then
            DTHLUPCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHLUPCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-D' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHLUPCOL.Items.Add(ListItem1)
                Next
            End If
        End If

        '縫上-色番(左右)
        If pFieldName = "THLRUPCOL" Then
            DTHLRUPCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHLRUPCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-D' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHLRUPCOL.Items.Add(ListItem1)
                Next
            End If
        End If


        '縫上-色番(右左)
        If pFieldName = "THRUPCOL" Then
            DTHRUPCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHRUPCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-D' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHRUPCOL.Items.Add(ListItem1)
                Next
            End If
        End If

        '縫上-色番(右右)
        If pFieldName = "THRRUPCOL" Then
            DTHRRUPCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHRRUPCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-D' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHRRUPCOL.Items.Add(ListItem1)
                Next
            End If
        End If


        '縫下-色番(同)
        If pFieldName = "THLOCOL" Then
            DTHLOCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHLOCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-D' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHLOCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        '縫下-色番(左左)
        If pFieldName = "THLLOCOL" Then
            DTHLLOCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHLLOCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-D' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHLLOCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        '縫下-色番(左右)
        If pFieldName = "THLRLOCOL" Then
            DTHLRLOCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHLRLOCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-D' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHLRLOCOL.Items.Add(ListItem1)
                Next
            End If
        End If

        '縫下-色番(右左)
        If pFieldName = "THRLOCOL" Then
            DTHRLOCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHRLOCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-D' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHRLOCOL.Items.Add(ListItem1)
                Next
            End If
        End If

        '縫下-色番(右右)
        If pFieldName = "THRRLOCOL" Then
            DTHRRLOCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHRRLOCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-D' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHRRLOCOL.Items.Add(ListItem1)
                Next
            End If
        End If

        'X-色番
        If pFieldName = "XMCOL" Then
            DXMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DXMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DXMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'A-色番
        If pFieldName = "AMCOL" Then
            DAMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'B-色番
        If pFieldName = "BMCOL" Then
            DBMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'C-色番
        If pFieldName = "CMCOL" Then
            DCMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'D-色番
        If pFieldName = "DMCOL" Then
            DDMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'E-色番
        If pFieldName = "EMCOL" Then
            DEMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DEMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DEMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'F-色番
        If pFieldName = "FMCOL" Then
            DFMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DFMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DFMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'G-色番
        If pFieldName = "GMCOL" Then
            DGMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DGMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DGMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'H-色番
        If pFieldName = "HMCOL" Then
            DHMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DHMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DHMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        '緯紗-色番
        If pFieldName = "LYCOL" Then
            DLYCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLYCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLYCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'QC人員
        If pFieldName = "QCPEOPLE" Then
            DQCPEOPLE.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCPEOPLE.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='QCPEOPLE' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCPEOPLE.Items.Add(ListItem1)
                Next
            End If
        End If
        '製圖者
        If pFieldName = "MAKEMAP" Then
            DMAKEMAP.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMAKEMAP.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='MAKEMAP' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DMAKEMAP.Items.Add(ListItem1)
                Next
            End If
        End If
        '難易度
        If pFieldName = "LEVEL" Then
            DLEVEL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLEVEL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='LEVEL' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLEVEL.Items.Add(ListItem1)
                Next
            End If
        End If
        '-----------------------------------------------------------------
        '-- 製造委託
        '-----------------------------------------------------------------
        'OP1
        '----遲納原因類別1
        If pFieldName = "OP1DELAYC1" Then
            DOP1DELAYC1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP1DELAYC1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp "
                sql &= "Where Cat  = '2002' "
                sql &= "  and DKey = '" & "DELAYC1-" + HOP1.Text & "' "
                sql &= "Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP1DELAYC1.Items.Add(ListItem1)
                Next
            End If
        End If
        '----遲納原因類別2
        If pFieldName = "OP1DELAYC2" Then
            DOP1DELAYC2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP1DELAYC2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp "
                sql &= "Where Cat  = '2002' "
                sql &= "  and DKey = '" & "DELAYC2-" + HOP1.Text + "-" + DOP1DELAYC1.SelectedValue & "' "
                sql &= "Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP1DELAYC2.Items.Add(ListItem1)
                Next
            End If
        End If
        'OP2
        '----遲納原因類別1
        If pFieldName = "OP2DELAYC1" Then
            DOP2DELAYC1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP2DELAYC1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp "
                sql &= "Where Cat  = '2002' "
                sql &= "  and DKey = '" & "DELAYC1-" + HOP2.Text & "' "
                sql &= "Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP2DELAYC1.Items.Add(ListItem1)
                Next
            End If
        End If
        '----遲納原因類別2
        If pFieldName = "OP2DELAYC2" Then
            DOP2DELAYC2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP2DELAYC2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp "
                sql &= "Where Cat  = '2002' "
                sql &= "  and DKey = '" & "DELAYC2-" + HOP2.Text + "-" + DOP2DELAYC1.SelectedValue & "' "
                sql &= "Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP2DELAYC2.Items.Add(ListItem1)
                Next
            End If
        End If
        'OP3
        '----遲納原因類別1
        If pFieldName = "OP3DELAYC1" Then
            DOP3DELAYC1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP3DELAYC1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp "
                sql &= "Where Cat  = '2002' "
                sql &= "  and DKey = '" & "DELAYC1-" + HOP3.Text & "' "
                sql &= "Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP3DELAYC1.Items.Add(ListItem1)
                Next
            End If
        End If
        '----遲納原因類別2
        If pFieldName = "OP3DELAYC2" Then
            DOP3DELAYC2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP3DELAYC2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp "
                sql &= "Where Cat  = '2002' "
                sql &= "  and DKey = '" & "DELAYC2-" + HOP3.Text + "-" + DOP3DELAYC1.SelectedValue & "' "
                sql &= "Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP3DELAYC2.Items.Add(ListItem1)
                Next
            End If
        End If
        'OP4
        '----遲納原因類別1
        If pFieldName = "OP4DELAYC1" Then
            DOP4DELAYC1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP4DELAYC1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp "
                sql &= "Where Cat  = '2002' "
                sql &= "  and DKey = '" & "DELAYC1-" + HOP4.Text & "' "
                sql &= "Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP4DELAYC1.Items.Add(ListItem1)
                Next
            End If
        End If
        '----遲納原因類別2
        If pFieldName = "OP4DELAYC2" Then
            DOP4DELAYC2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP4DELAYC2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp "
                sql &= "Where Cat  = '2002' "
                sql &= "  and DKey = '" & "DELAYC2-" + HOP4.Text + "-" + DOP4DELAYC1.SelectedValue & "' "
                sql &= "Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP4DELAYC2.Items.Add(ListItem1)
                Next
            End If
        End If
        'OP5
        '----遲納原因類別1
        If pFieldName = "OP5DELAYC1" Then
            DOP5DELAYC1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP5DELAYC1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp "
                sql &= "Where Cat  = '2002' "
                sql &= "  and DKey = '" & "DELAYC1-" + HOP5.Text & "' "
                sql &= "Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP5DELAYC1.Items.Add(ListItem1)
                Next
            End If
        End If
        '----遲納原因類別2
        If pFieldName = "OP5DELAYC2" Then
            DOP5DELAYC2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP5DELAYC2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp "
                sql &= "Where Cat  = '2002' "
                sql &= "  and DKey = '" & "DELAYC2-" + HOP5.Text + "-" + DOP5DELAYC1.SelectedValue & "' "
                sql &= "Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP5DELAYC2.Items.Add(ListItem1)
                Next
            End If
        End If
        'OP6
        '----遲納原因類別1
        If pFieldName = "OP6DELAYC1" Then
            DOP6DELAYC1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP6DELAYC1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp "
                sql &= "Where Cat  = '2002' "
                sql &= "  and DKey = '" & "DELAYC1-" + HOP6.Text & "' "
                sql &= "Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP6DELAYC1.Items.Add(ListItem1)
                Next
            End If
        End If
        '----遲納原因類別2
        If pFieldName = "OP6DELAYC2" Then
            DOP6DELAYC2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP6DELAYC2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp "
                sql &= "Where Cat  = '2002' "
                sql &= "  and DKey = '" & "DELAYC2-" + HOP6.Text + "-" + DOP6DELAYC1.SelectedValue & "' "
                sql &= "Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP6DELAYC2.Items.Add(ListItem1)
                Next
            End If
        End If
        'OP7
        '----遲納原因類別1
        If pFieldName = "OP7DELAYC1" Then
            DOP7DELAYC1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP7DELAYC1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp "
                sql &= "Where Cat  = '2002' "
                sql &= "  and DKey = '" & "DELAYC1-" + HOP7.Text & "' "
                sql &= "Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP7DELAYC1.Items.Add(ListItem1)
                Next
            End If
        End If
        '----遲納原因類別2
        If pFieldName = "OP7DELAYC2" Then
            DOP7DELAYC2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP7DELAYC2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp "
                sql &= "Where Cat  = '2002' "
                sql &= "  and DKey = '" & "DELAYC2-" + HOP7.Text + "-" + DOP7DELAYC1.SelectedValue & "' "
                sql &= "Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP7DELAYC2.Items.Add(ListItem1)
                Next
            End If
        End If
        'OP8
        '----遲納原因類別1
        If pFieldName = "OP8DELAYC1" Then
            DOP8DELAYC1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP8DELAYC1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp "
                sql &= "Where Cat  = '2002' "
                sql &= "  and DKey = '" & "DELAYC1-" + HOP8.Text & "' "
                sql &= "Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP8DELAYC1.Items.Add(ListItem1)
                Next
            End If
        End If
        '----遲納原因類別2
        If pFieldName = "OP8DELAYC2" Then
            DOP8DELAYC2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP8DELAYC2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp "
                sql &= "Where Cat  = '2002' "
                sql &= "  and DKey = '" & "DELAYC2-" + HOP8.Text + "-" + DOP8DELAYC1.SelectedValue & "' "
                sql &= "Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP8DELAYC2.Items.Add(ListItem1)
                Next
            End If
        End If
        '-----------------------------------------------------------------
        '-- 開發見本
        '-----------------------------------------------------------------
        '承認-作成者
        If pFieldName = "WF1" Then
            D3WF1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW1' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF1.Items.Add(ListItem1)
                Next
            End If
        End If
        '承認-責任者
        If pFieldName = "WF2" Then
            D3WF2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW2' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF2.Items.Add(ListItem1)
                Next
            End If
        End If
        '承認-製造1
        If pFieldName = "WF3NAME" Then
            D3WF3Name.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF3Name.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOWNAME' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF3Name.Items.Add(ListItem1)
                Next
            End If
        End If
        If pFieldName = "WF3" Then
            D3WF3.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF3.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF3.Items.Add(ListItem1)
                Next
            End If
        End If
        '承認-製造2
        If pFieldName = "WF4NAME" Then
            D3WF4Name.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF4Name.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOWNAME' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF4Name.Items.Add(ListItem1)
                Next
            End If
        End If
        If pFieldName = "WF4" Then
            D3WF4.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF4.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF4.Items.Add(ListItem1)
                Next
            End If
        End If
        '承認-製造3
        If pFieldName = "WF5NAME" Then
            D3WF5Name.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF5Name.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOWNAME' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF5Name.Items.Add(ListItem1)
                Next
            End If
        End If
        If pFieldName = "WF5" Then
            D3WF5.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF5.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF5.Items.Add(ListItem1)
                Next
            End If
        End If
        '承認-製造4
        If pFieldName = "WF6NAME" Then
            D3WF6Name.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF6Name.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOWNAME' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF6Name.Items.Add(ListItem1)
                Next
            End If
        End If
        If pFieldName = "WF6" Then
            D3WF6.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF6.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF6.Items.Add(ListItem1)
                Next
            End If
        End If
        '承認-廠長
        If pFieldName = "WF7NAME" Then
            D3WF7Name.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF7Name.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOWNAME' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF7Name.Items.Add(ListItem1)
                Next
            End If
        End If
        If pFieldName = "WF7" Then
            D3WF7.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF7.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF7.Items.Add(ListItem1)
                Next
            End If
        End If
        '-----------------------------------------------------------------
        '-- 原單純
        '-----------------------------------------------------------------
        '目前不控制
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetFunction)
    '**     表單功能按鈕顯示
    '**
    '*****************************************************************
    Sub ShowSheetFunction()
        Dim sql As String = ""
        Top = 1500
        '----說明設定---------------------------------------------
        If wFormSno > 0 And wStep > 3 Then    '判斷是否[簽核]
            sql = "Select * From T_WaitHandle "
            sql = sql & " Where Active = 1 "
            sql = sql & "   And FormNo =  '" & wFormNo & "'"
            sql = sql & "   And FormSno =  '" & CStr(wFormSno) & "'"
            sql = sql & "   And Step   =  '" & CStr(wStep) & "'"
            Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(sql)
            If dtWaitHandle.Rows.Count > 0 Then
                '說明設定
                DDescSheet.Visible = True
                DDecideDesc.Visible = True
                '位置設定
                DDescSheet.Style("top") = Top & "px"
                DDecideDesc.Style("top") = Top + 5 & "px"
                Top = Top + 80
                '說明欄位
                If dtWaitHandle.Rows(0)("FlowType") = 0 Then
                    DDescSheet.Visible = False
                    DDecideDesc.Visible = False
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    'ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "異常：需輸入說明")
                End If
                '遲納原因是否需輸入--控制欄位(DBEndTime)
                DBEndTime.Text = CDate(dtWaitHandle.Rows(0)("BEndTime")).ToString("yyyy/MM/dd HH:mm:ss")
            End If
        Else
            '說明欄位隱藏
            DDescSheet.Visible = False
            DDecideDesc.Visible = False
        End If
        '----按鈕/頁次設定---------------------------------------------
        sql = "Select * From M_Flow "
        sql &= " Where Active = 1 "
        sql &= "   And FormNo =  '" & wFormNo & "'"
        sql &= "   And Step   =  '" & wStep & "'"
        Dim dtFlow As DataTable = uDataBase.GetDataTable(sql)
        If dtFlow.Rows.Count > 0 Then
            '----未使用---------------------------------------------
            '電子簽章
            If dtFlow.Rows(0)("SignImage") = 1 Then
            Else
            End If
            '附加檔(由FormField中設定)
            If dtFlow.Rows(0)("Attach") = 1 Then
            Else
            End If
            '----按鈕設定---------------------------------------------
            '儲存按鈕
            If dtFlow.Rows(0)("SaveFun") = 1 Then
                BSAVE.Visible = True
                BSAVE.Text = dtFlow.Rows(0)("SaveDesc")
                'BSAVE.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BSAVE.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                BSAVE.Style("top") = Top & "px"
            Else
                BSAVE.Visible = False
            End If
            'NG-1按鈕
            If dtFlow.Rows(0)("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Text = dtFlow.Rows(0)("NGDesc1")
                'BNG1.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BNG1.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wNGSts1 = 2                             '狀態(0:開發中,1:OK,2:NG,3:取消)
                BNG1.Style("top") = Top & "px"
            Else
                BNG1.Visible = False
            End If
            'NG-2按鈕
            If dtFlow.Rows(0)("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Text = dtFlow.Rows(0)("NGDesc2")
                'BNG2.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BNG2.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wNGSts2 = 3                             '狀態(0:開發中,1:OK,2:NG,3:取消)
                BNG2.Style("top") = Top & "px"
            Else
                BNG2.Visible = False
            End If
            'OK按鈕
            If dtFlow.Rows(0)("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Text = dtFlow.Rows(0)("OKDesc")
                'BOK.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BOK.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wOKSts = 1                              '狀態(0:開發中,1:OK,2:NG,3:取消)
                BOK.Style("top") = Top & "px"
            Else
                BOK.Visible = False
            End If
            '頁次設定
            If Not IsPostBack Then
                '控制欄位(DPageIdx)
                DPageIdx.Text = dtFlow.Rows(0)("PageIdx").ToString
                Select Case CInt(DPageIdx.Text)
                    Case 0
                        DSIGNMARK.Style("left") = 122 & "px"
                    Case 1
                        DSIGNMARK.Style("left") = 249 & "px"
                    Case 2
                        DSIGNMARK.Style("left") = 377 & "px"
                    Case 3
                        DSIGNMARK.Style("left") = 504 & "px"
                    Case Else
                        DSIGNMARK.Style("left") = -100 & "px"
                End Select

                SetImageButtonImageFile(dtFlow.Rows(0)("PageIdx"))
                DModify.Style("left") = -100 & "px"         '需樣品/需登錄-修改可否
            End If
            '自動設定(適用 = 廠長)  Modify-2011/10/5
            If UCase(Request.QueryString("pUserID")) = "FADM02" Then
                '說明欄位內容
                If DDecideDesc.Visible = True Then
                    If DDecideDesc.Text = "" Then DDecideDesc.Text = "OK"
                End If
                '閱讀履歷
                DReadHistory.Checked = True
                '開發履歷
                GridView2.Style("top") = "1650px"
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("COMMISSIONFilePath")


        Dim SQL As String
        SQL = "Select * From F_CommissionSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtCommissionSheet As DataTable = uDataBase.GetDataTable(SQL)
        If dtCommissionSheet.Rows.Count > 0 Then
            '-----------------------------------------------------------------
            '-- 開發委託單
            '-----------------------------------------------------------------
            '----基本欄位設定-------------------------------------------------    
            DNO.Text = dtCommissionSheet.Rows(0).Item("NO")                             'NO
            If dtCommissionSheet.Rows(0).Item("REFNO") <> "" And _
               (dtCommissionSheet.Rows(0).Item("NeedSample") = 1 Or dtCommissionSheet.Rows(0).Item("NeedItemRegister") = 1) Then       '參考NO
                DREFNO.Text = dtCommissionSheet.Rows(0).Item("REFNO")
                LREFNO.Text = dtCommissionSheet.Rows(0).Item("REFNO")
                LREFNO.NavigateUrl = "CommissionSheet_03.aspx?pNo=" + dtCommissionSheet.Rows(0).Item("REFNO")
                ' uCommon.GetAppSetting("Http") & 
                LREFNO.Visible = True
            End If
            DAPPDATE.Text = dtCommissionSheet.Rows(0).Item("APPDATE")                   '申請日
            DAPPDEPT.Text = dtCommissionSheet.Rows(0).Item("APPDEPT")                   '部門
            DAPPPER.Text = dtCommissionSheet.Rows(0).Item("APPPER")                     '職稱
            SetFieldData(FindFieldInf("1-BASE"), "APPBUYER", dtCommissionSheet.Rows(0).Item("APPBUYER"))    'BUYER
            DSellVendor.Text = dtCommissionSheet.Rows(0).Item("SELLVENDOR")             '委託廠商
            DESYQTY.Text = dtCommissionSheet.Rows(0).Item("ESYQTY")                     '預估量
            DEXPDEL.Text = dtCommissionSheet.Rows(0).Item("EXPDEL")                     '希望交期
            DCUSTITEM.Text = dtCommissionSheet.Rows(0).Item("CUSTITEM")                 '客戶ITEM
            SetFieldData(FindFieldInf("1-BASE"), "USAGE", dtCommissionSheet.Rows(0).Item("USAGE"))          '用途
            DORNO.Text = dtCommissionSheet.Rows(0).Item("ORNO")                         'OR-NO
            SetFieldData(FindFieldInf("1-BASE"), "NEEDMAP", dtCommissionSheet.Rows(0).Item("NEEDMAP"))      '需圖面
            '草圖
            If dtCommissionSheet.Rows(0).Item("MAPREFFILE") <> "" Then
                LMAPREFFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("MAPREFFILE")
                LMAPREFFILE.Visible = True
            End If
            '----樣品欄位設定-------------------------------------------------    
            If dtCommissionSheet.Rows(0).Item("NeedSample") = 1 Then                                        '需樣品
                DNeedSample.Checked = True
            Else
                DNeedSample.Checked = False
                If FindFieldInf("1-SAMPLE") = 0 Then DNeedSample.Style("top") = -100 & "px"
            End If
            If dtCommissionSheet.Rows(0).Item("NeedItemRegister") = 1 Then                                  '需登錄
                DNeedItemRegister.Checked = True
            Else
                DNeedItemRegister.Checked = False
                If FindFieldInf("1-SAMPLE") = 0 Then DNeedItemRegister.Style("top") = -100 & "px"
            End If
            SetFieldData(FindFieldInf("1-SAMPLE"), "PRO", dtCommissionSheet.Rows(0).Item("PRO"))            '製品區分
            DOPPART.Text = dtCommissionSheet.Rows(0).Item("OPPART")                     '開具(色)
            DPLEN.Text = dtCommissionSheet.Rows(0).Item("PLEN")                         '長度(企)
            SetFieldData(FindFieldInf("1-SAMPLE"), "PLENUN", dtCommissionSheet.Rows(0).Item("PLENUN"))      '長度單位(企)
            DPQTY.Text = dtCommissionSheet.Rows(0).Item("PQTY")                         '數量(企)
            SetFieldData(FindFieldInf("1-SAMPLE"), "PQTYUN", dtCommissionSheet.Rows(0).Item("PQTYUN"))      '數量單位(企)
            DEALEN.Text = dtCommissionSheet.Rows(0).Item("EALEN")                        '長度(EA)
            SetFieldData(FindFieldInf("1-SAMPLE-EA"), "EALENUN", dtCommissionSheet.Rows(0).Item("EALENUN")) '長度單位(EA)
            DEAQTY.Text = dtCommissionSheet.Rows(0).Item("EAQTY")                        '數量(EA)
            SetFieldData(FindFieldInf("1-SAMPLE-EA"), "EAQTYUN", dtCommissionSheet.Rows(0).Item("EAQTYUN")) '數量單位(EA)
            DUPSLI.Text = dtCommissionSheet.Rows(0).Item("UPSLI")                       '拉頭(上)
            DLOSLI.Text = dtCommissionSheet.Rows(0).Item("LOSLI")                       '拉頭(下)
            DUPFIN.Text = dtCommissionSheet.Rows(0).Item("UPFIN")                       '表面處理(上)
            DLOFIN.Text = dtCommissionSheet.Rows(0).Item("LOFIN")                       '表面處理(下)
            DUPSTK.Text = dtCommissionSheet.Rows(0).Item("UPSTK")                       '上止種類
            DLOSTK.Text = dtCommissionSheet.Rows(0).Item("LOSTK")                       '下止種類
            DSPSPEC.Text = dtCommissionSheet.Rows(0).Item("SPSPEC")                     '特殊規格
            '----開發規格欄位設定-------------------------------------------------    
            SetFieldData(FindFieldInf("1-DEVELOP"), "SIZENO", dtCommissionSheet.Rows(0).Item("SIZENO"))     '型別
            SetFieldData(FindFieldInf("1-DEVELOP"), "ITEM", dtCommissionSheet.Rows(0).Item("ITEM"))         '鏈條型式
            SetFieldData(FindFieldInf("1-DEVELOP"), "TATYPE", dtCommissionSheet.Rows(0).Item("TATYPE"))     '布帶
            DTAWIDTH.Text = dtCommissionSheet.Rows(0).Item("TAWIDTH")                   '布帶寬度
            SetFieldData(FindFieldInf("1-DEVELOP"), "ECOLSEL", dtCommissionSheet.Rows(0).Item("ECOLSEL"))    '鏈齒顏色-SEL
            DECOL.Text = dtCommissionSheet.Rows(0).Item("ECOL")                         '鏈齒顏色
            SetFieldData(FindFieldInf("1-DEVELOP"), "CCOLSEL", dtCommissionSheet.Rows(0).Item("CCOLSEL"))   '丸紐-SEL
            DCCOL.Text = dtCommissionSheet.Rows(0).Item("CCOL")                         '丸紐
            SetFieldData(FindFieldInf("1-DEVELOP"), "TACOL", dtCommissionSheet.Rows(0).Item("TACOL"))       '布帶-色番(同)
            DTACOLNO.Text = dtCommissionSheet.Rows(0).Item("TACOLNO")                   '布帶-色番(同)
            DTAYCOLNO.Text = dtCommissionSheet.Rows(0).Item("TAYCOLNO")                 '布帶-YKK(同)
            SetFieldData(FindFieldInf("1-DEVELOP"), "TALCOL", dtCommissionSheet.Rows(0).Item("TALCOL"))     '布帶-色番(左)
            DTALCOLNO.Text = dtCommissionSheet.Rows(0).Item("TALCOLNO")                 '布帶-色番(左)
            DTALYCOLNO.Text = dtCommissionSheet.Rows(0).Item("TALYCOLNO")               '布帶-YKK(左)
            SetFieldData(FindFieldInf("1-DEVELOP"), "TARCOL", dtCommissionSheet.Rows(0).Item("TARCOL"))     '布帶-色番(右)
            DTARCOLNO.Text = dtCommissionSheet.Rows(0).Item("TARCOLNO")                 '布帶-色番(右)
            DTARYCOLNO.Text = dtCommissionSheet.Rows(0).Item("TARYCOLNO")               '布帶-YKK(右)
            SetFieldData(FindFieldInf("1-DEVELOP"), "THUPCOL", dtCommissionSheet.Rows(0).Item("THUPCOL"))   '縫上-色番(同)
            DTHUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THUPCOLNO")               '縫上-色番(同)
            DTHUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THUPYCOLNO")             '縫上-YKK(同)
            SetFieldData(FindFieldInf("1-DEVELOP"), "THLUPCOL", dtCommissionSheet.Rows(0).Item("THLUPCOL")) '縫上-色番(左左)
            DTHLUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLUPCOLNO")             '縫上-色番(左左)
            DTHLUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLUPYCOLNO")           '縫上-YKK(左左)
            SetFieldData(FindFieldInf("1-DEVELOP"), "THLRUPCOL", dtCommissionSheet.Rows(0).Item("THLRUPCOL")) '縫上-色番(左右)
            DTHLRUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLRUPCOLNO")             '縫上-色番(左右)
            DTHLRUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLRUPYCOLNO")           '縫上-YKK(左右)
            SetFieldData(FindFieldInf("1-DEVELOP"), "THRUPCOL", dtCommissionSheet.Rows(0).Item("THRUPCOL")) '縫上-色番(右左)
            DTHRUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRUPCOLNO")             '縫上-色番(右左)
            DTHRUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRUPYCOLNO")           '縫上-YKK(右左)
            SetFieldData(FindFieldInf("1-DEVELOP"), "THRRUPCOL", dtCommissionSheet.Rows(0).Item("THRRUPCOL")) '縫上-色番(右右)
            DTHRRUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRRUPCOLNO")             '縫上-色番(右右)
            DTHRRUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRRUPYCOLNO")           '縫上-YKK(右右)
            SetFieldData(FindFieldInf("1-DEVELOP"), "THLOCOL", dtCommissionSheet.Rows(0).Item("THLOCOL"))   '縫下-色番(同)
            DTHLOCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLOCOLNO")               '縫下-色番(同)
            DTHLOYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLOYCOLNO")             '縫下-YKK(同)
            SetFieldData(FindFieldInf("1-DEVELOP"), "THLLOCOL", dtCommissionSheet.Rows(0).Item("THLLOCOL")) '縫下-色番(左左)
            DTHLLOCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLLOCOLNO")             '縫下-色番(左左)
            DTHLLOYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLLOYCOLNO")           '縫下-YKK(左左)
            SetFieldData(FindFieldInf("1-DEVELOP"), "THLRLOCOL", dtCommissionSheet.Rows(0).Item("THLRLOCOL")) '縫下-色番(左右)
            DTHLRLOCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLRLOCOLNO")             '縫下-色番(左右)
            DTHLRLOYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLRLOYCOLNO")           '縫下-YKK(左右)

            SetFieldData(FindFieldInf("1-DEVELOP"), "THRLOCOL", dtCommissionSheet.Rows(0).Item("THRLOCOL")) '縫下-色番(右左)
            DTHRLOCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRLOCOLNO")             '縫下-色番(右左)
            DTHRLOYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRLOYCOLNO")           '縫下-YKK(右左)

            SetFieldData(FindFieldInf("1-DEVELOP"), "THRRLOCOL", dtCommissionSheet.Rows(0).Item("THRRLOCOL")) '縫下-色番(右右)
            DTHRRLOCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRRLOCOLNO")             '縫下-色番(右右)
            DTHRRLOYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRRLOYCOLNO")           '縫下-YKK(右右)


            DXMLEN.Text = dtCommissionSheet.Rows(0).Item("XMLEN")                       'X-尺寸
            SetFieldData(FindFieldInf("1-DEVELOP"), "XMCOL", dtCommissionSheet.Rows(0).Item("XMCOL"))       'X-色番
            DXMCOLNO.Text = dtCommissionSheet.Rows(0).Item("XMCOLNO")                   'X-色番
            DXMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("XMYCOLNO")                 'X-YKK
            DAMLEN.Text = dtCommissionSheet.Rows(0).Item("AMLEN")                       'A-尺寸
            SetFieldData(FindFieldInf("1-DEVELOP"), "AMCOL", dtCommissionSheet.Rows(0).Item("AMCOL"))       'A-色番
            DAMCOLNO.Text = dtCommissionSheet.Rows(0).Item("AMCOLNO")                   'A-色番
            DAMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("AMYCOLNO")                 'A-YKK
            DBMLEN.Text = dtCommissionSheet.Rows(0).Item("BMLEN")                       'B-尺寸
            SetFieldData(FindFieldInf("1-DEVELOP"), "BMCOL", dtCommissionSheet.Rows(0).Item("BMCOL"))       'B-色番
            DBMCOLNO.Text = dtCommissionSheet.Rows(0).Item("BMCOLNO")                   'B-色番
            DBMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("BMYCOLNO")                 'B-YKK
            DCMLEN.Text = dtCommissionSheet.Rows(0).Item("CMLEN")                       'C-尺寸
            SetFieldData(FindFieldInf("1-DEVELOP"), "CMCOL", dtCommissionSheet.Rows(0).Item("CMCOL"))       'C-色番
            DCMCOLNO.Text = dtCommissionSheet.Rows(0).Item("CMCOLNO")                   'C-色番
            DCMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("CMYCOLNO")                 'C-YKK
            DDMLEN.Text = dtCommissionSheet.Rows(0).Item("DMLEN")                       'D-尺寸
            SetFieldData(FindFieldInf("1-DEVELOP"), "DMCOL", dtCommissionSheet.Rows(0).Item("DMCOL"))       'D-色番
            DDMCOLNO.Text = dtCommissionSheet.Rows(0).Item("DMCOLNO")                   'D-色番
            DDMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("DMYCOLNO")                 'D-YKK
            DEMLEN.Text = dtCommissionSheet.Rows(0).Item("EMLEN")                       'E-尺寸
            SetFieldData(FindFieldInf("1-DEVELOP"), "EMCOL", dtCommissionSheet.Rows(0).Item("EMCOL"))       'E-色番
            DEMCOLNO.Text = dtCommissionSheet.Rows(0).Item("EMCOLNO")                   'E-色番
            DEMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("EMYCOLNO")                 'E-YKK
            DFMLEN.Text = dtCommissionSheet.Rows(0).Item("FMLEN")                       'F-尺寸
            SetFieldData(FindFieldInf("1-DEVELOP"), "FMCOL", dtCommissionSheet.Rows(0).Item("FMCOL"))       'F-色番
            DFMCOLNO.Text = dtCommissionSheet.Rows(0).Item("FMCOLNO")                   'F-色番
            DFMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("FMYCOLNO")                 'F-YKK
            DGMLEN.Text = dtCommissionSheet.Rows(0).Item("GMLEN")                       'G-尺寸
            SetFieldData(FindFieldInf("1-DEVELOP"), "GMCOL", dtCommissionSheet.Rows(0).Item("GMCOL"))       'G-色番
            DGMCOLNO.Text = dtCommissionSheet.Rows(0).Item("GMCOLNO")                   'G-色番
            DGMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("GMYCOLNO")                 'G-YKK
            DHMLEN.Text = dtCommissionSheet.Rows(0).Item("HMLEN")                       'H-尺寸
            SetFieldData(FindFieldInf("1-DEVELOP"), "HMCOL", dtCommissionSheet.Rows(0).Item("HMCOL"))       'H-色番
            DHMCOLNO.Text = dtCommissionSheet.Rows(0).Item("HMCOLNO")                   'H-色番
            DHMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("HMYCOLNO")                 'H-YKK
            DLYLEN.Text = dtCommissionSheet.Rows(0).Item("LYLEN")                       '緯紗-尺寸
            SetFieldData(FindFieldInf("1-DEVELOP"), "LYCOL", dtCommissionSheet.Rows(0).Item("LYCOL"))       '緯紗-色番
            DLYCOLNO.Text = dtCommissionSheet.Rows(0).Item("LYCOLNO")                   '緯紗-色番
            DLYYCOLNO.Text = dtCommissionSheet.Rows(0).Item("LYYCOLNO")                 '緯紗-YKK
            DOTCON.Text = dtCommissionSheet.Rows(0).Item("OTCON")                       '其他條件
            '----圖面欄位設定-------------------------------------------------    
            DMAPNO.Text = dtCommissionSheet.Rows(0).Item("MAPNO")                       '圖號

            SetFieldData(FindFieldInf("MAKEMAP"), "MAKEMAP", dtCommissionSheet.Rows(0).Item("MAKEMAP"))       '製圖者           
            SetFieldData(FindFieldInf("1-MAP"), "LEVEL", dtCommissionSheet.Rows(0).Item("LEVEL"))           '難易度
            '圖檔
            If dtCommissionSheet.Rows(0).Item("MAPFILE") <> "" Then
                LMAPFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("MAPFILE")
                LMAPFILE.Visible = True
            End If
            '----其他附件-------------------------------------------------    
            '適用型別檔
            If dtCommissionSheet.Rows(0).Item("FORTYPEFILE") <> "" Then
                LFORTYPEFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("FORTYPEFILE")
                LFORTYPEFILE.Visible = True
            End If
            '品質檢測項目檔
            If dtCommissionSheet.Rows(0).Item("QCCHECKFILE") <> "" Then
                LQCCHECKFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCCHECKFILE")
                LQCCHECKFILE.Visible = True
            End If
            'QC檢測文件
            If dtCommissionSheet.Rows(0).Item("QCFILE1") <> "" Then
                LQCFILE1.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE1")
                LQCFILE1.Visible = True
            End If
            If dtCommissionSheet.Rows(0).Item("QCFILE2") <> "" Then
                LQCFILE2.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE2")
                LQCFILE2.Visible = True
            End If
            If dtCommissionSheet.Rows(0).Item("QCFILE3") <> "" Then
                LQCFILE3.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE3")
                LQCFILE3.Visible = True
            End If
            If dtCommissionSheet.Rows(0).Item("QCFILE4") <> "" Then
                LQCFILE4.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE4")
                LQCFILE4.Visible = True
            End If
            If dtCommissionSheet.Rows(0).Item("QCFILE5") <> "" Then
                LQCFILE5.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE5")
                LQCFILE5.Visible = True
            End If
            If dtCommissionSheet.Rows(0).Item("QCFILE6") <> "" Then
                LQCFILE6.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE6")
                LQCFILE6.Visible = True
            End If


            If dtCommissionSheet.Rows(0).Item("GenFILE1") <> "" Then
                LGenFILE1.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("GenFILE1")
                LGenFILE1.Visible = True
            End If

            '客戶切結書
            If dtCommissionSheet.Rows(0).Item("CONTACTFILE") <> "" Then
                LCONTACTFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("CONTACTFILE")
                LCONTACTFILE.Visible = True
            End If
            '樣品確認書
            If dtCommissionSheet.Rows(0).Item("SAMPLECONFIRMFILE") <> "" Then
                LSAMPLECONFIRMFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("SAMPLECONFIRMFILE")
                LSAMPLECONFIRMFILE.Visible = True
            End If
            '製造授權書
            If dtCommissionSheet.Rows(0).Item("MANUFAUTHORITYFILE") <> "" Then
                LMANUFAUTHORITYFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("MANUFAUTHORITYFILE")
                LMANUFAUTHORITYFILE.Visible = True
            End If

            '外注加工費
            DMANUOUTPRICE.Text = dtCommissionSheet.Rows(0).Item("MANUOUTPRICE")

            '報價單       
            If dtCommissionSheet.Rows(0).Item("FORCASTFILE") <> "" Then
                LFORCASTFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("FORCASTFILE")
                LFORCASTFILE.Visible = True
            End If

            '-----------------------------------------------------------------
            '-- 製造委託
            '-----------------------------------------------------------------
            SQL = "Select * From FS_ManufactureSheet "
            SQL &= " Where NO =  '" & dtCommissionSheet.Rows(0).Item("NO") & "'"
            Dim dtManufactureSheet As DataTable = uDataBase.GetDataTable(SQL)
            If dtManufactureSheet.Rows.Count > 0 Then
                '----基本欄位設定-------------------------------------------------    
                DDEVTITLE.Text = dtManufactureSheet.Rows(0).Item("DEVTITLE")            '開發主題
                DDEVNO.Text = dtManufactureSheet.Rows(0).Item("DEVNO")                  '開發NO.
                DCODENO.Text = dtManufactureSheet.Rows(0).Item("CODENO")                'CODE NO.
                DISSDATE.Text = dtManufactureSheet.Rows(0).Item("ISSDATE")              '發行日
                DDEVPER1.Text = dtManufactureSheet.Rows(0).Item("DEVPER1")              '開發擔當

                '----開發內容-------------------------------------------------    
                '示意圖檔
                If dtManufactureSheet.Rows(0).Item("HINTFILE") <> "" Then
                    LHINTFILE.ImageUrl = Path & dtManufactureSheet.Rows(0).Item("HINTFILE")
                    LHINTFILE.Visible = True
                End If
                DUPSTK1.Text = dtManufactureSheet.Rows(0).Item("UPSTK")                  '上止
                DLOSTK1.Text = dtManufactureSheet.Rows(0).Item("LOSTK")                  '下止
                DOPPART1.Text = dtManufactureSheet.Rows(0).Item("OPPART")                '開具(色)
                DTASPEC.Text = dtManufactureSheet.Rows(0).Item("TASPEC")                '布帶
                DECOL1.Text = dtManufactureSheet.Rows(0).Item("ECOL")                    '鏈齒顏色
                DCCOL1.Text = dtManufactureSheet.Rows(0).Item("CCOL")                    '丸紐
                DTHSPEC.Text = dtManufactureSheet.Rows(0).Item("THSPEC")                '縫工線
                DPLEN1.Text = dtManufactureSheet.Rows(0).Item("PLEN")                    '長度(企)
                DPQTY1.Text = dtManufactureSheet.Rows(0).Item("PQTY")                    '數量(企)
                DEALEN1.Text = dtManufactureSheet.Rows(0).Item("EALEN")                  '長度(EA)
                DEAQTY1.Text = dtManufactureSheet.Rows(0).Item("EAQTY")                  '數量(EA)
                '----工程-------------------------------------------------    
                DMANUFTYPE.Text = dtManufactureSheet.Rows(0).Item("MANUFTYPE")          '內外製
                HOP1.Text = dtManufactureSheet.Rows(0).Item("OP1")                      'OP1-工程

                HOP1.NavigateUrl = "ManufactureHistory.aspx?&pFormNo=" & wFormNo & "&pFormSno=" & CStr(wFormSno) & "&pOPCode=" & fpObj.GetOPCode("CODE-" & HOP1.Text)

                'OP1-工程改超連結
                DOP1PER.Text = dtManufactureSheet.Rows(0).Item("OP1PER")                'OP1-擔當
                DOP1BTIME.Text = dtManufactureSheet.Rows(0).Item("OP1BTIME")            'OP1-預定納期
                DOP1BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP1BHOURS")          'OP1-預定時數
                DOP1ATIME.Text = dtManufactureSheet.Rows(0).Item("OP1ATIME")            'OP1-實際納期
                DOP1AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP1AHOURS")          'OP1-實際時數
                DOP1CON.Text = dtManufactureSheet.Rows(0).Item("OP1CON")                'OP1-作業內容
                SetFieldData(FindFieldInf("2-DELAYCAT1"), "OP1DELAYC1", dtManufactureSheet.Rows(0).Item("OP1DELAYC1"))       'OP1-遲納原因-1
                SetFieldData(FindFieldInf("2-DELAYCAT1"), "OP1DELAYC2", dtManufactureSheet.Rows(0).Item("OP1DELAYC2"))       'OP1-遲納原因-2
                DOP1REM.Text = dtManufactureSheet.Rows(0).Item("OP1REM")                'OP1-遲納原因
                If DOP1PER.Text = "" Then DOP40.Visible = False
                HOP2.Text = dtManufactureSheet.Rows(0).Item("OP2")                      'OP2-工程

                HOP2.NavigateUrl = "ManufactureHistory.aspx?&pFormNo=" & wFormNo & "&pFormSno=" & CStr(wFormSno) & "&pOPCode=" & fpObj.GetOPCode("CODE-" & HOP2.Text)

                'OP2-工程超連結
                DOP2PER.Text = dtManufactureSheet.Rows(0).Item("OP2PER")                'OP2-擔當
                DOP2BTIME.Text = dtManufactureSheet.Rows(0).Item("OP2BTIME")            'OP2-預定納期
                DOP2BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP2BHOURS")          'OP2-預定時數
                DOP2ATIME.Text = dtManufactureSheet.Rows(0).Item("OP2ATIME")            'OP2-實際納期
                DOP2AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP2AHOURS")          'OP2-實際時數
                DOP2CON.Text = dtManufactureSheet.Rows(0).Item("OP2CON")                'OP2-作業內容
                SetFieldData(FindFieldInf("2-DELAYCAT2"), "OP2DELAYC1", dtManufactureSheet.Rows(0).Item("OP2DELAYC1"))       'OP2-遲納原因-1
                SetFieldData(FindFieldInf("2-DELAYCAT2"), "OP2DELAYC2", dtManufactureSheet.Rows(0).Item("OP2DELAYC2"))       'OP2-遲納原因-2
                DOP2REM.Text = dtManufactureSheet.Rows(0).Item("OP2REM")                'OP2-遲納原因
                If DOP2PER.Text = "" Then DOP50.Visible = False

                HOP3.Text = dtManufactureSheet.Rows(0).Item("OP3")                      'OP3-工程

                HOP3.NavigateUrl = "ManufactureHistory.aspx?&pFormNo=" & wFormNo & "&pFormSno=" & CStr(wFormSno) & "&pOPCode=" & fpObj.GetOPCode("CODE-" & HOP3.Text)


                'OP3-工程超連結
                DOP3PER.Text = dtManufactureSheet.Rows(0).Item("OP3PER")                'OP3-擔當
                DOP3BTIME.Text = dtManufactureSheet.Rows(0).Item("OP3BTIME")            'OP3-預定納期
                DOP3BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP3BHOURS")          'OP3-預定時數
                DOP3ATIME.Text = dtManufactureSheet.Rows(0).Item("OP3ATIME")            'OP3-實際納期
                DOP3AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP3AHOURS")          'OP3-實際時數
                DOP3CON.Text = dtManufactureSheet.Rows(0).Item("OP3CON")                'OP3-作業內容
                SetFieldData(FindFieldInf("2-DELAYCAT3"), "OP3DELAYC1", dtManufactureSheet.Rows(0).Item("OP3DELAYC1"))       'OP3-遲納原因-1
                SetFieldData(FindFieldInf("2-DELAYCAT3"), "OP3DELAYC2", dtManufactureSheet.Rows(0).Item("OP3DELAYC2"))       'OP3-遲納原因-2
                DOP3REM.Text = dtManufactureSheet.Rows(0).Item("OP3REM")                'OP3-遲納原因
                If DOP3PER.Text = "" Then DOP60.Visible = False

                HOP4.Text = dtManufactureSheet.Rows(0).Item("OP4")                      'OP4-工程

                HOP4.NavigateUrl = "ManufactureHistory.aspx?&pFormNo=" & wFormNo & "&pFormSno=" & CStr(wFormSno) & "&pOPCode=" & fpObj.GetOPCode("CODE-" & HOP4.Text)

                'OP4-工程超連結
                DOP4PER.Text = dtManufactureSheet.Rows(0).Item("OP4PER")                'OP4-擔當
                DOP4BTIME.Text = dtManufactureSheet.Rows(0).Item("OP4BTIME")            'OP4-預定納期
                DOP4BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP4BHOURS")          'OP4-預定時數
                DOP4ATIME.Text = dtManufactureSheet.Rows(0).Item("OP4ATIME")            'OP4-實際納期
                DOP4AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP4AHOURS")          'OP4-實際時數
                DOP4CON.Text = dtManufactureSheet.Rows(0).Item("OP4CON")                'OP4-作業內容
                SetFieldData(FindFieldInf("2-DELAYCAT4"), "OP4DELAYC1", dtManufactureSheet.Rows(0).Item("OP4DELAYC1"))       'OP4-遲納原因-1
                SetFieldData(FindFieldInf("2-DELAYCAT4"), "OP4DELAYC2", dtManufactureSheet.Rows(0).Item("OP4DELAYC2"))       'OP4-遲納原因-2
                DOP4REM.Text = dtManufactureSheet.Rows(0).Item("OP4REM")                'OP4-遲納原因
                If DOP4PER.Text = "" Then DOP70.Visible = False

                HOP5.Text = dtManufactureSheet.Rows(0).Item("OP5")                      'OP5-工程

                HOP5.NavigateUrl = "ManufactureHistory.aspx?&pFormNo=" & wFormNo & "&pFormSno=" & CStr(wFormSno) & "&pOPCode=" & fpObj.GetOPCode("CODE-" & HOP5.Text)
                'OP5-工程超連結



                DOP5PER.Text = dtManufactureSheet.Rows(0).Item("OP5PER")                'OP5-擔當
                DOP5BTIME.Text = dtManufactureSheet.Rows(0).Item("OP5BTIME")            'OP5-預定納期
                DOP5BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP5BHOURS")          'OP5-預定時數
                DOP5ATIME.Text = dtManufactureSheet.Rows(0).Item("OP5ATIME")            'OP5-實際納期
                DOP5AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP5AHOURS")          'OP5-實際時數
                DOP5CON.Text = dtManufactureSheet.Rows(0).Item("OP5CON")                'OP5-作業內容
                SetFieldData(FindFieldInf("2-DELAYCAT5"), "OP5DELAYC1", dtManufactureSheet.Rows(0).Item("OP5DELAYC1"))       'OP5-遲納原因-1
                SetFieldData(FindFieldInf("2-DELAYCAT5"), "OP5DELAYC2", dtManufactureSheet.Rows(0).Item("OP5DELAYC2"))       'OP5-遲納原因-2
                DOP5REM.Text = dtManufactureSheet.Rows(0).Item("OP5REM")                'OP5-遲納原因
                If DOP5PER.Text = "" Then DOP80.Visible = False

                HOP6.Text = dtManufactureSheet.Rows(0).Item("OP6")                      'OP6-工程

                HOP6.NavigateUrl = "ManufactureHistory.aspx?&pFormNo=" & wFormNo & "&pFormSno=" & CStr(wFormSno) & "&pOPCode=" & fpObj.GetOPCode("CODE-" & HOP6.Text)


                'OP6-工程超連結
                DOP6PER.Text = dtManufactureSheet.Rows(0).Item("OP6PER")                'OP6-擔當
                DOP6BTIME.Text = dtManufactureSheet.Rows(0).Item("OP6BTIME")            'OP6-預定納期
                DOP6BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP6BHOURS")          'OP6-預定時數
                DOP6ATIME.Text = dtManufactureSheet.Rows(0).Item("OP6ATIME")            'OP6-實際納期
                DOP6AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP6AHOURS")          'OP6-實際時數
                DOP6CON.Text = dtManufactureSheet.Rows(0).Item("OP6CON")                'OP6-作業內容
                SetFieldData(FindFieldInf("2-DELAYCAT6"), "OP6DELAYC1", dtManufactureSheet.Rows(0).Item("OP6DELAYC1"))       'OP6-遲納原因-1
                SetFieldData(FindFieldInf("2-DELAYCAT6"), "OP6DELAYC2", dtManufactureSheet.Rows(0).Item("OP6DELAYC2"))       'OP6-遲納原因-2
                DOP6REM.Text = dtManufactureSheet.Rows(0).Item("OP6REM")                'OP6-遲納原因
                If DOP6PER.Text = "" Then DOP90.Visible = False

                HOP7.Text = dtManufactureSheet.Rows(0).Item("OP7")                      'OP7-工程

                HOP7.NavigateUrl = "ManufactureHistory.aspx?&pFormNo=" & wFormNo & "&pFormSno=" & CStr(wFormSno) & "&pOPCode=" & fpObj.GetOPCode("CODE-" & HOP7.Text)
                'OP7-工程超連結


                DOP7PER.Text = dtManufactureSheet.Rows(0).Item("OP7PER")                'OP7-擔當
                DOP7BTIME.Text = dtManufactureSheet.Rows(0).Item("OP7BTIME")            'OP7-預定納期
                DOP7BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP7BHOURS")          'OP7-預定時數
                DOP7ATIME.Text = dtManufactureSheet.Rows(0).Item("OP7ATIME")            'OP7-實際納期
                DOP7AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP7AHOURS")          'OP7-實際時數
                DOP7CON.Text = dtManufactureSheet.Rows(0).Item("OP7CON")                'OP7-作業內容
                SetFieldData(FindFieldInf("2-DELAYCAT7"), "OP7DELAYC1", dtManufactureSheet.Rows(0).Item("OP7DELAYC1"))       'OP7-遲納原因-1
                SetFieldData(FindFieldInf("2-DELAYCAT7"), "OP7DELAYC2", dtManufactureSheet.Rows(0).Item("OP7DELAYC2"))       'OP7-遲納原因-2
                DOP7REM.Text = dtManufactureSheet.Rows(0).Item("OP7REM")                'OP7-遲納原因
                If DOP7PER.Text = "" Then DOP100.Visible = False

                HOP8.NavigateUrl = "ManufactureHistory.aspx?&pFormNo=" & wFormNo & "&pFormSno=" & CStr(wFormSno) & "&pOPCode=" & fpObj.GetOPCode("CODE-" & HOP8.Text)
                'OP8-工程超連結

                HOP8.Text = dtManufactureSheet.Rows(0).Item("OP8")                      'OP8-工程


                DOP8PER.Text = dtManufactureSheet.Rows(0).Item("OP8PER")                'OP8-擔當
                DOP8BTIME.Text = dtManufactureSheet.Rows(0).Item("OP8BTIME")            'OP8-預定納期
                DOP8BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP8BHOURS")          'OP8-預定時數
                DOP8ATIME.Text = dtManufactureSheet.Rows(0).Item("OP8ATIME")            'OP8-實際納期
                DOP8AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP8AHOURS")          'OP8-實際時數
                DOP8CON.Text = dtManufactureSheet.Rows(0).Item("OP8CON")                'OP8-作業內容
                SetFieldData(FindFieldInf("2-DELAYCAT8"), "OP8DELAYC1", dtManufactureSheet.Rows(0).Item("OP8DELAYC1"))       'OP8-遲納原因-1
                SetFieldData(FindFieldInf("2-DELAYCAT8"), "OP8DELAYC2", dtManufactureSheet.Rows(0).Item("OP8DELAYC2"))       'OP8-遲納原因-2
                DOP8REM.Text = dtManufactureSheet.Rows(0).Item("OP8REM")                'OP8-遲納原因
                If DOP8PER.Text = "" Then DOP110.Visible = False

                '最後試作工程表示(CheckBox)-試作NG時使用
                SQL = "Select Top 1 Step, DecideID From T_WaitHandle "
                SQL &= "Where FormNo  =  '" & wFormNo & "' "
                SQL &= "  And FormSno =  '" & CStr(wFormSno) & "' "
                SQL &= "  And Step    <> '" & CStr(wStep) & "' "
                SQL &= "  And Active  = '0' "
                SQL &= "  And FlowType <> '0' "
                SQL &= "Order by Unique_ID Desc "
                Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(SQL)
                If dtWaitHandle.Rows.Count > 0 Then
                    Select Case dtWaitHandle.Rows(0).Item("Step")
                        Case 40
                            DOP40.Checked = True
                        Case 50
                            DOP50.Checked = True
                        Case 60
                            DOP60.Checked = True
                        Case 70
                            DOP70.Checked = True
                        Case 80
                            DOP80.Checked = True
                        Case 90
                            DOP90.Checked = True
                        Case 100
                            DOP100.Checked = True
                        Case 110
                            DOP110.Checked = True
                        Case Else
                            DOP40.Checked = True
                    End Select
                End If
            End If
            '-----------------------------------------------------------------
            '-- 開發見本
            '-----------------------------------------------------------------
            SQL = "Select * From FS_SampleSheet "
            SQL &= " Where NO =  '" & dtCommissionSheet.Rows(0).Item("NO") & "'"
            Dim dtSampleSheet As DataTable = uDataBase.GetDataTable(SQL)
            If dtSampleSheet.Rows.Count > 0 Then
                '----基本欄位設定-------------------------------------------------
                D3APPBUYER.Text = dtSampleSheet.Rows(0).Item("AppBuyer")                 'Customer
                D3DATE.Text = dtSampleSheet.Rows(0).Item("Date")                         '發行日
                D3SIZENO.Text = dtSampleSheet.Rows(0).Item("SizeNo")                     'Size
                D3ITEM.Text = dtSampleSheet.Rows(0).Item("Item")                         'Item
                D3CODENO.Text = dtSampleSheet.Rows(0).Item("CodeNo")                     'Code No
                '實際樣品
                If dtSampleSheet.Rows(0).Item("SampleFile") <> "" Then
                    LSAMPLEFILE.ImageUrl = Path & dtSampleSheet.Rows(0).Item("SampleFile")
                    LSAMPLEFILE.Visible = True
                End If
                '----開發規格-------------------------------------------------
                D3TAWIDTH.Text = dtSampleSheet.Rows(0).Item("TAWidth")                   '布帶寬度
                D3DEVNO.Text = dtSampleSheet.Rows(0).Item("DevNo")                       '開發No
                D3DEVPRD.Text = dtSampleSheet.Rows(0).Item("DevPrd")                     '開發期間
                D3TACOL.Text = dtSampleSheet.Rows(0).Item("TACol")                       '布帶Color
                D3TALINE.Text = dtSampleSheet.Rows(0).Item("TALine")                     '條紋線Color
                D3ECOL.Text = dtSampleSheet.Rows(0).Item("ECol")                         '務齒
                D3CCOL.Text = dtSampleSheet.Rows(0).Item("CCol")                         '丸紐
                D3THCOL.Text = dtSampleSheet.Rows(0).Item("THCol")                       '縫工線
                D3OTHER.Text = dtSampleSheet.Rows(0).Item("Other")                       '其他
                '----QC附檔-------------------------------------------------
                If dtSampleSheet.Rows(0).Item("QCFile1") <> "" Then                     '品測報告1
                    L3QCFILE1.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile1")
                    L3QCFILE1.Visible = True
                End If
                If dtSampleSheet.Rows(0).Item("QCFile2") <> "" Then                     '品測報告2
                    L3QCFILE2.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile2")
                    L3QCFILE2.Visible = True
                End If
                If dtSampleSheet.Rows(0).Item("QCFile3") <> "" Then                     '品測報告3
                    L3QCFILE3.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile3")
                    L3QCFILE3.Visible = True
                End If
                If dtSampleSheet.Rows(0).Item("QCFile4") <> "" Then                     '品測報告4
                    L3QCFILE4.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile4")
                    L3QCFILE4.Visible = True
                End If
                If dtSampleSheet.Rows(0).Item("QCFile5") <> "" Then                     '品測報告5
                    L3QCFILE5.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile5")
                    L3QCFILE5.Visible = True
                End If
                '----Wave's-------------------------------------------------
                D3TNLITEM.Text = dtSampleSheet.Rows(0).Item("TNLItem")                   'TAPE NAT-左
                D3TNRITEM.Text = dtSampleSheet.Rows(0).Item("TNRItem")                   'TAPE NAT-右
                D3TSLITEM.Text = dtSampleSheet.Rows(0).Item("TSLItem")                   'TAPE SET-左
                D3TSRITEM.Text = dtSampleSheet.Rows(0).Item("TSRItem")                   'TAPE SET-右
                D3TDLITEM.Text = dtSampleSheet.Rows(0).Item("TDLItem")                   'TAPE DYED-左
                D3TDRITEM.Text = dtSampleSheet.Rows(0).Item("TDRItem")                   'TAPE DYED-右
                D3CNITEM.Text = dtSampleSheet.Rows(0).Item("CNItem")                     'CHAIN NAT
                D3CSITEM.Text = dtSampleSheet.Rows(0).Item("CSItem")                     'CHAIN SET
                D3CDITEM.Text = dtSampleSheet.Rows(0).Item("CDItem")                     'CHAIN DYED
                D31Other.Text = dtSampleSheet.Rows(0).Item("Other1")                     'Other1
                D32Other.Text = dtSampleSheet.Rows(0).Item("Other2")                     'Other2
                D3O1ITEM.Text = dtSampleSheet.Rows(0).Item("O1Item")                     'O1Item
                D3O2ITEM.Text = dtSampleSheet.Rows(0).Item("O2Item")                     'O2Item
                D3CITEM.Text = dtSampleSheet.Rows(0).Item("CItem")                       'CORD
                '----FLOW-------------------------------------------------
                SetFieldData(FindFieldInf("3-FLOW"), "WF1", dtSampleSheet.Rows(0).Item("WF1"))          '承認WF1
                SetFieldData(FindFieldInf("3-FLOW"), "WF2", dtSampleSheet.Rows(0).Item("WF2"))          '承認WF2
                SetFieldData(FindFieldInf("3-FLOW"), "WF3", dtSampleSheet.Rows(0).Item("WF3"))          '承認WF3
                SetFieldData(FindFieldInf("3-FLOW"), "WF4", dtSampleSheet.Rows(0).Item("WF4"))          '承認WF4
                SetFieldData(FindFieldInf("3-FLOW"), "WF5", dtSampleSheet.Rows(0).Item("WF5"))          '承認WF5
                SetFieldData(FindFieldInf("3-FLOW"), "WF6", dtSampleSheet.Rows(0).Item("WF6"))          '承認WF6
                SetFieldData(FindFieldInf("3-FLOW"), "WF7", dtSampleSheet.Rows(0).Item("WF7"))          '承認WF7
                SetFieldData(FindFieldInf("3-FLOW"), "WF3NAME", dtSampleSheet.Rows(0).Item("WF3Name"))  '承認者部門WF3
                SetFieldData(FindFieldInf("3-FLOW"), "WF4NAME", dtSampleSheet.Rows(0).Item("WF4Name"))  '承認者部門WF4
                SetFieldData(FindFieldInf("3-FLOW"), "WF5NAME", dtSampleSheet.Rows(0).Item("WF5Name"))  '承認者部門WF5
                SetFieldData(FindFieldInf("3-FLOW"), "WF6NAME", dtSampleSheet.Rows(0).Item("WF6Name"))  '承認者部門WF6
                SetFieldData(FindFieldInf("3-FLOW"), "WF7NAME", dtSampleSheet.Rows(0).Item("WF7Name"))  '承認者部門WF7
            End If
            '-----------------------------------------------------------------
            '-- 原單位
            '-----------------------------------------------------------------
            SQL = "Select * From FS_GentaniSheet "
            SQL &= " Where NO =  '" & dtCommissionSheet.Rows(0).Item("NO") & "'"
            SQL &= " Union all "
            SQL &= "Select * From FS_GentaniSheetNew "
            SQL &= " Where NO =  '" & dtCommissionSheet.Rows(0).Item("NO") & "'"
            Dim dtGentaniSheet As DataTable = uDataBase.GetDataTable(SQL)
            If dtGentaniSheet.Rows.Count > 0 Then
                '
                Dim s As New List(Of String)
                s.Add("TNLITEM")
                s.Add("TNRITEM")
                s.Add("ECOL")
                s.Add("EITEM")
                '
                For Each dc As DataColumn In dtGentaniSheet.Columns
                    Dim l As Object = Me.form1.FindControl("D4" & dc.ColumnName)
                    Dim v As String = uCommon.ReplaceDBnull(dtGentaniSheet.Rows(0)(dc.ColumnName), "")
                    If l IsNot Nothing Then
                        l.Text = v
                    Else
                        'If dc.ColumnName = "TNLITEM" Then
                        '    D4TNLITEM1.Text = v
                        'ElseIf dc.ColumnName = "TNRITEM" Then
                        '    D4TNRITEM1.Text = v
                        'ElseIf dc.ColumnName = "EITEM" Then
                        '    D4EITEM1.Text = v

                        'End If


                        If s.Contains(dc.ColumnName) Then
                            l = Me.form1.FindControl("D4" & dc.ColumnName & "1")
                            l.Text = v
                            ' l = Me.form1.FindControl("D4" & dc.ColumnName & "2")
                            ' l.Text = v
                        End If
                    End If
                Next
                '
            End If
            '-----------------------------------------------------------------
            '-- 核定履歷
            '-----------------------------------------------------------------
            '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            'Modify-Start by Joy  2012/7/31 新納期對應
            '
            'SQL = "SELECT "
            'SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
            'SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
            'SQL = SQL + "'預定：[' + BStartTimeDesc + ' ~ ' + BEndTimeDesc + '], ' + "
            'SQL = SQL + "'實際：[' + AStartTimeDesc + ' ~ ' + AEndTimeDesc + '] ' As Description, "
            'SQL = SQL + "URL "
            'SQL = SQL + "FROM V_WaitHandle_01 "
            'SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
            'SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
            'SQL = SQL + "Order by Unique_ID Desc "
            'GridView1.DataSource = uDataBase.GetDataTable(SQL)
            'GridView1.DataBind()
            '
            SQL = "SELECT "
            SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
            SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
            SQL = SQL + "'' as Description, BStartTimeDesc, BEndTimeDesc, AStartTimeDesc, AEndTimeDesc, WorkID, "
            SQL = SQL + "'' As ViewOP, '' AS URL, Active "
            SQL = SQL + "FROM V_WaitHandle_01 "
            SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
            SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
            SQL = SQL + "  And Active <> '9' "
            'SQL = SQL + "Order by Unique_ID Desc "
            SQL = SQL + "Order by CreateTime Desc, Step Desc, SeqNo Desc "
            Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
            If dt_WaitHandle.Rows.Count > 0 Then
                For i As Integer = 0 To dt_WaitHandle.Rows.Count - 1
                    ' StepNameDesc(9)
                    dt_WaitHandle.Rows(i)(9) = "[" + dt_WaitHandle.Rows(i)(2).ToString + "]-" + Replace(dt_WaitHandle.Rows(i)(9), "開發委託_", "")
                    '
                    ' Description(12)
                    ' 預：[2012-04-25 16:00 ~ 2012-04-25 17:00](380),
                    ' 實：[2012-04-25 16:00 ~ 2012-04-26 14:10](380)
                    Dim xTime As Integer = 0
                    ' 預
                    xTime = 0
                    If InStr(dt_WaitHandle.Rows(i)(14), "無記錄") > 0 Then
                    Else
                        dt_WaitHandle.Rows(i)(12) = "預：[" + CDate(dt_WaitHandle.Rows(i)(13)).ToString("yyyy/MM/dd HH:mm") + " ~ " + CDate(dt_WaitHandle.Rows(i)(14)).ToString("yyyy/MM/dd HH:mm")
                        oSchedule.GetDevelopTime(CDate(dt_WaitHandle.Rows(i)(13)).ToString("yyyy/MM/dd HH:mm"), _
                                                 CDate(dt_WaitHandle.Rows(i)(14)).ToString("yyyy/MM/dd HH:mm"), _
                                                 xTime, _
                                                 oCommon.GetCalendarGroup(dt_WaitHandle.Rows(i)(17)))
                        dt_WaitHandle.Rows(i)(12) = dt_WaitHandle.Rows(i)(12) + "  (" + CStr(xTime) + ")]"
                    End If
                    ' 實
                    xTime = 0
                    If InStr(dt_WaitHandle.Rows(i)(16), "無記錄") > 0 Then
                    Else
                        dt_WaitHandle.Rows(i)(12) = dt_WaitHandle.Rows(i)(12) + ", " + _
                                                    "實：[" + CDate(dt_WaitHandle.Rows(i)(15)).ToString("yyyy/MM/dd HH:mm") + " ~ " + CDate(dt_WaitHandle.Rows(i)(16)).ToString("yyyy/MM/dd HH:mm")
                        oSchedule.GetDevelopTime(CDate(dt_WaitHandle.Rows(i)(15)).ToString("yyyy/MM/dd HH:mm"), _
                                                 CDate(dt_WaitHandle.Rows(i)(16)).ToString("yyyy/MM/dd HH:mm"), _
                                                 xTime, _
                                                 oCommon.GetCalendarGroup(dt_WaitHandle.Rows(i)(17)))
                        dt_WaitHandle.Rows(i)(12) = dt_WaitHandle.Rows(i)(12) + "  (" + CStr(xTime) + ")]"
                    End If
                    ' ViewOP(18), URL(19)
                    If GetFlowLoading(dt_WaitHandle.Rows(i)(2)) = 1 Then
                        dt_WaitHandle.Rows(i)(18) = "@"
                        If dt_WaitHandle.Rows(i)(20) <> 1 Then
                            dt_WaitHandle.Rows(i)(19) = "CommissionSheet_Load.aspx?pFormNo=" + dt_WaitHandle.Rows(i)(0).ToString + _
                                                                                 "&pFormSno=" + dt_WaitHandle.Rows(i)(1).ToString + _
                                                                                 "&pWID=" + dt_WaitHandle.Rows(i)(17).ToString + _
                                                                                 "&pBEndTime=" + CDate(dt_WaitHandle.Rows(i)(14)).ToString("yyyyMMddHHmm") + _
                                                                                 "&pAEndTime=" + CDate(dt_WaitHandle.Rows(i)(16)).ToString("yyyyMMddHHmm")
                        Else
                            dt_WaitHandle.Rows(i)(19) = "CommissionSheet_Load.aspx?pFormNo=" + dt_WaitHandle.Rows(i)(0).ToString + _
                                                                                 "&pFormSno=" + dt_WaitHandle.Rows(i)(1).ToString + _
                                                                                 "&pWID=" + dt_WaitHandle.Rows(i)(17).ToString + _
                                                                                 "&pBEndTime=" + CDate(dt_WaitHandle.Rows(i)(14)).ToString("yyyyMMddHHmm") + _
                                                                                 "&pAEndTime=" + ""
                        End If
                    End If
                Next
            End If
            GridView1.DataSource = dt_WaitHandle
            GridView1.DataBind()
            'Modify-End
            '
            '廠長用開發履歷  ADD-2011/10/15
            If UCase(Request.QueryString("pUserID")) = "FADM02" Then
                GridView2.DataSource = uDataBase.GetDataTable(SQL)
                GridView2.DataBind()
            End If
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
    '-------------------------------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     儲存按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BSAVE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSAVE.Click
        If InputDataOK(3) Then
            DisabledButton()   '停止Button運作
            'Modify-2012/3/16-Joy 160工程追加[開發中止] (因Button已滿所以使用 Save Button)
            '
            If wStep = 160 Then
                FlowControl("NG2", 3, "8")      'pFun=NG2, pAction=3, pSts=8 (T_Waithandle.Sts=8 SaveButton)                
            ElseIf wStep = 120 Then
                FlowControl("OK", 3, "8")      'pFun=NG2, pAction=3, pSts=8 (T_Waithandle.Sts=8 SaveButton)          
            ElseIf wStep = 32 Then
                FlowControl("OK", 3, "8")      'pFun=NG2, pAction=3, pSts=8 (T_Waithandle.Sts=8 SaveButton)        
            Else
                ModifyData("SAVE", "0")           '更新表單資料 Sts=0(未結)
                ModifyTranData("SAVE", "0")       '更新交易資料
                Dim URL As String = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=S&pFormNo=" & wFormNo & "&pStep=" & wStep & _
                                                    "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
            End If
            '
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
        Dim ErrCode As Integer = 0
        Dim Message As String = ""
        '
        '20工程製圖者不能空白    Jessica 2012/11/12
        If (wStep = 20) And (pAction = 1) Then
            If DMAKEMAP.Text = "" Then
                ErrCode = 9201

            End If
        End If



        '----表單-附加檔案-----------------------------------------------
        '開發委託-草圖
        If ErrCode = 0 Then
            If DMAPREFFILE.Visible Then
                If DMAPREFFILE.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DMAPREFFILE)
                End If
            End If
        End If
        '開發委託-圖檔
        If ErrCode = 0 Then
            If DMAPFILE.Visible Then
                If DMAPFILE.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DMAPFILE)
                End If
            End If
        End If
        '開發委託-適用型別檔
        If ErrCode = 0 Then
            If DFORTYPEFILE.Visible Then
                If DFORTYPEFILE.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DFORTYPEFILE)
                End If
            End If
        End If
        '開發委託-品質檢測項目檔
        If ErrCode = 0 Then
            If DQCCHECKFILE.Visible Then
                If DQCCHECKFILE.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DQCCHECKFILE)
                End If
            End If
        End If
        '開發委託-QC檢測文件
        If ErrCode = 0 Then
            If DQCFILE1.Visible Then
                If DQCFILE1.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DQCFILE1)
                End If
            End If
        End If
        If ErrCode = 0 Then
            If DQCFILE2.Visible Then
                If DQCFILE2.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DQCFILE2)
                End If
            End If
        End If
        If ErrCode = 0 Then
            If DQCFILE3.Visible Then
                If DQCFILE3.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DQCFILE3)
                End If
            End If
        End If
        If ErrCode = 0 Then
            If DQCFILE4.Visible Then
                If DQCFILE4.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DQCFILE4)
                End If
            End If
        End If
        If ErrCode = 0 Then
            If DQCFILE5.Visible Then
                If DQCFILE5.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DQCFILE5)
                End If
            End If
        End If
        If ErrCode = 0 Then
            If DQCFILE6.Visible Then
                If DQCFILE6.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DQCFILE6)
                End If
            End If
        End If

        If ErrCode = 0 Then
            If DGenFILE1.Visible Then
                If DGenFILE1.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DGenFILE1)
                End If
            End If
        End If

        '開發委託-客戶切結書
        If ErrCode = 0 Then
            If DCONTACTFILE.Visible Then
                If DCONTACTFILE.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DCONTACTFILE)
                End If
            End If
        End If
        '開發委託_切結書
        If ErrCode = 0 Then
            If DCONTACTFILE.Visible Then
                If DCONTACTFILE.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DCONTACTFILE)
                End If
            End If
        End If
        '開發委託-樣品確認書
        If ErrCode = 0 Then
            If DSAMPLECONFIRMFILE.Visible Then
                If DSAMPLECONFIRMFILE.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DSAMPLECONFIRMFILE)
                End If
            End If
        End If
        '開發委託-製造授權書
        If ErrCode = 0 Then
            If DMANUFAUTHORITYFILE.Visible Then
                If DMANUFAUTHORITYFILE.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DMANUFAUTHORITYFILE)
                End If
            End If
        End If
        '開發見本-樣品圖
        If ErrCode = 0 Then
            If D3SAMPLEFILE.Visible Then
                If D3SAMPLEFILE.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(D3SAMPLEFILE)
                End If
            End If
        End If
        '開發見本-品質-吋法
        If ErrCode = 0 Then
            If D3QCFILE1.Visible Then
                If D3QCFILE1.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(D3QCFILE1)
                End If
            End If
        End If
        '開發見本-品質-強度
        If ErrCode = 0 Then
            If D3QCFILE2.Visible Then
                If D3QCFILE2.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(D3QCFILE2)
                End If
            End If
        End If
        '開發見本-品質-往覆測試
        If ErrCode = 0 Then
            If D3QCFILE3.Visible Then
                If D3QCFILE3.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(D3QCFILE3)
                End If
            End If
        End If
        '開發見本-品質-式樣／組織
        If ErrCode = 0 Then
            If D3QCFILE4.Visible Then
                If D3QCFILE4.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(D3QCFILE4)
                End If
            End If
        End If
        '開發見本-品質-其他
        If ErrCode = 0 Then
            If D3QCFILE5.Visible Then
                If D3QCFILE5.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(D3QCFILE5)
                End If
            End If
        End If

        '20180807 選廠長只能選井田憲政
        Dim FADM As String = ""
        Dim sql2 As String = ""
        sql2 = " select * from  M_referp "
        sql2 &= " where CAT = 2002 AND DKEY = 'FADM' "
        Dim dtFADM As DataTable = uDataBase.GetDataTable(sql2)
        If dtFADM.Rows.Count > 0 Then
            FADM = dtFADM.Rows(0).Item("Data")
        End If


        If D3WF3Name.SelectedValue = "廠長" Then
            If D3WF3.SelectedValue <> FADM Then
                ErrCode = 9203
            End If
        End If

        If D3WF4Name.SelectedValue = "廠長" Then
            If D3WF4.SelectedValue <> FADM Then
                ErrCode = 9203
            End If
        End If


        If D3WF5Name.SelectedValue = "廠長" Then
            If D3WF5.SelectedValue <> FADM Then
                ErrCode = 9203
            End If
        End If

        If D3WF6Name.SelectedValue = "廠長" Then
            If D3WF6.SelectedValue <> FADM Then
                ErrCode = 9203
            End If
        End If



        If D3WF7Name.SelectedValue = "廠長" Then
            If D3WF7.SelectedValue <> FADM Then
                ErrCode = 9203
            End If
        End If


        If D3WF6.SelectedValue = FADM Then
            If D3WF6Name.SelectedValue <> "廠長" Then
                ErrCode = 9203
            End If
        End If


        If D3WF3.SelectedValue = FADM Then
            If D3WF3Name.SelectedValue <> "廠長" Then
                ErrCode = 9204
            End If
        End If

        If D3WF4.SelectedValue = FADM Then
            If D3WF4Name.SelectedValue <> "廠長" Then
                ErrCode = 9204
            End If
        End If

        If D3WF5.SelectedValue = FADM Then
            If D3WF5Name.SelectedValue <> "廠長" Then
                ErrCode = 9204
            End If
        End If

        If D3WF6.SelectedValue = FADM Then
            If D3WF6Name.SelectedValue <> "廠長" Then
                ErrCode = 9204
            End If
        End If

        If D3WF7.SelectedValue = FADM Then
            If D3WF7Name.SelectedValue <> "廠長" Then
                ErrCode = 9204
            End If
        End If

    

        '不是儲存按鈕需檢查資料
        If pAction < 3 Then
            '----表單-數字資料-----------------------------------------------
            '開發委託-企-樣品
            If ErrCode = 0 Then
                'If DPLEN.Text = "0" And DPQTY.Text = "0" Then ErrCode = 9090
            End If
            '開發委託-EA-樣品
            If ErrCode = 0 Then
                If wStep = 20 Then
                    If DEALEN.Text = "0" And DEAQTY.Text = "0" Then ErrCode = 9090
                End If
            End If
            '開發委託-企-長度
            If ErrCode = 0 Then
                If Not IsNumeric(DPLEN.Text) Then ErrCode = 9010
                'If InStr(DPLEN.Text, ".") > 0 Then ErrCode = 9040
            End If
            '開發委託-企-數量
            If ErrCode = 0 Then
                If Not IsNumeric(DPQTY.Text) Then ErrCode = 9010
                If InStr(DPQTY.Text, ".") > 0 Then ErrCode = 9040
            End If
            '開發委託-EA-長度
            If ErrCode = 0 Then
                If Not IsNumeric(DEALEN.Text) Then ErrCode = 9010
                'If InStr(DEALEN.Text, ".") > 0 Then ErrCode = 9040
            End If
            '開發委託-EA-數量
            If ErrCode = 0 Then
                If Not IsNumeric(DEAQTY.Text) Then ErrCode = 9010
                If InStr(DEAQTY.Text, ".") > 0 Then ErrCode = 9040
            End If
            '開發委託-布帶寬度
            If ErrCode = 0 Then
                If Not IsNumeric(DTAWIDTH.Text) Then ErrCode = 9010
                If InStr(DTAWIDTH.Text, ".") > 0 Then ErrCode = 9040
            End If
            '開發委託-X尺吋
            If ErrCode = 0 Then
                If Not IsNumeric(DXMLEN.Text) Then ErrCode = 9010
            End If
            '開發委託-A尺吋
            If ErrCode = 0 Then
                If Not IsNumeric(DAMLEN.Text) Then ErrCode = 9010
            End If
            '開發委託-B尺吋
            If ErrCode = 0 Then
                If Not IsNumeric(DBMLEN.Text) Then ErrCode = 9010
            End If
            '開發委託-C尺吋
            If ErrCode = 0 Then
                If Not IsNumeric(DCMLEN.Text) Then ErrCode = 9010
            End If
            '開發委託-D尺吋
            If ErrCode = 0 Then
                If Not IsNumeric(DDMLEN.Text) Then ErrCode = 9010
            End If
            '開發委託-E尺吋
            If ErrCode = 0 Then
                If Not IsNumeric(DEMLEN.Text) Then ErrCode = 9010
            End If
            '開發委託-F尺吋
            If ErrCode = 0 Then
                If Not IsNumeric(DFMLEN.Text) Then ErrCode = 9010
            End If
            '開發委託-E尺吋
            If ErrCode = 0 Then
                If Not IsNumeric(DGMLEN.Text) Then ErrCode = 9010
            End If
            '開發委託-E尺吋
            If ErrCode = 0 Then
                If Not IsNumeric(DHMLEN.Text) Then ErrCode = 9010
            End If
            '開發委託-E尺吋
            If ErrCode = 0 Then
                If Not IsNumeric(DLYLEN.Text) Then ErrCode = 9010
            End If
            '----表單-鏈齒/丸扭-----------------------------------------------
            If ErrCode = 0 Then
                If DECOLSEL.SelectedValue = "其他" Then
                    If DECOL.Text = "" Then ErrCode = 9190
                Else
                    If DECOL.Text <> "" Then ErrCode = 9190
                End If
                If DCCOLSEL.SelectedValue = "MF昭安" Then
                    If DCCOL.Text = "" Then ErrCode = 9200
                Else
                    If DCCOL.Text <> "" Then ErrCode = 9200
                End If
            End If
            '----表單-需樣品/需登錄-----------------------------------------------
            If ErrCode = 0 Then
                If DNeedSample.Checked = True Or DNeedItemRegister.Checked = True Then
                    If DREFNO.Text = "" Then
                        ErrCode = 9180
                    Else
                        Dim sql As String = ""
                        sql = "Select Sts From F_CommissionSheet "
                        sql &= "Where No  = '" & DREFNO.Text & "' "
                        sql &= "  And Sts = '2' "
                        Dim dtCommissionSheet As DataTable = uDataBase.GetDataTable(sql)
                        If dtCommissionSheet.Rows.Count <= 0 Then ErrCode = 9180
                    End If
                End If
            End If
            '----流程-資料-----------------------------------------------
            '開發委託_安排圖面或試作(Step=20)
            If ErrCode = 0 Then
                If wStep = 20 Then
                    If pAction = 0 Then
                        If DDEVTITLE.Text = "" Or DOP1PER.Text = "" Then ErrCode = 9080
                    End If
                End If
            End If
            '開發委託_試作NG對應(Step=120)
            If ErrCode = 0 Then
                If wStep = 120 Then
                    If pAction = 0 Then
                        Dim sql As String = ""
                        Dim wNextStep As Integer = 0
                        If DOP40.Checked = True Then wNextStep = 40 '40-工程
                        If DOP50.Checked = True Then wNextStep = 50 '50-工程
                        If DOP60.Checked = True Then wNextStep = 60 '60-工程
                        If DOP70.Checked = True Then wNextStep = 70 '70-工程
                        If DOP80.Checked = True Then wNextStep = 80 '80-工程
                        If DOP90.Checked = True Then wNextStep = 90 '90-工程
                        If DOP100.Checked = True Then wNextStep = 100 '100-工程
                        If DOP110.Checked = True Then wNextStep = 110 '110-工程
                        '
                        sql = "Select Top 1 Step, DecideID From T_WaitHandle "
                        sql &= "Where FormNo  =  '" & wFormNo & "' "
                        sql &= "  And FormSno =  '" & CStr(wFormSno) & "' "
                        sql &= "  And Step    = '" & CStr(wNextStep) & "' "
                        sql &= "  And Active  = '0' "
                        sql &= "  And FlowType <> '0' "
                        sql &= "Order by Unique_ID Desc "
                        Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(sql)
                        If dtWaitHandle.Rows.Count <= 0 Then ErrCode = 9150
                    End If
                End If
            End If
            '開發委託_客戶確認(Step=180)
            '--2012/10/9 Delete-Start by Joy
            'If ErrCode = 0 Then
            '    If wStep = 180 Then
            '        Dim sql As String = ""
            '        sql = "Select * From T_WaitHandle "
            '        sql &= "Where FormNo  =  '" & wFormNo & "' "
            '        sql &= "  And FormSno =  '" & CStr(wFormSno) & "' "
            '        sql &= "  And Step    = '260' "
            '        sql &= "Order by Unique_ID Desc "
            '        Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(sql)
            '        If dtWaitHandle.Rows.Count <= 0 Then ErrCode = 9170
            '    End If
            'End If
            '--2012/10/9 Delete-Start by Joy
            '開發委託_發行原單位及見本(Step=190)
            If ErrCode = 0 Then
                If wStep = 190 Then
                    If D3APPBUYER.Text = "" Or D3DEVNO.Text = "" Then ErrCode = 9110
                    If pAction = 0 Then
                        If D4NO.Text = "" Or D4DEVNO.Text = "" Then ErrCode = 9120
                    End If
                End If
            End If
            '----延遲-資料-----------------------------------------------
            '延遲理由內容
            If ErrCode = 0 Then
                Select Case wStep
                    Case 40
                        If DBEndTime.Text < NowDateTime Then
                            If DOP1DELAYC1.Text = "" Or DOP1DELAYC2.Text = "" Or DOP1REM.Text = "" Then ErrCode = 9100
                        End If
                    Case 50
                        If DBEndTime.Text < NowDateTime Then
                            If DOP2DELAYC1.Text = "" Or DOP2DELAYC2.Text = "" Or DOP2REM.Text = "" Then ErrCode = 9100
                        End If
                    Case 60
                        If DBEndTime.Text < NowDateTime Then
                            If DOP3DELAYC1.Text = "" Or DOP3DELAYC2.Text = "" Or DOP3REM.Text = "" Then ErrCode = 9100
                        End If
                    Case 70
                        If DBEndTime.Text < NowDateTime Then
                            If DOP4DELAYC1.Text = "" Or DOP4DELAYC2.Text = "" Or DOP4REM.Text = "" Then ErrCode = 9100
                        End If
                    Case 80
                        If DBEndTime.Text < NowDateTime Then
                            If DOP5DELAYC1.Text = "" Or DOP5DELAYC2.Text = "" Or DOP5REM.Text = "" Then ErrCode = 9100
                        End If
                    Case 90
                        If DBEndTime.Text < NowDateTime Then
                            If DOP6DELAYC1.Text = "" Or DOP6DELAYC2.Text = "" Or DOP6REM.Text = "" Then ErrCode = 9100
                        End If
                    Case 100
                        If DBEndTime.Text < NowDateTime Then
                            If DOP7DELAYC1.Text = "" Or DOP7DELAYC2.Text = "" Or DOP7REM.Text = "" Then ErrCode = 9100
                        End If
                    Case 110
                        If DBEndTime.Text < NowDateTime Then
                            If DOP8DELAYC1.Text = "" Or DOP8DELAYC2.Text = "" Or DOP8REM.Text = "" Then ErrCode = 9100
                        End If
                    Case Else
                End Select
            End If
            '----品質測試-資料-----------------------------------------------
            'QC-Leadtime
            If ErrCode = 0 Then
                If Not IsNumeric(DQCLT.Text) Then ErrCode = 9010
                If InStr(DQCLT.Text, ".") > 0 Then ErrCode = 9040
            End If



            If ErrCode = 0 Then
                If wStep = 130 And (pAction <> 2) Then   ' Jessica 2015/04/13增加非開發中止判斷 
                    If wStep = 130 And (pAction <> 1) Then   '  Jessica 2016/04/14增加非試作NG   
                        If CInt(DQCLT.Text) <= 0 Then ErrCode = 9160
                    End If
                End If


            End If


            ' Jessica 2015/04/13增加非開發中止判斷  '  Jessica 2016/04/14增加非試作NG   
            If ErrCode = 0 Then
                If (wStep = 130) And (pAction <> 2) Then
                    If (wStep = 130) And (pAction <> 1) Then
                        If DQCCHECKFILE.PostedFile.FileName = "" Then
                            ErrCode = 9202
                        End If
                    End If

                End If

            End If


            '----系統檢查-------------------------------------------------------
            '閱讀核定履歷
            If ErrCode = 0 Then
                If wFormSno > 0 And wStep > 3 Then
                    If Not DReadHistory.Checked Then ErrCode = 9070
                End If
            End If
            '核定說明
            If ErrCode = 0 Then
                If DDecideDesc.Visible = True Then
                    If DDecideDesc.Text = "" Then ErrCode = 9060
                End If
            End If
            '檢查圖號
            If ErrCode = 0 Then
                If DMAPNO.Text <> "" Then
                    ErrCode = CheckMapNo(DMAPNO.Text)
                    If ErrCode <> 0 Then
                        ErrCode = 9140
                    End If
                End If
            End If
            '檢查委託書No
            If ErrCode = 0 Then
                If DNO.Text <> "" Then
                    ErrCode = oCommon.CommissionNo("002002", wFormSno, wStep, DNO.Text) '表單號碼, 表單流水號, 工程, 委託書No
                    If ErrCode <> 0 Then
                        ErrCode = 9050
                    End If
                End If
            End If
        End If
        '----異常訊息處理----------------------------------------------------
        'ErrCode = 9999
        If ErrCode <> 0 Then
            If ErrCode = 9010 Then Message = "非有效數值,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案格式有誤,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案Size有誤,請確認!"
            If ErrCode = 9040 Then Message = "非整數值,請確認!"
            If ErrCode = 9050 Then Message = "委託書No.重覆,請確認委託書No.!"
            If ErrCode = 9060 Then Message = "核定說明需輸入,請確認!"
            If ErrCode = 9070 Then Message = "需閱讀核定履歷,請確認!"
            If ErrCode = 9080 Then Message = "製造委託書未輸入,請確認!"
            If ErrCode = 9090 Then Message = "樣品長度或數量未輸入,請確認!"
            If ErrCode = 9100 Then Message = "遲納原因及內容未輸入,請確認!"
            If ErrCode = 9110 Then Message = "開發見本或樣品圖未輸入,請確認!"
            If ErrCode = 9120 Then Message = "原單位未輸入,請確認!"
            If ErrCode = 9130 Then Message = "客戶切結書未輸入,請確認!"
            If ErrCode = 9140 Then Message = "圖號已存在,請確認!"
            If ErrCode = 9150 Then Message = "無法回試作工程,請確認!"
            If ErrCode = 9160 Then Message = "不可[<=0],請確認!"
            If ErrCode = 9170 Then Message = "開發見本廠長未簽核完成,請確認!"
            If ErrCode = 9180 Then Message = "不可選擇需樣品/需登錄,請確認!"
            If ErrCode = 9190 Then Message = "鏈齒輸入資料異常,請確認!"
            If ErrCode = 9200 Then Message = "丸扭輸入資料異常,請確認!"
            If ErrCode = 9201 Then Message = "製圖者請輸入,請確認!"
            If ErrCode = 9202 Then Message = "品質檢測項目檔未上傳,請確認!"
            If ErrCode = 9203 Then Message = "廠長姓名有誤!"
            If ErrCode = 9204 Then Message = "工程長姓名有誤!"
            If ErrCode = 9999 Then Message = "測試,請確認!"
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
    '**     檢查圖號是否重覆
    '**
    '*****************************************************************
    Function CheckMapNo(ByVal pMapNo As String) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        Try
            If wFormSno = 0 Then
                SQL = "SELECT MapNo From F_CommissionSheet "
                SQL &= " Where MapNo = '" & pMapNo & "' "
                Dim dt_FormTable As DataTable = uDataBase.GetDataTable(SQL)
                If dt_FormTable.Rows.Count > 0 Then
                    RtnCode = 1
                End If
            Else
                SQL = "SELECT MapNo From F_CommissionSheet "
                SQL &= " Where MapNo = '" & pMapNo & "' "
                SQL &= "   And FormSno <> '" & CStr(wFormSno) & "' "
                Dim dt_FormTable As DataTable = uDataBase.GetDataTable(SQL)
                If dt_FormTable.Rows.Count > 0 Then
                    RtnCode = 1
                End If
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
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
        Dim ErrCode As Integer = 0                  '宣告一個Error用來判別是否資料正常，預設為False
        Dim Message As String = ""
        Dim wQCLT As Integer = CInt(DQCLT.Text)     'QC-L/T
        Dim Run As Boolean = True                   '是否執行
        Dim RepeatRun As Boolean = False            '是否重覆執行
        Dim MultiJob As Integer = 0                 '多人核定
        Dim wLevel As String = DLEVEL.Text          '難易度
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
                        If wStep <> 260 Then pRunNextStep = 1 '是通知的重覆執行 260工程特殊處理
                    End If
                End If
            End If

            '取得下一關
            If ErrCode = 0 And pRunNextStep = 1 Then
                Dim SQL As String = ""
                Dim wAllocateID As String = ""
                '--指定簽核者處分---------------------------------------------------------------
                '[製造委託-指定簽核者/變更Action]
                Select Case wStep
                    Case 10         '-->營業窗口(已開需登錄)
                        If pAction = 0 Then
                            If DNeedItemRegister.Checked = True Then pAction = 2
                        End If
                        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                        'Add-Start by Joy  2012/7/31 新納期對應
                    Case 20
                        If pAction = 0 Then
                            ExpandFlowControl(wFormNo, wFormSno, wApplyID, Request.QueryString("pUserID"), 20, 130)
                        End If

                        'Add-End
                    Case 29         '-->OP1
                        If pAction = 0 Then
                            wAllocateID = oCommon.GetUserID(DOP1PER.Text)
                            wLevel = GetLevel(wLevel, HOP1.Text)            '取得難易度
                            UpdateLoading(40, HOP1.Text)                    '更新M_Flow負荷考量
                        End If
                    Case 32         '-->OP1
                        If pAction = 3 Then
                            wAllocateID = oCommon.GetUserID(DOP1PER.Text)
                            wLevel = GetLevel(wLevel, HOP1.Text)            '取得難易度
                            UpdateLoading(40, HOP1.Text)                    '更新M_Flow負荷考量
                        End If

                    Case 40         '-->OP2
                        If pAction = 0 Then
                            If DOP2PER.Text <> "" Then
                                wAllocateID = oCommon.GetUserID(DOP2PER.Text)
                            Else
                                If DNeedSample.Checked = False Then
                                    pAction = 2
                                Else    '已開需樣品
                                    pAction = 3
                                End If
                            End If
                            wLevel = GetLevel(wLevel, HOP2.Text)            '取得難易度
                            UpdateLoading(50, HOP2.Text)                    '更新M_Flow負荷考量
                        End If
                        'Add-Start by Joy  2012/7/31 新納期對應
                        'NG=動態產生Group成員(114工程=試作工程擔當)
                        If pAction = 1 Then
                            CreateGroupMember()
                        End If
                        'Add-End
                    Case 50         '-->OP3
                        If pAction = 0 Then
                            If DOP3PER.Text <> "" Then
                                wAllocateID = oCommon.GetUserID(DOP3PER.Text)
                            Else
                                If DNeedSample.Checked = False Then
                                    pAction = 2
                                Else    '已開需樣品
                                    pAction = 3
                                End If
                            End If
                            wLevel = GetLevel(wLevel, HOP3.Text)            '取得難易度
                            UpdateLoading(60, HOP3.Text)                    '更新M_Flow負荷考量
                        End If
                        'Add-Start by Joy  2012/7/31 新納期對應
                        'NG=動態產生Group成員(114工程=試作工程擔當)
                        If pAction = 1 Then
                            CreateGroupMember()
                        End If
                        'Add-End
                    Case 60         '-->OP4
                        If pAction = 0 Then
                            If DOP4PER.Text <> "" Then
                                wAllocateID = oCommon.GetUserID(DOP4PER.Text)
                            Else
                                If DNeedSample.Checked = False Then
                                    pAction = 2
                                Else    '已開需樣品
                                    pAction = 3
                                End If
                            End If
                            wLevel = GetLevel(wLevel, HOP4.Text)            '取得難易度
                            UpdateLoading(70, HOP4.Text)                    '更新M_Flow負荷考量
                        End If
                        'Add-Start by Joy  2012/7/31 新納期對應
                        'NG=動態產生Group成員(114工程=試作工程擔當)
                        If pAction = 1 Then
                            CreateGroupMember()
                        End If
                        'Add-End
                    Case 70         '-->OP5
                        If pAction = 0 Then
                            If DOP5PER.Text <> "" Then
                                wAllocateID = oCommon.GetUserID(DOP5PER.Text)
                            Else
                                pAction = 2
                            End If
                            wLevel = GetLevel(wLevel, HOP5.Text)            '取得難易度
                            UpdateLoading(80, HOP5.Text)                    '更新M_Flow負荷考量
                        End If
                        'Add-Start by Joy  2012/7/31 新納期對應
                        'NG=動態產生Group成員(114工程=試作工程擔當)
                        If pAction = 1 Then
                            CreateGroupMember()
                        End If
                        'Add-End
                    Case 80         '-->OP6
                        If pAction = 0 Then
                            If DOP6PER.Text <> "" Then
                                wAllocateID = oCommon.GetUserID(DOP6PER.Text)
                            Else
                                If DNeedSample.Checked = False Then
                                    pAction = 2
                                Else    '已開需樣品
                                    pAction = 3
                                End If
                            End If
                            wLevel = GetLevel(wLevel, HOP6.Text)            '取得難易度
                            UpdateLoading(90, HOP6.Text)                    '更新M_Flow負荷考量
                        End If
                        'Add-Start by Joy  2012/7/31 新納期對應
                        'NG=動態產生Group成員(114工程=試作工程擔當)
                        If pAction = 1 Then
                            CreateGroupMember()
                        End If
                        'Add-End
                    Case 90         '-->OP7
                        If pAction = 0 Then
                            If DOP7PER.Text <> "" Then
                                wAllocateID = oCommon.GetUserID(DOP7PER.Text)
                            Else
                                If DNeedSample.Checked = False Then
                                    pAction = 2
                                Else    '已開需樣品
                                    pAction = 3
                                End If
                            End If
                            wLevel = GetLevel(wLevel, HOP7.Text)            '取得難易度
                            UpdateLoading(100, HOP7.Text)                   '更新M_Flow負荷考量
                        End If
                        'Add-Start by Joy  2012/7/31 新納期對應
                        'NG=動態產生Group成員(114工程=試作工程擔當)
                        If pAction = 1 Then
                            CreateGroupMember()
                        End If
                        'Add-End
                    Case 100        '-->OP8
                        If pAction = 0 Then
                            If DOP8PER.Text <> "" Then
                                wAllocateID = oCommon.GetUserID(DOP8PER.Text)
                            Else
                                If DNeedSample.Checked = False Then
                                    pAction = 2
                                Else    '已開需樣品
                                    pAction = 3
                                End If
                            End If
                            wLevel = GetLevel(wLevel, HOP8.Text)            '取得難易度
                            UpdateLoading(110, HOP8.Text)                   '更新M_Flow負荷考量
                        End If
                        'Add-Start by Joy  2012/7/31 新納期對應
                        'NG=動態產生Group成員(114工程=試作工程擔當)
                        If pAction = 1 Then
                            CreateGroupMember()
                        End If
                        'Add-End
                    Case 110                                                '最後工程直跳170工程
                        If pAction = 0 Then
                            If DNeedSample.Checked = True Then pAction = 3 '已開需樣品
                        End If
                        'Add-Start by Joy  2012/7/31 新納期對應
                        'NG=動態產生Group成員(114工程=試作工程擔當)
                        If pAction = 1 Then
                            CreateGroupMember()
                        End If
                        'Add-End
                    Case 130        '-->準備品質檢測
                        'If InStr(DQCPEOPLE.SelectedValue, "會簽") <= 0 Then pAction = 1
                        'jessica Modity 20150413 
                        If pAction <> 2 And pAction <> 1 Then '非開發中止
                            If InStr(DQCPEOPLE.SelectedValue, "會簽") <= 0 Then pAction = 3
                        End If

                    Case 132        '-->準備品質檢測
                        If InStr(DQCPEOPLE.SelectedValue, "會簽") <= 0 Then
                            wAllocateID = oCommon.GetUserID(DQCPEOPLE.SelectedValue)
                        End If
                        '
                        'Add-Start by Joy  2012/7/31 新納期對應
                        '品質判定NG=動態產生Group成員(114工程=試作工程擔當)
                    Case 160
                        If pAction = 2 Then
                            CreateGroupMember()
                        End If
                        'Add-End
                    Case 170
                        If pAction = 0 Then
                            If DNeedSample.Checked = True Then pAction = 3 '已開需樣品(181)
                        End If
                    Case Else
                End Select
                '[製造委託-特殊跳躍]
                Select Case wStep
                    Case 20
                        If pAction = 1 Then  '2012/11/08   jessica
                            wAllocateID = oCommon.GetUserID(DMAKEMAP.Text)
                        End If
                    Case 32
                        If pAction = 2 Then   '2012/12/12   jessica
                            wAllocateID = oCommon.GetUserID(DMAKEMAP.Text)
                        End If
                    Case 34
                        If pAction = 1 Then  '2012/11/08   jessica
                            wAllocateID = oCommon.GetUserID(DMAKEMAP.Text)
                        End If
                    Case 120         '跳至指定工程==>自訂條件=[S]-001
                        If pAction = 0 Then
                            MultiJob = wFormSno                         'MultiJob==>FormSno
                            'pNextStep==>指定工程
                            If DOP40.Checked = True Then
                                pNextStep = 40                          '40-工程
                                wLevel = GetLevel(wLevel, HOP1.Text)    '取得難易度
                                UpdateLoading(40, HOP1.Text)            '更新M_Flow負荷考量
                            End If
                            If DOP50.Checked = True Then
                                pNextStep = 50                          '50-工程
                                wLevel = GetLevel(wLevel, HOP2.Text)    '取得難易度
                                UpdateLoading(50, HOP2.Text)            '更新M_Flow負荷考量
                            End If
                            If DOP60.Checked = True Then
                                pNextStep = 60                          '60-工程
                                wLevel = GetLevel(wLevel, HOP3.Text)    '取得難易度
                                UpdateLoading(60, HOP3.Text)            '更新M_Flow負荷考量
                            End If
                            If DOP70.Checked = True Then
                                pNextStep = 70                          '70-工程
                                wLevel = GetLevel(wLevel, HOP4.Text)    '取得難易度
                                UpdateLoading(70, HOP4.Text)            '更新M_Flow負荷考量
                            End If
                            If DOP80.Checked = True Then
                                pNextStep = 80                          '80-工程
                                wLevel = GetLevel(wLevel, HOP5.Text)    '取得難易度
                                UpdateLoading(80, HOP5.Text)            '更新M_Flow負荷考量
                            End If
                            If DOP90.Checked = True Then
                                pNextStep = 90                          '90-工程
                                wLevel = GetLevel(wLevel, HOP6.Text)    '取得難易度
                                UpdateLoading(90, HOP6.Text)            '更新M_Flow負荷考量
                            End If
                            If DOP100.Checked = True Then
                                pNextStep = 100                         '100-工程
                                wLevel = GetLevel(wLevel, HOP7.Text)    '取得難易度
                                UpdateLoading(100, HOP7.Text)           '更新M_Flow負荷考量
                            End If
                            If DOP110.Checked = True Then
                                pNextStep = 110                         '110-工程
                                wLevel = GetLevel(wLevel, HOP8.Text)    '取得難易度
                                UpdateLoading(110, HOP8.Text)           '更新M_Flow負荷考量
                            End If
                            UpdateManufInf(pNextStep)                   '更新試作工程資訊
                        End If
                        If pAction = 3 Then  '2012/11/08   jessica
                            wAllocateID = oCommon.GetUserID(DMAKEMAP.Text)
                        End If
                        If pAction = 1 Then UpdateManufInf(40) '更新試作工程資訊

                    Case Else
                End Select
                '[開發見本-指定簽核者]
                Select Case wStep
                    Case 190         '-->EA責任者
                        If pAction = 0 Then wAllocateID = oCommon.GetUserID(D3WF2.Text)
                    Case 200         '-->工程1責任者
                        If pAction = 0 Then
                            If D3WF3.Text <> "" Then
                                If D3WF3Name.Text = "廠長" Then
                                    pAction = 3
                                Else
                                    wAllocateID = oCommon.GetUserID(D3WF3.Text)
                                End If
                            Else
                                pAction = 2
                            End If
                        End If
                    Case 210         '-->工程2責任者
                        If pAction = 0 Then
                            If D3WF4.Text <> "" Then
                                If D3WF4Name.Text = "廠長" Then
                                    pAction = 3
                                Else
                                    wAllocateID = oCommon.GetUserID(D3WF4.Text)
                                End If

                            Else
                                pAction = 2
                            End If
                        End If
                    Case 211         '-->工程2責任者
                        If pAction = 0 Then
                            If D3WF4.Text <> "" Then
                                wAllocateID = oCommon.GetUserID(D3WF4.Text)
                            Else
                                pAction = 2
                            End If
                        End If
                    Case 220         '-->工程3責任者
                        If pAction = 0 Then
                            If D3WF5.Text <> "" Then
                                If D3WF5Name.Text = "廠長" Then
                                    pAction = 3
                                Else
                                    wAllocateID = oCommon.GetUserID(D3WF5.Text)
                                End If

                            Else
                                pAction = 2
                            End If
                        End If
                    Case 221         '-->工程2責任者
                        If pAction = 0 Then
                            If D3WF5.Text <> "" Then
                                wAllocateID = oCommon.GetUserID(D3WF5.Text)
                            Else
                                pAction = 2
                            End If
                        End If
                    Case 230         '-->工程4責任者
                        If pAction = 0 Then
                            If D3WF6.Text <> "" Then
                                If D3WF6Name.Text = "廠長" Then
                                    pAction = 3
                                Else
                                    wAllocateID = oCommon.GetUserID(D3WF6.Text)
                                End If

                            Else
                                pAction = 2
                            End If
                        End If
                    Case 231         '-->工程2責任者
                        If pAction = 0 Then
                            If D3WF6.Text <> "" Then
                                wAllocateID = oCommon.GetUserID(D3WF6.Text)
                            Else
                                pAction = 2
                            End If
                        End If
                    Case 240         '-->工程5責任者
                        If pAction = 0 Then
                            If D3WF7.Text <> "" Then
                                If D3WF7Name.Text = "廠長" Then
                                    pAction = 3
                                Else
                                    wAllocateID = oCommon.GetUserID(D3WF7.Text)
                                End If
                            Else
                                pAction = 2
                            End If
                        End If
                    Case 241         '-->工程2責任者
                        If pAction = 0 Then
                            If D3WF7.Text <> "" Then
                                wAllocateID = oCommon.GetUserID(D3WF7.Text)
                            Else
                                pAction = 2
                            End If
                        End If

                    Case Else
                End Select
                '[270工程:建置275工程(通知試作工程擔當)-動態產生Group成員]
                If wStep = 270 Then
                    If pAction = 0 Then
                        CreateGroupMember()
                    End If
                End If
                '--指定簽核者處分---------------------------------------------------------------
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
                        '
                        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                        'Modify-Start by Joy  2012/7/31 新納期對應
                        '
                        '取得工程負荷最後日時(核定者, 表單號碼, 工程號碼, 類別(0:通知,1:核定), 開始日時, 最後日時, 件數)
                        'oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)

                        '取得預定開始,完成日程計算(表單號碼,工程號碼,難易度,QC-L/T,現在時間, 預定開始日時, 預定完成日時, 行事曆)
                        'oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wNextGateCalendar)

                        '設定試作工程開始,預定時間及時數
                        'SetOPWorkTime("B", pNextStep, pStartTime, pEndTime, wNextGateCalendar)

                        '建置流程資料(表單號碼,表單流水號,工程關卡號碼,序號,申請者ID,行事曆,建置者, 簽核者, 被代理者, 申請者, 預定開始日, 預定完成日, 重要性)
                        'RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                        '
                        '------------------------------------------------------------------------
                        '
                        '取得工程負荷最後日時(核定者, 表單號碼, 工程號碼, 類別(0:通知,1:核定), 開始日時, 最後日時, 件數)
                        oSchedule.GetLastTimeCustom(pNextGate(i), wFormNo, wFormSno, pNextStep, pFlowType, NowDateTime, 0, pLastTime, pCount1)

                        '取得預定開始,完成日程計算(表單號碼,工程號碼,難易度,QC-L/T,現在時間, 預定開始日時, 預定完成日時, 行事曆)
                        oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wNextGateCalendar)

                        If pFlowType <> 0 Then      '非通知
                            '調整LeatTime
                            If (wStep = 32 And pNextStep = 20) Or (wStep = 34 And pNextStep = 20) Then
                                '製圖工程(32/34)回20工程時不需-調整LeatTime
                            Else
                                '取得工程負荷最後日時(表單號碼,表單流水號,下工程號碼,序號,行事曆,現在日時,預定開始日時,預定完成日時)
                                oSchedule.AdjustLeadTime(wFormNo, wFormSno, pNextStep, i, wNextGateCalendar, NowDateTime, pStartTime, pEndTime)
                            End If
                        End If

                        '設定試作工程開始,預定時間及時數
                        SetOPWorkTime("B", pNextStep, pStartTime, pEndTime, wNextGateCalendar)

                        '建置流程資料(表單號碼,表單流水號,工程關卡號碼,序號,申請者ID,行事曆,建置者, 簽核者, 被代理者, 申請者, 預定開始日, 預定完成日, 重要性)
                        RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                        '
                        '納期對應-刪除預定工程(Active=9 and FlowType=9)
                        oFlow.DeleteArrangFlow(wFormNo, wFormSno, pNextStep, i)
                        '
                        'Modify-End

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
                        ModifyData(pFun, "0")                       '更新表單資料 Sts=0(未結)
                    End If
                    AddCommissionNo(wFormNo, wFormSno)
                    '製造委託工程履歷  ADD-Start Joy 2012/2/2
                    If wStep >= 40 And wStep <= 110 Then
                        AddManufactureHistory(wFormNo, wFormSno, wStep, wSeqNo)
                    End If
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
    '**(SetOPWorkTime)
    '**     設定各工程開發時間
    '**
    '*****************************************************************
    Sub SetOPWorkTime(ByVal pType As String, ByVal pNextStep As Integer, _
                      ByVal pStartTime As DateTime, ByVal pEndTime As DateTime, _
                      ByVal pCalendar As String)
        Dim wDevelopTime As Integer = 0
        '
        oSchedule.GetDevelopTime(pStartTime.ToString("yyyy/MM/dd HH:mm:ss"), _
                                 pEndTime.ToString("yyyy/MM/dd HH:mm:ss"), _
                                 wDevelopTime, _
                                 pCalendar)
        If wDevelopTime < 0 Then wDevelopTime = 0
        '
        If pType = "B" Then
            '
            Select Case pNextStep
                Case 40
                    DOP1BTIME.Text = pStartTime.ToString("MM/dd HH:mm") + "~" + pEndTime.ToString("MM/dd HH:mm")    '預定開始~預定完成
                    DOP1BHOURS.Text = CStr(wDevelopTime)                                                            '預定時數
                Case 50
                    DOP2BTIME.Text = pStartTime.ToString("MM/dd HH:mm") + "~" + pEndTime.ToString("MM/dd HH:mm")    '預定開始~預定完成
                    DOP2BHOURS.Text = wDevelopTime.ToString                                                         '預定時數
                Case 60
                    DOP3BTIME.Text = pStartTime.ToString("MM/dd HH:mm") + "~" + pEndTime.ToString("MM/dd HH:mm")    '預定開始~預定完成
                    DOP3BHOURS.Text = wDevelopTime.ToString                                                         '預定時數
                Case 70
                    DOP4BTIME.Text = pStartTime.ToString("MM/dd HH:mm") + "~" + pEndTime.ToString("MM/dd HH:mm")    '預定開始~預定完成
                    DOP4BHOURS.Text = wDevelopTime.ToString                                                         '預定時數
                Case 80
                    DOP5BTIME.Text = pStartTime.ToString("MM/dd HH:mm") + "~" + pEndTime.ToString("MM/dd HH:mm")    '預定開始~預定完成
                    DOP5BHOURS.Text = wDevelopTime.ToString                                                         '預定時數
                Case 90
                    DOP6BTIME.Text = pStartTime.ToString("MM/dd HH:mm") + "~" + pEndTime.ToString("MM/dd HH:mm")    '預定開始~預定完成
                    DOP6BHOURS.Text = wDevelopTime.ToString                                                         '預定時數
                Case 100
                    DOP7BTIME.Text = pStartTime.ToString("MM/dd HH:mm") + "~" + pEndTime.ToString("MM/dd HH:mm")    '預定開始~預定完成
                    DOP7BHOURS.Text = wDevelopTime.ToString                                                         '預定時數
                Case 110
                    DOP8BTIME.Text = pStartTime.ToString("MM/dd HH:mm") + "~" + pEndTime.ToString("MM/dd HH:mm")    '預定開始~預定完成
                    DOP8BHOURS.Text = wDevelopTime.ToString                                                         '預定時數
                Case Else
            End Select
        Else
            Select Case pNextStep
                Case 40
                    DOP1ATIME.Text = pStartTime.ToString("MM/dd HH:mm") + "~" + pEndTime.ToString("MM/dd HH:mm")    '預定開始~實際完成
                    DOP1AHOURS.Text = wDevelopTime.ToString                                                         '實際時數
                Case 50
                    DOP2ATIME.Text = pStartTime.ToString("MM/dd HH:mm") + "~" + pEndTime.ToString("MM/dd HH:mm")    '預定開始~實際完成
                    DOP2AHOURS.Text = wDevelopTime.ToString                                                         '實際時數
                Case 60
                    DOP3ATIME.Text = pStartTime.ToString("MM/dd HH:mm") + "~" + pEndTime.ToString("MM/dd HH:mm")    '預定開始~實際完成
                    DOP3AHOURS.Text = wDevelopTime.ToString                                                         '實際時數
                Case 70
                    DOP4ATIME.Text = pStartTime.ToString("MM/dd HH:mm") + "~" + pEndTime.ToString("MM/dd HH:mm")    '預定開始~實際完成
                    DOP4AHOURS.Text = wDevelopTime.ToString                                                         '實際時數
                Case 80
                    DOP5ATIME.Text = pStartTime.ToString("MM/dd HH:mm") + "~" + pEndTime.ToString("MM/dd HH:mm")    '預定開始~實際完成
                    DOP5AHOURS.Text = wDevelopTime.ToString                                                         '實際時數
                Case 90
                    DOP6ATIME.Text = pStartTime.ToString("MM/dd HH:mm") + "~" + pEndTime.ToString("MM/dd HH:mm")    '預定開始~實際完成
                    DOP6AHOURS.Text = wDevelopTime.ToString                                                         '實際時數
                Case 100
                    DOP7ATIME.Text = pStartTime.ToString("MM/dd HH:mm") + "~" + pEndTime.ToString("MM/dd HH:mm")    '預定開始~實際完成
                    DOP7AHOURS.Text = wDevelopTime.ToString                                                         '實際時數
                Case 110
                    DOP8ATIME.Text = pStartTime.ToString("MM/dd HH:mm") + "~" + pEndTime.ToString("MM/dd HH:mm")    '預定開始~實際完成
                    DOP8AHOURS.Text = wDevelopTime.ToString                                                         '實際時數
                Case Else
            End Select
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     新增表單資料
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)
        Dim Path As String = Server.MapPath(uCommon.GetAppSetting("UploadFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim sql As String = ""
        '---------------------------------------------------------------------------------
        '-- 開發委託
        '---------------------------------------------------------------------------------
        sql = "INSERT INTO F_CommissionSheet ( " & _
              "Sts, CompletedTime, FormNo, FormSno, " & _
              "NO, REFNO, APPDATE, APPDEPT, APPPER, APPBUYER, " & _
              "SellVendor, ESYQTY, EXPDEL, CUSTITEM, USAGE, " & _
              "ORNO, NEEDMAP, MAPREFFILE, " & _
              "NeedSample, NeedItemRegister, " & _
              "PRO, OPPART, PLEN, PLENUN, PQTY, " & _
              "PQTYUN, EALEN, EALENUN, EAQTY, EAQTYUN, " & _
              "UPSLI, LOSLI, UPFIN, LOFIN, UPSTK, " & _
              "LOSTK, SPSPEC, " & _
              "SIZENO, ITEM, TATYPE, TAWIDTH, " & _
              "ECOLSEL, ECOL, CCOLSEL, CCOL, " & _
              "TACOL, TACOLNO, TAYCOLNO, " & _
              "TALCOL, TALCOLNO, TALYCOLNO, " & _
              "TARCOL, TARCOLNO, TARYCOLNO, " & _
              "THUPCOL, THUPCOLNO, THUPYCOLNO, " & _
              "THLUPCOL, THLUPCOLNO, THLUPYCOLNO, " & _
              "THLRUPCOL, THLRUPCOLNO, THLRUPYCOLNO, " & _
              "THRUPCOL, THRUPCOLNO, THRUPYCOLNO, " & _
              "THRRUPCOL, THRRUPCOLNO, THRRUPYCOLNO, " & _
              "THLOCOL, THLOCOLNO, THLOYCOLNO, " & _
              "THLLOCOL, THLLOCOLNO, THLLOYCOLNO, " & _
              "THLRLOCOL, THLRLOCOLNO, THLRLOYCOLNO, " & _
              "THRLOCOL, THRLOCOLNO, THRLOYCOLNO, " & _
              "THRRLOCOL, THRRLOCOLNO, THRRLOYCOLNO, " & _
              "XMLEN, XMCOL, XMCOLNO, XMYCOLNO, " & _
              "AMLEN, AMCOL, AMCOLNO, AMYCOLNO, " & _
              "BMLEN, BMCOL, BMCOLNO, BMYCOLNO, " & _
              "CMLEN, CMCOL, CMCOLNO, CMYCOLNO, " & _
              "DMLEN, DMCOL, DMCOLNO, DMYCOLNO, " & _
              "EMLEN, EMCOL, EMCOLNO, EMYCOLNO, " & _
              "FMLEN, FMCOL, FMCOLNO, FMYCOLNO, " & _
              "GMLEN, GMCOL, GMCOLNO, GMYCOLNO, " & _
              "HMLEN, HMCOL, HMCOLNO, HMYCOLNO, " & _
              "LYLEN, LYCOL, LYCOLNO, LYYCOLNO, " & _
              "OTCON, " & _
              "MAPNO, MAKEMAP, LEVEL, MAPFILE, " & _
              "FORTYPEFILE, " & _
              "QCCHECKFILE, " & _
              "QCFILE1, " & _
              "QCFILE2, " & _
              "QCFILE3, " & _
              "QCFILE4, " & _
              "QCFILE5, " & _
              "QCFILE6, " & _
              "GenFILE1, " & _
              "CONTACTFILE, " & _
              "SAMPLECONFIRMFILE, " & _
              "MANUFAUTHORITYFILE, " & _
              "MANUOUTPRICE, " & _
              "FORCASTFILE, " & _
              "CreateUser, CreateTime, ModifyUser, ModifyTime) " & _
        "VALUES("
        sql &= "'0' ,"
        sql &= "'" & NowDateTime & "' ,"
        sql &= "'" & wFormNo & "' ,"
        sql &= "'" & CStr(NewFormSno) & "' ,"
        '----------------------------------------------------------------------------------
        '--基本
        sql &= "'" & DNO.Text & "' ,"
        sql &= "'" & DREFNO.Text & "' ,"
        sql &= "'" & DAPPDATE.Text & "' ,"
        sql &= "'" & DAPPDEPT.Text & "' ,"
        sql &= "'" & DAPPPER.Text & "' ,"
        sql &= "'" & DAPPBUYER.SelectedValue & "' ,"
        sql &= "'" & DSellVendor.Text & "' ,"
        sql &= "'" & DESYQTY.Text & "' ,"
        sql &= "'" & DEXPDEL.Text & "' ,"
        sql &= "'" & DCUSTITEM.Text & "' ,"
        sql &= "'" & DUSAGE.SelectedValue & "' ,"
        sql &= "'" & DORNO.Text & "' ,"
        sql &= "'" & DNEEDMAP.SelectedValue & "' ,"
        Dim FileName As String = ""
        If DMAPREFFILE.Visible Then
            If DMAPREFFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(NewFormSno) & "-" & "MAPREFFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DMAPREFFILE.PostedFile.FileName)
                Try    '上傳圖檔
                    DMAPREFFILE.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "
        '--樣品
        If DNeedSample.Checked = True Then
            sql &= "'" & "1" & "' ,"
        Else
            sql &= "'" & "0" & "' ,"
        End If
        If DNeedItemRegister.Checked = True Then
            sql &= "'" & "1" & "' ,"
        Else
            sql &= "'" & "0" & "' ,"
        End If
        sql &= "'" & DPRO.SelectedValue & "' ,"
        sql &= "'" & DOPPART.Text & "' ,"
        sql &= "'" & DPLEN.Text & "' ,"
        sql &= "'" & DPLENUN.SelectedValue & "' ,"
        sql &= "'" & DPQTY.Text & "' ,"
        sql &= "'" & DPQTYUN.SelectedValue & "' ,"
        sql &= "'" & DEALEN.Text & "' ,"
        sql &= "'" & DEALENUN.SelectedValue & "' ,"
        sql &= "'" & DEAQTY.Text & "' ,"
        sql &= "'" & DEAQTYUN.SelectedValue & "' ,"
        sql &= "'" & DUPSLI.Text & "' ,"
        sql &= "'" & DLOSLI.Text & "' ,"
        sql &= "'" & DUPFIN.Text & "' ,"
        sql &= "'" & DLOFIN.Text & "' ,"
        sql &= "'" & DUPSTK.Text & "' ,"
        sql &= "'" & DLOSTK.Text & "' ,"
        sql &= "'" & DSPSPEC.Text & "' ,"
        '--開發
        sql &= "'" & DSIZENO.SelectedValue & "' ,"
        sql &= "'" & DITEM.SelectedValue & "' ,"
        sql &= "'" & DTATYPE.SelectedValue & "' ,"
        sql &= "'" & DTAWIDTH.Text & "' ,"
        sql &= "'" & DECOLSEL.SelectedValue & "' ,"
        sql &= "'" & DECOL.Text & "' ,"
        sql &= "'" & DCCOLSEL.SelectedValue & "' ,"
        sql &= "'" & DCCOL.Text & "' ,"
        sql &= "'" & DTACOL.SelectedValue & "' ,"
        sql &= "'" & DTACOLNO.Text & "' ,"
        sql &= "'" & DTAYCOLNO.Text & "' ,"
        sql &= "'" & DTALCOL.SelectedValue & "' ,"
        sql &= "'" & DTALCOLNO.Text & "' ,"
        sql &= "'" & DTALYCOLNO.Text & "' ,"
        sql &= "'" & DTARCOL.SelectedValue & "' ,"
        sql &= "'" & DTARCOLNO.Text & "' ,"
        sql &= "'" & DTARYCOLNO.Text & "' ,"
        sql &= "'" & DTHUPCOL.SelectedValue & "' ,"
        sql &= "'" & DTHUPCOLNO.Text & "' ,"
        sql &= "'" & DTHUPYCOLNO.Text & "' ,"
        sql &= "'" & DTHLUPCOL.SelectedValue & "' ,"
        sql &= "'" & DTHLUPCOLNO.Text & "' ,"
        sql &= "'" & DTHLUPYCOLNO.Text & "' ,"
        sql &= "'" & DTHLRUPCOL.SelectedValue & "' ,"
        sql &= "'" & DTHLRUPCOLNO.Text & "' ,"
        sql &= "'" & DTHLRUPYCOLNO.Text & "' ,"
        sql &= "'" & DTHRUPCOL.SelectedValue & "' ,"
        sql &= "'" & DTHRUPCOLNO.Text & "' ,"
        sql &= "'" & DTHRUPYCOLNO.Text & "' ,"
        sql &= "'" & DTHRRUPCOL.SelectedValue & "' ,"
        sql &= "'" & DTHRRUPCOLNO.Text & "' ,"
        sql &= "'" & DTHRRUPYCOLNO.Text & "' ,"
        sql &= "'" & DTHLOCOL.SelectedValue & "' ,"
        sql &= "'" & DTHLOCOLNO.Text & "' ,"
        sql &= "'" & DTHLOYCOLNO.Text & "' ,"
        sql &= "'" & DTHLLOCOL.SelectedValue & "' ,"
        sql &= "'" & DTHLLOCOLNO.Text & "' ,"
        sql &= "'" & DTHLLOYCOLNO.Text & "' ,"
        sql &= "'" & DTHLRLOCOL.SelectedValue & "' ,"
        sql &= "'" & DTHLRLOCOLNO.Text & "' ,"
        sql &= "'" & DTHLRLOYCOLNO.Text & "' ,"
        sql &= "'" & DTHRLOCOL.SelectedValue & "' ,"
        sql &= "'" & DTHRLOCOLNO.Text & "' ,"
        sql &= "'" & DTHRLOYCOLNO.Text & "' ,"
        sql &= "'" & DTHRRLOCOL.SelectedValue & "' ,"
        sql &= "'" & DTHRRLOCOLNO.Text & "' ,"
        sql &= "'" & DTHRRLOYCOLNO.Text & "' ,"
        sql &= "'" & DXMLEN.Text & "' ,"
        sql &= "'" & DXMCOL.SelectedValue & "' ,"
        sql &= "'" & DXMCOLNO.Text & "' ,"
        sql &= "'" & DXMYCOLNO.Text & "' ,"
        sql &= "'" & DAMLEN.Text & "' ,"
        sql &= "'" & DAMCOL.SelectedValue & "' ,"
        sql &= "'" & DAMCOLNO.Text & "' ,"
        sql &= "'" & DAMYCOLNO.Text & "' ,"
        sql &= "'" & DBMLEN.Text & "' ,"
        sql &= "'" & DBMCOL.SelectedValue & "' ,"
        sql &= "'" & DBMCOLNO.Text & "' ,"
        sql &= "'" & DBMYCOLNO.Text & "' ,"
        sql &= "'" & DCMLEN.Text & "' ,"
        sql &= "'" & DCMCOL.SelectedValue & "' ,"
        sql &= "'" & DCMCOLNO.Text & "' ,"
        sql &= "'" & DCMYCOLNO.Text & "' ,"
        sql &= "'" & DDMLEN.Text & "' ,"
        sql &= "'" & DDMCOL.SelectedValue & "' ,"
        sql &= "'" & DDMCOLNO.Text & "' ,"
        sql &= "'" & DDMYCOLNO.Text & "' ,"
        sql &= "'" & DEMLEN.Text & "' ,"
        sql &= "'" & DEMCOL.SelectedValue & "' ,"
        sql &= "'" & DEMCOLNO.Text & "' ,"
        sql &= "'" & DEMYCOLNO.Text & "' ,"
        sql &= "'" & DFMLEN.Text & "' ,"
        sql &= "'" & DFMCOL.SelectedValue & "' ,"
        sql &= "'" & DFMCOLNO.Text & "' ,"
        sql &= "'" & DFMYCOLNO.Text & "' ,"
        sql &= "'" & DGMLEN.Text & "' ,"
        sql &= "'" & DGMCOL.SelectedValue & "' ,"
        sql &= "'" & DGMCOLNO.Text & "' ,"
        sql &= "'" & DGMYCOLNO.Text & "' ,"
        sql &= "'" & DHMLEN.Text & "' ,"
        sql &= "'" & DHMCOL.SelectedValue & "' ,"
        sql &= "'" & DHMCOLNO.Text & "' ,"
        sql &= "'" & DHMYCOLNO.Text & "' ,"
        sql &= "'" & DLYLEN.Text & "' ,"
        sql &= "'" & DLYCOL.SelectedValue & "' ,"
        sql &= "'" & DLYCOLNO.Text & "' ,"
        sql &= "'" & DLYYCOLNO.Text & "' ,"
        sql &= "'" & DOTCON.Text & "' ,"
        '--圖面
        sql &= "'" & DMAPNO.Text & "' ,"
        sql &= "'" & DMAKEMAP.SelectedValue & "' ,"
        sql &= "'" & DLEVEL.SelectedValue & "' ,"
        FileName = ""
        If DMAPFILE.Visible Then
            If DMAPFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(NewFormSno) & "-" & "MAPFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DMAPFILE.PostedFile.FileName)
                Try    '上傳圖檔
                    DMAPFILE.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "
        '--適用型別
        FileName = ""
        If DFORTYPEFILE.Visible Then
            If DFORTYPEFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(NewFormSno) & "-" & "FORTYPEFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DFORTYPEFILE.PostedFile.FileName)
                Try    '上傳圖檔
                    DFORTYPEFILE.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "
        '--品質檢測項目
        FileName = ""
        If DQCCHECKFILE.Visible Then
            If DQCCHECKFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(NewFormSno) & "-" & "QCCHECKFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCCHECKFILE.PostedFile.FileName)
                Try    '上傳圖檔
                    DQCCHECKFILE.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "
        '--QC檢測文件
        FileName = ""
        If DQCFILE1.Visible Then
            If DQCFILE1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(NewFormSno) & "-" & "QCFILE1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE1.PostedFile.FileName)
                Try    '上傳圖檔
                    DQCFILE1.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "
        '
        FileName = ""
        If DQCFILE2.Visible Then
            If DQCFILE2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(NewFormSno) & "-" & "QCFILE2" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE2.PostedFile.FileName)
                Try    '上傳圖檔
                    DQCFILE2.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "
        '
        FileName = ""
        If DQCFILE3.Visible Then
            If DQCFILE3.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(NewFormSno) & "-" & "QCFILE3" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE3.PostedFile.FileName)
                Try    '上傳圖檔
                    DQCFILE3.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "
        '
        FileName = ""
        If DQCFILE4.Visible Then
            If DQCFILE4.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(NewFormSno) & "-" & "QCFILE4" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE4.PostedFile.FileName)
                Try    '上傳圖檔
                    DQCFILE4.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "
        '
        FileName = ""
        If DQCFILE5.Visible Then
            If DQCFILE5.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(NewFormSno) & "-" & "QCFILE5" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE5.PostedFile.FileName)
                Try    '上傳圖檔
                    DQCFILE5.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "
        '
        FileName = ""
        If DQCFILE6.Visible Then
            If DQCFILE6.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(NewFormSno) & "-" & "QCFILE6" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE6.PostedFile.FileName)
                Try    '上傳圖檔
                    DQCFILE6.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "


        '
        FileName = ""
        If DGenFILE1.Visible Then
            If DGenFILE1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(NewFormSno) & "-" & "GenFILE1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DGenFILE1.PostedFile.FileName)
                Try    '上傳圖檔
                    DGenFILE1.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "


        '--客戶切結書
        FileName = ""
        If DCONTACTFILE.Visible Then
            If DCONTACTFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(NewFormSno) & "-" & "CONTACTFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DCONTACTFILE.PostedFile.FileName)
                Try    '上傳圖檔
                    DCONTACTFILE.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "
        '--樣品確認書
        FileName = ""
        If DSAMPLECONFIRMFILE.Visible Then
            If DSAMPLECONFIRMFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(NewFormSno) & "-" & "SAMPLECONFIRMFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSAMPLECONFIRMFILE.PostedFile.FileName)
                Try    '上傳圖檔
                    DSAMPLECONFIRMFILE.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "
        '--製造授權書
        FileName = ""
        If DMANUFAUTHORITYFILE.Visible Then
            If DMANUFAUTHORITYFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(NewFormSno) & "-" & "MANUFAUTHORITYFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DMANUFAUTHORITYFILE.PostedFile.FileName)
                Try    '上傳圖檔
                    DMANUFAUTHORITYFILE.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "

        sql &= "'" & DMANUOUTPRICE.Text & "' ,"
        FileName = ""
        If DFORCASTFILE.Visible Then
            If DFORCASTFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(NewFormSno) & "-" & "FORCASTFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DFORCASTFILE.PostedFile.FileName)
                Try    '上傳圖檔
                    DFORCASTFILE.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        End If
        sql = sql + " '" + FileName + "', "
        '----------------------------------------------------------------------------------
        sql &= "'" & Request.QueryString("pUserID") & "' ,"
        sql &= "'" & NowDateTime & "' ,"
        sql &= "'" & Request.QueryString("pUserID") & "' ,"
        sql &= "'" & NowDateTime & "' )"
        uDataBase.ExecuteNonQuery(sql)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AddManufactureHistory)
    '**     製造委託工程履歷
    '**
    '*****************************************************************
    Sub AddManufactureHistory(ByVal pFormNo As String, ByVal pFormSno As Integer, ByVal pStep As Integer, ByVal pSeqNo As Integer)
        Dim sql As String = ""
        Dim wID As Integer = 0
        '
        '取得開發履歷ID
        sql = "Select Unique_ID From T_WaitHandle "
        sql &= "Where FormNo  = '" & pFormNo & "' "
        sql &= "  And FormSno = '" & CStr(pFormSno) & "' "
        sql &= "  And Step    = '" & CStr(pStep) & "' "
        sql &= "  And SeqNo   = '" & CStr(pSeqNo) & "' "
        sql &= "Order by Unique_ID Desc "
        Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(sql)
        If dt_WaitHandle.Rows.Count > 0 Then
            wID = dt_WaitHandle.Rows(0).Item("Unique_ID")
        End If
        '刪除 同開發履歷ID之製造委託工程履歷(避免重覆)
        sql = "Delete From FS_ManufactureHistory "
        sql &= "Where ID = '" & CStr(wID) & "' "
        uDataBase.ExecuteNonQuery(sql)
        '新增 製造委託工程履歷
        sql = "INSERT INTO FS_ManufactureHistory ( " & _
                "ID, FormNo, FormSno, Step, SeqNo, OPCode, " & _
                "OP, OPPER, OPBTIME, OPBHOURS, OPATIME, " & _
                "OPAHOURS, OPCON, OPDELAYC1, OPDELAYC2, OPREM, " & _
                "CreateUser, CreateTime, ModifyUser, ModifyTime) " & _
          "VALUES("
        sql &= "'" & CStr(wID) & "' ,"
        sql &= "'" & pFormNo & "' ,"
        sql &= "'" & CStr(pFormSno) & "' ,"
        sql &= "'" & CStr(pStep) & "' ,"
        sql &= "'" & CStr(pSeqNo) & "' ,"
        '----------------------------------------------------------------------------------
        Select Case pStep
            Case 40         ' OP1
                sql &= "'" & fpObj.GetOPCode("CODE-" & HOP1.Text) & "' ,"
                sql &= "'" & HOP1.Text & "' ,"
                sql &= "'" & DOP1PER.Text & "' ,"
                sql &= "'" & DOP1BTIME.Text & "' ,"
                sql &= "'" & DOP1BHOURS.Text & "' ,"
                sql &= "'" & DOP1ATIME.Text & "' ,"
                sql &= "'" & DOP1AHOURS.Text & "' ,"
                sql &= "'" & fpObj.ReplaceString(DOP1CON.Text) & "' ,"
                sql &= "'" & DOP1DELAYC1.SelectedValue & "' ,"
                sql &= "'" & DOP1DELAYC2.SelectedValue & "' ,"
                sql &= "'" & fpObj.ReplaceString(DOP1REM.Text) & "' ,"
            Case 50         ' OP2
                sql &= "'" & fpObj.GetOPCode("CODE-" & HOP2.Text) & "' ,"
                sql &= "'" & HOP2.Text & "' ,"
                sql &= "'" & DOP2PER.Text & "' ,"
                sql &= "'" & DOP2BTIME.Text & "' ,"
                sql &= "'" & DOP2BHOURS.Text & "' ,"
                sql &= "'" & DOP2ATIME.Text & "' ,"
                sql &= "'" & DOP2AHOURS.Text & "' ,"
                sql &= "'" & fpObj.ReplaceString(DOP2CON.Text) & "' ,"
                sql &= "'" & DOP2DELAYC1.SelectedValue & "' ,"
                sql &= "'" & DOP2DELAYC2.SelectedValue & "' ,"
                sql &= "'" & fpObj.ReplaceString(DOP2REM.Text) & "' ,"
            Case 60         ' OP3
                sql &= "'" & fpObj.GetOPCode("CODE-" & HOP3.Text) & "' ,"
                sql &= "'" & HOP3.Text & "' ,"
                sql &= "'" & DOP3PER.Text & "' ,"
                sql &= "'" & DOP3BTIME.Text & "' ,"
                sql &= "'" & DOP3BHOURS.Text & "' ,"
                sql &= "'" & DOP3ATIME.Text & "' ,"
                sql &= "'" & DOP3AHOURS.Text & "' ,"
                sql &= "'" & fpObj.ReplaceString(DOP3CON.Text) & "' ,"
                sql &= "'" & DOP3DELAYC1.SelectedValue & "' ,"
                sql &= "'" & DOP3DELAYC2.SelectedValue & "' ,"
                sql &= "'" & fpObj.ReplaceString(DOP3REM.Text) & "' ,"
            Case 70         ' OP4
                sql &= "'" & fpObj.GetOPCode("CODE-" & HOP4.Text) & "' ,"
                sql &= "'" & HOP4.Text & "' ,"
                sql &= "'" & DOP4PER.Text & "' ,"
                sql &= "'" & DOP4BTIME.Text & "' ,"
                sql &= "'" & DOP4BHOURS.Text & "' ,"
                sql &= "'" & DOP4ATIME.Text & "' ,"
                sql &= "'" & DOP4AHOURS.Text & "' ,"
                sql &= "'" & fpObj.ReplaceString(DOP4CON.Text) & "' ,"
                sql &= "'" & DOP4DELAYC1.SelectedValue & "' ,"
                sql &= "'" & DOP4DELAYC2.SelectedValue & "' ,"
                sql &= "'" & fpObj.ReplaceString(DOP4REM.Text) & "' ,"
            Case 80         ' OP5
                sql &= "'" & fpObj.GetOPCode("CODE-" & HOP5.Text) & "' ,"
                sql &= "'" & HOP5.Text & "' ,"
                sql &= "'" & DOP5PER.Text & "' ,"
                sql &= "'" & DOP5BTIME.Text & "' ,"
                sql &= "'" & DOP5BHOURS.Text & "' ,"
                sql &= "'" & DOP5ATIME.Text & "' ,"
                sql &= "'" & DOP5AHOURS.Text & "' ,"
                sql &= "'" & fpObj.ReplaceString(DOP5CON.Text) & "' ,"
                sql &= "'" & DOP5DELAYC1.SelectedValue & "' ,"
                sql &= "'" & DOP5DELAYC2.SelectedValue & "' ,"
                sql &= "'" & fpObj.ReplaceString(DOP5REM.Text) & "' ,"
            Case 90         ' OP6
                sql &= "'" & fpObj.GetOPCode("CODE-" & HOP6.Text) & "' ,"
                sql &= "'" & HOP6.Text & "' ,"
                sql &= "'" & DOP6PER.Text & "' ,"
                sql &= "'" & DOP6BTIME.Text & "' ,"
                sql &= "'" & DOP6BHOURS.Text & "' ,"
                sql &= "'" & DOP6ATIME.Text & "' ,"
                sql &= "'" & DOP6AHOURS.Text & "' ,"
                sql &= "'" & fpObj.ReplaceString(DOP6CON.Text) & "' ,"
                sql &= "'" & DOP6DELAYC1.SelectedValue & "' ,"
                sql &= "'" & DOP6DELAYC2.SelectedValue & "' ,"
                sql &= "'" & fpObj.ReplaceString(DOP6REM.Text) & "' ,"
            Case 100        ' OP7
                sql &= "'" & fpObj.GetOPCode("CODE-" & HOP7.Text) & "' ,"
                sql &= "'" & HOP7.Text & "' ,"
                sql &= "'" & DOP7PER.Text & "' ,"
                sql &= "'" & DOP7BTIME.Text & "' ,"
                sql &= "'" & DOP7BHOURS.Text & "' ,"
                sql &= "'" & DOP7ATIME.Text & "' ,"
                sql &= "'" & DOP7AHOURS.Text & "' ,"
                sql &= "'" & fpObj.ReplaceString(DOP7CON.Text) & "' ,"
                sql &= "'" & DOP7DELAYC1.SelectedValue & "' ,"
                sql &= "'" & DOP7DELAYC2.SelectedValue & "' ,"
                sql &= "'" & fpObj.ReplaceString(DOP7REM.Text) & "' ,"
            Case 110        ' OP8
                sql &= "'" & fpObj.GetOPCode("CODE-" & HOP8.Text) & "' ,"
                sql &= "'" & HOP8.Text & "' ,"
                sql &= "'" & DOP8PER.Text & "' ,"
                sql &= "'" & DOP8BTIME.Text & "' ,"
                sql &= "'" & DOP8BHOURS.Text & "' ,"
                sql &= "'" & DOP8ATIME.Text & "' ,"
                sql &= "'" & DOP8AHOURS.Text & "' ,"
                sql &= "'" & fpObj.ReplaceString(DOP8CON.Text) & "' ,"
                sql &= "'" & DOP8DELAYC1.SelectedValue & "' ,"
                sql &= "'" & DOP8DELAYC2.SelectedValue & "' ,"
                sql &= "'" & fpObj.ReplaceString(DOP8REM.Text) & "' ,"
        End Select
        '----------------------------------------------------------------------------------
        sql &= "'" & Request.QueryString("pUserID") & "' ,"
        sql &= "'" & NowDateTime & "' ,"
        sql &= "'" & Request.QueryString("pUserID") & "' ,"
        sql &= "'" & NowDateTime & "' )"
        uDataBase.ExecuteNonQuery(sql)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BGetYColor)
    '**     取得YKK色號
    '**
    '*****************************************************************
    Protected Sub BGetYColor_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGetYColor.Click
        '布帶
        DTAYCOLNO.Text = GetColorNo(DTACOL.SelectedValue, DTACOLNO.Text)
        DTALYCOLNO.Text = GetColorNo(DTALCOL.SelectedValue, DTALCOLNO.Text)
        DTARYCOLNO.Text = GetColorNo(DTARCOL.SelectedValue, DTARCOLNO.Text)
        '縫工上線
        DTHUPYCOLNO.Text = GetColorNo(DTHUPCOL.SelectedValue, DTHUPCOLNO.Text)
        DTHLUPYCOLNO.Text = GetColorNo(DTHLUPCOL.SelectedValue, DTHLUPCOLNO.Text)
        DTHLRUPYCOLNO.Text = GetColorNo(DTHLRUPCOL.SelectedValue, DTHLRUPCOLNO.Text)
        DTHRUPYCOLNO.Text = GetColorNo(DTHRUPCOL.SelectedValue, DTHRUPCOLNO.Text)
        DTHRRUPYCOLNO.Text = GetColorNo(DTHRRUPCOL.SelectedValue, DTHRRUPCOLNO.Text)
        '縫工下線
        DTHLOYCOLNO.Text = GetColorNo(DTHLOCOL.SelectedValue, DTHLOCOLNO.Text)
        DTHLLOYCOLNO.Text = GetColorNo(DTHLLOCOL.SelectedValue, DTHLLOCOLNO.Text)
        DTHLRLOYCOLNO.Text = GetColorNo(DTHLRLOCOL.SelectedValue, DTHLRLOCOLNO.Text)
        DTHRLOYCOLNO.Text = GetColorNo(DTHRLOCOL.SelectedValue, DTHRLOCOLNO.Text)
        DTHRRLOYCOLNO.Text = GetColorNo(DTHRRLOCOL.SelectedValue, DTHRRLOCOLNO.Text)
        '
        DXMYCOLNO.Text = GetColorNo(DXMCOL.SelectedValue, DXMCOLNO.Text)
        DAMYCOLNO.Text = GetColorNo(DAMCOL.SelectedValue, DAMCOLNO.Text)
        DBMYCOLNO.Text = GetColorNo(DBMCOL.SelectedValue, DBMCOLNO.Text)
        DCMYCOLNO.Text = GetColorNo(DCMCOL.SelectedValue, DCMCOLNO.Text)
        DDMYCOLNO.Text = GetColorNo(DDMCOL.SelectedValue, DDMCOLNO.Text)
        DEMYCOLNO.Text = GetColorNo(DEMCOL.SelectedValue, DEMCOLNO.Text)
        DFMYCOLNO.Text = GetColorNo(DFMCOL.SelectedValue, DFMCOLNO.Text)
        DGMYCOLNO.Text = GetColorNo(DGMCOL.SelectedValue, DGMCOLNO.Text)
        DHMYCOLNO.Text = GetColorNo(DHMCOL.SelectedValue, DHMCOLNO.Text)
        DLYYCOLNO.Text = GetColorNo(DLYCOL.SelectedValue, DLYCOLNO.Text)
        '顯示訊息
        If DMAPREFFILE.Visible Then
            If DMAPREFFILE.PostedFile.FileName <> "" Then
                uJavaScript.PopMsg(Me, "[注意]:草圖需重新指定,請確認!")
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetLevel)
    '**     取得難易度
    '**
    '*****************************************************************
    Public Function GetLevel(ByVal pLevel As String, ByVal pOP As String) As String
        Dim RtnString As String = "0"
        Dim OPKey As String = "OPLEVEL-" & pOP
        Dim SQL As String
        Try
            SQL = "Select DATA From M_REFERP "
            SQL = SQL & " Where CAT = '2002' "
            SQL = SQL & "   And DKEY = '" & OPKey & "' "
            Dim dtREFERP As DataTable = uDataBase.GetDataTable(SQL)
            If dtREFERP.Rows.Count > 0 Then
                RtnString = "0_" + dtREFERP.Rows(0).Item("DATA").ToString
            End If
        Catch ex As Exception
            RtnString = "0"
        End Try
        '
        Return RtnString
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(UpdateLoading)
    '**     更新M_Flow負荷考量
    '**
    '*****************************************************************
    Protected Sub UpdateLoading(ByVal pStep As Integer, ByVal pOP As String)
        Dim OPKey As String = "OPLOAD-" & pOP
        Dim wLoading As String = "0"
        Dim SQL As String
        Try
            '取得負荷考慮
            SQL = "Select DATA From M_REFERP "
            SQL = SQL & " Where CAT = '2002' "
            SQL = SQL & "   And DKEY = '" & OPKey & "' "
            Dim dtREFERP As DataTable = uDataBase.GetDataTable(SQL)
            If dtREFERP.Rows.Count > 0 Then
                wLoading = dtREFERP.Rows(0).Item("DATA").ToString
                '
                '更新M_Flow
                SQL = "Update M_Flow Set "
                SQL = SQL + " Loading = '" & wLoading & "',"
                SQL = SQL + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
                SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
                SQL = SQL + " Where FormNo = '" & "002002" & "'"
                SQL = SQL + "   And Step   = '" & CStr(pStep) & "'"
                uDataBase.ExecuteNonQuery(SQL)
            End If
        Catch ex As Exception
        End Try
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(UpdateManufInf)
    '**     更新試作工程資訊
    '**
    '*****************************************************************
    Protected Sub UpdateManufInf(ByVal pStep As Integer)
        Dim ListItem1 As New ListItem
        ListItem1.Text = ""
        ListItem1.Value = ""
        ListItem1.Selected = True
        '更新
        If pStep <= 40 Then
            DOP1BTIME.Text = ""
            DOP1BHOURS.Text = "0"
            DOP1ATIME.Text = ""
            DOP1AHOURS.Text = "0"
            DOP1DELAYC1.Items.Clear()
            DOP1DELAYC1.Items.Add(ListItem1)
            DOP1DELAYC2.Items.Clear()
            DOP1DELAYC2.Items.Add(ListItem1)
            DOP1REM.Text = ""
        End If
        '
        If pStep <= 50 Then
            DOP2BTIME.Text = ""
            DOP2BHOURS.Text = "0"
            DOP2ATIME.Text = ""
            DOP2AHOURS.Text = "0"
            DOP2DELAYC1.Items.Clear()
            DOP2DELAYC1.Items.Add(ListItem1)
            DOP2DELAYC2.Items.Clear()
            DOP2DELAYC2.Items.Add(ListItem1)
            DOP2REM.Text = ""
        End If
        '
        If pStep <= 60 Then
            DOP3BTIME.Text = ""
            DOP3BHOURS.Text = "0"
            DOP3ATIME.Text = ""
            DOP3AHOURS.Text = "0"
            DOP3DELAYC1.Items.Clear()
            DOP3DELAYC1.Items.Add(ListItem1)
            DOP3DELAYC2.Items.Clear()
            DOP3DELAYC2.Items.Add(ListItem1)
            DOP3REM.Text = ""
        End If
        '
        If pStep <= 70 Then
            DOP4BTIME.Text = ""
            DOP4BHOURS.Text = "0"
            DOP4ATIME.Text = ""
            DOP4AHOURS.Text = "0"
            DOP4DELAYC1.Items.Clear()
            DOP4DELAYC1.Items.Add(ListItem1)
            DOP4DELAYC2.Items.Clear()
            DOP4DELAYC2.Items.Add(ListItem1)
            DOP4REM.Text = ""
        End If
        '
        If pStep <= 80 Then
            DOP5BTIME.Text = ""
            DOP5BHOURS.Text = "0"
            DOP5ATIME.Text = ""
            DOP5AHOURS.Text = "0"
            DOP5DELAYC1.Items.Clear()
            DOP5DELAYC1.Items.Add(ListItem1)
            DOP5DELAYC2.Items.Clear()
            DOP5DELAYC2.Items.Add(ListItem1)
            DOP5REM.Text = ""
        End If
        '
        If pStep <= 90 Then
            DOP6BTIME.Text = ""
            DOP6BHOURS.Text = "0"
            DOP6ATIME.Text = ""
            DOP6AHOURS.Text = "0"
            DOP6DELAYC1.Items.Clear()
            DOP6DELAYC1.Items.Add(ListItem1)
            DOP6DELAYC2.Items.Clear()
            DOP6DELAYC2.Items.Add(ListItem1)
            DOP6REM.Text = ""
        End If
        '
        If pStep <= 100 Then
            DOP7BTIME.Text = ""
            DOP7BHOURS.Text = "0"
            DOP7ATIME.Text = ""
            DOP7AHOURS.Text = "0"
            DOP7DELAYC1.Items.Clear()
            DOP7DELAYC1.Items.Add(ListItem1)
            DOP7DELAYC2.Items.Clear()
            DOP7DELAYC2.Items.Add(ListItem1)
            DOP7REM.Text = ""
        End If
        '
        If pStep <= 110 Then
            DOP8BTIME.Text = ""
            DOP8BHOURS.Text = "0"
            DOP8ATIME.Text = ""
            DOP8AHOURS.Text = "0"
            DOP8DELAYC1.Items.Clear()
            DOP8DELAYC1.Items.Add(ListItem1)
            DOP8DELAYC2.Items.Clear()
            DOP8DELAYC2.Items.Add(ListItem1)
            DOP8REM.Text = ""
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetColorNo)
    '**     取得YKK色號
    '**
    '*****************************************************************
    Public Function GetColorNo(ByVal pSuppiler As String, ByVal pColorNo As String) As String
        Dim RtnString As String = ""
        Dim SQL As String
        Try
            If pSuppiler = "東隆" Or pSuppiler = "昭安" Then
                SQL = "Select YCOLNO From M_ColorRelation "
                SQL = SQL & " Where STS = '0' "
                SQL = SQL & "   And SUP = '" & pSuppiler & "' "
                SQL = SQL & "   And COLNO =  '" & pColorNo & "' "
                Dim dtColorRelation As DataTable = uDataBase.GetDataTable(SQL)
                If dtColorRelation.Rows.Count > 0 Then
                    RtnString = dtColorRelation.Rows(0).Item("YCOLNO").ToString
                End If
            End If
        Catch ex As Exception
            RtnString = ""
        End Try
        '
        Return RtnString
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BCREATE)
    '**     新增開發見本
    '**
    '*****************************************************************
    Protected Sub BCREATE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BCREATE.Click
        Dim Path As String = Server.MapPath(uCommon.GetAppSetting("UploadFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                     CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim sql As String
        '設定開發見本欄位
        D3DATE.Text = CDate(NowDateTime).ToString("yyyy/MM/dd")
        D3APPBUYER.Text = DAPPBUYER.Text
        If DCUSTITEM.Text <> "" Then D3APPBUYER.Text = D3APPBUYER.Text + " (" + DCUSTITEM.Text + ")"
        D3SIZENO.Text = DSIZENO.Text
        D3ITEM.Text = DITEM.Text
        D3CODENO.Text = DCODENO.Text
        D3TAWIDTH.Text = DTAWIDTH.Text
        D3DEVNO.Text = DDEVNO.Text
        D3DEVPRD.Text = DAPPDATE.Text + "~" + D3DATE.Text

        If InStr(DTASPEC.Text, "],[") <= 0 Then
            D3TACOL.Text = DTASPEC.Text
            D3TALINE.Text = ""
        Else
            D3TACOL.Text = Mid(DTASPEC.Text, 1, InStr(DTASPEC.Text, "],["))
            D3TALINE.Text = Mid(DTASPEC.Text, InStr(DTASPEC.Text, "],[") + 2, Len(DTASPEC.Text))
        End If

        D3ECOL.Text = DECOLSEL.Text + " " + DECOL.Text
        D3CCOL.Text = DCCOLSEL.Text + " " + DCCOL.Text
        D3THCOL.Text = DTHSPEC.Text
        D3TNLITEM.Text = D4TNLITEM1.Text
        D3TNRITEM.Text = D4TNRITEM1.Text
        D3TSLITEM.Text = D4TSLITEM.Text
        D3TSRITEM.Text = D4TSRITEM.Text
        D3TDLITEM.Text = D4TDLITEM.Text
        D3TDRITEM.Text = D4TDRITEM.Text
        D3CNITEM.Text = D4CNITEM.Text
        D3CSITEM.Text = D4CSITEM.Text
        D3CDITEM.Text = D4CDITEM.Text
        D3CITEM.Text = D4CITEM.Text
        D31Other.Text = D4OTHER1.Text
        D32Other.Text = D4OTHER2.Text
        D3O1ITEM.Text = D4O1ITEM.Text
        D3O2ITEM.Text = D4O2ITEM.Text
        '新增開發見本
        sql = "Select * From FS_SampleSheet "
        sql &= " Where NO =  '" & DNO.Text & "'"
        Dim dtSampleSheet As DataTable = uDataBase.GetDataTable(sql)
        If dtSampleSheet.Rows.Count <= 0 Then
            sql = "Insert into FS_SampleSheet "
            sql = sql + "( "
            sql = sql + "Date, AppBuyer, SizeNo, Item, CodeNo, "
            sql = sql + "SampleFile, TAWidth, DevNo, DevPrd, TACol, "
            sql = sql + "TALine, ECol, CCol, THCol, Other, "
            sql = sql + "QCFile1,QCFile2,QCFile3,QCFile4,QCFile5,TNLItem, TNRItem, TSLItem, TSRItem, "
            sql = sql + "TDLItem, TDRItem, CNITem, CSItem, CDItem, "
            sql = sql + "CItem, "
            sql = sql + "WF1, WF2, WF3, WF4, WF5 , WF6, WF7, "
            sql = sql + "WF3Name, WF4Name, WF5Name, WF6Name, WF7Name, "
            sql = sql + "No, Other1, Other2, O1Item, O2Item, "
            sql = sql + "CreateUser, CreateTime, ModifyUser, ModifyTime "
            sql = sql + ")  "
            sql = sql + "VALUES( "
            sql = sql + " '" + D3DATE.Text + "', "                       '發行日期
            sql = sql + " '" + D3APPBUYER.Text + "', "                   '客戶
            sql = sql + " '" + D3SIZENO.Text + "', "                     'Size
            sql = sql + " '" + D3ITEM.Text + "', "                       'Item
            sql = sql + " '" + D3CODENO.Text + "', "                     'Code-No
            '樣品檔
            Dim FileName As String = ""
            If D3SAMPLEFILE.Visible Then
                If D3SAMPLEFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    FileName = CStr(wFormSno) & "-" & "SampleFile" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(D3SAMPLEFILE.PostedFile.FileName)
                    Try    '上傳圖檔
                        D3SAMPLEFILE.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                End If
            Else
                FileName = ""
            End If
            sql = sql + " '" + FileName + "', "                         'Sample-File
            sql = sql + " '" + D3TAWIDTH.Text + "', "                    '布帶寬度
            sql = sql + " '" + D3DEVNO.Text + "', "                      '開發No
            sql = sql + " '" + D3DEVPRD.Text + "', "                     '開發期間
            sql = sql + " '" + D3TACOL.Text + "', "                      '布帶
            sql = sql + " '" + D3TALINE.Text + "', "                     '條紋線
            sql = sql + " '" + D3ECOL.Text + "', "                       '務齒
            sql = sql + " '" + D3CCOL.Text + "', "                       '丸紐
            sql = sql + " '" + D3THCOL.Text + "', "                      '縫工線
            sql = sql + " N'" + D3OTHER.Text + "', "                     '其他
            '品測檔案1
            FileName = ""
            If D3QCFILE1.Visible Then
                If D3QCFILE1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    FileName = CStr(wFormSno) & "-" & "QCFile1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(D3QCFILE1.PostedFile.FileName)
                    Try    '上傳圖檔
                        D3QCFILE1.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                End If
            Else
                FileName = ""
            End If
            sql = sql + " '" + FileName + "', "
            '品測檔案2
            FileName = ""
            If D3QCFILE2.Visible Then
                If D3QCFILE2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    FileName = CStr(wFormSno) & "-" & "QCFile2" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(D3QCFILE2.PostedFile.FileName)
                    Try    '上傳圖檔
                        D3QCFILE2.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                End If
            Else
                FileName = ""
            End If
            sql = sql + " '" + FileName + "', "
            '品測檔案3
            FileName = ""
            If D3QCFILE3.Visible Then
                If D3QCFILE3.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    FileName = CStr(wFormSno) & "-" & "QCFile3" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(D3QCFILE3.PostedFile.FileName)
                    Try    '上傳圖檔
                        D3QCFILE3.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                End If
            Else
                FileName = ""
            End If
            sql = sql + " '" + FileName + "', "
            '品測檔案4
            FileName = ""
            If D3QCFILE4.Visible Then
                If D3QCFILE4.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    FileName = CStr(wFormSno) & "-" & "QCFile4" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(D3QCFILE4.PostedFile.FileName)
                    Try    '上傳圖檔
                        D3QCFILE4.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                End If
            Else
                FileName = ""
            End If
            sql = sql + " '" + FileName + "', "
            '品測檔案5
            FileName = ""
            If D3QCFILE5.Visible Then
                If D3QCFILE5.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    FileName = CStr(wFormSno) & "-" & "QCFile5" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(D3QCFILE5.PostedFile.FileName)
                    Try    '上傳圖檔
                        D3QCFILE5.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                End If
            Else
                FileName = ""
            End If
            sql = sql + " '" + FileName + "', "
            sql = sql + " '" + D3TNLITEM.Text + "', "                    '
            sql = sql + " '" + D3TNRITEM.Text + "', "                    '
            sql = sql + " '" + D3TSLITEM.Text + "', "                    '
            sql = sql + " '" + D3TSRITEM.Text + "', "                    '
            sql = sql + " '" + D3TDLITEM.Text + "', "                    '
            sql = sql + " '" + D3TDRITEM.Text + "', "                    '
            sql = sql + " '" + D3CNITEM.Text + "', "                     '
            sql = sql + " '" + D3CSITEM.Text + "', "                     '
            sql = sql + " '" + D3CDITEM.Text + "', "                     '
            sql = sql + " '" + D3CITEM.Text + "', "                      '
            sql = sql + " N'" + D3WF1.SelectedValue + "', "               '
            sql = sql + " N'" + D3WF2.SelectedValue + "', "               '
            sql = sql + " N'" + D3WF3.SelectedValue + "', "               '
            sql = sql + " N'" + D3WF4.SelectedValue + "', "               '
            sql = sql + " N'" + D3WF5.SelectedValue + "', "               '
            sql = sql + " N'" + D3WF6.SelectedValue + "', "               '
            sql = sql + " N'" + D3WF7.SelectedValue + "', "               '
            sql = sql + " '" + D3WF3Name.SelectedValue + "', "           '
            sql = sql + " '" + D3WF4Name.SelectedValue + "', "           '
            sql = sql + " '" + D3WF5Name.SelectedValue + "', "           '
            sql = sql + " '" + D3WF6Name.SelectedValue + "', "           '
            sql = sql + " '" + D3WF7Name.SelectedValue + "', "           '
            sql = sql + " '" + DNO.Text + "', "                    '
            sql = sql + " '" + D31Other.Text + "', "                    'Other1
            sql = sql + " '" + D32Other.Text + "', "                    'Other2
            sql = sql + " '" + D3O1ITEM.Text + "', "                    'Item1
            sql = sql + " '" + D3O2ITEM.Text + "', "                    'Item2
            sql = sql + " '" + Request.QueryString("pUserID") + "', "  '作成者
            sql = sql + " '" + NowDateTime + "', "                      '作成時間
            sql = sql + " '" + "" + "', "                               '修改者
            sql = sql + " '" + NowDateTime + "' "                       '修改時間
            sql = sql + " ) "
            uDataBase.ExecuteNonQuery(sql)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim Path As String = Server.MapPath(uCommon.GetAppSetting("UploadFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                     CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim sql As String
        '---------------------------------------------------------------------------------
        '-- 開發委託
        '---------------------------------------------------------------------------------
        sql = "Update F_CommissionSheet Set "
        If pFun <> "SAVE" Then
            sql &= " Sts = '" & pSts & "',"
            sql &= " CompletedTime = '" & NowDateTime & "',"
        End If
        '--基本
        sql &= " NO = '" & DNO.Text & "' ,"
        sql &= " REFNO = '" & DREFNO.Text & "' ,"
        sql &= " APPDATE = '" & DAPPDATE.Text & "' ,"
        sql &= " APPDEPT = '" & DAPPDEPT.Text & "' ,"
        sql &= " APPPER = '" & DAPPPER.Text & "' ,"
        sql &= " APPBUYER = '" & DAPPBUYER.SelectedValue & "' ,"
        sql &= " SellVendor = '" & DSellVendor.Text & "' ,"
        sql &= " ESYQTY = '" & DESYQTY.Text & "' ,"
        sql &= " EXPDEL = '" & DEXPDEL.Text & "' ,"
        sql &= " CUSTITEM = '" & DCUSTITEM.Text & "' ,"
        sql &= " USAGE = '" & DUSAGE.SelectedValue & "' ,"
        sql &= " ORNO = '" & DORNO.Text & "' ,"
        sql &= " NEEDMAP = '" & DNEEDMAP.SelectedValue & "' ,"
        If DMAPREFFILE.Visible Then
            If DMAPREFFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "MAPREFFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DMAPREFFILE.PostedFile.FileName)
                Try    '上傳圖檔
                    DMAPREFFILE.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " MAPREFFILE = '" + FileName + "', "
            End If
        End If
        '--樣品
        If DNeedSample.Checked = True Then
            sql &= " NeedSample = '" & "1" & "' ,"
        Else
            sql &= " NeedSample = '" & "0" & "' ,"
        End If
        If DNeedItemRegister.Checked = True Then
            sql &= " NeedItemRegister = '" & "1" & "' ,"
        Else
            sql &= " NeedItemRegister = '" & "0" & "' ,"
        End If
        sql &= " PRO = '" & DPRO.SelectedValue & "' ,"
        sql &= " OPPART = '" & DOPPART.Text & "' ,"
        sql &= " PLEN = '" & DPLEN.Text & "' ,"
        sql &= " PLENUN = '" & DPLENUN.SelectedValue & "' ,"
        sql &= " PQTY = '" & DPQTY.Text & "' ,"
        sql &= " PQTYUN = '" & DPQTYUN.SelectedValue & "' ,"
        sql &= " EALEN = '" & DEALEN.Text & "' ,"
        sql &= " EALENUN = '" & DEALENUN.SelectedValue & "' ,"
        sql &= " EAQTY = '" & DEAQTY.Text & "' ,"
        sql &= " EAQTYUN = '" & DEAQTYUN.SelectedValue & "' ,"
        sql &= " UPSLI = '" & DUPSLI.Text & "' ,"
        sql &= " LOSLI = '" & DLOSLI.Text & "' ,"
        sql &= " UPFIN = '" & DUPFIN.Text & "' ,"
        sql &= " LOFIN = '" & DLOFIN.Text & "' ,"
        sql &= " UPSTK = '" & DUPSTK.Text & "' ,"
        sql &= " LOSTK = '" & DLOSTK.Text & "' ,"
        sql &= " SPSPEC = '" & DSPSPEC.Text & "' ,"
        '--開發
        sql &= " SIZENO = '" & DSIZENO.SelectedValue & "' ,"
        sql &= " ITEM = '" & DITEM.SelectedValue & "' ,"
        sql &= " TATYPE = '" & DTATYPE.SelectedValue & "' ,"
        sql &= " TAWIDTH = '" & DTAWIDTH.Text & "' ,"
        sql &= " ECOLSEL = '" & DECOLSEL.SelectedValue & "' ,"
        sql &= " ECOL = '" & DECOL.Text & "' ,"
        sql &= " CCOLSEL = '" & DCCOLSEL.SelectedValue & "' ,"
        sql &= " CCOL = '" & DCCOL.Text & "' ,"
        sql &= " TACOL = '" & DTACOL.SelectedValue & "' ,"
        sql &= " TACOLNO = '" & DTACOLNO.Text & "' ,"
        sql &= " TAYCOLNO = '" & DTAYCOLNO.Text & "' ,"
        sql &= " TALCOL = '" & DTALCOL.SelectedValue & "' ,"
        sql &= " TALCOLNO = '" & DTALCOLNO.Text & "' ,"
        sql &= " TALYCOLNO = '" & DTALYCOLNO.Text & "' ,"
        sql &= " TARCOL = '" & DTARCOL.SelectedValue & "' ,"
        sql &= " TARCOLNO = '" & DTARCOLNO.Text & "' ,"
        sql &= " TARYCOLNO = '" & DTARYCOLNO.Text & "' ,"
        sql &= " THUPCOL = '" & DTHUPCOL.SelectedValue & "' ,"
        sql &= " THUPCOLNO = '" & DTHUPCOLNO.Text & "' ,"
        sql &= " THUPYCOLNO = '" & DTHUPYCOLNO.Text & "' ,"
        sql &= " THLUPCOL = '" & DTHLUPCOL.SelectedValue & "' ,"
        sql &= " THLUPCOLNO = '" & DTHLUPCOLNO.Text & "' ,"
        sql &= " THLUPYCOLNO = '" & DTHLUPYCOLNO.Text & "' ,"
        sql &= " THLRUPCOL = '" & DTHLRUPCOL.SelectedValue & "' ,"
        sql &= " THLRUPCOLNO = '" & DTHLRUPCOLNO.Text & "' ,"
        sql &= " THLRUPYCOLNO = '" & DTHLRUPYCOLNO.Text & "' ,"
        sql &= " THRUPCOL = '" & DTHRUPCOL.SelectedValue & "' ,"
        sql &= " THRUPCOLNO = '" & DTHRUPCOLNO.Text & "' ,"
        sql &= " THRUPYCOLNO = '" & DTHRUPYCOLNO.Text & "' ,"
        sql &= " THRRUPCOL = '" & DTHRRUPCOL.SelectedValue & "' ,"
        sql &= " THRRUPCOLNO = '" & DTHRRUPCOLNO.Text & "' ,"
        sql &= " THRRUPYCOLNO = '" & DTHRRUPYCOLNO.Text & "' ,"
        sql &= " THLOCOL = '" & DTHLOCOL.SelectedValue & "' ,"
        sql &= " THLOCOLNO = '" & DTHLOCOLNO.Text & "' ,"
        sql &= " THLOYCOLNO = '" & DTHLOYCOLNO.Text & "' ,"
        sql &= " THLLOCOL = '" & DTHLLOCOL.SelectedValue & "' ,"
        sql &= " THLLOCOLNO = '" & DTHLLOCOLNO.Text & "' ,"
        sql &= " THLLOYCOLNO = '" & DTHLLOYCOLNO.Text & "' ,"
        sql &= " THLRLOCOL = '" & DTHLRLOCOL.SelectedValue & "' ,"
        sql &= " THLRLOCOLNO = '" & DTHLRLOCOLNO.Text & "' ,"
        sql &= " THLRLOYCOLNO = '" & DTHLRLOYCOLNO.Text & "' ,"
        sql &= " THRLOCOL = '" & DTHRLOCOL.SelectedValue & "' ,"
        sql &= " THRLOCOLNO = '" & DTHRLOCOLNO.Text & "' ,"
        sql &= " THRLOYCOLNO = '" & DTHRLOYCOLNO.Text & "' ,"
        sql &= " THRRLOCOL = '" & DTHRRLOCOL.SelectedValue & "' ,"
        sql &= " THRRLOCOLNO = '" & DTHRRLOCOLNO.Text & "' ,"
        sql &= " THRRLOYCOLNO = '" & DTHRRLOYCOLNO.Text & "' ,"
        sql &= " XMLEN = '" & DXMLEN.Text & "' ,"
        sql &= " XMCOL = '" & DXMCOL.SelectedValue & "' ,"
        sql &= " XMCOLNO = '" & DXMCOLNO.Text & "' ,"
        sql &= " XMYCOLNO = '" & DXMYCOLNO.Text & "' ,"
        sql &= " AMLEN = '" & DAMLEN.Text & "' ,"
        sql &= " AMCOL = '" & DAMCOL.SelectedValue & "' ,"
        sql &= " AMCOLNO = '" & DAMCOLNO.Text & "' ,"
        sql &= " AMYCOLNO = '" & DAMYCOLNO.Text & "' ,"
        sql &= " BMLEN = '" & DBMLEN.Text & "' ,"
        sql &= " BMCOL = '" & DBMCOL.SelectedValue & "' ,"
        sql &= " BMCOLNO = '" & DBMCOLNO.Text & "' ,"
        sql &= " BMYCOLNO = '" & DBMYCOLNO.Text & "' ,"
        sql &= " CMLEN = '" & DCMLEN.Text & "' ,"
        sql &= " CMCOL = '" & DCMCOL.SelectedValue & "' ,"
        sql &= " CMCOLNO = '" & DCMCOLNO.Text & "' ,"
        sql &= " CMYCOLNO = '" & DCMYCOLNO.Text & "' ,"
        sql &= " DMLEN = '" & DDMLEN.Text & "' ,"
        sql &= " DMCOL = '" & DDMCOL.SelectedValue & "' ,"
        sql &= " DMCOLNO = '" & DDMCOLNO.Text & "' ,"
        sql &= " DMYCOLNO = '" & DDMYCOLNO.Text & "' ,"
        sql &= " EMLEN = '" & DEMLEN.Text & "' ,"
        sql &= " EMCOL = '" & DEMCOL.SelectedValue & "' ,"
        sql &= " EMCOLNO = '" & DEMCOLNO.Text & "' ,"
        sql &= " EMYCOLNO = '" & DEMYCOLNO.Text & "' ,"
        sql &= " FMLEN = '" & DFMLEN.Text & "' ,"
        sql &= " FMCOL = '" & DFMCOL.SelectedValue & "' ,"
        sql &= " FMCOLNO = '" & DFMCOLNO.Text & "' ,"
        sql &= " FMYCOLNO = '" & DFMYCOLNO.Text & "' ,"
        sql &= " GMLEN = '" & DGMLEN.Text & "' ,"
        sql &= " GMCOL = '" & DGMCOL.SelectedValue & "' ,"
        sql &= " GMCOLNO = '" & DGMCOLNO.Text & "' ,"
        sql &= " GMYCOLNO = '" & DGMYCOLNO.Text & "' ,"
        sql &= " HMLEN = '" & DHMLEN.Text & "' ,"
        sql &= " HMCOL = '" & DHMCOL.SelectedValue & "' ,"
        sql &= " HMCOLNO = '" & DHMCOLNO.Text & "' ,"
        sql &= " HMYCOLNO = '" & DHMYCOLNO.Text & "' ,"
        sql &= " LYLEN = '" & DLYLEN.Text & "' ,"
        sql &= " LYCOL = '" & DLYCOL.SelectedValue & "' ,"
        sql &= " LYCOLNO = '" & DLYCOLNO.Text & "' ,"
        sql &= " LYYCOLNO = '" & DLYYCOLNO.Text & "' ,"
        sql &= " OTCON = '" & DOTCON.Text & "' ,"
        '--圖面
        sql &= " MAPNO = '" & DMAPNO.Text & "' ,"
        sql &= " MAKEMAP = '" & DMAKEMAP.SelectedValue & "' ,"
        sql &= " LEVEL = '" & DLEVEL.SelectedValue & "' ,"
        If DMAPFILE.Visible Then
            If DMAPFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "MAPFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DMAPFILE.PostedFile.FileName)
                Try    '上傳圖檔
                    DMAPFILE.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " MAPFILE = '" + FileName + "', "
            End If
        End If
        '--適用型別
        If DFORTYPEFILE.Visible Then
            If DFORTYPEFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "FORTYPEFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DFORTYPEFILE.PostedFile.FileName)
                Try    '上傳圖檔
                    DFORTYPEFILE.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " FORTYPEFILE = '" + FileName + "', "
            End If
        End If
        '--品質檢測
        If DQCCHECKFILE.Visible Then
            If DQCCHECKFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "QCCHECKFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCCHECKFILE.PostedFile.FileName)
                Try    '上傳圖檔
                    DQCCHECKFILE.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " QCCHECKFILE = '" + FileName + "', "
            End If
        End If
        '--QC檢測文件
        If DQCFILE1.Visible Then
            If DQCFILE1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "QCFILE1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE1.PostedFile.FileName)
                Try    '上傳圖檔
                    DQCFILE1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " QCFILE1 = '" + FileName + "', "
            End If
        End If
        '
        If DQCFILE2.Visible Then
            If DQCFILE2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "QCFILE2" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE2.PostedFile.FileName)
                Try    '上傳圖檔
                    DQCFILE2.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " QCFILE2 = '" + FileName + "', "
            End If
        End If
        '
        If DQCFILE3.Visible Then
            If DQCFILE3.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "QCFILE3" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE3.PostedFile.FileName)
                Try    '上傳圖檔
                    DQCFILE3.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " QCFILE3 = '" + FileName + "', "
            End If
        End If
        '
        If DQCFILE4.Visible Then
            If DQCFILE4.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "QCFILE4" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE4.PostedFile.FileName)
                Try    '上傳圖檔
                    DQCFILE4.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " QCFILE4 = '" + FileName + "', "
            End If
        End If
        '
        If DQCFILE5.Visible Then
            If DQCFILE5.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "QCFILE5" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE5.PostedFile.FileName)
                Try    '上傳圖檔
                    DQCFILE5.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " QCFILE5 = '" + FileName + "', "
            End If
        End If
        '
        If DQCFILE6.Visible Then
            If DQCFILE6.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "QCFILE6" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE6.PostedFile.FileName)
                Try    '上傳圖檔
                    DQCFILE6.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " QCFILE6 = '" + FileName + "', "
            End If
        End If

        '
        If DGenFILE1.Visible Then
            If DGenFILE1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "GenFILE1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DGenFILE1.PostedFile.FileName)
                Try    '上傳圖檔
                    DGenFILE1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " GenFILE1 = '" + FileName + "', "
            End If
        End If


        '--客戶切結書
        If DCONTACTFILE.Visible Then
            If DCONTACTFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "CONTACTFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DCONTACTFILE.PostedFile.FileName)
                Try    '上傳圖檔
                    DCONTACTFILE.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " CONTACTFILE = '" + FileName + "', "
            End If
        End If
        '--樣品確認書
        If DSAMPLECONFIRMFILE.Visible Then
            If DSAMPLECONFIRMFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "SAMPLECONFIRMFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSAMPLECONFIRMFILE.PostedFile.FileName)
                Try    '上傳圖檔
                    DSAMPLECONFIRMFILE.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " SAMPLECONFIRMFILE = '" + FileName + "', "
            End If
        End If
        '--製造授權書
        If DMANUFAUTHORITYFILE.Visible Then
            If DMANUFAUTHORITYFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "MANUFAUTHORITYFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DMANUFAUTHORITYFILE.PostedFile.FileName)
                Try    '上傳圖檔
                    DMANUFAUTHORITYFILE.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " MANUFAUTHORITYFILE = '" + FileName + "', "
            End If
        End If

        sql &= " MANUOUTPRICE = '" & DMANUOUTPRICE.Text & "' ,"
        '--報價單
        If DFORCASTFILE.Visible Then
            If DFORCASTFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "DFORCASTFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DFORCASTFILE.PostedFile.FileName)
                Try    '上傳圖檔
                    DFORCASTFILE.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " FORCASTFILE = '" + FileName + "', "
            End If
        End If
        sql &= " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        sql &= " ModifyTime = '" & NowDateTime & "' "
        sql &= " Where FormNo  =  '" & wFormNo & "'"
        sql &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        uDataBase.ExecuteNonQuery(sql)
        '---------------------------------------------------------------------------------
        '-- 製造委託
        '---------------------------------------------------------------------------------
        sql = "Select * From FS_ManufactureSheet "
        sql &= " Where NO =  '" & DNO.Text & "'"
        Dim dtManufactureSheet As DataTable = uDataBase.GetDataTable(sql)
        If dtManufactureSheet.Rows.Count > 0 Then
            sql = "Update FS_ManufactureSheet Set "
            '--示意圖
            If DHINTFILE.Visible Then
                If DHINTFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    Dim FileName As String = CStr(wFormSno) & "-" & "HINTFILE" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DHINTFILE.PostedFile.FileName)
                    Try    '上傳圖檔
                        DHINTFILE.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                    sql = sql + " HINTFILE = '" + FileName + "', "
                End If
            End If
            '--各工程日程
            sql &= " OP1BTIME = '" & DOP1BTIME.Text & "' ,"
            sql &= " OP1BHOURS = '" & DOP1BHOURS.Text & "' ,"
            sql &= " OP1ATIME = '" & DOP1ATIME.Text & "' ,"
            sql &= " OP1AHOURS = '" & DOP1AHOURS.Text & "' ,"

            sql &= " OP2BTIME = '" & DOP2BTIME.Text & "' ,"
            sql &= " OP2BHOURS = '" & DOP2BHOURS.Text & "' ,"
            sql &= " OP2ATIME = '" & DOP2ATIME.Text & "' ,"
            sql &= " OP2AHOURS = '" & DOP2AHOURS.Text & "' ,"

            sql &= " OP3BTIME = '" & DOP3BTIME.Text & "' ,"
            sql &= " OP3BHOURS = '" & DOP3BHOURS.Text & "' ,"
            sql &= " OP3ATIME = '" & DOP3ATIME.Text & "' ,"
            sql &= " OP3AHOURS = '" & DOP3AHOURS.Text & "' ,"

            sql &= " OP4BTIME = '" & DOP4BTIME.Text & "' ,"
            sql &= " OP4BHOURS = '" & DOP4BHOURS.Text & "' ,"
            sql &= " OP4ATIME = '" & DOP4ATIME.Text & "' ,"
            sql &= " OP4AHOURS = '" & DOP4AHOURS.Text & "' ,"

            sql &= " OP5BTIME = '" & DOP5BTIME.Text & "' ,"
            sql &= " OP5BHOURS = '" & DOP5BHOURS.Text & "' ,"
            sql &= " OP5ATIME = '" & DOP5ATIME.Text & "' ,"
            sql &= " OP5AHOURS = '" & DOP5AHOURS.Text & "' ,"

            sql &= " OP6BTIME = '" & DOP6BTIME.Text & "' ,"
            sql &= " OP6BHOURS = '" & DOP6BHOURS.Text & "' ,"
            sql &= " OP6ATIME = '" & DOP6ATIME.Text & "' ,"
            sql &= " OP6AHOURS = '" & DOP6AHOURS.Text & "' ,"

            sql &= " OP7BTIME = '" & DOP7BTIME.Text & "' ,"
            sql &= " OP7BHOURS = '" & DOP7BHOURS.Text & "' ,"
            sql &= " OP7ATIME = '" & DOP7ATIME.Text & "' ,"
            sql &= " OP7AHOURS = '" & DOP7AHOURS.Text & "' ,"

            sql &= " OP8BTIME = '" & DOP8BTIME.Text & "' ,"
            sql &= " OP8BHOURS = '" & DOP8BHOURS.Text & "' ,"
            sql &= " OP8ATIME = '" & DOP8ATIME.Text & "' ,"
            sql &= " OP8AHOURS = '" & DOP8AHOURS.Text & "' ,"
            '--各工程遲納
            sql &= " OP1DELAYC1 = '" & DOP1DELAYC1.SelectedValue & "' ,"
            sql &= " OP1DELAYC2 = '" & DOP1DELAYC2.SelectedValue & "' ,"
            sql &= " OP1REM = '" & DOP1REM.Text & "' ,"
            sql &= " OP2DELAYC1 = '" & DOP2DELAYC1.SelectedValue & "' ,"
            sql &= " OP2DELAYC2 = '" & DOP2DELAYC2.SelectedValue & "' ,"
            sql &= " OP2REM = '" & DOP2REM.Text & "' ,"
            sql &= " OP3DELAYC1 = '" & DOP3DELAYC1.SelectedValue & "' ,"
            sql &= " OP3DELAYC2 = '" & DOP3DELAYC2.SelectedValue & "' ,"
            sql &= " OP3REM = '" & DOP3REM.Text & "' ,"
            sql &= " OP4DELAYC1 = '" & DOP4DELAYC1.SelectedValue & "' ,"
            sql &= " OP4DELAYC2 = '" & DOP4DELAYC2.SelectedValue & "' ,"
            sql &= " OP4REM = '" & DOP4REM.Text & "' ,"
            sql &= " OP5DELAYC1 = '" & DOP5DELAYC1.SelectedValue & "' ,"
            sql &= " OP5DELAYC2 = '" & DOP5DELAYC2.SelectedValue & "' ,"
            sql &= " OP5REM = '" & DOP5REM.Text & "' ,"
            sql &= " OP6DELAYC1 = '" & DOP6DELAYC1.SelectedValue & "' ,"
            sql &= " OP6DELAYC2 = '" & DOP6DELAYC2.SelectedValue & "' ,"
            sql &= " OP6REM = '" & DOP6REM.Text & "' ,"
            sql &= " OP7DELAYC1 = '" & DOP7DELAYC1.SelectedValue & "' ,"
            sql &= " OP7DELAYC2 = '" & DOP7DELAYC2.SelectedValue & "' ,"
            sql &= " OP7REM = '" & DOP7REM.Text & "' ,"
            sql &= " OP8DELAYC1 = '" & DOP8DELAYC1.SelectedValue & "' ,"
            sql &= " OP8DELAYC2 = '" & DOP8DELAYC2.SelectedValue & "' ,"
            sql &= " OP8REM = '" & DOP8REM.Text & "' ,"
            sql &= " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            sql &= " ModifyTime = '" & NowDateTime & "' "
            sql &= " Where NO =  '" & DNO.Text & "'"
            uDataBase.ExecuteNonQuery(sql)
        End If
        '---------------------------------------------------------------------------------
        '-- 開發見本
        '---------------------------------------------------------------------------------
        sql = "Select * From FS_SampleSheet "
        sql &= " Where NO =  '" & DNO.Text & "'"
        Dim dtSampleSheet As DataTable = uDataBase.GetDataTable(sql)
        If dtSampleSheet.Rows.Count > 0 Then
            sql = "Update FS_SampleSheet Set "
            sql &= " DATE = '" & D3DATE.Text & "' ,"
            sql &= " APPBUYER = '" & D3APPBUYER.Text & "' ,"
            sql &= " SIZENO = '" & D3SIZENO.Text & "' ,"
            sql &= " ITEM = '" & D3ITEM.Text & "' ,"
            sql &= " CODENO = '" & D3CODENO.Text & "' ,"
            '樣品檔
            If D3SAMPLEFILE.Visible Then
                If D3SAMPLEFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    Dim FileName As String = CStr(wFormSno) & "-" & "SampleFile" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(D3SAMPLEFILE.PostedFile.FileName)
                    Try    '上傳圖檔
                        D3SAMPLEFILE.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                    sql = sql + " SAMPLEFILE = '" + FileName + "', "
                End If
            End If
            sql &= " TAWIDTH = '" & D3TAWIDTH.Text & "' ,"
            sql &= " DEVNO = '" & D3DEVNO.Text & "' ,"
            sql &= " DEVPRD = '" & D3DEVPRD.Text & "' ,"
            sql &= " TACOL = '" & D3TACOL.Text & "' ,"
            sql &= " TALINE = '" & D3TALINE.Text & "' ,"
            sql &= " ECOL = '" & D3ECOL.Text & "' ,"
            sql &= " CCOL = '" & D3CCOL.Text & "' ,"
            sql &= " THCOL = '" & D3THCOL.Text & "' ,"
            sql &= " OTHER = N'" & D3OTHER.Text & "' ,"
            '品測檔案1
            If D3QCFILE1.Visible Then
                If D3QCFILE1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    Dim FileName As String = CStr(wFormSno) & "-" & "QCFile1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(D3QCFILE1.PostedFile.FileName)
                    Try    '上傳圖檔
                        D3QCFILE1.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                    sql = sql + " QCFile1 = '" + FileName + "', "
                End If
            End If
            '品測檔案2
            If D3QCFILE2.Visible Then
                If D3QCFILE2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    Dim FileName As String = CStr(wFormSno) & "-" & "QCFile2" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(D3QCFILE2.PostedFile.FileName)
                    Try    '上傳圖檔
                        D3QCFILE2.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                    sql = sql + " QCFile2 = '" + FileName + "', "
                End If
            End If
            '品測檔案3
            If D3QCFILE3.Visible Then
                If D3QCFILE3.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    Dim FileName As String = CStr(wFormSno) & "-" & "QCFile3" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(D3QCFILE3.PostedFile.FileName)
                    Try    '上傳圖檔
                        D3QCFILE3.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                    sql = sql + " QCFile3 = '" + FileName + "', "
                End If
            End If
            '品測檔案4
            If D3QCFILE4.Visible Then
                If D3QCFILE4.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    Dim FileName As String = CStr(wFormSno) & "-" & "QCFile4" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(D3QCFILE4.PostedFile.FileName)
                    Try    '上傳圖檔
                        D3QCFILE4.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                    sql = sql + " QCFile4 = '" + FileName + "', "
                End If
            End If
            '品測檔案5
            If D3QCFILE5.Visible Then
                If D3QCFILE5.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    Dim FileName As String = CStr(wFormSno) & "-" & "QCFile5" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(D3QCFILE5.PostedFile.FileName)
                    Try    '上傳圖檔
                        D3QCFILE5.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                    sql = sql + " QCFile5 = '" + FileName + "', "
                End If
            End If
            sql &= " TNLITEM = '" & D3TNLITEM.Text & "' ,"
            sql &= " TNRITEM = '" & D3TNRITEM.Text & "' ,"
            sql &= " TSLITEM = '" & D3TSLITEM.Text & "' ,"
            sql &= " TSRITEM = '" & D3TSRITEM.Text & "' ,"

            sql &= " TDLITEM = '" & D3TDLITEM.Text & "' ,"
            sql &= " TDRITEM = '" & D3TDRITEM.Text & "' ,"
            sql &= " CNITEM = '" & D3CNITEM.Text & "' ,"
            sql &= " CSITEM = '" & D3CSITEM.Text & "' ,"
            sql &= " CDITEM = '" & D3CDITEM.Text & "' ,"
            sql &= " CITEM = '" & D3CITEM.Text & "' ,"
            sql &= " WF1 = N'" & D3WF1.SelectedValue & "' ,"
            sql &= " WF2 = N'" & D3WF2.SelectedValue & "' ,"
            sql &= " WF3 = N'" & D3WF3.SelectedValue & "' ,"
            sql &= " WF4 = N'" & D3WF4.SelectedValue & "' ,"
            sql &= " WF5 = N'" & D3WF5.SelectedValue & "' ,"
            sql &= " WF6 = N'" & D3WF6.SelectedValue & "' ,"
            sql &= " WF7 = N'" & D3WF7.SelectedValue & "' ,"
            sql &= " WF3Name = '" & D3WF3Name.SelectedValue & "' ,"
            sql &= " WF4Name = '" & D3WF4Name.SelectedValue & "' ,"
            sql &= " WF5Name = '" & D3WF5Name.SelectedValue & "' ,"
            sql &= " WF6Name = '" & D3WF6Name.SelectedValue & "' ,"
            sql &= " WF7Name = '" & D3WF7Name.SelectedValue & "' ,"
            sql &= " Other1 = '" & D31Other.Text & "' ,"
            sql &= " Other2 = '" & D32Other.Text & "' ,"
            sql &= " O1Item = '" & D3O1ITEM.Text & "' ,"
            sql &= " O2Item = '" & D3O2ITEM.Text & "' ,"
            sql &= " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            sql &= " ModifyTime = '" & NowDateTime & "' "
            sql &= " Where NO =  '" & DNO.Text & "'"
            uDataBase.ExecuteNonQuery(sql)
        End If
        '---------------------------------------------------------------------------------
        '-- 原單位
        '---------------------------------------------------------------------------------
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyTranData)
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub ModifyTranData(ByVal pFun As String, ByVal pSts As String)
        Dim SQl As String
        Dim wStartTime As DateTime
        If pFun <> "SAVE" Then      '<> Save
            SQl = "Update T_WaitHandle Set "
            SQl = SQl + " Active = '" & "0" & "',"
            SQl = SQl + " Sts = '" & pSts & "',"
            If pSts = "1" Then SQl = SQl + " StsDes = '" & BOK.Text & "',"
            If pSts = "2" Then SQl = SQl + " StsDes = '" & BNG1.Text & "',"
            If pSts = "3" Then SQl = SQl + " StsDes = '" & BNG2.Text & "',"
            'Add-2012/3/16-Joy 160工程追加[開發中止] (因Button已滿所以使用 Save Button)
            '
            If pSts = "8" Then SQl = SQl + " StsDes = '" & BSAVE.Text & "',"
            '
            '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            'Add-Start by Joy  2012/7/31 新納期對應
            Dim RtnCode As Integer = 0
            Dim xAStartTime As String = ""
            RtnCode = oSchedule.GetActualStartTime(wFormNo, _
                                                   wFormSno, _
                                                   GetFlowLoading(wStep), _
                                                   oCommon.GetWorkID(Request.QueryString("pUserID")), _
                                                   oCommon.GetCalendarGroup(Request.QueryString("pUserID")), _
                                                   xAStartTime)
            If RtnCode = 0 Then
                SQl = SQl + " AStartTime = '" & xAStartTime & "',"
            End If
            'Add-End
            '
            SQl = SQl + " AEndTime = '" & NowDateTime & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
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
            SQl = SQl + " DecideDesc = N'" & uCommon.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step =  '" & CStr(wStep) & "'"
            SQl = SQl + "   And SeqNo =  '" & CStr(wSeqNo) & "'"
            SQl = SQl + "   And Active =  '1' "
        End If
        uDataBase.ExecuteNonQuery(SQl)
        '取得實際開始日期
        SQl = "Select AStartTime From T_WaitHandle "
        SQl = SQl + " Where FormNo  =  '" & wFormNo & "' "
        SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "' "
        SQl = SQl + "   And Step    =  '" & CStr(wStep) & "' "
        SQl = SQl + "   And SeqNo   =  '" & CStr(wSeqNo) & "' "
        SQl = SQl + "Order by Unique_ID Desc "      ' ADD-Start Joy 2012/2/2
        Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQl)
        If dt_WaitHandle.Rows.Count > 0 Then
            wStartTime = dt_WaitHandle.Rows(0).Item("AStartTime")
        End If
        '設定試作工程開始,預定時間及時數
        SetOPWorkTime("A", wStep, wStartTime, CDate(NowDateTime), wDecideCalendar)
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
                SQl = SQl + " '" + DMAPNO.Text + "', "
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
                    SQl = SQl + " MapNo = '" & DMAPNO.Text & "',"
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
    '**(GridView1_RowDataBound)
    '**     延遲處理
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim Pos As Integer = InStr(e.Row.Cells(7).Text.ToString, "],")
            '
            '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            'Modify-Start by Joy  2012/7/31 新納期對應
            '
            'Dim Str1 As String = Mid(e.Row.Cells(7).Text.ToString, 1, Pos)
            'Dim Str2 As String = Mid(e.Row.Cells(7).Text.ToString, Pos + 3, Len(e.Row.Cells(7).Text.ToString))
            'e.Row.Cells(7).Text = Str1 + "<br/>" + Str2
            If Pos > 0 Then
                Dim Str1 As String = Mid(e.Row.Cells(7).Text.ToString, 1, Pos)
                Dim Str2 As String = Mid(e.Row.Cells(7).Text.ToString, Pos + 3, Len(e.Row.Cells(7).Text.ToString))
                e.Row.Cells(7).Text = Str1 + "<br/>" + Str2
            End If
            '
            'Modify-End
            '
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(DOP1DELAYC1_SelectedIndexChanged)
    '**     OP1~OP8延遲原因2
    '**
    '*****************************************************************
    Protected Sub DOP1DELAYC1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DOP1DELAYC1.SelectedIndexChanged
        SetFieldData(1, "OP1DELAYC2", "")       'OP1-遲納原因-2
    End Sub
    Protected Sub DOP2DELAYC1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DOP2DELAYC1.SelectedIndexChanged
        SetFieldData(1, "OP2DELAYC2", "")       'OP2-遲納原因-2
    End Sub
    Protected Sub DOP3DELAYC1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DOP3DELAYC1.SelectedIndexChanged
        SetFieldData(1, "OP3DELAYC2", "")       'OP3-遲納原因-2
    End Sub
    Protected Sub DOP4DELAYC1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DOP4DELAYC1.SelectedIndexChanged
        SetFieldData(1, "OP4DELAYC2", "")       'OP4-遲納原因-2
    End Sub
    Protected Sub DOP5DELAYC1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DOP5DELAYC1.SelectedIndexChanged
        SetFieldData(1, "OP5DELAYC2", "")       'OP5-遲納原因-2
    End Sub
    Protected Sub DOP6DELAYC1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DOP6DELAYC1.SelectedIndexChanged
        SetFieldData(1, "OP6DELAYC2", "")       'OP6-遲納原因-2
    End Sub
    Protected Sub DOP7DELAYC1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DOP7DELAYC1.SelectedIndexChanged
        SetFieldData(1, "OP7DELAYC2", "")       'OP7-遲納原因-2
    End Sub
    Protected Sub DOP8DELAYC1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DOP8DELAYC1.SelectedIndexChanged
        SetFieldData(1, "OP8DELAYC2", "")       'OP8-遲納原因-2
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     檢查上傳檔案
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As FileUpload) As Integer
        Dim fileExtension As String     '宣告一個變數存放檔案格式(副檔名)
        Dim allowedExtensions As String() = {".jpg", ".jpeg", ".gif", ".pdf", ".xls", ".doc", ".ppt", ".docx", ".xlsx", ".pptx"}   '定義允許的檔案格式
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
            If UPFile.PostedFile.ContentLength <= 2000 * 1024 Then
                UPFileIsNormal = 0
            Else
                UPFileIsNormal = 9030
            End If
        End If
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
    '*****************************************************************
    '**
    '**     設定ImageButton的圖檔
    '**
    '*****************************************************************
    Protected Sub SetImageButtonImageFile(ByVal pIndex As Integer)
        MultiView1.ActiveViewIndex = pIndex
        '設定執行按鈕
        If pIndex <> CInt(DPageIdx.Text) Then
            BOK.Enabled = False
            BSAVE.Enabled = False
            BNG1.Enabled = False
            BNG2.Enabled = False
        Else
            BOK.Enabled = True
            BSAVE.Enabled = True
            BNG1.Enabled = True
            BNG2.Enabled = True
        End If
        '設定頁次影像
        If pIndex = 0 Then
            ImageButton1.ImageUrl = "Images/DevelopmentCommission_Blank.jpg"
            ImageButton2.ImageUrl = "Images/ManufactureCommission_Button.jpg"
            ImageButton3.ImageUrl = "Images/DevelopmentSample_Button.jpg"
            ImageButton4.ImageUrl = "Images/DevelopmentGentani_Button.jpg"
            ImageButton5.ImageUrl = "Images/SignHistory_Button.jpg"
        End If
        If pIndex = 1 Then
            ImageButton1.ImageUrl = "Images/DevelopmentCommission_Button.jpg"
            ImageButton2.ImageUrl = "Images/ManufactureCommission_Blank.jpg"
            ImageButton3.ImageUrl = "Images/DevelopmentSample_Button.jpg"
            ImageButton4.ImageUrl = "Images/DevelopmentGentani_Button.jpg"
            ImageButton5.ImageUrl = "Images/SignHistory_Button.jpg"
        End If
        If pIndex = 2 Then
            ImageButton1.ImageUrl = "Images/DevelopmentCommission_Button.jpg"
            ImageButton2.ImageUrl = "Images/ManufactureCommission_Button.jpg"
            ImageButton3.ImageUrl = "Images/DevelopmentSample_Blank.jpg"
            ImageButton4.ImageUrl = "Images/DevelopmentGentani_Button.jpg"
            ImageButton5.ImageUrl = "Images/SignHistory_Button.jpg"
        End If
        If pIndex = 3 Then
            ImageButton1.ImageUrl = "Images/DevelopmentCommission_Button.jpg"
            ImageButton2.ImageUrl = "Images/ManufactureCommission_Button.jpg"
            ImageButton3.ImageUrl = "Images/DevelopmentSample_Button.jpg"
            ImageButton4.ImageUrl = "Images/DevelopmentGentani_Blank.jpg"
            ImageButton5.ImageUrl = "Images/SignHistory_Button.jpg"
        End If
        If pIndex = 4 Then
            ImageButton1.ImageUrl = "Images/DevelopmentCommission_Button.jpg"
            ImageButton2.ImageUrl = "Images/ManufactureCommission_Button.jpg"
            ImageButton3.ImageUrl = "Images/DevelopmentSample_Button.jpg"
            ImageButton4.ImageUrl = "Images/DevelopmentGentani_Button.jpg"
            ImageButton5.ImageUrl = "Images/SignHistory_Blank.jpg"
        End If
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton1 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton1.Click
        '廠長用開發履歷  ADD-2011/10/15
        GridView2.Visible = True
        '
        SetImageButtonImageFile(0)
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton2 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton2.Click
        '廠長用開發履歷  ADD-2011/10/15
        GridView2.Visible = True
        '
        SetImageButtonImageFile(1)
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton3 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton3_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton3.Click
        '廠長用開發履歷  ADD-2011/10/15
        GridView2.Visible = True
        '
        SetImageButtonImageFile(2)
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton4 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton4_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton4.Click
        '廠長用開發履歷  ADD-2011/10/15
        GridView2.Visible = True
        '
        SetImageButtonImageFile(3)
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton5 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton5_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton5.Click
        '廠長用開發履歷  ADD-2011/10/15
        GridView2.Visible = False
        '
        DReadHistory.Checked = True
        SetImageButtonImageFile(4)
    End Sub
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    'Add-Start by Joy  2012/7/31 新納期對應
    '
    '*****************************************************************
    '**
    '**     動態產生Group成員(CreateGroupMember) 
    '**
    '*****************************************************************
    Protected Sub CreateGroupMember()
        Dim sql As String
        Dim i As Integer
        Dim wOPList(8) As String
        '
        For i = 1 To 8
            wOPList(i) = ""
        Next i
        '
        If DOP1PER.Text <> "" Then wOPList(1) = DOP1PER.Text
        If DOP2PER.Text <> "" Then wOPList(2) = DOP2PER.Text
        If DOP3PER.Text <> "" Then wOPList(3) = DOP3PER.Text
        If DOP4PER.Text <> "" Then wOPList(4) = DOP4PER.Text
        If DOP5PER.Text <> "" Then wOPList(5) = DOP5PER.Text
        If DOP6PER.Text <> "" Then wOPList(6) = DOP6PER.Text
        If DOP7PER.Text <> "" Then wOPList(7) = DOP7PER.Text
        If DOP8PER.Text <> "" Then wOPList(8) = DOP8PER.Text
        '
        '刪除
        Sql = "Delete From M_GROUP Where GROUPID = '002002-N5' "
        uDataBase.ExecuteNonQuery(Sql)
        '
        '增加
        For i = 1 To 8
            If wOPList(i) <> "" Then
                Sql = "Insert into M_GROUP "
                Sql &= "( "
                Sql &= "Active, GroupID, GroupName, UserID, UserName, Description, CreateUser, CreateTime, ModifyUser, ModifyTime "
                Sql &= ")  "
                Sql &= "VALUES( "
                Sql &= " '" + "1" + "', "
                Sql &= " '" + "002002-N5" + "', "
                Sql &= " '" + "試作工程擔當" + "', "
                Sql &= " '" + oCommon.GetUserID(wOPList(i)) + "', "
                Sql &= " '" + wOPList(i) + "', "
                Sql &= " '" + "" + "', "
                sql &= "'" & Request.QueryString("pUserID") & "' ,"
                sql &= "'" & NowDateTime & "' ,"
                sql &= "'" & Request.QueryString("pUserID") & "' ,"
                Sql &= "'" & NowDateTime & "'"
                Sql &= " ) "
                uDataBase.ExecuteNonQuery(Sql)
            End If
        Next i
    End Sub
    '*****************************************************************
    '**(GetFlowLoading)
    '**     取得負荷狀態
    '**
    '*****************************************************************
    Function GetFlowLoading(ByVal pStep As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim xKey As String = ""
        Dim sql As String
        '
        Select Case pStep
            Case 40
                xKey = "OPLOAD-" & HOP1.Text
            Case 50
                xKey = "OPLOAD-" & HOP2.Text
            Case 60
                xKey = "OPLOAD-" & HOP3.Text
            Case 70
                xKey = "OPLOAD-" & HOP4.Text
            Case 80
                xKey = "OPLOAD-" & HOP5.Text
            Case 90
                xKey = "OPLOAD-" & HOP6.Text
            Case 100
                xKey = "OPLOAD-" & HOP7.Text
            Case Else
                xKey = "OPLOAD-" & HOP8.Text
        End Select
        '
        sql = "Select * From M_Referp "
        sql &= "Where Cat  = '2002' "
        sql &= "  And DKey = '" & xKey & "' "
        Dim dt_Referp = uDataBase.GetDataTable(sql)
        If dt_Referp.Rows.Count > 0 Then
            RtnCode = 1
        End If
        '
        Return RtnCode
    End Function
    '*****************************************************************
    '**
    '**     展開預定工程(ExpandFlowControl) 
    '**
    '*****************************************************************
    Protected Sub ExpandFlowControl(ByVal pFormNo As String, _
                                    ByVal pFormSno As Integer, _
                                    ByVal pApplyID As String, _
                                    ByVal pUserID As String, _
                                    ByVal pStartStep As Integer, _
                                    ByVal pEndStep As Integer)
        Dim xStep As Integer = pStartStep
        Dim xAction, xNextStep As Integer
        Dim xLevel, xAllocateID As String
        '
        Dim sql As String
        sql = "Select * From T_WaitHandle "
        sql &= "Where FormNo   = '" & pFormNo & "' "
        sql &= "  And FormSno  = '" & CStr(pFormSno) & "' "
        sql &= "  And Active   = '9' "
        sql &= "  And FlowType = '9' "
        Dim dt_WaitHandle = uDataBase.GetDataTable(sql)
        If dt_WaitHandle.Rows.Count <= 0 Then
            '
            While xStep < pEndStep
                '初值
                xAction = 0
                xAllocateID = ""
                xLevel = DLEVEL.Text
                '1.取得Action
                xAction = oCommon.GetNextAction(pFormNo, xStep)
                '2.指定人員  
                '3.取得難易度
                '4.更新M_Flow負荷考量
                Select Case xStep
                    Case 20         '-->OP1
                        xAllocateID = oCommon.GetUserID(DOP1PER.Text)
                        xLevel = GetLevel(xLevel, HOP1.Text)
                        UpdateLoading(40, HOP1.Text)
                        If xAllocateID = "" Then xAction = 2
                    Case 40         '-->OP2
                        xAllocateID = oCommon.GetUserID(DOP2PER.Text)
                        xLevel = GetLevel(xLevel, HOP2.Text)
                        UpdateLoading(50, HOP2.Text)
                        If xAllocateID = "" Then xAction = 2
                    Case 50         '-->OP3
                        xAllocateID = oCommon.GetUserID(DOP3PER.Text)
                        xLevel = GetLevel(xLevel, HOP3.Text)
                        UpdateLoading(60, HOP3.Text)
                        If xAllocateID = "" Then xAction = 2
                    Case 60         '-->OP4
                        xAllocateID = oCommon.GetUserID(DOP4PER.Text)
                        xLevel = GetLevel(xLevel, HOP4.Text)
                        UpdateLoading(70, HOP4.Text)
                        If xAllocateID = "" Then xAction = 2
                    Case 70         '-->OP5
                        xAllocateID = oCommon.GetUserID(DOP5PER.Text)
                        xLevel = GetLevel(xLevel, HOP5.Text)
                        UpdateLoading(80, HOP5.Text)
                        If xAllocateID = "" Then xAction = 2
                    Case 80         '-->OP6
                        xAllocateID = oCommon.GetUserID(DOP6PER.Text)
                        xLevel = GetLevel(xLevel, HOP6.Text)
                        UpdateLoading(90, HOP6.Text)
                        If xAllocateID = "" Then xAction = 2
                    Case 90         '-->OP7
                        xAllocateID = oCommon.GetUserID(DOP7PER.Text)
                        xLevel = GetLevel(xLevel, HOP7.Text)
                        UpdateLoading(100, HOP7.Text)
                        If xAllocateID = "" Then xAction = 2
                    Case 100        '-->OP8
                        xAllocateID = oCommon.GetUserID(DOP8PER.Text)
                        xLevel = GetLevel(xLevel, HOP8.Text)
                        UpdateLoading(110, HOP8.Text)
                        If xAllocateID = "" Then xAction = 2
                    Case Else
                End Select
                '5.安排預定工程
                oCommon.ArrangFlowControl(pFormNo, pFormSno, xStep, xLevel, pApplyID, pUserID, xAllocateID, xAction, xNextStep)
                '
                'MsgBox("Step[" + CStr(xStep) + "]" + Chr(13) + _
                '       "NextStep[" + CStr(xNextStep) + "]" + Chr(13) + _
                '       "StartStep[" + CStr(pStartStep) + "]" + Chr(13) + _
                '       "EndStep[" + CStr(pEndStep) + "]")
                '下一工程
                xStep = xNextStep
            End While
            '
        End If
    End Sub
    'Add-End

End Class
