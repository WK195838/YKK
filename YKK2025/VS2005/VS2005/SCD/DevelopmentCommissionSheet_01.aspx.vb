Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Collections.Generic

Partial Class DevelopmentCommissionSheet_01
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
        '75
        SetParameter()                              '設定參數
        '100
        SetPopupFunction()                          '設定彈出視窗事件
        '3800
        ShowSheetFunction()                         '表單功能按鈕/欄位顯示
        '
        If Not IsPostBack Then                      'PostBack
            '110
            ShowSheetField("New")                   '表單欄位顯示及欄位輸入檢查
            '
            If wFormSno > 0 And wStep > 3 Then      '起單/簽核
                '3900
                ShowFormData()                      '顯示表單資料
                '4300
                UpdateTranFile()                    '更新交易資料
            End If
        Else
            '110
            ShowSheetField("PostBack")              '表單欄位顯示及欄位輸入檢查
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
        Response.Cookies("PGM").Value = "DevelopmentCommissionSheet_01.aspx"                        '程式名
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
        BCREATE.Attributes("onclick") = "CreateSampleSheet();"                                      '製作開發樣品
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
        '-- 開發委託
        '-----------------------------------------------------------------
        '----基本欄位設定-------------------------------------------------    
        'No
        Select Case FindFieldInf("1-NO")
            Case 0  '顯示
                DRNO.BackColor = Color.LightGray
                DRNO.Visible = True
                DRNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRNO.Visible = True
                DRNO.BackColor = Color.GreenYellow
                DRNO.ReadOnly = False
                ShowRequiredFieldValidator("DRNORqd", "DRNO", "異常：需輸入ＮＯ")
            Case 2  '修改
                DRNO.Visible = True
                DRNO.BackColor = Color.Yellow
                DRNO.ReadOnly = False
            Case Else   '隱藏
                DRNO.Visible = False
        End Select
        If pPost = "New" Then DRNO.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-BASE"), "APPBUYER", "ZZZZZZ")
        '委託廠商
        Select Case FindFieldInf("1-BASE")
            Case 0  '顯示
                DSELLVENDOR.BackColor = Color.LightGray
                DSELLVENDOR.Visible = True
                DSELLVENDOR.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSELLVENDOR.Visible = True
                DSELLVENDOR.BackColor = Color.GreenYellow
                DSELLVENDOR.ReadOnly = False
                ShowRequiredFieldValidator("DSELLVENDORRqd", "DSELLVENDOR", "異常：需輸入委託廠商")
            Case 2  '修改
                DSELLVENDOR.Visible = True
                DSELLVENDOR.BackColor = Color.Yellow
                DSELLVENDOR.ReadOnly = False
            Case Else   '隱藏
                DSELLVENDOR.Visible = False
        End Select
        If pPost = "New" Then DSELLVENDOR.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-BASE"), "USAGE", "ZZZZZZ")
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-BASE"), "NEEDMAP", "ZZZZZZ")
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-SAMPLE"), "PRO", "ZZZZZZ")
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
        If pPost = "New" Then DPLEN.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-SAMPLE"), "PLENUN", "ZZZZZZ")
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
        If pPost = "New" Then DPQTY.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-SAMPLE"), "PQTYUN", "ZZZZZZ")
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
        If pPost = "New" Then DEALEN.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-SAMPLE-EA"), "EALENUN", "ZZZZZZ")
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
        If pPost = "New" Then DEAQTY.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-SAMPLE-EA"), "EAQTYUN", "ZZZZZZ")
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
            Case 1  '修改+檢查
                DSIZENO.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSIZENORqd", "DSIZENO", "異常：需輸入型別")
                DSIZENO.Visible = True
            Case 2  '修改
                DSIZENO.BackColor = Color.Yellow
                DSIZENO.Visible = True
            Case Else   '隱藏
                DSIZENO.Visible = False
        End Select
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "SIZENO", "ZZZZZZ")
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "ITEM", "ZZZZZZ")
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "TATYPE", "ZZZZZZ")
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
        If pPost = "New" Then DTAWIDTH.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "ECOLSEL", "ZZZZZZ")
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "CCOLSEL", "ZZZZZZ")
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "TACOL", "ZZZZZZ")
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "TALCOL", "ZZZZZZ")
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "TARCOL", "ZZZZZZ")
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "THUPCOL", "ZZZZZZ")
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
        '縫上-色番(左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLUPCOL.BackColor = Color.LightGray
                DTHLUPCOL.Visible = True
            Case 1  '修改+檢查
                DTHLUPCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTHLUPCOLRqd", "DTHLUPCOL", "異常：需輸入縫上-色番(左)")
                DTHLUPCOL.Visible = True
            Case 2  '修改
                DTHLUPCOL.BackColor = Color.Yellow
                DTHLUPCOL.Visible = True
            Case Else   '隱藏
                DTHLUPCOL.Visible = False
        End Select
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "THLUPCOL", "ZZZZZZ")
        '縫上-色番(左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLUPCOLNO.BackColor = Color.LightGray
                DTHLUPCOLNO.Visible = True
                DTHLUPCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHLUPCOLNO.Visible = True
                DTHLUPCOLNO.BackColor = Color.GreenYellow
                DTHLUPCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHLUPCOLNORqd", "DTHLUPCOLNO", "異常：需輸入縫上-色番(左)")
            Case 2  '修改
                DTHLUPCOLNO.Visible = True
                DTHLUPCOLNO.BackColor = Color.Yellow
                DTHLUPCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHLUPCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHLUPCOLNO.Text = ""
        '縫上-YKK(左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLUPYCOLNO.BackColor = Color.LightGray
                DTHLUPYCOLNO.Visible = True
                DTHLUPYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHLUPYCOLNO.Visible = True
                DTHLUPYCOLNO.BackColor = Color.GreenYellow
                DTHLUPYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHLUPYCOLNORqd", "DTHLUPYCOLNO", "異常：需輸入縫上-YKK(左)")
            Case 2  '修改
                DTHLUPYCOLNO.Visible = True
                DTHLUPYCOLNO.BackColor = Color.Yellow
                DTHLUPYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHLUPYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHLUPYCOLNO.Text = ""
        '縫上-色番(右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRUPCOL.BackColor = Color.LightGray
                DTHRUPCOL.Visible = True
            Case 1  '修改+檢查
                DTHRUPCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTHRUPCOLRqd", "DTHRUPCOL", "異常：需輸入縫上-色番(右)")
                DTHRUPCOL.Visible = True
            Case 2  '修改
                DTHRUPCOL.BackColor = Color.Yellow
                DTHRUPCOL.Visible = True
            Case Else   '隱藏
                DTHRUPCOL.Visible = False
        End Select
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "THRUPCOL", "ZZZZZZ")
        '縫上-色番(右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRUPCOLNO.BackColor = Color.LightGray
                DTHRUPCOLNO.Visible = True
                DTHRUPCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHRUPCOLNO.Visible = True
                DTHRUPCOLNO.BackColor = Color.GreenYellow
                DTHRUPCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHRUPCOLNORqd", "DTHRUPCOLNO", "異常：需輸入縫上-色番(右)")
            Case 2  '修改
                DTHRUPCOLNO.Visible = True
                DTHRUPCOLNO.BackColor = Color.Yellow
                DTHRUPCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHRUPCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHRUPCOLNO.Text = ""
        '縫上-YKK(右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRUPYCOLNO.BackColor = Color.LightGray
                DTHRUPYCOLNO.Visible = True
                DTHRUPYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHRUPYCOLNO.Visible = True
                DTHRUPYCOLNO.BackColor = Color.GreenYellow
                DTHRUPYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHRUPYCOLNORqd", "DTHRUPYCOLNO", "異常：需輸入縫上-YKK(右)")
            Case 2  '修改
                DTHRUPYCOLNO.Visible = True
                DTHRUPYCOLNO.BackColor = Color.Yellow
                DTHRUPYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHRUPYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHRUPYCOLNO.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "THLOCOL", "ZZZZZZ")
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
        '縫下-色番(左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLLOCOL.BackColor = Color.LightGray
                DTHLLOCOL.Visible = True
            Case 1  '修改+檢查
                DTHLLOCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTHLLOCOLRqd", "DTHLLOCOL", "異常：需輸入縫下-色番(左)")
                DTHLLOCOL.Visible = True
            Case 2  '修改
                DTHLLOCOL.BackColor = Color.Yellow
                DTHLLOCOL.Visible = True
            Case Else   '隱藏
                DTHLLOCOL.Visible = False
        End Select
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "THLLOCOL", "ZZZZZZ")
        '縫下-色番(左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLLOCOLNO.BackColor = Color.LightGray
                DTHLLOCOLNO.Visible = True
                DTHLLOCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHLLOCOLNO.Visible = True
                DTHLLOCOLNO.BackColor = Color.GreenYellow
                DTHLLOCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHLLOCOLNORqd", "DTHLLOCOLNO", "異常：需輸入縫下-色番(左)")
            Case 2  '修改
                DTHLLOCOLNO.Visible = True
                DTHLLOCOLNO.BackColor = Color.Yellow
                DTHLLOCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHLLOCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHLLOCOLNO.Text = ""
        '縫下-YKK(左)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHLLOYCOLNO.BackColor = Color.LightGray
                DTHLLOYCOLNO.Visible = True
                DTHLLOYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHLLOYCOLNO.Visible = True
                DTHLLOYCOLNO.BackColor = Color.GreenYellow
                DTHLLOYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHLLOYCOLNORqd", "DTHLLOYCOLNO", "異常：需輸入縫下-YKK(左)")
            Case 2  '修改
                DTHLLOYCOLNO.Visible = True
                DTHLLOYCOLNO.BackColor = Color.Yellow
                DTHLLOYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHLLOYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHLLOYCOLNO.Text = ""
        '縫下-色番(右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRLOCOL.BackColor = Color.LightGray
                DTHRLOCOL.Visible = True
            Case 1  '修改+檢查
                DTHRLOCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTHRLOCOLRqd", "DTHRLOCOL", "異常：需輸入縫下-色番(右)")
                DTHRLOCOL.Visible = True
            Case 2  '修改
                DTHRLOCOL.BackColor = Color.Yellow
                DTHRLOCOL.Visible = True
            Case Else   '隱藏
                DTHRLOCOL.Visible = False
        End Select
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "THRLOCOL", "ZZZZZZ")
        '縫下-色番(右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRLOCOLNO.BackColor = Color.LightGray
                DTHRLOCOLNO.Visible = True
                DTHRLOCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHRLOCOLNO.Visible = True
                DTHRLOCOLNO.BackColor = Color.GreenYellow
                DTHRLOCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHRLOCOLNORqd", "DTHRLOCOLNO", "異常：需輸入縫下-色番(右)")
            Case 2  '修改
                DTHRLOCOLNO.Visible = True
                DTHRLOCOLNO.BackColor = Color.Yellow
                DTHRLOCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHRLOCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHRLOCOLNO.Text = ""
        '縫下-YKK(右)
        Select Case FindFieldInf("1-DEVELOP")
            Case 0  '顯示
                DTHRLOYCOLNO.BackColor = Color.LightGray
                DTHRLOYCOLNO.Visible = True
                DTHRLOYCOLNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DTHRLOYCOLNO.Visible = True
                DTHRLOYCOLNO.BackColor = Color.GreenYellow
                DTHRLOYCOLNO.ReadOnly = False
                ShowRequiredFieldValidator("DTHRLOYCOLNORqd", "DTHRLOYCOLNO", "異常：需輸入縫下-YKK(右)")
            Case 2  '修改
                DTHRLOYCOLNO.Visible = True
                DTHRLOYCOLNO.BackColor = Color.Yellow
                DTHRLOYCOLNO.ReadOnly = False
            Case Else   '隱藏
                DTHRLOYCOLNO.Visible = False
        End Select
        If pPost = "New" Then DTHRLOYCOLNO.Text = ""
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
        If pPost = "New" Then DXMLEN.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "XMCOL", "ZZZZZZ")
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
        If pPost = "New" Then DAMLEN.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "AMCOL", "ZZZZZZ")
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
        If pPost = "New" Then DBMLEN.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "BMCOL", "ZZZZZZ")
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
        If pPost = "New" Then DCMLEN.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "CMCOL", "ZZZZZZ")
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
        If pPost = "New" Then DDMLEN.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "DMCOL", "ZZZZZZ")
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
        If pPost = "New" Then DEMLEN.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "EMCOL", "ZZZZZZ")
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
        If pPost = "New" Then DFMLEN.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "FMCOL", "ZZZZZZ")
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
        If pPost = "New" Then DGMLEN.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "GMCOL", "ZZZZZZ")
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
        If pPost = "New" Then DHMLEN.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "HMCOL", "ZZZZZZ")
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
        If pPost = "New" Then DLYLEN.Text = ""
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-DEVELOP"), "LYCOL", "ZZZZZZ")
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
        '製圖者
        Select Case FindFieldInf("1-MAP")
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-MAP"), "MAKEMAP", "ZZZZZZ")
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
        If Not IsPostBack Then SetFieldData(FindFieldInf("1-MAP"), "LEVEL", "ZZZZZZ")
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
        'OP1
        '----遲納原因類別1
        DOP1DELAYC1.BackColor = Color.Yellow
        DOP1DELAYC1.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("2-DELAYCAT"), "OP1DELAYC1", "ZZZZZZ")
        '----遲納原因類別2
        DOP1DELAYC2.BackColor = Color.Yellow
        DOP1DELAYC2.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("2-DELAYCAT"), "OP1DELAYC2", "ZZZZZZ")
        '----遲納原因內容
        DOP1REM.Visible = True
        DOP1REM.BackColor = Color.Yellow
        DOP1REM.ReadOnly = False
        If pPost = "New" Then DOP1REM.Text = ""
        'OP2
        '----遲納原因類別1
        DOP2DELAYC1.BackColor = Color.Yellow
        DOP2DELAYC1.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("2-DELAYCAT"), "OP2DELAYC1", "ZZZZZZ")
        '----遲納原因類別2
        DOP2DELAYC2.BackColor = Color.Yellow
        DOP2DELAYC2.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("2-DELAYCAT"), "OP2DELAYC2", "ZZZZZZ")
        '----遲納原因內容
        DOP2REM.Visible = True
        DOP2REM.BackColor = Color.Yellow
        DOP2REM.ReadOnly = False
        If pPost = "New" Then DOP2REM.Text = ""
        'OP3
        '----遲納原因類別1
        DOP3DELAYC1.BackColor = Color.Yellow
        DOP3DELAYC1.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("2-DELAYCAT"), "OP3DELAYC1", "ZZZZZZ")
        '----遲納原因類別2
        DOP3DELAYC2.BackColor = Color.Yellow
        DOP3DELAYC2.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("2-DELAYCAT"), "OP3DELAYC2", "ZZZZZZ")
        '----遲納原因內容
        DOP3REM.Visible = True
        DOP3REM.BackColor = Color.Yellow
        DOP3REM.ReadOnly = False
        If pPost = "New" Then DOP3REM.Text = ""
        'OP4
        '----遲納原因類別1
        DOP4DELAYC1.BackColor = Color.Yellow
        DOP4DELAYC1.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("2-DELAYCAT"), "OP4DELAYC1", "ZZZZZZ")
        '----遲納原因類別2
        DOP4DELAYC2.BackColor = Color.Yellow
        DOP4DELAYC2.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("2-DELAYCAT"), "OP4DELAYC2", "ZZZZZZ")
        '----遲納原因內容
        DOP4REM.Visible = True
        DOP4REM.BackColor = Color.Yellow
        DOP4REM.ReadOnly = False
        If pPost = "New" Then DOP4REM.Text = ""
        'OP5
        '----遲納原因類別1
        DOP5DELAYC1.BackColor = Color.Yellow
        DOP5DELAYC1.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("2-DELAYCAT"), "OP5DELAYC1", "ZZZZZZ")
        '----遲納原因類別2
        DOP5DELAYC2.BackColor = Color.Yellow
        DOP5DELAYC2.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("2-DELAYCAT"), "OP5DELAYC2", "ZZZZZZ")
        '----遲納原因內容
        DOP5REM.Visible = True
        DOP5REM.BackColor = Color.Yellow
        DOP5REM.ReadOnly = False
        If pPost = "New" Then DOP5REM.Text = ""
        'OP6
        '----遲納原因類別1
        DOP6DELAYC1.BackColor = Color.Yellow
        DOP6DELAYC1.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("2-DELAYCAT"), "OP6DELAYC1", "ZZZZZZ")
        '----遲納原因類別2
        DOP6DELAYC2.BackColor = Color.Yellow
        DOP6DELAYC2.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("2-DELAYCAT"), "OP6DELAYC2", "ZZZZZZ")
        '----遲納原因內容
        DOP6REM.Visible = True
        DOP6REM.BackColor = Color.Yellow
        DOP6REM.ReadOnly = False
        If pPost = "New" Then DOP6REM.Text = ""
        'OP7
        '----遲納原因類別1
        DOP7DELAYC1.BackColor = Color.Yellow
        DOP7DELAYC1.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("2-DELAYCAT"), "OP7DELAYC1", "ZZZZZZ")
        '----遲納原因類別2
        DOP7DELAYC2.BackColor = Color.Yellow
        DOP7DELAYC2.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("2-DELAYCAT"), "OP7DELAYC2", "ZZZZZZ")
        '----遲納原因內容
        DOP7REM.Visible = True
        DOP7REM.BackColor = Color.Yellow
        DOP7REM.ReadOnly = False
        If pPost = "New" Then DOP7REM.Text = ""
        'OP8
        '----遲納原因類別1
        DOP8DELAYC1.BackColor = Color.Yellow
        DOP8DELAYC1.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("2-DELAYCAT"), "OP8DELAYC1", "ZZZZZZ")
        '----遲納原因類別2
        DOP8DELAYC2.BackColor = Color.Yellow
        DOP8DELAYC2.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("2-DELAYCAT"), "OP8DELAYC2", "ZZZZZZ")
        '----遲納原因內容
        DOP8REM.Visible = True
        DOP8REM.BackColor = Color.Yellow
        DOP8REM.ReadOnly = False
        If pPost = "New" Then DOP8REM.Text = ""
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
        If pPost = "New" Then LQCFILE1.Visible = False
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
        If pPost = "New" Then LQCFILE2.Visible = False
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
        If pPost = "New" Then LQCFILE3.Visible = False
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
        If pPost = "New" Then LQCFILE4.Visible = False
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
        If pPost = "New" Then LQCFILE5.Visible = False
        '承認-作成者
        D3WF1.BackColor = Color.Yellow
        D3WF1.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("3-FLOW"), "WF1", "ZZZZZZ")
        '承認-責任者
        D3WF2.BackColor = Color.Yellow
        D3WF2.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("3-FLOW"), "WF2", "ZZZZZZ")
        '承認-製造1
        D3WF3Name.BackColor = Color.Yellow
        D3WF3Name.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("3-FLOW"), "WF3NAME", "ZZZZZZ")
        '
        D3WF3.BackColor = Color.Yellow
        D3WF3.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("3-FLOW"), "WF3", "ZZZZZZ")
        '承認-製造2
        D3WF4Name.BackColor = Color.Yellow
        D3WF4Name.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("3-FLOW"), "WF4NAME", "ZZZZZZ")
        '
        D3WF4.BackColor = Color.Yellow
        D3WF4.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("3-FLOW"), "WF4", "ZZZZZZ")
        '承認-製造3
        D3WF5Name.BackColor = Color.Yellow
        D3WF5Name.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("3-FLOW"), "WF5NAME", "ZZZZZZ")
        '
        D3WF5.BackColor = Color.Yellow
        D3WF5.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("3-FLOW"), "WF5", "ZZZZZZ")
        '承認-製造4
        D3WF6Name.BackColor = Color.Yellow
        D3WF6Name.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("3-FLOW"), "WF6NAME", "ZZZZZZ")
        '
        D3WF6.BackColor = Color.Yellow
        D3WF6.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("3-FLOW"), "WF6", "ZZZZZZ")
        '承認-廠長
        D3WF7Name.BackColor = Color.Yellow
        D3WF7Name.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("3-FLOW"), "WF7NAME", "ZZZZZZ")
        '
        D3WF7.BackColor = Color.Yellow
        D3WF7.Visible = True
        If Not IsPostBack Then SetFieldData(FindFieldInf("3-FLOW"), "WF7", "ZZZZZZ")
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-A' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
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
        '縫上-色番(左)
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
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
        '縫上-色番(右)
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
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
        '縫下-色番(左)
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
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
        '縫下-色番(右)
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC1' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC2' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC1' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC2' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC1' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC2' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC1' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC2' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC1' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC2' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC1' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC2' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC1' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC2' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC1' Order by Data "
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
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC2' Order by Data "
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
        Top = 1420
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
                If dtWaitHandle.Rows(0)("FlowType") = 0 Then
                    DDecideDesc.BackColor = Color.LightGray
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "異常：需輸入說明")
                End If
                '位置設定
                DDescSheet.Style("top") = Top & "px"
                DDecideDesc.Style("top") = Top + 5 & "px"
                Top = Top + 80
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
                wNGSts1 = dtFlow.Rows(0)("NGSts1") + 1
                BNG1.Style("top") = Top & "px"
            Else
                BNG1.Visible = False
            End If
            'NG-2按鈕
            If dtFlow.Rows(0)("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Text = dtFlow.Rows(0)("NGDesc2")
                'BNG2.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BNG2.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wNGSts2 = dtFlow.Rows(0)("NGSts2") + 1
                BNG2.Style("top") = Top & "px"
            Else
                BNG2.Visible = False
            End If
            'OK按鈕
            If dtFlow.Rows(0)("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Text = dtFlow.Rows(0)("OKDesc")
                'BOK.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BOK.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wOKSts = dtFlow.Rows(0)("OKSts") + 1
                BOK.Style("top") = Top & "px"
            Else
                BOK.Visible = False
            End If
            '頁次設定
            SetImageButtonImageFile(dtFlow.Rows(0)("PageIdx"))
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
        SQL = "Select * From F_DevelopmentCommissionSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtCommissionSheet As DataTable = uDataBase.GetDataTable(SQL)
        If dtCommissionSheet.Rows.Count > 0 Then
            '-----------------------------------------------------------------
            '-- 開發委託單
            '-----------------------------------------------------------------
            '----基本欄位設定-------------------------------------------------    
            DRNO.Text = dtCommissionSheet.Rows(0).Item("RNO")                           'NO
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
            Else
                LMAPREFFILE.Visible = False
            End If
            '----樣品欄位設定-------------------------------------------------    
            SetFieldData(FindFieldInf("1-SAMPLE"), "PRO", dtCommissionSheet.Rows(0).Item("NEEDMAP"))        '製品區分
            DOPPART.Text = dtCommissionSheet.Rows(0).Item("OPPART")                     '開具(色)
            DPLEN.Text = dtCommissionSheet.Rows(0).Item("PLEN")                         '長度(企)
            SetFieldData(FindFieldInf("1-SAMPLE"), "PLENUN", dtCommissionSheet.Rows(0).Item("PLENUN"))      '長度單位(企)
            DPQTY.Text = dtCommissionSheet.Rows(0).Item("PQTY")                         '數量(企)
            SetFieldData(FindFieldInf("1-SAMPLE"), "PQTYUN", dtCommissionSheet.Rows(0).Item("PQTYUN"))      '數量單位(企)
            DPLEN.Text = dtCommissionSheet.Rows(0).Item("EALEN")                        '長度(EA)
            SetFieldData(FindFieldInf("1-SAMPLE-EA"), "EALENUN", dtCommissionSheet.Rows(0).Item("EALENUN")) '長度單位(EA)
            DPQTY.Text = dtCommissionSheet.Rows(0).Item("EAQTY")                        '數量(EA)
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
            DECOL.Text = dtCommissionSheet.Rows(0).Item("CCOL")                         '丸紐
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
            SetFieldData(FindFieldInf("1-DEVELOP"), "THLUPCOL", dtCommissionSheet.Rows(0).Item("THLUPCOL")) '縫上-色番(左)
            DTHLUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLUPCOLNO")             '縫上-色番(左)
            DTHLUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLUPYCOLNO")           '縫上-YKK(左)
            SetFieldData(FindFieldInf("1-DEVELOP"), "THRUPCOL", dtCommissionSheet.Rows(0).Item("THRUPCOL")) '縫上-色番(右)
            DTHRUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRUPCOLNO")             '縫上-色番(右)
            DTHRUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRUPYCOLNO")           '縫上-YKK(右)
            SetFieldData(FindFieldInf("1-DEVELOP"), "THLOCOL", dtCommissionSheet.Rows(0).Item("THLOCOL"))   '縫下-色番(同)
            DTHLOCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLOCOLNO")               '縫下-色番(同)
            DTHLOYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLOYCOLNO")             '縫下-YKK(同)
            SetFieldData(FindFieldInf("1-DEVELOP"), "THLLOCOL", dtCommissionSheet.Rows(0).Item("THLLOCOL")) '縫下-色番(左)
            DTHLLOCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLLOCOLNO")             '縫下-色番(左)
            DTHLLOYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLLOYCOLNO")           '縫下-YKK(左)
            SetFieldData(FindFieldInf("1-DEVELOP"), "THRLOCOL", dtCommissionSheet.Rows(0).Item("THRLOCOL")) '縫下-色番(右)
            DTHRLOCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRLOCOLNO")             '縫下-色番(右)
            DTHRLOYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRLOYCOLNO")           '縫下-YKK(右)
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
            SetFieldData(FindFieldInf("1-MAP"), "MAKEMAP", dtCommissionSheet.Rows(0).Item("MAKEMAP"))       '製圖者
            SetFieldData(FindFieldInf("1-MAP"), "LEVEL", dtCommissionSheet.Rows(0).Item("LEVEL"))           '難易度
            '圖檔
            If dtCommissionSheet.Rows(0).Item("MAPFILE") <> "" Then
                LMAPFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("MAPFILE")
                LMAPFILE.Visible = True
            Else
                LMAPFILE.Visible = False
            End If
            '----適用型別欄位設定-------------------------------------------------    
            '適用型別檔
            If dtCommissionSheet.Rows(0).Item("FORTYPEFILE") <> "" Then
                LFORTYPEFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("FORTYPEFILE")
                LFORTYPEFILE.Visible = True
            Else
                LFORTYPEFILE.Visible = False
            End If
            '-----------------------------------------------------------------
            '-- 製造委託
            '-----------------------------------------------------------------
            SQL = "Select * From ManufactureSheet "
            SQL &= " Where RNO =  '" & dtCommissionSheet.Rows(0).Item("RNO") & "'"
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
                Else
                    LHINTFILE.Visible = False
                End If
                DUPSTK.Text = dtManufactureSheet.Rows(0).Item("UPSTK")                  '上止
                DLOSTK.Text = dtManufactureSheet.Rows(0).Item("LOSTK")                  '下止
                DOPPART.Text = dtManufactureSheet.Rows(0).Item("OPPART")                '開具(色)
                DTASPEC.Text = dtManufactureSheet.Rows(0).Item("TASPEC")                '布帶
                DECOL.Text = dtManufactureSheet.Rows(0).Item("ECOL")                    '鏈齒顏色
                DCCOL.Text = dtManufactureSheet.Rows(0).Item("CCOL")                    '丸紐
                DTHSPEC.Text = dtManufactureSheet.Rows(0).Item("THSPEC")                '縫工線
                DPLEN.Text = dtManufactureSheet.Rows(0).Item("PLEN")                    '長度(企)
                DPQTY.Text = dtManufactureSheet.Rows(0).Item("PQTY")                    '數量(企)
                DEALEN.Text = dtManufactureSheet.Rows(0).Item("EALEN")                  '長度(EA)
                DEAQTY.Text = dtManufactureSheet.Rows(0).Item("EAQTY")                  '數量(EA)
                '----工程-------------------------------------------------    
                DMANUFTYPE.Text = dtManufactureSheet.Rows(0).Item("MANUFTYPE")          '內外製
                DOP1.Text = dtManufactureSheet.Rows(0).Item("OP1")                      'OP1-工程
                DOP1PER.Text = dtManufactureSheet.Rows(0).Item("OP1PER")                'OP1-擔當
                DOP1BTIME.Text = dtManufactureSheet.Rows(0).Item("OP1BTIME")            'OP1-預定納期
                DOP1BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP1BHOURS")          'OP1-預定時數
                DOP1ATIME.Text = dtManufactureSheet.Rows(0).Item("OP1ATIME")            'OP1-實際納期
                DOP1AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP1AHOURS")          'OP1-實際時數
                DOP1CON.Text = dtManufactureSheet.Rows(0).Item("OP1CON")                'OP1-作業內容
                SetFieldData(FindFieldInf("2-DELAYCAT"), "OP1DELAYC1", dtManufactureSheet.Rows(0).Item("OP1DELAYC1"))       'OP1-遲納原因-1
                SetFieldData(FindFieldInf("2-DELAYCAT"), "OP1DELAYC2", dtManufactureSheet.Rows(0).Item("OP1DELAYC2"))       'OP1-遲納原因-2
                DOP1REM.Text = dtManufactureSheet.Rows(0).Item("OP1REM")                'OP1-遲納原因

                DOP2.Text = dtManufactureSheet.Rows(0).Item("OP2")                      'OP2-工程
                DOP2PER.Text = dtManufactureSheet.Rows(0).Item("OP2PER")                'OP2-擔當
                DOP2BTIME.Text = dtManufactureSheet.Rows(0).Item("OP2BTIME")            'OP2-預定納期
                DOP2BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP2BHOURS")          'OP2-預定時數
                DOP2ATIME.Text = dtManufactureSheet.Rows(0).Item("OP2ATIME")            'OP2-實際納期
                DOP2AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP2AHOURS")          'OP2-實際時數
                DOP2CON.Text = dtManufactureSheet.Rows(0).Item("OP2CON")                'OP2-作業內容
                SetFieldData(FindFieldInf("2-DELAYCAT"), "OP2DELAYC1", dtManufactureSheet.Rows(0).Item("OP2DELAYC1"))       'OP2-遲納原因-1
                SetFieldData(FindFieldInf("2-DELAYCAT"), "OP2DELAYC2", dtManufactureSheet.Rows(0).Item("OP2DELAYC2"))       'OP2-遲納原因-2
                DOP2REM.Text = dtManufactureSheet.Rows(0).Item("OP2REM")                'OP2-遲納原因

                DOP3.Text = dtManufactureSheet.Rows(0).Item("OP3")                      'OP3-工程
                DOP3PER.Text = dtManufactureSheet.Rows(0).Item("OP3PER")                'OP3-擔當
                DOP3BTIME.Text = dtManufactureSheet.Rows(0).Item("OP3BTIME")            'OP3-預定納期
                DOP3BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP3BHOURS")          'OP3-預定時數
                DOP3ATIME.Text = dtManufactureSheet.Rows(0).Item("OP3ATIME")            'OP3-實際納期
                DOP3AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP3AHOURS")          'OP3-實際時數
                DOP3CON.Text = dtManufactureSheet.Rows(0).Item("OP3CON")                'OP3-作業內容
                SetFieldData(FindFieldInf("2-DELAYCAT"), "OP3DELAYC1", dtManufactureSheet.Rows(0).Item("OP3DELAYC1"))       'OP3-遲納原因-1
                SetFieldData(FindFieldInf("2-DELAYCAT"), "OP3DELAYC2", dtManufactureSheet.Rows(0).Item("OP3DELAYC2"))       'OP3-遲納原因-2
                DOP3REM.Text = dtManufactureSheet.Rows(0).Item("OP3REM")                'OP3-遲納原因

                DOP4.Text = dtManufactureSheet.Rows(0).Item("OP4")                      'OP4-工程
                DOP4PER.Text = dtManufactureSheet.Rows(0).Item("OP4PER")                'OP4-擔當
                DOP4BTIME.Text = dtManufactureSheet.Rows(0).Item("OP4BTIME")            'OP4-預定納期
                DOP4BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP4BHOURS")          'OP4-預定時數
                DOP4ATIME.Text = dtManufactureSheet.Rows(0).Item("OP4ATIME")            'OP4-實際納期
                DOP4AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP4AHOURS")          'OP4-實際時數
                DOP4CON.Text = dtManufactureSheet.Rows(0).Item("OP4CON")                'OP4-作業內容
                SetFieldData(FindFieldInf("2-DELAYCAT"), "OP4DELAYC1", dtManufactureSheet.Rows(0).Item("OP4DELAYC1"))       'OP4-遲納原因-1
                SetFieldData(FindFieldInf("2-DELAYCAT"), "OP4DELAYC2", dtManufactureSheet.Rows(0).Item("OP4DELAYC2"))       'OP4-遲納原因-2
                DOP4REM.Text = dtManufactureSheet.Rows(0).Item("OP4REM")                'OP4-遲納原因

                DOP5.Text = dtManufactureSheet.Rows(0).Item("OP5")                      'OP5-工程
                DOP5PER.Text = dtManufactureSheet.Rows(0).Item("OP5PER")                'OP5-擔當
                DOP5BTIME.Text = dtManufactureSheet.Rows(0).Item("OP5BTIME")            'OP5-預定納期
                DOP5BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP5BHOURS")          'OP5-預定時數
                DOP5ATIME.Text = dtManufactureSheet.Rows(0).Item("OP5ATIME")            'OP5-實際納期
                DOP5AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP5AHOURS")          'OP5-實際時數
                DOP5CON.Text = dtManufactureSheet.Rows(0).Item("OP5CON")                'OP5-作業內容
                SetFieldData(FindFieldInf("2-DELAYCAT"), "OP5DELAYC1", dtManufactureSheet.Rows(0).Item("OP5DELAYC1"))       'OP5-遲納原因-1
                SetFieldData(FindFieldInf("2-DELAYCAT"), "OP5DELAYC2", dtManufactureSheet.Rows(0).Item("OP5DELAYC2"))       'OP5-遲納原因-2
                DOP5REM.Text = dtManufactureSheet.Rows(0).Item("OP5REM")                'OP5-遲納原因

                DOP6.Text = dtManufactureSheet.Rows(0).Item("OP6")                      'OP6-工程
                DOP6PER.Text = dtManufactureSheet.Rows(0).Item("OP6PER")                'OP6-擔當
                DOP6BTIME.Text = dtManufactureSheet.Rows(0).Item("OP6BTIME")            'OP6-預定納期
                DOP6BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP6BHOURS")          'OP6-預定時數
                DOP6ATIME.Text = dtManufactureSheet.Rows(0).Item("OP6ATIME")            'OP6-實際納期
                DOP6AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP6AHOURS")          'OP6-實際時數
                DOP6CON.Text = dtManufactureSheet.Rows(0).Item("OP6CON")                'OP6-作業內容
                SetFieldData(FindFieldInf("2-DELAYCAT"), "OP6DELAYC1", dtManufactureSheet.Rows(0).Item("OP6DELAYC1"))       'OP6-遲納原因-1
                SetFieldData(FindFieldInf("2-DELAYCAT"), "OP6DELAYC2", dtManufactureSheet.Rows(0).Item("OP6DELAYC2"))       'OP6-遲納原因-2
                DOP6REM.Text = dtManufactureSheet.Rows(0).Item("OP6REM")                'OP6-遲納原因

                DOP7.Text = dtManufactureSheet.Rows(0).Item("OP7")                      'OP7-工程
                DOP7PER.Text = dtManufactureSheet.Rows(0).Item("OP7PER")                'OP7-擔當
                DOP7BTIME.Text = dtManufactureSheet.Rows(0).Item("OP7BTIME")            'OP7-預定納期
                DOP7BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP7BHOURS")          'OP7-預定時數
                DOP7ATIME.Text = dtManufactureSheet.Rows(0).Item("OP7ATIME")            'OP7-實際納期
                DOP7AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP7AHOURS")          'OP7-實際時數
                DOP7CON.Text = dtManufactureSheet.Rows(0).Item("OP7CON")                'OP7-作業內容
                SetFieldData(FindFieldInf("2-DELAYCAT"), "OP7DELAYC1", dtManufactureSheet.Rows(0).Item("OP7DELAYC1"))       'OP7-遲納原因-1
                SetFieldData(FindFieldInf("2-DELAYCAT"), "OP7DELAYC2", dtManufactureSheet.Rows(0).Item("OP7DELAYC2"))       'OP7-遲納原因-2
                DOP7REM.Text = dtManufactureSheet.Rows(0).Item("OP7REM")                'OP7-遲納原因

                DOP8.Text = dtManufactureSheet.Rows(0).Item("OP8")                      'OP8-工程
                DOP8PER.Text = dtManufactureSheet.Rows(0).Item("OP8PER")                'OP8-擔當
                DOP8BTIME.Text = dtManufactureSheet.Rows(0).Item("OP8BTIME")            'OP8-預定納期
                DOP8BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP8BHOURS")          'OP8-預定時數
                DOP8ATIME.Text = dtManufactureSheet.Rows(0).Item("OP8ATIME")            'OP8-實際納期
                DOP8AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP8AHOURS")          'OP8-實際時數
                DOP8CON.Text = dtManufactureSheet.Rows(0).Item("OP8CON")                'OP8-作業內容
                SetFieldData(FindFieldInf("2-DELAYCAT"), "OP8DELAYC1", dtManufactureSheet.Rows(0).Item("OP8DELAYC1"))       'OP8-遲納原因-1
                SetFieldData(FindFieldInf("2-DELAYCAT"), "OP8DELAYC2", dtManufactureSheet.Rows(0).Item("OP8DELAYC2"))       'OP8-遲納原因-2
                DOP8REM.Text = dtManufactureSheet.Rows(0).Item("OP8REM")                'OP8-遲納原因
            End If
            '-----------------------------------------------------------------
            '-- 開發見本
            '-----------------------------------------------------------------
            SQL = "Select * From SampleSheet "
            SQL &= " Where RNO =  '" & dtCommissionSheet.Rows(0).Item("RNO") & "'"
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
                Else
                    LSAMPLEFILE.Visible = False
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
                    LQCFILE1.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile1")
                Else
                    LQCFILE1.Visible = False
                End If
                If dtSampleSheet.Rows(0).Item("QCFile2") <> "" Then                     '品測報告2
                    LQCFILE2.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile2")
                Else
                    LQCFILE2.Visible = False
                End If
                If dtSampleSheet.Rows(0).Item("QCFile3") <> "" Then                     '品測報告3
                    LQCFILE3.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile3")
                Else
                    LQCFILE3.Visible = False
                End If
                If dtSampleSheet.Rows(0).Item("QCFile4") <> "" Then                     '品測報告4
                    LQCFILE4.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile4")
                Else
                    LQCFILE4.Visible = False
                End If
                If dtSampleSheet.Rows(0).Item("QCFile5") <> "" Then                     '品測報告5
                    LQCFILE5.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile5")
                Else
                    LQCFILE5.Visible = False
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
            SQL = "Select * From GentaniSheet "
            SQL &= " Where RNO =  '" & dtCommissionSheet.Rows(0).Item("RNO") & "'"
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
                        If s.Contains(dc.ColumnName) Then
                            l = Me.form1.FindControl("D4" & dc.ColumnName & "1")
                            l.Text = v
                            l = Me.form1.FindControl("D4" & dc.ColumnName & "2")
                            l.Text = v
                        End If
                    End If
                Next
                '
            End If
            '-----------------------------------------------------------------
            '-- 核定履歷
            '-----------------------------------------------------------------
            SQL = "SELECT "
            SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
            SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
            SQL = SQL + "AEndTimeDesc As Description, "
            SQL = SQL + "URL "
            SQL = SQL + "FROM V_WaitHandle_01 "
            SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
            SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
            'SQL = SQL + "Order by Unique_ID Desc "
            SQL = SQL + "Order by CreateTime Desc, Step Desc, SeqNo Desc "
            GridView2.DataSource = uDataBase.GetDataTable(SQL)
            GridView2.DataBind()
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
        'If InputDataOK(0) Then
        '    DisabledButton()   '停止Button運作
        '    FlowControl("OK", 0, "1")
        'Else
        '    EnabledButton()   '起動Button運作
        'End If
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

        '----附加檔案檢查----------------------------------------------------------------------------------
        '--開發委託
        '草圖
        If ErrCode = 0 Then
            If DMAPREFFILE.Visible Then
                If DMAPREFFILE.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DMAPREFFILE)
                End If
            End If
        End If
        '圖檔
        If ErrCode = 0 Then
            If DMAPFILE.Visible Then
                If DMAPFILE.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DMAPFILE)
                End If
            End If
        End If
        '適用型別檔
        If ErrCode = 0 Then
            If DFORTYPEFILE.Visible Then
                If DFORTYPEFILE.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DFORTYPEFILE)
                End If
            End If
        End If
        '--開發見本
        '品質-吋法
        If ErrCode = 0 Then
            If D3QCFILE1.Visible Then
                If D3QCFILE1.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(D3QCFILE1)
                End If
            End If
        End If
        '品質-強度
        If ErrCode = 0 Then
            If D3QCFILE2.Visible Then
                If D3QCFILE2.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(D3QCFILE2)
                End If
            End If
        End If
        '品質-往覆測試
        If ErrCode = 0 Then
            If D3QCFILE3.Visible Then
                If D3QCFILE3.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(D3QCFILE3)
                End If
            End If
        End If
        '品質-式樣／組織
        If ErrCode = 0 Then
            If D3QCFILE4.Visible Then
                If D3QCFILE4.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(D3QCFILE4)
                End If
            End If
        End If
        '品質-其他
        If ErrCode = 0 Then
            If D3QCFILE5.Visible Then
                If D3QCFILE5.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(D3QCFILE5)
                End If
            End If
        End If
        '----系統檢查----------------------------------------------------------------------------------
        '檢查委託書No
        If ErrCode = 0 Then
            If DRNO.Text <> "" Then
                ErrCode = oCommon.CommissionNo("002002", wFormSno, wStep, DRNO.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9050
                End If
            End If
        End If
        '----異常訊息處理----------------------------------------------------------------------------------
        If ErrCode <> 0 Then
            If ErrCode = 9010 Then Message = ""
            If ErrCode = 9020 Then Message = "上傳檔案格式有誤,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案Size有誤,請確認!"
            If ErrCode = 9040 Then Message = ""
            If ErrCode = 9050 Then Message = "委託書No.重覆,請確認委託書No.!"
            If ErrCode = 9060 Then Message = ""
            If ErrCode = 9070 Then Message = ""
            If ErrCode = 9071 Then Message = ""
            If ErrCode = 9072 Then Message = ""
            uJavaScript.PopMsg(Me, Message)
        Else
            isOK = True
        End If
        '
        Return isOK
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
        Dim Path As String = Server.MapPath(uCommon.GetAppSetting("COMMISSIONFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim sql As String = ""
        '---------------------------------------------------------------------------------
        '-- 開發委託
        '---------------------------------------------------------------------------------
        sql = "INSERT INTO F_DevelopmentCommissionSheet ( " & _
              "Sts, CompletedTime, FormNo, FormSno, " & _
              "RNO, APPDATE, APPDEPT, APPPER, APPBUYER, " & _
              "SellVendor, ESYQTY, EXPDEL, CUSTITEM, USAGE, " & _
              "ORNO, NEEDMAP, MAPREFFILE, " & _
              "PRO, OPPART, PLEN, PLENUN, PQTY, " & _
              "PQTYUN, EALEN, EALENUN, EAQTY, EAQTYUN, " & _
              "UPSLI, LOSLI, UPFIN, LOFIN, UPSTK, " & _
              "LOSTK, SPSPEC, " & _
              "SIZENO, ITEM, TATYPE, TAWIDTH, " & _
              "ECOLSEL, ECOL, CCOLSEL, ECOL, " & _
              "TACOL, TACOLNO, TAYCOLNO, " & _
              "TALCOL, TALCOLNO, TALYCOLNO, " & _
              "TARCOL, TARCOLNO, TARYCOLNO, " & _
              "THUPCOL, THUPCOLNO, THUPYCOLNO, " & _
              "THLUPCOL, THLUPCOLNO, THLUPYCOLNO, " & _
              "THRUPCOL, THRUPCOLNO, THRUPYCOLNO, " & _
              "THLOCOL, THLOCOLNO, THLOYCOLNO, " & _
              "THLLOCOL, THLLOCOLNO, THLLOYCOLNO, " & _
              "THRLOCOL, THRLOCOLNO, THRLOYCOLNO, " & _
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
              "CreateUser, CreateTime, ModifyUser, ModifyTime) " & _
        "VALUES("
        sql &= "'0' ,"
        sql &= "'" & NowDateTime & "' ,"
        sql &= "'" & wFormNo & "' ,"
        sql &= "'" & CStr(NewFormSno) & "' ,"
        '----------------------------------------------------------------------------------
        '--基本
        sql &= "'" & DRNO.Text & "' ,"
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
        sql &= "'" & DTHRUPCOL.SelectedValue & "' ,"
        sql &= "'" & DTHRUPCOLNO.Text & "' ,"
        sql &= "'" & DTHRUPYCOLNO.Text & "' ,"
        sql &= "'" & DTHLOCOL.SelectedValue & "' ,"
        sql &= "'" & DTHLOCOLNO.Text & "' ,"
        sql &= "'" & DTHLOYCOLNO.Text & "' ,"
        sql &= "'" & DTHLLOCOL.SelectedValue & "' ,"
        sql &= "'" & DTHLLOCOLNO.Text & "' ,"
        sql &= "'" & DTHLLOYCOLNO.Text & "' ,"
        sql &= "'" & DTHRLOCOL.SelectedValue & "' ,"
        sql &= "'" & DTHRLOCOLNO.Text & "' ,"
        sql &= "'" & DTHRLOYCOLNO.Text & "' ,"
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
        '----------------------------------------------------------------------------------
        sql &= "'" & Request.QueryString("pUserID") & "' ,"
        sql &= "'" & NowDateTime & "' ,"
        sql &= "'" & Request.QueryString("pUserID") & "' ,"
        sql &= "'" & NowDateTime & "' )"
        uDataBase.ExecuteNonQuery(sql)
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim Path As String = Server.MapPath(uCommon.GetAppSetting("COMMISSIONFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                     CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim sql As String
        '---------------------------------------------------------------------------------
        '-- 開發委託
        '---------------------------------------------------------------------------------
        sql = "Update F_DevelopmentCommissionSheet Set "
        If pFun <> "SAVE" Then
            sql &= " Sts = '" & pSts & "',"
            sql &= " CompletedTime = '" & NowDateTime & "',"
        End If
        '--基本
        sql &= " RNO = '" & DRNO.Text & "' ,"
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
        sql &= " THRUPCOL = '" & DTHRUPCOL.SelectedValue & "' ,"
        sql &= " THRUPCOLNO = '" & DTHRUPCOLNO.Text & "' ,"
        sql &= " THRUPYCOLNO = '" & DTHRUPYCOLNO.Text & "' ,"
        sql &= " THLOCOL = '" & DTHLOCOL.SelectedValue & "' ,"
        sql &= " THLOCOLNO = '" & DTHLOCOLNO.Text & "' ,"
        sql &= " THLOYCOLNO = '" & DTHLOYCOLNO.Text & "' ,"
        sql &= " THLOCOL = '" & DTHLOCOL.SelectedValue & "' ,"
        sql &= " THLOCOLNO = '" & DTHLOCOLNO.Text & "' ,"
        sql &= " THLOYCOLNO = '" & DTHLOYCOLNO.Text & "' ,"
        sql &= " THLLOCOL = '" & DTHLLOCOL.SelectedValue & "' ,"
        sql &= " THLLOCOLNO = '" & DTHLLOCOLNO.Text & "' ,"
        sql &= " THLLOYCOLNO = '" & DTHLLOYCOLNO.Text & "' ,"
        sql &= " THRLOCOL = '" & DTHRLOCOL.SelectedValue & "' ,"
        sql &= " THRLOCOLNO = '" & DTHRLOCOLNO.Text & "' ,"
        sql &= " THRLOYCOLNO = '" & DTHRLOYCOLNO.Text & "' ,"
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
        sql &= " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        sql &= " ModifyTime = '" & NowDateTime & "' "
        sql &= " Where FormNo  =  '" & wFormNo & "'"
        sql &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        uDataBase.ExecuteNonQuery(sql)
        '---------------------------------------------------------------------------------
        '-- 製造委託
        '---------------------------------------------------------------------------------
        sql = "Select * From ManufactureSheet "
        sql &= " Where RNO =  '" & DRNO.Text & "'"
        Dim dtManufactureSheet As DataTable = uDataBase.GetDataTable(sql)
        If dtManufactureSheet.Rows.Count > 0 Then
            sql = "Update ManufactureSheet Set "
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
            '--各工程遲納理由
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
            sql &= " Where RNO =  '" & DRNO.Text & "'"
            uDataBase.ExecuteNonQuery(sql)
        End If
        '---------------------------------------------------------------------------------
        '-- 開發見本
        '---------------------------------------------------------------------------------
        If D3APPBUYER.Text <> "" Then
            sql = "Select * From SampleSheet "
            sql &= " Where RNO =  '" & DRNO.Text & "'"
            Dim dtSampleSheet As DataTable = uDataBase.GetDataTable(sql)
            If dtSampleSheet.Rows.Count <= 0 Then
                sql = "Insert into SampleSheet "
                sql = sql + "( "
                sql = sql + "Date, AppBuyer, SizeNo, Item, CodeNo, "
                sql = sql + "SampleFile, TAWidth, DevNo, DevPrd, TACol, "
                sql = sql + "TALine, ECol, CCol, THCol, Other, "
                sql = sql + "QCFile1,QCFile2,QCFile3,QCFile4,QCFile5,TNLItem, TNRItem, TSLItem, TSRItem, "
                sql = sql + "TDLItem, TDRItem, CNITem, CSItem, CDItem, "
                sql = sql + "CItem, "
                sql = sql + "WF1, WF2, WF3, WF4, WF5 , WF6, WF7, "
                sql = sql + "WF3Name, WF4Name, WF5Name, WF6Name, WF7Name, "
                sql = sql + "Rno, "
                sql = sql + "CreateUser, CreateTime, ModifyUser, ModifyTime ,Other1,Other2,O1Item,O2Item "
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
                sql = sql + " '" + D3WF1.SelectedValue + "', "               '
                sql = sql + " '" + D3WF2.SelectedValue + "', "               '
                sql = sql + " '" + D3WF3.SelectedValue + "', "               '
                sql = sql + " '" + D3WF4.SelectedValue + "', "               '
                sql = sql + " '" + D3WF5.SelectedValue + "', "               '
                sql = sql + " '" + D3WF6.SelectedValue + "', "               '
                sql = sql + " '" + D3WF7.SelectedValue + "', "               '
                sql = sql + " '" + D3WF3Name.SelectedValue + "', "           '
                sql = sql + " '" + D3WF4Name.SelectedValue + "', "           '
                sql = sql + " '" + D3WF5Name.SelectedValue + "', "           '
                sql = sql + " '" + D3WF6Name.SelectedValue + "', "           '
                sql = sql + " '" + D3WF7Name.SelectedValue + "', "           '
                sql = sql + " '" + DRNO.Text + "', "                    '
                sql = sql + " '" + Request.QueryString("pUserID") + "', "  '作成者
                sql = sql + " '" + NowDateTime + "', "                      '作成時間
                sql = sql + " '" + "" + "', "                               '修改者
                sql = sql + " '" + NowDateTime + "', "                      '修改時間
                sql = sql + " '" + D31Other.Text + "', "                    'Other1
                sql = sql + " '" + D32Other.Text + "', "                    'Other2
                sql = sql + " '" + D3O1ITEM.Text + "', "                    'Item1
                sql = sql + " '" + D3O2ITEM.Text + "' "                     'Item2
                sql = sql + " ) "
                uDataBase.ExecuteNonQuery(sql)
            End If
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
        If pFun <> "SAVE" Then      '<> Save
            SQl = "Update T_WaitHandle Set "
            SQl = SQl + " Active = '" & "0" & "',"
            SQl = SQl + " Sts = '" & pSts & "',"
            If pSts = "1" Then SQl = SQl + " StsDes = '" & BOK.Text & "',"
            If pSts = "2" Then SQl = SQl + " StsDes = '" & BNG1.Text & "',"
            If pSts = "3" Then SQl = SQl + " StsDes = '" & BNG2.Text & "',"

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
        Dim dtCommissionNo As DataTable = uDataBase.GetDataTable(SQl)

        If dtCommissionNo.Rows.Count <= 0 Then
            If DRNO.Text <> "" Then
                SQl = "Insert into T_CommissionNo "
                SQl = SQl + "( "
                SQl = SQl + "FormNo, FormSno, No, MapNo, CreateUser, CreateTime "    '1~5
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '" + pFormNo + "', "
                SQl = SQl + " '" + CStr(pFormSno) + "', "
                SQl = SQl + " '" + DRNO.Text + "', "
                SQl = SQl + " '" + "" + "', "
                SQl = SQl + " '" + Request.QueryString("pUserID") + "', "
                SQl = SQl + " '" + NowDateTime + "' "
                SQl = SQl + " ) "
                uDataBase.ExecuteNonQuery(SQl)
            End If
        Else
            If DRNO.Text <> "" Then
                If DRNO.Text <> dtCommissionNo.Rows(0)("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DRNO.Text & "',"
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
    '**
    '**     檢查上傳檔案
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As FileUpload) As Integer
        Dim fileExtension As String     '宣告一個變數存放檔案格式(副檔名)
        Dim allowedExtensions As String() = {".jpg", ".jpeg", ".gif", ".pdf", ".xls", ".doc", ".ppt"}   '定義允許的檔案格式
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
        SetImageButtonImageFile(0)
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton2 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton2.Click
        SetImageButtonImageFile(1)
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton3 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton3_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton3.Click
        SetImageButtonImageFile(2)
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton4 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton4_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton4.Click
        SetImageButtonImageFile(3)
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton5 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton5_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton5.Click
        SetImageButtonImageFile(4)
    End Sub
End Class
