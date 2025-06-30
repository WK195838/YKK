Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Collections.Generic

Partial Class SampleSheet_01
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
            ShowSheetField("IsPostBack")            '表單欄位顯示及欄位輸入檢查
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
        Response.Cookies("PGM").Value = "SampleSheet_01.aspx"                                       '程式名
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
        BDataPump.Attributes("onclick") = "window.open('OpenNoPickerForSampleSheet.aspx','NoPicker','status=0,toolbar=0,width=620,height=650,resizable=yes,scrollbars=yes');"    '選取參考開發資料
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
        'No
        Select Case FindFieldInf("NO")
            Case 0  '顯示
                DNO.BackColor = Color.White
                DNO.Visible = True
                DNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DNO.Visible = True
                DNO.BackColor = Color.GreenYellow
                DNO.ReadOnly = False
                ShowRequiredFieldValidator("DNORqd", "DNO", "異常：NO")
            Case 2  '修改
                DNO.Visible = True
                DNO.BackColor = Color.Yellow
                DNO.ReadOnly = False
            Case Else   '隱藏
                DNO.Visible = False
        End Select
        If pPost = "New" Then DNO.Text = ""
        'DataPump
        Select Case FindFieldInf("DATAPUMP")
            Case 0  '顯示
                BDataPump.Visible = True
            Case Else   '隱藏
                BDataPump.Visible = False
        End Select
        If pPost = "New" Then DNO.Text = ""
        '客戶
        DAPPBUYER.BackColor = Color.LightGray
        DAPPBUYER.Visible = True
        DAPPBUYER.Attributes.Add("readonly", "true")
        '發行日
        DDATE.BackColor = Color.LightGray
        DDATE.Visible = True
        DDATE.Attributes.Add("readonly", "true")
        If pPost = "New" Then DDATE.Text = CDate(NowDateTime).ToString("yyyy/MM/dd")
        'SIZE
        DSIZENO.BackColor = Color.LightGray
        DSIZENO.Visible = True
        DSIZENO.Attributes.Add("readonly", "true")
        'ITEM
        DITEM.BackColor = Color.LightGray
        DITEM.Visible = True
        DITEM.Attributes.Add("readonly", "true")
        'TAPE(CODENO)
        DCODENO.BackColor = Color.LightGray
        DCODENO.Visible = True
        DCODENO.Attributes.Add("readonly", "true")
        '樣品圖-表
        Select Case FindFieldInf("SAMPLEFILE-1")
            Case 0  '顯示
                DSAMPLEFILE1.Visible = False
                DSAMPLEFILE1.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DSAMPLEFILE1Rqd", "DSAMPLEFILE1", "異常：需輸入樣品圖-表")
                DSAMPLEFILE1.Visible = True
                DSAMPLEFILE1.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DSAMPLEFILE1.Visible = True
                DSAMPLEFILE1.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DSAMPLEFILE1.Visible = False
        End Select
        If pPost = "New" Then LSAMPLEFILE1.Visible = False
        '樣品圖-裏
        Select Case FindFieldInf("SAMPLEFILE-2")
            Case 0  '顯示
                DSAMPLEFILE2.Visible = False
                DSAMPLEFILE2.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DSAMPLEFILE2Rqd", "DSAMPLEFILE2", "異常：需輸入樣品圖-裏")
                DSAMPLEFILE2.Visible = True
                DSAMPLEFILE2.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DSAMPLEFILE2.Visible = True
                DSAMPLEFILE2.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DSAMPLEFILE2.Visible = False
        End Select
        If pPost = "New" Then LSAMPLEFILE2.Visible = False
        '布帶寬度
        DTAWIDTH.BackColor = Color.LightGray
        DTAWIDTH.Visible = True
        DTAWIDTH.Attributes.Add("readonly", "true")
        '開發NO
        DDEVNO.BackColor = Color.LightGray
        DDEVNO.Visible = True
        DDEVNO.Attributes.Add("readonly", "true")
        '開發期間
        DDEVPRD.BackColor = Color.LightGray
        DDEVPRD.Visible = True
        DDEVPRD.Attributes.Add("readonly", "true")
        '布帶
        DTACOL.BackColor = Color.LightGray
        DTACOL.Visible = True
        DTACOL.Attributes.Add("readonly", "true")
        '條紋線
        DTALINE.BackColor = Color.LightGray
        DTALINE.Visible = True
        DTALINE.Attributes.Add("readonly", "true")
        '務齒
        DECOL.BackColor = Color.LightGray
        DECOL.Visible = True
        DECOL.Attributes.Add("readonly", "true")
        '丸紐
        DCCOL.BackColor = Color.LightGray
        DCCOL.Visible = True
        DCCOL.Attributes.Add("readonly", "true")
        '縫工上線
        DTHUPCOL.BackColor = Color.LightGray
        DTHUPCOL.Visible = True
        DTHUPCOL.Attributes.Add("readonly", "true")
        '縫工下線
        DTHLOCOL.BackColor = Color.LightGray
        DTHLOCOL.Visible = True
        DTHLOCOL.Attributes.Add("readonly", "true")
        '生產注意事項-工程-1
        Select Case FindFieldInf("MANUFOP-1")
            Case 0  '顯示
                DMOP1.BackColor = Color.LightGray
                DMOP1.Visible = True
                DMOP1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DMOP1.Visible = True
                DMOP1.BackColor = Color.GreenYellow
                DMOP1.ReadOnly = False
                ShowRequiredFieldValidator("DMOP1Rqd", "DMOP1", "異常：生產注意事項-工程-1")
            Case 2  '修改
                DMOP1.Visible = True
                DMOP1.BackColor = Color.Yellow
                DMOP1.ReadOnly = False
            Case Else   '隱藏
                DMOP1.Visible = False
        End Select
        If pPost = "New" Then DMOP1.Text = ""
        '生產注意事項-工程-2
        Select Case FindFieldInf("MANUFOP-2")
            Case 0  '顯示
                DMOP2.BackColor = Color.LightGray
                DMOP2.Visible = True
                DMOP2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DMOP2.Visible = True
                DMOP2.BackColor = Color.GreenYellow
                DMOP2.ReadOnly = False
                ShowRequiredFieldValidator("DMOP2Rqd", "DMOP2", "異常：生產注意事項-工程-2")
            Case 2  '修改
                DMOP2.Visible = True
                DMOP2.BackColor = Color.Yellow
                DMOP2.ReadOnly = False
            Case Else   '隱藏
                DMOP2.Visible = False
        End Select
        If pPost = "New" Then DMOP2.Text = ""
        '生產注意事項-工程-3
        Select Case FindFieldInf("MANUFOP-3")
            Case 0  '顯示
                DMOP3.BackColor = Color.LightGray
                DMOP3.Visible = True
                DMOP3.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DMOP3.Visible = True
                DMOP3.BackColor = Color.GreenYellow
                DMOP3.ReadOnly = False
                ShowRequiredFieldValidator("DMOP3Rqd", "DMOP3", "異常：生產注意事項-工程-3")
            Case 2  '修改
                DMOP3.Visible = True
                DMOP3.BackColor = Color.Yellow
                DMOP3.ReadOnly = False
            Case Else   '隱藏
                DMOP3.Visible = False
        End Select
        If pPost = "New" Then DMOP3.Text = ""
        '生產注意事項-工程-4
        Select Case FindFieldInf("MANUFOP-4")
            Case 0  '顯示
                DMOP4.BackColor = Color.LightGray
                DMOP4.Visible = True
                DMOP4.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DMOP4.Visible = True
                DMOP4.BackColor = Color.GreenYellow
                DMOP4.ReadOnly = False
                ShowRequiredFieldValidator("DMOP4Rqd", "DMOP4", "異常：生產注意事項-工程-4")
            Case 2  '修改
                DMOP4.Visible = True
                DMOP4.BackColor = Color.Yellow
                DMOP4.ReadOnly = False
            Case Else   '隱藏
                DMOP4.Visible = False
        End Select
        If pPost = "New" Then DMOP4.Text = ""
        '生產注意事項-說明-1
        Select Case FindFieldInf("MANUFNOTE-1")
            Case 0  '顯示
                DMNote1.BackColor = Color.LightGray
                DMNote1.Visible = True
                DMNote1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DMNote1.Visible = True
                DMNote1.BackColor = Color.GreenYellow
                DMNote1.ReadOnly = False
                ShowRequiredFieldValidator("DMNote1Rqd", "DMNote1", "異常：生產注意事項-說明-1")
            Case 2  '修改
                DMNote1.Visible = True
                DMNote1.BackColor = Color.Yellow
                DMNote1.ReadOnly = False
            Case Else   '隱藏
                DMNote1.Visible = False
        End Select
        If pPost = "New" Then DMNote1.Text = ""
        '生產注意事項-說明-2
        Select Case FindFieldInf("MANUFNOTE-2")
            Case 0  '顯示
                DMNote2.BackColor = Color.LightGray
                DMNote2.Visible = True
                DMNote2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DMNote2.Visible = True
                DMNote2.BackColor = Color.GreenYellow
                DMNote2.ReadOnly = False
                ShowRequiredFieldValidator("DMNote2Rqd", "DMNote2", "異常：生產注意事項-說明-2")
            Case 2  '修改
                DMNote2.Visible = True
                DMNote2.BackColor = Color.Yellow
                DMNote2.ReadOnly = False
            Case Else   '隱藏
                DMNote2.Visible = False
        End Select
        If pPost = "New" Then DMNote2.Text = ""
        '生產注意事項-說明-3
        Select Case FindFieldInf("MANUFNOTE-3")
            Case 0  '顯示
                DMNote3.BackColor = Color.LightGray
                DMNote3.Visible = True
                DMNote3.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DMNote3.Visible = True
                DMNote3.BackColor = Color.GreenYellow
                DMNote3.ReadOnly = False
                ShowRequiredFieldValidator("DMNote3Rqd", "DMNote3", "異常：生產注意事項-說明-3")
            Case 2  '修改
                DMNote3.Visible = True
                DMNote3.BackColor = Color.Yellow
                DMNote3.ReadOnly = False
            Case Else   '隱藏
                DMNote3.Visible = False
        End Select
        If pPost = "New" Then DMNote3.Text = ""
        '生產注意事項-說明-4
        Select Case FindFieldInf("MANUFNOTE-4")
            Case 0  '顯示
                DMNote4.BackColor = Color.LightGray
                DMNote4.Visible = True
                DMNote4.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DMNote4.Visible = True
                DMNote4.BackColor = Color.GreenYellow
                DMNote4.ReadOnly = False
                ShowRequiredFieldValidator("DMNote4Rqd", "DMNote4", "異常：生產注意事項-說明-4")
            Case 2  '修改
                DMNote4.Visible = True
                DMNote4.BackColor = Color.Yellow
                DMNote4.ReadOnly = False
            Case Else   '隱藏
                DMNote4.Visible = False
        End Select
        If pPost = "New" Then DMNote4.Text = ""
        'WAVE'S
        DTNLITEM.BackColor = Color.LightGray
        DTNLITEM.Visible = True
        DTNLITEM.Attributes.Add("readonly", "true")
        '
        DTNRITEM.BackColor = Color.LightGray
        DTNRITEM.Visible = True
        DTNRITEM.Attributes.Add("readonly", "true")
        '
        DTSLITEM.BackColor = Color.LightGray
        DTSLITEM.Visible = True
        DTSLITEM.Attributes.Add("readonly", "true")
        '
        DTSRITEM.BackColor = Color.LightGray
        DTSRITEM.Visible = True
        DTSRITEM.Attributes.Add("readonly", "true")
        '
        DTDLITEM.BackColor = Color.LightGray
        DTDLITEM.Visible = True
        DTDLITEM.Attributes.Add("readonly", "true")
        '
        DTDRITEM.BackColor = Color.LightGray
        DTDRITEM.Visible = True
        DTDRITEM.Attributes.Add("readonly", "true")
        '
        DCNITEM.BackColor = Color.LightGray
        DCNITEM.Visible = True
        DCNITEM.Attributes.Add("readonly", "true")
        '
        DCSITEM.BackColor = Color.LightGray
        DCSITEM.Visible = True
        DCSITEM.Attributes.Add("readonly", "true")
        '
        DCDITEM.BackColor = Color.LightGray
        DCDITEM.Visible = True
        DCDITEM.Attributes.Add("readonly", "true")
        '
        DCITEM.BackColor = Color.LightGray
        DCITEM.Visible = True
        DCITEM.Attributes.Add("readonly", "true")
        '
        'DODESCR1.BackColor = Color.LightGray
        DODESCR1.Visible = True
        DODESCR1.Attributes.Add("readonly", "true")
        '
        DO1ITEM.BackColor = Color.LightGray
        DO1ITEM.Visible = True
        DO1ITEM.Attributes.Add("readonly", "true")
        '
        'DODESCR2.BackColor = Color.LightGray
        DODESCR2.Visible = True
        DODESCR2.Attributes.Add("readonly", "true")
        '
        DO2ITEM.BackColor = Color.LightGray
        DO2ITEM.Visible = True
        DO2ITEM.Attributes.Add("readonly", "true")
        '
        DOP1.BackColor = Color.LightGray
        DOP1.Visible = True
        DOP1.Attributes.Add("readonly", "true")
        '
        DOP2.BackColor = Color.LightGray
        DOP2.Visible = True
        DOP2.Attributes.Add("readonly", "true")
        '
        DOP3.BackColor = Color.LightGray
        DOP3.Visible = True
        DOP3.Attributes.Add("readonly", "true")
        '
        DOP4.BackColor = Color.LightGray
        DOP4.Visible = True
        DOP4.Attributes.Add("readonly", "true")
        '
        DOP5.BackColor = Color.LightGray
        DOP5.Visible = True
        DOP5.Attributes.Add("readonly", "true")
        '
        DOP6.BackColor = Color.LightGray
        DOP6.Visible = True
        DOP6.Attributes.Add("readonly", "true")
    End Sub
    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pIdx As Integer, ByVal pFieldName As String, ByVal pName As String)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetFunction)
    '**     表單功能按鈕顯示
    '**
    '*****************************************************************
    Sub ShowSheetFunction()
        Dim sql As String = ""
        Top = 1150
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
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("SampleFilePath")
        Dim SQL As String
        SQL = "Select * From F_FactorySampleSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtSampleSheet As DataTable = uDataBase.GetDataTable(SQL)
        If dtSampleSheet.Rows.Count > 0 Then
            '----基本欄位設定-------------------------------------------------
            DNO.Text = dtSampleSheet.Rows(0).Item("No")                              'No
            DAPPBUYER.Text = dtSampleSheet.Rows(0).Item("AppBuyer")                 'Customer
            DDATE.Text = dtSampleSheet.Rows(0).Item("Date")                         '發行日
            DSIZENO.Text = dtSampleSheet.Rows(0).Item("SizeNo")                     'Size
            DITEM.Text = dtSampleSheet.Rows(0).Item("Item")                         'Item
            DCODENO.Text = dtSampleSheet.Rows(0).Item("CodeNo")                     'Code No
            '實際樣品-表
            If dtSampleSheet.Rows(0).Item("SampleFile1") <> "" Then
                LSAMPLEFILE1.ImageUrl = Path & dtSampleSheet.Rows(0).Item("SampleFile1")
                LSAMPLEFILE1.Visible = True
            End If
            '實際樣品-裏
            If dtSampleSheet.Rows(0).Item("SampleFile2") <> "" Then
                LSAMPLEFILE2.ImageUrl = Path & dtSampleSheet.Rows(0).Item("SampleFile2")
                LSAMPLEFILE2.Visible = True
            End If
            '----開發規格-------------------------------------------------
            DTAWIDTH.Text = dtSampleSheet.Rows(0).Item("TAWidth")                   '布帶寬度
            DDEVNO.Text = dtSampleSheet.Rows(0).Item("DevNo")                       '開發No
            DDEVPRD.Text = dtSampleSheet.Rows(0).Item("DevPrd")                     '開發期間
            DTACOL.Text = dtSampleSheet.Rows(0).Item("TACol")                       '布帶Color
            DTALINE.Text = dtSampleSheet.Rows(0).Item("TALine")                     '條紋線Color
            DECOL.Text = dtSampleSheet.Rows(0).Item("ECol")                         '務齒
            DCCOL.Text = dtSampleSheet.Rows(0).Item("CCol")                         '丸紐
            DTHUPCOL.Text = dtSampleSheet.Rows(0).Item("THUPCol")                   '縫工上線
            DTHLOCOL.Text = dtSampleSheet.Rows(0).Item("THLOCol")                   '縫工下線
            '----生產注意-------------------------------------------------
            DMOP1.Text = dtSampleSheet.Rows(0).Item("MOP1")                         '工程-1
            DMOP2.Text = dtSampleSheet.Rows(0).Item("MOP2")                         '工程-2
            DMOP3.Text = dtSampleSheet.Rows(0).Item("MOP3")                         '工程-3
            DMOP4.Text = dtSampleSheet.Rows(0).Item("MOP4")                         '工程-4
            DMNote1.Text = dtSampleSheet.Rows(0).Item("MNote1")                     '說明-1
            DMNote2.Text = dtSampleSheet.Rows(0).Item("MNote2")                     '說明-2
            DMNote3.Text = dtSampleSheet.Rows(0).Item("MNote3")                     '說明-3
            DMNote4.Text = dtSampleSheet.Rows(0).Item("MNote4")                     '說明-4
            '----Wave's-------------------------------------------------
            DTNLITEM.Text = dtSampleSheet.Rows(0).Item("TNLItem")                   'TAPE NAT-左
            DTNRITEM.Text = dtSampleSheet.Rows(0).Item("TNRItem")                   'TAPE NAT-右
            DTSLITEM.Text = dtSampleSheet.Rows(0).Item("TSLItem")                   'TAPE SET-左
            DTSRITEM.Text = dtSampleSheet.Rows(0).Item("TSRItem")                   'TAPE SET-右
            DTDLITEM.Text = dtSampleSheet.Rows(0).Item("TDLItem")                   'TAPE DYED-左
            DTDRITEM.Text = dtSampleSheet.Rows(0).Item("TDRItem")                   'TAPE DYED-右
            DCNITEM.Text = dtSampleSheet.Rows(0).Item("CNItem")                     'CHAIN NAT
            DCSITEM.Text = dtSampleSheet.Rows(0).Item("CSItem")                     'CHAIN SET
            DCDITEM.Text = dtSampleSheet.Rows(0).Item("CDItem")                     'CHAIN DYED
            DODESCR1.Text = dtSampleSheet.Rows(0).Item("Other1")                    'Other1
            DODESCR2.Text = dtSampleSheet.Rows(0).Item("Other2")                    'Other2
            DO1ITEM.Text = dtSampleSheet.Rows(0).Item("O1Item")                     'O1Item
            DO2ITEM.Text = dtSampleSheet.Rows(0).Item("O2Item")                     'O2Item
            DCITEM.Text = dtSampleSheet.Rows(0).Item("CItem")                       'CORD
            '----FLOW-------------------------------------------------
            DOP1.Text = dtSampleSheet.Rows(0).Item("OP1")                           'OP1
            DOP2.Text = dtSampleSheet.Rows(0).Item("OP2")                           'OP2
            DOP3.Text = dtSampleSheet.Rows(0).Item("OP3")                           'OP3
            DOP4.Text = dtSampleSheet.Rows(0).Item("OP4")                           'OP4
            DOP5.Text = dtSampleSheet.Rows(0).Item("OP5")                           'OP5
            DOP6.Text = dtSampleSheet.Rows(0).Item("OP6")                           'OP6
            '-----------------------------------------------------------------
            '-- 核定履歷
            '-----------------------------------------------------------------
            SQL = "SELECT "
            SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
            SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
            SQL = SQL + "'預定：[' + BStartTimeDesc + ' ~ ' + BEndTimeDesc + '], ' + "
            SQL = SQL + "'實際：[' + AStartTimeDesc + ' ~ ' + AEndTimeDesc + '] ' As Description, "
            SQL = SQL + "URL "
            SQL = SQL + "FROM V_WaitHandle_01 "
            SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
            SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
            'SQL = SQL + "Order by Unique_ID Desc "
            SQL = SQL + "Order by CreateTime Desc, Step Desc, SeqNo Desc "
            GridView1.DataSource = uDataBase.GetDataTable(SQL)
            GridView1.DataBind()
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
        '----表單-附加檔案-----------------------------------------------
        '樣品圖-表
        If ErrCode = 0 Then
            If DSAMPLEFILE1.Visible Then
                If DSAMPLEFILE1.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DSAMPLEFILE1)
                End If
            End If
        End If
        '樣品圖-裏
        If ErrCode = 0 Then
            If DSAMPLEFILE2.Visible Then
                If DSAMPLEFILE2.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DSAMPLEFILE2)
                End If
            End If
        End If
        If pAction <> 3 Then    '不是儲存
            '----生產注意-------------------------------------------------------
            '生產工程
            If ErrCode = 0 Then
                If DMOP1.Text = "" And DMOP2.Text = "" And DMOP3.Text = "" And DMOP4.Text = "" Then ErrCode = 9010
            End If
            '注意內容
            If ErrCode = 0 Then
                If DMNote1.Text = "" And DMNote2.Text = "" And DMNote3.Text = "" And DMNote4.Text = "" Then ErrCode = 9040
            End If
            '----系統檢查-------------------------------------------------------
            '核定說明
            If ErrCode = 0 Then
                If DDecideDesc.Visible = True Then
                    If DDecideDesc.Text = "" Then ErrCode = 9060
                End If
            End If
            '檢查委託書No
            If ErrCode = 0 Then
                If DNO.Text <> "" Then
                    ErrCode = oCommon.CommissionNo("002003", wFormSno, wStep, DNO.Text) '表單號碼, 表單流水號, 工程, 委託書No
                    If ErrCode <> 0 Then
                        ErrCode = 9050
                    End If
                End If
            End If

        End If
        '----異常訊息處理----------------------------------------------------
        'ErrCode = 9999
        If ErrCode <> 0 Then
            If ErrCode = 9010 Then Message = "生產注意內容需輸入,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案格式有誤,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案Size有誤,請確認!"
            If ErrCode = 9040 Then Message = "生產注意內容需輸入,請確認!"
            If ErrCode = 9050 Then Message = "委託書No.重覆,請確認委託書No.!"
            If ErrCode = 9060 Then Message = "核定說明需輸入,請確認!"
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
        Dim wQCLT As Integer = 0                    'QC-L/T
        Dim Run As Boolean = True                   '是否執行
        Dim RepeatRun As Boolean = False            '是否重覆執行
        Dim MultiJob As Integer = 0                 '多人核定
        Dim wLevel As String = ""                   '難易度
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
                    End If
                End If
            End If

            '取得下一關
            If ErrCode = 0 And pRunNextStep = 1 Then
                Dim SQL As String = ""
                Dim wAllocateID As String = ""
                '--指定簽核者處分---------------------------------------------------------------
                '
                '取得簽核者
                '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
                '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 
                RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, wAgentID, wAllocateID, MultiJob, _
                                              pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
                '
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
                        ModifyData(pFun, "0")                       '更新表單資料 Sts=0(未結)
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
        Dim Path As String = Server.MapPath(uCommon.GetAppSetting("UploadSampleFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        '
        Dim sql As String = ""
        sql = "Insert into F_FactorySampleSheet "
        sql &= "( "
        sql &= "Sts, CompletedTime, FormNo, FormSno, "
        sql &= "No, AppBuyer, Date, SizeNo, Item, "
        sql &= "CodeNo, SampleFile1, SampleFile2, TAWidth, DevNo, "
        sql &= "DevPrd, TACol, TALine, ECol, CCol, "
        sql &= "THUPCol, THLOCol, "
        sql &= "MOP1, MOP2, MOP3, MOP4, "
        sql &= "MNote1, MNote2, MNote3, MNote4, "
        sql &= "TNLItem, TNRItem, TSLItem, TSRItem, TDLItem, "
        sql &= "TDRItem, CNITem, CSItem, CDItem, CItem, "
        sql &= "Other1, Other2, O1Item, O2Item, "
        sql &= "OP1, OP2, OP3, OP4, OP5, OP6, "
        sql &= "ONo, "
        sql &= "CreateUser, CreateTime, ModifyUser, ModifyTime "
        sql &= ")  "
        sql &= "VALUES ( "
        '----系統------------------------------------------------------------------------------
        sql &= "'0' ,"                              '狀態
        sql &= "'" & NowDateTime & "' ,"            '日時
        sql &= "'" & wFormNo & "' ,"                '表單No
        sql &= "'" & CStr(NewFormSno) & "' ,"       '表單流水號
        '----表單-基本--------------------------------------------------------------------------
        sql &= "'" & DNO.Text & "' ,"               'NO
        sql &= " '" + DAPPBUYER.Text + "', "        '客戶
        sql &= " '" + DDATE.Text + "', "            '發行日期
        sql &= " '" + DSIZENO.Text + "', "          'Size
        sql &= " '" + DITEM.Text + "', "            'Item
        sql &= " '" + DCODENO.Text + "', "          'Code-No
        '----表單-樣品檔--------------------------------------------------------------------------
        Dim FileName As String = ""
        If DSAMPLEFILE1.Visible Then
            If DSAMPLEFILE1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(wFormSno) & "-" & "SampleFile1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSAMPLEFILE1.PostedFile.FileName)
                Try    '上傳圖檔
                    DSAMPLEFILE1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        sql &= " '" + FileName + "', "              '樣品檔-左
        FileName = ""
        If DSAMPLEFILE2.Visible Then
            If DSAMPLEFILE2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(wFormSno) & "-" & "SampleFile2" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSAMPLEFILE2.PostedFile.FileName)
                Try    '上傳圖檔
                    DSAMPLEFILE2.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        sql &= " '" + FileName + "', "              '樣品檔-右
        '----表單-開發規格---------------------------------------------------------------------
        sql &= " '" + DTAWIDTH.Text + "', "         '布帶寬度
        sql &= " '" + DDEVNO.Text + "', "           '開發No
        sql &= " '" + DDEVPRD.Text + "', "          '開發期間
        sql &= " '" + DTACOL.Text + "', "           '布帶
        sql &= " '" + DTALINE.Text + "', "          '條紋線
        sql &= " '" + DECOL.Text + "', "            '務齒
        sql &= " '" + DCCOL.Text + "', "            '丸紐
        sql &= " '" + DTHUPCOL.Text + "', "         '縫工上線
        sql &= " '" + DTHLOCOL.Text + "', "         '縫工下線
        '----表單-生產注意---------------------------------------------------------------------
        sql &= " '" + DMOP1.Text + "', "            '工程-1
        sql &= " '" + DMOP2.Text + "', "            '工程-2
        sql &= " '" + DMOP3.Text + "', "            '工程-3
        sql &= " '" + DMOP4.Text + "', "            '工程-4
        sql &= " '" + DMNote1.Text + "', "          '說明-1
        sql &= " '" + DMNote2.Text + "', "          '說明-2
        sql &= " '" + DMNote3.Text + "', "          '說明-3
        sql &= " '" + DMNote4.Text + "', "          '說明-4
        '----表單-Wave's-------------------------------------------------
        sql &= " '" + DTNLITEM.Text + "', "         'TAPE NAT-左
        sql &= " '" + DTNRITEM.Text + "', "         'TAPE NAT-右
        sql &= " '" + DTSLITEM.Text + "', "         'TAPE SET-左
        sql &= " '" + DTSRITEM.Text + "', "         'TAPE SET-右
        sql &= " '" + DTDLITEM.Text + "', "         'TAPE DYED-左
        sql &= " '" + DTDRITEM.Text + "', "         'TAPE DYED-右
        sql &= " '" + DCNITEM.Text + "', "          'CHAIN NAT
        sql &= " '" + DCSITEM.Text + "', "          'CHAIN SET
        sql &= " '" + DCDITEM.Text + "', "          'CHAIN DYED
        sql &= " '" + DCITEM.Text + "', "           'CORD
        sql &= " '" + DODESCR1.Text + "', "         'Other1
        sql &= " '" + DODESCR2.Text + "', "         'Other2
        sql &= " '" + DO1ITEM.Text + "', "          'Item1
        sql &= " '" + DO2ITEM.Text + "', "          'Item2
        '----表單-OP-------------------------------------------------
        sql &= " '" + DOP1.Text + "', "             'OP1
        sql &= " '" + DOP2.Text + "', "             'OP2
        sql &= " '" + DOP3.Text + "', "             'OP3
        sql &= " '" + DOP4.Text + "', "             'OP4
        sql &= " '" + DOP5.Text + "', "             'OP5
        sql &= " '" + DOP6.Text + "', "             'OP6
        '----連結------------------------------------------------------------------------------
        sql &= "'" & DONO.Text & "' ,"              '原NO
        '----系統------------------------------------------------------------------------------
        sql &= " '" + Request.QueryString("pUserID") + "', "  '作成者
        sql &= " '" + NowDateTime + "', "                      '作成時間
        sql &= " '" + "" + "', "                               '修改者
        sql &= " '" + NowDateTime + "' "                       '修改時間
        sql &= " ) "
        uDataBase.ExecuteNonQuery(sql)
        '更新原開發委託書Status
        UpdateCommissionSheet(DONO.Text)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim Path As String = Server.MapPath(uCommon.GetAppSetting("UploadSampleFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                     CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim sql As String
        sql = "Update F_FactorySampleSheet Set "
        If pFun <> "SAVE" Then
            sql &= " Sts = '" & pSts & "',"
            sql &= " CompletedTime = '" & NowDateTime & "',"
        End If
        '----表單-基本--------------------------------------------------------------------------
        sql &= " DATE = '" & DDATE.Text & "' ,"
        sql &= " APPBUYER = '" & DAPPBUYER.Text & "' ,"
        sql &= " SIZENO = '" & DSIZENO.Text & "' ,"
        sql &= " ITEM = '" & DITEM.Text & "' ,"
        sql &= " CODENO = '" & DCODENO.Text & "' ,"
        '----表單-樣品檔--------------------------------------------------------------------------
        '樣品檔-左
        If DSAMPLEFILE1.Visible Then
            If DSAMPLEFILE1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "SampleFile1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSAMPLEFILE1.PostedFile.FileName)
                Try    '上傳圖檔
                    DSAMPLEFILE1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " SAMPLEFILE1 = '" + FileName + "', "
            End If
        End If
        '樣品檔-右
        If DSAMPLEFILE2.Visible Then
            If DSAMPLEFILE2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                Dim FileName As String = CStr(wFormSno) & "-" & "SampleFile2" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSAMPLEFILE2.PostedFile.FileName)
                Try    '上傳圖檔
                    DSAMPLEFILE2.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                sql = sql + " SAMPLEFILE2 = '" + FileName + "', "
            End If
        End If
        '----表單-開發規格---------------------------------------------------------------------
        sql &= " TAWIDTH = '" & DTAWIDTH.Text & "' ,"
        sql &= " DEVNO = '" & DDEVNO.Text & "' ,"
        sql &= " DEVPRD = '" & DDEVPRD.Text & "' ,"
        sql &= " TACOL = '" & DTACOL.Text & "' ,"
        sql &= " TALINE = '" & DTALINE.Text & "' ,"
        sql &= " ECOL = '" & DECOL.Text & "' ,"
        sql &= " CCOL = '" & DCCOL.Text & "' ,"
        sql &= " THUPCOL = '" & DTHUPCOL.Text & "' ,"
        sql &= " THLOCOL = '" & DTHLOCOL.Text & "' ,"
        '----表單-生產注意---------------------------------------------------------------------
        sql &= " MOP1 = '" & DMOP1.Text & "' ,"
        sql &= " MOP2 = '" & DMOP2.Text & "' ,"
        sql &= " MOP3 = '" & DMOP3.Text & "' ,"
        sql &= " MOP4 = '" & DMOP4.Text & "' ,"
        sql &= " MNote1 = '" & DMNote1.Text & "' ,"
        sql &= " MNote2 = '" & DMNote2.Text & "' ,"
        sql &= " MNote3 = '" & DMNote3.Text & "' ,"
        sql &= " MNote4 = '" & DMNote4.Text & "' ,"
        '----表單-Wave's-------------------------------------------------
        sql &= " TNLITEM = '" & DTNLITEM.Text & "' ,"
        sql &= " TNRITEM = '" & DTNRITEM.Text & "' ,"
        sql &= " TSLITEM = '" & DTSLITEM.Text & "' ,"
        sql &= " TSRITEM = '" & DTSRITEM.Text & "' ,"
        sql &= " TDLITEM = '" & DTDLITEM.Text & "' ,"
        sql &= " TDRITEM = '" & DTDRITEM.Text & "' ,"
        sql &= " CNITEM = '" & DCNITEM.Text & "' ,"
        sql &= " CSITEM = '" & DCSITEM.Text & "' ,"
        sql &= " CDITEM = '" & DCDITEM.Text & "' ,"
        sql &= " CITEM = '" & DCITEM.Text & "' ,"
        sql &= " Other1 = '" & DODESCR1.Text & "' ,"
        sql &= " Other2 = '" & DODESCR2.Text & "' ,"
        sql &= " O1Item = '" & DO1ITEM.Text & "' ,"
        sql &= " O2Item = '" & DO2ITEM.Text & "' ,"
        '----表單-OP-------------------------------------------------
        sql &= " OP1 = '" & DOP1.Text & "' ,"
        sql &= " OP2 = '" & DOP2.Text & "' ,"
        sql &= " OP3 = '" & DOP3.Text & "' ,"
        sql &= " OP4 = '" & DOP4.Text & "' ,"
        sql &= " OP5 = '" & DOP5.Text & "' ,"
        sql &= " OP6 = '" & DOP6.Text & "' ,"
        '----------------------------------------------------------------------------------
        sql &= " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        sql &= " ModifyTime = '" & NowDateTime & "' "
        sql &= " Where FormNo  =  '" & wFormNo & "'"
        sql &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        uDataBase.ExecuteNonQuery(sql)
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
    '**(GridView1_RowDataBound)
    '**     延遲處理
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim Pos As Integer = InStr(e.Row.Cells(7).Text.ToString, "],")
            Dim Str1 As String = Mid(e.Row.Cells(7).Text.ToString, 1, Pos)
            Dim Str2 As String = Mid(e.Row.Cells(7).Text.ToString, Pos + 3, Len(e.Row.Cells(7).Text.ToString))
            '
            e.Row.Cells(7).Text = Str1 + "<br/>" + Str2
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
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(UpdateCommissionSheet)
    '**     更新原開發委託書Status
    '**
    '*****************************************************************
    Protected Sub UpdateCommissionSheet(ByVal pNo As String)
        Dim SQL As String
        Try
            SQL = "Update F_CommissionSheet Set "
            SQL = SQL + " Sts_FactorySample = '1', "
            SQL = SQL + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
            SQL = SQL + " Where No = '" & pNo & "'"
            uDataBase.ExecuteNonQuery(SQL)
        Catch ex As Exception
        End Try
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
End Class
