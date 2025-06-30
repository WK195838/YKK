Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb

Partial Class SPD_YKKGroupCopySheet_01
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
    Dim wNo As String = ""          '表單-委託單No
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    Dim wUserName As String = ""    '姓名代理用

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900 ' 設定逾時時間
        SetParameter()                  ' 設定共用參數
        TopPosition()                   ' 設定top
        SetPopupFunction()              ' 設定彈出視窗事件
        ShowSheetField("New")           ' 表單欄位顯示及欄位輸入檢查
        ShowSheetFunction()             ' 表單功能按鈕顯示
        If Not IsPostBack Then
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
        Top = 450
        Top = Top + 85
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
        Response.Cookies("PGM").Value = "SPD_YKKGroupSheet_01.aspx"
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
                BSAVE.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BSAVE.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
            Else
                BSAVE.Visible = False
            End If
            'NG-1按鈕
            If dtFlow.Rows(0)("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Text = dtFlow.Rows(0)("NGDesc1")
                BNG1.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BNG1.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wNGSts1 = dtFlow.Rows(0)("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2按鈕
            If dtFlow.Rows(0)("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Text = dtFlow.Rows(0)("NGDesc2")
                BNG2.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BNG2.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wNGSts2 = dtFlow.Rows(0)("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK按鈕
            If dtFlow.Rows(0)("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Text = dtFlow.Rows(0)("OKDesc")
                ' BOK.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BOK.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
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
        If wFormSno > 0 Then    '判斷是否[簽核]

            sql = "Select * From T_WaitHandle "
            sql = sql & " Where Active = 1 "
            sql = sql & "   And FormNo =  '" & wFormNo & "'"
            sql = sql & "   And FormSno =  '" & CStr(wFormSno) & "'"
            sql = sql & "   And Step   =  '" & CStr(wStep) & "'"
            Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(sql)
            If dtWaitHandle.Rows.Count > 0 Then
                '核定說明
                DDescSheet.Visible = True       '說明Sheet
                DDecideDesc.Visible = True      '說明欄位顯示
                DDescSheet.Style("top") = Top & "px"
                DDecideDesc.Style("top") = Top + 6 & "px"
                If dtWaitHandle.Rows(0)("FlowType") = 0 Then        '通知
                    DDecideDesc.BackColor = Color.LightGray
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "異常：需輸入說明")
                End If
                Top = Top + 74
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
                        'Sheet顯示
                        DDelay.Style("top") = Top & "px"
                        DReasonCode.Style("top") = Top & "px"
                        DReason.Style("top") = Top & "px"
                        DReasonDesc.Style("top") = Top & "px"
                    Else
                        DDelay.Visible = False  '延遲Sheet
                        DReasonCode.Visible = False     '延遲理由代碼
                        DReason.Visible = False         '延遲理由
                        DReasonDesc.Visible = False     '延遲其他說明
                    End If
                End If
            End If

        Else
            'Sheet隱藏
            DDescSheet.Visible = False  '說明Sheet
            DDelay.Visible = False      '延遲Sheet
            '欄位隱藏
            DDecideDesc.Visible = False    '說明
            DReasonCode.Visible = False    '延遲理由代碼
            DReason.Visible = False        '延遲理由
            DReasonDesc.Visible = False    '延遲其他說明
            DHistoryLabel.Visible = False  '核定履歷

            'Top設定(TopPosition中設定) 
        End If
        '按鈕及超連結設值
        BSAVE.Style("top") = Top & "px"
        BNG1.Style("top") = Top & "px"
        BNG2.Style("top") = Top & "px"
        BOK.Style("top") = Top & "px"
        DHistoryLabel.Style("top") = Top + 32 & "px"
        GridView2.Style("top") = Top + 32 + 16 & "px"
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
        If pPost = "New" Then
            DDate.Text = CDate(NowDateTime).ToString("yyyy/MM/dd")
            Dim sql As String = "Select e.Com_Code,e.Com_Name,e.Dep_Code,e.Dep_Name,e.[ID],e.[Name],e.Job_Title_Code,e.Job_Title from M_Users u inner join M_Emp e on u.EmpID = e.[ID] where u.UserID='" & wApplyID & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(sql)
            '取得申請者資訊
            If dt.Rows.Count > 0 Then
                DDivision.Text = dt.Rows(0)("Dep_Name").ToString.Trim
                DPerson.Text = dt.Rows(0)("Name").ToString.Trim

            End If
        End If

        'No
        Select Case FindFieldInf("No")
            Case 0  '顯示
                DNo.BackColor = Color.LightGray
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

        'Person
        Select Case FindFieldInf("Person")
            Case 0  '顯示
                DPerson.BackColor = Color.LightGray
                DPerson.Visible = True
                DPerson.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DPerson.Visible = True
                DPerson.BackColor = Color.GreenYellow
                DPerson.ReadOnly = False
                ShowRequiredFieldValidator("DNameRqd", "DName", "異常：需輸入姓名")
            Case 2  '修改
                DPerson.Visible = True
                DPerson.BackColor = Color.Yellow
                DPerson.ReadOnly = False
            Case Else   '隱藏
                DPerson.Visible = False
        End Select

        'Division
        'Select Case FindFieldInf("Division")
        Select Case FindFieldInf("Division")
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



        'MapNo
        Select Case FindFieldInf("MapNo")
            Case 0  '顯示
                DMapNo.BackColor = Color.LightGray
                DMapNo.Visible = True
                DMapNo.Attributes.Add("readonly", "true")
                BMapNo.Visible = False
            Case 1  '修改+檢查
                DMapNo.Visible = True
                DMapNo.BackColor = Color.GreenYellow
                DMapNo.ReadOnly = False
                ShowRequiredFieldValidator("DMapNoRqd", "DMapNo", "異常：需輸入拉頭/拉片品名")
                BMapNo.Visible = True
            Case 2  '修改
                DMapNo.Visible = True
                DMapNo.BackColor = Color.Yellow
                DMapNo.ReadOnly = False
                BMapNo.Visible = True
            Case Else   '隱藏
                DMapNo.Visible = False
                BMapNo.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("MapNo", "ZZZZZZ")

        'Buyer
        Select Case FindFieldInf("Buyer")
            Case 0  '顯示
                DBuyer.BackColor = Color.LightGray
                DBuyer.Visible = True
                DBuyer.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DBuyer.Visible = True
                DBuyer.BackColor = Color.GreenYellow
                DBuyer.ReadOnly = False
                ShowRequiredFieldValidator("DBuyerRqd", "DBuyer", "異常：需輸入Buyer")
            Case 2  '修改
                DBuyer.Visible = True
                DBuyer.BackColor = Color.Yellow
                DBuyer.ReadOnly = False
            Case Else   '隱藏
                DBuyer.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Buyer", "ZZZZZZ")


        'OFormNo
        Select Case FindFieldInf("OFormNo")
            Case 0  '顯示
                DOFormNo.BackColor = Color.LightGray
                DOFormNo.Visible = True
                DOFormNo.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DOFormNo.Visible = True
                DOFormNo.BackColor = Color.GreenYellow
                DOFormNo.ReadOnly = False
                ShowRequiredFieldValidator("DOFormNoRqd", "DOFormNo", "異常：需輸入OFormNo")
            Case 2  '修改
                DOFormNo.Visible = True
                DOFormNo.BackColor = Color.Yellow
                DOFormNo.ReadOnly = False
            Case Else   '隱藏
                DOFormNo.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("OFormNo", "ZZZZZZ")


        'OFormSNo
        Select Case FindFieldInf("OFormSNo")
            Case 0  '顯示
                DOFormSNo.BackColor = Color.LightGray
                DOFormSNo.Visible = True
                DOFormSNo.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DOFormSNo.Visible = True
                DOFormSNo.BackColor = Color.GreenYellow
                DOFormSNo.ReadOnly = False
                ShowRequiredFieldValidator("DOFormSNoRqd", "DOFormSNo", "異常：需輸入OFormSNo")
            Case 2  '修改
                DOFormSNo.Visible = True
                DOFormSNo.BackColor = Color.Yellow
                DOFormSNo.ReadOnly = False
            Case Else   '隱藏
                DOFormSno.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("OFormSNo", "ZZZZZZ")




        'ProvideDate
        Select Case FindFieldInf("ProvideDate")
            Case 0  '顯示
                DProvideDate.BackColor = Color.LightGray
                DProvideDate.Visible = True
                DProvideDate.Attributes.Add("readonly", "true")
                BDate.Visible = False
            Case 1  '修改+檢查
                DProvideDate.Visible = True
                DProvideDate.BackColor = Color.GreenYellow
                DProvideDate.ReadOnly = False
                ShowRequiredFieldValidator("DProvideDateRqd", "DProvideDate", "異常：需輸入ProvideDate")
                BDate.Visible = True
            Case 2  '修改
                DProvideDate.Visible = True
                DProvideDate.BackColor = Color.Yellow
                DProvideDate.ReadOnly = False
                BDate.Visible = True
            Case Else   '隱藏
                DProvideDate.Visible = False
                BDate.Visible = True
        End Select
        If pPost = "New" Then SetFieldData("ProvideDate", "ZZZZZZ")



        'SliderCode
        Select Case FindFieldInf("SliderCode")
            Case 0  '顯示
                DSliderCode.BackColor = Color.LightGray
                DSliderCode.Visible = True
                DSliderCode.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSliderCode.Visible = True
                DSliderCode.BackColor = Color.GreenYellow
                DSliderCode.ReadOnly = False
                ShowRequiredFieldValidator("DSliderCodeRqd", "DSliderCode", "異常：需輸入拉頭/拉片品名")
            Case 2  '修改
                DSliderCode.Visible = True
                DSliderCode.BackColor = Color.Yellow
                DSliderCode.ReadOnly = False
            Case Else   '隱藏
                DSliderCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SliderCode", "ZZZZZZ")


        'WaveCode
        Select Case FindFieldInf("WaveCode")
            Case 0  '顯示
                DWaveCode.BackColor = Color.LightGray
                DWaveCode.Visible = True
                DWaveCode.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DWaveCode.Visible = True
                DWaveCode.BackColor = Color.GreenYellow
                DWaveCode.ReadOnly = False
                ShowRequiredFieldValidator("DWaveCodeRqd", "DWaveCode", "異常：需輸入WAVE'S CODE")
            Case 2  '修改
                DWaveCode.Visible = True
                DWaveCode.BackColor = Color.Yellow
                DWaveCode.ReadOnly = False
            Case Else   '隱藏
                DWaveCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WaveCode", "ZZZZZZ")




        'YKKGroup
        Select Case FindFieldInf("YKKGroup")
            Case 0  '顯示
                DYKKGroup.BackColor = Color.LightGray
                DYKKGroup.Visible = True
            Case 1  '修改+檢查
                DYKKGroup.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DYKKGroupRqd", "DYKKGroup", "異常：需輸入姐妹社")
                DYKKGroup.Visible = True
            Case 2  '修改
                DYKKGroup.BackColor = Color.Yellow
                DYKKGroup.Visible = True
            Case Else   '隱藏
                DYKKGroup.Visible = False
        End Select
        If Not IsPostBack Then SetFieldData("YKKGroup", "ZZZZZZ")



        'CopyReason
        Select Case FindFieldInf("CopyReason")
            Case 0  '顯示
                DCopyReason.BackColor = Color.LightGray
                DCopyReason.Visible = True
                DCopyReason.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCopyReason.Visible = True
                DCopyReason.BackColor = Color.GreenYellow
                DCopyReason.ReadOnly = False
                ShowRequiredFieldValidator("DCopyReasonRqd", "DCopyReason", "異常：需輸入複製理由")
            Case 2  '修改
                DCopyReason.Visible = True
                DCopyReason.BackColor = Color.Yellow
                DCopyReason.ReadOnly = False
            Case Else   '隱藏
                DCopyReason.Visible = False
        End Select
        If Not IsPostBack Then SetFieldData("CopyReason", "ZZZZZZ")

        'Forcast
        Select Case FindFieldInf("Forcast")
            Case 0  '顯示
                DForcast.BackColor = Color.LightGray
                DForcast.Visible = True
                DForcast.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DForcast.Visible = True
                DForcast.BackColor = Color.GreenYellow
                DForcast.ReadOnly = False
                ShowRequiredFieldValidator("DForcastRqd", "DForcast", "異常：需輸入預估訂單量")
            Case 2  '修改
                DForcast.Visible = True
                DForcast.BackColor = Color.Yellow
                DForcast.ReadOnly = False
            Case Else   '隱藏
                DForcast.Visible = False
        End Select
        If Not IsPostBack Then SetFieldData("Forcast", "ZZZZZZ")


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
        If Not IsPostBack Then SetFieldData("Remark", "ZZZZZZ")

        '複製理由
        Select Case FindFieldInf("CopyCheck1")
            Case 0  '顯示
                ' DCopyCheck1.BackColor = Color.LightGray
                DCopyCheck1.Enabled = False

            Case 1  '修改+檢查
                DCopyCheck1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCopyCheck1Rqd", "DCopyCheck1", "異常：需輸入複製理由")

            Case 2  '修改
                DCopyCheck1.BackColor = Color.Yellow

            Case Else   '隱藏
                DCopyCheck1.Visible = False
        End Select
        '    If pPost = "New" Then DCopyCheck1.Checked = False

        '複製理由
        Select Case FindFieldInf("CopyCheck2")
            Case 0  '顯示
                ' DCopyCheck1.BackColor = Color.LightGray
                DCopyCheck2.Enabled = False

            Case 1  '修改+檢查
                DCopyCheck2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCopyCheck2Rqd", "DCopyCheck2", "異常：需輸入複製理由")

            Case 2  '修改
                DCopyCheck2.BackColor = Color.Yellow

            Case Else   '隱藏
                DCopyCheck2.Visible = False
        End Select
        '  If pPost = "New" Then DCopyCheck2.Checked = False

        '複製理由
        Select Case FindFieldInf("CopyCheck3")
            Case 0  '顯示
                ' DCopyCheck1.BackColor = Color.LightGray
                DCopyCheck3.Enabled = False

            Case 1  '修改+檢查
                DCopyCheck3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCopyCheck3Rqd", "DCopyCheck3", "異常：需輸入複製理由")

            Case 2  '修改
                DCopyCheck3.BackColor = Color.Yellow

            Case Else   '隱藏
                DCopyCheck3.Visible = False
        End Select
        ' If pPost = "New" Then DCopyCheck3.Checked = False
    End Sub
    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim sql As String = ""
        Dim idx As Integer = FindFieldInf(pFieldName)


        'ykk姐妹社
        If pFieldName = "YKKGroup" Then
            DYKKGroup.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DYKKGroup.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='700' and dkey = 'YKKGROUP'  Order by DKey, Data "
                Dim dtReasonCode As DataTable = uDataBase.GetDataTable(sql)
                DYKKGroup.Items.Add("")
                For i As Integer = 0 To dtReasonCode.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReasonCode.Rows(i)("Data")
                    ListItem1.Value = dtReasonCode.Rows(i)("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DYKKGroup.Items.Add(ListItem1)
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
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = uCommon.GetAppSetting("Http")
        Dim SQL As String
        SQL = "Select * From F_YKKGroupCopySheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dt As DataTable = uDataBase.GetDataTable(SQL)
        If dt.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dt.Rows(0).Item("No")                           ' No
            DDate.Text = dt.Rows(0).Item("Date")                       ' 申請日
            DPerson.Text = dt.Rows(0).Item("Person")                       ' 申請人姓名
            DDivision.Text = dt.Rows(0).Item("Division")               ' 部門
            DBuyer.Text = dt.Rows(0).Item("Buyer")                   ' BUYER
            DProvideDate.Text = dt.Rows(0).Item("ProvideDate")                   ' ProvideDate
            DMapNo.Text = dt.Rows(0).Item("MapNo")                   ' MAPNO
            DSliderCode.Text = dt.Rows(0).Item("SliderCode")                   ' SliderCode
            DWaveCode.Text = dt.Rows(0).Item("WaveCode")                   ' WAVECODE
            SetFieldData("YKKGroup", dt.Rows(0).Item("YKKGroup"))    'YKKGROUP
            DCopyReason.Text = dt.Rows(0).Item("CopyReason")                   ' copyReason
            DForcast.Text = dt.Rows(0).Item("Forcast")                   ' Forcast
            DRemark.Text = dt.Rows(0).Item("Remark")                   ' Remark
            DOFormNo.Text = dt.Rows(0).Item("OFormNo")                   ' Forcast
            DOFormSno.Text = dt.Rows(0).Item("OFormSNo")                   ' Remark

            If dt.Rows(0).Item("OFormNo") = "000003" Then
                LFormsno.NavigateUrl = Path & "WorkFlow/ManufInSheet_02.aspx?pFormNo=" & dt.Rows(0).Item("OFormNo") & "&pFormSno=" & dt.Rows(0).Item("OFormSNo")
                LFormsno.Visible = True
            ElseIf dt.Rows(0).Item("OFormNo") = "000007" Then
                LFormsno.NavigateUrl = Path & "WorkFlow/ManufOutSheet_02.aspx?pFormNo=" & dt.Rows(0).Item("OFormNo") & "&pFormSno=" & dt.Rows(0).Item("OFormSNo")
                LFormsno.Visible = True
            Else
                LFormsno.Visible = False
            End If

            If dt.Rows(0).Item("CopyCheck1") = 1 Then
                DCopyCheck1.Checked = True
            ElseIf dt.Rows(0).Item("CopyCheck2") = 1 Then
                DCopyCheck2.Checked = True
            Else
                DCopyCheck3.Checked = True
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
    '**(ShowRequiredFieldValidator)
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator
        Dim sql As String = ""
        Dim xTop As Integer = 0
        '
        If wFormSno > 0 And wStep > 3 Then    '判斷是否[簽核]
            xTop = 74
            sql = "Select * From M_Flow "
            sql &= " Where Active = 1 "
            sql &= "   And FormNo =  '" & wFormNo & "'"
            sql &= "   And Step   =  '" & wStep & "'"
            Dim dtFlow As DataTable = uDataBase.GetDataTable(sql)
            If dtFlow.Rows.Count > 0 Then
                If dtFlow.Rows(0)("Delay") = 1 Then xTop = xTop + 110 '遲納管理
            End If
        End If
        '
        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.Text = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", Top + xTop & "px")
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
        BDate.Attributes.Add("onClick", "calendarPicker('DProvideDate')")
        BMapNo.Attributes.Add("onClick", "GetMapNo()")
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
                        DNo.Text = SetNo(NewFormSno)
                        '
                        '將待簽核的資料轉入暫存檔 T_OverTimeAutoJob
                        Dim sql As String
                        sql = " insert into Q_WaitAutoApprove (formno,formsno,seqno,applyid,DecideID,step,Division,createtime) "
                        sql = sql + "values('" + wFormNo + "','" + CStr(NewFormSno) + "'," + "1" + ",'" + wApplyID + "','" + Request.QueryString("pUserID") + "'," + "1" + ",'" + "" + "',getdate())"
                        uDataBase.ExecuteNonQuery(sql)
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

        Dim sql As String = ""
        sql = "INSERT INTO F_YKKGroupCopySheet ( " & _
              "Sts, CompletedTime, FormNo, FormSno, " & _
              "No, Date, Division,Person,Buyer,ProvideDate, " & _
              "MapNo,SliderCode,WaveCode,YKKGroup,CopyCheck1,CopyCheck2,CopyCheck3," & _
              "CopyReason,Forcast,Remark,OFormNo,OFormSNo, " & _
              "CreateUser, CreateTime, ModifyUser, ModifyTime) " & _
        "VALUES("
        sql &= "'0' ,"
        sql &= "'" & NowDateTime & "' ,"
        sql &= "'" & wFormNo & "' ,"
        sql &= "'" & CStr(NewFormSno) & "' ,"
        sql &= "'" & DNo.Text & "' ,"
        sql &= "'" & DDate.Text & "' ,"
        sql &= "N'" & DDivision.Text & "' ,"
        sql &= "N'" & DPerson.Text & "' ,"
        sql &= "N'" & DBuyer.Text & "' ,"
        sql &= "N'" & DProvideDate.Text & "' ,"
        sql &= "N'" & DMapNo.Text & "' ,"
        sql &= "N'" & DSliderCode.Text & "' ,"
        sql &= "N'" & DWaveCode.Text & "' ,"
        sql &= "N'" & DYKKGroup.SelectedValue & "' ,"
        If DCopyCheck1.Checked = True Then
            sql &= "1,"
        Else
            sql &= "0,"
        End If
        If DCopyCheck2.Checked = True Then
            sql &= "1,"
        Else
            sql &= "0,"
        End If
        If DCopyCheck3.Checked = True Then
            sql &= "1,"
        Else
            sql &= "0,"
        End If
        sql &= "N'" & DCopyReason.Text & "' ,"
        sql &= "N'" & DForcast.Text & "' ,"
        sql &= "N'" & DRemark.Text & "' ,"
        sql &= "N'" & DOFormNo.Text & "' ,"
        sql &= "N'" & DOFormSno.Text & "' ,"
        sql &= "N'" & Request.QueryString("pUserID") & "' ,"
        sql &= "N'" & NowDateTime & "' ,"
        sql &= "'" & Request.QueryString("pUserID") & "' ,"
        sql &= "'" & NowDateTime & "' )"
        uDataBase.ExecuteNonQuery(sql)


        '更新內製或外注姐妹社履歷註記
        If DOFormNo.Text = "000003" Then  '更新內製
            sql = " Update F_ManufInSheet"
            sql &= "  set YKKGroupCopy =1 "
            sql &= " Where FormNo  =  '" & DOFormNo.Text & "'"
            sql &= "   And FormSno =  '" & DOFormSno.Text & "'"
            uDataBase.ExecuteNonQuery(sql)
        End If

        '更新內製或外注姐妹社履歷註記
        If DOFormNo.Text = "000007" Then  '更新外注
            sql = " Update F_ManufOutSheet"
            sql &= "  set YKKGroupCopy =1 "
            sql &= " Where FormNo  =  '" & DOFormNo.Text & "'"
            sql &= "   And FormSno =  '" & DOFormSno.Text & "'"
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

        Dim sql As String

        sql = "Update F_YKKGroupCopySheet Set "
        If pFun <> "SAVE" Then
            sql &= " Sts = '" & pSts & "',"
            sql &= " CompletedTime = '" & NowDateTime & "',"
        End If
        sql &= " FormNo = '" & wFormNo & "',"
        sql &= " FormsNo = '" & wFormSno & "',"
        sql &= " No = '" & DNo.Text & "',"
        sql &= " Date = '" & DDate.Text & "',"
        sql &= " Division = '" & DDivision.Text & "',"
        sql &= " Person = '" & DPerson.Text & "',"
        sql &= " Buyer = '" & DBuyer.Text & "',"
        sql &= " ProvideDate = '" & DProvideDate.Text & "',"
        sql &= " SliderCode = '" & DSliderCode.Text & "',"
        sql &= " WaveCode = '" & DWaveCode.Text & "',"
        sql &= " YKKGroup= '" & DYKKGroup.SelectedValue & "',"

        If DCopyCheck1.Checked = True Then
            sql &= " CopyCheck1= 1,"
        Else
            sql &= " CopyCheck1= 0,"
        End If

        If DCopyCheck2.Checked = True Then

            sql &= " CopyCheck2 = 1,"
        Else

            sql &= " CopyCheck2 = 0,"
        End If
        If DCopyCheck3.Checked = True Then
            sql &= " CopyCheck3 = 1,"
        Else
            sql &= " CopyCheck3 = 0,"
        End If

        sql &= " CopyReason = '" & DCopyReason.Text & "',"
        sql &= " ForCast = '" & DForcast.Text & "',"
        sql &= " Remark = '" & DRemark.Text & "',"
        sql &= " OFormNo = '" & DOFormNo.Text & "',"
        sql &= " OFormSNo = '" & DOFormSno.Text & "',"
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

        If DCopyCheck1.Checked = False And DCopyCheck2.Checked = False And DCopyCheck3.Checked = False Then
            ErrCode = 9070
        End If

        If DCopyCheck3.Checked Then
            If DCopyReason.Text = "" Then
                ErrCode = 9060
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


        If ErrCode <> 0 Then
            If ErrCode = 9010 Then Message = "無附件檔案,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案格式有誤,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案Size有誤,請確認!"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9050 Then Message = "委託書No.重覆,請確認委託書No.!"
            If ErrCode = 9060 Then Message = "請輸入複製理由其它的原因!"
            If ErrCode = 9070 Then Message = "請選撢複製理由!"
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
        Dim allowedExtensions As String() = {".xls"}   '定義允許的檔案格式
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
 
    Protected Sub DCopyCheck1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DCopyCheck1.CheckedChanged
        If DCopyCheck1.Checked Then
            DCopyCheck2.Checked = False
            DCopyCheck3.Checked = False
        End If
    End Sub

    Protected Sub DCopyCheck2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DCopyCheck2.CheckedChanged
        If DCopyCheck2.Checked Then
            DCopyCheck3.Checked = False
            DCopyCheck1.Checked = False
        End If
    End Sub

    Protected Sub DCopyCheck3_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DCopyCheck3.CheckedChanged
        If DCopyCheck3.Checked Then
            DCopyCheck1.Checked = False
            DCopyCheck2.Checked = False
            DCopyReason.BackColor = Color.GreenYellow
        End If
    End Sub
End Class
