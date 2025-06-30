Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

Partial Class PriceInforSheet_01
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
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID
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

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim HourList, MinList, AdjList As New ListItemCollection
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    Dim aAgentID As String


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SetParameter()                  ' 設定共用參數
        TopPosition()                   ' 設定top

        If Not IsPostBack Then
            Server.ScriptTimeout = 900  ' 設定逾時時間

            ShowSheetField("New")       ' 表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()         ' 表單功能按鈕顯示
            ShowFormData()              ' 顯示表單資料   
            If wFormSno > 0 And wStep > 3 Then
                UpdateTranFile()        ' 更新交易資料
            End If
            SetControlPosition()

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
        wSeqNo = Request.QueryString("pSeqNo")      '序號
        wApplyID = Request.QueryString("pApplyID")  '申請者ID
        wAgentID = Request.QueryString("pAgentID")  '被代理人ID

        'Add-Start by Joy 2009/12/2(2010行事曆對應)
        '申請者-群組行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        wDecideCalendar = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))
        'Add-End

        Response.Cookies("PGM").Value = "PriceInforSheet_01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetControlPosition)
    '**    設定按鈕,說明,延遲,核定履歷的位置
    '**
    '*****************************************************************
    Sub SetControlPosition()
        If grd.Visible = True Then

            Dim GVTop As Integer = grd.Style("top").Replace("px", "")
            Dim GVHeight As Integer = grd.Rows.Count * 2 * 31 + 10          ' 55是列高
            Dim ControlTop As Integer = (GVTop + GVHeight) + 10

            If DDescSheet.Visible Then                                      ' 說明
                DDescSheet.Style("top") = ControlTop & "px"
                DDecideDesc.Style("top") = ControlTop + 6 & "px"
                ControlTop += 101
            End If

            If DDelay.Visible Then                                          ' 延遲說明
                DDelay.Style("top") = ControlTop & "px"
                DReasonCode.Style("top") = ControlTop + 9 & "px"
                DReason.Style("top") = ControlTop + 6 & "px"
                DReasonDesc.Style("top") = ControlTop + 39 & "px"
                ControlTop += 161
            End If

            BSAVE.Style("top") = ControlTop + 10 & "px"
            BNG1.Style("top") = ControlTop + 10 & "px"
            BNG2.Style("top") = ControlTop + 10 & "px"
            BOK.Style("top") = ControlTop + 10 & "px"
            ControlTop += 48

            If GridView2.Rows.Count > 0 Then                                ' 核定履歷
                DHistoryLabel.Style("top") = ControlTop & "px"
                ControlTop += 20
                GridView2.Style("top") = ControlTop & "px"
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
        Top = "500"
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

    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        If wFormSno > 0 And wStep > 3 Then
            SQL = "Select * From  F_PriceInforSheetH "
            SQL &= " Where FormNo =  '" & wFormNo & "'"
            SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(SQL)

            DDate.Text = CDate(dt.Rows(0)("Date").ToString).ToString("yyyy/MM/dd")
            DDepoName.Text = dt.Rows(0)("DepoName").ToString.Trim
            DDepoCode.Text = dt.Rows(0)("DepoCode").ToString.Trim
            DDivision.Text = dt.Rows(0)("Division").ToString.Trim
            DDivisionCode.Text = dt.Rows(0)("DivisionCode").ToString.Trim
            DName.Text = dt.Rows(0)("Name").ToString.Trim
            DEmpID.Text = dt.Rows(0)("EMPID").ToString.Trim
            DJobTitle.Text = dt.Rows(0)("JobTitle").ToString.Trim
            DJobCode.Text = dt.Rows(0)("JobCode").ToString.Trim
            DNo.Text = dt.Rows(0)("No").ToString.Trim()
            Dremark.Text = dt.Rows(0)("Remark").ToString.Trim()

            SQL = " Select ItemNo As No, * From  F_PriceInforSheetDt "
            SQL &= " Where FormNo =  '" & wFormNo & "'"
            SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL &= " Order by Code, Version "
            Dim dtPriceInforDT As DataTable = uDataBase.GetDataTable(SQL)
            grd.DataSource = dtPriceInforDT
            grd.DataBind()
        Else
            OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("TP_Conn")  'SQL連結設定
            SQL = " SELECT ApplyName,CType+Version As Version,No, Code, ItemName1 + ItemName2 + ItemName3 AS ItemName,"
            SQL = SQL + " A1, B1, B2, B4, C1, C2, C3, D1 FROM WFS_PriceInforH WHERE (WFS_No = '') AND (WFS_Sts = 9) "
            SQL = SQL + " Order by Code, Version "
            OleDbConnection1.Open()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "WFS_Price")
            grd.DataSource = DBDataSet1
            grd.DataBind()
        End If

        ' 帶入子表資訊到員工清單

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



        OleDbConnection1.Close()
    End Sub


    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)

        Dim sql As String = ""
        Dim idx As Integer = FindFieldInf(pFieldName)

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
                'BSAVE.Attributes("onclick") = "return Button('SAVE', '" + BSAVE.Text + "');"
            Else
                BSAVE.Visible = False
            End If
            'NG-1按鈕
            If dtFlow.Rows(0)("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Text = dtFlow.Rows(0)("NGDesc1")
                'BNG1.Attributes("onclick") = "return Button('NG1', '" + BNG1.Text + "');"
                wNGSts1 = dtFlow.Rows(0)("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2按鈕
            If dtFlow.Rows(0)("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Text = dtFlow.Rows(0)("NGDesc2")
                'BNG2.Attributes("onclick") = "return Button('NG2', '" + BNG2.Text + "');"
                wNGSts2 = dtFlow.Rows(0)("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK按鈕
            If dtFlow.Rows(0)("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Text = dtFlow.Rows(0)("OKDesc")
                'BOK.Attributes("onclick") = "return Button('OK', '" + BOK.Text + "');"
                wOKSts = dtFlow.Rows(0)("OKSts") + 1
            Else
                BOK.Visible = False
            End If
            '遲納管理
            If dtFlow.Rows(0)("Delay") = 1 Then
                wDelay = 1
            End If
        End If

        If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
            sql = "Select * From T_WaitHandle "
            sql = sql & " Where Active = 1 "
            sql = sql & "   And FormNo =  '" & wFormNo & "'"
            sql = sql & "   And FormSno =  '" & CStr(wFormSno) & "'"
            sql = sql & "   And Step   =  '" & CStr(wStep) & "'"

            Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(sql)

            If dtWaitHandle.Rows.Count > 0 Then
                'Sheet顯示
                DDescSheet.Visible = True        '說明Sheet
                ' 代理人顯示 Modify by Jessica 2010/06/07
                aAgentID = dtWaitHandle.Rows(0)("AgentID")
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
                    Else
                        DDelay.Visible = False  '延遲Sheet
                        DReasonCode.Visible = False     '延遲理由代碼
                        DReason.Visible = False         '延遲理由
                        DReasonDesc.Visible = False     '延遲其他說明
                    End If
                End If

                '欄位顯示
                DDecideDesc.Visible = True      '說明
                '說明需輸入
                If dtWaitHandle.Rows(0)("FlowType") = 0 Then
                    DDecideDesc.BackColor = Color.LightGray
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "異常：需輸入說明")
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

        End If


    End Sub



    '*****************************************************************
    '**(SetFieldAttribute)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        'DoverTimeDate.Text = Now.ToShortDateString

        If pPost = "New" Then
            Dim sql As String = "Select e.Com_Code,e.Com_Name,e.Dep_Code,e.Dep_Name,e.[ID],e.[Name],e.Job_Title_Code,e.Job_Title from M_Users u inner join M_Emp e on u.EmpID = e.[ID] where u.UserID='" & wApplyID & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(sql)

            '取得申請者資訊

            DDate.Text = CDate(NowDateTime).ToString("yyyy/MM/dd")
            DDepoName.Text = dt.Rows(0)("Com_Name").ToString.Trim
            DDepoCode.Text = dt.Rows(0)("Com_Code").ToString.Trim
            DDivision.Text = dt.Rows(0)("Dep_Name").ToString.Trim
            DDivisionCode.Text = dt.Rows(0)("Dep_Code").ToString.Trim
            DName.Text = dt.Rows(0)("Name").ToString.Trim
            DEmpID.Text = dt.Rows(0)("ID").ToString.Trim
            DJobTitle.Text = dt.Rows(0)("Job_Title").ToString.Trim
            DJobCode.Text = dt.Rows(0)("Job_Title_Code").ToString.Trim


            ' DOverDivision.DataSource = fpObj.GetApplyDiv(wApplyID)      ' 取得可申請加班的部門
            ' DOverDivision.DataBind()



        End If
        'No
        Select Case FindFieldInf("No")
            Case 0  '顯示
                DNo.BackColor = Color.White
                DNo.Visible = True
                DNo.ReadOnly = True
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
                DDate.ReadOnly = True
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

        'DepoName
        Select Case FindFieldInf("DepoName")
            Case 0  '顯示
                DDepoName.BackColor = Color.LightGray
                DDepoName.Visible = True
                DDepoName.ReadOnly = True
            Case 1  '修改+檢查
                DDepoName.Visible = True
                DDepoName.BackColor = Color.GreenYellow
                DDepoName.ReadOnly = False
                ShowRequiredFieldValidator("DDepoNameRqd", "DDepoName", "異常：需輸入公司")
            Case 2  '修改
                DDepoName.Visible = True
                DDepoName.BackColor = Color.Yellow
                DDepoName.ReadOnly = False
            Case Else   '隱藏
                DDepoName.Visible = False
        End Select

        'DepoCode
        Select Case FindFieldInf("DepoCode")
            Case 0  '顯示
                DDepoCode.BackColor = Color.LightGray
                DDepoCode.Visible = True
                DDepoCode.ReadOnly = True
            Case 1  '修改+檢查
                DDepoCode.Visible = True
                DDepoCode.BackColor = Color.GreenYellow
                DDepoCode.ReadOnly = False
                ShowRequiredFieldValidator("DDepoCodeRqd", "DDepoCode", "異常：需輸入公司代碼")
            Case 2  '修改
                DDepoCode.Visible = True
                DDepoCode.BackColor = Color.Yellow
                DDepoCode.ReadOnly = False
            Case Else   '隱藏
                DDepoCode.Visible = False
        End Select

        'Division
        Select Case FindFieldInf("Division")
            Case 0  '顯示
                DDivision.BackColor = Color.LightGray
                DDivision.Visible = True
                DDivision.ReadOnly = True
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

        'DivisionCode
        Select Case FindFieldInf("DivisionCode")
            Case 0  '顯示
                DDivisionCode.BackColor = Color.LightGray
                DDivisionCode.Visible = True
                DDivisionCode.ReadOnly = True
            Case 1  '修改+檢查
                DDivisionCode.Visible = True
                DDivisionCode.BackColor = Color.GreenYellow
                DDivisionCode.ReadOnly = False
                ShowRequiredFieldValidator("DDivisionCodeRqd", "DDivisionCode", "異常：需輸入部門代碼")
            Case 2  '修改
                DDivisionCode.Visible = True
                DDivisionCode.BackColor = Color.Yellow
                DDivisionCode.ReadOnly = False
            Case Else   '隱藏
                DDivisionCode.Visible = False
        End Select

        'Name
        Select Case FindFieldInf("Name")
            Case 0  '顯示
                DName.BackColor = Color.LightGray
                DName.Visible = True
                DName.ReadOnly = True
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

        'EmpID
        Select Case FindFieldInf("EmpId")
            Case 0  '顯示
                DEmpID.BackColor = Color.LightGray
                DEmpID.Visible = True
                DEmpID.ReadOnly = True
            Case 1  '修改+檢查
                DEmpID.Visible = True
                DEmpID.BackColor = Color.GreenYellow
                DEmpID.ReadOnly = False
                ShowRequiredFieldValidator("DEmpIDRqd", "DEmpID", "異常：需輸入卡號")
            Case 2  '修改
                DEmpID.Visible = True
                DEmpID.BackColor = Color.Yellow
                DEmpID.ReadOnly = False
            Case Else   '隱藏
                DEmpID.Visible = False
        End Select

        'JobTitle
        Select Case FindFieldInf("JobTitle")
            Case 0  '顯示
                DJobTitle.BackColor = Color.LightGray
                DJobTitle.Visible = True
                DJobTitle.ReadOnly = True
            Case 1  '修改+檢查
                DJobTitle.Visible = True
                DJobTitle.BackColor = Color.GreenYellow
                DJobTitle.ReadOnly = False
                ShowRequiredFieldValidator("DJobTitleRqd", "DJobTitle", "異常：需輸入職稱")
            Case 2  '修改
                DJobTitle.Visible = True
                DJobTitle.BackColor = Color.Yellow
                DJobTitle.ReadOnly = False
            Case Else   '隱藏
                DJobTitle.Visible = False
        End Select

        'JobCode
        Select Case FindFieldInf("JobCode")
            Case 0  '顯示
                DJobCode.BackColor = Color.LightGray
                DJobCode.Visible = True
                DJobCode.ReadOnly = True
            Case 1  '修改+檢查
                DJobCode.Visible = True
                DJobCode.BackColor = Color.GreenYellow
                DJobCode.ReadOnly = False
                ShowRequiredFieldValidator("DJobCodeRqd", "DJobCode", "異常：需輸入職稱代碼")
            Case 2  '修改
                DJobCode.Visible = True
                DJobCode.BackColor = Color.Yellow
                DJobCode.ReadOnly = False
            Case Else   '隱藏
                DJobCode.Visible = False
        End Select

        'Remark
        Select Case FindFieldInf("Remark")
            Case 0  '顯示
                Dremark.BackColor = Color.LightGray
                Dremark.Visible = True
                Dremark.ReadOnly = True
            Case 1  '修改+檢查
                Dremark.Visible = True
                Dremark.BackColor = Color.GreenYellow
                Dremark.ReadOnly = False
                ShowRequiredFieldValidator("DJobCodeRqd", "DJobCode", "異常：需輸入職稱代碼")
            Case 2  '修改
                Dremark.Visible = True
                Dremark.BackColor = Color.Yellow
                Dremark.ReadOnly = False
            Case Else   '隱藏
                DJobCode.Visible = False
        End Select


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
        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.ErrorMessage = pMessage
        rqdVal.Display = ValidatorDisplay.None  ' 因使用摘要驗證控制項 , 所以不須顯示一般驗證控制項
        rqdVal.Style.Add("Top", Top & "px")
        rqdVal.Style.Add("Left", "10px")
        rqdVal.Style.Add("z-index", "999")
        rqdVal.Style.Add("POSITION", "absolute")
        Me.Form.Controls.Add(rqdVal)
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

        '檢查委託書No
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                ErrCode = oCommon.CommissionNo("001161", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9060
                End If
            End If
        End If

        If ErrCode = 0 Then
            Dim Run As Boolean = True           '是否執行
            Dim RepeatRun As Boolean = False    '是否重覆執行
            Dim MultiJob As Integer = 0         '多人核定
            Dim wLevel As String = ""           '難易度

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
                            '申請流程資料建置
                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                            '
                            oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
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
                            'RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDepo, Request.QueryString("pUserID"), pRunNextStep)
                            '
                            RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, Request.QueryString("pUserID"), pRunNextStep)
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
                    '找出簽核者User ID
                    Dim wAllocateID As String = ""

                    '取得簽核者
                    RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, wAgentID, wAllocateID, MultiJob, _
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
                            'RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wDepo, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
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
                            RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '表單號碼,表單流水號,工程關卡號碼,序號,申請者ID,行事曆,建置者, 簽核者, 被代理者, 申請者, 預定開始日, 預定完成日, 重要性
                            'Modify-End

                            If RtnCode <> 0 Then
                                ErrCode = 9150
                                Exit For
                            End If
                        Next i
                    Else
                        'Modify-Start by 2009/11/20(2010行事曆對應)
                        'RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                        '
                        RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者
                        'Modify-End

                        If RtnCode <> 0 Then ErrCode = 9160
                    End If
                End If
                '當工程日程調整
                If ErrCode = 0 Then
                    If RepeatRun = False Then
                        'Modify-Start by 2009/11/20(2010行事曆對應)
                        'RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDepo)
                        '
                        RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDecideCalendar)
                        '簽核者,表單號碼,表單流水號,工程關卡號碼,序號,現在日時,難易度,行事曆
                        'Modify-End
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

                Dim URL As String = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                    "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
            End If
        Else
            EnabledButton()   '起動Button運作
            If ErrCode = 9010 Then Message = "承認者指定有誤,請確認!"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9060 Then Message = "委託書No.重覆,請確認委託書No.!"
            uJavaScript.PopMsg(Me, Message)
        End If      '上傳檔案ErrCode=0
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim sql As String

        sql = "Update  F_PriceInforSheetH Set "
        If pFun <> "SAVE" Then
            sql &= " Sts = '" & pSts & "',"
            sql &= " CompletedTime = '" & NowDateTime & "', "
        End If
        sql &= " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        sql &= " ModifyTime = '" & NowDateTime & "' "
        sql &= " Where FormNo  =  '" & wFormNo & "'"
        sql &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        uDataBase.ExecuteNonQuery(sql)

        sql = "Update  F_PriceInforSheetDt Set "
        If pFun <> "SAVE" Then
            sql &= " Sts = '" & pSts & "',"
            sql &= " CompletedTime = '" & NowDateTime & "',"
        End If
        sql &= " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        sql &= " ModifyTime = '" & NowDateTime & "' "
        sql &= " Where FormNo  =  '" & wFormNo & "'"
        sql &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        uDataBase.ExecuteNonQuery(sql)

        ChangeWFSPriceinfo() '更新wfs_priceInfo





    End Sub

    Sub ChangeWFSPriceinfo()
        '更新價格檔
        Dim SQL As String
        Dim wfssts As Integer
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("TP_Conn")  'SQL連結設定
        '
        If wStep = 1 Or wStep = 20 Or wStep = 500 Then
            If wStep = 1 Then
                wfssts = 0
                SQL = " UPDATE  WFS_PriceInforH"
                SQL = SQL + " SET WFS_NO = '" + DNo.Text + "'"
                SQL = SQL + ", WFS_STs = 0"
                SQL = SQL + "  WHERE WFS_No = '' AND WFS_Sts =9"
                '
                Dim cmd1 As New OleDb.OleDbCommand(SQL, OleDbConnection1)
                OleDbConnection1.Open()
                cmd1.ExecuteReader()
                OleDbConnection1.Dispose()
            End If
            '
            If wStep = 20 Then
                wfssts = 1
                SQL = " UPDATE  WFS_PriceInforH"
                SQL = SQL + " SET WFS_STs = '" + CStr(wfssts) + "', "
                SQL = SQL + " ModifyUser  = 'WFS-PriceSheet', "
                SQL = SQL + " ModifyTime  = '" + NowDateTime + "' "
                SQL = SQL + " WHERE WFS_No = '" + DNo.Text + "' AND WFS_Sts =0 "
                '
                Dim cmd1 As New OleDb.OleDbCommand(SQL, OleDbConnection1)
                OleDbConnection1.Open()
                cmd1.ExecuteReader()
                OleDbConnection1.Dispose()
            End If
            '
            If wStep = 500 Then
                If Request.Cookies("RunBNG1").Value = True Then
                    wfssts = 2
                    SQL = " UPDATE  WFS_PriceInforH"
                    SQL = SQL + " SET WFS_STs = '" + CStr(wfssts) + "', "
                    SQL = SQL + " ModifyUser  = 'WFS-PriceSheet', "
                    SQL = SQL + " ModifyTime  = '" + NowDateTime + "' "
                    SQL = SQL + " WHERE WFS_No = '" + DNo.Text + "' AND WFS_Sts =0 "
                    '
                    Dim cmd1 As New OleDb.OleDbCommand(SQL, OleDbConnection1)
                    OleDbConnection1.Open()
                    cmd1.ExecuteReader()
                    OleDbConnection1.Dispose()
                End If
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


        Dim SQL As String


        SQL = "INSERT INTO F_PriceInforSheetH (Sts, CompletedTime, FormNo, FormSno, [No], [Date], DepoCode, DepoName, " & _
              " Division, DivisionCode, Name, EmpID, JobTitle, JobCode,Remark, CreateUser, CreateTime, ModifyUser, ModifyTime) " & _
        "VALUES("

        SQL &= "'0' ,"
        SQL &= "'" & NowDateTime & "' ,"
        SQL &= "'" & wFormNo & "' ,"
        SQL &= "'" & CStr(NewFormSno) & "' ,"
        SQL &= "'" & DNo.Text & "' ,"
        SQL &= "'" & DDate.Text & "' ,"
        SQL &= "'" & DDepoCode.Text & "' ,"
        SQL &= "'" & DDepoName.Text & "' ,"
        SQL &= "'" & DDivision.Text & "' ,"
        SQL &= "'" & DDivisionCode.Text & "' ,"
        SQL &= "'" & DName.Text & "' ,"
        SQL &= "'" & DEmpID.Text & "' ,"
        SQL &= "'" & DJobTitle.Text & "' ,"
        SQL &= "'" & DJobCode.Text & "' ,"
        SQL &= "'" & Dremark.Text & "' ,"
        SQL &= "'" & Request.QueryString("pUserID") & "' ,"
        SQL &= "'" & NowDateTime & "' ,"
        SQL &= "'" & Request.QueryString("pUserID") & "' ,"
        SQL &= "'" & NowDateTime & "' )"
        uDataBase.ExecuteNonQuery(SQL)

        '更新主檔細項


        For Each di As GridViewRow In grd.Rows
            Dim chk As CheckBox = di.FindControl("IsSelect")

            Dim ItemNo As String = CType(di.FindControl("ItemNo"), TextBox).Text.Trim
            Dim Code As String = CType(di.FindControl("Code"), TextBox).Text.Trim
            Dim ItemName As String = CType(di.FindControl("ItemName"), TextBox).Text.Trim
            Dim ApplyName As String = CType(di.FindControl("ApplyName"), TextBox).Text.Trim
            Dim Version As String = CType(di.FindControl("Version"), TextBox).Text.Trim
            Dim A1 As String = CType(di.FindControl("A1"), TextBox).Text.Trim
            Dim B1 As String = CType(di.FindControl("B1"), TextBox).Text.Trim
            Dim B2 As String = CType(di.FindControl("B2"), TextBox).Text.Trim
            Dim B4 As String = CType(di.FindControl("B4"), TextBox).Text.Trim
            Dim C1 As String = CType(di.FindControl("C1"), TextBox).Text.Trim
            Dim C2 As String = CType(di.FindControl("C2"), TextBox).Text.Trim
            Dim C3 As String = CType(di.FindControl("C3"), TextBox).Text.Trim
            Dim D1 As String = CType(di.FindControl("D1"), TextBox).Text.Trim



            SQL = "Insert into F_PriceInforSheetDT (Sts, CompletedTime, FormNo, FormSno, ItemNo, Code, " & _
                    "ItemName, ApplyName, Version, A1, B1, B2, B4, C1, C2,C3, " & _
                    "D1, CreateUser, CreateTime, ModifyUser, ModifyTime) " & _
                    "Values (@Sts, @CompletedTime, @FormNo, @FormSno,@ItemNo, @Code,@ItemName, @ApplyName, " & _
                    "@Version, @A1, @B1, @B2, @B4, @C1, @C2,@C3,@D1,@CreateUser, @CreateTime, @ModifyUser, @ModifyTime)"
            Dim sqlcmd As New SqlCommand(SQL)
            sqlcmd.Parameters.AddWithValue("@Sts", "0")
            sqlcmd.Parameters.AddWithValue("@CompletedTime", NowDateTime)
            sqlcmd.Parameters.AddWithValue("@FormNo", wFormNo)
            sqlcmd.Parameters.AddWithValue("@FormSno", CStr(NewFormSno))
            sqlcmd.Parameters.AddWithValue("@ItemNo", ItemNo)
            sqlcmd.Parameters.AddWithValue("@Code", Code)
            sqlcmd.Parameters.AddWithValue("@ItemName", ItemName)
            sqlcmd.Parameters.AddWithValue("@ApplyName", ApplyName)
            sqlcmd.Parameters.AddWithValue("@Version", Version)
            sqlcmd.Parameters.AddWithValue("@A1", A1)
            sqlcmd.Parameters.AddWithValue("@B1", B1)
            sqlcmd.Parameters.AddWithValue("@B2", B2)
            sqlcmd.Parameters.AddWithValue("@B4", B4)
            sqlcmd.Parameters.AddWithValue("@C1", C1)
            sqlcmd.Parameters.AddWithValue("@C2", C2)
            sqlcmd.Parameters.AddWithValue("@C3", C3)
            sqlcmd.Parameters.AddWithValue("@D1", D1)
            sqlcmd.Parameters.AddWithValue("@CreateUser", Request.QueryString("pUserID"))
            sqlcmd.Parameters.AddWithValue("@CreateTime", NowDateTime)
            sqlcmd.Parameters.AddWithValue("@ModifyUser", Request.QueryString("pUserID"))
            sqlcmd.Parameters.AddWithValue("@ModifyTime", NowDateTime)

            uDataBase.ExecuteNonQuery(sqlcmd)



        Next
        ChangeWFSPriceinfo() '更新wfs_priceInfo

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
        If pFun <> "SAVE" Then
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
    '**     儲存按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BSAVE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSAVE.Click
        If Request.Cookies("RunBSAVE").Value = True Then
            DisabledButton()   '停止Button運作

            Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
            Dim Message As String = ""

            'Check理由
            If ErrCode = 0 Then
                If DReasonCode.Visible = True Then
                    If DReasonCode.SelectedValue = "99" Then
                        If DReasonDesc.Text = "" Then ErrCode = 9210
                    End If
                End If
            End If
            '儲存資料
            If ErrCode = 0 Then
                ModifyData("SAVE", "0")           '更新表單資料 Sts=0(未結)
                ModifyTranData("SAVE", "0")       '更新交易資料
            Else
                If ErrCode = 9210 Then Message = "延遲理由為其他時需填寫說明,請確認!"
                uJavaScript.PopMsg(Me, Message)
            End If      '上傳檔案ErrCode=0

            If ErrCode = 0 Then
                Dim URL As String = uCommon.GetAppSetting("RedirectURL.aspx") & "?pMSGID=S&pFormNo=" & wFormNo & "&pStep=" & wStep & _
                                                    "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
            End If

        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-2按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BNG2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNG2.Click
        If Request.Cookies("RunBNG2").Value = True Then
            DisabledButton()   '停止Button運作
            FlowControl("NG2", 2, "3")
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-1按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BNG1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNG1.Click
        If Request.Cookies("RunBNG1").Value = True Then
            DisabledButton()   '停止Button運作
            FlowControl("NG1", 1, "2")
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.Click
        If Request.Cookies("RunBOK").Value = True Then
            DisabledButton()   '停止Button運作
            FlowControl("OK", 0, "1")
        End If
    End Sub



End Class
