Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class HR_AwaySheet_01
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BADays As System.Web.UI.WebControls.Button
    Protected WithEvents BBDays As System.Web.UI.WebControls.Button
    Protected WithEvents BAEndDate As System.Web.UI.WebControls.Button
    Protected WithEvents BAStartDate As System.Web.UI.WebControls.Button
    Protected WithEvents DAEndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BBEndDate As System.Web.UI.WebControls.Button
    Protected WithEvents DBEndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BBStartDate As System.Web.UI.WebControls.Button
    Protected WithEvents DBStartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAEndH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAStartH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBEndH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBStartH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDepoCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSalaryYM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DHistoryLabel As System.Web.UI.WebControls.Label
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDescSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DJobCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataGrid9 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DDivisionCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobTitle As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEmpID As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents DPlace As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAHour As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobAgent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOtherPlace As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBHour As System.Web.UI.WebControls.TextBox
    Protected WithEvents DADay As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBDay As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAwaySheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents BCardTime As System.Web.UI.WebControls.Button

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim FieldName(60) As String     '各欄位
    Dim Attribute(60) As Integer    '各欄位屬性    
    Dim Top As Integer              '動態元件的Top位置
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
    'Dim wDepo As String = "TP"      '台北行事曆(CL->中壢, TP->台北)
    '
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    'Modify-End
    Dim wUserName As String = ""    '姓名代理用

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置

        If Not Me.IsPostBack Then   '不是PostBack
            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()     '表單功能按鈕顯示
            If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
                ShowFormData()      '顯示表單資料
                UpdateTranFile()    '更新交易資料
            End If
            SetPopupFunction()      '設定彈出視窗事件
        Else
            ShowSheetField("Post")  '表單欄位顯示及欄位輸入檢查
            ShowMessage()           '上傳資料檢查及顯示訊息
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '現在日時
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
        wStep = Request.QueryString("pStep")        '工程代碼
        wSeqNo = Request.QueryString("pSeqNo")      '序號
        wApplyID = Request.QueryString("pApplyID")  '申請者ID
        wAgentID = Request.QueryString("pAgentID")  '被代理人ID
        'Add-Start by Joy 2009/11/20(2010行事曆對應)
        '申請者-群組行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        wDecideCalendar = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))
        'Add-End

        Response.Cookies("PGM").Value = "HR_AwaySheet_01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     '上傳資料檢查及顯示訊息
    '**
    '*****************************************************************
    Sub ShowMessage()
        'Dim Message As String = ""

        'Check切結書
        'If DContactFile.Visible Then
        'If DContactFile.PostedFile.FileName <> "" Then
        'If Message = "" Then
        'Message = "切結書"
        'Else
        '    Message = Message & ", " & "切結書"
        'End If
        'End If
        'End If

        'If Message <> "" Then
        'Message = "下列所設定的附加檔案將會遺失 (" & Message & ") " & ",請重新設定!"
        'Response.Write(YKK.ShowMessage(Message))
        'End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub UpdateTranFile()
        'Modify-Start by Joy 2009/11/20(2010行事曆對應)
        'oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDepo, Request.QueryString("pUserID"))
        '
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, Request.QueryString("pUserID"))
        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者
        'Modify-End
    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("AwayFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet9 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_AwaySheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_AwaySheet")
        If DBDataSet1.Tables("F_AwaySheet").Rows.Count > 0 Then
            '表單資料
            DNo.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("Date")                   '申請日期
            DSalaryYM.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("SalaryYM")           '所屬年月

            SetFieldData("Name", DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("Name"))          '姓名
            DEmpID.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("EmpID")                 'EMPID
            DJobTitle.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("JobTitle")           '職稱
            DJobCode.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("JobCode")             '職稱代碼
            DDepoName.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("DepoName")           '公司別
            DDepoCode.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("DepoCode")           '公司別代碼
            DDivision.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("Division")           '部門
            DDivisionCode.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("DivisionCode")   '部門代碼

            SetFieldData("Place", DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("Place"))        '目的地
            DOtherPlace.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("OtherPlace")       '目的地(其他)
            DJobAgent.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("JobAgent")           '代理人

            DBStartDate.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("BStartDate")               '預定開始日期
            SetFieldData("BStartH", DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("BStartH").ToString)   '預定開始時
            DBEndDate.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("BEndDate")                   '預定結束日期
            SetFieldData("BEndH", DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("BEndH").ToString)       '預定結束時
            DBDay.Text = CStr(DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("BDay"))     '預定天
            DBHour.Text = CStr(DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("BHour"))   '預定時

            DAStartDate.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("AStartDate")               '實際開始日期
            SetFieldData("AStartH", DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("AStartH").ToString)   '實際開始時
            DAEndDate.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("AEndDate")                   '實際結束日期
            SetFieldData("AEndH", DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("AEndH").ToString)       '實際結束時
            DADay.Text = CStr(DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("ADay"))     '實際天
            DAHour.Text = CStr(DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("AHour"))   '實際時

            DFReason.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("FReason")             '外出理由
            '交易資料
            DBDataSet1.Clear()
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                DDecideDesc.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("DecideDesc")       '說明

                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") < NowDateTime Then
                    If DDelay.Visible = True Then
                        SetFieldData("ReasonCode", DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonCode"))    '延遲理由代碼
                        If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonCode") = "" Then
                            SetFieldData("Reason", DReasonCode.SelectedValue)    '延遲理由
                            DReasonDesc.Text = ""   '延遲其他說明
                        Else
                            DReason.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("Reason")  '延遲理由
                            DReasonDesc.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonDesc")     '延遲其他說明
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
            'SQL = SQL + "Order by Unique_ID Desc "
            SQL = SQL + "Order by CreateTime Desc, Step Desc, SeqNo Desc "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet9, "DecideHistory")
            DataGrid9.DataSource = DBDataSet9
            DataGrid9.DataBind()
        End If
        'DB連結關閉
        OleDbConnection1.Close()
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
    '**(ShowSheetFunction)
    '**     表單功能按鈕顯示
    '**
    '*****************************************************************
    Sub ShowSheetFunction()
        Dim wDelay As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Flow")
        If DBDataSet1.Tables("M_Flow").Rows.Count > 0 Then
            '電子簽章未使用
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("SignImage") = 1 Then
            Else
            End If
            '附加檔案未使用(由FormField中設定)
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("Attach") = 1 Then
                'DRefMapFile.Visible = True
                'DMapFile.Visible = True
            Else
                'DRefMapFile.Visible = False
                'DMapFile.Visible = False
            End If
            '儲存按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("SaveFun") = 1 Then
                BSAVE.Visible = True
                BSAVE.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("SaveDesc")
                '-- 防止按鈕按二次 JOY 16/08/17
                '舊
                'BSAVE.Attributes("onclick") = "Button('SAVE', '" + BSAVE.Value + "');"
                '新
                BSAVE.Attributes("onclick") = "this.disabled = true;" & "Button('SAVE', '" + BSAVE.Value + "');"
                '--
            Else
                BSAVE.Visible = False
            End If
            'NG-1按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGDesc1")
                '-- 防止按鈕按二次 JOY 16/08/17
                '舊
                'BNG1.Attributes("onclick") = "Button('NG1', '" + BNG1.Value + "');"
                '新
                BNG1.Attributes("onclick") = "this.disabled = true;" & "Button('NG1', '" + BNG1.Value + "');"
                '--
                wNGSts1 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGDesc2")
                '-- 防止按鈕按二次 JOY 16/08/17
                '舊
                'BNG2.Attributes("onclick") = "Button('NG2', '" + BNG2.Value + "');"
                '新
                BNG2.Attributes("onclick") = "this.disabled = true;" & "Button('NG2', '" + BNG2.Value + "');"
                '--
                wNGSts2 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("OKDesc")
                '-- 防止按鈕按二次 JOY 16/08/17
                '舊
                'BOK.Attributes("onclick") = "Button('OK', '" + BOK.Value + "');"
                '新
                BOK.Attributes("onclick") = "this.disabled = true;" & "Button('OK', '" + BOK.Value + "');"
                '--
                wOKSts = DBDataSet1.Tables("M_Flow").Rows(0).Item("OKSts") + 1
            Else
                BOK.Visible = False
            End If
            '遲納管理
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("Delay") = 1 Then
                wDelay = 1
            End If
        End If

        If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                'Sheet顯示
                DAwaySheet1.Visible = True   '表單Sheet-1
                DDescSheet.Visible = True        '說明Sheet

                '遲納管理
                If wDelay = 1 Then
                    If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") < NowDateTime Then
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
                    Top = 576
                Else
                    DReasonCode.Visible = False     '延遲理由代碼
                    DReason.Visible = False         '延遲理由
                    DReasonDesc.Visible = False     '延遲其他說明
                    Top = 464
                End If

                '欄位顯示
                DDecideDesc.Visible = True      '說明
                '說明需輸入
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("FlowType") = 0 Then
                    DDecideDesc.BackColor = Color.LightGray
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "異常：需輸入說明")
                End If

                '連結顯示---需再修改

                '按鈕位置
                BNG1.Style.Add("Top", Top)     'NG1按鈕
                BNG2.Style.Add("Top", Top)     'NG2按鈕
                BSAVE.Style.Add("Top", Top)    '儲存按鈕
                BOK.Style.Add("Top", Top)      'OK按鈕
                '核定履歷
                DHistoryLabel.Style.Add("Top", Top + 24)  '核定履歷
                DataGrid9.Style.Add("Top", Top + 48)     '核定履歷
            End If
        Else
            Top = 384
            'Sheet隱藏
            DDescSheet.Visible = False  '說明Sheet
            DDelay.Visible = False      '延遲Sheet
            '欄位隱藏
            DDecideDesc.Visible = False    '說明
            DReasonCode.Visible = False    '延遲理由代碼
            DReason.Visible = False        '延遲理由
            DReasonDesc.Visible = False    '延遲其他說明
            DHistoryLabel.Visible = False  '核定履歷
            '連結顯示---需再修改

            '按鈕位置
            BNG1.Style.Add("Top", Top)     'NG1按鈕
            BNG2.Style.Add("Top", Top)     'NG2按鈕
            BSAVE.Style.Add("Top", Top)    '儲存按鈕
            BOK.Style.Add("Top", Top)      'OK按鈕
        End If
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        'Modify-Start by Joy 2009/11/20(2010行事曆對應)
        'BBStartDate.Attributes("onclick") = "CalendarPicker('" + wDepo + "', 'Form1.DBStartDate', 'Form1.DSalaryYM');"  
        'BBEndDate.Attributes("onclick") = "CalendarPicker('" + wDepo + "', 'Form1.DBEndDate', '');"                     
        'BAStartDate.Attributes("onclick") = "CalendarPicker('" + wDepo + "', 'Form1.DAStartDate', 'Form1.DSalaryYM');"  
        'BAEndDate.Attributes("onclick") = "CalendarPicker('" + wDepo + "', 'Form1.DAEndDate', '');"                     
        '
        BBStartDate.Attributes("onclick") = "CalendarPicker('" + wApplyCalendar + "', 'Form1.DBStartDate', 'Form1.DSalaryYM');"  '外出預定開始日期
        BBEndDate.Attributes("onclick") = "CalendarPicker('" + wApplyCalendar + "', 'Form1.DBEndDate', '');"                     '外出預定結束日期
        BAStartDate.Attributes("onclick") = "CalendarPicker('" + wApplyCalendar + "', 'Form1.DAStartDate', 'Form1.DSalaryYM');"  '外出實際開始日期
        BAEndDate.Attributes("onclick") = "CalendarPicker('" + wApplyCalendar + "', 'Form1.DAEndDate', '');"                     '外出實際結束日期
        'Modify-End
        BCardTime.Attributes("onclick") = "ShowCardTime();"    '刷卡記錄
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     按鈕及RequestedField的Top位置
    '**
    '*****************************************************************
    Sub TopPosition()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        '按鈕及RequestedField的Top位置
        If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") >= NowDateTime Then
                    Top = 464
                Else
                    If DDelay.Visible = True Then
                        Top = 576
                    Else
                        Top = 464
                    End If
                End If
            End If
        Else
            Top = 384
        End If
        '----
        'DB連結關閉
        OleDbConnection1.Close()
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
        '表單各欄位屬性及欄位輸入檢查等設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim DBDataSet1 As New DataSet
        Dim SQL As String
        Dim wEmpID, wJobTitle, wJobCode, wDivision, wDivisionCode, wDepoName, wDepoCode, wSalaryYM As String

        OleDbConnection1.Open()
        '取得申請者資訊
        SQL = "Select UserName, EmpID, JobName, JobID, DivName, DivID, DepoID, DepoName From M_Users "
        SQL = SQL & " Where UserID = '" & Request.QueryString("pUserID") & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
            wDepoCode = DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID")
            wDepoName = DBDataSet1.Tables("M_Users").Rows(0).Item("DepoName")
            'Delete-Start by Joy 2009/11/20(2010行事曆對應)
            'If DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID") = "01" Then wDepo = "CL"
            'If DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID") = "10" Then wDepo = "TP"
            'If DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID") = "11" Then wDepo = "TP"
            'If DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID") = "51" Then wDepo = "YA"
            'Delete-End
            wUserName = DBDataSet1.Tables("M_Users").Rows(0).Item("UserName")
            wEmpID = DBDataSet1.Tables("M_Users").Rows(0).Item("EmpID")
            wJobTitle = DBDataSet1.Tables("M_Users").Rows(0).Item("JobName")
            wJobCode = DBDataSet1.Tables("M_Users").Rows(0).Item("JobID")
            wDivision = DBDataSet1.Tables("M_Users").Rows(0).Item("DivName")
            wDivisionCode = DBDataSet1.Tables("M_Users").Rows(0).Item("DivID")
        End If
        OleDbConnection1.Close()
        '取得所屬年月
        If DateTime.Now.Month < 10 Then
            wSalaryYM = CStr(DateTime.Now.Year) + "/0" + CStr(DateTime.Now.Month)
        Else
            wSalaryYM = CStr(DateTime.Now.Year) + "/" + CStr(DateTime.Now.Month)
        End If
        '------------------------------------------------------------------------------------------
        'No
        Select Case FindFieldInf("No")
            Case 0  '顯示
                'DNo.BackColor = Color.LightGray
                DNo.BackColor = Color.White
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
        '申請日期
        Select Case FindFieldInf("Date")
            Case 0  '顯示
                DDate.BackColor = Color.LightGray
                DDate.ReadOnly = True
                DDate.Visible = True
            Case 1  '修改+檢查
                DDate.BackColor = Color.GreenYellow
                DDate.ReadOnly = True
                ShowRequiredFieldValidator("DDateRqd", "DDate", "異常：需輸入申請日期")
                DDate.Visible = True
            Case 2  '修改
                DDate.BackColor = Color.Yellow
                DDate.ReadOnly = True
                DDate.Visible = True
            Case Else   '隱藏
                DDate.Visible = False
        End Select
        If pPost = "New" Then DDate.Text = CStr(DateTime.Now.Today)
        '所屬年月
        Select Case FindFieldInf("SalaryYM")
            Case 0  '顯示
                DSalaryYM.BackColor = Color.LightGray
                DSalaryYM.ReadOnly = True
                DSalaryYM.Visible = True
            Case 1  '修改+檢查
                DSalaryYM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSalaryYMRqd", "DSalaryYM", "異常：需輸入所屬年月")
                DSalaryYM.ReadOnly = False
                DSalaryYM.Visible = True
            Case 2  '修改
                DSalaryYM.BackColor = Color.Yellow
                DSalaryYM.ReadOnly = False
                DSalaryYM.Visible = True
            Case Else   '隱藏
                DSalaryYM.Visible = False
        End Select
        If pPost = "New" Then DSalaryYM.Text = wSalaryYM
        '姓名
        Select Case FindFieldInf("Name")
            Case 0  '顯示
                DName.BackColor = Color.LightGray
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
        'EmpID
        Select Case FindFieldInf("EmpID")
            Case 0  '顯示
                DEmpID.BackColor = Color.LightGray
                DEmpID.ReadOnly = True
                DEmpID.Visible = True
            Case 1  '修改+檢查
                DEmpID.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEmpIDRqd", "DEmpID", "異常：需輸入卡號")
                DEmpID.Visible = True
            Case 2  '修改
                DEmpID.BackColor = Color.Yellow
                DEmpID.Visible = True
            Case Else   '隱藏
                DEmpID.Visible = False
        End Select
        If pPost = "New" Then DEmpID.Text = wEmpID
        '職稱
        Select Case FindFieldInf("JobTitle")
            Case 0  '顯示
                DJobTitle.BackColor = Color.LightGray
                DJobTitle.ReadOnly = True
                DJobTitle.Visible = True
            Case 1  '修改+檢查
                DJobTitle.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DJobTitleRqd", "DJobTitle", "異常：需輸入職稱")
                DJobTitle.Visible = True
            Case 2  '修改
                DJobTitle.BackColor = Color.Yellow
                DJobTitle.Visible = True
            Case Else   '隱藏
                DJobTitle.Visible = False
        End Select
        If pPost = "New" Then DJobTitle.Text = wJobTitle
        '職稱代碼
        Select Case FindFieldInf("JobCode")
            Case 0  '顯示
                DJobCode.BackColor = Color.LightGray
                DJobCode.ReadOnly = True
                DJobCode.Visible = True
            Case 1  '修改+檢查
                DJobCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DJobCodeRqd", "DJobCode", "異常：需輸入職稱代碼")
                DJobCode.Visible = True
            Case 2  '修改
                DJobCode.BackColor = Color.Yellow
                DJobCode.Visible = True
            Case Else   '隱藏
                DJobCode.Visible = False
        End Select
        If pPost = "New" Then DJobCode.Text = wJobCode
        'Depo Name
        Select Case FindFieldInf("DepoName")
            Case 0  '顯示
                DDepoName.BackColor = Color.LightGray
                DDepoName.ReadOnly = True
                DDepoName.Visible = True
            Case 1  '修改+檢查
                DDepoName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDepoNameRqd", "DDepoName", "異常：需輸入公司")
                DDepoName.Visible = True
            Case 2  '修改
                DDepoName.BackColor = Color.Yellow
                DDepoName.Visible = True
            Case Else   '隱藏
                DDepoName.Visible = False
        End Select
        If pPost = "New" Then DDepoName.Text = wDepoName
        'Depo Code
        Select Case FindFieldInf("DepoCode")
            Case 0  '顯示
                DDepoCode.BackColor = Color.LightGray
                DDepoCode.ReadOnly = True
                DDepoCode.Visible = True
            Case 1  '修改+檢查
                DDepoCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDepoCodeRqd", "DDepoCode", "異常：需輸入公司代碼")
                DDepoCode.Visible = True
            Case 2  '修改
                DDepoCode.BackColor = Color.Yellow
                DDepoCode.Visible = True
            Case Else   '隱藏
                DDepoCode.Visible = False
        End Select
        If pPost = "New" Then DDepoCode.Text = wDepoCode
        '部門
        Select Case FindFieldInf("Division")
            Case 0  '顯示
                DDivision.BackColor = Color.LightGray
                DDivision.ReadOnly = True
                DDivision.Visible = True
            Case 1  '修改+檢查
                DDivision.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDivisionRqd", "DDivision", "異常：需輸入部門")
                DDivision.Visible = True
            Case 2  '修改
                DDivision.BackColor = Color.Yellow
                DDivision.Visible = True
            Case Else   '隱藏
                DDivision.Visible = False
        End Select
        If pPost = "New" Then DDivision.Text = wDivision
        '部門代碼
        Select Case FindFieldInf("DivisionCode")
            Case 0  '顯示
                DDivisionCode.BackColor = Color.LightGray
                DDivisionCode.ReadOnly = True
                DDivisionCode.Visible = True
            Case 1  '修改+檢查
                DDivisionCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDivisionCodeRqd", "DDivisionCode", "異常：需輸入部門代碼")
                DDivisionCode.Visible = True
            Case 2  '修改
                DDivisionCode.BackColor = Color.Yellow
                DDivisionCode.Visible = True
            Case Else   '隱藏
                DDivisionCode.Visible = False
        End Select
        If pPost = "New" Then DDivisionCode.Text = wDivisionCode
        '目的地
        Select Case FindFieldInf("Place")
            Case 0  '顯示
                DPlace.BackColor = Color.LightGray
                DPlace.Visible = True
            Case 1  '修改+檢查
                DPlace.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPlaceRqd", "DPlace", "異常：需輸入目的地")
                DPlace.Visible = True
            Case 2  '修改
                DPlace.BackColor = Color.Yellow
                DPlace.Visible = True
            Case Else   '隱藏
                DPlace.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Place", "ZZZZZZ")
        '目的地(其他)
        Select Case FindFieldInf("OtherPlace")
            Case 0  '顯示
                DOtherPlace.BackColor = Color.LightGray
                DOtherPlace.Visible = True
            Case 1  '修改+檢查
                DOtherPlace.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOtherPlaceRqd", "DOtherPlace", "異常：需輸入目的地")
                DOtherPlace.Visible = True
            Case 2  '修改
                DOtherPlace.BackColor = Color.Yellow
                DOtherPlace.Visible = True
            Case Else   '隱藏
                DOtherPlace.Visible = False
        End Select
        If pPost = "New" Then DOtherPlace.Text = ""
        '職務代理人
        Select Case FindFieldInf("JobAgent")
            Case 0  '顯示
                DJobAgent.BackColor = Color.LightGray
                DJobAgent.Visible = True
            Case 1  '修改+檢查
                DJobAgent.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DJobAgentRqd", "DJobAgent", "異常：需輸入職務代理人")
                DJobAgent.Visible = True
            Case 2  '修改
                DJobAgent.BackColor = Color.Yellow
                DJobAgent.Visible = True
            Case Else   '隱藏
                DJobAgent.Visible = False
        End Select
        If pPost = "New" Then DJobAgent.Text = ""
        '預定開始日期
        Select Case FindFieldInf("BStartDate")
            Case 0  '顯示
                DBStartDate.BackColor = Color.LightGray
                DBStartDate.ReadOnly = True
                DBStartDate.Visible = True
                BBStartDate.Visible = False
            Case 1  '修改+檢查
                DBStartDate.BackColor = Color.GreenYellow
                DBStartDate.ReadOnly = True
                ShowRequiredFieldValidator("DBStartDateRqd", "DBStartDate", "異常：需輸入預定開始日期")
                DBStartDate.Visible = True
                BBStartDate.Visible = True
            Case 2  '修改
                DBStartDate.BackColor = Color.Yellow
                DBStartDate.ReadOnly = True
                DBStartDate.Visible = True
                BBStartDate.Visible = True
            Case Else   '隱藏
                DBStartDate.Visible = False
                BBStartDate.Visible = False
        End Select
        If pPost = "New" Then DBStartDate.Text = CStr(DateTime.Now.Today)
        '預定開始-時
        Select Case FindFieldInf("BStartH")
            Case 0  '顯示
                DBStartH.BackColor = Color.LightGray
                DBStartH.Visible = True
                BBDays.Visible = False
            Case 1  '修改+檢查
                DBStartH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBStartHRqd", "DBStartH", "異常：需輸入預定開始-時")
                DBStartH.Visible = True
                BBDays.Visible = True
            Case 2  '修改
                DBStartH.BackColor = Color.Yellow
                DBStartH.Visible = True
                BBDays.Visible = True
            Case Else   '隱藏
                DBStartH.Visible = False
                BBDays.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BStartH", "ZZZZZZ")
        '預定結束日期
        Select Case FindFieldInf("BEndDate")
            Case 0  '顯示
                DBEndDate.BackColor = Color.LightGray
                DBEndDate.ReadOnly = True
                DBEndDate.Visible = True
                BBEndDate.Visible = False
            Case 1  '修改+檢查
                DBEndDate.BackColor = Color.GreenYellow
                DBEndDate.ReadOnly = True
                ShowRequiredFieldValidator("DBEndDateRqd", "DBEndDate", "異常：需輸入預定結束日期")
                DBEndDate.Visible = True
                BBEndDate.Visible = True
            Case 2  '修改
                DBEndDate.BackColor = Color.Yellow
                DBEndDate.ReadOnly = True
                DBEndDate.Visible = True
                BBEndDate.Visible = True
            Case Else   '隱藏
                DBEndDate.Visible = False
                BBEndDate.Visible = False
        End Select
        If pPost = "New" Then DBEndDate.Text = CStr(DateTime.Now.Today)
        '預定結束-時
        Select Case FindFieldInf("BEndH")
            Case 0  '顯示
                DBEndH.BackColor = Color.LightGray
                DBEndH.Visible = True
            Case 1  '修改+檢查
                DBEndH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBEndHRqd", "DBEndH", "異常：需輸入預定結束-時")
                DBEndH.Visible = True
            Case 2  '修改
                DBEndH.BackColor = Color.Yellow
                DBEndH.Visible = True
            Case Else   '隱藏
                DBEndH.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BEndH", "ZZZZZZ")
        '預定天數
        Select Case FindFieldInf("BDay")
            Case 0  '顯示
                DBDay.BackColor = Color.LightGray
                DBDay.Visible = True
            Case 1  '修改+檢查
                DBDay.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBDayRqd", "DBDay", "異常：需輸入預定天數")
                DBDay.Visible = True
            Case 2  '修改
                DBDay.BackColor = Color.Yellow
                DBDay.Visible = True
            Case Else   '隱藏
                DBDay.Visible = False
        End Select
        If pPost = "New" Then DBDay.Text = "0"
        '預定時數
        Select Case FindFieldInf("BHour")
            Case 0  '顯示
                DBHour.BackColor = Color.LightGray
                DBHour.Visible = True
            Case 1  '修改+檢查
                DBHour.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBHourRqd", "DBHour", "異常：需輸入預定時數")
                DBHour.Visible = True
            Case 2  '修改
                DBHour.BackColor = Color.Yellow
                DBHour.Visible = True
            Case Else   '隱藏
                DBHour.Visible = False
        End Select
        If pPost = "New" Then DBHour.Text = "0"
        '實際開始日期
        Select Case FindFieldInf("AStartDate")
            Case 0  '顯示
                DAStartDate.BackColor = Color.LightGray
                DAStartDate.ReadOnly = True
                DAStartDate.Visible = True
                BAStartDate.Visible = False
            Case 1  '修改+檢查
                DAStartDate.BackColor = Color.GreenYellow
                DAStartDate.ReadOnly = True
                ShowRequiredFieldValidator("DAStartDateRqd", "DAStartDate", "異常：需輸入實際開始日期")
                DAStartDate.Visible = True
                BAStartDate.Visible = True
            Case 2  '修改
                DAStartDate.BackColor = Color.Yellow
                DAStartDate.ReadOnly = True
                DAStartDate.Visible = True
                BAStartDate.Visible = True
            Case Else   '隱藏
                DAStartDate.Visible = False
                BAStartDate.Visible = False
        End Select
        If pPost = "New" Then DAStartDate.Text = CStr(DateTime.Now.Today)
        '實際開始-時
        Select Case FindFieldInf("AStartH")
            Case 0  '顯示
                DAStartH.BackColor = Color.LightGray
                DAStartH.Visible = True
                BADays.Visible = False
            Case 1  '修改+檢查
                DAStartH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAStartHRqd", "DAStartH", "異常：需輸入實際開始-時")
                DAStartH.Visible = True
                BADays.Visible = True
            Case 2  '修改
                DAStartH.BackColor = Color.Yellow
                DAStartH.Visible = True
                BADays.Visible = True
            Case Else   '隱藏
                DAStartH.Visible = False
                BADays.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AStartH", "ZZZZZZ")
        '實際結束日期
        Select Case FindFieldInf("AEndDate")
            Case 0  '顯示
                DAEndDate.BackColor = Color.LightGray
                DAEndDate.ReadOnly = True
                DAEndDate.Visible = True
                BAEndDate.Visible = False
            Case 1  '修改+檢查
                DAEndDate.BackColor = Color.GreenYellow
                DAEndDate.ReadOnly = True
                ShowRequiredFieldValidator("DAEndDateRqd", "DAEndDate", "異常：需輸入實際結束日期")
                DAEndDate.Visible = True
                BAEndDate.Visible = True
            Case 2  '修改
                DAEndDate.BackColor = Color.Yellow
                DAEndDate.ReadOnly = True
                DAEndDate.Visible = True
                BAEndDate.Visible = True
            Case Else   '隱藏
                DAEndDate.Visible = False
                BAEndDate.Visible = False
        End Select
        If pPost = "New" Then DAEndDate.Text = CStr(DateTime.Now.Today)
        '實際結束-時
        Select Case FindFieldInf("AEndH")
            Case 0  '顯示
                DAEndH.BackColor = Color.LightGray
                DAEndH.Visible = True
            Case 1  '修改+檢查
                DAEndH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAEndHRqd", "DAEndH", "異常：需輸入實際結束-時")
                DAEndH.Visible = True
            Case 2  '修改
                DAEndH.BackColor = Color.Yellow
                DAEndH.Visible = True
            Case Else   '隱藏
                DAEndH.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AEndH", "ZZZZZZ")
        '實際天數
        Select Case FindFieldInf("ADay")
            Case 0  '顯示
                DADay.BackColor = Color.LightGray
                DADay.Visible = True
            Case 1  '修改+檢查
                DADay.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DADayRqd", "DADay", "異常：需輸入實際天數")
                DADay.Visible = True
            Case 2  '修改
                DADay.BackColor = Color.Yellow
                DADay.Visible = True
            Case Else   '隱藏
                DADay.Visible = False
        End Select
        If pPost = "New" Then DADay.Text = "0"
        '實際時數
        Select Case FindFieldInf("AHour")
            Case 0  '顯示
                DAHour.BackColor = Color.LightGray
                DAHour.Visible = True
            Case 1  '修改+檢查
                DAHour.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAHourRqd", "DAHour", "異常：需輸入實際時數")
                DAHour.Visible = True
            Case 2  '修改
                DAHour.BackColor = Color.Yellow
                DAHour.Visible = True
            Case Else   '隱藏
                DAHour.Visible = False
        End Select
        If pPost = "New" Then DAHour.Text = "0"
        '外出理由
        Select Case FindFieldInf("FReason")
            Case 0  '顯示
                DFReason.BackColor = Color.LightGray
                DFReason.ReadOnly = True
                DFReason.Visible = True
            Case 1  '修改+檢查
                DFReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFReasonRqd", "DFReason", "異常：需輸入外出理由")
                DFReason.Visible = True
            Case 2  '修改
                DFReason.BackColor = Color.Yellow
                DFReason.Visible = True
            Case Else   '隱藏
                DFReason.Visible = False
        End Select
        If pPost = "New" Then DFReason.Text = ""
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
        rqdVal.ErrorMessage = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", CStr(Top))
        rqdVal.Style.Add("Left", "8")
        rqdVal.Style.Add("Height", "20")
        rqdVal.Style.Add("Width", "250")
        rqdVal.Style.Add("POSITION", "absolute")
        Page.Controls(1).Controls.Add(rqdVal)
    End Sub

    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim idx As Integer
        Dim i As Integer
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        idx = FindFieldInf(pFieldName)

        '姓名
        If pFieldName = "Name" Then
            DName.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DName.Items.Add(ListItem1)
                End If
            Else
                '登入者
                Dim ListItem1 As New ListItem
                ListItem1.Text = wUserName
                ListItem1.Value = Request.QueryString("pUserID")
                ListItem1.Selected = True
                DName.Items.Add(ListItem1)
                '全表單代理
                SQL = "Select UserName, UserID From M_Agent  "
                SQL = SQL + "Where Active = '1' "
                SQL = SQL + "  And AllForm = '0' "
                SQL = SQL + "  And AgentID = '" + Request.QueryString("pUserID") + "' "
                SQL = SQL + "  And StartDate <= '" + NowDateTime + "' "
                SQL = SQL + "  And EndDate >= '" + NowDateTime + "' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Agent")
                DBTable1 = DBDataSet1.Tables("M_Agent")
                If DBTable1.Rows.Count <= 0 Then
                    '單一表單代理
                    DBDataSet1.Clear()
                    SQL = "Select UserName, UserID From M_Agent  "
                    SQL = SQL + "Where Active = '1' "
                    SQL = SQL + "  And AllForm = '1' "
                    SQL = SQL + "  And AgentID = '" + Request.QueryString("pUserID") + "' "
                    SQL = SQL + "  And StartDate <= '" + NowDateTime + "' "
                    SQL = SQL + "  And EndDate >= '" + NowDateTime + "' "
                    SQL = SQL + "  And FormNo = '001003' "
                    Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter2.Fill(DBDataSet1, "M_Agent")
                    DBTable1 = DBDataSet1.Tables("M_Agent")
                    If DBTable1.Rows.Count <= 0 Then
                    Else
                        For i = 0 To DBTable1.Rows.Count - 1
                            Dim ListItem2 As New ListItem
                            ListItem2.Text = DBTable1.Rows(i).Item("UserName")
                            ListItem2.Value = DBTable1.Rows(i).Item("UserID")
                            DName.Items.Add(ListItem2)
                        Next
                    End If
                Else
                    For i = 0 To DBTable1.Rows.Count - 1
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("UserName")
                        ListItem2.Value = DBTable1.Rows(i).Item("UserID")
                        DName.Items.Add(ListItem2)
                    Next
                End If
            End If
        End If
        '目的地
        If pFieldName = "Place" Then
            DPlace.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPlace.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='1003' and DKey='PLACE' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPlace.Items.Add(ListItem1)
                Next
            End If
        End If
        '預定開始-時
        If pFieldName = "BStartH" Then
            DBStartH.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBStartH.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1003' and DKey='HOUR' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBStartH.Items.Add(ListItem1)
                Next
            End If
        End If
        '預定結束-時
        If pFieldName = "BEndH" Then
            DBEndH.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBEndH.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1003' and DKey='HOUR' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBEndH.Items.Add(ListItem1)
                Next
            End If
        End If
        '實際開始-時
        If pFieldName = "AStartH" Then
            DAStartH.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1003' and DKey='HOUR' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAStartH.Items.Add(ListItem1)
                Next
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1003' and DKey='HOUR' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAStartH.Items.Add(ListItem1)
                Next
            End If
        End If
        '實際結束-時
        If pFieldName = "AEndH" Then
            DAEndH.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1003' and DKey='HOUR' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAEndH.Items.Add(ListItem1)
                Next
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1003' and DKey='HOUR' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAEndH.Items.Add(ListItem1)
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
                SQL = "Select * From M_Referp Where Cat='014' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("DKey")
                    ListItem1.Value = DBTable1.Rows(i).Item("DKey")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DReasonCode.Items.Add(ListItem1)
                Next
            End If
        End If
        '延遲理由
        If pFieldName = "Reason" Then
            SQL = "Select * From M_Referp Where Cat='014' and DKey= '" & DReasonCode.SelectedValue & "' Order by Data "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                DReason.Text = DBTable1.Rows(i).Item("Data")
            Next
        End If
        'DB連結關閉
        OleDbConnection1.Close()
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
    '**
    '**     儲存按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BSAVE_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSAVE.ServerClick
        If Request.Cookies("RunBSAVE").Value = True Then
            If InputCheck() = 0 Then
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
                    Response.Write(YKK.ShowMessage(Message))
                End If      '上傳檔案ErrCode=0

                If ErrCode = 0 Then
                    Dim URL As String = "http://10.245.1.6/WorkFlowSub/MessagePage.aspx?pMSGID=S&pFormNo=" & wFormNo & "&pStep=" & wStep & _
                                                        "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                    Response.Redirect(URL)
                End If
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BOK_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
        If Request.Cookies("RunBOK").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '停止Button運作
                FlowControl("OK", 0, "1")
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-1按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BNG1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG1.ServerClick
        If Request.Cookies("RunBNG1").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '停止Button運作
                FlowControl("NG1", 1, "2")
            End If
        End If

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-2按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BNG2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG2.ServerClick
        If Request.Cookies("RunBNG2").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '停止Button運作
                FlowControl("NG2", 2, "3")
            End If
        End If
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
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim Message As String = ""
        Dim wQCLT As Integer = 0 'QC-L/T

        GetDataStatus()  '取得此關卡各按鈕的Data Status

        '--------------------------------------------------------------------------
        '--  檢查表單各欄位資料
        '--------------------------------------------------------------------------
        '外出天數,時數
        If ErrCode = 0 Then
            If DBDay.Text = "0" And DBHour.Text = "0" Then ErrCode = 9051
            If DADay.Text = "0" And DAHour.Text = "0" Then ErrCode = 9051
        End If
        '目的地
        If ErrCode = 0 Then
            If Left(DPlace.SelectedValue, 2) <= "10" Or Left(DPlace.SelectedValue, 2) = "99" Then
                If DOtherPlace.Text = "" Then
                    ErrCode = 9052
                End If
            End If
        End If
        '起迄日期檢查
        If ErrCode = 0 Then
            '預定
            If CDate(DBEndDate.Text) < CDate(DBStartDate.Text) Then ErrCode = 8010
            If CDate(DBEndDate.Text) = CDate(DBStartDate.Text) Then
                If DBEndH.SelectedValue <= DBStartH.SelectedValue Then ErrCode = 8010
            End If
            '實際
            If CDate(DAEndDate.Text) < CDate(DAStartDate.Text) Then ErrCode = 8010
            If CDate(DAEndDate.Text) = CDate(DAStartDate.Text) Then
                If DAEndH.SelectedValue <= DAStartH.SelectedValue Then ErrCode = 8010
            End If
        End If
        '起迄日期同月檢查
        If ErrCode = 0 Then
            '預定
            If CStr(Year(CDate(DBStartDate.Text))) + "/" + CStr(Month(CDate(DBStartDate.Text))) <> CStr(Year(CDate(DBEndDate.Text))) + "/" + CStr(Month(CDate(DBEndDate.Text))) Then
                ErrCode = 8020
            End If
            '實際
            If CStr(Year(CDate(DAStartDate.Text))) + "/" + CStr(Month(CDate(DAStartDate.Text))) <> CStr(Year(CDate(DAEndDate.Text))) + "/" + CStr(Month(CDate(DAEndDate.Text))) Then
                ErrCode = 8020
            End If
        End If
        '日數結算結果檢查
        If ErrCode = 0 Then
            '#預定
            Dim Hours As Integer = 0
            Dim VHours As Integer = 0
            Dim MHours As Integer = 0
            Dim wTotal As Integer = 0

            Dim SQL As String
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

            OleDbConnection1.Open()
            SQL = "Select Count(*) As Days From M_Vacation "
            SQL = SQL + "Where Active = '1' "
            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
            'SQL = SQL + "  and Depo = '" + wDepo + "' "
            '
            SQL = SQL + "  and Depo = '" + wApplyCalendar + "' "
            'Modify-End
            SQL = SQL + "  and YMD  >= '" + DBStartDate.Text + "' "
            SQL = SQL + "  and YMD  <= '" + DBEndDate.Text + "' "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "Vacation")
            MHours = DBDataSet1.Tables("Vacation").Rows(0).Item("Days") * 8
            OleDbConnection1.Close()
            '外出時間
            If CDate(DBEndDate.Text) > CDate(DBStartDate.Text) Then
                VHours = DateDiff("d", CDate(DBStartDate.Text), CDate(DBEndDate.Text)) * 8
            End If
            '外出時數差時間
            If DBEndH.SelectedValue >= DBStartH.SelectedValue Then
                If (DBStartH.SelectedValue <= "12" And DBEndH.SelectedValue <= "12") Then
                    Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue)
                Else
                    If (DBStartH.SelectedValue >= "13" And DBEndH.SelectedValue >= "13") Then
                        Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue)
                    Else
                        Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue) - 1
                    End If
                End If
                If Hours > 8 Then Hours = 8
            Else
                If (DBStartH.SelectedValue <= "12" And DBEndH.SelectedValue <= "12") Then
                    Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue)
                Else
                    If (DBStartH.SelectedValue >= "13" And DBEndH.SelectedValue >= "13") Then
                        Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue)
                    Else
                        Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue) + 1
                    End If
                End If
                If Hours > 8 Then Hours = 8
            End If
            '計算時間
            wTotal = VHours + Hours - MHours
            If DBDay.Text <> CStr(Fix(wTotal / 8)) Then
                ErrCode = 8040
            End If
            If DBHour.Text <> CStr(wTotal - (Fix(wTotal / 8) * 8)) Then
                ErrCode = 8040
            End If
            DBDataSet1.Clear()
            '----------------------------------------------------------------------------------------------
            '#實際
            Hours = 0
            VHours = 0
            MHours = 0
            wTotal = 0
            OleDbConnection1.Open()
            SQL = "Select Count(*) As Days From M_Vacation "
            SQL = SQL + "Where Active = '1' "
            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
            'SQL = SQL + "  and Depo = '" + wDepo + "' "
            '
            SQL = SQL + "  and Depo = '" + wApplyCalendar + "' "
            'Modify-End
            SQL = SQL + "  and YMD  >= '" + DAStartDate.Text + "' "
            SQL = SQL + "  and YMD  <= '" + DAEndDate.Text + "' "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "Vacation")
            MHours = DBDataSet1.Tables("Vacation").Rows(0).Item("Days") * 8
            OleDbConnection1.Close()
            '外出時間
            If CDate(DAEndDate.Text) > CDate(DAStartDate.Text) Then
                VHours = DateDiff("d", CDate(DAStartDate.Text), CDate(DAEndDate.Text)) * 8
            End If
            '外出時數差時間
            If DAEndH.SelectedValue >= DAStartH.SelectedValue Then
                If (DAStartH.SelectedValue <= "12" And DAEndH.SelectedValue <= "12") Then
                    Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue)
                Else
                    If (DAStartH.SelectedValue >= "13" And DAEndH.SelectedValue >= "13") Then
                        Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue)
                    Else
                        Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue) - 1
                    End If
                End If
                If Hours > 8 Then Hours = 8
            Else
                If (DAStartH.SelectedValue <= "12" And DAEndH.SelectedValue <= "12") Then
                    Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue)
                Else
                    If (DAStartH.SelectedValue >= "13" And DAEndH.SelectedValue >= "13") Then
                        Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue)
                    Else
                        Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue) + 1
                    End If
                End If
                If Hours > 8 Then Hours = 8
            End If
            '計算時間
            wTotal = VHours + Hours - MHours
            If DADay.Text <> CStr(Fix(wTotal / 8)) Then
                ErrCode = 8040
            End If
            If DAHour.Text <> CStr(wTotal - (Fix(wTotal / 8) * 8)) Then
                ErrCode = 8040
            End If
        End If
        '--------------------------------------------------------------------------
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

                ErrCode = oCommon.CommissionNo("001003", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No

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
                Dim SQL As String

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
                    Dim wAllocateID As String = ""

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
                    Response.Write(YKK.ShowMessage(Message))
                End If      '儲存表單ErrCode=0
            End While       '重覆執行

            If ErrCode = 0 Then
                '--郵件傳送---------
                oCommon.SendMail()

                Dim URL As String = "http://10.245.1.6/WorkFlowSub/MessagePage.aspx?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                    "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
            End If
        Else
            EnabledButton()   '起動Button運作
            If ErrCode = 9010 Then Message = "上傳檔案格式錯誤,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案Size超過1024KB,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案沒指定,請確認!"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9051 Then Message = "外出時間需填寫,請確認!"
            If ErrCode = 9052 Then Message = "目的地(客戶,外注商,其他)時需填寫內容,請確認!"
            If ErrCode = 8010 Then Message = "結束日期不可小於開始日期,請確認!"
            If ErrCode = 8020 Then Message = "開始日期與結束日期不可跨月,請確認!"
            If ErrCode = 8040 Then Message = "外出結算時間不符合,請確認!"
            If ErrCode = 9060 Then Message = "委託書No.重覆,請確認委託書No.!"
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
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Flow")
        If DBDataSet1.Tables("M_Flow").Rows.Count > 0 Then
            'NG-1按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun1") = 1 Then
                wNGSts1 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts1") + 1
            End If
            'NG-2按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun2") = 1 Then
                wNGSts2 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts2") + 1
            End If
            'OK按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("OKFun") = 1 Then
                wOKSts = DBDataSet1.Tables("M_Flow").Rows(0).Item("OKSts") + 1
            End If
        End If
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     新增表單資料
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("AwayFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Insert into F_AwaySheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, No, "                    '1~5
        SQl = SQl + "Date, SalaryYM, Name, EmpID, "                                '6~9
        SQl = SQl + "JobTitle, JobCode, DepoName, DepoCode, Division, "            '10~14
        SQl = SQl + "DivisionCode, Place, OtherPlace, JobAgent, "                  '15~18
        SQl = SQl + "BStartDate, BStartH, BEndDate, BEndH, BDay, BHour, "          '19~24
        SQl = SQl + "AStartDate, AStartH, AEndDate, AEndH, ADay, AHour, "          '25~30
        SQl = SQl + "FReason, "                                                    '31
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "              '32~35
        SQl = SQl + ")  "
        SQl = SQl + "VALUES( "
        '1~5
        SQl = SQl + " '0', "                                '狀態(0:未結,1:已結NG,2:已結OK)
        SQl = SQl + " '" + NowDateTime + "', "              '結案日
        SQl = SQl + " '001003', "                           '表單代號
        SQl = SQl + " '" + CStr(NewFormSno) + "', "         '表單流水號
        SQl = SQl + " '" + DNo.Text + "', "                 'NO
        '6~9
        SQl = SQl + " '" + DDate.Text + "', "               '申請日期
        SQl = SQl + " '" + DSalaryYM.Text + "', "           '所屬年月
        SQl = SQl + " N'" + DName.SelectedItem.Text + "', " '姓名
        SQl = SQl + " '" + DEmpID.Text + "', "              'EMP-ID
        '10~15
        SQl = SQl + " N'" + DJobTitle.Text + "', "          '職稱
        SQl = SQl + " '" + DJobCode.Text + "', "            '職稱代碼
        SQl = SQl + " N'" + DDepoName.Text + "', "          '公司別
        SQl = SQl + " '" + DDepoCode.Text + "', "           '公司別Code
        SQl = SQl + " N'" + DDivision.Text + "', "          '部門
        SQl = SQl + " '" + DDivisionCode.Text + "', "       '部門代碼
        '16~18
        SQl = SQl + " N'" + DPlace.SelectedValue + "', "    '目的地
        SQl = SQl + " N'" + DOtherPlace.Text + "', "        '目的地-其他
        SQl = SQl + " N'" + DJobAgent.Text + "', "          '職務代理人
        '19~24
        SQl = SQl + " '" + DBStartDate.Text + "', "                       '預定-開始日期
        SQl = SQl + " '" + CStr(CInt(DBStartH.SelectedValue)) + "', "     '預定-開始時
        SQl = SQl + " '" + DBEndDate.Text + "', "                         '預定-結束日期
        SQl = SQl + " '" + CStr(CInt(DBEndH.SelectedValue)) + "', "       '預定-結束時
        SQl = SQl + " '" + CStr(CInt(DBDay.Text)) + "', "                 '預定-天數
        SQl = SQl + " '" + CStr(CInt(DBHour.Text)) + "', "                '預定-時數
        '25~30
        SQl = SQl + " '" + DAStartDate.Text + "', "                       '實際-開始日期
        SQl = SQl + " '" + CStr(CInt(DAStartH.SelectedValue)) + "', "     '實際-開始時
        SQl = SQl + " '" + DAEndDate.Text + "', "                         '實際-結束日期
        SQl = SQl + " '" + CStr(CInt(DAEndH.SelectedValue)) + "', "       '實際-結束時
        SQl = SQl + " '" + CStr(CInt(DADay.Text)) + "', "                 '實際-天數
        SQl = SQl + " '" + CStr(CInt(DAHour.Text)) + "', "                '實際-時數
        '31
        SQl = SQl + " N'" + YKK.ReplaceString(DFReason.Text) + "', "      '外出理由
        '--------------------------------------------
        SQl = SQl + " '" + Request.QueryString("pUserID") + "', "        '作成者
        SQl = SQl + " '" + NowDateTime + "', "                            '作成時間
        SQl = SQl + " '" + "" + "', "                                     '修改者
        SQl = SQl + " '" + NowDateTime + "' "                             '修改時間
        SQl = SQl + " ) "
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("AwayFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Update F_AwaySheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = '" & DNo.Text & "',"
        SQl = SQl + " Date = '" & DDate.Text & "',"
        SQl = SQl + " SalaryYM = '" & DSalaryYM.Text & "',"
        SQl = SQl + " Name = N'" & DName.SelectedItem.Text & "',"
        SQl = SQl + " EmpID = '" & DEmpID.Text & "',"

        SQl = SQl + " JobTitle = N'" & DJobTitle.Text & "',"
        SQl = SQl + " JobCode = '" & DJobCode.Text & "',"
        SQl = SQl + " DepoName = N'" & DDepoName.Text & "',"
        SQl = SQl + " DepoCode = '" & DDepoCode.Text & "',"
        SQl = SQl + " Division = N'" & DDivision.Text & "',"
        SQl = SQl + " DivisionCode = '" & DDivisionCode.Text & "',"

        SQl = SQl + " Place = N'" & DPlace.SelectedValue & "',"
        SQl = SQl + " OtherPlace = N'" & DOtherPlace.Text & "',"
        SQl = SQl + " JobAgent = N'" & DJobAgent.Text & "',"

        SQl = SQl + " BStartDate = '" & DBStartDate.Text & "',"
        SQl = SQl + " BStartH = '" & CStr(CInt(DBStartH.SelectedValue)) & "',"
        SQl = SQl + " BEndDate = '" & DBEndDate.Text & "',"
        SQl = SQl + " BEndH = '" & CStr(CInt(DBEndH.SelectedValue)) & "',"
        SQl = SQl + " BDay = '" & CStr(CInt(DBDay.Text)) & "',"
        SQl = SQl + " BHour = '" & CStr(CInt(DBHour.Text)) & "',"

        SQl = SQl + " AStartDate = '" & DAStartDate.Text & "',"
        SQl = SQl + " AStartH = '" & CStr(CInt(DAStartH.SelectedValue)) & "',"
        SQl = SQl + " AEndDate = '" & DAEndDate.Text & "',"
        SQl = SQl + " AEndH = '" & CStr(CInt(DAEndH.SelectedValue)) & "',"
        SQl = SQl + " ADay = '" & CStr(CInt(DADay.Text)) & "',"
        SQl = SQl + " AHour = '" & CStr(CInt(DAHour.Text)) & "',"

        SQl = SQl + " FReason = N'" & YKK.ReplaceString(DFReason.Text) & "',"

        SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
        SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
        SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()
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

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        If pFun <> "SAVE" Then      '<> Save
            SQl = "Update T_WaitHandle Set "
            SQl = SQl + " Active = '" & "0" & "',"
            SQl = SQl + " Sts = '" & pSts & "',"
            If pSts = "1" Then SQl = SQl + " StsDes = '" & BOK.Value & "',"
            If pSts = "2" Then SQl = SQl + " StsDes = '" & BNG1.Value & "',"
            If pSts = "3" Then SQl = SQl + " StsDes = '" & BNG2.Value & "',"

            SQl = SQl + " AEndTime = '" & NowDateTime & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
            If DDelay.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(DDecideDesc.Text) & "',"
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
            SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step =  '" & CStr(wStep) & "'"
            SQl = SQl + "   And SeqNo =  '" & CStr(wSeqNo) & "'"
        End If
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Add T_CommissionNo)
    '**     追加交易資料和委託單對照表
    '**
    '*****************************************************************
    Sub AddCommissionNo(ByVal pFormNo As String, ByVal pFormSno As Integer)
        Dim SQl As String
        Dim DBDataSet1 As New DataSet

        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand
        OleDbConnection1.Open()

        SQl = "Select * From T_CommissionNo "
        SQl = SQl & " Where FormNo =  '" & pFormNo & "'"
        SQl = SQl & "   And FormSno =  '" & CStr(pFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQl, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "T_CommissionNo")

        If DBDataSet1.Tables("T_CommissionNo").Rows.Count <= 0 Then
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
                OleDBCommand1.Connection = OleDbConnection1
                OleDBCommand1.CommandText = SQl
                OleDBCommand1.ExecuteNonQuery()
            End If
        Else
            If DNo.Text <> "" Then
                If DNo.Text <> DBDataSet1.Tables("T_CommissionNo").Rows(0).Item("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNo.Text & "',"
                    SQl = SQl + " MapNo = '" & "" & "',"
                    SQl = SQl + " CreateUser = '" & Request.QueryString("pUserID") & "',"
                    SQl = SQl + " CreateTime = '" & NowDateTime & "' "
                    SQl = SQl + " Where FormNo  =  '" & pFormNo & "'"
                    SQl = SQl + "   And FormSno =  '" & CStr(pFormSno) & "'"
                    OleDBCommand1.Connection = OleDbConnection1
                    OleDBCommand1.CommandText = SQl
                    OleDBCommand1.ExecuteNonQuery()
                End If
            End If
        End If

        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     →Button計算預定日數 
    '**
    '*****************************************************************
    Private Sub BBDays_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BBDays.Click
        Dim Hours As Integer = 0
        Dim VHours As Integer = 0
        Dim MHours As Integer = 0
        Dim wTotal As Integer = 0
        Dim ErrCode As Integer = 0
        Dim SQL, Message As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        '起迄日期檢查
        If ErrCode = 0 Then
            If CDate(DBEndDate.Text) < CDate(DBStartDate.Text) Then ErrCode = 8010
            If CDate(DBEndDate.Text) = CDate(DBStartDate.Text) Then
                If DBEndH.SelectedValue <= DBStartH.SelectedValue Then ErrCode = 8010
            End If
        End If
        '起迄日期同月檢查
        If ErrCode = 0 Then
            If CStr(Year(CDate(DBStartDate.Text))) + "/" + CStr(Month(CDate(DBStartDate.Text))) <> CStr(Year(CDate(DBEndDate.Text))) + "/" + CStr(Month(CDate(DBEndDate.Text))) Then
                ErrCode = 8020
            End If
        End If
        '假日時間
        If ErrCode = 0 Then
            OleDbConnection1.Open()
            SQL = "Select Count(*) As Days From M_Vacation "
            SQL = SQL + "Where Active = '1' "
            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
            'SQL = SQL + "  and Depo = '" + wDepo + "' "
            '
            SQL = SQL + "  and Depo = '" + wApplyCalendar + "' "
            'Modify-End
            SQL = SQL + "  and YMD  >= '" + DBStartDate.Text + "' "
            SQL = SQL + "  and YMD  <= '" + DBEndDate.Text + "' "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "Vacation")
            MHours = DBDataSet1.Tables("Vacation").Rows(0).Item("Days") * 8
            OleDbConnection1.Close()
            '外出時間
            If CDate(DBEndDate.Text) > CDate(DBStartDate.Text) Then
                VHours = DateDiff("d", CDate(DBStartDate.Text), CDate(DBEndDate.Text)) * 8
            End If
            '外出時數差時間
            If DBEndH.SelectedValue >= DBStartH.SelectedValue Then
                If (DBStartH.SelectedValue <= "12" And DBEndH.SelectedValue <= "12") Then
                    Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue)
                Else
                    If (DBStartH.SelectedValue >= "13" And DBEndH.SelectedValue >= "13") Then
                        Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue)
                    Else
                        Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue) - 1
                    End If
                End If
                If Hours > 8 Then Hours = 8
            Else
                If (DBStartH.SelectedValue <= "12" And DBEndH.SelectedValue <= "12") Then
                    Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue)
                Else
                    If (DBStartH.SelectedValue >= "13" And DBEndH.SelectedValue >= "13") Then
                        Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue)
                    Else
                        Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue) + 1
                    End If
                End If
                If Hours > 8 Then Hours = 8
            End If
            '計算時間
            wTotal = VHours + Hours - MHours
            DBDay.Text = CStr(Fix(wTotal / 8))
            DBHour.Text = CStr(wTotal - (Fix(wTotal / 8) * 8))
            '實際同步
            DAStartDate.Text = DBStartDate.Text
            DAStartH.SelectedIndex = DAStartH.Items.IndexOf(DAStartH.Items.FindByValue(DBStartH.SelectedValue))
            DAEndDate.Text = DBEndDate.Text
            DAEndH.SelectedIndex = DAEndH.Items.IndexOf(DAEndH.Items.FindByValue(DBEndH.SelectedValue))
            DADay.Text = DBDay.Text
            DAHour.Text = DBHour.Text
        End If

        If ErrCode > 0 Then
            If ErrCode = 8010 Then Message = "結束日期不可小於開始日期,請確認!"
            If ErrCode = 8020 Then Message = "開始日期與結束日期不可跨月,請確認!"
            Response.Write(YKK.ShowMessage(Message))
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     →Button計算實際日數 
    '**
    '*****************************************************************
    Private Sub BADays_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BADays.Click
        Dim Hours As Integer = 0
        Dim VHours As Integer = 0
        Dim MHours As Integer = 0
        Dim wTotal As Integer = 0
        Dim ErrCode As Integer = 0
        Dim SQL, Message As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        '起迄日期檢查
        If ErrCode = 0 Then
            If CDate(DAEndDate.Text) < CDate(DAStartDate.Text) Then ErrCode = 8010
            If CDate(DAEndDate.Text) = CDate(DAStartDate.Text) Then
                If DAEndH.SelectedValue <= DAStartH.SelectedValue Then ErrCode = 8010
            End If
        End If
        '起迄日期同月檢查
        If ErrCode = 0 Then
            If CStr(Year(CDate(DAStartDate.Text))) + "/" + CStr(Month(CDate(DAStartDate.Text))) <> CStr(Year(CDate(DAEndDate.Text))) + "/" + CStr(Month(CDate(DAEndDate.Text))) Then
                ErrCode = 8020
            End If
        End If
        '假日時間
        If ErrCode = 0 Then
            OleDbConnection1.Open()
            SQL = "Select Count(*) As Days From M_Vacation "
            SQL = SQL + "Where Active = '1' "
            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
            'SQL = SQL + "  and Depo = '" + wDepo + "' "
            '
            SQL = SQL + "  and Depo = '" + wApplyCalendar + "' "
            'Modify-End
            SQL = SQL + "  and YMD  >= '" + DAStartDate.Text + "' "
            SQL = SQL + "  and YMD  <= '" + DAEndDate.Text + "' "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "Vacation")
            MHours = DBDataSet1.Tables("Vacation").Rows(0).Item("Days") * 8
            OleDbConnection1.Close()
            '外出時間
            If CDate(DAEndDate.Text) > CDate(DAStartDate.Text) Then
                VHours = DateDiff("d", CDate(DAStartDate.Text), CDate(DAEndDate.Text)) * 8
            End If
            '外出時數差時間
            If DAEndH.SelectedValue >= DAStartH.SelectedValue Then
                If (DAStartH.SelectedValue <= "12" And DAEndH.SelectedValue <= "12") Then
                    Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue)
                Else
                    If (DAStartH.SelectedValue >= "13" And DAEndH.SelectedValue >= "13") Then
                        Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue)
                    Else
                        Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue) - 1
                    End If
                End If
                If Hours > 8 Then Hours = 8
            Else
                If (DAStartH.SelectedValue <= "12" And DAEndH.SelectedValue <= "12") Then
                    Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue)
                Else
                    If (DAStartH.SelectedValue >= "13" And DAEndH.SelectedValue >= "13") Then
                        Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue)
                    Else
                        Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue) + 1
                    End If
                End If
                If Hours > 8 Then Hours = 8
            End If
            '計算時間
            wTotal = VHours + Hours - MHours
            DADay.Text = CStr(Fix(wTotal / 8))
            DAHour.Text = CStr(wTotal - (Fix(wTotal / 8) * 8))
        End If

        If ErrCode > 0 Then
            If ErrCode = 8010 Then Message = "結束日期不可小於開始日期,請確認!"
            If ErrCode = 8020 Then Message = "開始日期與結束日期不可跨月,請確認!"
            Response.Write(YKK.ShowMessage(Message))
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     姓名變更 
    '**
    '*****************************************************************
    Private Sub DName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DName.SelectedIndexChanged
        Dim DBDataSet1 As New DataSet
        Dim SQL As String
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        '取得申請者資訊
        SQL = "Select UserName, EmpID, JobName, JobID, DivName, DivID, DepoID, DepoName From M_Users "
        SQL = SQL & " Where UserID = '" & DName.SelectedValue & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
            DDepoCode.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID")
            DDepoName.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("DepoName")
            DEmpID.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("EmpID")
            DJobTitle.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("JobName")
            DJobCode.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("JobID")
            DDivision.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("DivName")
            DDivisionCode.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("DivID")
        End If
        OleDbConnection1.Close()
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
            If UPFile.PostedFile.ContentLength <= 2000 * 1024 Then
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
        BOK.Disabled = True
        BNG1.Disabled = True
        BNG2.Disabled = True
        BSAVE.Disabled = True
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     起動Button(OK, Ng1, NG2, Save)運作 
    '**
    '*****************************************************************
    Private Sub EnabledButton()
        BOK.Disabled = False
        BNG1.Disabled = False
        BNG2.Disabled = False
        BSAVE.Disabled = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     輸入檢查
    '**
    '*****************************************************************
    Function InputCheck() As Integer
        InputCheck = 0
        'No
        If InputCheck = 0 Then
            If FindFieldInf("No") = 1 Then
                If DNo.Text = "" Then InputCheck = 1
            End If
        End If
        '日期
        If InputCheck = 0 Then
            If FindFieldInf("Date") = 1 Then
                If DDate.Text = "" Then InputCheck = 1
            End If
        End If
        '姓名
        If InputCheck = 0 Then
            If FindFieldInf("Name") = 1 Then
                If DName.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'EmpID
        If InputCheck = 0 Then
            If FindFieldInf("EmpID") = 1 Then
                If DEmpID.Text = "" Then InputCheck = 1
            End If
        End If
        '職稱
        If InputCheck = 0 Then
            If FindFieldInf("JobTitle") = 1 Then
                If DJobTitle.Text = "" Then InputCheck = 1
            End If
        End If
        '職稱代碼
        If InputCheck = 0 Then
            If FindFieldInf("JobCode") = 1 Then
                If DJobCode.Text = "" Then InputCheck = 1
            End If
        End If
        'Depo Name
        If InputCheck = 0 Then
            If FindFieldInf("DepoName") = 1 Then
                If DDepoName.Text = "" Then InputCheck = 1
            End If
        End If
        'Depo Code
        If InputCheck = 0 Then
            If FindFieldInf("DepoCode") = 1 Then
                If DDepoCode.Text = "" Then InputCheck = 1
            End If
        End If
        '部門
        If InputCheck = 0 Then
            If FindFieldInf("Division") = 1 Then
                If DDivision.Text = "" Then InputCheck = 1
            End If
        End If
        '部門代碼
        If InputCheck = 0 Then
            If FindFieldInf("DivisionCode") = 1 Then
                If DDivisionCode.Text = "" Then InputCheck = 1
            End If
        End If
        '所屬年月
        If InputCheck = 0 Then
            If FindFieldInf("SalaryYM") = 1 Then
                If DSalaryYM.Text = "" Then InputCheck = 1
            End If
        End If
        '目的地
        If InputCheck = 0 Then
            If FindFieldInf("Place") = 1 Then
                If DPlace.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '目的地(其他)
        If InputCheck = 0 Then
            If FindFieldInf("OtherPlace") = 1 Then
                If DOtherPlace.Text = "" Then InputCheck = 1
            End If
        End If
        '職務代理人
        If InputCheck = 0 Then
            If FindFieldInf("JobAgent") = 1 Then
                If DJobAgent.Text = "" Then InputCheck = 1
            End If
        End If
        '預定開始日期
        If InputCheck = 0 Then
            If FindFieldInf("BStartDate") = 1 Then
                If DBStartDate.Text = "" Then InputCheck = 1
            End If
        End If
        '預定開始-時
        If InputCheck = 0 Then
            If FindFieldInf("BStartH") = 1 Then
                If DBStartH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '預定結束日期
        If InputCheck = 0 Then
            If FindFieldInf("BEndDate") = 1 Then
                If DBEndDate.Text = "" Then InputCheck = 1
            End If
        End If
        '預定結束-時
        If InputCheck = 0 Then
            If FindFieldInf("BEndH") = 1 Then
                If DBEndH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '預定天數
        If InputCheck = 0 Then
            If FindFieldInf("BDay") = 1 Then
                If DBDay.Text = "" Then InputCheck = 1
            End If
        End If
        '預定時數
        If InputCheck = 0 Then
            If FindFieldInf("BHour") = 1 Then
                If DBHour.Text = "" Then InputCheck = 1
            End If
        End If
        '實際開始日期
        If InputCheck = 0 Then
            If FindFieldInf("AStartDate") = 1 Then
                If DAStartDate.Text = "" Then InputCheck = 1
            End If
        End If
        '實際開始-時
        If InputCheck = 0 Then
            If FindFieldInf("AStartH") = 1 Then
                If DAStartH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '實際結束日期
        If InputCheck = 0 Then
            If FindFieldInf("AEndDate") = 1 Then
                If DAEndDate.Text = "" Then InputCheck = 1
            End If
        End If
        '實際結束-時
        If InputCheck = 0 Then
            If FindFieldInf("AEndH") = 1 Then
                If DAEndH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '實際天數
        If InputCheck = 0 Then
            If FindFieldInf("ADay") = 1 Then
                If DADay.Text = "" Then InputCheck = 1
            End If
        End If
        '實際時數
        If InputCheck = 0 Then
            If FindFieldInf("AHour") = 1 Then
                If DAHour.Text = "" Then InputCheck = 1
            End If
        End If
        '外出理由
        If InputCheck = 0 Then
            If FindFieldInf("FReason") = 1 Then
                If DFReason.Text = "" Then InputCheck = 1
            End If
        End If

    End Function

End Class
