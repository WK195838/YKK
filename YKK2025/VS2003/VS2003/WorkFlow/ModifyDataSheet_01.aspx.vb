Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ModifyDataSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DModifyDataSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DOPReadyDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOPReady As System.Web.UI.WebControls.TextBox
    Protected WithEvents BFlow As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents LBefOP As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMapSheet3 As System.Web.UI.WebControls.Image
    Protected WithEvents DMapSheet4 As System.Web.UI.WebControls.Image
    Protected WithEvents DDelivery As System.Web.UI.WebControls.Image
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents BPrint As System.Web.UI.WebControls.ImageButton
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents DSheet As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMReasonType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DStatus As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMContent1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMContent2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DIContent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFDateTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DIPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DComNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents LSheet1 As System.Web.UI.WebControls.HyperLink

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
    'Dim wDepo As String = "CL"      '中壢行事曆(CL->中壢, TP->台北)
    '
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    'Modify-End

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
            If wFormSno > 0 And wStep > 2 Then      '判斷是否[簽核]
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

        Response.Cookies("PGM").Value = "ModifyDataSheet_01.aspx"
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
        Dim Message As String = ""
        If Message <> "" Then
            Message = "下列所設定的附加檔案將會遺失 (" & Message & ") " & ",請重新設定!"
            Response.Write(YKK.ShowMessage(Message))
        End If
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
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_ModifyDataSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ModifyDataSheet")
        If DBDataSet1.Tables("F_ModifyDataSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("No")                         'No
            DDate.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("Date")                     '日期
            SetFieldData("Division", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("Division"))    '委託部門
            SetFieldData("Person", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("Person"))        '委託擔當
            SetFieldData("MDivision", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MDivision"))  '修改部門
            SetFieldData("MPerson", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MPerson"))      '修改擔當
            SetFieldData("Sheet", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("Sheet"))          '委託單
            DComNo.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("ComNo")                   '委託No
            SetFieldData("Status", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("Status"))        '委託單
            SetFieldData("WDivision", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("WDivision"))  '待處理工程部門
            SetFieldData("WPerson", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("WPerson"))      '待處理工程擔當

            SetFieldData("MReasonType", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MReasonType"))   '修改理由類別
            DMReason.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MReason")               '修改理由
            DMContent1.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MContent1")           '修改內容-1
            DMContent2.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MContent2")           '修改內容-2

            DIContent.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("IContent")             '實際修改內容
            DFDateTime.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("FDateTime")           '預定完成時間
            SetFieldData("IPerson", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("IPerson"))      '實際修改擔當

            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                DDecideDesc.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("DecideDesc")       '說明
                DBStartTime.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BStartTime")       '預定開始
                DBEndTime.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime")           '預定完成
                DAStartTime.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("AStartTime")       '實際開始
                If Not IsDBNull(DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("AEndTime")) Then
                    DAEndTime.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("AEndTime")           '實際完成
                Else
                    DAEndTime.Text = ""
                End If
                LBefOP.NavigateUrl = "BefOPList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno & "&pStep=" & wStep

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
                DFormSno.Text = "單號：" & DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("FormSno")       '單號
            End If
        End If
        '表單連結
        SQL = " select  formno,tablename1  from M_form"
        SQL &= " where formname = '" + Trim(DSheet.SelectedValue) + "'"
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        Dim ds As New Data.DataSet
        DBAdapter3.Fill(ds)
        Dim SheetFormNo As String
        Dim SheetFormSno As String

        Dim dt As DataTable = ds.Tables(0)
        If dt.Rows.Count > 0 Then
            SheetFormNo = dt.Rows(0)("formno")
            SQL = " select formsno from f_" + dt.Rows(0)("tablename1")
            SQL &= " where no='" + Trim(DComNo.Text) + "'"
            Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
            Dim ds1 As New Data.DataSet
            DBAdapter4.Fill(ds1)
            Dim dt1 As DataTable = ds1.Tables(0)
            Dim sUrl As String
            Dim cun As Integer

            If dt1.Rows.Count > 0 Then
                SheetFormSno = dt1.Rows(0)("formsno")
                sUrl = dt.Rows(0)("tablename1") + "_02.aspx?pFormNo=" & SheetFormNo & "&pFormSno=" & SheetFormSno
                'Dim sUrlStr = Mid(sUrl, 11, sUrl.Length - 1)
                LSheet1.NavigateUrl = sUrl

            End If
        End If
        'DB連結關閉
        OleDbConnection1.Close()
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
            Else
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

        If wFormSno > 0 And wStep > 2 Then      '判斷是否[簽核]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                'Sheet顯示
                DMapSheet3.Visible = True   '圖面Sheet
                DMapSheet4.Visible = True   '說明Sheet
                DDelivery.Visible = True    '交期Sheet

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
                    Top = 992
                Else
                    DReasonCode.Visible = False     '延遲理由代碼
                    DReason.Visible = False         '延遲理由
                    DReasonDesc.Visible = False     '延遲其他說明
                    Top = 880
                End If

                '交期
                DBStartTime.Visible = True      '預定開始
                DBEndTime.Visible = True        '預定完成
                DAStartTime.Visible = True      '實際開始
                DAEndTime.Visible = True        '實際完成
                '工程履歷閱讀
                DOPReady.Visible = True
                DOPReadyDesc.Visible = True
                '需閱讀
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("FlowType") = 0 Then
                    DOPReady.BackColor = Color.LightGray
                Else
                    '20190430 JESSICA 本田
                    DOPReady.BackColor = Color.LightGray

                    'DOPReady.BackColor = Color.GreenYellow
                    'ShowRequiredFieldValidator("DOPReadyRqd", "DOPReady", "異常：需閱讀工程履歷")
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
                LBefOP.Visible = True          '工程履歷
                LSheet1.Visible = True          '工程履歷
                '按鈕位置
                BNG1.Style.Add("Top", Top)      'NG按鈕
                BNG2.Style.Add("Top", Top)     'NG按鈕
                BSAVE.Style.Add("Top", Top)    '儲存按鈕
                BOK.Style.Add("Top", Top)      'OK按鈕
                DFormSno.Style.Add("Top", Top) '單號
            End If
        Else
            Top = 696
            'Sheet隱藏
            DMapSheet3.Visible = False  '圖面Sheet
            DMapSheet4.Visible = False  '說明Sheet
            DDelivery.Visible = False   '交期Sheet

            DOPReady.Visible = False    '工程履歷閱讀
            DOPReadyDesc.Visible = False

            DDelay.Visible = False      '延遲Sheet
            '欄位隱藏
            DDecideDesc.Visible = False     '說明
            DBStartTime.Visible = False     '預定開始
            DBEndTime.Visible = False       '預定完成
            DAStartTime.Visible = False     '實際開始
            DAEndTime.Visible = False       '實際完成
            DReasonCode.Visible = False     '延遲理由代碼
            DReason.Visible = False         '延遲理由
            DReasonDesc.Visible = False     '延遲其他說明
            '連結隱藏
            LBefOP.Visible = False          '工程履歷
            LSheet1.Visible = False          '工程履歷
            '按鈕位置
            BNG1.Style.Add("Top", Top)      'NG按鈕
            BNG2.Style.Add("Top", Top)     'NG按鈕
            BSAVE.Style.Add("Top", Top)    '儲存按鈕
            BOK.Style.Add("Top", Top)      'OK按鈕
            DFormSno.Style.Add("Top", Top) '單號
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
        '日期選擇
        BDate.Attributes("onclick") = "calendarPicker('Form1.DDate');"
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
    '**
    '**     修改者部門點選後事件
    '**
    '*****************************************************************
    Private Sub DMDivision_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DMDivision.SelectedIndexChanged
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim i As Integer
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        DMPerson.Items.Clear()
        SQL = "Select UserName From M_Users "
        SQL = SQL & " Where DivName = '" & DMDivision.SelectedValue & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")

        For i = 0 To DBTable1.Rows.Count - 1
            If i = 0 Then
                Dim ListItem1 As New ListItem
                ListItem1.Text = ""
                ListItem1.Value = ""
                ListItem1.Selected = True
                DMPerson.Items.Add(ListItem1)
            End If
            Dim ListItem2 As New ListItem
            ListItem2.Text = DBTable1.Rows(i).Item("UserName")
            ListItem2.Value = DBTable1.Rows(i).Item("UserName")
            ListItem2.Selected = False
            DMPerson.Items.Add(ListItem2)
        Next
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     待處理工程部門點選後事件
    '**
    '*****************************************************************
    Private Sub DWDivision_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DWDivision.SelectedIndexChanged
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim i As Integer
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        DWPerson.Items.Clear()
        SQL = "Select UserName From M_Users "
        SQL = SQL & " Where DivName = '" & DWDivision.SelectedValue & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")

        For i = 0 To DBTable1.Rows.Count - 1
            If i = 0 Then
                Dim ListItem1 As New ListItem
                ListItem1.Text = ""
                ListItem1.Value = ""
                ListItem1.Selected = True
                DWPerson.Items.Add(ListItem1)
            End If
            Dim ListItem2 As New ListItem
            ListItem2.Text = DBTable1.Rows(i).Item("UserName")
            ListItem2.Value = DBTable1.Rows(i).Item("UserName")
            ListItem2.Selected = False
            DWPerson.Items.Add(ListItem2)
        Next
        'DB連結關閉
        OleDbConnection1.Close()
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
        If wFormSno > 0 And wStep > 2 Then      '判斷是否[簽核]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") >= NowDateTime Then
                    Top = 880
                Else
                    If DDelay.Visible = True Then
                        Top = 992
                    Else
                        Top = 880
                    End If
                End If
            End If
        Else
            Top = 696
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
    '**(ShowSheetField)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '表單各欄位屬性及欄位輸入檢查等設定
        'No
        Select Case FindFieldInf("No")
            Case 0  '顯示
                DNo.BackColor = Color.LightGray
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

        '日期
        Select Case FindFieldInf("Date")
            Case 0  '顯示
                DDate.BackColor = Color.LightGray
                DDate.ReadOnly = True
                DDate.Visible = True
                BDate.Visible = False
            Case 1  '修改+檢查
                DDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDateRqd", "DDate", "異常：需輸入日期")
                DDate.Visible = True
                BDate.Visible = True
            Case 2  '修改
                DDate.BackColor = Color.Yellow
                DDate.Visible = True
                BDate.Visible = True
            Case Else   '隱藏
                DDate.Visible = False
                BDate.Visible = False
        End Select
        If pPost = "New" Then DDate.Text = CStr(DateTime.Now.Today)

        '委託部門
        Select Case FindFieldInf("Division")
            Case 0  '顯示
                DDivision.BackColor = Color.LightGray
                'DDivision.Enabled = False
                DDivision.Visible = True
            Case 1  '修改+檢查
                DDivision.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDivisionRqd", "DDivision", "異常：需輸入委託部門")
                DDivision.Visible = True
            Case 2  '修改
                DDivision.BackColor = Color.Yellow
                DDivision.Visible = True
            Case Else   '隱藏
                DDivision.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Division", "ZZZZZZ")

        '委託擔當
        Select Case FindFieldInf("Person")
            Case 0  '顯示
                DPerson.BackColor = Color.LightGray
                'DPerson.Enabled = False
                DPerson.Visible = True
            Case 1  '修改+檢查
                DPerson.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPersonRqd", "DPerson", "異常：需輸入委託擔當")
                DPerson.Visible = True
            Case 2  '修改
                DPerson.BackColor = Color.Yellow
                DPerson.Visible = True
            Case Else   '隱藏
                DPerson.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Person", "ZZZZZZ")

        '修改部門
        Select Case FindFieldInf("MDivision")
            Case 0  '顯示
                DMDivision.BackColor = Color.LightGray
                'DDivision.Enabled = False
                DMDivision.Visible = True
            Case 1  '修改+檢查
                DMDivision.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMDivisionRqd", "DMDivision", "異常：需輸入修改部門")
                DMDivision.Visible = True
            Case 2  '修改
                DMDivision.BackColor = Color.Yellow
                DMDivision.Visible = True
            Case Else   '隱藏
                DMDivision.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("MDivision", "ZZZZZZ")

        '修改擔當
        Select Case FindFieldInf("MPerson")
            Case 0  '顯示
                DMPerson.BackColor = Color.LightGray
                'DPerson.Enabled = False
                DMPerson.Visible = True
            Case 1  '修改+檢查
                DMPerson.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMPersonRqd", "DMPerson", "異常：需輸入修改擔當")
                DMPerson.Visible = True
            Case 2  '修改
                DMPerson.BackColor = Color.Yellow
                DMPerson.Visible = True
            Case Else   '隱藏
                DMPerson.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("MPerson", "ZZZZZZ")

        '委託書
        Select Case FindFieldInf("Sheet")
            Case 0  '顯示
                DSheet.BackColor = Color.LightGray
                DSheet.Visible = True
            Case 1  '修改+檢查
                DSheet.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSheetRqd", "DSheet", "異常：需輸入委託書")
                DSheet.Visible = True
            Case 2  '修改
                DSheet.BackColor = Color.Yellow
                DSheet.Visible = True
            Case Else   '隱藏
                DSheet.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Sheet", "ZZZZZZ")

        '委託No
        Select Case FindFieldInf("ComNo")
            Case 0  '顯示
                DComNo.BackColor = Color.LightGray
                DComNo.ReadOnly = True
                DComNo.Visible = True
            Case 1  '修改+檢查
                DComNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DComNoRqd", "DComNo", "異常：需輸入Ｎｏ")
                DComNo.Visible = True
            Case 2  '修改
                DComNo.BackColor = Color.Yellow
                DComNo.Visible = True
            Case Else   '隱藏
                DComNo.Visible = False
        End Select
        If pPost = "New" Then DComNo.Text = ""

        '狀態
        Select Case FindFieldInf("Status")
            Case 0  '顯示
                DStatus.BackColor = Color.LightGray
                DStatus.Visible = True
            Case 1  '修改+檢查
                DStatus.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DStatusRqd", "DStatus", "異常：需輸入狀態")
                DStatus.Visible = True
            Case 2  '修改
                DStatus.BackColor = Color.Yellow
                DStatus.Visible = True
            Case Else   '隱藏
                DStatus.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Status", "ZZZZZZ")

        '待處理工程部門
        Select Case FindFieldInf("WDivision")
            Case 0  '顯示
                DWDivision.BackColor = Color.LightGray
                DWDivision.Visible = True
            Case 1  '修改+檢查
                DWDivision.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWDivisionRqd", "DWDivision", "異常：需輸入待處理工程部門")
                DWDivision.Visible = True
            Case 2  '修改
                DWDivision.BackColor = Color.Yellow
                DWDivision.Visible = True
            Case Else   '隱藏
                DWDivision.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WDivision", "ZZZZZZ")

        '待處理工程擔當
        Select Case FindFieldInf("WPerson")
            Case 0  '顯示
                DWPerson.BackColor = Color.LightGray
                DWPerson.Visible = True
            Case 1  '修改+檢查
                DWPerson.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWPersonRqd", "DWPerson", "異常：需輸入待處理工程擔當")
                DWPerson.Visible = True
            Case 2  '修改
                DWPerson.BackColor = Color.Yellow
                DWPerson.Visible = True
            Case Else   '隱藏
                DWPerson.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WPerson", "ZZZZZZ")

        '修改理由類別
        Select Case FindFieldInf("MReasonType")
            Case 0  '顯示
                DMReasonType.BackColor = Color.LightGray
                DMReasonType.Visible = True
            Case 1  '修改+檢查
                DMReasonType.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMReasonTypeRqd", "DMReasonType", "異常：需輸入修改理由類別")
                DMReasonType.Visible = True
            Case 2  '修改
                DMReasonType.BackColor = Color.Yellow
                DMReasonType.Visible = True
            Case Else   '隱藏
                DMReasonType.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("MReasonType", "ZZZZZZ")

        '修改理由
        Select Case FindFieldInf("MReason")
            Case 0  '顯示
                DMReason.BackColor = Color.LightGray
                DMReason.ReadOnly = True
                DMReason.Visible = True
            Case 1  '修改+檢查
                DMReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMReasonRqd", "DMReason", "異常：需輸入修改理由")
                DMReason.Visible = True
            Case 2  '修改
                DMReason.BackColor = Color.Yellow
                DMReason.Visible = True
            Case Else   '隱藏
                DMReason.Visible = False
        End Select
        If pPost = "New" Then DMReason.Text = ""

        '修改內容-1
        Select Case FindFieldInf("MContent1")
            Case 0  '顯示
                DMContent1.BackColor = Color.LightGray
                DMContent1.ReadOnly = True
                DMContent1.Visible = True
            Case 1  '修改+檢查
                DMContent1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMContent1Rqd", "DMContent1", "異常：需輸入修改內容-1")
                DMContent1.Visible = True
            Case 2  '修改
                DMContent1.BackColor = Color.Yellow
                DMContent1.Visible = True
            Case Else   '隱藏
                DMContent1.Visible = False
        End Select
        If pPost = "New" Then DMContent1.Text = ""

        '修改內容-2
        Select Case FindFieldInf("MContent2")
            Case 0  '顯示
                DMContent2.BackColor = Color.LightGray
                DMContent2.ReadOnly = True
                DMContent2.Visible = True
            Case 1  '修改+檢查
                DMContent2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMContent2Rqd", "DMContent2", "異常：需輸入修改內容-2")
                DMContent2.Visible = True
            Case 2  '修改
                DMContent2.BackColor = Color.Yellow
                DMContent2.Visible = True
            Case Else   '隱藏
                DMContent2.Visible = False
        End Select
        If pPost = "New" Then DMContent2.Text = ""

        '實際修改內容
        Select Case FindFieldInf("IContent")
            Case 0  '顯示
                DIContent.BackColor = Color.LightGray
                DIContent.ReadOnly = True
                DIContent.Visible = True
            Case 1  '修改+檢查
                DIContent.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DIContentRqd", "DIContent", "異常：需輸入實際修改內容")
                DIContent.Visible = True
            Case 2  '修改
                DIContent.BackColor = Color.Yellow
                DIContent.Visible = True
            Case Else   '隱藏
                DIContent.Visible = False
        End Select
        If pPost = "New" Then DIContent.Text = ""

        '預定完成時間
        Select Case FindFieldInf("FDateTime")
            Case 0  '顯示
                DFDateTime.BackColor = Color.LightGray
                DFDateTime.ReadOnly = True
                DFDateTime.Visible = True
            Case 1  '修改+檢查
                DFDateTime.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFDateTimeRqd", "DFDateTime", "異常：需輸入預定完成時間")
                DFDateTime.Visible = True
            Case 2  '修改
                DFDateTime.BackColor = Color.Yellow
                DFDateTime.Visible = True
            Case Else   '隱藏
                DFDateTime.Visible = False
        End Select
        If pPost = "New" Then DFDateTime.Text = ""

        '資料修改擔當
        Select Case FindFieldInf("IPerson")
            Case 0  '顯示
                DIPerson.BackColor = Color.LightGray
                DIPerson.Visible = True
            Case 1  '修改+檢查
                DIPerson.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DIPersonRqd", "DIPerson", "異常：需輸入資料修改擔當")
                DIPerson.Visible = True
            Case 2  '修改
                DIPerson.BackColor = Color.Yellow
                DIPerson.Visible = True
            Case Else   '隱藏
                DIPerson.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("IPerson", "ZZZZZZ")

        '單號
        If pPost = "New" Then DFormSno.Text = ""
    End Sub
    '*****************************************************************
    '**(ShowSheetField)
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator
        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.ErrorMessage = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", CStr(Top + 25))
        rqdVal.Style.Add("Left", "8")
        rqdVal.Style.Add("Height", "20")
        rqdVal.Style.Add("Width", "250")
        rqdVal.Style.Add("POSITION", "absolute")
        Page.Controls(1).Controls.Add(rqdVal)
    End Sub

    '*****************************************************************
    '**(ShowSheetField)
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

        '委託部門
        If pFieldName = "Division" Then
            DDivision.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDivision.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select DivName From M_Users "
                SQL = SQL & " Where UserID = '" & Request.QueryString("pUserID") & "'"
                SQL = SQL & "   And Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("DivName")
                    ListItem1.Value = DBTable1.Rows(i).Item("DivName")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDivision.Items.Add(ListItem1)
                Next
            End If
        End If

        '委託擔當
        If pFieldName = "Person" Then
            DPerson.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPerson.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select UserName From M_Users "
                SQL = SQL & " Where UserID = '" & Request.QueryString("pUserID") & "'"
                SQL = SQL & "   And Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("UserName")
                    ListItem1.Value = DBTable1.Rows(i).Item("UserName")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPerson.Items.Add(ListItem1)
                Next
            End If
        End If

        '修改部門
        If pFieldName = "MDivision" Then
            DMDivision.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMDivision.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select DivName From M_Users "
                SQL = SQL & " Where Active = '1' "
                SQL = SQL & " Group By DivName "
                SQL = SQL & " Order By DivName "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("DivName")
                    ListItem1.Value = DBTable1.Rows(i).Item("DivName")
                    If ListItem1.Value = DDivision.SelectedValue Then ListItem1.Selected = True
                    DMDivision.Items.Add(ListItem1)
                Next
            End If
        End If

        '修改擔當
        If pFieldName = "MPerson" Then
            DMPerson.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMPerson.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select UserName From M_Users "
                SQL = SQL & " Where DivName = '" & DMDivision.SelectedValue & "'"
                SQL = SQL & "   And Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("UserName")
                    ListItem1.Value = DBTable1.Rows(i).Item("UserName")
                    If ListItem1.Value = DPerson.SelectedValue Then ListItem1.Selected = True
                    DMPerson.Items.Add(ListItem1)
                Next
            End If
        End If

        '待處理工程部門
        If pFieldName = "WDivision" Then
            DWDivision.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWDivision.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select DivName From M_Users "
                SQL = SQL & " Where Active = '1' "
                SQL = SQL & " Group By DivName "
                SQL = SQL & " Order By DivName "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    If pName <> "ZZZZZZ" Then
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("DivName")
                        ListItem2.Value = DBTable1.Rows(i).Item("DivName")
                        If ListItem2.Value = pName Then ListItem2.Selected = True
                        DWDivision.Items.Add(ListItem2)
                    Else
                        If i = 0 Then
                            Dim ListItem1 As New ListItem
                            ListItem1.Text = ""
                            ListItem1.Value = ""
                            ListItem1.Selected = True
                            DWDivision.Items.Add(ListItem1)
                        End If
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("DivName")
                        ListItem2.Value = DBTable1.Rows(i).Item("DivName")
                        ListItem2.Selected = False
                        DWDivision.Items.Add(ListItem2)
                    End If
                Next
            End If
        End If

        '待處理工程擔當
        If pFieldName = "WPerson" Then
            DWPerson.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWPerson.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select UserName From M_Users "
                SQL = SQL & " Where Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")

                For i = 0 To DBTable1.Rows.Count - 1
                    If pName <> "ZZZZZZ" Then
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("UserName")
                        ListItem2.Value = DBTable1.Rows(i).Item("UserName")
                        If ListItem2.Value = pName Then ListItem2.Selected = True
                        DWPerson.Items.Add(ListItem2)
                    Else
                        If i = 0 Then
                            Dim ListItem1 As New ListItem
                            ListItem1.Text = ""
                            ListItem1.Value = ""
                            ListItem1.Selected = True
                            DWPerson.Items.Add(ListItem1)
                        End If
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("UserName")
                        ListItem2.Value = DBTable1.Rows(i).Item("UserName")
                        ListItem2.Selected = False
                        DWPerson.Items.Add(ListItem2)
                    End If
                Next
            End If
        End If

        '委託單
        If pFieldName = "Sheet" Then
            DSheet.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSheet.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select FormName From M_Form "
                SQL = SQL & " Where FormNo < 1000 "
                SQL = SQL & "   And Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Form")
                DBTable1 = DBDataSet1.Tables("M_Form")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("FormName")
                    ListItem1.Value = DBTable1.Rows(i).Item("FormName")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSheet.Items.Add(ListItem1)
                Next
            End If
        End If

        '狀態
        If pFieldName = "Status" Then
            DStatus.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DStatus.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='900' and DKey='STATUS' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DStatus.Items.Add(ListItem1)
                Next
            End If
        End If

        '修改理由類別
        If pFieldName = "MReasonType" Then
            DMReasonType.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMReasonType.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='900' and DKey='MODIFYREASON' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DMReasonType.Items.Add(ListItem1)
                Next
            End If
        End If

        '實際修改擔當
        If pFieldName = "IPerson" Then
            DIPerson.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DIPerson.Items.Add(ListItem1)

                End If
            Else
                SQL = "Select * From M_Referp Where Cat='900' and DKey='MODIFYPERSON' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DIPerson.Items.Add(ListItem1)
                Next

                SQL = " Select  data  From M_Referp a,M_users b  Where Cat='900' and DKey='MODIFYPERSON'  "
                SQL &= " and userid ='" & Request.QueryString("pUserID") & "'  and  substring(a.data,4,6)=b.username "
                SQL &= " Order by DKey, Data"

                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                DIPerson.SelectedValue = DBTable1.Rows(i).Item("Data")
                DFDateTime.Text = Format(Now, "MM/dd hh:mm")


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
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DReason.Text = pName
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='014' and DKey= '" & DReasonCode.SelectedValue & "' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    DReason.Text = DBTable1.Rows(i).Item("Data")
                Next
            End If
        End If
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
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
                    If ErrCode = 9010 Then Message = "上傳檔案格式錯誤,請確認!"
                    If ErrCode = 9020 Then Message = "上傳檔案Size超過1024KB,請確認!"
                    If ErrCode = 9030 Then Message = "上傳檔案沒指定,請確認!"
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
    Private Sub BOK_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
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
    '**     NG1按鈕點選後事件
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
    '**     NG2按鈕點選後事件
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
        Dim Message As String = ""
        Dim wQCLT As Integer = 0 'QC-L/T

        GetDataStatus()  '取得此關卡各按鈕的Data Status

        'Check待處理工程
        If ErrCode = 0 Then
            'If DStatus.SelectedValue = "01-開發中" Then
            If DWDivision.SelectedValue = "" Or DWPerson.SelectedValue = "" Then
                ErrCode = 9070
            End If
            'End If
        End If

        'Check延遲理由
        If ErrCode = 0 Then
            If DReasonCode.Visible = True Then
                If DReasonCode.SelectedValue = "99" Then
                    If DReasonDesc.Text = "" Then ErrCode = 9040
                End If
            End If
        End If

        '--檢查委託書No---------
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                ErrCode = oCommon.CommissionNo("800001", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9060
                End If
            End If
        End If

        If ErrCode = 0 Then
            Dim Run As Boolean = True           '是否執行
            Dim RepeatRun As Boolean = False    '是否重覆執行
            Dim MultiJob As Integer = 0         '多人核定
            Dim wLevel As String = ""           '樣品難易度

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
                Dim i, RtnCode As Integer
                Dim NewFormSno As Integer = wFormSno    '表單流水號
                Dim pRunNextStep As Integer = 0         '是否執行計算下一關(會簽)
                Dim SQL As String

                '取得表單流水號或更新交易資料
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then      '判斷是否起單
                        '取得表單流水號
                        RtnCode = oCommon.GetFormSeqNo(wFormNo, NewFormSno) '表單號碼, 表單流水號
                        If RtnCode <> 0 Then
                            ErrCode = 9110
                        Else
                            '委託No自動編列
                            DNo.Text = DComNo.Text & "-" & CStr(NewFormSno)

                            '申請流程資料建置
                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                            '
                            oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                            '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者
                            'Modify-End
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
                    Dim DBDataSet1 As New DataSet
                    Dim wAllocateID As String = ""
                    Dim OleDbConnection1 As New OleDbConnection
                    OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
                    OleDbConnection1.Open()

                    'jessica 20150209 修改增加
                    SQL = "Select UserID From M_Users "
                    SQL = SQL & " Where Active = 1 "
                    If wStep = 20 Or wStep = 22 Or wStep = 27 Or wStep = 46 Then
                        SQL = SQL & "   And UserName =  '" & DWPerson.SelectedValue & "'"
                    Else
                        SQL = SQL & "   And UserName =  '" & DMPerson.SelectedValue & "'"
                    End If
                    Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter1.Fill(DBDataSet1, "M_Users")
                    If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then wAllocateID = DBDataSet1.Tables("M_Users").Rows(0).Item("UserID")

                    OleDbConnection1.Close()
                    '
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
                    If wFormSno = 0 And wStep < 3 Then      '判斷是否起單
                        If NewFormSno <> 0 Then
                            AppendData(pFun, NewFormSno)  'Insert
                            AddCommissionNo(wFormNo, NewFormSno)
                        End If  'pSeqno <> 0
                    Else    '判斷是否起單
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
            If ErrCode = 9050 Then Message = "材質為其他時需填寫說明,請確認!"
            If ErrCode = 9060 Then Message = "委託書No.重覆,請確認委託書No.!"
            If ErrCode = 9070 Then Message = "需填入工程待處理部門及工程待處理者.!"
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("MapFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Insert into F_ModifyDataSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, No, "                  '1~5
        SQl = SQl + "Date, Division, Person, MDivision, MPerson, "               '6~10
        SQl = SQl + "Sheet, ComNo, Status, WDivision, WPerson, "                 '11~15
        SQl = SQl + "MReasonType, MReason, MContent1, MContent2, IContent, FDateTime, "   '16~21
        SQl = SQl + "IPerson, "                                                  '22
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "            '23~26
        SQl = SQl + ")  "

        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        SQl = SQl + " '" + NowDateTime + "', "        '結案日
        SQl = SQl + " '800001', "                     '表單代號
        SQl = SQl + " '" + CStr(NewFormSno) + "', "   '表單流水號
        SQl = SQl + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO

        SQl = SQl + " '" + DDate.Text + "', "                '日期
        SQl = SQl + " '" + DDivision.SelectedValue + "', "   '委託部門
        SQl = SQl + " '" + DPerson.SelectedValue + "', "     '委託擔當
        SQl = SQl + " '" + DMDivision.SelectedValue + "', "  '修改部門
        SQl = SQl + " '" + DMPerson.SelectedValue + "', "    '修改擔當

        SQl = SQl + " '" + DSheet.SelectedValue + "', "      '委託單
        SQl = SQl + " N'" + YKK.ReplaceString(DComNo.Text) + "', "   '委託NO
        SQl = SQl + " '" + DStatus.SelectedValue + "', "     '狀態
        SQl = SQl + " '" + DWDivision.SelectedValue + "', "  '待處理工程部門
        SQl = SQl + " '" + DWPerson.SelectedValue + "', "    '待處理工程擔當

        SQl = SQl + " '" + DMReasonType.SelectedValue + "', "   '修改理由類別
        SQl = SQl + " N'" + YKK.ReplaceString(DMReason.Text) + "', "     '修改理由
        SQl = SQl + " N'" + YKK.ReplaceString(DMContent1.Text) + "', "   '修改內容-1
        SQl = SQl + " N'" + YKK.ReplaceString(DMContent2.Text) + "', "   '修改內容-2
        SQl = SQl + " N'" + YKK.ReplaceString(DIContent.Text) + "', "    '實際修改內容
        SQl = SQl + " N'" + YKK.ReplaceString(DFDateTime.Text) + "', "   '實際修改時間

        SQl = SQl + " '" + DIPerson.SelectedValue + "', "    '實際修改擔當

        SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '作成者
        SQl = SQl + " '" + NowDateTime + "', "       '作成時間
        SQl = SQl + " '" + "" + "', "                       '修改者
        SQl = SQl + " '" + NowDateTime + "' "       '修改時間
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("MapFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Update F_ModifyDataSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"

        SQl = SQl + " Date = '" & DDate.Text & "',"
        SQl = SQl + " Division = '" & DDivision.SelectedValue & "',"
        SQl = SQl + " Person = '" & DPerson.SelectedValue & "',"
        SQl = SQl + " MDivision = '" & DMDivision.SelectedValue & "',"
        SQl = SQl + " MPerson = '" & DMPerson.SelectedValue & "',"

        SQl = SQl + " Sheet = '" & DSheet.SelectedValue & "',"
        SQl = SQl + " ComNo = N'" & YKK.ReplaceString(DComNo.Text) & "',"
        SQl = SQl + " Status = '" & DStatus.SelectedValue & "',"
        SQl = SQl + " WDivision = '" & DWDivision.SelectedValue & "',"
        SQl = SQl + " WPerson = '" & DWPerson.SelectedValue & "',"

        SQl = SQl + " MReasonType = '" & DMReasonType.SelectedValue & "',"
        SQl = SQl + " MReason = N'" & YKK.ReplaceString(DMReason.Text) & "',"
        SQl = SQl + " MContent1 = N'" & YKK.ReplaceString(DMContent1.Text) & "',"
        SQl = SQl + " MContent2 = N'" & YKK.ReplaceString(DMContent2.Text) & "',"
        SQl = SQl + " IContent = N'" & YKK.ReplaceString(DIContent.Text) & "',"
        SQl = SQl + " FDateTime = N'" & YKK.ReplaceString(DFDateTime.Text) & "',"

        SQl = SQl + " IPerson = '" & DIPerson.SelectedValue & "',"

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
                SQl = SQl + "FormNo, FormSno, No, CreateUser, CreateTime "    '1~5
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '" + pFormNo + "', "
                SQl = SQl + " '" + CStr(pFormSno) + "', "
                SQl = SQl + " '" + DNo.Text + "', "
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
            If UPFile.PostedFile.ContentLength <= 1000 * 1024 Then
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
    '**     列印表單
    '**
    '*****************************************************************
    Private Sub BPrint_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BPrint.Click
        Dim URL As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()
        SQL = "Select ViewURL From V_WaitHandle_01 "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "URL")
        If DBDataSet1.Tables("URL").Rows.Count > 0 Then
            URL = DBDataSet1.Tables("URL").Rows(0).Item("ViewURL")
        End If
        'DB連結關閉
        OleDbConnection1.Close()

        'Call JavaScript
        Dim scriptString As String = "<script language=JavaScript> "
        scriptString = scriptString & "OpenPrintSheet('" & URL & "'); "
        scriptString = scriptString & "</script>"
        If (Not Me.IsStartupScriptRegistered("Startup")) Then
            Me.RegisterStartupScript("Startup", scriptString)
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     流程模擬表單
    '**
    '*****************************************************************
    Private Sub BFlow_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BFlow.Click
        Dim URL As String = ""
        Dim SQL As String = ""
        Dim wLevel As String = ""           '樣品難易度
        Dim wPerson As String = ""         '製圖者

        If wStep = 1 Or wStep = 46 Then
            If DWPerson.SelectedValue <> "" Then
                Dim DBDataSet1 As New DataSet
                Dim OleDbConnection1 As New OleDbConnection
                OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
                OleDbConnection1.Open()

                SQL = "Select UserID From M_Users "
                SQL = SQL & " Where Active = 1 "
                SQL = SQL & "   And UserName =  '" & DWPerson.SelectedValue & "'"
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then wPerson = DBDataSet1.Tables("M_Users").Rows(0).Item("UserID")

                OleDbConnection1.Close()
            End If
        Else
            If DMPerson.SelectedValue <> "" Then
                Dim DBDataSet1 As New DataSet
                Dim OleDbConnection1 As New OleDbConnection
                OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
                OleDbConnection1.Open()

                SQL = "Select UserID From M_Users "
                SQL = SQL & " Where Active = 1 "
                SQL = SQL & "   And UserName =  '" & DMPerson.SelectedValue & "'"
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then wPerson = DBDataSet1.Tables("M_Users").Rows(0).Item("UserID")

                OleDbConnection1.Close()
            End If
        End If


        URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                              "&pFormSno=" & wFormSno & _
                              "&pStep=" & wStep & _
                              "&pLevel=" & wLevel & _
                              "&pAllocateID=" & wPerson

        'Call JavaScript
        Dim scriptString As String = "<script language=JavaScript> "
        scriptString = scriptString & "OpenSimulationSheet('" & URL & "'); "
        scriptString = scriptString & "</script>"
        If (Not Me.IsStartupScriptRegistered("Startup")) Then
            Me.RegisterStartupScript("Startup", scriptString)
        End If

    End Sub

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
        '部門
        If InputCheck = 0 Then
            If FindFieldInf("Division") = 1 Then
                If DDivision.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '擔當
        If InputCheck = 0 Then
            If FindFieldInf("Person") = 1 Then
                If DPerson.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '修改部門
        If InputCheck = 0 Then
            If FindFieldInf("MDivision") = 1 Then
                If DMDivision.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '修改擔當
        If InputCheck = 0 Then
            If FindFieldInf("MPerson") = 1 Then
                If DMPerson.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '委託書
        If InputCheck = 0 Then
            If FindFieldInf("Sheet") = 1 Then
                If DSheet.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '委託No
        If InputCheck = 0 Then
            If FindFieldInf("ComNo") = 1 Then
                If DComNo.Text = "" Then InputCheck = 1
            End If
        End If
        '狀態
        If InputCheck = 0 Then
            If FindFieldInf("Status") = 1 Then
                If DStatus.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '待處理工程部門
        If InputCheck = 0 Then
            If FindFieldInf("WDivision") = 1 Then
                If DWDivision.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '待處理工程擔當
        If InputCheck = 0 Then
            If FindFieldInf("WPerson") = 1 Then
                If DWPerson.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '修改理由類別
        If InputCheck = 0 Then
            If FindFieldInf("MReasonType") = 1 Then
                If DMReasonType.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '修改理由
        If InputCheck = 0 Then
            If FindFieldInf("MReason") = 1 Then
                If DMReason.Text = "" Then InputCheck = 1
            End If
        End If
        '修改內容-1
        If InputCheck = 0 Then
            If FindFieldInf("MContent1") = 1 Then
                If DMContent1.Text = "" Then InputCheck = 1
            End If
        End If
        '修改內容-2
        If InputCheck = 0 Then
            If FindFieldInf("MContent2") = 1 Then
                If DMContent2.Text = "" Then InputCheck = 1
            End If
        End If
        '實際修改內容
        If InputCheck = 0 Then
            If FindFieldInf("IContent") = 1 Then
                If DIContent.Text = "" Then InputCheck = 1
            End If
        End If
        '預定完成時間
        If InputCheck = 0 Then
            If FindFieldInf("FDateTime") = 1 Then
                If DFDateTime.Text = "" Then InputCheck = 1
            End If
        End If
        '資料修改擔當
        If InputCheck = 0 Then
            If FindFieldInf("IPerson") = 1 Then
                If DIPerson.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'Add-Start 2011/12/2 對應:濫用資料修改委託書
        '                    對策:1人限1張(Active=1)
        If wFormSno = 0 And wStep < 3 Then      ' 起單
            If oCommon.CanUseSheet("800001", Request.QueryString("pUserID")) = 1 Then
                InputCheck = 1
                Response.Write(YKK.ShowMessage("上一張資料修改委託書未完成 ! "))
            End If
        End If

    End Function
End Class
