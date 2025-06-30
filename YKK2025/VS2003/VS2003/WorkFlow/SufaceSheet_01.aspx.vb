Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class SufaceSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DOrderTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCResult3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCResult2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCRemark3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCRemark2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCDate3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCDate2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCRemark1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCResult1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCDate1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEADesc1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEACheck1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck12 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck11 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck10 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck9 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck8 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck5 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck4 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LFinalSampleFile As System.Web.UI.WebControls.Image
    Protected WithEvents DEnglishName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents BBFinalDate As System.Web.UI.WebControls.Button
    Protected WithEvents DBFinalDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAllowSample As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQty As System.Web.UI.WebControls.TextBox
    Protected WithEvents DColor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DManufType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BOrderTime As System.Web.UI.WebControls.Button
    Protected WithEvents BReqDelDate As System.Web.UI.WebControls.Button
    Protected WithEvents DSliderSample As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReqQty As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReqDelDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDevReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOPReadyDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOPReady As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSufaceSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DSuppiler As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBuyer As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LQCReqFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOPManualFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LEACheckFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFinalFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents LContactFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LBefOP As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DAEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents LForcastFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LManufFlowFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAttachSample As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DORNO As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSpec As System.Web.UI.WebControls.Button
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDelivery As System.Web.UI.WebControls.Image
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DDescSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DSufaceSheet2 As System.Web.UI.WebControls.Image
    Protected WithEvents LCustSampleFile As System.Web.UI.WebControls.Image
    Protected WithEvents DFinalSampleFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQCReqFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DCustSampleFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DOPManualFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQCFinalFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DContactFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DForcastFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DManufFlowFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DEACheckFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BFlow As System.Web.UI.WebControls.ImageButton
    Protected WithEvents BPrint As System.Web.UI.WebControls.ImageButton
    Protected WithEvents BQCDate1 As System.Web.UI.WebControls.Button
    Protected WithEvents BQCDate2 As System.Web.UI.WebControls.Button
    Protected WithEvents BQCDate3 As System.Web.UI.WebControls.Button
    Protected WithEvents DQCLT As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents BIn As System.Web.UI.WebControls.Button
    Protected WithEvents BOut As System.Web.UI.WebControls.Button
    Protected WithEvents BImport As System.Web.UI.WebControls.Button
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCap As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSchedule As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEACheckFile1 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents LEACheckFile1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DQCCheck13 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck14 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DLOSS As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCCHECK15 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DYearSeason As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCCHECK16 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LPFASFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DPFASFile As System.Web.UI.HtmlControls.HtmlInputFile

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

        Response.Cookies("DevNo").Value = ""        '開發No, DevNoPicker使用
        Response.Cookies("PGM").Value = "SufaceSheet_01.aspx"
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
        'Check客戶樣品圖
        If DCustSampleFile.Visible Then
            If DCustSampleFile.PostedFile.FileName <> "" Then
                Message = "客戶樣品圖"
            End If
        End If
        'Check最終樣品圖
        If DFinalSampleFile.Visible Then
            If DFinalSampleFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "最終樣品圖"
                Else
                    Message = Message & ", " & "最終樣品圖"
                End If
            End If
        End If
        'Check品質依賴書
        If DQCReqFile.Visible Then
            If DQCReqFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "品質依賴書"
                Else
                    Message = Message & ", " & "品質依賴書"
                End If
            End If
        End If
        'Check電鍍膜厚
        ' If DQCCheck6File.Visible Then
        ' If DQCCheck6File.PostedFile.FileName <> "" Then
        'If Message = "" Then
        'Message = "電鍍膜厚"
        'Else
        '   Message = Message & ", " & "電鍍膜厚"
        'End If
        ' End If
        ' End If
        'Check Oeko-tex有害物質報告
        If DEACheckFile.Visible Then
            If DEACheckFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "Oeko-tex有害物質報告"
                Else
                    Message = Message & ", " & "Oeko-tex有害物質報告"
                End If
            End If
        End If
        'Check A01有害物質報告
        If DEACheckFile1.Visible Then
            If DEACheckFile1.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "A01有害物質報告"
                Else
                    Message = Message & ", " & "A01有害物質報告"
                End If
            End If
        End If
        'Check測試報告書
        If DQCFinalFile.Visible Then
            If DQCFinalFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "測試報告書"
                Else
                    Message = Message & ", " & "測試報告書"
                End If
            End If
        End If

        'Check測試報告書
        If DPFASFile.Visible Then
            If DPFASFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "PFAS報告"
                Else
                    Message = Message & ", " & "PFAS報告"
                End If
            End If
        End If


        'Check製造流程表
        If DManufFlowFile.Visible Then
            If DManufFlowFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "製造流程表"
                Else
                    Message = Message & ", " & "製造流程表"
                End If
            End If
        End If
        'Check報價單
        If DForcastFile.Visible Then
            If DForcastFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "報價單"
                Else
                    Message = Message & ", " & "報價單"
                End If
            End If
        End If
        'Check作業標準書
        If DOPManualFile.Visible Then
            If DOPManualFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "作業標準書"
                Else
                    Message = Message & ", " & "作業標準書"
                End If
            End If
        End If
        'Check切結書
        If DContactFile.Visible Then
            If DContactFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "切結書"
                Else
                    Message = Message & ", " & "切結書"
                End If
            End If
        End If

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("SufaceFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_SufaceSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_SufaceSheet")
        If DBDataSet1.Tables("F_SufaceSheet").Rows.Count > 0 Then
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("CustSampleFile") <> "" Then        '客戶樣品圖
                LCustSampleFile.ImageUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("CustSampleFile")
            Else
                LCustSampleFile.Visible = False
            End If
            DNo.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Date")                   '日期
            SetFieldData("Division", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Division"))  '部門
            SetFieldData("Person", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Person"))      '擔當
            DSpec.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Spec")                   '規格
            SetFieldData("Buyer", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Buyer"))        'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("SellVendor")       '委託廠商
            DReqDelDate.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ReqDelDate")       '希望交期
            DReqQty.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ReqQty")               '預估量
            DSliderSample.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("SliderSample")   '樣品拉頭
            SetFieldData("AttachSample", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("AttachSample"))    '附樣
            DORNO.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ORNO")                   'OR-NO
            DOrderTime.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OrderTime")         '下單時間
            DPrice.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Price")                 '價格
            DDevReason.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("DevReason")         '開發理由
            SetFieldData("YearSeason", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("YearSeason"))    '年季
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FinalSampleFile") <> "" Then       '最終樣品圖
                LFinalSampleFile.ImageUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FinalSampleFile")
            Else
                LFinalSampleFile.Visible = False
            End If
            SetFieldData("ManufType", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ManufType"))  '內製/外注
            SetFieldData("Suppiler", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Suppiler"))    '外注商
            DColor.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Color")                   'Color
            DQty.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Qty")                       '數量
            DCap.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Cap")
            DSchedule.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Schedule")
            DFReason.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FReason")


            SetFieldData("AllowSample", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("AllowSample"))    '限度樣品
            DBFinalDate.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("BFinalDate")        '預定完成日
            DCode.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Code")                    'Code
            DEnglishName.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EnglishName")      '英文名稱
            DLOSS.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("LOSS")                    'LOSS
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCReqFile") <> "" Then              '品質依賴書
                LQCReqFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCReqFile")
            Else
                LQCReqFile.Visible = False
            End If

            SetFieldData("QCCheck1", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck1"))   '口徑寸法
            SetFieldData("QCCheck2", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck2"))   '摺動抵抗
            SetFieldData("QCCheck3", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck3"))   'LOCK強度
            SetFieldData("QCCheck4", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck4"))   '90度強度
            SetFieldData("QCCheck5", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck5"))   '扭力
            SetFieldData("QCCheck15", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck15"))   'N-ANTI
            SetFieldData("QCCheck16", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck16")) 'FAFS


            '  If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck6File") <> "" Then           '電鍍膜厚
            '  LQCCheck6File.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck6File")
            'Else
            '   LQCCheck6File.Visible = False
            'End If

            SetFieldData("QCCheck7", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck7"))   '檢針
            SetFieldData("QCCheck8", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck8"))   'AATCC
            SetFieldData("QCCheck9", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck9"))   '乾洗
            SetFieldData("QCCheck10", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck10")) '鹽水噴霧
            SetFieldData("QCCheck11", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck11")) '一次密著
            SetFieldData("QCCheck12", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck12")) '二次密著
            SetFieldData("QCCheck13", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck13")) 'Oeko-tex
            SetFieldData("QCCheck14", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck14")) 'A01
            SetFieldData("EACheck1", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheck1"))   '有害物質
            DEADesc1.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EADesc1")              '有害物質備註

            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile") <> "" Then            'Oeko-tex有害物質報告
                LEACheckFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile")
            Else
                LEACheckFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile1") <> "" Then            'A01有害物質報告
                LEACheckFile1.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile1")
            Else
                LEACheckFile1.Visible = False
            End If

            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCFinalFile") <> "" Then            '測試報告書
                LQCFinalFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCFinalFile")
            Else
                LQCFinalFile.Visible = False
            End If


            DQCDate1.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCDate1")               '日期-1
            SetFieldData("QCResult1", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCResult1"))  '檢測結果-1
            DQCRemark1.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCRemark1")           '備註-1
            DQCDate2.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCDate2")               '日期-2
            SetFieldData("QCResult2", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCResult2"))  '檢測結果-2
            DQCRemark2.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCRemark2")           '備註-2
            DQCDate3.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCDate3")               '日期-3
            SetFieldData("QCResult3", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCResult3"))  '檢測結果-3
            DQCRemark3.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCRemark3")           '備註-3

            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ManufFlowFile") <> "" Then           '製造流程表
                LManufFlowFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ManufFlowFile")
            Else
                LManufFlowFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ForcastFile") <> "" Then             '報價單
                LForcastFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ForcastFile")
            Else
                LForcastFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OPManualFile") <> "" Then            '作業標準書
                LOPManualFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OPManualFile")
            Else
                LOPManualFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ContactFile") <> "" Then             '切結書
                LContactFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ContactFile")
            Else
                LContactFile.Visible = False
            End If

            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("PFASFile") <> "" Then             'PFAS報告
                LPFASFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("PFASFile")
            Else
                LPFASFile.Visible = False
            End If

            DOFormNo.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OFormNo")             '原表單
            DOFormSno.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OFormSno")           '原單號

            DQCLT.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCLT")    'QC-L/T

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
                DSufaceSheet1.Visible = True   '表單Sheet-1
                DSufaceSheet2.Visible = True   '表單Sheet-2
                DDescSheet.Visible = True       '說明Sheet
                DDelivery.Visible = True        '交期Sheet

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
                    Top = 1496
                Else
                    DReasonCode.Visible = False     '延遲理由代碼
                    DReason.Visible = False         '延遲理由
                    DReasonDesc.Visible = False     '延遲其他說明
                    Top = 1384
                End If

                '交期
                DBStartTime.Visible = True      '預定開始
                DBEndTime.Visible = True        '預定完成
                DAStartTime.Visible = True      '實際開始
                DAEndTime.Visible = True        '實際完成
                DOPReady.Visible = True     '工程履歷閱讀
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
                LCustSampleFile.Visible = True   '客戶樣品圖
                LFinalSampleFile.Visible = True  '最終樣品圖
                LQCReqFile.Visible = True        '品質依賴書
                '  LQCCheck6File.Visible = True     '電鍍膜厚
                LEACheckFile.Visible = True      'Oeko-tex有害物質報告
                LEACheckFile1.Visible = True     'A01有害物質報告
                LQCFinalFile.Visible = True      '測試報告書
                LPFASFile.Visible = True      '測試報告書

                LManufFlowFile.Visible = True    '製造流程表
                LForcastFile.Visible = True      '報價單
                LOPManualFile.Visible = True     '作業標準書
                LContactFile.Visible = True      '切結書
                LBefOP.Visible = True            '工程履歷
                '按鈕位置
                BNG1.Style.Add("Top", Top)     'NG1按鈕
                BNG2.Style.Add("Top", Top)     'NG2按鈕
                BSAVE.Style.Add("Top", Top)    '儲存按鈕
                BOK.Style.Add("Top", Top)      'OK按鈕
                DFormSno.Style.Add("Top", Top) '單號
            End If
        Else
            Top = 1200
            'Sheet隱藏
            DDescSheet.Visible = False  '說明Sheet
            DDelivery.Visible = False   '交期Sheet
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

            DOPReady.Visible = False    '工程履歷閱讀
            DOPReadyDesc.Visible = False

            '連結隱藏
            LCustSampleFile.Visible = False  '客戶樣品圖
            LFinalSampleFile.Visible = False '最終樣品圖
            LQCReqFile.Visible = False       '品質依賴書
            '  LQCCheck6File.Visible = False    '電鍍膜厚
            LEACheckFile.Visible = False     'Oeko-tex有害物質報告
            LEACheckFile1.Visible = False    'A01有害物質報告
            LQCFinalFile.Visible = False     '測試報告書
            LPFASFile.Visible = False     'FAfS
            LManufFlowFile.Visible = False   '製造流程表
            LForcastFile.Visible = False     '報價單
            LOPManualFile.Visible = False    '作業標準書
            LContactFile.Visible = False     '切結書
            LBefOP.Visible = False           '工程履歷
            '按鈕位置
            BNG1.Style.Add("Top", Top)     'NG1按鈕
            BNG2.Style.Add("Top", Top)     'NG2按鈕
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
        BIn.Attributes("onclick") = "DevNoPicker('In','000014');"       '內製委託書-No
        BOut.Attributes("onclick") = "DevNoPicker('Out','000014');"     '外注委託書-No
        BImport.Attributes("onclick") = "DevNoPicker('Import','000014');"     '進口/客供拉頭委託書-No

        BDate.Attributes("onclick") = "CalendarPicker('Form1.DDate');"              '日期
        BQCDate1.Attributes("onclick") = "CalendarPicker('Form1.DQCDate1');"        '日期
        BQCDate2.Attributes("onclick") = "CalendarPicker('Form1.DQCDate2');"        '日期
        BQCDate3.Attributes("onclick") = "CalendarPicker('Form1.DQCDate3');"        '日期

        BSpec.Attributes("onclick") = "SpecPicker('Form1.DSpec', 'SUFACE');"        'Spec
        BReqDelDate.Attributes("onclick") = "CalendarPicker('Form1.DReqDelDate');"  '希望交期
        BOrderTime.Attributes("onclick") = "CalendarPicker('Form1.DOrderTime');"    '下單時間
        BBFinalDate.Attributes("onclick") = "CalendarPicker('Form1.DBFinalDate');"  '預定完成日
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
                    Top = 1384
                Else
                    If DDelay.Visible = True Then
                        Top = 1496
                    Else
                        Top = 1384
                    End If
                End If
            End If
        Else
            Top = 1200
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
        '客戶樣品圖
        Select Case FindFieldInf("CustSampleFile")
            Case 0  '顯示
                DCustSampleFile.Visible = False
                DCustSampleFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DCustSampleFileRqd", "DCustSampleFile", "異常：需輸入客戶樣品圖")
                DCustSampleFile.Visible = True
                DCustSampleFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DCustSampleFile.Visible = True
                DCustSampleFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DCustSampleFile.Visible = False
        End Select

        'No
        Select Case FindFieldInf("No")
            Case 0  '顯示
                DNo.BackColor = Color.LightGray
                DNo.ReadOnly = True
                DNo.Visible = True
                BIn.Visible = False
                BOut.Visible = False
                BImport.Visible = False
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

        '部門
        Select Case FindFieldInf("Division")
            Case 0  '顯示
                DDivision.BackColor = Color.LightGray
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
        If pPost = "New" Then SetFieldData("Division", "ZZZZZZ")

        '擔當
        Select Case FindFieldInf("Person")
            Case 0  '顯示
                DPerson.BackColor = Color.LightGray
                DPerson.Visible = True
            Case 1  '修改+檢查
                DPerson.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPersonRqd", "DPerson", "異常：需輸入擔當")
                DPerson.Visible = True
            Case 2  '修改
                DPerson.BackColor = Color.Yellow
                DPerson.Visible = True
            Case Else   '隱藏
                DPerson.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Person", "ZZZZZZ")

        'Spec
        Select Case FindFieldInf("Spec")
            Case 0  '顯示
                DSpec.BackColor = Color.LightGray
                DSpec.ReadOnly = True
                DSpec.Visible = True
                BSpec.Visible = False
            Case 1  '修改+檢查
                DSpec.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSpecRqd", "DSpec", "異常：需輸入規格(Size, Chain Type, 胴體)")
                DSpec.Visible = True
                BSpec.Visible = True
            Case 2  '修改
                DSpec.BackColor = Color.Yellow
                DSpec.Visible = True
                BSpec.Visible = True
            Case Else   '隱藏
                DSpec.Visible = False
                BSpec.Visible = False
        End Select
        If pPost = "New" Then DSpec.Text = ""

        'Buyer
        Select Case FindFieldInf("Buyer")
            Case 0  '顯示
                DBuyer.BackColor = Color.LightGray
                DBuyer.Visible = True
            Case 1  '修改+檢查
                DBuyer.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerRqd", "DBuyer", "異常：需輸入Buyer")
                DBuyer.Visible = True
            Case 2  '修改
                DBuyer.BackColor = Color.Yellow
                DBuyer.Visible = True
            Case Else   '隱藏
                DBuyer.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Buyer", "ZZZZZZ")

        '委託廠商
        Select Case FindFieldInf("SellVendor")
            Case 0  '顯示
                DSellVendor.BackColor = Color.LightGray
                DSellVendor.ReadOnly = True
                DSellVendor.Visible = True
            Case 1  '修改+檢查
                DSellVendor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSellVendorRqd", "DSellVendor", "異常：需輸入委託廠商")
                DSellVendor.Visible = True
            Case 2  '修改
                DSellVendor.BackColor = Color.Yellow
                DSellVendor.Visible = True
            Case Else   '隱藏
                DSellVendor.Visible = False
        End Select
        If pPost = "New" Then DSellVendor.Text = ""

        '希望交期
        Select Case FindFieldInf("ReqDelDate")
            Case 0  '顯示
                DReqDelDate.BackColor = Color.LightGray
                DReqDelDate.ReadOnly = True
                DReqDelDate.Visible = True
                BReqDelDate.Visible = False
            Case 1  '修改+檢查
                DReqDelDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DReqDelDateRqd", "DReqDelDate", "異常：需輸入希望交期")
                DReqDelDate.Visible = True
                BReqDelDate.Visible = True
            Case 2  '修改
                DReqDelDate.BackColor = Color.Yellow
                DReqDelDate.Visible = True
                BReqDelDate.Visible = True
            Case Else   '隱藏
                DReqDelDate.Visible = False
                BReqDelDate.Visible = False
        End Select
        If pPost = "New" Then DReqDelDate.Text = CStr(DateTime.Now.Today)

        '預估量
        Select Case FindFieldInf("ReqQty")
            Case 0  '顯示
                DReqQty.BackColor = Color.LightGray
                DReqQty.ReadOnly = True
                DReqQty.Visible = True
            Case 1  '修改+檢查
                DReqQty.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DReqQtyRqd", "DReqQty", "異常：需輸入預估量")
                DReqQty.Visible = True
            Case 2  '修改
                DReqQty.BackColor = Color.Yellow
                DReqQty.Visible = True
            Case Else   '隱藏
                DReqQty.Visible = False
        End Select
        If pPost = "New" Then DReqQty.Text = ""

        '樣品拉頭
        Select Case FindFieldInf("SliderSample")
            Case 0  '顯示
                DSliderSample.BackColor = Color.LightGray
                DSliderSample.ReadOnly = True
                DSliderSample.Visible = True
            Case 1  '修改+檢查
                DSliderSample.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderSampleRqd", "DSliderSample", "異常：需輸入樣品拉頭")
                DSliderSample.Visible = True
            Case 2  '修改
                DSliderSample.BackColor = Color.Yellow
                DSliderSample.Visible = True
            Case Else   '隱藏
                DSliderSample.Visible = False
        End Select
        If pPost = "New" Then DSliderSample.Text = ""

        '附樣
        Select Case FindFieldInf("AttachSample")
            Case 0  '顯示
                DAttachSample.BackColor = Color.LightGray
                DAttachSample.Visible = True
            Case 1  '修改+檢查
                DAttachSample.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAttachSampleRqd", "DAttachSample", "異常：需輸入附樣")
                DAttachSample.Visible = True
            Case 2  '修改
                DAttachSample.BackColor = Color.Yellow
                DAttachSample.Visible = True
            Case Else   '隱藏
                DAttachSample.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AttachSample", "ZZZZZZ")

        '年季
        Select Case FindFieldInf("YearSeason")
            Case 0  '顯示
                DYearSeason.BackColor = Color.LightGray
                DYearSeason.Visible = True
            Case 1  '修改+檢查
                DYearSeason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DYearSeasonRqd", "DYearSeason", "異常：需輸年或季")
                DYearSeason.Visible = True
            Case 2  '修改
                DYearSeason.BackColor = Color.Yellow
                DYearSeason.Visible = True
            Case Else   '隱藏
                DYearSeason.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("YearSeason", "ZZZZZZ")

        'OR-NO
        Select Case FindFieldInf("ORNO")
            Case 0  '顯示
                DORNO.BackColor = Color.LightGray
                DORNO.ReadOnly = True
                DORNO.Visible = True
            Case 1  '修改+檢查
                DORNO.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DORNORqd", "DORNO", "異常：需輸入OR-NO")
                DORNO.Visible = True
            Case 2  '修改
                DORNO.BackColor = Color.Yellow
                DORNO.Visible = True
            Case Else   '隱藏
                DORNO.Visible = False
        End Select
        If pPost = "New" Then DORNO.Text = ""

        '下單時間
        Select Case FindFieldInf("OrderTime")
            Case 0  '顯示
                DOrderTime.BackColor = Color.LightGray
                DOrderTime.ReadOnly = True
                DOrderTime.Visible = True
                BOrderTime.Visible = False
            Case 1  '修改+檢查
                DOrderTime.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOrderTimeRqd", "DOrderTime", "異常：需輸入下單時間")
                DOrderTime.Visible = True
            Case 2  '修改
                DOrderTime.BackColor = Color.Yellow
                DOrderTime.Visible = True
            Case Else   '隱藏
                DOrderTime.Visible = False
        End Select
        If pPost = "New" Then DOrderTime.Text = ""

        '價格
        Select Case FindFieldInf("Price")
            Case 0  '顯示
                DPrice.BackColor = Color.LightGray
                DPrice.ReadOnly = True
                DPrice.Visible = True
            Case 1  '修改+檢查
                DPrice.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPriceRqd", "DPrice", "異常：需輸入價格")
                DPrice.Visible = True
            Case 2  '修改
                DPrice.BackColor = Color.Yellow
                DPrice.Visible = True
            Case Else   '隱藏
                DPrice.Visible = False
        End Select
        If pPost = "New" Then DPrice.Text = ""

        '開發理由
        Select Case FindFieldInf("DevReason")
            Case 0  '顯示
                DDevReason.BackColor = Color.LightGray
                DDevReason.ReadOnly = True
                DDevReason.Visible = True
            Case 1  '修改+檢查
                DDevReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDevReasonRqd", "DDevReason", "異常：需輸入開發理由")
                DDevReason.Visible = True
            Case 2  '修改
                DDevReason.BackColor = Color.Yellow
                DDevReason.Visible = True
            Case Else   '隱藏
                DDevReason.Visible = False
        End Select
        If pPost = "New" Then DDevReason.Text = ""

        '最終樣品圖
        Select Case FindFieldInf("FinalSampleFile")
            Case 0  '顯示
                DFinalSampleFile.Visible = False
                DFinalSampleFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DFinalSampleFileRqd", "DFinalSampleFile", "異常：需輸入最終樣品圖")
                DFinalSampleFile.Visible = True
                DFinalSampleFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DFinalSampleFile.Visible = True
                DFinalSampleFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DFinalSampleFile.Visible = False
        End Select

        '內製/外注
        Select Case FindFieldInf("ManufType")
            Case 0  '顯示
                DManufType.BackColor = Color.LightGray
                DManufType.Visible = True
            Case 1  '修改+檢查
                DManufType.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DManufTypeRqd", "DManufType", "異常：需輸入內製/外注")
                DManufType.Visible = True
            Case 2  '修改
                DManufType.BackColor = Color.Yellow
                DManufType.Visible = True
            Case Else   '隱藏
                DManufType.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ManufType", "ZZZZZZ")

        '外注商
        Select Case FindFieldInf("Suppiler")
            Case 0  '顯示
                DSuppiler.BackColor = Color.LightGray
                DSuppiler.Visible = True
            Case 1  '修改+檢查
                DSuppiler.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSuppilerRqd", "DSuppiler", "異常：需輸入外注商")
                DSuppiler.Visible = True
            Case 2  '修改
                DSuppiler.BackColor = Color.Yellow
                DSuppiler.Visible = True
            Case Else   '隱藏
                DSuppiler.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Suppiler", "ZZZZZZ")

        '顏色
        Select Case FindFieldInf("Color")
            Case 0  '顯示
                DColor.BackColor = Color.LightGray
                DColor.ReadOnly = True
                DColor.Visible = True
            Case 1  '修改+檢查
                DColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DColorRqd", "DColor", "異常：需輸入顏色")
                DColor.Visible = True
            Case 2  '修改
                DColor.BackColor = Color.Yellow
                DColor.Visible = True
            Case Else   '隱藏
                DColor.Visible = False
        End Select
        If pPost = "New" Then DColor.Text = ""

        '數量
        Select Case FindFieldInf("Qty")
            Case 0  '顯示
                DQty.BackColor = Color.LightGray
                DQty.ReadOnly = True
                DQty.Visible = True
            Case 1  '修改+檢查
                DQty.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQtyRqd", "DQty", "異常：需輸入數量")
                DQty.Visible = True
            Case 2  '修改
                DQty.BackColor = Color.Yellow
                DQty.Visible = True
            Case Else   '隱藏
                DQty.Visible = False
        End Select
        If pPost = "New" Then DQty.Text = ""

        '日產能
        Select Case FindFieldInf("Cap")
            Case 0  '顯示
                DCap.BackColor = Color.LightGray
                DCap.ReadOnly = True
                DCap.Visible = True
            Case 1  '修改+檢查
                DCap.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCapRqd", "DCap", "異常：需輸入數量")
                DCap.Visible = True
            Case 2  '修改
                DCap.BackColor = Color.Yellow
                DCap.Visible = True
            Case Else   '隱藏
                DCap.Visible = False
        End Select
        If pPost = "New" Then DCap.Text = ""

        '日產能
        Select Case FindFieldInf("Schedule")
            Case 0  '顯示
                DSchedule.BackColor = Color.LightGray
                DSchedule.ReadOnly = True
                DSchedule.Visible = True
            Case 1  '修改+檢查
                DSchedule.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DScheduleRqd", "DSchedule", "異常：需輸入數量")
                DSchedule.Visible = True
            Case 2  '修改
                DSchedule.BackColor = Color.Yellow
                DSchedule.Visible = True
            Case Else   '隱藏
                DSchedule.Visible = False
        End Select
        If pPost = "New" Then DSchedule.Text = ""

        '理由
        Select Case FindFieldInf("Schedule")
            Case 0  '顯示
                DFReason.BackColor = Color.LightGray
                DFReason.ReadOnly = True
                DFReason.Visible = True
            Case 1  '修改+檢查
                DFReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFReasonRqd", "DFReason", "異常：需輸入理由")
                DFReason.Visible = True
            Case 2  '修改
                DFReason.BackColor = Color.Yellow
                DFReason.Visible = True
            Case Else   '隱藏
                DFReason.Visible = False
        End Select
        If pPost = "New" Then DFReason.Text = ""

        '限度樣品
        Select Case FindFieldInf("AllowSample")
            Case 0  '顯示
                DAllowSample.BackColor = Color.LightGray
                DAllowSample.Visible = True
            Case 1  '修改+檢查
                DAllowSample.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAllowSampleRqd", "DAllowSample", "異常：需輸入限度樣品")
                DAllowSample.Visible = True
            Case 2  '修改
                DAllowSample.BackColor = Color.Yellow
                DAllowSample.Visible = True
            Case Else   '隱藏
                DAllowSample.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AllowSample", "ZZZZZZ")

        '預定完成日
        Select Case FindFieldInf("BFinalDate")
            Case 0  '顯示
                DBFinalDate.BackColor = Color.LightGray
                DBFinalDate.ReadOnly = True
                DBFinalDate.Visible = True
                BBFinalDate.Visible = False
            Case 1  '修改+檢查
                DBFinalDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBFinalDateRqd", "DBFinalDate", "異常：需輸入預定完成日")
                DBFinalDate.Visible = True
                BBFinalDate.Visible = True
            Case 2  '修改
                DBFinalDate.BackColor = Color.Yellow
                DBFinalDate.Visible = True
                BBFinalDate.Visible = True
            Case Else   '隱藏
                DBFinalDate.Visible = False
                BBFinalDate.Visible = False
        End Select
        If pPost = "New" Then DBFinalDate.Text = ""

        'Code
        Select Case FindFieldInf("Code")
            Case 0  '顯示
                DCode.BackColor = Color.LightGray
                DCode.ReadOnly = True
                DCode.Visible = True
            Case 1  '修改+檢查
                DCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCodeRqd", "DCode", "異常：需輸入Code")
                DCode.Visible = True
            Case 2  '修改
                DCode.BackColor = Color.Yellow
                DCode.Visible = True
            Case Else   '隱藏
                DCode.Visible = False
        End Select
        If pPost = "New" Then DCode.Text = ""

        '英文名稱
        Select Case FindFieldInf("EnglishName")
            Case 0  '顯示
                DEnglishName.BackColor = Color.LightGray
                DEnglishName.ReadOnly = True
                DEnglishName.Visible = True
            Case 1  '修改+檢查
                DEnglishName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEnglishNameRqd", "DEnglishName", "異常：需輸入英文名稱")
                DEnglishName.Visible = True
            Case 2  '修改
                DEnglishName.BackColor = Color.Yellow
                DEnglishName.Visible = True
            Case Else   '隱藏
                DEnglishName.Visible = False
        End Select
        If pPost = "New" Then DEnglishName.Text = ""

        'LOSS
        Select Case FindFieldInf("LOSS")
            Case 0  '顯示
                DLOSS.BackColor = Color.LightGray
                DLOSS.ReadOnly = True
                DLOSS.Visible = True
            Case 1  '修改+檢查
                DLOSS.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DLOSSRqd", "DLOSS", "異常：需輸入LOSS")
                DLOSS.Visible = True
            Case 2  '修改
                DLOSS.BackColor = Color.Yellow
                DLOSS.Visible = True
            Case Else   '隱藏
                DLOSS.Visible = False
        End Select
        If pPost = "New" Then DLOSS.Text = ""

        '品質依賴書
        Select Case FindFieldInf("QCReqFile")
            Case 0  '顯示
                DQCReqFile.Visible = False
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQCReqFileRqd", "DQCReqFile", "異常：需輸入品質依賴書")
                DQCReqFile.Visible = True
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DQCReqFile.Visible = True
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DQCReqFile.Visible = False
        End Select

        '口徑寸法
        Select Case FindFieldInf("QCCheck1")
            Case 0  '顯示
                DQCCheck1.BackColor = Color.LightGray
                DQCCheck1.Visible = True
            Case 1  '修改+檢查
                DQCCheck1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck1Rqd", "DQCCheck1", "異常：需輸入口徑寸法")
                DQCCheck1.Visible = True
            Case 2  '修改
                DQCCheck1.BackColor = Color.Yellow
                DQCCheck1.Visible = True
            Case Else   '隱藏
                DQCCheck1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck1", "ZZZZZZ")

        '摺動抵抗
        Select Case FindFieldInf("QCCheck2")
            Case 0  '顯示
                DQCCheck2.BackColor = Color.LightGray
                DQCCheck2.Visible = True
            Case 1  '修改+檢查
                DQCCheck2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck2Rqd", "DQCCheck2", "異常：需輸入摺動抵抗")
                DQCCheck2.Visible = True
            Case 2  '修改
                DQCCheck2.BackColor = Color.Yellow
                DQCCheck2.Visible = True
            Case Else   '隱藏
                DQCCheck2.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck2", "ZZZZZZ")

        'LOCK強度
        Select Case FindFieldInf("QCCheck3")
            Case 0  '顯示
                DQCCheck3.BackColor = Color.LightGray
                DQCCheck3.Visible = True
            Case 1  '修改+檢查
                DQCCheck3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck3Rqd", "DQCCheck3", "異常：需輸入LOCK強度")
                DQCCheck3.Visible = True
            Case 2  '修改
                DQCCheck3.BackColor = Color.Yellow
                DQCCheck3.Visible = True
            Case Else   '隱藏
                DQCCheck3.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck3", "ZZZZZZ")

        '90度強度
        Select Case FindFieldInf("QCCheck4")
            Case 0  '顯示
                DQCCheck4.BackColor = Color.LightGray
                DQCCheck4.Visible = True
            Case 1  '修改+檢查
                DQCCheck4.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck4Rqd", "DQCCheck4", "異常：需輸入90度強度")
                DQCCheck4.Visible = True
            Case 2  '修改
                DQCCheck4.BackColor = Color.Yellow
                DQCCheck4.Visible = True
            Case Else   '隱藏
                DQCCheck4.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck4", "ZZZZZZ")

        '扭力
        Select Case FindFieldInf("QCCheck5")
            Case 0  '顯示
                DQCCheck5.BackColor = Color.LightGray
                DQCCheck5.Visible = True
            Case 1  '修改+檢查
                DQCCheck5.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck5Rqd", "DQCCheck5", "異常：需輸入扭力")
                DQCCheck5.Visible = True
            Case 2  '修改
                DQCCheck5.BackColor = Color.Yellow
                DQCCheck5.Visible = True
            Case Else   '隱藏
                DQCCheck5.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck5", "ZZZZZZ")


        'N-ANTI
        Select Case FindFieldInf("QCCheck15")
            Case 0  '顯示
                DQCCHECK15.BackColor = Color.LightGray
                DQCCHECK15.Visible = True
            Case 1  '修改+檢查
                DQCCHECK15.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("QCCheck15Rqd", "DQCCheck15", "異常：需輸入扭力")
                DQCCHECK15.Visible = True
            Case 2  '修改
                DQCCHECK15.BackColor = Color.Yellow
                DQCCHECK15.Visible = True
            Case Else   '隱藏
                DQCCHECK15.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck15", "ZZZZZZ")


        'PFAS
        Select Case FindFieldInf("QCCheck15")
            Case 0  '顯示
                DQCCHECK16.BackColor = Color.LightGray
                DQCCHECK16.Visible = True
            Case 1  '修改+檢查
                DQCCHECK16.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("QCCheck16Rqd", "DQCCHECK16", "異常：需輸入PFAS")
                DQCCHECK16.Visible = True
            Case 2  '修改
                DQCCHECK16.BackColor = Color.Yellow
                DQCCHECK16.Visible = True
            Case Else   '隱藏
                DQCCHECK16.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck16", "ZZZZZZ")

        '電鍍膜厚
        ' Select Case FindFieldInf("QCCheck6File")
        '     Case 0  '顯示
        ' DQCCheck6File.Visible = False
        ' DQCCheck6File.Style.Add("BACKGROUND-COLOR", "LightGrey")
        '     Case 1  '修改+檢查
        ' ShowRequiredFieldValidator("DQCCheck6FileRqd", "DQCCheck6File", "異常：需輸入電鍍膜厚附件")
        ' DQCCheck6File.Visible = True
        ' DQCCheck6File.Style.Add("BACKGROUND-COLOR", "GreenYellow")
        '     Case 2  '修改
        ' DQCCheck6File.Visible = True
        ' DQCCheck6File.Style.Add("BACKGROUND-COLOR", "Yellow")
        '     Case Else   '隱藏
        ' DQCCheck6File.Visible = False
        ' End Select

        '檢針
        Select Case FindFieldInf("QCCheck7")
            Case 0  '顯示
                DQCCheck7.BackColor = Color.LightGray
                DQCCheck7.Visible = True
            Case 1  '修改+檢查
                DQCCheck7.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck7Rqd", "DQCCheck7", "異常：需輸入檢針")
                DQCCheck7.Visible = True
            Case 2  '修改
                DQCCheck7.BackColor = Color.Yellow
                DQCCheck7.Visible = True
            Case Else   '隱藏
                DQCCheck7.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck7", "ZZZZZZ")

        'AATCC
        Select Case FindFieldInf("QCCheck8")
            Case 0  '顯示
                DQCCheck8.BackColor = Color.LightGray
                DQCCheck8.Visible = True
            Case 1  '修改+檢查
                DQCCheck8.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck8Rqd", "DQCCheck8", "異常：需輸入AATCC")
                DQCCheck8.Visible = True
            Case 2  '修改
                DQCCheck8.BackColor = Color.Yellow
                DQCCheck8.Visible = True
            Case Else   '隱藏
                DQCCheck8.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck8", "ZZZZZZ")

        '乾洗
        Select Case FindFieldInf("QCCheck9")
            Case 0  '顯示
                DQCCheck9.BackColor = Color.LightGray
                DQCCheck9.Visible = True
            Case 1  '修改+檢查
                DQCCheck9.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck9Rqd", "DQCCheck9", "異常：需輸入乾洗")
                DQCCheck9.Visible = True
            Case 2  '修改
                DQCCheck9.BackColor = Color.Yellow
                DQCCheck9.Visible = True
            Case Else   '隱藏
                DQCCheck9.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck9", "ZZZZZZ")

        '鹽水噴霧
        Select Case FindFieldInf("QCCheck10")
            Case 0  '顯示
                DQCCheck10.BackColor = Color.LightGray
                DQCCheck10.Visible = True
            Case 1  '修改+檢查
                DQCCheck10.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck10Rqd", "DQCCheck10", "異常：需輸入鹽水噴霧")
                DQCCheck10.Visible = True
            Case 2  '修改
                DQCCheck10.BackColor = Color.Yellow
                DQCCheck10.Visible = True
            Case Else   '隱藏
                DQCCheck10.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck10", "ZZZZZZ")

        '一次密著
        Select Case FindFieldInf("QCCheck11")
            Case 0  '顯示
                DQCCheck11.BackColor = Color.LightGray
                DQCCheck11.Visible = True
            Case 1  '修改+檢查
                DQCCheck11.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck11Rqd", "DQCCheck11", "異常：需輸入一次密著")
                DQCCheck11.Visible = True
            Case 2  '修改
                DQCCheck11.BackColor = Color.Yellow
                DQCCheck11.Visible = True
            Case Else   '隱藏
                DQCCheck11.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck11", "ZZZZZZ")

        '二次密著,Oeko-tex,A01
        Select Case FindFieldInf("QCCheck12")
            Case 0  '顯示
                '二次密著
                DQCCheck12.BackColor = Color.LightGray
                DQCCheck12.Visible = True
                'Oeko-tex
                DQCCheck13.BackColor = Color.LightGray
                DQCCheck13.Visible = True
                'A01
                DQCCheck14.BackColor = Color.LightGray
                DQCCheck14.Visible = True
            Case 1  '修改+檢查
                '二次密著
                DQCCheck12.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck12Rqd", "DQCCheck12", "異常：需輸入二次密著")
                DQCCheck12.Visible = True
                'Oeko-tex
                DQCCheck13.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck13Rqd", "DQCCheck13", "異常：需輸入Oeko-Tex")
                DQCCheck13.Visible = True
                'A01
                DQCCheck14.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck14Rqd", "DQCCheck14", "異常：需輸入A01")
                DQCCheck14.Visible = True
            Case 2  '修改
                '二次密著
                DQCCheck12.BackColor = Color.Yellow
                DQCCheck12.Visible = True
                'Oeko-tex
                DQCCheck13.BackColor = Color.Yellow
                DQCCheck13.Visible = True
                'A01
                DQCCheck14.BackColor = Color.Yellow
                DQCCheck14.Visible = True
            Case Else   '隱藏
                '二次密著
                DQCCheck12.Visible = False
                'Oeko-tex
                DQCCheck13.Visible = False
                'A01
                DQCCheck14.Visible = False
        End Select
        If pPost = "New" Then
            SetFieldData("QCCheck12", "ZZZZZZ")
            SetFieldData("QCCheck13", "ZZZZZZ")
            SetFieldData("QCCheck14", "ZZZZZZ")
        End If

        '有害物質
        Select Case FindFieldInf("EACheck1")
            Case 0  '顯示
                DEACheck1.BackColor = Color.LightGray
                DEACheck1.Visible = True
            Case 1  '修改+檢查
                DEACheck1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEACheck1Rqd", "DEACheck1", "異常：需輸入有害物質")
                DEACheck1.Visible = True
            Case 2  '修改
                DEACheck1.BackColor = Color.Yellow
                DEACheck1.Visible = True
            Case Else   '隱藏
                DEACheck1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("EACheck1", "ZZZZZZ")

        '有害物質備註
        Select Case FindFieldInf("EADesc1")
            Case 0  '顯示
                DEADesc1.BackColor = Color.LightGray
                DEADesc1.ReadOnly = True
                DEADesc1.Visible = True
            Case 1  '修改+檢查
                DEADesc1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEADesc1Rqd", "DEADesc1", "異常：需輸入有害物質備註")
                DEADesc1.Visible = True
            Case 2  '修改
                DEADesc1.BackColor = Color.Yellow
                DEADesc1.Visible = True
            Case Else   '隱藏
                DEADesc1.Visible = False
        End Select
        If pPost = "New" Then DEADesc1.Text = ""

        'Oeko-tex,A01有害物質報告
        Select Case FindFieldInf("EACheckFile")
            Case 0  '顯示
                'Oeko-tex有害物質報告
                DEACheckFile.Visible = False
                DEACheckFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                'A01有害物質報告
                DEACheckFile1.Visible = False
                DEACheckFile1.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                'Oeko-tex有害物質報告
                ShowRequiredFieldValidator("DEACheckFileRqd", "DEACheckFile", "異常：需輸入Oeko-tex有害物質報告")
                DEACheckFile.Visible = True
                DEACheckFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                'A01有害物質報告
                ShowRequiredFieldValidator("DEACheckFile1Rqd", "DEACheckFile1", "異常：需輸入A01有害物質報告")
                DEACheckFile1.Visible = True
                DEACheckFile1.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                'Oeko-tex有害物質報告
                DEACheckFile.Visible = True
                DEACheckFile.Style.Add("BACKGROUND-COLOR", "Yellow")
                'A01有害物質報告
                DEACheckFile1.Visible = True
                DEACheckFile1.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                'Oeko-tex有害物質報告
                DEACheckFile.Visible = False
                'A01有害物質報告
                DEACheckFile1.Visible = False
        End Select

        '測試報告書
        Select Case FindFieldInf("QCFinalFile")
            Case 0  '顯示
                DQCFinalFile.Visible = False
                DQCFinalFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQCFinalFileRqd", "DQCFinalFile", "異常：需輸入測試報告書")
                DQCFinalFile.Visible = True
                DQCFinalFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DQCFinalFile.Visible = True
                DQCFinalFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DQCFinalFile.Visible = False
        End Select

        'FAFS
        Select Case FindFieldInf("QCFinalFile")
            Case 0  '顯示
                DPFASFile.Visible = False
                DPFASFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DPFASFileRqd", "DPFASFile", "異常：需輸入PFAS報告")
                DPFASFile.Visible = True
                DPFASFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DPFASFile.Visible = True
                DPFASFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DPFASFile.Visible = False
        End Select



        '日期-1
        Select Case FindFieldInf("QCDate1")
            Case 0  '顯示
                DQCDate1.BackColor = Color.LightGray
                DQCDate1.ReadOnly = True
                DQCDate1.Visible = True
                BQCDate1.Visible = False
            Case 1  '修改+檢查
                DQCDate1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCDate1Rqd", "DQCDate1", "異常：需輸入日期-1")
                DQCDate1.Visible = True
                BQCDate1.Visible = True
            Case 2  '修改
                DQCDate1.BackColor = Color.Yellow
                DQCDate1.Visible = True
                BQCDate1.Visible = True
            Case Else   '隱藏
                DQCDate1.Visible = False
                BQCDate1.Visible = False
        End Select
        If pPost = "New" Then DQCDate1.Text = ""

        '檢測結果-1
        Select Case FindFieldInf("QCResult1")
            Case 0  '顯示
                DQCResult1.BackColor = Color.LightGray
                DQCResult1.Visible = True
            Case 1  '修改+檢查
                DQCResult1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCResult1Rqd", "DQCResult1", "異常：需輸入檢測結果-1")
                DQCResult1.Visible = True
            Case 2  '修改
                DQCResult1.BackColor = Color.Yellow
                DQCResult1.Visible = True
            Case Else   '隱藏
                DQCResult1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCResult1", "ZZZZZZ")

        '備註-1
        Select Case FindFieldInf("QCRemark1")
            Case 0  '顯示
                DQCRemark1.BackColor = Color.LightGray
                DQCRemark1.ReadOnly = True
                DQCRemark1.Visible = True
            Case 1  '修改+檢查
                DQCRemark1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCRemark1Rqd", "DQCRemark1", "異常：需輸入備註-1")
                DQCRemark1.Visible = True
            Case 2  '修改
                DQCRemark1.BackColor = Color.Yellow
                DQCRemark1.Visible = True
            Case Else   '隱藏
                DQCRemark1.Visible = False
        End Select
        If pPost = "New" Then DQCRemark1.Text = ""

        '日期-2
        Select Case FindFieldInf("QCDate2")
            Case 0  '顯示
                DQCDate2.BackColor = Color.LightGray
                DQCDate2.ReadOnly = True
                DQCDate2.Visible = True
                BQCDate2.Visible = False

                DQCDate3.BackColor = Color.LightGray
                DQCDate3.ReadOnly = True
                DQCDate3.Visible = True
                BQCDate3.Visible = False
            Case 1  '修改+檢查
                DQCDate2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCDate2Rqd", "DQCDate2", "異常：需輸入日期-2")
                DQCDate2.Visible = True
                BQCDate2.Visible = True

                DQCDate3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCDate3Rqd", "DQCDate3", "異常：需輸入日期-3")
                DQCDate3.Visible = True
                BQCDate3.Visible = True
            Case 2  '修改
                DQCDate2.BackColor = Color.Yellow
                DQCDate2.Visible = True
                BQCDate2.Visible = True

                DQCDate3.BackColor = Color.Yellow
                DQCDate3.Visible = True
                BQCDate1.Visible = True
            Case Else   '隱藏
                DQCDate2.Visible = False
                BQCDate2.Visible = False
                DQCDate3.Visible = False
                BQCDate3.Visible = False

        End Select
        If pPost = "New" Then DQCDate2.Text = ""

        '檢測結果-2
        Select Case FindFieldInf("QCResult2")
            Case 0  '顯示
                DQCResult2.BackColor = Color.LightGray
                DQCResult2.Visible = True

                DQCResult3.BackColor = Color.LightGray
                DQCResult3.Visible = True
            Case 1  '修改+檢查
                DQCResult2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCResult2Rqd", "DQCResult2", "異常：需輸入檢測結果-2")
                DQCResult2.Visible = True

                DQCResult3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCResult3Rqd", "DQCResult3", "異常：需輸入檢測結果-3")
                DQCResult3.Visible = True
            Case 2  '修改
                DQCResult2.BackColor = Color.Yellow
                DQCResult2.Visible = True

                DQCResult3.BackColor = Color.Yellow
                DQCResult3.Visible = True
            Case Else   '隱藏
                DQCResult2.Visible = False
                DQCResult3.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCResult2", "ZZZZZZ")

        '備註-2 備註-3 寫在一起 20102/4 jessica 修改
        Select Case FindFieldInf("QCRemark2")
            Case 0  '顯示
                DQCRemark2.BackColor = Color.LightGray
                DQCRemark2.ReadOnly = True
                DQCRemark2.Visible = True

                DQCRemark3.BackColor = Color.LightGray
                DQCRemark3.ReadOnly = True
                DQCRemark3.Visible = True

            Case 1  '修改+檢查
                DQCRemark2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCRemark2Rqd", "DQCRemark2", "異常：需輸入備註-2")
                DQCRemark2.Visible = True

                DQCRemark3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCRemark3Rqd", "DQCRemark3", "異常：需輸入備註-3")
                DQCRemark3.Visible = True

            Case 2  '修改
                DQCRemark2.BackColor = Color.Yellow
                DQCRemark2.Visible = True

                DQCRemark3.BackColor = Color.Yellow
                DQCRemark3.Visible = True

            Case Else   '隱藏
                DQCRemark2.Visible = False
                DQCRemark3.Visible = False

        End Select
        If pPost = "New" Then DQCRemark2.Text = ""
        If pPost = "New" Then DQCDate3.Text = ""
        If pPost = "New" Then SetFieldData("QCResult3", "ZZZZZZ")
        If pPost = "New" Then DQCRemark3.Text = ""

        '製造流程表
        Select Case FindFieldInf("ManufFlowFile")
            Case 0  '顯示
                DManufFlowFile.Visible = False
                DManufFlowFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DManufFlowFileRqd", "DManufFlowFile", "異常：需輸入製造流程表")
                DManufFlowFile.Visible = True
                DManufFlowFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DManufFlowFile.Visible = True
                DManufFlowFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DManufFlowFile.Visible = False
        End Select

        '報價單
        Select Case FindFieldInf("ForcastFile")
            Case 0  '顯示
                DForcastFile.Visible = False
                DForcastFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DForcastFileRqd", "DForcastFile", "異常：需輸入報價單")
                DForcastFile.Visible = True
                DForcastFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DForcastFile.Visible = True
                DForcastFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DForcastFile.Visible = False
        End Select

        '作業標準書
        Select Case FindFieldInf("OPManualFile")
            Case 0  '顯示
                DOPManualFile.Visible = False
                DOPManualFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DOPManualFileRqd", "DOPManualFile", "異常：需輸入作業標準書")
                DOPManualFile.Visible = True
                DOPManualFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DOPManualFile.Visible = True
                DOPManualFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DOPManualFile.Visible = False
        End Select

        '切結書
        Select Case FindFieldInf("ContactFile")
            Case 0  '顯示
                DContactFile.Visible = False
                DContactFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DContactFileRqd", "DContactFile", "異常：需輸入切結書")
                DContactFile.Visible = True
                DContactFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DContactFile.Visible = True
                DContactFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DContactFile.Visible = False
        End Select

        'QC-L/T
        Select Case FindFieldInf("QCLT")
            Case 0  '顯示
                DQCLT.BackColor = Color.LightGray
                DQCLT.ReadOnly = True
                DQCLT.Visible = True
            Case 1  '修改+檢查
                DQCLT.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCLTRqd", "DQCLT", "異常：需輸入QC L/T")
                DQCLT.Visible = True
            Case 2  '修改
                DQCLT.BackColor = Color.Yellow
                DQCLT.Visible = True
            Case Else   '隱藏
                DQCLT.Visible = False
        End Select
        If pPost = "New" Then DQCLT.Text = ""

        ''**********************************************************************************************
        '難易度
        'Select Case FindFieldInf("Level")
        '    Case 0  '顯示
        'DLevel.BackColor = Color.LightGray
        'DLevel.Visible = True
        '    Case 1  '修改+檢查
        'DLevel.BackColor = Color.GreenYellow
        'ShowRequiredFieldValidator("DLevelRqd", "DLevel", "異常：需輸入難易度")
        'DLevel.Visible = True
        '    Case 2  '修改
        'DLevel.BackColor = Color.Yellow
        'DLevel.Visible = True
        '    Case Else   '隱藏
        'DLevel.Visible = False
        'End Select
        'If pPost = "New" Then SetFieldData("Level", "ZZZZZZ")
        ''**********************************************************************************************

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

        '部門
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

        '擔當
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

        'Buyer
        If pFieldName = "Buyer" Then
            DBuyer.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBuyer.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='700' and DKey='BUYER' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBuyer.Items.Add(ListItem1)
                Next
            End If
        End If

        '附樣
        If pFieldName = "AttachSample" Then
            DAttachSample.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAttachSample.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='ATTACHSAMPLE' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAttachSample.Items.Add(ListItem1)
                Next
            End If
        End If


        '年季
        If pFieldName = "YearSeason" Then
            DYearSeason.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DYearSeason.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='YearSeason' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                DYearSeason.Items.Add("")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DYearSeason.Items.Add(ListItem1)
                Next
            End If
        End If

        '內製/外注
        If pFieldName = "ManufType" Then
            DManufType.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DManufType.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='MANUFTYPE' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DManufType.Items.Add(ListItem1)
                Next
            End If
        End If

        'Suppiler
        If pFieldName = "Suppiler" Then
            DSuppiler.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSuppiler.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='700' and DKey='SUPPLIER' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSuppiler.Items.Add(ListItem1)
                Next
            End If
        End If

        '限度樣品
        If pFieldName = "AllowSample" Then
            DAllowSample.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAllowSample.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='ALLOWSAMPLE' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAllowSample.Items.Add(ListItem1)
                Next
            End If
        End If

        '口徑寸法
        If pFieldName = "QCCheck1" Then
            DQCCheck1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck1.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK1' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck1.Items.Add(ListItem1)
                Next
            End If
        End If

        '摺動抵抗
        If pFieldName = "QCCheck2" Then
            DQCCheck2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck2.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK2' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck2.Items.Add(ListItem1)
                Next
            End If
        End If

        'LOCK強度
        If pFieldName = "QCCheck3" Then
            DQCCheck3.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck3.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK3' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck3.Items.Add(ListItem1)
                Next
            End If
        End If

        '90度強度
        If pFieldName = "QCCheck4" Then
            DQCCheck4.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck4.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK4' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck4.Items.Add(ListItem1)
                Next
            End If
        End If

        '扭力
        If pFieldName = "QCCheck5" Then
            DQCCheck5.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck5.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK5' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck5.Items.Add(ListItem1)
                Next
            End If
        End If


        'N-ANTI
        If pFieldName = "QCCheck15" Then
            DQCCHECK15.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCHECK15.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK15' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCHECK15.Items.Add(ListItem1)
                Next
            End If
        End If

        'PFAS
        If pFieldName = "QCCheck16" Then
            DQCCHECK16.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCHECK16.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK16' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCHECK16.Items.Add(ListItem1)
                Next
            End If
        End If




        '檢針
        If pFieldName = "QCCheck7" Then
            DQCCheck7.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck7.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK7' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck7.Items.Add(ListItem1)
                Next
            End If
        End If

        'AATCC
        If pFieldName = "QCCheck8" Then
            DQCCheck8.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck8.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK8' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck8.Items.Add(ListItem1)
                Next
            End If
        End If

        '乾洗
        If pFieldName = "QCCheck9" Then
            DQCCheck9.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck9.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK9' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck9.Items.Add(ListItem1)
                Next
            End If
        End If

        '鹽水噴霧
        If pFieldName = "QCCheck10" Then
            DQCCheck10.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck10.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK10' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck10.Items.Add(ListItem1)
                Next
            End If
        End If

        '一次密著
        If pFieldName = "QCCheck11" Then
            DQCCheck11.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck11.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK11' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck11.Items.Add(ListItem1)
                Next
            End If
        End If

        '二次密著,Oeko-tex,A01
        If pFieldName = "QCCheck12" Then
            '二次密著
            DQCCheck12.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck12.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK12' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck12.Items.Add(ListItem1)
                Next
            End If
        End If

        'Oeko-tex
        If pFieldName = "QCCheck13" Then
            DQCCheck13.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck13.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK13' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck13.Items.Add(ListItem1)
                Next
            End If
        End If

        'A01
        If pFieldName = "QCCheck14" Then
            DQCCheck14.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck14.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK14' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck14.Items.Add(ListItem1)
                Next
            End If
        End If

        '有害物質
        If pFieldName = "EACheck1" Then
            DEACheck1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DEACheck1.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='EACHECK1' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DEACheck1.Items.Add(ListItem1)
                Next
            End If
        End If

        '檢測結果-1
        If pFieldName = "QCResult1" Then
            DQCResult1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCResult1.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCRESULT1' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCResult1.Items.Add(ListItem1)
                Next
            End If
        End If

        '檢測結果-2
        If pFieldName = "QCResult2" Then
            DQCResult2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCResult2.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCRESULT2' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCResult2.Items.Add(ListItem1)
                Next
            End If
        End If

        '檢測結果-3
        If pFieldName = "QCResult3" Then
            DQCResult3.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCResult3.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCRESULT3' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCResult3.Items.Add(ListItem1)
                Next
            End If
        End If

        '難易度
        'If pFieldName = "Level" Then
        'DLevel.Items.Clear()
        'If idx = 0 Then
        'If pName <> "ZZZZZZ" Then
        'Dim ListItem1 As New ListItem
        'ListItem1.Text = pName
        'ListItem1.Value = pName
        'DLevel.Items.Add(ListItem1)
        'End If
        'Else
        'SQL = "Select * From M_Referp Where Cat='007' and (DKey like 'In%' or DKey = '') Order by DKey, Data "
        'Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter1.Fill(DBDataSet1, "M_Referp")
        'DBTable1 = DBDataSet1.Tables("M_Referp")
        'For i = 0 To DBTable1.Rows.Count - 1
        'Dim ListItem1 As New ListItem
        'ListItem1.Text = DBTable1.Rows(i).Item("Data")
        'ListItem1.Value = DBTable1.Rows(i).Item("Data")
        'If ListItem1.Value = pName Then ListItem1.Selected = True
        'DLevel.Items.Add(ListItem1)
        'Next
        'End If
        'End If

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

                'Check客戶樣品圖-Size及格式
                If ErrCode = 0 Then
                    If DCustSampleFile.Visible Then
                        If DCustSampleFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DCustSampleFile)
                        End If
                    End If
                End If

                'Check最終樣品圖-Size及格式
                If ErrCode = 0 Then
                    If DFinalSampleFile.Visible Then
                        If DFinalSampleFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DFinalSampleFile)
                        End If
                    End If
                End If

                'Check品質依賴書-Size及格式
                If ErrCode = 0 Then
                    If DQCReqFile.Visible Then
                        If DQCReqFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DQCReqFile)
                        End If
                    End If
                End If

                'Check電鍍膜厚-Size及格式
                ' If ErrCode = 0 Then
                ' If DQCCheck6File.Visible Then
                ' If DQCCheck6File.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                ' ErrCode = UPFileIsNormal(DQCCheck6File)
                'End If
                'End If
                ' End If

                'Check Oeko-tex有害物質報告-Size及格式
                If ErrCode = 0 Then
                    If DEACheckFile.Visible Then
                        If DEACheckFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DEACheckFile)
                        End If
                    End If
                End If

                'Check A01有害物質報告-Size及格式
                If ErrCode = 0 Then
                    If DEACheckFile1.Visible Then
                        If DEACheckFile1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DEACheckFile1)
                        End If
                    End If
                End If

                'Check測試報告書-Size及格式
                If ErrCode = 0 Then
                    If DQCFinalFile.Visible Then
                        If DQCFinalFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DQCFinalFile)
                        End If
                    End If
                End If



                'Check FASFS-Size及格式
                If ErrCode = 0 Then
                    If DPFASFile.Visible Then
                        If DPFASFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DPFASFile)
                        End If
                    End If
                End If


                'Check製造流程表-Size及格式
                If ErrCode = 0 Then
                    If DManufFlowFile.Visible Then
                        If DManufFlowFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DManufFlowFile)
                        End If
                    End If
                End If

                'Check報價單-Size及格式
                If ErrCode = 0 Then
                    If DForcastFile.Visible Then
                        If DForcastFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DForcastFile)
                        End If
                    End If
                End If

                'Check作業標準書-Size及格式
                If ErrCode = 0 Then
                    If DOPManualFile.Visible Then
                        If DOPManualFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DOPManualFile)
                        End If
                    End If
                End If

                'Check切結書-Size及格式
                If ErrCode = 0 Then
                    If DContactFile.Visible Then
                        If DContactFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DContactFile)
                        End If
                    End If
                End If

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

        'Check客戶樣品圖-Size及格式
        If ErrCode = 0 Then
            If DCustSampleFile.Visible Then
                If DCustSampleFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DCustSampleFile)
                End If
            End If
        End If

        'Check最終樣品圖-Size及格式
        If ErrCode = 0 Then
            If DFinalSampleFile.Visible Then
                If DFinalSampleFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DFinalSampleFile)
                End If
            End If
        End If

        'Check品質依賴書-Size及格式
        If ErrCode = 0 Then
            If DQCReqFile.Visible Then
                If DQCReqFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DQCReqFile)
                End If
            End If
        End If

        'Check電鍍膜厚-Size及格式
        '  If ErrCode = 0 Then
        ' If DQCCheck6File.Visible Then
        ' If DQCCheck6File.PostedFile.FileName <> "" Then  '判斷有檔案上傳
        ' ErrCode = UPFileIsNormal(DQCCheck6File)
        'End If
        'End If
        '    End If

        'Check Oeko-tex有害物質報告-Size及格式
        If ErrCode = 0 Then
            If DEACheckFile.Visible Then
                If DEACheckFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DEACheckFile)
                End If
            End If
        End If

        'Check A01有害物質報告-Size及格式
        If ErrCode = 0 Then
            If DEACheckFile1.Visible Then
                If DEACheckFile1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DEACheckFile1)
                End If
            End If
        End If

        'Check測試報告書-Size及格式
        If ErrCode = 0 Then
            If DQCFinalFile.Visible Then
                If DQCFinalFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DQCFinalFile)
                End If
            End If
        End If


        'Check FAFS Size及格式
        If ErrCode = 0 Then
            If DPFASFile.Visible Then
                If DPFASFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DPFASFile)
                End If
            End If
        End If


        'Check製造流程表-Size及格式
        If ErrCode = 0 Then
            If DManufFlowFile.Visible Then
                If DManufFlowFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DManufFlowFile)
                End If
            End If
        End If

        'Check報價單-Size及格式
        If ErrCode = 0 Then
            If DForcastFile.Visible Then
                If DForcastFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DForcastFile)
                End If
            End If
        End If

        'Check作業標準書-Size及格式
        If ErrCode = 0 Then
            If DOPManualFile.Visible Then
                If DOPManualFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DOPManualFile)
                End If
            End If
        End If

        'Check切結書-Size及格式
        If ErrCode = 0 Then
            If DContactFile.Visible Then
                If DContactFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DContactFile)
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

        'Check-QC-L/T
        If ErrCode = 0 Then
            If DQCLT.Text <> "" Then
                If Not YKK.IntegerData(DQCLT.Text) Then
                    ErrCode = 9050
                Else
                    wQCLT = CInt(DQCLT.Text)
                End If
            End If
        End If

        '--檢查委託書No---------
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                ErrCode = oCommon.CommissionNo("000014", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
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
            If DAttachSample.SelectedValue = "YES" Then wLevel = "Z"


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
                    'Step=60 Action調整, 內製=不變更(0), 外注=變更(1)
                    If wStep = 60 Then
                        If DManufType.SelectedValue = "外注" Then pAction = 1
                    End If

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
                    If wFormSno = 0 And wStep < 3 Then '判斷是否起單
                        If NewFormSno <> 0 Then
                            AppendData(pFun, NewFormSno)  'Insert
                            AddCommissionNo(wFormNo, NewFormSno)
                            ModifyManufData(pFun, 1, 1)   '更新原表單狀態
                        End If
                    Else
                        If pNextStep = 999 Then     '工程結束嗎?
                            If pFun = "OK" Then ModifyData(pFun, CStr(wOKSts)) '更新表單資料
                            If pFun = "NG1" Then ModifyData(pFun, CStr(wNGSts1))
                            If pFun = "NG2" Then ModifyData(pFun, CStr(wNGSts2))
                            ModifyManufData(pFun, 0, 1)   '更新原表單狀態
                        Else
                            ModifyData(pFun, "0")         '更新表單資料 Sts=0(未結)
                        End If
                        AddCommissionNo(wFormNo, wFormSno)
                    End If
                    If wStep = 20 Then      '判斷是否為工程20
                        ModifyManufData(pFun, 1, 1)   '更新原表單狀態
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
            If ErrCode = 9050 Then Message = "QC L/T需為有效數字,請確認!"
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("SufaceFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String
        Dim RtnCode As Integer

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Insert into F_SufaceSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, No, "                    '1~5
        SQl = SQl + "Date, Division, Person, CustSampleFile, Spec, "               '6~10
        SQl = SQl + "Buyer, SellVendor, ReqDelDate, ReqQty, SliderSample, "        '11~15
        SQl = SQl + "AttachSample, ORNO, OrderTime, DevReason, FinalSampleFile, "  '16~20
        SQl = SQl + "ManufType, Suppiler, Color, Qty, AllowSample, "               '21~25
        SQl = SQl + "BFinalDate, Code, EnglishName, LOSS, QCReqFile, QCCheck1, "         '26~30
        SQl = SQl + "QCCheck2, QCCheck3, QCCheck4, QCCheck5, QCCheck15, QCCheck16,"       '31~35
        SQl = SQl + "QCCheck7, QCCheck8, QCCheck9, QCCheck10, QCCheck11,  "        '36~40
        SQl = SQl + "QCCheck12, QCCheck13, QCCheck14, EACheck1, EADesc1, EACheckFile, EACheckFile1, QCFinalFile, PFASFile,"     '41~47
        SQl = SQl + "QCDate1, QCResult1, QCRemark1, QCDate2, QCResult2, "          '48~52
        SQl = SQl + "QCRemark2, QCDate3, QCResult3, QCRemark3, ManufFlowFile, "    '53~57
        SQl = SQl + "ForcastFile, OPManualFile, ContactFile, Level, QCLT, "        '58~62
        SQl = SQl + "UpdSts, Suface, OFormNo, OFormSno, Price, YearSeason, "                   '63~67
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "              '68~71
        SQl = SQl + ")  "

        SQl = SQl + "VALUES( "
        '1~5
        SQl = SQl + " '0', "                                '狀態(0:未結,1:已結NG,2:已結OK)
        SQl = SQl + " '" + NowDateTime + "', "              '結案日
        SQl = SQl + " '000014', "                           '表單代號
        SQl = SQl + " '" + CStr(NewFormSno) + "', "         '表單流水號
        SQl = SQl + " '" + YKK.ReplaceString(DNo.Text) + "', "   'NO
        '6~10
        SQl = SQl + " N'" + DDate.Text + "', "                    '日期
        SQl = SQl + " N'" + DDivision.SelectedValue + "', "       '部門
        SQl = SQl + " N'" + DPerson.SelectedValue + "', "         '擔當
        FileName = ""
        If DCustSampleFile.Visible Then                          '客戶樣品圖
            If DCustSampleFile.PostedFile.FileName <> "" Then    '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "CustSample" & "-" & UploadDateTime & "-" & Right(DCustSampleFile.PostedFile.FileName, InStr(StrReverse(DCustSampleFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "CustSample" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DCustSampleFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DCustSampleFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "
        SQl = SQl + " N'" + YKK.ReplaceString(DSpec.Text) + "', "          '規格
        '11~15
        SQl = SQl + " N'" + DBuyer.SelectedValue + "', "                  'Buyer
        SQl = SQl + " N'" + YKK.ReplaceString(DSellVendor.Text) + "', "   '委託廠商
        SQl = SQl + " N'" + DReqDelDate.Text + "', "                       '希望交期
        SQl = SQl + " N'" + DReqQty.Text + "', "                           '預估量
        SQl = SQl + " N'" + DSliderSample.Text + "', "                    '樣品拉頭
        '16~20
        SQl = SQl + " N'" + DAttachSample.SelectedValue + "', "           '附樣
        SQl = SQl + " N'" + DORNO.Text + "', "                            'OR-NO
        SQl = SQl + " N'" + DOrderTime.Text + "', "                       '下單時間
        SQl = SQl + " N'" + DDevReason.Text + "', "                       '開發理由
        FileName = ""
        If DFinalSampleFile.Visible Then                                  '最終樣品圖
            If DFinalSampleFile.PostedFile.FileName <> "" Then    '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "FinalSample" & "-" & UploadDateTime & "-" & Right(DFinalSampleFile.PostedFile.FileName, InStr(StrReverse(DFinalSampleFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "FinalSample" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DFinalSampleFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DFinalSampleFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "
        '21~25
        SQl = SQl + " N'" + DManufType.SelectedValue + "', "              '內製/外注
        SQl = SQl + " N'" + DSuppiler.SelectedValue + "', "               '外注商
        SQl = SQl + " N'" + DColor.Text + "', "                           'Color
        SQl = SQl + " N'" + DQty.Text + "', "                             '數量
        SQl = SQl + " N'" + DAllowSample.SelectedValue + "', "            '限度樣品
        '26~30
        SQl = SQl + " N'" + DBFinalDate.Text + "', "                      '預定完成日
        SQl = SQl + " N'" + DCode.Text + "', "                            'Code
        SQl = SQl + " N'" + DEnglishName.Text + "', "                     '英文名稱
        SQl = SQl + " N'" + DLOSS.Text + "', "                            'LOSS
        FileName = ""
        If DQCReqFile.Visible Then                                        '品質依賴書
            If DQCReqFile.PostedFile.FileName <> "" Then    '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QCReq" & "-" & UploadDateTime & "-" & Right(DQCReqFile.PostedFile.FileName, InStr(StrReverse(DQCReqFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QCReq" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCReqFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQCReqFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "
        SQl = SQl + " N'" + DQCCheck1.SelectedValue + "', "               '口徑寸法
        '31~35
        SQl = SQl + " N'" + DQCCheck2.SelectedValue + "', "               '摺動抵抗
        SQl = SQl + " N'" + DQCCheck3.SelectedValue + "', "               'LOCK強度
        SQl = SQl + " N'" + DQCCheck4.SelectedValue + "', "               '90度強度
        SQl = SQl + " N'" + DQCCheck5.SelectedValue + "', "               '扭力
        SQl = SQl + " N'" + DQCCHECK15.SelectedValue + "', "              'N-ANTI
        SQl = SQl + " N'" + DQCCHECK16.SelectedValue + "', "              'PFAS

        'FileName = ""
        ' If DQCCheck6File.Visible Then                                     '電鍍膜厚
        ' If DQCCheck6File.PostedFile.FileName <> "" Then    '判斷有檔案上傳
        ' '*** IE8對應-Start 2011/1/4
        ' 'FileName = CStr(NewFormSno) & "-" & "QCCheck6" & "-" & UploadDateTime & "-" & Right(DQCCheck6File.PostedFile.FileName, InStr(StrReverse(DQCCheck6File.PostedFile.FileName), "\") - 1)
        ' FileName = CStr(NewFormSno) & "-" & "QCCheck6" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCCheck6File.PostedFile.FileName)
        ' '*** IE8對應-End
        ' Try    '上傳圖檔
        ' DQCCheck6File.PostedFile.SaveAs(Path + FileName)
        ' Catch ex As Exception
        ' End Try
        ' End If
        ' Else
        ' FileName = ""
        ' End If
        'SQl = SQl + " N'" + FileName + "', "
        '36~40
        SQl = SQl + " N'" + DQCCheck7.SelectedValue + "', "               '檢針
        SQl = SQl + " N'" + DQCCheck8.SelectedValue + "', "               'AATCC
        SQl = SQl + " N'" + DQCCheck9.SelectedValue + "', "               '乾洗
        SQl = SQl + " N'" + DQCCheck10.SelectedValue + "', "              '鹽水噴霧
        SQl = SQl + " N'" + DQCCheck11.SelectedValue + "', "              '一次密著
        '41~47
        SQl = SQl + " N'" + DQCCheck12.SelectedValue + "', "              '二次密著
        SQl = SQl + " N'" + DQCCheck13.SelectedValue + "', "              'Oeko-tex
        SQl = SQl + " N'" + DQCCheck14.SelectedValue + "', "              'A01
        SQl = SQl + " N'" + DEACheck1.SelectedValue + "', "               '有害物質
        SQl = SQl + " N'" + DEADesc1.Text + "', "                         '有害物質備註

        FileName = ""
        If DEACheckFile.Visible Then                                      'Oeko-tex有害物質報告
            If DEACheckFile.PostedFile.FileName <> "" Then    '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "EACheck" & "-" & UploadDateTime & "-" & Right(DEACheckFile.PostedFile.FileName, InStr(StrReverse(DEACheckFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "EACheck" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DEACheckFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DEACheckFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        FileName = ""
        If DEACheckFile1.Visible Then                                      'A01有害物質報告
            If DEACheckFile1.PostedFile.FileName <> "" Then    '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "EACheck1" & "-" & UploadDateTime & "-" & Right(DEACheckFile1.PostedFile.FileName, InStr(StrReverse(DEACheckFile1.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "EACheck1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DEACheckFile1.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DEACheckFile1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        FileName = ""
        If DQCFinalFile.Visible Then                                      '測試報告書
            If DQCFinalFile.PostedFile.FileName <> "" Then    '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QCFinal" & "-" & UploadDateTime & "-" & Right(DQCFinalFile.PostedFile.FileName, InStr(StrReverse(DQCFinalFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QCFinal" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFinalFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQCFinalFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "
        FileName = ""
        If DPFASFile.Visible Then                                      '測試報告書
            If DPFASFile.PostedFile.FileName <> "" Then    '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QCFinal" & "-" & UploadDateTime & "-" & Right(DPFASFile.PostedFile.FileName, InStr(StrReverse(DPFASFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "FAFS" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DPFASFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DPFASFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        '48~52
        SQl = SQl + " N'" + DQCDate1.Text + "', "                         '日期-1
        SQl = SQl + " N'" + DQCResult1.SelectedValue + "', "              '檢測結果-1
        SQl = SQl + " N'" + DQCRemark1.Text + "', "                       '備註-1
        SQl = SQl + " N'" + DQCDate2.Text + "', "                         '日期-2
        SQl = SQl + " N'" + DQCResult2.SelectedValue + "', "              '檢測結果-2
        '53~57
        SQl = SQl + " N'" + DQCRemark2.Text + "', "                       '備註-2
        SQl = SQl + " N'" + DQCDate3.Text + "', "                         '日期-3
        SQl = SQl + " N'" + DQCResult3.SelectedValue + "', "              '檢測結果-3
        SQl = SQl + " N'" + DQCRemark3.Text + "', "                       '備註-3
        FileName = ""
        If DManufFlowFile.Visible Then                                      '製造流程表
            If DManufFlowFile.PostedFile.FileName <> "" Then    '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "ManufFlow" & "-" & UploadDateTime & "-" & Right(DManufFlowFile.PostedFile.FileName, InStr(StrReverse(DManufFlowFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "ManufFlow" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DManufFlowFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DManufFlowFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "
        '58~62
        FileName = ""
        If DForcastFile.Visible Then                                      '報價單
            If DForcastFile.PostedFile.FileName <> "" Then    '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Forcast" & "-" & UploadDateTime & "-" & Right(DForcastFile.PostedFile.FileName, InStr(StrReverse(DForcastFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "Forcast" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DForcastFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DForcastFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "
        FileName = ""
        If DOPManualFile.Visible Then                                      '作業標準書
            If DOPManualFile.PostedFile.FileName <> "" Then    '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "OPManual" & "-" & UploadDateTime & "-" & Right(DOPManualFile.PostedFile.FileName, InStr(StrReverse(DOPManualFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "OPManual" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DOPManualFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DOPManualFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "
        FileName = ""
        If DContactFile.Visible Then                                      '切結書
            If DContactFile.PostedFile.FileName <> "" Then    '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Contact" & "-" & UploadDateTime & "-" & Right(DContactFile.PostedFile.FileName, InStr(StrReverse(DContactFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "Contact" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DContactFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DContactFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "
        SQl = SQl + " '" + "" + "', "
        SQl = SQl + " '" + DQCLT.Text + "', "            'QC-L/T
        '63~67
        SQl = SQl + " '" + "0" + "', "
        SQl = SQl + " '" + "0" + "', "
        SQl = SQl + " '" + DOFormNo.Text + "', "            '表單No
        If DOFormSno.Text = "" Then                         '單號
            SQl = SQl + " '0', "
        Else
            SQl = SQl + " '" + DOFormSno.Text + "', "
        End If

        SQl = SQl + " N'" + DPrice.Text + "', "             '價格
        SQl = SQl + " N'" + DYearSeason.SelectedValue + "', "  '年季
        '--------------------------------------------
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("SufaceFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String
        Dim RtnCode As Integer

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Update F_SufaceSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = '" & YKK.ReplaceString(DNo.Text) & "',"     'NO
        SQl = SQl + " Date = N'" & DDate.Text & "',"                   '日期
        '20140714
        'SQl = SQl + " Division = N'" & DDivision.SelectedValue & "',"  '部門
        'SQl = SQl + " Person = N'" & DPerson.SelectedValue & "',"      '擔當
        If DCustSampleFile.Visible Then                                '客戶樣品圖
            If DCustSampleFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "CustSample" & "-" & UploadDateTime & "-" & Right(DCustSampleFile.PostedFile.FileName, InStr(StrReverse(DCustSampleFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "CustSample" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DCustSampleFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DCustSampleFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " CustSampleFile = N'" & FileName & "',"
            End If
        End If
        SQl = SQl + " Spec = N'" & YKK.ReplaceString(DSpec.Text) & "',"              '規格
        SQl = SQl + " Buyer = N'" & DBuyer.SelectedValue & "',"                      'Buyer
        SQl = SQl + " SellVendor = N'" & YKK.ReplaceString(DSellVendor.Text) & "',"  '委託廠商
        SQl = SQl + " ReqDelDate = N'" & DReqDelDate.Text & "',"                     '希望交期
        SQl = SQl + " ReqQty = N'" & DReqQty.Text & "',"                             '預估量
        SQl = SQl + " SliderSample = N'" & DSliderSample.Text & "',"                 '樣品拉頭
        SQl = SQl + " AttachSample = N'" & DAttachSample.SelectedValue & "',"        '附樣
        SQl = SQl + " ORNO = N'" & DORNO.Text & "',"                                 'OR-NO
        SQl = SQl + " OrderTime = N'" & DOrderTime.Text & "',"                       '下單時間
        SQl = SQl + " Price = N'" & DPrice.Text & "',"                               '價格
        SQl = SQl + " DevReason = N'" & DDevReason.Text & "',"                       '開發理由
        If DFinalSampleFile.Visible Then                                             '最終樣品圖
            If DFinalSampleFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "FinalSample" & "-" & UploadDateTime & "-" & Right(DFinalSampleFile.PostedFile.FileName, InStr(StrReverse(DFinalSampleFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "FinalSample" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DFinalSampleFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DFinalSampleFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " FinalSampleFile = N'" & FileName & "',"
            End If
        End If
        SQl = SQl + " ManufType = N'" & DManufType.SelectedValue & "',"        '內製/外注
        SQl = SQl + " Suppiler = N'" & DSuppiler.SelectedValue & "',"          '外注商
        SQl = SQl + " Color = N'" & DColor.Text & "',"                         'Color
        SQl = SQl + " Qty = N'" & DQty.Text & "',"                             '數量
        SQl = SQl + " Cap = N'" & DCap.Text & "',"                             '日產能
        SQl = SQl + " Schedule = N'" & DSchedule.Text & "',"                   '基礎日程
        SQl = SQl + " FReason = N'" & DFReason.Text & "',"                     '理由
        SQl = SQl + " BFinalDate = N'" & DBFinalDate.Text & "',"               '預定完成日
        SQl = SQl + " Code = N'" & DCode.Text & "',"                           'Code
        SQl = SQl + " EnglishName = N'" & DEnglishName.Text & "',"             '英文名稱
        SQl = SQl + " LOSS = N'" & DLOSS.Text & "',"                           'LOSS
        If DQCReqFile.Visible Then                                             '品質依賴書
            If DQCReqFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QCReq" & "-" & UploadDateTime & "-" & Right(DQCReqFile.PostedFile.FileName, InStr(StrReverse(DQCReqFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QCReq" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCReqFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQCReqFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QCReqFile = N'" & FileName & "',"
            End If
        End If
        SQl = SQl + " QCCheck1 = N'" & DQCCheck1.SelectedValue & "',"    '口徑寸法
        SQl = SQl + " QCCheck2 = N'" & DQCCheck2.SelectedValue & "',"    '摺動抵抗
        SQl = SQl + " QCCheck3 = N'" & DQCCheck3.SelectedValue & "',"    'LOCK強度
        SQl = SQl + " QCCheck4 = N'" & DQCCheck4.SelectedValue & "',"    '90度強度
        SQl = SQl + " QCCheck5 = N'" & DQCCheck5.SelectedValue & "',"    '扭力
        SQl = SQl + " QCCheck15 = N'" & DQCCHECK15.SelectedValue & "',"    'N-ANTI
        SQl = SQl + " QCCheck16  = N'" & DQCCHECK16.SelectedValue & "',"    'PFAS

        ' If DQCCheck6File.Visible Then                                       '電鍍膜厚
        ' If DQCCheck6File.PostedFile.FileName <> "" Then  '判斷有檔案上傳
        ' '*** IE8對應-Start 2011/1/4
        ' 'FileName = CStr(wFormSno) & "-" & "QCCheck6" & "-" & UploadDateTime & "-" & Right(DQCCheck6File.PostedFile.FileName, InStr(StrReverse(DQCCheck6File.PostedFile.FileName), "\") - 1)
        ' FileName = CStr(wFormSno) & "-" & "QCCheck6" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCCheck6File.PostedFile.FileName)
        ' '*** IE8對應-End
        ' Try    '上傳圖檔
        ' DQCCheck6File.PostedFile.SaveAs(Path + FileName)
        ' Catch ex As Exception
        ' End Try
        ' SQl = SQl + " QCCheck6File = N'" & FileName & "',"
        ' End If
        ' End If

        SQl = SQl + " QCCheck7 = N'" & DQCCheck7.SelectedValue & "',"    '檢針
        SQl = SQl + " QCCheck8 = N'" & DQCCheck8.SelectedValue & "',"    'AATCC
        SQl = SQl + " QCCheck9 = N'" & DQCCheck9.SelectedValue & "',"    '乾洗
        SQl = SQl + " QCCheck10 = N'" & DQCCheck10.SelectedValue & "',"  '鹽水噴霧
        SQl = SQl + " QCCheck11 = N'" & DQCCheck11.SelectedValue & "',"  '一次密著
        SQl = SQl + " QCCheck12 = N'" & DQCCheck12.SelectedValue & "',"  '二次密著
        SQl = SQl + " QCCheck13 = N'" & DQCCheck13.SelectedValue & "',"  'Oeko-tex
        SQl = SQl + " QCCheck14 = N'" & DQCCheck14.SelectedValue & "',"  'A01
        SQl = SQl + " EACheck1 = N'" & DEACheck1.SelectedValue & "',"    '有害物質
        SQl = SQl + " EADesc1 = N'" & DEADesc1.Text & "',"               '有害物質備註

        If DEACheckFile.Visible Then                                     'Oeko-tex有害物質報告
            If DEACheckFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "EACheck" & "-" & UploadDateTime & "-" & Right(DEACheckFile.PostedFile.FileName, InStr(StrReverse(DEACheckFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "EACheck" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DEACheckFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DEACheckFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " EACheckFile = N'" & FileName & "',"
            End If
        End If

        If DEACheckFile1.Visible Then                                     'A01有害物質報告
            If DEACheckFile1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "EACheck1" & "-" & UploadDateTime & "-" & Right(DEACheckFile1.PostedFile.FileName, InStr(StrReverse(DEACheckFile1.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "EACheck1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DEACheckFile1.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DEACheckFile1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " EACheckFile1 = N'" & FileName & "',"
            End If
        End If

        If DQCFinalFile.Visible Then                                     '有害物質報告
            If DQCFinalFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QCFinal" & "-" & UploadDateTime & "-" & Right(DQCFinalFile.PostedFile.FileName, InStr(StrReverse(DQCFinalFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QCFinal" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFinalFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQCFinalFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QCFinalFile = N'" & FileName & "',"
            End If
        End If

        If DPFASFile.Visible Then                                     '有害物質報告
            If DPFASFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QCFinal" & "-" & UploadDateTime & "-" & Right(DPFASFile.PostedFile.FileName, InStr(StrReverse(DPFASFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "FAFS" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DPFASFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DPFASFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " PFASFile = N'" & FileName & "',"
            End If
        End If


        SQl = SQl + " QCDate1 = N'" & DQCDate1.Text & "',"               '日期-1
        SQl = SQl + " QCResult1 = N'" & DQCResult1.SelectedValue & "',"  '檢測結果-1
        SQl = SQl + " QCRemark1 = N'" & DQCRemark1.Text & "',"           '備註-1
        SQl = SQl + " QCDate2 = N'" & DQCDate2.Text & "',"               '日期-2
        SQl = SQl + " QCResult2 = N'" & DQCResult2.SelectedValue & "',"  '檢測結果-2
        SQl = SQl + " QCRemark2 = N'" & DQCRemark2.Text & "',"           '備註-2
        SQl = SQl + " QCDate3 = N'" & DQCDate3.Text & "',"               '日期-3
        SQl = SQl + " QCResult3 = N'" & DQCResult3.SelectedValue & "',"  '檢測結果-3
        SQl = SQl + " QCRemark3 = N'" & DQCRemark3.Text & "',"           '備註-3
        If DManufFlowFile.Visible Then                                   '製造流程表
            If DManufFlowFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "ManufFlow" & "-" & UploadDateTime & "-" & Right(DManufFlowFile.PostedFile.FileName, InStr(StrReverse(DManufFlowFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "ManufFlow" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DManufFlowFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DManufFlowFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " ManufFlowFile = N'" & FileName & "',"
            End If
        End If
        If DForcastFile.Visible Then                                   '製造流程表
            If DForcastFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Forcast" & "-" & UploadDateTime & "-" & Right(DForcastFile.PostedFile.FileName, InStr(StrReverse(DForcastFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Forcast" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DForcastFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DForcastFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " ForcastFile = N'" & FileName & "',"
            End If
        End If
        If DOPManualFile.Visible Then                                   '製造流程表
            If DOPManualFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "OPManual" & "-" & UploadDateTime & "-" & Right(DOPManualFile.PostedFile.FileName, InStr(StrReverse(DOPManualFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "OPManual" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DOPManualFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DOPManualFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " OPManualFile = N'" & FileName & "',"
            End If
        End If
        If DContactFile.Visible Then                                   '製造流程表
            If DContactFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Contact" & "-" & UploadDateTime & "-" & Right(DContactFile.PostedFile.FileName, InStr(StrReverse(DContactFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Contact" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DContactFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DContactFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " ContactFile = N'" & FileName & "',"
            End If
        End If

        SQl = SQl + " OFormNo = '" & DOFormNo.Text & "',"
        SQl = SQl + " OFormSno = '" & DOFormSno.Text & "',"
        SQl = SQl + " QCLT = '" & DQCLT.Text & "',"  'QC-L/T
        SQl = SQl + " YearSeason = '" & DYearSeason.SelectedValue & "',"  'QC-L/T

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
    '**     檢查上傳檔案
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As HtmlInputFile) As Integer
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
                UPFileIsNormal = 9010
            End If
        Next
        'Check上傳檔案Size
        If UPFileIsNormal = 0 Then
            If UPFile.PostedFile.ContentLength <= 2500 * 1024 Then
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
    Private Sub BPrint_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BPrint.Click
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
        Dim wLevel As String = ""   '難易度

        URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                              "&pFormSno=" & wFormSno & _
                              "&pStep=" & wStep & _
                              "&pUserID=" & Request.QueryString("pUserID") & _
                              "&pLevel=" & wLevel

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
    '**(ModifyManufData)
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub ModifyManufData(ByVal pFun As String, ByVal pSts As Integer, ByVal pSurface As Integer)
        If DOFormNo.Text = "000003" Or DOFormNo.Text = "000007" Or DOFormNo.Text = "000011" Then    ''判斷是否有選取NO
            Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                                 CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
            Dim SQl As String

            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
            Dim OleDBCommand1 As New OleDbCommand

            If DOFormNo.Text = "000003" Then
                SQl = "Update F_ManufInSheet Set "
            End If
            If DOFormNo.Text = "000007" Then
                SQl = "Update F_ManufOutSheet Set "
            End If
            If DOFormNo.Text = "000011" Then
                SQl = "Update F_ImportSheet Set "
            End If

            SQl = SQl + " SFSts = '" & CStr(pSts) & "',"
            SQl = SQl + " Surface = '" & CStr(pSurface) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & DOFormNo.Text & "'"
            SQl = SQl + "   And FormSno =  '" & DOFormSno.Text & "'"
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQl
            OleDbConnection1.Open()
            OleDBCommand1.ExecuteNonQuery()
            OleDbConnection1.Close()
        End If
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
        'Spec
        If InputCheck = 0 Then
            If FindFieldInf("Spec") = 1 Then
                If DSpec.Text = "" Then InputCheck = 1
            End If
        End If
        'Buyer
        If InputCheck = 0 Then
            If FindFieldInf("Buyer") = 1 Then
                If DBuyer.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '委託廠商
        If InputCheck = 0 Then
            If FindFieldInf("SellVendor") = 1 Then
                If DSellVendor.Text = "" Then InputCheck = 1
            End If
        End If
        '希望交期
        If InputCheck = 0 Then
            If FindFieldInf("ReqDelDate") = 1 Then
                If DReqDelDate.Text = "" Then InputCheck = 1
            End If
        End If
        '預估量
        If InputCheck = 0 Then
            If FindFieldInf("ReqQty") = 1 Then
                If DReqQty.Text = "" Then InputCheck = 1
            End If
        End If
        '樣品拉頭
        If InputCheck = 0 Then
            If FindFieldInf("SliderSample") = 1 Then
                If DSliderSample.Text = "" Then InputCheck = 1
            End If
        End If
        '附樣
        If InputCheck = 0 Then
            If FindFieldInf("AttachSample") = 1 Then
                If DAttachSample.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'OR-NO
        If InputCheck = 0 Then
            If FindFieldInf("ORNO") = 1 Then
                If DORNO.Text = "" Then InputCheck = 1
            End If
        End If
        '下單時間
        If InputCheck = 0 Then
            If FindFieldInf("OrderTime") = 1 Then
                If DOrderTime.Text = "" Then InputCheck = 1
            End If
        End If
        '價格
        If InputCheck = 0 Then
            If FindFieldInf("Price") = 1 Then
                If DPrice.Text = "" Then InputCheck = 1
            End If
        End If
        '開發理由
        If InputCheck = 0 Then
            If FindFieldInf("DevReason") = 1 Then
                If DDevReason.Text = "" Then InputCheck = 1
            End If
        End If
        '最終樣品圖
        If InputCheck = 0 Then
            If FindFieldInf("FinalSampleFile") = 1 Then
                If DFinalSampleFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '內製/外注
        If InputCheck = 0 Then
            If FindFieldInf("ManufType") = 1 Then
                If DManufType.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '外注商
        If InputCheck = 0 Then
            If FindFieldInf("Suppiler") = 1 Then
                If DSuppiler.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '顏色
        If InputCheck = 0 Then
            If FindFieldInf("Color") = 1 Then
                If DColor.Text = "" Then InputCheck = 1
            End If
        End If
        '數量
        If InputCheck = 0 Then
            If FindFieldInf("Qty") = 1 Then
                If DQty.Text = "" Then InputCheck = 1
            End If
        End If
        'QC-L/T
        If InputCheck = 0 Then
            If FindFieldInf("QCLT") = 1 Then
                If DQCLT.Text = "" Then InputCheck = 1
            End If
        End If
        '日產能
        If InputCheck = 0 Then
            If FindFieldInf("Cap") = 1 Then
                If DCap.Text = "" Then InputCheck = 1
            End If
        End If
        '日程
        If InputCheck = 0 Then
            If FindFieldInf("Schedule") = 1 Then
                If DSchedule.Text = "" Then InputCheck = 1
            End If
        End If
        '日程
        If InputCheck = 0 Then
            If FindFieldInf("FReason") = 1 Then
                If DFReason.Text = "" Then InputCheck = 1
            End If
        End If
        '限度樣品
        If InputCheck = 0 Then
            If FindFieldInf("AllowSample") = 1 Then
                If DAllowSample.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '預定完成日
        If InputCheck = 0 Then
            If FindFieldInf("BFinalDate") = 1 Then
                If DBFinalDate.Text = "" Then InputCheck = 1
            End If
        End If
        'Code
        If InputCheck = 0 Then
            If FindFieldInf("Code") = 1 Then
                If DCode.Text = "" Then InputCheck = 1
            End If
        End If
        '英文名稱
        If InputCheck = 0 Then
            If FindFieldInf("EnglishName") = 1 Then
                If DEnglishName.Text = "" Then InputCheck = 1
            End If
        End If
        'LOSS
        If InputCheck = 0 Then
            If FindFieldInf("LOSS") = 1 Then
                If DLOSS.Text = "" Then InputCheck = 1
            End If
        End If
        '品質依賴書
        If InputCheck = 0 Then
            If FindFieldInf("QCReqFile") = 1 Then
                If DQCReqFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '口徑寸法
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck1") = 1 Then
                If DQCCheck1.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '摺動抵抗
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck2") = 1 Then
                If DQCCheck2.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'LOCK強度
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck3") = 1 Then
                If DQCCheck3.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '90度強度
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck4") = 1 Then
                If DQCCheck4.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '扭力
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck5") = 1 Then
                If DQCCheck5.SelectedValue = "" Then InputCheck = 1
            End If
        End If

        'N-ANTI
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck15") = 1 Then
                If DQCCHECK15.SelectedValue = "" Then InputCheck = 1
            End If
        End If

        'N-ANTI
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck16") = 1 Then
                If DQCCHECK16.SelectedValue = "" Then InputCheck = 1
            End If
        End If



        '電鍍膜厚
        ' If InputCheck = 0 Then
        'If FindFieldInf("QCCheck6File") = 1 Then
        'If DQCCheck6File.PostedFile.FileName = "" Then InputCheck = 1
        'End If
        'End If

        '檢針
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck7") = 1 Then
                If DQCCheck7.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'AATCC
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck8") = 1 Then
                If DQCCheck8.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '乾洗
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck9") = 1 Then
                If DQCCheck9.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '鹽水噴霧
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck10") = 1 Then
                If DQCCheck10.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '一次密著
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck11") = 1 Then
                If DQCCheck11.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '二次密著,Oeko-tex,A01
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck12") = 1 Then
                If DQCCheck12.SelectedValue = "" Then InputCheck = 1
                If DQCCheck13.SelectedValue = "" Then InputCheck = 1
                If DQCCheck14.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '有害物質
        If InputCheck = 0 Then
            If FindFieldInf("EACheck1") = 1 Then
                If DEACheck1.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '有害物質備註
        If InputCheck = 0 Then
            If FindFieldInf("EADesc1") = 1 Then
                If DEADesc1.Text = "" Then InputCheck = 1
            End If
        End If
        'Oeko-tex,A01有害物質報告
        If InputCheck = 0 Then
            If FindFieldInf("EACheckFile") = 1 Then
                If DEACheckFile.PostedFile.FileName = "" Then InputCheck = 1
                If DEACheckFile1.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '測試報告書
        If InputCheck = 0 Then
            If FindFieldInf("QCFinalFile") = 1 Then
                If DQCFinalFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If

        'FAFS
        If InputCheck = 0 Then
            If FindFieldInf("QCFinalFile") = 1 Then
                If DPFASFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If


        '日期-1
        If InputCheck = 0 Then
            If FindFieldInf("QCDate1") = 1 Then
                If DQCDate1.Text = "" Then InputCheck = 1
            End If
        End If
        '檢測結果-1
        If InputCheck = 0 Then
            If FindFieldInf("QCResult1") = 1 Then
                If DQCResult1.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '備註-1
        If InputCheck = 0 Then
            If FindFieldInf("QCRemark1") = 1 Then
                If DQCRemark1.Text = "" Then InputCheck = 1
            End If
        End If
        '日期-2
        If InputCheck = 0 Then
            If FindFieldInf("QCDate2") = 1 Then
                If DQCDate2.Text = "" Then InputCheck = 1
                If DQCDate3.Text = "" Then InputCheck = 1
            End If
        End If
        '檢測結果-2
        If InputCheck = 0 Then
            If FindFieldInf("QCResult2") = 1 Then
                If DQCResult2.SelectedValue = "" Then InputCheck = 1
                If DQCResult3.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '備註-2 備註-3
        If InputCheck = 0 Then
            If FindFieldInf("QCRemark2") = 1 Then
                If DQCRemark2.Text = "" Then InputCheck = 1
                If DQCRemark3.Text = "" Then InputCheck = 1
            End If
        End If
        '製造流程表
        If InputCheck = 0 Then
            If FindFieldInf("ManufFlowFile") = 1 Then
                If DManufFlowFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '報價單
        If InputCheck = 0 Then
            If FindFieldInf("ForcastFile") = 1 Then
                If DForcastFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '作業標準書
        If InputCheck = 0 Then
            If FindFieldInf("OPManualFile") = 1 Then
                If DOPManualFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '切結書
        If InputCheck = 0 Then
            If FindFieldInf("ContactFile") = 1 Then
                If DContactFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '客戶樣品圖
        If InputCheck = 0 Then
            If FindFieldInf("CustSampleFile") = 1 Then
                If DCustSampleFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
    End Function

    Private Sub DAttachSample_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DAttachSample.SelectedIndexChanged

    End Sub
End Class
