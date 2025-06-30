Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ManufInBefSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DQAD13 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DCpsc As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMakeCAM As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC13 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB13 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD12 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC12 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB12 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD11 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC11 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB11 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA13 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA12 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA11 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD14 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD10 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD9 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD8 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD6 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAF4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAD5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAD4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAD3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAD1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAC14 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC10 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC9 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC8 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC6 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAF3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAC5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAC4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAC3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAC1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAB14 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB10 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB9 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAF2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAB8 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB6 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAB4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAB3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAB1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAA14 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA10 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA9 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA8 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA6 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAA4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAA3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAA2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAF1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DManuaInSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DSuppiler As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Label25 As System.Web.UI.WebControls.Label
    Protected WithEvents Label24 As System.Web.UI.WebControls.Label
    Protected WithEvents Label23 As System.Web.UI.WebControls.Label
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label22 As System.Web.UI.WebControls.Label
    Protected WithEvents Label21 As System.Web.UI.WebControls.Label
    Protected WithEvents Label20 As System.Web.UI.WebControls.Label
    Protected WithEvents Label19 As System.Web.UI.WebControls.Label
    Protected WithEvents Label18 As System.Web.UI.WebControls.Label
    Protected WithEvents Label17 As System.Web.UI.WebControls.Label
    Protected WithEvents Label16 As System.Web.UI.WebControls.Label
    Protected WithEvents Label15 As System.Web.UI.WebControls.Label
    Protected WithEvents Label14 As System.Web.UI.WebControls.Label
    Protected WithEvents Label13 As System.Web.UI.WebControls.Label
    Protected WithEvents Label26 As System.Web.UI.WebControls.Label
    Protected WithEvents Label12 As System.Web.UI.WebControls.Label
    Protected WithEvents Label11 As System.Web.UI.WebControls.Label
    Protected WithEvents Label10 As System.Web.UI.WebControls.Label
    Protected WithEvents Label9 As System.Web.UI.WebControls.Label
    Protected WithEvents Label8 As System.Web.UI.WebControls.Label
    Protected WithEvents Label7 As System.Web.UI.WebControls.Label
    Protected WithEvents Label6 As System.Web.UI.WebControls.Label
    Protected WithEvents Label5 As System.Web.UI.WebControls.Label
    Protected WithEvents Label4 As System.Web.UI.WebControls.Label
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents DBuyer As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LQAttachFile1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQAttachFile2 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LSAttachFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DFQAD3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFQAD2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFQAD1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFQAC1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFQAC2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFQAC3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFQAB3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFQAB2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFQAB1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents LMapNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DFQAA3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFQAA2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFQAA1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceJ1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceI3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceI2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceI1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceH3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceH2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceJ3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceJ2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceA3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceA2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceH1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceG3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceG2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceG1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceF3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceF2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceF1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceE3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceE2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceE1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceD3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceD2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceD1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceC3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceC2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceC1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceB3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceB2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceB1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleC3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleC2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleC1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleB3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleB2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleB1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceA1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleA1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSliderCode As System.Web.UI.WebControls.Button
    Protected WithEvents LContactFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LRefFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQAFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DSampleA2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleA3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents BMMapNo As System.Web.UI.WebControls.Button
    Protected WithEvents LSampleFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMoldPoint As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMoldQty As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPullerPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPurMold As System.Web.UI.WebControls.TextBox
    Protected WithEvents DArMoldFee As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDevReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents LAuthorizeFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LConfirmFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMaterial As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DManufPlace As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSliderType2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSliderType1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAssembler As System.Web.UI.WebControls.TextBox
    Protected WithEvents DLevel As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BOMapNo As System.Web.UI.WebControls.Button
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSpec As System.Web.UI.WebControls.Button
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderGRCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DManuaInSheet2 As System.Web.UI.WebControls.Image
    Protected WithEvents LMapFile As System.Web.UI.WebControls.Image
    Protected WithEvents DQAttachFile2 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQAttachFile1 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DSAttachFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DMapFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DContactFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQAFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DSampleFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DAuthorizeFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DConfirmFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DRefFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents DOPReadyDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOPReady As System.Web.UI.WebControls.TextBox
    Protected WithEvents BFlow As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LBefOP As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DAEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDelivery As System.Web.UI.WebControls.Image
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DDescSheet As System.Web.UI.WebControls.Image
    Protected WithEvents BPrint As System.Web.UI.WebControls.ImageButton
    Protected WithEvents Lable27 As System.Web.UI.WebControls.Label

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

        Response.Cookies("MapNo").Value = ""        '圖號, MapPicker使用

        Response.Cookies("PGM").Value = "ManufInSheet_01.aspx"
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
        'Check圖樣
        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then
                Message = "圖樣"
            End If
        End If
        'Check樣品-其他附件
        If DSAttachFile.Visible Then
            If DSAttachFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "樣品-其他附件"
                Else
                    Message = Message & ", " & "樣品-其他附件"
                End If
            End If
        End If
        'Check品質１-其他附件
        If DQAttachFile1.Visible Then
            If DQAttachFile1.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "品質１-其他附件"
                Else
                    Message = Message & ", " & "品質１-其他附件"
                End If
            End If
        End If
        'Check品質２-其他附件
        If DQAttachFile2.Visible Then
            If DQAttachFile2.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "品質２-其他附件"
                Else
                    Message = Message & ", " & "品質２-其他附件"
                End If
            End If
        End If
        'Check其他附件
        If DRefFile.Visible Then
            If DRefFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "其他附件"
                Else
                    Message = Message & ", " & "其他附件"
                End If
            End If
        End If
        'Check確認書
        If DConfirmFile.Visible Then
            If DConfirmFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "確認書"
                Else
                    Message = Message & ", " & "確認書"
                End If
            End If
        End If
        'Check授權書
        If DAuthorizeFile.Visible Then
            If DAuthorizeFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "授權書"
                Else
                    Message = Message & ", " & "授權書"
                End If
            End If
        End If
        'Check樣品圖
        If DSampleFile.Visible Then
            If DSampleFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "樣品圖"
                Else
                    Message = Message & ", " & "樣品圖"
                End If
            End If
        End If
        'Check測試報告
        If DQAFile.Visible Then
            If DQAFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "測試報告"
                Else
                    Message = Message & ", " & "測試報告"
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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ManufInFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_ManufInSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ManufInSheet")
        If DBDataSet1.Tables("F_ManufInSheet").Rows.Count > 0 Then
            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MapFile") <> "" Then          '圖樣
                LMapFile.ImageUrl = Path & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MapFile")
            Else
                LMapFile.Visible = False
            End If

            DNo.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Date")                   '日期
            SetFieldData("Division", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Division"))  '部門
            SetFieldData("Person", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Person"))      '擔當
            SetFieldData("Cpsc", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Cpsc"))          'CPSC
            DSliderCode.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SliderCode")       'Slider Code
            DSliderGRCode.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SliderGRCode")   'Slider Group Code
            DSpec.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Spec")                   '規格
            DMapNo.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MapNo")                 '圖號
            DOFormNo.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("OFormNo")             '原表單
            DOFormSno.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("OFormSno")           '原單號

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MapNo") <> "" Then                 '圖號
                LMapNo.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MapNo")
                If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("OFormNo") = "000001" Then
                    LMapNo.NavigateUrl = "MapSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("OFormNo") & _
                                                          "&pFormSno=" & CStr(DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("OFormSno"))
                Else
                    LMapNo.NavigateUrl = "MapModSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("OFormNo") & _
                                                          "&pFormSno=" & CStr(DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("OFormSno"))
                End If
            Else
                LMapNo.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SAttachFile") <> "" Then       '樣品-其他附件
                LSAttachFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SAttachFile")
            Else
                LSAttachFile.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("QAttachFile1") <> "" Then      '品質1-其他附件
                LQAttachFile1.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("QAttachFile1")
            Else
                LQAttachFile1.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("QAttachFile2") <> "" Then       '樣品2-其他附件
                LQAttachFile2.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("QAttachFile2")
            Else
                LQAttachFile2.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("RefFile") <> "" Then          '其他附件 
                LRefFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("RefFile")
            Else
                LRefFile.Visible = False
            End If

            SetFieldData("Level", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Level"))                '難易度
            DAssembler.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Assembler")                 '組立
            SetFieldData("SliderType1", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SliderType1"))    '拉頭種類1
            SetFieldData("SliderType2", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SliderType2"))    '拉頭種類2
            SetFieldData("ManufPlace", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("ManufPlace"))      '生產地
            SetFieldData("Material", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Material"))              '材質
            SetFieldData("Buyer", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Buyer"))    'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SellVendor")   '委託廠商

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("ConfirmFile") <> "" Then       '確認書
                LConfirmFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("ConfirmFile")
            Else
                LConfirmFile.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("AuthorizeFile") <> "" Then     '授權書
                LAuthorizeFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("AuthorizeFile")
            Else
                LAuthorizeFile.Visible = False
            End If

            DDevReason.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("DevReason")     '開發理由

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Sample") = 1 Then              '樣品
                Dim DBTable1 As DataTable
                SQL = "Select * From F_ManufCOSample "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "ManufCOSample")
                DBTable1 = DBDataSet1.Tables("ManufCOSample")
                For i = 0 To DBTable1.Rows.Count - 1
                    Select Case DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Seqno")
                        Case 1
                            DSampleA1.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Spec")   'Spec
                            DSampleA2.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Color")  'Color
                            DSampleA3.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Qty")    'Qty
                        Case 2
                            DSampleB1.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Spec")   'Spec
                            DSampleB2.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Color")  'Color
                            DSampleB3.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Qty")    'Qty
                        Case Else
                            DSampleC1.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Spec")   'Spec
                            DSampleC2.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Color")  'Color
                            DSampleC3.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Qty")    'Qty
                    End Select
                Next
            End If

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SampleFile") <> "" Then        '樣品檔
                LSampleFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SampleFile")
            Else
                LSampleFile.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Price") = 1 Then               '單價
                Dim DBTable1 As DataTable
                SQL = "Select * From F_ManufCOPrice "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "ManufCOPrice")
                DBTable1 = DBDataSet1.Tables("ManufCOPrice")
                For i = 0 To DBTable1.Rows.Count - 1
                    Select Case DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Seqno")
                        Case 1
                            DPriceA1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceA2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceA3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 2
                            DPriceB1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceB2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceB3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 3
                            DPriceC1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceC2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceC3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 4
                            DPriceD1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceD2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceD3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 5
                            DPriceE1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceE2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceE3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 6
                            DPriceF1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceF2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceF3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 7
                            DPriceG1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceG2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceG3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 8
                            DPriceH1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceH2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceH3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 9
                            DPriceI1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceI2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceI3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case Else
                            DPriceJ1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceJ2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceJ3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                    End Select
                Next
            End If

            DArMoldFee.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("ArMoldFee")     '應收模具費
            DPurMold.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("PurMold")         '模具購入費
            DPullerPrice.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("PullerPrice") '引手購入價
            SetFieldData("Suppiler", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Suppiler"))    '外注商
            SetFieldData("MakeCAM", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MakeCAM"))    'Make CAM
            DMoldQty.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MoldQty")         '模型
            DMoldPoint.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MoldPoint")     '穴取

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Quality1") = 1 Then            '品質1
                Dim DBTable1 As DataTable
                SQL = "Select * From F_ManufCOQA1 "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "ManufCOQA1")
                DBTable1 = DBDataSet1.Tables("ManufCOQA1")
                For i = 0 To DBTable1.Rows.Count - 1
                    Select Case DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Seqno")
                        Case 1
                            DQAA1.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Date")
                            DQAA2.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Spec")
                            SetFieldData("DQAA3", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Assembler"))
                            DQAA4.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Surface")
                            DQAA5.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("GenTani")
                            SetFieldData("DQAA6", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Kyoudo"))
                            SetFieldData("DQAA7", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Nyuryoku"))
                            SetFieldData("DQAA8", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Kensin"))
                            SetFieldData("DQAA9", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Water"))
                            SetFieldData("DQAA10", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Dry"))
                            SetFieldData("DQAA11", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Yellow"))
                            SetFieldData("DQAA12", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Mityaku1"))
                            SetFieldData("DQAA13", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Mityaku2"))
                            SetFieldData("DQAA14", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("CPSC"))
                        Case 2
                            DQAB1.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Date")
                            DQAB2.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Spec")
                            SetFieldData("DQAB3", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Assembler"))
                            DQAB4.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Surface")
                            DQAB5.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("GenTani")
                            SetFieldData("DQAB6", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Kyoudo"))
                            SetFieldData("DQAB7", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Nyuryoku"))
                            SetFieldData("DQAB8", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Kensin"))
                            SetFieldData("DQAB9", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Water"))
                            SetFieldData("DQAB10", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Dry"))
                            SetFieldData("DQAB11", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Yellow"))
                            SetFieldData("DQAB12", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Mityaku1"))
                            SetFieldData("DQAB13", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Mityaku2"))
                            SetFieldData("DQAB14", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("CPSC"))
                        Case 3
                            DQAC1.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Date")
                            DQAC2.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Spec")
                            SetFieldData("DQAC3", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Assembler"))
                            DQAC4.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Surface")
                            DQAC5.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("GenTani")
                            SetFieldData("DQAC6", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Kyoudo"))
                            SetFieldData("DQAC7", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Nyuryoku"))
                            SetFieldData("DQAC8", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Kensin"))
                            SetFieldData("DQAC9", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Water"))
                            SetFieldData("DQAC10", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Dry"))
                            SetFieldData("DQAC11", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Yellow"))
                            SetFieldData("DQAC12", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Mityaku1"))
                            SetFieldData("DQAC13", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Mityaku2"))
                            SetFieldData("DQAC14", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("CPSC"))
                        Case Else
                            DQAD1.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Date")
                            DQAD2.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Spec")
                            SetFieldData("DQAD3", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Assembler"))
                            DQAD4.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Surface")
                            DQAD5.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("GenTani")
                            SetFieldData("DQAD6", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Kyoudo"))
                            SetFieldData("DQAD7", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Nyuryoku"))
                            SetFieldData("DQAD8", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Kensin"))
                            SetFieldData("DQAD9", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Water"))
                            SetFieldData("DQAD10", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Dry"))
                            SetFieldData("DQAD11", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Yellow"))
                            SetFieldData("DQAD12", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Mityaku1"))
                            SetFieldData("DQAD13", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Mityaku2"))
                            SetFieldData("DQAD14", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("CPSC"))
                    End Select
                Next
            End If

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Quality2") = 1 Then            '品質2
                Dim DBTable1 As DataTable
                SQL = "Select * From F_ManufCOQA2 "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "ManufCOQA2")
                DBTable1 = DBDataSet1.Tables("ManufCOQA2")
                For i = 0 To DBTable1.Rows.Count - 1
                    Select Case DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Seqno")
                        Case 1
                            DFQAA1.Text = DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Date")
                            SetFieldData("DFQAA2", DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("QACheck"))
                            DFQAA3.Text = DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Remark")
                        Case 2
                            DFQAB1.Text = DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Date")
                            SetFieldData("DFQAB2", DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("QACheck"))
                            DFQAB3.Text = DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Remark")
                        Case 3
                            DFQAC1.Text = DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Date")
                            SetFieldData("DFQAC2", DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("QACheck"))
                            DFQAC3.Text = DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Remark")
                        Case Else
                            DFQAD1.Text = DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Date")
                            SetFieldData("DFQAD2", DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("QACheck"))
                            DFQAD3.Text = DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Remark")
                    End Select
                Next
            End If

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("QAFile") <> "" Then      '測試報告
                LQAFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("QAFile")
            Else
                LQAFile.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("ContactFile") <> "" Then      '切結書
                LContactFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("ContactFile")
            Else
                LContactFile.Visible = False
            End If

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
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

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
                DManuaInSheet1.Visible = True   '表單Sheet-1
                DManuaInSheet2.Visible = True   '表單Sheet-2
                DDescSheet.Visible = True       '說明Sheet
                DDelivery.Visible = True        '交期Sheet
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") < NowDateTime Then
                    'DDelay.Visible = True   '延遲Sheet
                    DDelay.Visible = False   '延遲Sheet
                Else
                    DDelay.Visible = False  '延遲Sheet
                End If

                DOPReady.Visible = True     '工程履歷閱讀
                DOPReadyDesc.Visible = True
                '需閱讀
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("FlowType") = 0 Then
                    DOPReady.BackColor = Color.LightGray
                Else
                    DOPReady.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DOPReadyRqd", "DOPReady", "異常：需閱讀工程履歷")
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
                DBStartTime.Visible = True      '預定開始
                DBEndTime.Visible = True        '預定完成
                DAStartTime.Visible = True      '實際開始
                DAEndTime.Visible = True        '實際完成
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") < NowDateTime Then
                    If DDelay.Visible = True Then
                        DReasonCode.Visible = True     '延遲理由代碼
                        DReasonCode.BackColor = Color.GreenYellow
                        ShowRequiredFieldValidator("DReasonCodeRqd", "DReasonCode", "異常：需輸入延遲理由")
                        DReason.Visible = True         '延遲理由
                        DReason.BackColor = Color.GreenYellow
                        ShowRequiredFieldValidator("DReasonRqd", "DReason", "異常：需輸入延遲理由")
                        DReasonDesc.Visible = True     '延遲其他說明

                    End If
                Else
                    DReasonCode.Visible = False     '延遲理由代碼
                    DReason.Visible = False         '延遲理由
                    DReasonDesc.Visible = False     '延遲其他說明
                End If
                '連結顯示---需再修改
                LMapNo.Visible = True           'MapNo
                LMapFile.Visible = True         '圖樣
                LSAttachFile.Visible = True     '樣品-其他附件
                LQAttachFile1.Visible = True    '品質1-其他附件
                LQAttachFile2.Visible = True    '品質2-其他附件
                LRefFile.Visible = True         '其他附件
                LConfirmFile.Visible = True     '確認書
                LAuthorizeFile.Visible = True   '授權書
                LSampleFile.Visible = True      '樣品圖
                LQAFile.Visible = True          '測試報告
                LContactFile.Visible = True     '切結書
                LBefOP.Visible = True           '工程履歷
                '按鈕位置
                BNG1.Style.Add("Top", Top)     'NG1按鈕
                BNG2.Style.Add("Top", Top)     'NG2按鈕
                BSAVE.Style.Add("Top", Top)    '儲存按鈕
                BOK.Style.Add("Top", Top)      'OK按鈕
                DFormSno.Style.Add("Top", Top) '單號
            End If
        Else
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
            LMapNo.Visible = False           'Map No
            LMapFile.Visible = False         '圖樣
            LSAttachFile.Visible = False     '樣品-其他附件
            LQAttachFile1.Visible = False    '品質1-其他附件
            LQAttachFile2.Visible = False    '品質2-其他附件
            LRefFile.Visible = False         '其他附件
            LConfirmFile.Visible = False     '確認書
            LAuthorizeFile.Visible = False   '授權書
            LSampleFile.Visible = False      '樣品圖
            LQAFile.Visible = False          '測試報告
            LContactFile.Visible = False     '切結書
            LBefOP.Visible = False           '工程履歷
            '按鈕位置
            BNG1.Style.Add("Top", Top)     'NG1按鈕
            BNG2.Style.Add("Top", Top)     'NG2按鈕
            BSAVE.Style.Add("Top", Top)    '儲存按鈕
            BOK.Style.Add("Top", Top)      'OK按鈕
            DFormSno.Style.Add("Top", Top) '單號
        End If

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
        BDate.Attributes("onclick") = "CalendarPicker('Form1.DDate');"  '日期選擇
        BSliderCode.Attributes("onclick") = "SliderPicker('Form1.DSliderCode');"  'Slider Code
        BSpec.Attributes("onclick") = "SpecPicker('Form1.DSpec', 'MANUF');"      '規格
        BOMapNo.Attributes("onclick") = "MapPicker('Ori');"             '原始圖號
        BMMapNo.Attributes("onclick") = "MapPicker('Mod');"             '修改圖號
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
                    Top = 1216
                Else
                    If DDelay.Visible = True Then
                        Top = 1328
                    Else
                        Top = 1216
                    End If
                End If
            End If
        Else
            Top = 1032
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
        'CPSC
        Select Case FindFieldInf("Cpsc")
            Case 0  '顯示
                DCpsc.BackColor = Color.LightGray
                'DCPSC.Enabled = False
                DCpsc.Visible = True
            Case 1  '修改+檢查
                DCpsc.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCpscRqd", "DCpsc", "異常：需輸入CPSC")
                DCpsc.Visible = True
            Case 2  '修改
                DCpsc.BackColor = Color.Yellow
                DCpsc.Visible = True
            Case Else   '隱藏
                DCpsc.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Cpsc", "ZZZZZZ")
        'Slider Group Code
        Select Case FindFieldInf("SliderGRCode")
            Case 0  '顯示
                DSliderGRCode.BackColor = Color.LightGray
                DSliderGRCode.ReadOnly = True
                DSliderGRCode.Visible = True
            Case 1  '修改+檢查
                DSliderGRCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderGRCodeRqd", "DSliderGRCode", "異常：需輸入 Slider Code")
                DSliderGRCode.Visible = True
            Case 2  '修改
                DSliderGRCode.BackColor = Color.Yellow
                DSliderGRCode.Visible = True
            Case Else   '隱藏
                DSliderGRCode.Visible = False
        End Select
        If pPost = "New" Then DSliderGRCode.Text = ""
        'Slider Code
        Select Case FindFieldInf("SliderCode")
            Case 0  '顯示
                DSliderCode.BackColor = Color.LightGray
                DSliderCode.ReadOnly = True
                DSliderCode.Visible = True
                BSliderCode.Visible = False
            Case 1  '修改+檢查
                DSliderCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderCodeRqd", "DSliderCode", "異常：需輸入 Slider Code")
                DSliderCode.Visible = True
                BSliderCode.Visible = True
            Case 2  '修改
                DSliderCode.BackColor = Color.Yellow
                DSliderCode.Visible = True
                BSliderCode.Visible = True
            Case Else   '隱藏
                DSliderCode.Visible = False
                BSliderCode.Visible = False
        End Select
        If pPost = "New" Then DSliderCode.Text = ""
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
        '圖號
        Select Case FindFieldInf("MapNo")
            Case 0  '顯示
                DMapNo.BackColor = Color.LightGray
                DMapNo.ReadOnly = True
                DMapNo.Visible = False
                BOMapNo.Visible = False
                BMMapNo.Visible = False
            Case 1  '修改+檢查
                DMapNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMapNoRqd", "DMapNo", "異常：需輸入圖號")
                DMapNo.Visible = True
                BOMapNo.Visible = True
                BMMapNo.Visible = True
            Case 2  '修改
                DMapNo.BackColor = Color.Yellow
                DMapNo.Visible = True
                BOMapNo.Visible = True
                BMMapNo.Visible = True
            Case Else   '隱藏
                DMapNo.Visible = False
                BOMapNo.Visible = False
                BMMapNo.Visible = False
        End Select
        If pPost = "New" Then DMapNo.Text = ""
        '原始表單號碼
        Select Case FindFieldInf("OFormNo")
            Case 0  '顯示
                DOFormNo.BackColor = Color.LightGray
                DOFormNo.ReadOnly = True
                DOFormNo.Visible = True
            Case 1  '修改+檢查
                DOFormNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOFormNoRqd", "DOFormNo", "異常：需輸入圖號")
                DOFormNo.Visible = True
            Case 2  '修改
                DOFormNo.BackColor = Color.Yellow
                DOFormNo.Visible = True
            Case Else   '隱藏
                DOFormNo.Visible = False
        End Select
        If pPost = "New" Then DOFormNo.Text = ""
        '原始單號
        Select Case FindFieldInf("OFormSno")
            Case 0  '顯示
                DOFormSno.BackColor = Color.LightGray
                DOFormSno.ReadOnly = True
                DOFormSno.Visible = True
            Case 1  '修改+檢查
                DOFormSno.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOFormSnoRqd", "DOFormSno", "異常：需輸入圖號")
                DOFormSno.Visible = True
            Case 2  '修改
                DOFormSno.BackColor = Color.Yellow
                DOFormSno.Visible = True
            Case Else   '隱藏
                DOFormSno.Visible = False
        End Select
        If pPost = "New" Then DOFormSno.Text = ""
        '難易度
        Select Case FindFieldInf("Level")
            Case 0  '顯示
                DLevel.BackColor = Color.LightGray
                DLevel.Visible = True
            Case 1  '修改+檢查
                DLevel.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DLevelRqd", "DLevel", "異常：需輸入難易度")
                DLevel.Visible = True
            Case 2  '修改
                DLevel.BackColor = Color.Yellow
                DLevel.Visible = True
            Case Else   '隱藏
                DLevel.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Level", "ZZZZZZ")
        '組立判別
        Select Case FindFieldInf("Assembler")
            Case 0  '顯示
                DAssembler.BackColor = Color.LightGray
                DAssembler.ReadOnly = True
                DAssembler.Visible = True
            Case 1  '修改+檢查
                DAssembler.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAssemblerRqd", "DAssembler", "異常：需輸入組立判定")
                DAssembler.Visible = True
            Case 2  '修改
                DAssembler.BackColor = Color.Yellow
                DAssembler.Visible = True
            Case Else   '隱藏
                DAssembler.Visible = False
        End Select
        If pPost = "New" Then DAssembler.Text = ""
        '拉頭種類(內製.外注...)
        Select Case FindFieldInf("SliderType1")
            Case 0  '顯示
                DSliderType1.BackColor = Color.LightGray
                DSliderType1.Visible = True
            Case 1  '修改+檢查
                DSliderType1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderType1Rqd", "DSliderType1", "異常：需輸入拉頭種類")
                DSliderType1.Visible = True
            Case 2  '修改
                DSliderType1.BackColor = Color.Yellow
                DSliderType1.Visible = True
            Case Else   '隱藏
                DSliderType1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SliderType1", "ZZZZZZ")
        '拉頭種類(半成品.成品...)
        Select Case FindFieldInf("SliderType2")
            Case 0  '顯示
                DSliderType2.BackColor = Color.LightGray
                DSliderType2.Visible = True
            Case 1  '修改+檢查
                DSliderType2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderType2Rqd", "DSliderType2", "異常：需輸入拉頭種類")
                DSliderType2.Visible = True
            Case 2  '修改
                DSliderType2.BackColor = Color.Yellow
                DSliderType2.Visible = True
            Case Else   '隱藏
                DSliderType2.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SliderType2", "ZZZZZZ")
        '生產地
        Select Case FindFieldInf("ManufPlace")
            Case 0  '顯示
                DManufPlace.BackColor = Color.LightGray
                DManufPlace.Visible = True
            Case 1  '修改+檢查
                DManufPlace.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DManufPlaceRqd", "DManufPlace", "異常：需輸入生產地")
                DManufPlace.Visible = True
            Case 2  '修改
                DManufPlace.BackColor = Color.Yellow
                DManufPlace.Visible = True
            Case Else   '隱藏
                DManufPlace.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ManufPlace", "ZZZZZZ")
        '材質
        Select Case FindFieldInf("Material")
            Case 0  '顯示
                DMaterial.BackColor = Color.LightGray
                DMaterial.Visible = True
            Case 1  '修改+檢查
                DMaterial.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMaterialRqd", "DMaterial", "異常：需輸入材質")
                DMaterial.Visible = True
            Case 2  '修改
                DMaterial.BackColor = Color.Yellow
                DMaterial.Visible = True
            Case Else   '隱藏
                DMaterial.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Material", "ZZZZZZ")
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
        '應收模具費
        Select Case FindFieldInf("ArMoldFee")
            Case 0  '顯示
                DArMoldFee.BackColor = Color.LightGray
                DArMoldFee.ReadOnly = True
                DArMoldFee.Visible = True
            Case 1  '修改+檢查
                DArMoldFee.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DArMoldFeeRqd", "DArMoldFee", "異常：需輸入應收模具費")
                DArMoldFee.Visible = True
            Case 2  '修改
                DArMoldFee.BackColor = Color.Yellow
                DArMoldFee.Visible = True
            Case Else   '隱藏
                DArMoldFee.Visible = False
        End Select
        If pPost = "New" Then DArMoldFee.Text = ""
        '模具購入價
        Select Case FindFieldInf("PurMold")
            Case 0  '顯示
                DPurMold.BackColor = Color.LightGray
                DPurMold.ReadOnly = True
                DPurMold.Visible = True
            Case 1  '修改+檢查
                DPurMold.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPurMoldRqd", "DPurMold", "異常：需輸入模具購入價")
                DPurMold.Visible = True
            Case 2  '修改
                DPurMold.BackColor = Color.Yellow
                DPurMold.Visible = True
            Case Else   '隱藏
                DPurMold.Visible = False
        End Select
        If pPost = "New" Then DPurMold.Text = ""
        '引手購入單價
        Select Case FindFieldInf("PullerPrice")
            Case 0  '顯示
                DPullerPrice.BackColor = Color.LightGray
                DPullerPrice.ReadOnly = True
                DPullerPrice.Visible = True
            Case 1  '修改+檢查
                DPullerPrice.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPullerPriceRqd", "DPullerPrice", "異常：需輸入引手購入單價")
                DPullerPrice.Visible = True
            Case 2  '修改
                DPullerPrice.BackColor = Color.Yellow
                DPullerPrice.Visible = True
            Case Else   '隱藏
                DPullerPrice.Visible = False
        End Select
        If pPost = "New" Then DPullerPrice.Text = ""

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

        '製CAM
        Select Case FindFieldInf("MakeCAM")
            Case 0  '顯示
                DMakeCAM.BackColor = Color.LightGray
                DMakeCAM.Visible = True
            Case 1  '修改+檢查
                DMakeCAM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMakeCAMRqd", "DMakeCAM", "異常：需輸入製CAM者")
                DMakeCAM.Visible = True
            Case 2  '修改
                DMakeCAM.BackColor = Color.Yellow
                DMakeCAM.Visible = True
            Case Else   '隱藏
                DMakeCAM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("MakeCAM", "ZZZZZZ")

        '模型個取-模型
        Select Case FindFieldInf("MoldQty")
            Case 0  '顯示
                DMoldQty.BackColor = Color.LightGray
                DMoldQty.ReadOnly = True
                DMoldQty.Visible = True
            Case 1  '修改+檢查
                DMoldQty.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMoldQtyRqd", "DMoldQty", "異常：需輸入模型個取-模型")
                DMoldQty.Visible = True
            Case 2  '修改
                DMoldQty.BackColor = Color.Yellow
                DMoldQty.Visible = True
            Case Else   '隱藏
                DMoldQty.Visible = False
        End Select
        If pPost = "New" Then DMoldQty.Text = ""
        '模型個取-穴取
        Select Case FindFieldInf("MoldPoint")
            Case 0  '顯示
                DMoldPoint.BackColor = Color.LightGray
                DMoldPoint.ReadOnly = True
                DMoldPoint.Visible = True
            Case 1  '修改+檢查
                DMoldPoint.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMoldPointRqd", "DMoldPoint", "異常：需輸入模型個取-穴取")
                DMoldPoint.Visible = True
            Case 2  '修改
                DMoldPoint.BackColor = Color.Yellow
                DMoldPoint.Visible = True
            Case Else   '隱藏
                DMoldPoint.Visible = False
        End Select
        If pPost = "New" Then DMoldPoint.Text = ""

        '樣品圖
        Select Case FindFieldInf("SampleFile")
            Case 0  '顯示
                DSampleFile.Visible = False
                DSampleFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DSampleFileRqd", "DSampleFile", "異常：需輸入樣品圖")
                DSampleFile.Visible = True
                DSampleFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DSampleFile.Visible = True
                DSampleFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DSampleFile.Visible = False
        End Select

        '圖樣
        Select Case FindFieldInf("MapFile")
            Case 0  '顯示
                DMapFile.Visible = False
                DMapFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DMapFileRqd", "DMapFile", "異常：需輸入圖檔")
                DMapFile.Visible = True
                DMapFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DMapFile.Visible = True
                DMapFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DMapFile.Visible = False
        End Select
        '樣品-其他附件
        Select Case FindFieldInf("SAttachFile")
            Case 0  '顯示
                DSAttachFile.Visible = False
                DSAttachFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DSAttachFileRqd", "DSAttachFile", "異常：需輸入其他附件")
                DSAttachFile.Visible = True
                DSAttachFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DSAttachFile.Visible = True
                DSAttachFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DSAttachFile.Visible = False
        End Select

        '品質1-其他附件
        Select Case FindFieldInf("QAttachFile1")
            Case 0  '顯示
                DQAttachFile1.Visible = False
                DQAttachFile1.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQAttachFile1Rqd", "DQAttachFile1", "異常：需輸入其他附件")
                DQAttachFile1.Visible = True
                DQAttachFile1.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DQAttachFile1.Visible = True
                DQAttachFile1.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DQAttachFile1.Visible = False
        End Select
        '品質2-其他附件
        Select Case FindFieldInf("QAttachFile2")
            Case 0  '顯示
                DQAttachFile2.Visible = False
                DQAttachFile2.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQAttachFile2Rqd", "DQAttachFile2", "異常：需輸入其他附件")
                DQAttachFile2.Visible = True
                DQAttachFile2.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DQAttachFile2.Visible = True
                DQAttachFile2.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DQAttachFile2.Visible = False
        End Select
        '其他附件
        Select Case FindFieldInf("RefFile")
            Case 0  '顯示
                DRefFile.Visible = False
                DRefFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DRefFileRqd", "DRefFile", "異常：需輸入其他附件")
                DRefFile.Visible = True
                DRefFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DRefFile.Visible = True
                DRefFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DRefFile.Visible = False
        End Select
        '確認書
        Select Case FindFieldInf("ConfirmFile")
            Case 0  '顯示
                DConfirmFile.Visible = False
                DConfirmFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DConfirmFileRqd", "DConfirmFile", "異常：需輸入確認書")
                DConfirmFile.Visible = True
                DConfirmFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DConfirmFile.Visible = True
                DConfirmFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DConfirmFile.Visible = False
        End Select
        '授權書
        Select Case FindFieldInf("AuthorizeFile")
            Case 0  '顯示
                DAuthorizeFile.Visible = False
                DAuthorizeFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DAuthorizeFileRqd", "DAuthorizeFile", "異常：需輸入授權書")
                DAuthorizeFile.Visible = True
                DAuthorizeFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DAuthorizeFile.Visible = True
                DAuthorizeFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DAuthorizeFile.Visible = False
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
        '測試報告
        Select Case FindFieldInf("QAFile")
            Case 0  '顯示
                DQAFile.Visible = False
                DQAFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQAFileRqd", "DQAFile", "異常：需輸入測試報告")
                DQAFile.Visible = True
                DQAFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DQAFile.Visible = True
                DQAFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DQAFile.Visible = False
        End Select

        'Sample
        Select Case FindFieldInf("Sample")
            Case 0  '顯示
                DSampleA1.BackColor = Color.LightGray
                DSampleA2.BackColor = Color.LightGray
                DSampleA3.BackColor = Color.LightGray
                DSampleB1.BackColor = Color.LightGray
                DSampleB2.BackColor = Color.LightGray
                DSampleB3.BackColor = Color.LightGray
                DSampleC1.BackColor = Color.LightGray
                DSampleC2.BackColor = Color.LightGray
                DSampleC3.BackColor = Color.LightGray
                DSampleA1.ReadOnly = True
                DSampleA2.ReadOnly = True
                DSampleA3.ReadOnly = True
                DSampleB1.ReadOnly = True
                DSampleB2.ReadOnly = True
                DSampleB3.ReadOnly = True
                DSampleC1.ReadOnly = True
                DSampleC2.ReadOnly = True
                DSampleC3.ReadOnly = True
                DSampleA1.Visible = True
                DSampleA2.Visible = True
                DSampleA3.Visible = True
                DSampleB1.Visible = True
                DSampleB2.Visible = True
                DSampleB3.Visible = True
                DSampleC1.Visible = True
                DSampleC2.Visible = True
                DSampleC3.Visible = True
            Case 1  '修改+檢查
                DSampleA1.BackColor = Color.GreenYellow
                DSampleA2.BackColor = Color.GreenYellow
                DSampleA3.BackColor = Color.GreenYellow
                DSampleB1.BackColor = Color.GreenYellow
                DSampleB2.BackColor = Color.GreenYellow
                DSampleB3.BackColor = Color.GreenYellow
                DSampleC1.BackColor = Color.GreenYellow
                DSampleC2.BackColor = Color.GreenYellow
                DSampleC3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSampleA1Rqd", "DSampleA1", "異常：需輸入樣品資訊(A1)")
                ShowRequiredFieldValidator("DSampleA2Rqd", "DSampleA2", "異常：需輸入樣品資訊(A2)")
                ShowRequiredFieldValidator("DSampleA3Rqd", "DSampleA3", "異常：需輸入樣品資訊(A3)")
                ShowRequiredFieldValidator("DSampleB1Rqd", "DSampleB1", "異常：需輸入樣品資訊(B1)")
                ShowRequiredFieldValidator("DSampleB2Rqd", "DSampleB2", "異常：需輸入樣品資訊(B2)")
                ShowRequiredFieldValidator("DSampleB3Rqd", "DSampleB3", "異常：需輸入樣品資訊(B3)")
                ShowRequiredFieldValidator("DSampleC1Rqd", "DSampleC1", "異常：需輸入樣品資訊(C1)")
                ShowRequiredFieldValidator("DSampleC2Rqd", "DSampleC2", "異常：需輸入樣品資訊(C2)")
                ShowRequiredFieldValidator("DSampleC3Rqd", "DSampleC3", "異常：需輸入樣品資訊(C3)")
                DSampleA1.Visible = True
                DSampleA2.Visible = True
                DSampleA3.Visible = True
                DSampleB1.Visible = True
                DSampleB2.Visible = True
                DSampleB3.Visible = True
                DSampleC1.Visible = True
                DSampleC2.Visible = True
                DSampleC3.Visible = True
            Case 2  '修改
                DSampleA1.BackColor = Color.Yellow
                DSampleA2.BackColor = Color.Yellow
                DSampleA3.BackColor = Color.Yellow
                DSampleB1.BackColor = Color.Yellow
                DSampleB2.BackColor = Color.Yellow
                DSampleB3.BackColor = Color.Yellow
                DSampleC1.BackColor = Color.Yellow
                DSampleC2.BackColor = Color.Yellow
                DSampleC3.BackColor = Color.Yellow
                DSampleA1.Visible = True
                DSampleA2.Visible = True
                DSampleA3.Visible = True
                DSampleB1.Visible = True
                DSampleB2.Visible = True
                DSampleB3.Visible = True
                DSampleC1.Visible = True
                DSampleC2.Visible = True
                DSampleC3.Visible = True
            Case Else   '隱藏
                DSampleA1.Visible = False
                DSampleA2.Visible = False
                DSampleA3.Visible = False
                DSampleB1.Visible = False
                DSampleB2.Visible = False
                DSampleB3.Visible = False
                DSampleC1.Visible = False
                DSampleC2.Visible = False
                DSampleC3.Visible = False
        End Select
        If pPost = "New" Then
            DSampleA1.Text = ""
            DSampleA2.Text = ""
            DSampleA3.Text = ""
            DSampleB1.Text = ""
            DSampleB2.Text = ""
            DSampleB3.Text = ""
            DSampleC1.Text = ""
            DSampleC2.Text = ""
            DSampleC3.Text = ""
        End If
        'Price
        Select Case FindFieldInf("Price")
            Case 0  '顯示
                DPriceA1.BackColor = Color.LightGray
                DPriceA2.BackColor = Color.LightGray
                DPriceA3.BackColor = Color.LightGray
                DPriceB1.BackColor = Color.LightGray
                DPriceB2.BackColor = Color.LightGray
                DPriceB3.BackColor = Color.LightGray
                DPriceC1.BackColor = Color.LightGray
                DPriceC2.BackColor = Color.LightGray
                DPriceC3.BackColor = Color.LightGray
                DPriceD1.BackColor = Color.LightGray
                DPriceD2.BackColor = Color.LightGray
                DPriceD3.BackColor = Color.LightGray
                DPriceE1.BackColor = Color.LightGray
                DPriceE2.BackColor = Color.LightGray
                DPriceE3.BackColor = Color.LightGray
                DPriceF1.BackColor = Color.LightGray
                DPriceF2.BackColor = Color.LightGray
                DPriceF3.BackColor = Color.LightGray
                DPriceG1.BackColor = Color.LightGray
                DPriceG2.BackColor = Color.LightGray
                DPriceG3.BackColor = Color.LightGray
                DPriceH1.BackColor = Color.LightGray
                DPriceH2.BackColor = Color.LightGray
                DPriceH3.BackColor = Color.LightGray
                DPriceI1.BackColor = Color.LightGray
                DPriceI2.BackColor = Color.LightGray
                DPriceI3.BackColor = Color.LightGray
                DPriceJ1.BackColor = Color.LightGray
                DPriceJ2.BackColor = Color.LightGray
                DPriceJ3.BackColor = Color.LightGray

                DPriceA1.ReadOnly = True
                DPriceA3.ReadOnly = True
                DPriceB1.ReadOnly = True
                DPriceB3.ReadOnly = True
                DPriceC1.ReadOnly = True
                DPriceC3.ReadOnly = True
                DPriceD1.ReadOnly = True
                DPriceD3.ReadOnly = True
                DPriceE1.ReadOnly = True
                DPriceE3.ReadOnly = True
                DPriceF1.ReadOnly = True
                DPriceF3.ReadOnly = True
                DPriceG1.ReadOnly = True
                DPriceG3.ReadOnly = True
                DPriceH1.ReadOnly = True
                DPriceH3.ReadOnly = True
                DPriceI1.ReadOnly = True
                DPriceI3.ReadOnly = True
                DPriceJ1.ReadOnly = True
                DPriceJ3.ReadOnly = True

                DPriceA1.Visible = True
                DPriceA2.Visible = True
                DPriceA3.Visible = True
                DPriceB1.Visible = True
                DPriceB2.Visible = True
                DPriceB3.Visible = True
                DPriceC1.Visible = True
                DPriceC2.Visible = True
                DPriceC3.Visible = True
                DPriceD1.Visible = True
                DPriceD2.Visible = True
                DPriceD3.Visible = True
                DPriceE1.Visible = True
                DPriceE2.Visible = True
                DPriceE3.Visible = True
                DPriceF1.Visible = True
                DPriceF2.Visible = True
                DPriceF3.Visible = True
                DPriceG1.Visible = True
                DPriceG2.Visible = True
                DPriceG3.Visible = True
                DPriceH1.Visible = True
                DPriceH2.Visible = True
                DPriceH3.Visible = True
                DPriceI1.Visible = True
                DPriceI2.Visible = True
                DPriceI3.Visible = True
                DPriceJ1.Visible = True
                DPriceJ2.Visible = True
                DPriceJ3.Visible = True
            Case 1  '修改+檢查
                DPriceA1.BackColor = Color.GreenYellow
                DPriceA2.BackColor = Color.GreenYellow
                DPriceA3.BackColor = Color.GreenYellow
                DPriceB1.BackColor = Color.GreenYellow
                DPriceB2.BackColor = Color.GreenYellow
                DPriceB3.BackColor = Color.GreenYellow
                DPriceC1.BackColor = Color.GreenYellow
                DPriceC2.BackColor = Color.GreenYellow
                DPriceC3.BackColor = Color.GreenYellow
                DPriceD1.BackColor = Color.GreenYellow
                DPriceD2.BackColor = Color.GreenYellow
                DPriceD3.BackColor = Color.GreenYellow
                DPriceE1.BackColor = Color.GreenYellow
                DPriceE2.BackColor = Color.GreenYellow
                DPriceE3.BackColor = Color.GreenYellow
                DPriceF1.BackColor = Color.GreenYellow
                DPriceF2.BackColor = Color.GreenYellow
                DPriceF3.BackColor = Color.GreenYellow
                DPriceG1.BackColor = Color.GreenYellow
                DPriceG2.BackColor = Color.GreenYellow
                DPriceG3.BackColor = Color.GreenYellow
                DPriceH1.BackColor = Color.GreenYellow
                DPriceH2.BackColor = Color.GreenYellow
                DPriceH3.BackColor = Color.GreenYellow
                DPriceI1.BackColor = Color.GreenYellow
                DPriceI2.BackColor = Color.GreenYellow
                DPriceI3.BackColor = Color.GreenYellow
                DPriceJ1.BackColor = Color.GreenYellow
                DPriceJ2.BackColor = Color.GreenYellow
                DPriceJ3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPriceA1Rqd", "DPriceA1", "異常：需輸入單價資訊(A1)")
                ShowRequiredFieldValidator("DPriceA2Rqd", "DPriceA2", "異常：需輸入單價資訊(A2)")
                ShowRequiredFieldValidator("DPriceA3Rqd", "DPriceA3", "異常：需輸入單價資訊(A3)")
                ShowRequiredFieldValidator("DPriceB1Rqd", "DPriceB1", "異常：需輸入單價資訊(B1)")
                ShowRequiredFieldValidator("DPriceB2Rqd", "DPriceB2", "異常：需輸入單價資訊(B2)")
                ShowRequiredFieldValidator("DPriceB3Rqd", "DPriceB3", "異常：需輸入單價資訊(B3)")
                ShowRequiredFieldValidator("DPriceC1Rqd", "DPriceC1", "異常：需輸入單價資訊(C1)")
                ShowRequiredFieldValidator("DPriceC2Rqd", "DPriceC2", "異常：需輸入單價資訊(C2)")
                ShowRequiredFieldValidator("DPriceC3Rqd", "DPriceC3", "異常：需輸入單價資訊(C3)")
                ShowRequiredFieldValidator("DPriceD1Rqd", "DPriceD1", "異常：需輸入單價資訊(D1)")
                ShowRequiredFieldValidator("DPriceD2Rqd", "DPriceD2", "異常：需輸入單價資訊(D2)")
                ShowRequiredFieldValidator("DPriceD3Rqd", "DPriceD3", "異常：需輸入單價資訊(D3)")
                ShowRequiredFieldValidator("DPriceE1Rqd", "DPriceE1", "異常：需輸入單價資訊(E1)")
                ShowRequiredFieldValidator("DPriceE2Rqd", "DPriceE2", "異常：需輸入單價資訊(E2)")
                ShowRequiredFieldValidator("DPriceE3Rqd", "DPriceE3", "異常：需輸入單價資訊(E3)")
                ShowRequiredFieldValidator("DPriceF1Rqd", "DPriceF1", "異常：需輸入單價資訊(F1)")
                ShowRequiredFieldValidator("DPriceF2Rqd", "DPriceF2", "異常：需輸入單價資訊(F2)")
                ShowRequiredFieldValidator("DPriceF3Rqd", "DPriceF3", "異常：需輸入單價資訊(F3)")
                ShowRequiredFieldValidator("DPriceG1Rqd", "DPriceG1", "異常：需輸入單價資訊(G1)")
                ShowRequiredFieldValidator("DPriceG2Rqd", "DPriceG2", "異常：需輸入單價資訊(G2)")
                ShowRequiredFieldValidator("DPriceG3Rqd", "DPriceG3", "異常：需輸入單價資訊(G3)")
                ShowRequiredFieldValidator("DPriceH1Rqd", "DPriceH1", "異常：需輸入單價資訊(H1)")
                ShowRequiredFieldValidator("DPriceH2Rqd", "DPriceH2", "異常：需輸入單價資訊(H2)")
                ShowRequiredFieldValidator("DPriceH3Rqd", "DPriceH3", "異常：需輸入單價資訊(H3)")
                ShowRequiredFieldValidator("DPriceI1Rqd", "DPriceI1", "異常：需輸入單價資訊(I1)")
                ShowRequiredFieldValidator("DPriceI2Rqd", "DPriceI2", "異常：需輸入單價資訊(I2)")
                ShowRequiredFieldValidator("DPriceI3Rqd", "DPriceI3", "異常：需輸入單價資訊(I3)")
                ShowRequiredFieldValidator("DPriceJ1Rqd", "DPriceJ1", "異常：需輸入單價資訊(J1)")
                ShowRequiredFieldValidator("DPriceJ2Rqd", "DPriceJ2", "異常：需輸入單價資訊(J2)")
                ShowRequiredFieldValidator("DPriceJ3Rqd", "DPriceJ3", "異常：需輸入單價資訊(J3)")
                DPriceA1.Visible = True
                DPriceA2.Visible = True
                DPriceA3.Visible = True
                DPriceB1.Visible = True
                DPriceB2.Visible = True
                DPriceB3.Visible = True
                DPriceC1.Visible = True
                DPriceC2.Visible = True
                DPriceC3.Visible = True
                DPriceD1.Visible = True
                DPriceD2.Visible = True
                DPriceD3.Visible = True
                DPriceE1.Visible = True
                DPriceE2.Visible = True
                DPriceE3.Visible = True
                DPriceF1.Visible = True
                DPriceF2.Visible = True
                DPriceF3.Visible = True
                DPriceG1.Visible = True
                DPriceG2.Visible = True
                DPriceG3.Visible = True
                DPriceH1.Visible = True
                DPriceH2.Visible = True
                DPriceH3.Visible = True
                DPriceI1.Visible = True
                DPriceI2.Visible = True
                DPriceI3.Visible = True
                DPriceJ1.Visible = True
                DPriceJ2.Visible = True
                DPriceJ3.Visible = True
            Case 2  '修改
                DPriceA1.BackColor = Color.Yellow
                DPriceA2.BackColor = Color.Yellow
                DPriceA3.BackColor = Color.Yellow
                DPriceB1.BackColor = Color.Yellow
                DPriceB2.BackColor = Color.Yellow
                DPriceB3.BackColor = Color.Yellow
                DPriceC1.BackColor = Color.Yellow
                DPriceC2.BackColor = Color.Yellow
                DPriceC3.BackColor = Color.Yellow
                DPriceD1.BackColor = Color.Yellow
                DPriceD2.BackColor = Color.Yellow
                DPriceD3.BackColor = Color.Yellow
                DPriceE1.BackColor = Color.Yellow
                DPriceE2.BackColor = Color.Yellow
                DPriceE3.BackColor = Color.Yellow
                DPriceF1.BackColor = Color.Yellow
                DPriceF2.BackColor = Color.Yellow
                DPriceF3.BackColor = Color.Yellow
                DPriceG1.BackColor = Color.Yellow
                DPriceG2.BackColor = Color.Yellow
                DPriceG3.BackColor = Color.Yellow
                DPriceH1.BackColor = Color.Yellow
                DPriceH2.BackColor = Color.Yellow
                DPriceH3.BackColor = Color.Yellow
                DPriceI1.BackColor = Color.Yellow
                DPriceI2.BackColor = Color.Yellow
                DPriceI3.BackColor = Color.Yellow
                DPriceJ1.BackColor = Color.Yellow
                DPriceJ2.BackColor = Color.Yellow
                DPriceJ3.BackColor = Color.Yellow
                DPriceA1.Visible = True
                DPriceA2.Visible = True
                DPriceA3.Visible = True
                DPriceB1.Visible = True
                DPriceB2.Visible = True
                DPriceB3.Visible = True
                DPriceC1.Visible = True
                DPriceC2.Visible = True
                DPriceC3.Visible = True
                DPriceD1.Visible = True
                DPriceD2.Visible = True
                DPriceD3.Visible = True
                DPriceE1.Visible = True
                DPriceE2.Visible = True
                DPriceE3.Visible = True
                DPriceF1.Visible = True
                DPriceF2.Visible = True
                DPriceF3.Visible = True
                DPriceG1.Visible = True
                DPriceG2.Visible = True
                DPriceG3.Visible = True
                DPriceH1.Visible = True
                DPriceH2.Visible = True
                DPriceH3.Visible = True
                DPriceI1.Visible = True
                DPriceI2.Visible = True
                DPriceI3.Visible = True
                DPriceJ1.Visible = True
                DPriceJ2.Visible = True
                DPriceJ3.Visible = True
            Case Else   '隱藏
                DPriceA1.Visible = False
                DPriceA2.Visible = False
                DPriceA3.Visible = False
                DPriceB1.Visible = False
                DPriceB2.Visible = False
                DPriceB3.Visible = False
                DPriceC1.Visible = False
                DPriceC2.Visible = False
                DPriceC3.Visible = False
                DPriceD1.Visible = False
                DPriceD2.Visible = False
                DPriceD3.Visible = False
                DPriceE1.Visible = False
                DPriceE2.Visible = False
                DPriceE3.Visible = False
                DPriceF1.Visible = False
                DPriceF2.Visible = False
                DPriceF3.Visible = False
                DPriceG1.Visible = False
                DPriceG2.Visible = False
                DPriceG3.Visible = False
                DPriceH1.Visible = False
                DPriceH2.Visible = False
                DPriceH3.Visible = False
                DPriceI1.Visible = False
                DPriceI2.Visible = False
                DPriceI3.Visible = False
                DPriceJ1.Visible = False
                DPriceJ2.Visible = False
                DPriceJ3.Visible = False
        End Select
        If pPost = "New" Then
            DPriceA1.Text = ""
            DPriceA3.Text = ""
            DPriceB1.Text = ""
            DPriceB3.Text = ""
            DPriceC1.Text = ""
            DPriceC3.Text = ""
            DPriceD1.Text = ""
            DPriceD3.Text = ""
            DPriceE1.Text = ""
            DPriceE3.Text = ""
            DPriceF1.Text = ""
            DPriceF3.Text = ""
            DPriceG1.Text = ""
            DPriceG3.Text = ""
            DPriceH1.Text = ""
            DPriceH3.Text = ""
            DPriceI1.Text = ""
            DPriceI3.Text = ""
            DPriceJ1.Text = ""
            DPriceJ3.Text = ""
        End If
        'Quality1
        Select Case FindFieldInf("Quality1")
            Case 0  '顯示
                DQAA1.BackColor = Color.LightGray
                DQAA2.BackColor = Color.LightGray
                DQAA3.BackColor = Color.LightGray
                DQAA4.BackColor = Color.LightGray
                DQAA5.BackColor = Color.LightGray
                DQAA6.BackColor = Color.LightGray
                DQAA7.BackColor = Color.LightGray
                DQAA8.BackColor = Color.LightGray
                DQAA9.BackColor = Color.LightGray
                DQAA10.BackColor = Color.LightGray
                DQAA11.BackColor = Color.LightGray
                DQAA12.BackColor = Color.LightGray
                DQAA13.BackColor = Color.LightGray
                DQAA14.BackColor = Color.LightGray
                DQAB1.BackColor = Color.LightGray
                DQAB2.BackColor = Color.LightGray
                DQAB3.BackColor = Color.LightGray
                DQAB4.BackColor = Color.LightGray
                DQAB5.BackColor = Color.LightGray
                DQAB6.BackColor = Color.LightGray
                DQAB7.BackColor = Color.LightGray
                DQAB8.BackColor = Color.LightGray
                DQAB9.BackColor = Color.LightGray
                DQAB10.BackColor = Color.LightGray
                DQAB11.BackColor = Color.LightGray
                DQAB12.BackColor = Color.LightGray
                DQAB13.BackColor = Color.LightGray
                DQAB14.BackColor = Color.LightGray
                DQAC1.BackColor = Color.LightGray
                DQAC2.BackColor = Color.LightGray
                DQAC3.BackColor = Color.LightGray
                DQAC4.BackColor = Color.LightGray
                DQAC5.BackColor = Color.LightGray
                DQAC6.BackColor = Color.LightGray
                DQAC7.BackColor = Color.LightGray
                DQAC8.BackColor = Color.LightGray
                DQAC9.BackColor = Color.LightGray
                DQAC10.BackColor = Color.LightGray
                DQAC11.BackColor = Color.LightGray
                DQAC12.BackColor = Color.LightGray
                DQAC13.BackColor = Color.LightGray
                DQAC14.BackColor = Color.LightGray
                DQAD1.BackColor = Color.LightGray
                DQAD2.BackColor = Color.LightGray
                DQAD3.BackColor = Color.LightGray
                DQAD4.BackColor = Color.LightGray
                DQAD5.BackColor = Color.LightGray
                DQAD6.BackColor = Color.LightGray
                DQAD7.BackColor = Color.LightGray
                DQAD8.BackColor = Color.LightGray
                DQAD9.BackColor = Color.LightGray
                DQAD10.BackColor = Color.LightGray
                DQAD11.BackColor = Color.LightGray
                DQAD12.BackColor = Color.LightGray
                DQAD13.BackColor = Color.LightGray
                DQAD14.BackColor = Color.LightGray
                DQAA1.ReadOnly = True
                DQAA2.ReadOnly = True
                DQAA4.ReadOnly = True
                DQAA5.ReadOnly = True
                DQAB1.ReadOnly = True
                DQAB2.ReadOnly = True
                DQAB4.ReadOnly = True
                DQAB5.ReadOnly = True
                DQAC1.ReadOnly = True
                DQAC2.ReadOnly = True
                DQAC4.ReadOnly = True
                DQAC5.ReadOnly = True
                DQAD1.ReadOnly = True
                DQAD2.ReadOnly = True
                DQAD4.ReadOnly = True
                DQAD5.ReadOnly = True
                DQAA1.Visible = True
                DQAA2.Visible = True
                DQAA3.Visible = True
                DQAA4.Visible = True
                DQAA5.Visible = True
                DQAA6.Visible = True
                DQAA7.Visible = True
                DQAA8.Visible = True
                DQAA9.Visible = True
                DQAA10.Visible = True
                DQAA11.Visible = True
                DQAA12.Visible = True
                DQAA13.Visible = True
                DQAA14.Visible = True
                DQAB1.Visible = True
                DQAB2.Visible = True
                DQAB3.Visible = True
                DQAB4.Visible = True
                DQAB5.Visible = True
                DQAB6.Visible = True
                DQAB7.Visible = True
                DQAB8.Visible = True
                DQAB9.Visible = True
                DQAB10.Visible = True
                DQAB11.Visible = True
                DQAB12.Visible = True
                DQAB13.Visible = True
                DQAB14.Visible = True
                DQAC1.Visible = True
                DQAC2.Visible = True
                DQAC3.Visible = True
                DQAC4.Visible = True
                DQAC5.Visible = True
                DQAC6.Visible = True
                DQAC7.Visible = True
                DQAC8.Visible = True
                DQAC9.Visible = True
                DQAC10.Visible = True
                DQAC11.Visible = True
                DQAC12.Visible = True
                DQAC13.Visible = True
                DQAC14.Visible = True
                DQAD1.Visible = True
                DQAD2.Visible = True
                DQAD3.Visible = True
                DQAD4.Visible = True
                DQAD5.Visible = True
                DQAD6.Visible = True
                DQAD7.Visible = True
                DQAD8.Visible = True
                DQAD9.Visible = True
                DQAD10.Visible = True
                DQAD11.Visible = True
                DQAD12.Visible = True
                DQAD13.Visible = True
                DQAD14.Visible = True
            Case 1  '修改+檢查
                DQAA1.BackColor = Color.GreenYellow
                DQAA2.BackColor = Color.GreenYellow
                DQAA3.BackColor = Color.GreenYellow
                DQAA4.BackColor = Color.GreenYellow
                DQAA5.BackColor = Color.GreenYellow
                DQAA6.BackColor = Color.GreenYellow
                DQAA7.BackColor = Color.GreenYellow
                DQAA8.BackColor = Color.GreenYellow
                DQAA9.BackColor = Color.GreenYellow
                DQAA10.BackColor = Color.GreenYellow
                DQAA11.BackColor = Color.GreenYellow
                DQAA12.BackColor = Color.GreenYellow
                DQAA13.BackColor = Color.GreenYellow
                DQAA14.BackColor = Color.GreenYellow
                DQAB1.BackColor = Color.GreenYellow
                DQAB2.BackColor = Color.GreenYellow
                DQAB3.BackColor = Color.GreenYellow
                DQAB4.BackColor = Color.GreenYellow
                DQAB5.BackColor = Color.GreenYellow
                DQAB6.BackColor = Color.GreenYellow
                DQAB7.BackColor = Color.GreenYellow
                DQAB8.BackColor = Color.GreenYellow
                DQAB9.BackColor = Color.GreenYellow
                DQAB10.BackColor = Color.GreenYellow
                DQAB11.BackColor = Color.GreenYellow
                DQAB12.BackColor = Color.GreenYellow
                DQAB13.BackColor = Color.GreenYellow
                DQAB14.BackColor = Color.GreenYellow
                DQAC1.BackColor = Color.GreenYellow
                DQAC2.BackColor = Color.GreenYellow
                DQAC3.BackColor = Color.GreenYellow
                DQAC4.BackColor = Color.GreenYellow
                DQAC5.BackColor = Color.GreenYellow
                DQAC6.BackColor = Color.GreenYellow
                DQAC7.BackColor = Color.GreenYellow
                DQAC8.BackColor = Color.GreenYellow
                DQAC9.BackColor = Color.GreenYellow
                DQAC10.BackColor = Color.GreenYellow
                DQAC11.BackColor = Color.GreenYellow
                DQAC12.BackColor = Color.GreenYellow
                DQAC13.BackColor = Color.GreenYellow
                DQAC14.BackColor = Color.GreenYellow
                DQAD1.BackColor = Color.GreenYellow
                DQAD2.BackColor = Color.GreenYellow
                DQAD3.BackColor = Color.GreenYellow
                DQAD4.BackColor = Color.GreenYellow
                DQAD5.BackColor = Color.GreenYellow
                DQAD6.BackColor = Color.GreenYellow
                DQAD7.BackColor = Color.GreenYellow
                DQAD8.BackColor = Color.GreenYellow
                DQAD9.BackColor = Color.GreenYellow
                DQAD10.BackColor = Color.GreenYellow
                DQAD11.BackColor = Color.GreenYellow
                DQAD12.BackColor = Color.GreenYellow
                DQAD13.BackColor = Color.GreenYellow
                DQAD14.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQAA1Rqd", "DQAA1", "異常：需輸入品質資訊(A1)")
                ShowRequiredFieldValidator("DQAA2Rqd", "DQAA2", "異常：需輸入品質資訊(A2)")
                ShowRequiredFieldValidator("DQAA3Rqd", "DQAA3", "異常：需輸入品質資訊(A3)")
                ShowRequiredFieldValidator("DQAA4Rqd", "DQAA4", "異常：需輸入品質資訊(A4)")
                ShowRequiredFieldValidator("DQAA5Rqd", "DQAA5", "異常：需輸入品質資訊(A5)")
                ShowRequiredFieldValidator("DQAA6Rqd", "DQAA6", "異常：需輸入品質資訊(A6)")
                ShowRequiredFieldValidator("DQAA7Rqd", "DQAA7", "異常：需輸入品質資訊(A7)")
                ShowRequiredFieldValidator("DQAA8Rqd", "DQAA8", "異常：需輸入品質資訊(A8)")
                ShowRequiredFieldValidator("DQAA9Rqd", "DQAA9", "異常：需輸入品質資訊(A9)")
                ShowRequiredFieldValidator("DQAA10Rqd", "DQAA10", "異常：需輸入品質資訊(A10)")
                ShowRequiredFieldValidator("DQAA11Rqd", "DQAA11", "異常：需輸入品質資訊(A11)")
                ShowRequiredFieldValidator("DQAA12Rqd", "DQAA12", "異常：需輸入品質資訊(A12)")
                ShowRequiredFieldValidator("DQAA13Rqd", "DQAA13", "異常：需輸入品質資訊(A13)")
                ShowRequiredFieldValidator("DQAA14Rqd", "DQAA14", "異常：需輸入品質資訊(A14)")
                ShowRequiredFieldValidator("DQAB1Rqd", "DQAB1", "異常：需輸入品質資訊(B1)")
                ShowRequiredFieldValidator("DQAB2Rqd", "DQAB2", "異常：需輸入品質資訊(B2)")
                ShowRequiredFieldValidator("DQAB3Rqd", "DQAB3", "異常：需輸入品質資訊(B3)")
                ShowRequiredFieldValidator("DQAB4Rqd", "DQAB4", "異常：需輸入品質資訊(B4)")
                ShowRequiredFieldValidator("DQAB5Rqd", "DQAB5", "異常：需輸入品質資訊(B5)")
                ShowRequiredFieldValidator("DQAB6Rqd", "DQAB6", "異常：需輸入品質資訊(B6)")
                ShowRequiredFieldValidator("DQAB7Rqd", "DQAB7", "異常：需輸入品質資訊(B7)")
                ShowRequiredFieldValidator("DQAB8Rqd", "DQAB8", "異常：需輸入品質資訊(B8)")
                ShowRequiredFieldValidator("DQAB9Rqd", "DQAB9", "異常：需輸入品質資訊(B9)")
                ShowRequiredFieldValidator("DQAB10Rqd", "DQAB10", "異常：需輸入品質資訊(B10)")
                ShowRequiredFieldValidator("DQAB11Rqd", "DQAB11", "異常：需輸入品質資訊(B11)")
                ShowRequiredFieldValidator("DQAB12Rqd", "DQAB12", "異常：需輸入品質資訊(B12)")
                ShowRequiredFieldValidator("DQAB13Rqd", "DQAB13", "異常：需輸入品質資訊(B13)")
                ShowRequiredFieldValidator("DQAB14Rqd", "DQAB14", "異常：需輸入品質資訊(B14)")
                ShowRequiredFieldValidator("DQAC1Rqd", "DQAC1", "異常：需輸入品質資訊(C1)")
                ShowRequiredFieldValidator("DQAC2Rqd", "DQAC2", "異常：需輸入品質資訊(C2)")
                ShowRequiredFieldValidator("DQAC3Rqd", "DQAC3", "異常：需輸入品質資訊(C3)")
                ShowRequiredFieldValidator("DQAC4Rqd", "DQAC4", "異常：需輸入品質資訊(C4)")
                ShowRequiredFieldValidator("DQAC5Rqd", "DQAC5", "異常：需輸入品質資訊(C5)")
                ShowRequiredFieldValidator("DQAC6Rqd", "DQAC6", "異常：需輸入品質資訊(C6)")
                ShowRequiredFieldValidator("DQAC7Rqd", "DQAC7", "異常：需輸入品質資訊(C7)")
                ShowRequiredFieldValidator("DQAC8Rqd", "DQAC8", "異常：需輸入品質資訊(C8)")
                ShowRequiredFieldValidator("DQAC9Rqd", "DQAC9", "異常：需輸入品質資訊(C9)")
                ShowRequiredFieldValidator("DQAC10Rqd", "DQAC10", "異常：需輸入品質資訊(C10)")
                ShowRequiredFieldValidator("DQAC11Rqd", "DQAC11", "異常：需輸入品質資訊(C11)")
                ShowRequiredFieldValidator("DQAC12Rqd", "DQAC12", "異常：需輸入品質資訊(C12)")
                ShowRequiredFieldValidator("DQAC13Rqd", "DQAC13", "異常：需輸入品質資訊(C13)")
                ShowRequiredFieldValidator("DQAC14Rqd", "DQAC14", "異常：需輸入品質資訊(C14)")
                ShowRequiredFieldValidator("DQAD1Rqd", "DQAD1", "異常：需輸入品質資訊(D1)")
                ShowRequiredFieldValidator("DQAD2Rqd", "DQAD2", "異常：需輸入品質資訊(D2)")
                ShowRequiredFieldValidator("DQAD3Rqd", "DQAD3", "異常：需輸入品質資訊(D3)")
                ShowRequiredFieldValidator("DQAD4Rqd", "DQAD4", "異常：需輸入品質資訊(D4)")
                ShowRequiredFieldValidator("DQAD5Rqd", "DQAD5", "異常：需輸入品質資訊(D5)")
                ShowRequiredFieldValidator("DQAD6Rqd", "DQAD6", "異常：需輸入品質資訊(D6)")
                ShowRequiredFieldValidator("DQAD7Rqd", "DQAD7", "異常：需輸入品質資訊(D7)")
                ShowRequiredFieldValidator("DQAD8Rqd", "DQAD8", "異常：需輸入品質資訊(D8)")
                ShowRequiredFieldValidator("DQAD9Rqd", "DQAD9", "異常：需輸入品質資訊(D9)")
                ShowRequiredFieldValidator("DQAD10Rqd", "DQAD10", "異常：需輸入品質資訊(D10)")
                ShowRequiredFieldValidator("DQAD11Rqd", "DQAD11", "異常：需輸入品質資訊(D11)")
                ShowRequiredFieldValidator("DQAD12Rqd", "DQAD12", "異常：需輸入品質資訊(D12)")
                ShowRequiredFieldValidator("DQAD13Rqd", "DQAD13", "異常：需輸入品質資訊(D13)")
                ShowRequiredFieldValidator("DQAD14Rqd", "DQAD14", "異常：需輸入品質資訊(D14)")
                DQAA1.Visible = True
                DQAA2.Visible = True
                DQAA3.Visible = True
                DQAA4.Visible = True
                DQAA5.Visible = True
                DQAA6.Visible = True
                DQAA7.Visible = True
                DQAA8.Visible = True
                DQAA9.Visible = True
                DQAA10.Visible = True
                DQAA11.Visible = True
                DQAA12.Visible = True
                DQAA13.Visible = True
                DQAA14.Visible = True
                DQAB1.Visible = True
                DQAB2.Visible = True
                DQAB3.Visible = True
                DQAB4.Visible = True
                DQAB5.Visible = True
                DQAB6.Visible = True
                DQAB7.Visible = True
                DQAB8.Visible = True
                DQAB9.Visible = True
                DQAB10.Visible = True
                DQAB11.Visible = True
                DQAB12.Visible = True
                DQAB13.Visible = True
                DQAB14.Visible = True
                DQAC1.Visible = True
                DQAC2.Visible = True
                DQAC3.Visible = True
                DQAC4.Visible = True
                DQAC5.Visible = True
                DQAC6.Visible = True
                DQAC7.Visible = True
                DQAC8.Visible = True
                DQAC9.Visible = True
                DQAC10.Visible = True
                DQAC11.Visible = True
                DQAC12.Visible = True
                DQAC13.Visible = True
                DQAC14.Visible = True
                DQAD1.Visible = True
                DQAD2.Visible = True
                DQAD3.Visible = True
                DQAD4.Visible = True
                DQAD5.Visible = True
                DQAD6.Visible = True
                DQAD7.Visible = True
                DQAD8.Visible = True
                DQAD9.Visible = True
                DQAD10.Visible = True
                DQAD11.Visible = True
                DQAD12.Visible = True
                DQAD13.Visible = True
                DQAD14.Visible = True
            Case 2  '修改
                DQAA1.BackColor = Color.Yellow
                DQAA2.BackColor = Color.Yellow
                DQAA3.BackColor = Color.Yellow
                DQAA4.BackColor = Color.Yellow
                DQAA5.BackColor = Color.Yellow
                DQAA6.BackColor = Color.Yellow
                DQAA7.BackColor = Color.Yellow
                DQAA8.BackColor = Color.Yellow
                DQAA9.BackColor = Color.Yellow
                DQAA10.BackColor = Color.Yellow
                DQAA11.BackColor = Color.Yellow
                DQAA12.BackColor = Color.Yellow
                DQAA13.BackColor = Color.Yellow
                DQAA14.BackColor = Color.Yellow
                DQAB1.BackColor = Color.Yellow
                DQAB2.BackColor = Color.Yellow
                DQAB3.BackColor = Color.Yellow
                DQAB4.BackColor = Color.Yellow
                DQAB5.BackColor = Color.Yellow
                DQAB6.BackColor = Color.Yellow
                DQAB7.BackColor = Color.Yellow
                DQAB8.BackColor = Color.Yellow
                DQAB9.BackColor = Color.Yellow
                DQAB10.BackColor = Color.Yellow
                DQAB11.BackColor = Color.Yellow
                DQAB12.BackColor = Color.Yellow
                DQAB13.BackColor = Color.Yellow
                DQAB14.BackColor = Color.Yellow
                DQAC1.BackColor = Color.Yellow
                DQAC2.BackColor = Color.Yellow
                DQAC3.BackColor = Color.Yellow
                DQAC4.BackColor = Color.Yellow
                DQAC5.BackColor = Color.Yellow
                DQAC6.BackColor = Color.Yellow
                DQAC7.BackColor = Color.Yellow
                DQAC8.BackColor = Color.Yellow
                DQAC9.BackColor = Color.Yellow
                DQAC10.BackColor = Color.Yellow
                DQAC11.BackColor = Color.Yellow
                DQAC12.BackColor = Color.Yellow
                DQAC13.BackColor = Color.Yellow
                DQAC14.BackColor = Color.Yellow
                DQAD1.BackColor = Color.Yellow
                DQAD2.BackColor = Color.Yellow
                DQAD3.BackColor = Color.Yellow
                DQAD4.BackColor = Color.Yellow
                DQAD5.BackColor = Color.Yellow
                DQAD6.BackColor = Color.Yellow
                DQAD7.BackColor = Color.Yellow
                DQAD8.BackColor = Color.Yellow
                DQAD9.BackColor = Color.Yellow
                DQAD10.BackColor = Color.Yellow
                DQAD11.BackColor = Color.Yellow
                DQAD12.BackColor = Color.Yellow
                DQAD13.BackColor = Color.Yellow
                DQAD14.BackColor = Color.Yellow
                DQAA1.Visible = True
                DQAA2.Visible = True
                DQAA3.Visible = True
                DQAA4.Visible = True
                DQAA5.Visible = True
                DQAA6.Visible = True
                DQAA7.Visible = True
                DQAA8.Visible = True
                DQAA9.Visible = True
                DQAA10.Visible = True
                DQAA11.Visible = True
                DQAA12.Visible = True
                DQAA13.Visible = True
                DQAA14.Visible = True
                DQAB1.Visible = True
                DQAB2.Visible = True
                DQAB3.Visible = True
                DQAB4.Visible = True
                DQAB5.Visible = True
                DQAB6.Visible = True
                DQAB7.Visible = True
                DQAB8.Visible = True
                DQAB9.Visible = True
                DQAB10.Visible = True
                DQAB11.Visible = True
                DQAB12.Visible = True
                DQAB13.Visible = True
                DQAB14.Visible = True
                DQAC1.Visible = True
                DQAC2.Visible = True
                DQAC3.Visible = True
                DQAC4.Visible = True
                DQAC5.Visible = True
                DQAC6.Visible = True
                DQAC7.Visible = True
                DQAC8.Visible = True
                DQAC9.Visible = True
                DQAC10.Visible = True
                DQAC11.Visible = True
                DQAC12.Visible = True
                DQAC13.Visible = True
                DQAC14.Visible = True
                DQAD1.Visible = True
                DQAD2.Visible = True
                DQAD3.Visible = True
                DQAD4.Visible = True
                DQAD5.Visible = True
                DQAD6.Visible = True
                DQAD7.Visible = True
                DQAD8.Visible = True
                DQAD9.Visible = True
                DQAD10.Visible = True
                DQAD11.Visible = True
                DQAD12.Visible = True
                DQAD13.Visible = True
                DQAD14.Visible = True
            Case Else   '隱藏
                DQAA1.Visible = False
                DQAA2.Visible = False
                DQAA3.Visible = False
                DQAA4.Visible = False
                DQAA5.Visible = False
                DQAA6.Visible = False
                DQAA7.Visible = False
                DQAA8.Visible = False
                DQAA9.Visible = False
                DQAA10.Visible = False
                DQAA11.Visible = False
                DQAA12.Visible = False
                DQAA13.Visible = False
                DQAA14.Visible = False
                DQAB1.Visible = False
                DQAB2.Visible = False
                DQAB3.Visible = False
                DQAB4.Visible = False
                DQAB5.Visible = False
                DQAB6.Visible = False
                DQAB7.Visible = False
                DQAB8.Visible = False
                DQAB9.Visible = False
                DQAB10.Visible = False
                DQAB11.Visible = False
                DQAB12.Visible = False
                DQAB13.Visible = False
                DQAB14.Visible = False
                DQAC1.Visible = False
                DQAC2.Visible = False
                DQAC3.Visible = False
                DQAC4.Visible = False
                DQAC5.Visible = False
                DQAC6.Visible = False
                DQAC7.Visible = False
                DQAC8.Visible = False
                DQAC9.Visible = False
                DQAC10.Visible = False
                DQAC11.Visible = False
                DQAC12.Visible = False
                DQAC13.Visible = False
                DQAC14.Visible = False
                DQAD1.Visible = False
                DQAD2.Visible = False
                DQAD3.Visible = False
                DQAD4.Visible = False
                DQAD5.Visible = False
                DQAD6.Visible = False
                DQAD7.Visible = False
                DQAD8.Visible = False
                DQAD9.Visible = False
                DQAD10.Visible = False
                DQAD11.Visible = False
                DQAD12.Visible = False
                DQAD13.Visible = False
                DQAD14.Visible = False
                DQAF1.Visible = False
                DQAF2.Visible = False
                DQAF3.Visible = False
                DQAF4.Visible = False
        End Select
        If pPost = "New" Then
            DQAA1.Text = ""
            DQAA2.Text = ""
            DQAA4.Text = ""
            DQAA5.Text = ""
            DQAB1.Text = ""
            DQAB2.Text = ""
            DQAB4.Text = ""
            DQAB5.Text = ""
            DQAC1.Text = ""
            DQAC2.Text = ""
            DQAC4.Text = ""
            DQAC5.Text = ""
            DQAD1.Text = ""
            DQAD2.Text = ""
            DQAD4.Text = ""
            DQAD5.Text = ""
        End If
        'Quality2
        Select Case FindFieldInf("Quality2")
            Case 0  '顯示
                DFQAA1.BackColor = Color.LightGray
                DFQAA2.BackColor = Color.LightGray
                DFQAA3.BackColor = Color.LightGray
                DFQAB1.BackColor = Color.LightGray
                DFQAB2.BackColor = Color.LightGray
                DFQAB3.BackColor = Color.LightGray
                DFQAC1.BackColor = Color.LightGray
                DFQAC2.BackColor = Color.LightGray
                DFQAC3.BackColor = Color.LightGray
                DFQAD1.BackColor = Color.LightGray
                DFQAD2.BackColor = Color.LightGray
                DFQAD3.BackColor = Color.LightGray
                DFQAA1.ReadOnly = True
                DFQAA3.ReadOnly = True
                DFQAB1.ReadOnly = True
                DFQAB3.ReadOnly = True
                DFQAC1.ReadOnly = True
                DFQAC3.ReadOnly = True
                DFQAD1.ReadOnly = True
                DFQAD3.ReadOnly = True
                DFQAA1.Visible = True
                DFQAA2.Visible = True
                DFQAA3.Visible = True
                DFQAB1.Visible = True
                DFQAB2.Visible = True
                DFQAB3.Visible = True
                DFQAC1.Visible = True
                DFQAC2.Visible = True
                DFQAC3.Visible = True
                DFQAD1.Visible = True
                DFQAD2.Visible = True
                DFQAD3.Visible = True
            Case 1  '修改+檢查
                DFQAA1.BackColor = Color.GreenYellow
                DFQAA2.BackColor = Color.GreenYellow
                DFQAA3.BackColor = Color.GreenYellow
                DFQAB1.BackColor = Color.GreenYellow
                DFQAB2.BackColor = Color.GreenYellow
                DFQAB3.BackColor = Color.GreenYellow
                DFQAC1.BackColor = Color.GreenYellow
                DFQAC2.BackColor = Color.GreenYellow
                DFQAC3.BackColor = Color.GreenYellow
                DFQAD1.BackColor = Color.GreenYellow
                DFQAD2.BackColor = Color.GreenYellow
                DFQAD3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFQAA1Rqd", "DFQAA1", "異常：需輸入品質分析資訊(FA1)")
                ShowRequiredFieldValidator("DFQAA2Rqd", "DFQAA2", "異常：需輸入品質分析資訊(FA2)")
                ShowRequiredFieldValidator("DFQAA3Rqd", "DFQAA3", "異常：需輸入品質分析資訊(FA3)")
                ShowRequiredFieldValidator("DFQAB1Rqd", "DFQAB1", "異常：需輸入品質分析資訊(FB1)")
                ShowRequiredFieldValidator("DFQAB2Rqd", "DFQAB2", "異常：需輸入品質分析資訊(FB2)")
                ShowRequiredFieldValidator("DFQAB3Rqd", "DFQAB3", "異常：需輸入品質分析資訊(FB3)")
                ShowRequiredFieldValidator("DFQAC1Rqd", "DFQAC1", "異常：需輸入品質分析資訊(FC1)")
                ShowRequiredFieldValidator("DFQAC2Rqd", "DFQAC2", "異常：需輸入品質分析資訊(FC2)")
                ShowRequiredFieldValidator("DFQAC3Rqd", "DFQAC3", "異常：需輸入品質分析資訊(FC3)")
                ShowRequiredFieldValidator("DFQAD1Rqd", "DFQAD1", "異常：需輸入品質分析資訊(FD1)")
                ShowRequiredFieldValidator("DFQAD2Rqd", "DFQAD2", "異常：需輸入品質分析資訊(FD2)")
                ShowRequiredFieldValidator("DFQAD3Rqd", "DFQAD3", "異常：需輸入品質分析資訊(FD3)")
                DFQAA1.Visible = True
                DFQAA2.Visible = True
                DFQAA3.Visible = True
                DFQAB1.Visible = True
                DFQAB2.Visible = True
                DFQAB3.Visible = True
                DFQAC1.Visible = True
                DFQAC2.Visible = True
                DFQAC3.Visible = True
                DFQAD1.Visible = True
                DFQAD2.Visible = True
                DFQAD3.Visible = True
            Case 2  '修改
                DFQAA1.BackColor = Color.Yellow
                DFQAA2.BackColor = Color.Yellow
                DFQAA3.BackColor = Color.Yellow
                DFQAB1.BackColor = Color.Yellow
                DFQAB2.BackColor = Color.Yellow
                DFQAB3.BackColor = Color.Yellow
                DFQAC1.BackColor = Color.Yellow
                DFQAC2.BackColor = Color.Yellow
                DFQAC3.BackColor = Color.Yellow
                DFQAD1.BackColor = Color.Yellow
                DFQAD2.BackColor = Color.Yellow
                DFQAD3.BackColor = Color.Yellow
                DFQAA1.Visible = True
                DFQAA2.Visible = True
                DFQAA3.Visible = True
                DFQAB1.Visible = True
                DFQAB2.Visible = True
                DFQAB3.Visible = True
                DFQAC1.Visible = True
                DFQAC2.Visible = True
                DFQAC3.Visible = True
                DFQAD1.Visible = True
                DFQAD2.Visible = True
                DFQAD3.Visible = True
            Case Else   '隱藏
                DFQAA1.Visible = False
                DFQAA2.Visible = False
                DFQAA3.Visible = False
                DFQAB1.Visible = False
                DFQAB2.Visible = False
                DFQAB3.Visible = False
                DFQAC1.Visible = False
                DFQAC2.Visible = False
                DFQAC3.Visible = False
                DFQAD1.Visible = False
                DFQAD2.Visible = False
                DFQAD3.Visible = False
        End Select
        If pPost = "New" Then
            DFQAA1.Text = ""
            DFQAA3.Text = ""
            DFQAB1.Text = ""
            DFQAB3.Text = ""
            DFQAC1.Text = ""
            DFQAC3.Text = ""
            DFQAD1.Text = ""
            DFQAD3.Text = ""
        End If
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

        If Left(pFieldName, 3) = "DQA" Then
            idx = FindFieldInf("Quality1")
        Else
            idx = FindFieldInf(pFieldName)
        End If

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
        'CPSC
        If pFieldName = "Cpsc" Then
            DCpsc.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCpsc.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='100' and DKey='CPSC' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCpsc.Items.Add(ListItem1)
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

        'Make CAM
        If pFieldName = "MakeCAM" Then
            DMakeCAM.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMakeCAM.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='200' and DKey='MAKECAM' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DMakeCAM.Items.Add(ListItem1)
                Next
            End If
        End If

        '難易度
        If pFieldName = "Level" Then
            DLevel.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLevel.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='007' and (DKey like 'In%' or DKey = '') Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLevel.Items.Add(ListItem1)
                Next
            End If
        End If
        '拉頭種類(內製.外注...)
        If pFieldName = "SliderType1" Then
            DSliderType1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSliderType1.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='200' and DKey='SLIDERTYPE1' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSliderType1.Items.Add(ListItem1)
                Next
            End If
        End If
        '拉頭種類(半成品.成品...)
        If pFieldName = "SliderType2" Then
            DSliderType2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSliderType2.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='200' and DKey='SLIDERTYPE2' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSliderType2.Items.Add(ListItem1)
                Next
            End If
        End If
        '生產地
        If pFieldName = "ManufPlace" Then
            DManufPlace.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DManufPlace.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='200' and DKey='MANUFPLACE' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DManufPlace.Items.Add(ListItem1)
                Next
            End If
        End If
        '材質
        If pFieldName = "Material" Then
            DMaterial.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMaterial.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='101'  Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DMaterial.Items.Add(ListItem1)
                Next
            End If
        End If
        'QA--------------------------------------
        'DQAA3
        If pFieldName = "DQAA3" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAA3.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAA3.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAA3.Items.Count - 1
                    If DQAA3.Items(i).Value = pName Then
                        DQAA3.Items(i).Selected = True
                    Else
                        DQAA3.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAA7
        If pFieldName = "DQAA7" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAA7.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAA7.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAA7.Items.Count - 1
                    If DQAA7.Items(i).Value = pName Then
                        DQAA7.Items(i).Selected = True
                    Else
                        DQAA7.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAA8
        If pFieldName = "DQAA8" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAA8.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAA8.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAA8.Items.Count - 1
                    If DQAA8.Items(i).Value = pName Then
                        DQAA8.Items(i).Selected = True
                    Else
                        DQAA8.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAA9
        If pFieldName = "DQAA9" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAA9.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAA9.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAA9.Items.Count - 1
                    If DQAA9.Items(i).Value = pName Then
                        DQAA9.Items(i).Selected = True
                    Else
                        DQAA9.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAA10
        If pFieldName = "DQAA10" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAA10.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAA10.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAA10.Items.Count - 1
                    If DQAA10.Items(i).Value = pName Then
                        DQAA10.Items(i).Selected = True
                    Else
                        DQAA10.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAA11
        If pFieldName = "DQAA11" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAA11.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAA11.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAA11.Items.Count - 1
                    If DQAA11.Items(i).Value = pName Then
                        DQAA11.Items(i).Selected = True
                    Else
                        DQAA11.Items(i).Selected = False
                    End If
                Next
            End If
        End If

        'DQAA6
        If pFieldName = "DQAA6" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAA6.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAA6.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAA6.Items.Count - 1
                    If DQAA6.Items(i).Value = pName Then
                        DQAA6.Items(i).Selected = True
                    Else
                        DQAA6.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAA12
        If pFieldName = "DQAA12" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAA12.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAA12.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAA12.Items.Count - 1
                    If DQAA12.Items(i).Value = pName Then
                        DQAA12.Items(i).Selected = True
                    Else
                        DQAA12.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAA13
        If pFieldName = "DQAA13" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAA13.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAA13.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAA13.Items.Count - 1
                    If DQAA13.Items(i).Value = pName Then
                        DQAA13.Items(i).Selected = True
                    Else
                        DQAA13.Items(i).Selected = False
                    End If
                Next
            End If
        End If

        'DQAA14
        If pFieldName = "DQAA14" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAA14.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAA14.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAA14.Items.Count - 1
                    If DQAA14.Items(i).Value = pName Then
                        DQAA14.Items(i).Selected = True
                    Else
                        DQAA14.Items(i).Selected = False
                    End If
                Next
            End If
        End If

        'DQAB3
        If pFieldName = "DQAB3" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAB3.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAB3.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAB3.Items.Count - 1
                    If DQAB3.Items(i).Value = pName Then
                        DQAB3.Items(i).Selected = True
                    Else
                        DQAB3.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAB7
        If pFieldName = "DQAB7" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAB7.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAB7.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAB7.Items.Count - 1
                    If DQAB7.Items(i).Value = pName Then
                        DQAB7.Items(i).Selected = True
                    Else
                        DQAB7.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAB8
        If pFieldName = "DQAB8" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAB8.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAB8.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAB8.Items.Count - 1
                    If DQAB8.Items(i).Value = pName Then
                        DQAB8.Items(i).Selected = True
                    Else
                        DQAB8.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAB9
        If pFieldName = "DQAB9" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAB9.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAB9.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAB9.Items.Count - 1
                    If DQAB9.Items(i).Value = pName Then
                        DQAB9.Items(i).Selected = True
                    Else
                        DQAB9.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAB10
        If pFieldName = "DQAB10" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAB10.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAB10.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAB10.Items.Count - 1
                    If DQAB10.Items(i).Value = pName Then
                        DQAB10.Items(i).Selected = True
                    Else
                        DQAB10.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAB11
        If pFieldName = "DQAB11" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAB11.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAB11.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAB11.Items.Count - 1
                    If DQAB11.Items(i).Value = pName Then
                        DQAB11.Items(i).Selected = True
                    Else
                        DQAB11.Items(i).Selected = False
                    End If
                Next
            End If
        End If

        'DQAB6
        If pFieldName = "DQAB6" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAB6.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAB6.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAB6.Items.Count - 1
                    If DQAB6.Items(i).Value = pName Then
                        DQAB6.Items(i).Selected = True
                    Else
                        DQAB6.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAB12
        If pFieldName = "DQAB12" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAB12.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAB12.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAB12.Items.Count - 1
                    If DQAB12.Items(i).Value = pName Then
                        DQAB12.Items(i).Selected = True
                    Else
                        DQAB12.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAB13
        If pFieldName = "DQAB13" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAB13.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAB13.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAB13.Items.Count - 1
                    If DQAB13.Items(i).Value = pName Then
                        DQAB13.Items(i).Selected = True
                    Else
                        DQAB13.Items(i).Selected = False
                    End If
                Next
            End If
        End If

        'DQAB14
        If pFieldName = "DQAB14" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAB14.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAB14.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAB14.Items.Count - 1
                    If DQAB14.Items(i).Value = pName Then
                        DQAB14.Items(i).Selected = True
                    Else
                        DQAB14.Items(i).Selected = False
                    End If
                Next
            End If
        End If

        'DQAC3
        If pFieldName = "DQAC3" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAC3.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAC3.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAC3.Items.Count - 1
                    If DQAC3.Items(i).Value = pName Then
                        DQAC3.Items(i).Selected = True
                    Else
                        DQAC3.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAC7
        If pFieldName = "DQAC7" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAC7.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAC7.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAC7.Items.Count - 1
                    If DQAC7.Items(i).Value = pName Then
                        DQAC7.Items(i).Selected = True
                    Else
                        DQAC7.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAC8
        If pFieldName = "DQAC8" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAC8.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAC8.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAC8.Items.Count - 1
                    If DQAC8.Items(i).Value = pName Then
                        DQAC8.Items(i).Selected = True
                    Else
                        DQAC8.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAC9
        If pFieldName = "DQAC9" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAC9.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAC9.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAC9.Items.Count - 1
                    If DQAC9.Items(i).Value = pName Then
                        DQAC9.Items(i).Selected = True
                    Else
                        DQAC9.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAC10
        If pFieldName = "DQAC10" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAC10.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAC10.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAC10.Items.Count - 1
                    If DQAC10.Items(i).Value = pName Then
                        DQAC10.Items(i).Selected = True
                    Else
                        DQAC10.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAC11
        If pFieldName = "DQAC11" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAC11.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAC11.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAC11.Items.Count - 1
                    If DQAC11.Items(i).Value = pName Then
                        DQAC11.Items(i).Selected = True
                    Else
                        DQAC11.Items(i).Selected = False
                    End If
                Next
            End If
        End If

        'DQAC6
        If pFieldName = "DQAC6" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAC6.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAC6.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAC6.Items.Count - 1
                    If DQAC6.Items(i).Value = pName Then
                        DQAC6.Items(i).Selected = True
                    Else
                        DQAC6.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAC12
        If pFieldName = "DQAC12" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAC12.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAC12.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAC12.Items.Count - 1
                    If DQAC12.Items(i).Value = pName Then
                        DQAC12.Items(i).Selected = True
                    Else
                        DQAC12.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAC13
        If pFieldName = "DQAC13" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAC13.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAC13.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAC13.Items.Count - 1
                    If DQAC13.Items(i).Value = pName Then
                        DQAC13.Items(i).Selected = True
                    Else
                        DQAC13.Items(i).Selected = False
                    End If
                Next
            End If
        End If

        'DQAC14
        If pFieldName = "DQAC14" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAC14.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAC14.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAC14.Items.Count - 1
                    If DQAC14.Items(i).Value = pName Then
                        DQAC14.Items(i).Selected = True
                    Else
                        DQAC14.Items(i).Selected = False
                    End If
                Next
            End If
        End If

        'DQAD3
        If pFieldName = "DQAD3" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAD3.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAD3.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAD3.Items.Count - 1
                    If DQAD3.Items(i).Value = pName Then
                        DQAD3.Items(i).Selected = True
                    Else
                        DQAD3.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAD7
        If pFieldName = "DQAD7" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAD7.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAD7.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAD7.Items.Count - 1
                    If DQAD7.Items(i).Value = pName Then
                        DQAD7.Items(i).Selected = True
                    Else
                        DQAD7.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAD8
        If pFieldName = "DQAD8" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAD8.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAD8.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAD8.Items.Count - 1
                    If DQAD8.Items(i).Value = pName Then
                        DQAD8.Items(i).Selected = True
                    Else
                        DQAD8.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAD9
        If pFieldName = "DQAD9" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAD9.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAD9.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAD9.Items.Count - 1
                    If DQAD9.Items(i).Value = pName Then
                        DQAD9.Items(i).Selected = True
                    Else
                        DQAD9.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAD10
        If pFieldName = "DQAD10" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAD10.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAD10.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAD10.Items.Count - 1
                    If DQAD10.Items(i).Value = pName Then
                        DQAD10.Items(i).Selected = True
                    Else
                        DQAD10.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAD11
        If pFieldName = "DQAD11" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAD11.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAD11.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAD11.Items.Count - 1
                    If DQAD11.Items(i).Value = pName Then
                        DQAD11.Items(i).Selected = True
                    Else
                        DQAD11.Items(i).Selected = False
                    End If
                Next
            End If
        End If

        'DQAD6
        If pFieldName = "DQAD6" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAD6.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAD6.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAD6.Items.Count - 1
                    If DQAD6.Items(i).Value = pName Then
                        DQAD6.Items(i).Selected = True
                    Else
                        DQAD6.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAD12
        If pFieldName = "DQAD12" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAD12.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAD12.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAD12.Items.Count - 1
                    If DQAD12.Items(i).Value = pName Then
                        DQAD12.Items(i).Selected = True
                    Else
                        DQAD12.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'DQAD13
        If pFieldName = "DQAD13" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAD13.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAD13.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAD13.Items.Count - 1
                    If DQAD13.Items(i).Value = pName Then
                        DQAD13.Items(i).Selected = True
                    Else
                        DQAD13.Items(i).Selected = False
                    End If
                Next
            End If
        End If

        'DQAD14
        If pFieldName = "DQAD14" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DQAD14.Items.Clear()
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQAD14.Items.Add(ListItem1)
                End If
            Else
                For i = 0 To DQAD14.Items.Count - 1
                    If DQAD14.Items(i).Value = pName Then
                        DQAD14.Items(i).Selected = True
                    Else
                        DQAD14.Items(i).Selected = False
                    End If
                Next
            End If
        End If
        'QA-1 
        'DFQAA2
        If pFieldName = "DFQAA2" Then
            DFQAA2.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DFQAA2.Items.Add(ListItem1)
        End If
        'DFQAB2
        If pFieldName = "DFQAB2" Then
            DFQAB2.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DFQAB2.Items.Add(ListItem1)
        End If
        'DFQAC2
        If pFieldName = "DFQAC2" Then
            DFQAC2.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DFQAC2.Items.Add(ListItem1)
        End If
        'DFQAD2
        If pFieldName = "DFQAD2" Then
            DFQAD2.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DFQAD2.Items.Add(ListItem1)
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
                Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
                Dim Message As String = ""

                'Check附圖-1Size及格式
                If ErrCode = 0 Then
                    If DMapFile.Visible Then
                        If DMapFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DMapFile)
                        End If
                    End If
                End If
                'Check樣品其他附件-Size及格式
                If ErrCode = 0 Then
                    If DSAttachFile.Visible Then
                        If DSAttachFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DSAttachFile)
                        End If
                    End If
                End If
                'Check品質1其他附件-Size及格式
                If ErrCode = 0 Then
                    If DQAttachFile1.Visible Then
                        If DQAttachFile1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DQAttachFile1)
                        End If
                    End If
                End If
                'Check品質2其他附件-Size及格式
                If ErrCode = 0 Then
                    If DQAttachFile2.Visible Then
                        If DQAttachFile2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DQAttachFile2)
                        End If
                    End If
                End If
                'Check參考附件-Size及格式
                If ErrCode = 0 Then
                    If DRefFile.Visible Then
                        If DRefFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DRefFile)
                        End If
                    End If
                End If
                'Check確認書Size及格式
                If ErrCode = 0 Then
                    If DConfirmFile.Visible Then
                        If DConfirmFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DConfirmFile)
                        End If
                    End If
                End If
                'Check授權書Size及格式
                If ErrCode = 0 Then
                    If DAuthorizeFile.Visible Then
                        If DAuthorizeFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DAuthorizeFile)
                        End If
                    End If
                End If
                'Check樣品圖Size及格式
                If ErrCode = 0 Then
                    If DSampleFile.Visible Then
                        If DSampleFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DSampleFile)
                        End If
                    End If
                End If
                'CheckQA附件Size及格式
                If ErrCode = 0 Then
                    If DQAFile.Visible Then
                        If DQAFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DQAFile)
                        End If
                    End If
                End If
                'Check切結書
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

        'Check附圖-1Size及格式
        If ErrCode = 0 Then
            If DMapFile.Visible Then
                If DMapFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DMapFile)
                End If
            End If
        End If
        'Check樣品其他附件-Size及格式
        If ErrCode = 0 Then
            If DSAttachFile.Visible Then
                If DSAttachFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DSAttachFile)
                End If
            End If
        End If
        'Check品質1其他附件-Size及格式
        If ErrCode = 0 Then
            If DQAttachFile1.Visible Then
                If DQAttachFile1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DQAttachFile1)
                End If
            End If
        End If
        'Check品質2其他附件-Size及格式
        If ErrCode = 0 Then
            If DQAttachFile2.Visible Then
                If DQAttachFile2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DQAttachFile2)
                End If
            End If
        End If
        'Check附圖-2Size及格式
        If ErrCode = 0 Then
            If DRefFile.Visible Then
                If DRefFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DRefFile)
                End If
            End If
        End If
        'Check確認書Size及格式
        If ErrCode = 0 Then
            If DConfirmFile.Visible Then
                If DConfirmFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DConfirmFile)
                End If
            End If
        End If
        'Check授權書Size及格式
        If ErrCode = 0 Then
            If DAuthorizeFile.Visible Then
                If DAuthorizeFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DAuthorizeFile)
                End If
            End If
        End If
        'Check樣品圖Size及格式
        If ErrCode = 0 Then
            If DSampleFile.Visible Then
                If DSampleFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DSampleFile)
                End If
            End If
        End If
        'CheckQA附件Size及格式
        If ErrCode = 0 Then
            If DQAFile.Visible Then
                If DQAFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DQAFile)
                End If
            End If
        End If
        'Check切結書
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


        '--檢查委託書No---------
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                ErrCode = oCommon.CommissionNo("000003", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9060
                End If
            End If
        End If

        If ErrCode = 0 Then
            Dim Run As Boolean = True           '是否執行
            Dim RepeatRun As Boolean = False    '是否重覆執行
            Dim MultiJob As Integer = 0         '多人核定
            Dim wLevel As String = DLevel.SelectedValue     '難易度

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
                            'oFlow.NewFlow("000003", NewFormSno, wStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                            '
                            oFlow.NewFlow("000003", NewFormSno, wStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
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

                    SQL = "Select UserID From M_Users "
                    SQL = SQL & " Where Active = 1 "
                    SQL = SQL & "   And UserName =  '" & DMakeCAM.SelectedValue & "'"
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
                        'RtnCode = oFlow.EndFlow("000003", NewFormSno, pNextStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                        '
                        RtnCode = oFlow.EndFlow("000003", NewFormSno, pNextStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
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
                            AddCommissionNo("000003", NewFormSno)
                        End If
                    Else
                        If pNextStep = 999 Then     '工程結束嗎?
                            If pFun = "OK" Then ModifyData(pFun, CStr(wOKSts)) '更新表單資料
                            If pFun = "NG1" Then ModifyData(pFun, CStr(wNGSts1))
                            If pFun = "NG2" Then ModifyData(pFun, CStr(wNGSts2))
                        Else
                            ModifyData(pFun, "0")         '更新表單資料 Sts=0(未結)
                        End If
                        AddCommissionNo("000003", wFormSno)
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("ManufInFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String
        Dim RtnCode As Integer

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        OleDbConnection1.Open()
        SQl = "Insert into F_ManufInSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, "                         '1~4
        SQl = SQl + "No, Date, Division, Person, SliderCode, "                      '5~9
        SQl = SQl + "SliderGRCode, Spec, Cpsc, MapNo, MapFile, RefFile, "               '10~14
        SQl = SQl + "Level, Assembler, SliderType1, SliderType2, ManufPlace, "      '15~19
        SQl = SQl + "Material, MaterialOther, SellVendor, Buyer, "  '20~24
        SQl = SQl + "ConfirmFile, AuthorizeFile, DevReason, Sample, Price, "        '25~29
        SQl = SQl + "ArMoldFee, PurMold, PullerPrice, Suppiler, MakeCAM, MoldQty, "          '29~33
        SQl = SQl + "MoldPoint, Quality1, Quality2, QAFile, SampleFile,  "    '34~38
        SQl = SQl + "ContactFile, OFormNo, OFormSno, QCLT, "    '34~38
        SQl = SQl + "SAttachFile, QAttachFile1, QAttachFile2, "               '39~42
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "               '39~42
        SQl = SQl + ")  "
        SQl = SQl + "VALUES( "
        SQl = SQl + " '1', "                                '狀態(0:未結,1:已結NG,2:已結OK)
        SQl = SQl + " '" + NowDateTime + "', "              '結案日
        SQl = SQl + " '000003', "                           '表單代號
        SQl = SQl + " '" + CStr(NewFormSno) + "', "         '表單流水號

        SQl = SQl + " '" + YKK.ReplaceString(DNo.Text) + "', "                 'NO
        SQl = SQl + " '" + DDate.Text + "', "               '日期
        SQl = SQl + " '" + DDivision.SelectedValue + "', "  '部門
        SQl = SQl + " '" + DPerson.SelectedValue + "', "    '擔當
        SQl = SQl + " '" + YKK.ReplaceString(DSliderCode.Text) + "', "         'Slider Code(顯示用)
        SQl = SQl + " '" + YKK.ReplaceString(DSliderGRCode.Text) + "', "       'Slider Group Code
        SQl = SQl + " '" + YKK.ReplaceString(DSpec.Text) + "', "               '規格
        SQl = SQl + " '" + DCpsc.SelectedValue + "', "      'CPSC
        SQl = SQl + " '" + YKK.ReplaceString(DMapNo.Text) + "', "              '圖號

        FileName = ""
        If DMapFile.Visible Then                           '圖檔1
            If DMapFile.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Map" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "Map" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DMapFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        FileName = ""
        If DRefFile.Visible Then                           '圖檔2
            If DRefFile.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Ref" & "-" & UploadDateTime & "-" & Right(DRefFile.PostedFile.FileName, InStr(StrReverse(DRefFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "Ref" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DRefFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DRefFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        SQl = SQl + " '" + DLevel.SelectedValue + "', "           '難易度
        SQl = SQl + " N'" + YKK.ReplaceString(DAssembler.Text) + "', "                '組立判定
        SQl = SQl + " '" + DSliderType1.SelectedValue + "', "     '拉頭種類-1
        SQl = SQl + " '" + DSliderType2.SelectedValue + "', "     '拉頭種類-2
        SQl = SQl + " '" + DManufPlace.SelectedValue + "', "      '生產地
        SQl = SQl + " '" + DMaterial.SelectedValue + "', "        '材質
        SQl = SQl + " '" + "" + "', "            '材質其他說明
        SQl = SQl + " N'" + YKK.ReplaceString(DSellVendor.Text) + "', "               '委託廠商
        SQl = SQl + " N'" + DBuyer.SelectedValue + "', "                    'Buyer

        FileName = ""
        If DConfirmFile.Visible Then                              '確認書
            If DConfirmFile.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Confirm" & "-" & UploadDateTime & "-" & Right(DConfirmFile.PostedFile.FileName, InStr(StrReverse(DConfirmFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "Confirm" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DConfirmFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DConfirmFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        FileName = ""
        If DAuthorizeFile.Visible Then                            '授權書
            If DAuthorizeFile.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Authorize" & "-" & UploadDateTime & "-" & Right(DAuthorizeFile.PostedFile.FileName, InStr(StrReverse(DAuthorizeFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "Authorize" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DAuthorizeFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DAuthorizeFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        SQl = SQl + " N'" + YKK.ReplaceString(DDevReason.Text) + "', "          '開發理由
        If DSampleA1.Text <> "" Then                          'Sample(顯示用)
            SQl = SQl + " '1', "
        Else
            SQl = SQl + " '0', "
        End If
        If DPriceA1.Text <> "" Then                           'Price(顯示用)
            SQl = SQl + " '1', "
        Else
            SQl = SQl + " '0', "
        End If
        SQl = SQl + " '" + YKK.ReplaceString(DArMoldFee.Text) + "', "          '應收模具費
        SQl = SQl + " '" + YKK.ReplaceString(DPurMold.Text) + "', "            '模具購入費
        SQl = SQl + " '" + YKK.ReplaceString(DPullerPrice.Text) + "', "        '引手購入單價
        SQl = SQl + " N'" + DSuppiler.SelectedValue + "', " '外注商
        SQl = SQl + " N'" + DMakeCAM.SelectedValue + "', "  'Make CAM
        SQl = SQl + " '" + YKK.ReplaceString(DMoldQty.Text) + "', "            '模型數
        SQl = SQl + " '" + YKK.ReplaceString(DMoldPoint.Text) + "', "          '穴取數
        If DQAA1.Text <> "" Then                        '品質分析-1(顯示用)
            SQl = SQl + " '1', "
        Else
            SQl = SQl + " '0', "
        End If
        If DFQAA1.Text <> "" Then                        '品質分析-2(顯示用)
            SQl = SQl + " '1', "
        Else
            SQl = SQl + " '0', "
        End If

        FileName = ""
        If DQAFile.Visible Then                      '品質分析附件
            If DQAFile.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QA" & "-" & UploadDateTime & "-" & Right(DQAFile.PostedFile.FileName, InStr(StrReverse(DQAFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QA" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQAFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQAFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        FileName = ""
        If DSampleFile.Visible Then                      '樣品圖檔
            If DSampleFile.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Sample" & "-" & UploadDateTime & "-" & Right(DSampleFile.PostedFile.FileName, InStr(StrReverse(DSampleFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "Sample" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSampleFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DSampleFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        FileName = ""
        If DContactFile.Visible Then                      '切結書
            If DContactFile.PostedFile.FileName <> "" Then     '判斷有檔案上傳
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
        SQl = SQl + " '" + DOFormNo.Text + "', "         '表單
        SQl = SQl + " '" + DOFormSno.Text + "', "        '單號
        SQl = SQl + " '" + "" + "', "            'QC-L/T

        FileName = ""
        If DSAttachFile.Visible Then                      '樣品-其他附件
            If DSAttachFile.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "SAttachFile" & "-" & UploadDateTime & "-" & Right(DSAttachFile.PostedFile.FileName, InStr(StrReverse(DSAttachFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "SAttachFile" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSAttachFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DSAttachFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        FileName = ""
        If DQAttachFile1.Visible Then                      '品質1-其他附件
            If DQAttachFile1.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QAttachFile1" & "-" & UploadDateTime & "-" & Right(DQAttachFile1.PostedFile.FileName, InStr(StrReverse(DQAttachFile1.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QAttachFile1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQAttachFile1.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQAttachFile1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        FileName = ""
        If DQAttachFile2.Visible Then                      '品質2-其他附件
            If DQAttachFile2.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QAttachFile2" & "-" & UploadDateTime & "-" & Right(DQAttachFile2.PostedFile.FileName, InStr(StrReverse(DQAttachFile2.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QAttachFile2" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQAttachFile2.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQAttachFile2.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '作成者
        SQl = SQl + " '" + NowDateTime + "', "       '作成時間
        SQl = SQl + " '" + "" + "', "                       '修改者
        SQl = SQl + " '" + NowDateTime + "' "       '修改時間
        SQl = SQl + " ) "
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDBCommand1.ExecuteNonQuery()
        '
        'Modify-Start by Joy 09-10-03
        'Sample Table處理
        Dim i As Integer
        For i = 1 To 3
            SQl = "Insert into F_ManufCOSample "
            SQl = SQl + "( "
            SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, SeqNo, "
            SQl = SQl + "Spec, Color, Qty, "
            SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "
            SQl = SQl + ")  "
            SQl = SQl + "VALUES( "
            SQl = SQl + " '0', "
            SQl = SQl + " '" + NowDateTime + "', "
            SQl = SQl + " '000003', "
            SQl = SQl + " '" + CStr(NewFormSno) + "', "
            SQl = SQl + " '" + CStr(i) + "', "
            Select Case i
                Case Is = 1
                    SQl = SQl + " '" + YKK.ReplaceString(DSampleA1.Text) + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DSampleA2.Text) + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DSampleA3.Text) + "', "
                Case Is = 2
                    SQl = SQl + " '" + YKK.ReplaceString(DSampleB1.Text) + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DSampleB2.Text) + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DSampleB3.Text) + "', "
                Case Is = 3
                    SQl = SQl + " '" + YKK.ReplaceString(DSampleC1.Text) + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DSampleC2.Text) + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DSampleC3.Text) + "', "
            End Select
            SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '作成者
            SQl = SQl + " '" + NowDateTime + "', "       '作成時間
            SQl = SQl + " '" + "" + "', "                       '修改者
            SQl = SQl + " '" + NowDateTime + "' "       '修改時間
            SQl = SQl + " ) "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQl
            OleDBCommand1.ExecuteNonQuery()
        Next
        'Price Table處理
        For i = 1 To 10
            SQl = "Insert into F_ManufCOPrice "
            SQl = SQl + "( "
            SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, SeqNo, "
            SQl = SQl + "Spec, Currency, Price, "
            SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "
            SQl = SQl + ")  "
            SQl = SQl + "VALUES( "
            SQl = SQl + " '0', "
            SQl = SQl + " '" + NowDateTime + "', "
            SQl = SQl + " '000003', "
            SQl = SQl + " '" + CStr(NewFormSno) + "', "
            SQl = SQl + " '" + CStr(i) + "', "
            Select Case i
                Case Is = 1
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceA1.Text) + "', "
                    SQl = SQl + " '" + DPriceA2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceA3.Text) + "', "
                Case Is = 2
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceB1.Text) + "', "
                    SQl = SQl + " '" + DPriceB2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceB3.Text) + "', "
                Case Is = 3
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceC1.Text) + "', "
                    SQl = SQl + " '" + DPriceC2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceC3.Text) + "', "
                Case Is = 4
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceD1.Text) + "', "
                    SQl = SQl + " '" + DPriceD2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceD3.Text) + "', "
                Case Is = 5
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceE1.Text) + "', "
                    SQl = SQl + " '" + DPriceE2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceE3.Text) + "', "
                Case Is = 6
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceF1.Text) + "', "
                    SQl = SQl + " '" + DPriceF2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceF3.Text) + "', "
                Case Is = 7
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceG1.Text) + "', "
                    SQl = SQl + " '" + DPriceG2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceG3.Text) + "', "
                Case Is = 8
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceH1.Text) + "', "
                    SQl = SQl + " '" + DPriceH2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceH3.Text) + "', "
                Case Is = 9
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceI1.Text) + "', "
                    SQl = SQl + " '" + DPriceI2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceI3.Text) + "', "
                Case Is = 10
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceJ1.Text) + "', "
                    SQl = SQl + " '" + DPriceJ2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceJ3.Text) + "', "
            End Select
            SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '作成者
            SQl = SQl + " '" + NowDateTime + "', "       '作成時間
            SQl = SQl + " '" + "" + "', "                       '修改者
            SQl = SQl + " '" + NowDateTime + "' "       '修改時間
            SQl = SQl + " ) "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQl
            OleDBCommand1.ExecuteNonQuery()
        Next
        'QA1 Table處理
        For i = 1 To 4
            SQl = "Insert into F_ManufCOQA1 "
            SQl = SQl + "( "
            SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, SeqNo, "
            SQl = SQl + "Date, Spec, Assembler, Surface, GenTani, Kyoudo, Nyuryoku, Kensin, Water, Dry, Yellow, Mityaku1,  Mityaku2, CPSC, "
            SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "
            SQl = SQl + ")  "
            SQl = SQl + "VALUES( "
            SQl = SQl + " '0', "
            SQl = SQl + " '" + NowDateTime + "', "
            SQl = SQl + " '000003', "
            SQl = SQl + " '" + CStr(NewFormSno) + "', "
            SQl = SQl + " '" + CStr(i) + "', "
            Select Case i
                Case Is = 1
                    SQl = SQl + " '" + YKK.ReplaceString(DQAA1.Text) + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DQAA2.Text) + "', "
                    SQl = SQl + " '" + DQAA3.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DQAA4.Text) + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DQAA5.Text) + "', "
                    SQl = SQl + " '" + DQAA6.SelectedValue + "', "
                    SQl = SQl + " '" + DQAA7.SelectedValue + "', "
                    SQl = SQl + " '" + DQAA8.SelectedValue + "', "
                    SQl = SQl + " '" + DQAA9.SelectedValue + "', "
                    SQl = SQl + " '" + DQAA10.SelectedValue + "', "
                    SQl = SQl + " '" + DQAA11.SelectedValue + "', "
                    SQl = SQl + " '" + DQAA12.SelectedValue + "', "
                    SQl = SQl + " '" + DQAA13.SelectedValue + "', "
                    SQl = SQl + " '" + DQAA14.SelectedValue + "', "
                Case Is = 2
                    SQl = SQl + " '" + YKK.ReplaceString(DQAB1.Text) + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DQAB2.Text) + "', "
                    SQl = SQl + " '" + DQAB3.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DQAB4.Text) + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DQAB5.Text) + "', "
                    SQl = SQl + " '" + DQAB6.SelectedValue + "', "
                    SQl = SQl + " '" + DQAB7.SelectedValue + "', "
                    SQl = SQl + " '" + DQAB8.SelectedValue + "', "
                    SQl = SQl + " '" + DQAB9.SelectedValue + "', "
                    SQl = SQl + " '" + DQAB10.SelectedValue + "', "
                    SQl = SQl + " '" + DQAB11.SelectedValue + "', "
                    SQl = SQl + " '" + DQAB12.SelectedValue + "', "
                    SQl = SQl + " '" + DQAB13.SelectedValue + "', "
                    SQl = SQl + " '" + DQAB14.SelectedValue + "', "
                Case Is = 3
                    SQl = SQl + " '" + YKK.ReplaceString(DQAC1.Text) + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DQAC2.Text) + "', "
                    SQl = SQl + " '" + DQAC3.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DQAC4.Text) + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DQAC5.Text) + "', "
                    SQl = SQl + " '" + DQAC6.SelectedValue + "', "
                    SQl = SQl + " '" + DQAC7.SelectedValue + "', "
                    SQl = SQl + " '" + DQAC8.SelectedValue + "', "
                    SQl = SQl + " '" + DQAC9.SelectedValue + "', "
                    SQl = SQl + " '" + DQAC10.SelectedValue + "', "
                    SQl = SQl + " '" + DQAC11.SelectedValue + "', "
                    SQl = SQl + " '" + DQAC12.SelectedValue + "', "
                    SQl = SQl + " '" + DQAC13.SelectedValue + "', "
                    SQl = SQl + " '" + DQAC14.SelectedValue + "', "
                Case Is = 4
                    SQl = SQl + " '" + YKK.ReplaceString(DQAD1.Text) + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DQAD2.Text) + "', "
                    SQl = SQl + " '" + DQAD3.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DQAD4.Text) + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DQAD5.Text) + "', "
                    SQl = SQl + " '" + DQAD6.SelectedValue + "', "
                    SQl = SQl + " '" + DQAD7.SelectedValue + "', "
                    SQl = SQl + " '" + DQAD8.SelectedValue + "', "
                    SQl = SQl + " '" + DQAD9.SelectedValue + "', "
                    SQl = SQl + " '" + DQAD10.SelectedValue + "', "
                    SQl = SQl + " '" + DQAD11.SelectedValue + "', "
                    SQl = SQl + " '" + DQAD12.SelectedValue + "', "
                    SQl = SQl + " '" + DQAD13.SelectedValue + "', "
                    SQl = SQl + " '" + DQAD14.SelectedValue + "', "
            End Select
            SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '作成者
            SQl = SQl + " '" + NowDateTime + "', "       '作成時間
            SQl = SQl + " '" + "" + "', "                       '修改者
            SQl = SQl + " '" + NowDateTime + "' "       '修改時間
            SQl = SQl + " ) "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQl
            OleDBCommand1.ExecuteNonQuery()
        Next
        'QA2 Table處理
        For i = 1 To 4
            SQl = "Insert into F_ManufCOQA2 "
            SQl = SQl + "( "
            SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, SeqNo, "
            SQl = SQl + "Date, QACheck, Remark, "
            SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "
            SQl = SQl + ")  "
            SQl = SQl + "VALUES( "
            SQl = SQl + " '0', "
            SQl = SQl + " '" + NowDateTime + "', "
            SQl = SQl + " '000003', "
            SQl = SQl + " '" + CStr(NewFormSno) + "', "
            SQl = SQl + " '" + CStr(i) + "', "
            Select Case i
                Case Is = 1
                    SQl = SQl + " '" + YKK.ReplaceString(DFQAA1.Text) + "', "
                    SQl = SQl + " '" + DFQAA2.SelectedValue + "', "
                    SQl = SQl + " N'" + YKK.ReplaceString(DFQAA3.Text) + "', "
                Case Is = 2
                    SQl = SQl + " '" + YKK.ReplaceString(DFQAB1.Text) + "', "
                    SQl = SQl + " '" + DFQAB2.SelectedValue + "', "
                    SQl = SQl + " N'" + YKK.ReplaceString(DFQAB3.Text) + "', "
                Case Is = 3
                    SQl = SQl + " '" + YKK.ReplaceString(DFQAC1.Text) + "', "
                    SQl = SQl + " '" + DFQAC2.SelectedValue + "', "
                    SQl = SQl + " N'" + YKK.ReplaceString(DFQAC3.Text) + "', "
                Case Is = 4
                    SQl = SQl + " '" + YKK.ReplaceString(DFQAD1.Text) + "', "
                    SQl = SQl + " '" + DFQAD2.SelectedValue + "', "
                    SQl = SQl + " N'" + YKK.ReplaceString(DFQAD3.Text) + "', "
            End Select
            SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '作成者
            SQl = SQl + " '" + NowDateTime + "', "       '作成時間
            SQl = SQl + " '" + "" + "', "                       '修改者
            SQl = SQl + " '" + NowDateTime + "' "       '修改時間
            SQl = SQl + " ) "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQl
            OleDBCommand1.ExecuteNonQuery()
        Next
        OleDbConnection1.Close()
        'Modify-end by Joy
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("ManufInFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String
        Dim DBDataSet1 As New DataSet
        Dim RtnCode As Integer

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand


        OleDbConnection1.Open()
        SQl = "Update F_ManufInSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = '" & YKK.ReplaceString(DNo.Text) & "',"
        SQl = SQl + " Date = '" & DDate.Text & "',"
        SQl = SQl + " Division = '" & DDivision.SelectedValue & "',"
        SQl = SQl + " Person = '" & DPerson.SelectedValue & "',"
        SQl = SQl + " SliderCode = '" & YKK.ReplaceString(DSliderCode.Text) & "',"
        SQl = SQl + " SliderGRCode = '" & YKK.ReplaceString(DSliderGRCode.Text) & "',"
        SQl = SQl + " Spec = '" & YKK.ReplaceString(DSpec.Text) & "',"
        SQl = SQl + " Cpsc = '" + DCpsc.SelectedValue + "', "            'CPSC
        SQl = SQl + " MapNo = '" & YKK.ReplaceString(DMapNo.Text) & "',"

        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Map" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Map" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DMapFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " MapFile = N'" & FileName & "',"
            End If
        End If

        If DRefFile.Visible Then
            If DRefFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Ref" & "-" & UploadDateTime & "-" & Right(DRefFile.PostedFile.FileName, InStr(StrReverse(DRefFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Ref" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DRefFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DRefFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " RefFile = N'" & FileName & "',"
            End If
        End If

        SQl = SQl + " Level = '" & DLevel.SelectedValue & "',"
        SQl = SQl + " Assembler = N'" & YKK.ReplaceString(DAssembler.Text) & "',"
        SQl = SQl + " SliderType1 = '" & DSliderType1.SelectedValue & "',"
        SQl = SQl + " SliderType2 = '" & DSliderType2.SelectedValue & "',"
        SQl = SQl + " ManufPlace = '" & DManufPlace.SelectedValue & "',"
        SQl = SQl + " Material = '" & DMaterial.SelectedValue & "',"
        SQl = SQl + " MaterialOther = '" & "" & "',"
        SQl = SQl + " SellVendor = N'" & YKK.ReplaceString(DSellVendor.Text) & "',"
        SQl = SQl + " Buyer = N'" & DBuyer.SelectedValue & "',"

        If DConfirmFile.Visible Then
            If DConfirmFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Confirm" & "-" & UploadDateTime & "-" & Right(DConfirmFile.PostedFile.FileName, InStr(StrReverse(DConfirmFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Confirm" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DConfirmFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DConfirmFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " ConfirmFile = N'" & FileName & "',"
            End If
        End If

        If DAuthorizeFile.Visible Then
            If DAuthorizeFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Authorize" & "-" & UploadDateTime & "-" & Right(DAuthorizeFile.PostedFile.FileName, InStr(StrReverse(DAuthorizeFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Authorize" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DAuthorizeFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DAuthorizeFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " AuthorizeFile = N'" & FileName & "',"
            End If
        End If

        SQl = SQl + " DevReason = N'" & YKK.ReplaceString(DDevReason.Text) & "',"
        If DSampleA1.Text <> "" Then
            SQl = SQl + " Sample = '" & "1" & "',"
        Else
            SQl = SQl + " Sample = '" & "0" & "',"
        End If
        If DPriceA1.Text <> "" Then
            SQl = SQl + " Price = '" & "1" & "',"
        Else
            SQl = SQl + " Price = '" & "0" & "',"
        End If
        SQl = SQl + " ArMoldFee = '" & YKK.ReplaceString(DArMoldFee.Text) & "',"
        SQl = SQl + " PurMold = '" & YKK.ReplaceString(DPurMold.Text) & "',"
        SQl = SQl + " PullerPrice = '" & YKK.ReplaceString(DPullerPrice.Text) & "',"
        SQl = SQl + " Suppiler = N'" & DSuppiler.SelectedValue & "',"
        SQl = SQl + " MakeCAM = N'" & DMakeCAM.SelectedValue & "',"

        SQl = SQl + " MoldQty = '" & YKK.ReplaceString(DMoldQty.Text) & "',"
        SQl = SQl + " MoldPoint = '" & YKK.ReplaceString(DMoldPoint.Text) & "',"
        If DQAA1.Text <> "" Then
            SQl = SQl + " Quality1 = '" & "1" & "',"
        Else
            SQl = SQl + " Quality1 = '" & "0" & "',"
        End If
        If DFQAA1.Text <> "" Then
            SQl = SQl + " Quality2 = '" & "1" & "',"
        Else
            SQl = SQl + " Quality2 = '" & "0" & "',"
        End If

        If DQAFile.Visible Then
            If DQAFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QA" & "-" & UploadDateTime & "-" & Right(DQAFile.PostedFile.FileName, InStr(StrReverse(DQAFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QA" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQAFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQAFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QAFile = N'" & FileName & "',"
            End If
        End If

        If DSampleFile.Visible Then
            If DSampleFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Sample" & "-" & UploadDateTime & "-" & Right(DSampleFile.PostedFile.FileName, InStr(StrReverse(DSampleFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Sample" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSampleFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DSampleFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " SampleFile = N'" & FileName & "',"
            End If
        End If

        If DContactFile.Visible Then
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
        SQl = SQl + " QCLT = '" & "" & "',"

        If DSAttachFile.Visible Then
            If DSAttachFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "SAttachFile" & "-" & UploadDateTime & "-" & Right(DSAttachFile.PostedFile.FileName, InStr(StrReverse(DSAttachFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "SAttachFile" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSAttachFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DSAttachFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " SAttachFile = N'" & FileName & "',"
            End If
        End If

        If DQAttachFile1.Visible Then
            If DQAttachFile1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QAttachFile1" & "-" & UploadDateTime & "-" & Right(DQAttachFile1.PostedFile.FileName, InStr(StrReverse(DQAttachFile1.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QAttachFile1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQAttachFile1.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQAttachFile1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QAttachFile1 = N'" & FileName & "',"
            End If
        End If

        If DQAttachFile2.Visible Then
            If DQAttachFile2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QAttachFile2" & "-" & UploadDateTime & "-" & Right(DQAttachFile2.PostedFile.FileName, InStr(StrReverse(DQAttachFile2.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QAttachFile2" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQAttachFile2.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQAttachFile2.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QAttachFile2 = N'" & FileName & "',"
            End If
        End If

        SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
        SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
        SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDBCommand1.ExecuteNonQuery()
        '
        'Modify-Start by Joy 09-10-03
        'Sample Table處理
        Dim i As Integer
        For i = 1 To 3
            DBDataSet1.Clear()
            SQl = "Select * From F_ManufCOSample "
            SQl = SQl + " Where FormNo  =  '" & "000003" & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And SeqNo   =  '" & CStr(i) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQl, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "F_ManufCOSample")
            If DBDataSet1.Tables("F_ManufCOSample").Rows.Count > 0 Then
                SQl = "Update F_ManufCOSample Set "
                Select Case i
                    Case Is = 1
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DSampleA1.Text) + "', "
                        SQl = SQl + " Color = '" + YKK.ReplaceString(DSampleA2.Text) + "', "
                        SQl = SQl + " Qty = '" + YKK.ReplaceString(DSampleA3.Text) + "', "
                    Case Is = 2
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DSampleB1.Text) + "', "
                        SQl = SQl + " Color = '" + YKK.ReplaceString(DSampleB2.Text) + "', "
                        SQl = SQl + " Qty = '" + YKK.ReplaceString(DSampleB3.Text) + "', "
                    Case Is = 3
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DSampleC1.Text) + "', "
                        SQl = SQl + " Color = '" + YKK.ReplaceString(DSampleC2.Text) + "', "
                        SQl = SQl + " Qty = '" + YKK.ReplaceString(DSampleC3.Text) + "', "
                End Select
                SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
                SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
                SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
                SQl = SQl + " Where FormNo  =  '" & "000003" & "'"
                SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQl = SQl + "   And SeqNo   =  '" & CStr(i) & "'"
            Else
                SQl = "Insert into F_ManufCOSample "
                SQl = SQl + "( "
                SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, SeqNo, "
                SQl = SQl + "Spec, Color, Qty, "
                SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '0', "
                SQl = SQl + " '" + NowDateTime + "', "
                SQl = SQl + " '000003', "
                SQl = SQl + " '" + CStr(wFormSno) + "', "
                SQl = SQl + " '" + CStr(i) + "', "
                Select Case i
                    Case Is = 1
                        SQl = SQl + " '" + YKK.ReplaceString(DSampleA1.Text) + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DSampleA2.Text) + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DSampleA3.Text) + "', "
                    Case Is = 2
                        SQl = SQl + " '" + YKK.ReplaceString(DSampleB1.Text) + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DSampleB2.Text) + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DSampleB3.Text) + "', "
                    Case Is = 3
                        SQl = SQl + " '" + YKK.ReplaceString(DSampleC1.Text) + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DSampleC2.Text) + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DSampleC3.Text) + "', "
                End Select
                SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '作成者
                SQl = SQl + " '" + NowDateTime + "', "       '作成時間
                SQl = SQl + " '" + "" + "', "                       '修改者
                SQl = SQl + " '" + NowDateTime + "' "       '修改時間
                SQl = SQl + " ) "
            End If
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQl
            OleDBCommand1.ExecuteNonQuery()
        Next
        'Price Table處理
        For i = 1 To 10
            DBDataSet1.Clear()
            SQl = "Select * From F_ManufCOPrice "
            SQl = SQl + " Where FormNo  =  '" & "000003" & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And SeqNo   =  '" & CStr(i) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQl, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "F_ManufCOPrice")
            If DBDataSet1.Tables("F_ManufCOPrice").Rows.Count > 0 Then
                SQl = "Update F_ManufCOPrice Set "
                Select Case i
                    Case Is = 1
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceA1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceA2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceA3.Text) + "', "
                    Case Is = 2
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceB1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceB2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceB3.Text) + "', "
                    Case Is = 3
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceC1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceC2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceC3.Text) + "', "
                    Case Is = 4
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceD1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceD2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceD3.Text) + "', "
                    Case Is = 5
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceE1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceE2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceE3.Text) + "', "
                    Case Is = 6
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceF1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceF2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceF3.Text) + "', "
                    Case Is = 7
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceG1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceG2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceG3.Text) + "', "
                    Case Is = 8
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceH1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceH2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceH3.Text) + "', "
                    Case Is = 9
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceI1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceI2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceI3.Text) + "', "
                    Case Is = 10
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceJ1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceJ2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceJ3.Text) + "', "
                End Select
                SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
                SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
                SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
                SQl = SQl + " Where FormNo  =  '" & "000003" & "'"
                SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQl = SQl + "   And SeqNo   =  '" & CStr(i) & "'"
            Else
                SQl = "Insert into F_ManufCOPrice "
                SQl = SQl + "( "
                SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, SeqNo, "
                SQl = SQl + "Spec, Currency, Price, "
                SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '0', "
                SQl = SQl + " '" + NowDateTime + "', "
                SQl = SQl + " '000003', "
                SQl = SQl + " '" + CStr(wFormSno) + "', "
                SQl = SQl + " '" + CStr(i) + "', "
                Select Case i
                    Case Is = 1
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceA1.Text) + "', "
                        SQl = SQl + " '" + DPriceA2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceA3.Text) + "', "
                    Case Is = 2
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceB1.Text) + "', "
                        SQl = SQl + " '" + DPriceB2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceB3.Text) + "', "
                    Case Is = 3
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceC1.Text) + "', "
                        SQl = SQl + " '" + DPriceC2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceC3.Text) + "', "
                    Case Is = 4
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceD1.Text) + "', "
                        SQl = SQl + " '" + DPriceD2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceD3.Text) + "', "
                    Case Is = 5
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceE1.Text) + "', "
                        SQl = SQl + " '" + DPriceE2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceE3.Text) + "', "
                    Case Is = 6
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceF1.Text) + "', "
                        SQl = SQl + " '" + DPriceF2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceF3.Text) + "', "
                    Case Is = 7
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceG1.Text) + "', "
                        SQl = SQl + " '" + DPriceG2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceG3.Text) + "', "
                    Case Is = 8
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceH1.Text) + "', "
                        SQl = SQl + " '" + DPriceH2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceH3.Text) + "', "
                    Case Is = 9
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceI1.Text) + "', "
                        SQl = SQl + " '" + DPriceI2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceI3.Text) + "', "
                    Case Is = 10
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceJ1.Text) + "', "
                        SQl = SQl + " '" + DPriceJ2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceJ3.Text) + "', "
                End Select
                SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '作成者
                SQl = SQl + " '" + NowDateTime + "', "       '作成時間
                SQl = SQl + " '" + "" + "', "                       '修改者
                SQl = SQl + " '" + NowDateTime + "' "       '修改時間
                SQl = SQl + " ) "
            End If
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQl
            OleDBCommand1.ExecuteNonQuery()
        Next
        'QA1 Table處理
        For i = 1 To 4
            DBDataSet1.Clear()
            SQl = "Select * From F_ManufCOQA1 "
            SQl = SQl + " Where FormNo  =  '" & "000003" & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And SeqNo   =  '" & CStr(i) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQl, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "F_ManufCOQA1")
            If DBDataSet1.Tables("F_ManufCOQA1").Rows.Count > 0 Then
                SQl = "Update F_ManufCOQA1 Set "
                Select Case i
                    Case Is = 1
                        SQl = SQl + " Date ='" + YKK.ReplaceString(DQAA1.Text) + "', "
                        SQl = SQl + " Spec ='" + YKK.ReplaceString(DQAA2.Text) + "', "
                        SQl = SQl + " Assembler = '" + DQAA3.SelectedValue + "', "
                        SQl = SQl + " Surface = '" + YKK.ReplaceString(DQAA4.Text) + "', "
                        SQl = SQl + " GenTani = '" + YKK.ReplaceString(DQAA5.Text) + "', "
                        SQl = SQl + " Kyoudo = '" + DQAA6.SelectedValue + "', "
                        SQl = SQl + " Nyuryoku = '" + DQAA7.SelectedValue + "', "
                        SQl = SQl + " Kensin = '" + DQAA8.SelectedValue + "', "
                        SQl = SQl + " Water = '" + DQAA9.SelectedValue + "', "
                        SQl = SQl + " Dry = '" + DQAA10.SelectedValue + "', "
                        SQl = SQl + " Yellow = '" + DQAA11.SelectedValue + "', "
                        SQl = SQl + " Mityaku1 = '" + DQAA12.SelectedValue + "', "
                        SQl = SQl + " Mityaku2 = '" + DQAA13.SelectedValue + "', "
                        SQl = SQl + " CPSC = '" + DQAA14.SelectedValue + "', "
                    Case Is = 2
                        SQl = SQl + " Date ='" + YKK.ReplaceString(DQAB1.Text) + "', "
                        SQl = SQl + " Spec ='" + YKK.ReplaceString(DQAB2.Text) + "', "
                        SQl = SQl + " Assembler = '" + DQAB3.SelectedValue + "', "
                        SQl = SQl + " Surface = '" + YKK.ReplaceString(DQAB4.Text) + "', "
                        SQl = SQl + " GenTani = '" + YKK.ReplaceString(DQAB5.Text) + "', "
                        SQl = SQl + " Kyoudo = '" + DQAB6.SelectedValue + "', "
                        SQl = SQl + " Nyuryoku = '" + DQAB7.SelectedValue + "', "
                        SQl = SQl + " Kensin = '" + DQAB8.SelectedValue + "', "
                        SQl = SQl + " Water = '" + DQAB9.SelectedValue + "', "
                        SQl = SQl + " Dry = '" + DQAB10.SelectedValue + "', "
                        SQl = SQl + " Yellow = '" + DQAB11.SelectedValue + "', "
                        SQl = SQl + " Mityaku1 = '" + DQAB12.SelectedValue + "', "
                        SQl = SQl + " Mityaku2 = '" + DQAB13.SelectedValue + "', "
                        SQl = SQl + " CPSC = '" + DQAB14.SelectedValue + "', "
                    Case Is = 3
                        SQl = SQl + " Date ='" + YKK.ReplaceString(DQAC1.Text) + "', "
                        SQl = SQl + " Spec ='" + YKK.ReplaceString(DQAC2.Text) + "', "
                        SQl = SQl + " Assembler = '" + DQAC3.SelectedValue + "', "
                        SQl = SQl + " Surface = '" + YKK.ReplaceString(DQAC4.Text) + "', "
                        SQl = SQl + " GenTani = '" + YKK.ReplaceString(DQAC5.Text) + "', "
                        SQl = SQl + " Kyoudo = '" + DQAC6.SelectedValue + "', "
                        SQl = SQl + " Nyuryoku = '" + DQAC7.SelectedValue + "', "
                        SQl = SQl + " Kensin = '" + DQAC8.SelectedValue + "', "
                        SQl = SQl + " Water = '" + DQAC9.SelectedValue + "', "
                        SQl = SQl + " Dry = '" + DQAC10.SelectedValue + "', "
                        SQl = SQl + " Yellow = '" + DQAC11.SelectedValue + "', "
                        SQl = SQl + " Mityaku1 = '" + DQAC12.SelectedValue + "', "
                        SQl = SQl + " Mityaku2 = '" + DQAC13.SelectedValue + "', "
                        SQl = SQl + " CPSC = '" + DQAC14.SelectedValue + "', "
                    Case Is = 4
                        SQl = SQl + " Date ='" + YKK.ReplaceString(DQAD1.Text) + "', "
                        SQl = SQl + " Spec ='" + YKK.ReplaceString(DQAD2.Text) + "', "
                        SQl = SQl + " Assembler = '" + DQAD3.SelectedValue + "', "
                        SQl = SQl + " Surface = '" + YKK.ReplaceString(DQAD4.Text) + "', "
                        SQl = SQl + " GenTani = '" + YKK.ReplaceString(DQAD5.Text) + "', "
                        SQl = SQl + " Kyoudo = '" + DQAD6.SelectedValue + "', "
                        SQl = SQl + " Nyuryoku = '" + DQAD7.SelectedValue + "', "
                        SQl = SQl + " Kensin = '" + DQAD8.SelectedValue + "', "
                        SQl = SQl + " Water = '" + DQAD9.SelectedValue + "', "
                        SQl = SQl + " Dry = '" + DQAD10.SelectedValue + "', "
                        SQl = SQl + " Yellow = '" + DQAD11.SelectedValue + "', "
                        SQl = SQl + " Mityaku1 = '" + DQAD12.SelectedValue + "', "
                        SQl = SQl + " Mityaku2 = '" + DQAD13.SelectedValue + "', "
                        SQl = SQl + " CPSC = '" + DQAD14.SelectedValue + "', "
                End Select
                SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
                SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
                SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
                SQl = SQl + " Where FormNo  =  '" & "000003" & "'"
                SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQl = SQl + "   And SeqNo   =  '" & CStr(i) & "'"
            Else
                SQl = "Insert into F_ManufCOQA1 "
                SQl = SQl + "( "
                SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, SeqNo, "
                SQl = SQl + "Date, Spec, Assembler, Surface, GenTani, Kyoudo, Nyuryoku, Kensin, Water, Dry, Yellow, Mityaku1,  Mityaku2, CPSC, "
                SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '0', "
                SQl = SQl + " '" + NowDateTime + "', "
                SQl = SQl + " '000003', "
                SQl = SQl + " '" + CStr(wFormSno) + "', "
                SQl = SQl + " '" + CStr(i) + "', "
                Select Case i
                    Case Is = 1
                        SQl = SQl + " '" + YKK.ReplaceString(DQAA1.Text) + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DQAA2.Text) + "', "
                        SQl = SQl + " '" + DQAA3.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DQAA4.Text) + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DQAA5.Text) + "', "
                        SQl = SQl + " '" + DQAA6.SelectedValue + "', "
                        SQl = SQl + " '" + DQAA7.SelectedValue + "', "
                        SQl = SQl + " '" + DQAA8.SelectedValue + "', "
                        SQl = SQl + " '" + DQAA9.SelectedValue + "', "
                        SQl = SQl + " '" + DQAA10.SelectedValue + "', "
                        SQl = SQl + " '" + DQAA11.SelectedValue + "', "
                        SQl = SQl + " '" + DQAA12.SelectedValue + "', "
                        SQl = SQl + " '" + DQAA13.SelectedValue + "', "
                        SQl = SQl + " '" + DQAA14.SelectedValue + "', "
                    Case Is = 2
                        SQl = SQl + " '" + YKK.ReplaceString(DQAB1.Text) + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DQAB2.Text) + "', "
                        SQl = SQl + " '" + DQAB3.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DQAB4.Text) + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DQAB5.Text) + "', "
                        SQl = SQl + " '" + DQAB6.SelectedValue + "', "
                        SQl = SQl + " '" + DQAB7.SelectedValue + "', "
                        SQl = SQl + " '" + DQAB8.SelectedValue + "', "
                        SQl = SQl + " '" + DQAB9.SelectedValue + "', "
                        SQl = SQl + " '" + DQAB10.SelectedValue + "', "
                        SQl = SQl + " '" + DQAB11.SelectedValue + "', "
                        SQl = SQl + " '" + DQAB12.SelectedValue + "', "
                        SQl = SQl + " '" + DQAB13.SelectedValue + "', "
                        SQl = SQl + " '" + DQAB14.SelectedValue + "', "
                    Case Is = 3
                        SQl = SQl + " '" + YKK.ReplaceString(DQAC1.Text) + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DQAC2.Text) + "', "
                        SQl = SQl + " '" + DQAC3.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DQAC4.Text) + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DQAC5.Text) + "', "
                        SQl = SQl + " '" + DQAC6.SelectedValue + "', "
                        SQl = SQl + " '" + DQAC7.SelectedValue + "', "
                        SQl = SQl + " '" + DQAC8.SelectedValue + "', "
                        SQl = SQl + " '" + DQAC9.SelectedValue + "', "
                        SQl = SQl + " '" + DQAC10.SelectedValue + "', "
                        SQl = SQl + " '" + DQAC11.SelectedValue + "', "
                        SQl = SQl + " '" + DQAC12.SelectedValue + "', "
                        SQl = SQl + " '" + DQAC13.SelectedValue + "', "
                        SQl = SQl + " '" + DQAC14.SelectedValue + "', "
                    Case Is = 4
                        SQl = SQl + " '" + YKK.ReplaceString(DQAD1.Text) + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DQAD2.Text) + "', "
                        SQl = SQl + " '" + DQAD3.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DQAD4.Text) + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DQAD5.Text) + "', "
                        SQl = SQl + " '" + DQAD6.SelectedValue + "', "
                        SQl = SQl + " '" + DQAD7.SelectedValue + "', "
                        SQl = SQl + " '" + DQAD8.SelectedValue + "', "
                        SQl = SQl + " '" + DQAD9.SelectedValue + "', "
                        SQl = SQl + " '" + DQAD10.SelectedValue + "', "
                        SQl = SQl + " '" + DQAD11.SelectedValue + "', "
                        SQl = SQl + " '" + DQAD12.SelectedValue + "', "
                        SQl = SQl + " '" + DQAD13.SelectedValue + "', "
                        SQl = SQl + " '" + DQAD14.SelectedValue + "', "
                End Select
                SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '作成者
                SQl = SQl + " '" + NowDateTime + "', "       '作成時間
                SQl = SQl + " '" + "" + "', "                       '修改者
                SQl = SQl + " '" + NowDateTime + "' "       '修改時間
                SQl = SQl + " ) "
            End If
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQl
            OleDBCommand1.ExecuteNonQuery()
        Next
        'QA2 Table處理
        For i = 1 To 4
            DBDataSet1.Clear()
            SQl = "Select * From F_ManufCOQA2 "
            SQl = SQl + " Where FormNo  =  '" & "000003" & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And SeqNo   =  '" & CStr(i) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQl, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "F_ManufCOQA2")
            If DBDataSet1.Tables("F_ManufCOQA2").Rows.Count > 0 Then
                SQl = "Update F_ManufCOQA2 Set "
                Select Case i
                    Case Is = 1
                        SQl = SQl + " Date = '" + YKK.ReplaceString(DFQAA1.Text) + "', "
                        SQl = SQl + " QACheck = '" + DFQAA2.SelectedValue + "', "
                        SQl = SQl + " Remark = N'" + YKK.ReplaceString(DFQAA3.Text) + "', "
                    Case Is = 2
                        SQl = SQl + " Date = '" + YKK.ReplaceString(DFQAB1.Text) + "', "
                        SQl = SQl + " QACheck = '" + DFQAB2.SelectedValue + "', "
                        SQl = SQl + " Remark = N'" + YKK.ReplaceString(DFQAB3.Text) + "', "
                    Case Is = 3
                        SQl = SQl + " Date = '" + YKK.ReplaceString(DFQAC1.Text) + "', "
                        SQl = SQl + " QACheck = '" + DFQAC2.SelectedValue + "', "
                        SQl = SQl + " Remark = N'" + YKK.ReplaceString(DFQAC3.Text) + "', "
                    Case Is = 4
                        SQl = SQl + " Date = '" + YKK.ReplaceString(DFQAD1.Text) + "', "
                        SQl = SQl + " QACheck = '" + DFQAD2.SelectedValue + "', "
                        SQl = SQl + " Remark = N'" + YKK.ReplaceString(DFQAD3.Text) + "', "
                End Select
                SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
                SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
                SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
                SQl = SQl + " Where FormNo  =  '" & "000003" & "'"
                SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQl = SQl + "   And SeqNo   =  '" & CStr(i) & "'"
            Else
                SQl = "Insert into F_ManufCOQA2 "
                SQl = SQl + "( "
                SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, SeqNo, "
                SQl = SQl + "Date, QACheck, Remark, "
                SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '0', "
                SQl = SQl + " '" + NowDateTime + "', "
                SQl = SQl + " '000003', "
                SQl = SQl + " '" + CStr(wFormSno) + "', "
                SQl = SQl + " '" + CStr(i) + "', "
                Select Case i
                    Case Is = 1
                        SQl = SQl + " '" + YKK.ReplaceString(DFQAA1.Text) + "', "
                        SQl = SQl + " '" + DFQAA2.SelectedValue + "', "
                        SQl = SQl + " N'" + YKK.ReplaceString(DFQAA3.Text) + "', "
                    Case Is = 2
                        SQl = SQl + " '" + YKK.ReplaceString(DFQAB1.Text) + "', "
                        SQl = SQl + " '" + DFQAB2.SelectedValue + "', "
                        SQl = SQl + " N'" + YKK.ReplaceString(DFQAB3.Text) + "', "
                    Case Is = 3
                        SQl = SQl + " '" + YKK.ReplaceString(DFQAC1.Text) + "', "
                        SQl = SQl + " '" + DFQAC2.SelectedValue + "', "
                        SQl = SQl + " N'" + YKK.ReplaceString(DFQAC3.Text) + "', "
                    Case Is = 4
                        SQl = SQl + " '" + YKK.ReplaceString(DFQAD1.Text) + "', "
                        SQl = SQl + " '" + DFQAD2.SelectedValue + "', "
                        SQl = SQl + " N'" + YKK.ReplaceString(DFQAD3.Text) + "', "
                End Select
                SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '作成者
                SQl = SQl + " '" + NowDateTime + "', "       '作成時間
                SQl = SQl + " '" + "" + "', "                       '修改者
                SQl = SQl + " '" + NowDateTime + "' "       '修改時間
                SQl = SQl + " ) "
            End If
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQl
            OleDBCommand1.ExecuteNonQuery()
        Next
        OleDbConnection1.Close()
        'Modify-end by Joy
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
        Dim wLevel As String = DLevel.SelectedValue '難易度

        URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                              "&pFormSno=" & wFormSno & _
                              "&pStep=" & wStep & _
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
        'CPSC
        If InputCheck = 0 Then
            If FindFieldInf("Cpsc") = 1 Then
                If DCpsc.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'Slider Group Code
        If InputCheck = 0 Then
            If FindFieldInf("SliderGRCode") = 1 Then
                If DSliderGRCode.Text = "" Then InputCheck = 1
            End If
        End If
        'Slider Code
        If InputCheck = 0 Then
            If FindFieldInf("SliderCode") = 1 Then
                If DSliderCode.Text = "" Then InputCheck = 1
            End If
        End If
        'Spec
        If InputCheck = 0 Then
            If FindFieldInf("Spec") = 1 Then
                If DSpec.Text = "" Then InputCheck = 1
            End If
        End If
        '圖號
        If InputCheck = 0 Then
            If FindFieldInf("MapNo") = 1 Then
                If DMapNo.Text = "" Then InputCheck = 1
            End If
        End If
        '原始表單號碼
        If InputCheck = 0 Then
            If FindFieldInf("OFormNo") = 1 Then
                If DOFormNo.Text = "" Then InputCheck = 1
            End If
        End If
        '原始單號
        If InputCheck = 0 Then
            If FindFieldInf("OFormSno") = 1 Then
                If DOFormSno.Text = "" Then InputCheck = 1
            End If
        End If
        '難易度
        If InputCheck = 0 Then
            If FindFieldInf("Level") = 1 Then
                If DLevel.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '組立判別
        If InputCheck = 0 Then
            If FindFieldInf("Assembler") = 1 Then
                If DAssembler.Text = "" Then InputCheck = 1
            End If
        End If
        '拉頭種類(內製.外注...)
        If InputCheck = 0 Then
            If FindFieldInf("SliderType1") = 1 Then
                If DSliderType1.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '拉頭種類(半成品.成品...)
        If InputCheck = 0 Then
            If FindFieldInf("SliderType2") = 1 Then
                If DSliderType2.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '生產地
        If InputCheck = 0 Then
            If FindFieldInf("ManufPlace") = 1 Then
                If DManufPlace.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '材質
        If InputCheck = 0 Then
            If FindFieldInf("Material") = 1 Then
                If DMaterial.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '委託廠商
        If InputCheck = 0 Then
            If FindFieldInf("SellVendor") = 1 Then
                If DSellVendor.Text = "" Then InputCheck = 1
            End If
        End If
        'Buyer
        If InputCheck = 0 Then
            If FindFieldInf("Buyer") = 1 Then
                If DBuyer.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '開發理由
        If InputCheck = 0 Then
            If FindFieldInf("DevReason") = 1 Then
                If DDevReason.Text = "" Then InputCheck = 1
            End If
        End If
        '應收模具費
        If InputCheck = 0 Then
            If FindFieldInf("ArMoldFee") = 1 Then
                If DArMoldFee.Text = "" Then InputCheck = 1
            End If
        End If
        '模具購入價
        If InputCheck = 0 Then
            If FindFieldInf("PurMold") = 1 Then
                If DPurMold.Text = "" Then InputCheck = 1
            End If
        End If
        '引手購入單價
        If InputCheck = 0 Then
            If FindFieldInf("PullerPrice") = 1 Then
                If DPullerPrice.Text = "" Then InputCheck = 1
            End If
        End If
        '外注商
        If InputCheck = 0 Then
            If FindFieldInf("Suppiler") = 1 Then
                If DSuppiler.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '製CAM
        If InputCheck = 0 Then
            If FindFieldInf("MakeCAM") = 1 Then
                If DMakeCAM.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '模型個取-模型
        If InputCheck = 0 Then
            If FindFieldInf("MoldQty") = 1 Then
                If DMoldQty.Text = "" Then InputCheck = 1
            End If
        End If
        '模型個取-穴取
        If InputCheck = 0 Then
            If FindFieldInf("MoldPoint") = 1 Then
                If DMoldPoint.Text = "" Then InputCheck = 1
            End If
        End If
        '樣品圖
        If InputCheck = 0 Then
            If FindFieldInf("SampleFile") = 1 Then
                If DSampleFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '圖樣
        If InputCheck = 0 Then
            If FindFieldInf("MapFile") = 1 Then
                If DMapFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '樣品-其他附件
        If InputCheck = 0 Then
            If FindFieldInf("SAttachFile") = 1 Then
                If DSAttachFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '品質1-其他附件
        If InputCheck = 0 Then
            If FindFieldInf("QAttachFile1") = 1 Then
                If DQAttachFile1.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '品質2-其他附件
        If InputCheck = 0 Then
            If FindFieldInf("QAttachFile2") = 1 Then
                If DQAttachFile2.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '其他附件
        If InputCheck = 0 Then
            If FindFieldInf("RefFile") = 1 Then
                If DRefFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '確認書
        If InputCheck = 0 Then
            If FindFieldInf("ConfirmFile") = 1 Then
                If DConfirmFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '授權書
        If InputCheck = 0 Then
            If FindFieldInf("AuthorizeFile") = 1 Then
                If DAuthorizeFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '切結書
        If InputCheck = 0 Then
            If FindFieldInf("ContactFile") = 1 Then
                If DContactFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '測試報告
        If InputCheck = 0 Then
            If FindFieldInf("QAFile") = 1 Then
                If DQAFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        'Sample
        If InputCheck = 0 Then
            If FindFieldInf("Sample") = 1 Then
                If DSampleA1.Text = "" Then InputCheck = 1
                If DSampleA2.Text = "" Then InputCheck = 1
                If DSampleA3.Text = "" Then InputCheck = 1
                If DSampleB1.Text = "" Then InputCheck = 1
                If DSampleB2.Text = "" Then InputCheck = 1
                If DSampleB3.Text = "" Then InputCheck = 1
                If DSampleC1.Text = "" Then InputCheck = 1
                If DSampleC2.Text = "" Then InputCheck = 1
                If DSampleC3.Text = "" Then InputCheck = 1
            End If
        End If
        'Price
        If InputCheck = 0 Then
            If FindFieldInf("Price") = 1 Then
                If DPriceA1.Text = "" Then InputCheck = 1
                If DPriceA2.SelectedValue = "" Then InputCheck = 1
                If DPriceA3.Text = "" Then InputCheck = 1
                If DPriceB1.Text = "" Then InputCheck = 1
                If DPriceB2.SelectedValue = "" Then InputCheck = 1
                If DPriceB3.Text = "" Then InputCheck = 1
                If DPriceC1.Text = "" Then InputCheck = 1
                If DPriceC2.SelectedValue = "" Then InputCheck = 1
                If DPriceC3.Text = "" Then InputCheck = 1
                If DPriceD1.Text = "" Then InputCheck = 1
                If DPriceD2.SelectedValue = "" Then InputCheck = 1
                If DPriceD3.Text = "" Then InputCheck = 1
                If DPriceE1.Text = "" Then InputCheck = 1
                If DPriceE2.SelectedValue = "" Then InputCheck = 1
                If DPriceE3.Text = "" Then InputCheck = 1
                If DPriceF1.Text = "" Then InputCheck = 1
                If DPriceF2.SelectedValue = "" Then InputCheck = 1
                If DPriceF3.Text = "" Then InputCheck = 1
                If DPriceG1.Text = "" Then InputCheck = 1
                If DPriceG2.SelectedValue = "" Then InputCheck = 1
                If DPriceG3.Text = "" Then InputCheck = 1
                If DPriceH1.Text = "" Then InputCheck = 1
                If DPriceH2.SelectedValue = "" Then InputCheck = 1
                If DPriceH3.Text = "" Then InputCheck = 1
                If DPriceI1.Text = "" Then InputCheck = 1
                If DPriceI2.SelectedValue = "" Then InputCheck = 1
                If DPriceI3.Text = "" Then InputCheck = 1
                If DPriceJ1.Text = "" Then InputCheck = 1
                If DPriceJ2.SelectedValue = "" Then InputCheck = 1
                If DPriceJ3.Text = "" Then InputCheck = 1
            End If
        End If
        'Quality1
        If InputCheck = 0 Then
            If FindFieldInf("Quality1") = 1 Then
                If DQAA1.Text = "" Then InputCheck = 1
                If DQAA2.Text = "" Then InputCheck = 1
                If DQAA3.SelectedValue = "" Then InputCheck = 1
                If DQAA4.Text = "" Then InputCheck = 1
                If DQAA5.Text = "" Then InputCheck = 1
                If DQAA6.SelectedValue = "" Then InputCheck = 1
                If DQAA7.SelectedValue = "" Then InputCheck = 1
                If DQAA8.SelectedValue = "" Then InputCheck = 1
                If DQAA9.SelectedValue = "" Then InputCheck = 1
                If DQAA10.SelectedValue = "" Then InputCheck = 1
                If DQAA11.SelectedValue = "" Then InputCheck = 1
                If DQAA12.SelectedValue = "" Then InputCheck = 1
                If DQAA13.SelectedValue = "" Then InputCheck = 1
                If DQAA14.SelectedValue = "" Then InputCheck = 1
                If DQAB1.Text = "" Then InputCheck = 1
                If DQAB2.Text = "" Then InputCheck = 1
                If DQAB3.SelectedValue = "" Then InputCheck = 1
                If DQAB4.Text = "" Then InputCheck = 1
                If DQAB5.Text = "" Then InputCheck = 1
                If DQAB6.SelectedValue = "" Then InputCheck = 1
                If DQAB7.SelectedValue = "" Then InputCheck = 1
                If DQAB8.SelectedValue = "" Then InputCheck = 1
                If DQAB9.SelectedValue = "" Then InputCheck = 1
                If DQAB10.SelectedValue = "" Then InputCheck = 1
                If DQAB11.SelectedValue = "" Then InputCheck = 1
                If DQAB12.SelectedValue = "" Then InputCheck = 1
                If DQAB13.SelectedValue = "" Then InputCheck = 1
                If DQAB14.SelectedValue = "" Then InputCheck = 1
                If DQAC1.Text = "" Then InputCheck = 1
                If DQAC2.Text = "" Then InputCheck = 1
                If DQAC3.SelectedValue = "" Then InputCheck = 1
                If DQAC4.Text = "" Then InputCheck = 1
                If DQAC5.Text = "" Then InputCheck = 1
                If DQAC6.SelectedValue = "" Then InputCheck = 1
                If DQAC7.SelectedValue = "" Then InputCheck = 1
                If DQAC8.SelectedValue = "" Then InputCheck = 1
                If DQAC9.SelectedValue = "" Then InputCheck = 1
                If DQAC10.SelectedValue = "" Then InputCheck = 1
                If DQAC11.SelectedValue = "" Then InputCheck = 1
                If DQAC12.SelectedValue = "" Then InputCheck = 1
                If DQAC13.SelectedValue = "" Then InputCheck = 1
                If DQAC14.SelectedValue = "" Then InputCheck = 1
                If DQAD1.Text = "" Then InputCheck = 1
                If DQAD2.Text = "" Then InputCheck = 1
                If DQAD3.SelectedValue = "" Then InputCheck = 1
                If DQAD4.Text = "" Then InputCheck = 1
                If DQAD5.Text = "" Then InputCheck = 1
                If DQAD6.SelectedValue = "" Then InputCheck = 1
                If DQAD7.SelectedValue = "" Then InputCheck = 1
                If DQAD8.SelectedValue = "" Then InputCheck = 1
                If DQAD9.SelectedValue = "" Then InputCheck = 1
                If DQAD10.SelectedValue = "" Then InputCheck = 1
                If DQAD11.SelectedValue = "" Then InputCheck = 1
                If DQAD12.SelectedValue = "" Then InputCheck = 1
                If DQAD13.SelectedValue = "" Then InputCheck = 1
                If DQAD14.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'Quality2
        If InputCheck = 0 Then
            If FindFieldInf("Quality2") = 1 Then
                If DFQAA1.Text = "" Then InputCheck = 1
                If DFQAA2.SelectedValue = "" Then InputCheck = 1
                If DFQAA3.Text = "" Then InputCheck = 1
                If DFQAB1.Text = "" Then InputCheck = 1
                If DFQAB2.SelectedValue = "" Then InputCheck = 1
                If DFQAB3.Text = "" Then InputCheck = 1
                If DFQAC1.Text = "" Then InputCheck = 1
                If DFQAC2.SelectedValue = "" Then InputCheck = 1
                If DFQAC3.Text = "" Then InputCheck = 1
                If DFQAD1.Text = "" Then InputCheck = 1
                If DFQAD2.SelectedValue = "" Then InputCheck = 1
                If DFQAD3.Text = "" Then InputCheck = 1
            End If
        End If
    End Function

End Class
