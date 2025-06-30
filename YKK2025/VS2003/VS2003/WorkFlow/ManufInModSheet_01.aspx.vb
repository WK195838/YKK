Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ManufInModSheet
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DManuaInSheet1 As System.Web.UI.WebControls.Image
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
    Protected WithEvents DPriceJ3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceJ2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceJ1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceI3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceI2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceI1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceH3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceH2 As System.Web.UI.WebControls.DropDownList
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
    Protected WithEvents DPriceA3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceA2 As System.Web.UI.WebControls.DropDownList
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
    Protected WithEvents DSuppiler As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPullerPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPurMold As System.Web.UI.WebControls.TextBox
    Protected WithEvents DArMoldFee As System.Web.UI.WebControls.TextBox
    Protected WithEvents LAuthorizeFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LConfirmFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
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
    Protected WithEvents DManuaInSheet2 As System.Web.UI.WebControls.Image
    Protected WithEvents LMapFile As System.Web.UI.WebControls.Image
    Protected WithEvents DMapFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DContactFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQAFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DSampleFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DAuthorizeFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DConfirmFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DRefFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents DSAttachFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents LSAttachFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DQAttachFile1 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents LQAttachFile1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DQAttachFile2 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents LQAttachFile2 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents Label25 As System.Web.UI.WebControls.Label
    Protected WithEvents Label24 As System.Web.UI.WebControls.Label
    Protected WithEvents Label23 As System.Web.UI.WebControls.Label
    Protected WithEvents Label22 As System.Web.UI.WebControls.Label
    Protected WithEvents Label21 As System.Web.UI.WebControls.Label
    Protected WithEvents Label20 As System.Web.UI.WebControls.Label
    Protected WithEvents Label19 As System.Web.UI.WebControls.Label
    Protected WithEvents Label18 As System.Web.UI.WebControls.Label
    Protected WithEvents Label17 As System.Web.UI.WebControls.Label
    Protected WithEvents Label16 As System.Web.UI.WebControls.Label
    Protected WithEvents Label15 As System.Web.UI.WebControls.Label
    Protected WithEvents Label14 As System.Web.UI.WebControls.Label
    Protected WithEvents DMakeCAM As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents DCpsc As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Label27 As System.Web.UI.WebControls.Label
    Protected WithEvents DQCLT As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label28 As System.Web.UI.WebControls.Label
    Protected WithEvents DManufFlow As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDevReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCustomerCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DLogo As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents Label4 As System.Web.UI.WebControls.Label
    Protected WithEvents Label5 As System.Web.UI.WebControls.Label
    Protected WithEvents Label6 As System.Web.UI.WebControls.Label
    Protected WithEvents Label7 As System.Web.UI.WebControls.Label
    Protected WithEvents Label8 As System.Web.UI.WebControls.Label
    Protected WithEvents Label9 As System.Web.UI.WebControls.Label
    Protected WithEvents Label10 As System.Web.UI.WebControls.Label
    Protected WithEvents Label11 As System.Web.UI.WebControls.Label
    Protected WithEvents Label12 As System.Web.UI.WebControls.Label
    Protected WithEvents Label26 As System.Web.UI.WebControls.Label
    Protected WithEvents Label13 As System.Web.UI.WebControls.Label
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents DQAA2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAA1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAA3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAA5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAA6 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA8 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA9 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA10 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA14 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAB2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAB3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAB5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAB6 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB8 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB9 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB10 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB14 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAC2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAC3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAC5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAC6 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC8 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC9 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC10 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC14 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAD2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAD3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAD5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAD6 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD8 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD9 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD10 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD14 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA11 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA12 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA13 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB11 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC11 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD11 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB12 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC12 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD12 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB13 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC13 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD13 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA15 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Label29 As System.Web.UI.WebControls.Label
    Protected WithEvents DQAB15 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC15 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD15 As System.Web.UI.WebControls.DropDownList

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
    Dim wOFormNo As String          '原表單號碼
    Dim wOFormSno As Integer        '原表單流水號
    Dim rFormNo As String           '傳回表單號碼
    Dim rFormSno As Integer         '傳回表單流水號
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
    Dim wDepo As String = "CL1"      '中壢行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
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
            'If wFormSno > 0 And wStep > 2 Then    '判斷是否有資料
            If wFormSno > 0 Then    '判斷是否有資料
                ShowFormData()      '顯示表單資料
            Else
                ShowOriFormData()   '顯示原表單資料
            End If
            SetPopupFunction()      '設定彈出視窗事件
        Else
            ShowSheetField("Post")  '表單欄位顯示及欄位輸入檢查
            'ShowMessage()           '上傳資料檢查及顯示訊息
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
        wOFormNo = Request.QueryString("pOFormNo")    '原表單號碼
        wOFormSno = Request.QueryString("pOFormSno")  '原表單流水號
        wStep = Request.QueryString("pStep")        '工程代碼
        wSeqNo = Request.QueryString("pSeqNo")      '序號
        wApplyID = Request.QueryString("pApplyID")  '申請者ID
        wAgentID = Request.QueryString("pAgentID")  '被代理人ID
        Response.Cookies("MapNo").Value = ""        '圖號, MapPicker使用

        Response.Cookies("PGM").Value = "ManufInModSheet_01.aspx"
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

        SQL = "Select * From F_ManufInModSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ManufInModSheet")
        If DBDataSet1.Tables("F_ManufInModSheet").Rows.Count > 0 Then
            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MapFile") <> "" Then          '圖樣
                LMapFile.ImageUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MapFile")
            Else
                LMapFile.Visible = False
            End If

            DNo.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Date")                   '日期
            SetFieldData("Division", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Division"))  '部門
            SetFieldData("Person", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Person"))      '擔當
            SetFieldData("Cpsc", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Cpsc"))          'Cpsc
            SetFieldData("Logo", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Logo"))          'Logo

            DSliderCode.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SliderCode")       'Slider Code
            DSliderGRCode.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SliderGRCode")   'Slider Group Code
            DSpec.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Spec")                   '規格
            DMapNo.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MapNo")                 '圖號
            DOFormNo.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("OFormNo")             '原表單
            DOFormSno.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("OFormSno")           '原單號

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MapNo") <> "" Then                 '圖號
                LMapNo.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MapNo")
                If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("OFormNo") = "000001" Then
                    LMapNo.NavigateUrl = "MapSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("OFormNo") & _
                                                          "&pFormSno=" & CStr(DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("OFormSno"))
                Else
                    LMapNo.NavigateUrl = "MapSheetMod_02.aspx?pFormNo=" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("OFormNo") & _
                                                          "&pFormSno=" & CStr(DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("OFormSno"))
                End If
            Else
                LMapNo.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SAttachFile") <> "" Then       '樣品-其他附件
                LSAttachFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SAttachFile")
            Else
                LSAttachFile.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("QAttachFile1") <> "" Then      '品質1-其他附件
                LQAttachFile1.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("QAttachFile1")
            Else
                LQAttachFile1.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("QAttachFile2") <> "" Then       '樣品2-其他附件
                LQAttachFile2.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("QAttachFile2")
            Else
                LQAttachFile2.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("RefFile") <> "" Then          '其他附件 
                LRefFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("RefFile")
            Else
                LRefFile.Visible = False
            End If

            SetFieldData("Level", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Level"))                '難易度
            DAssembler.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Assembler")                 '組立
            SetFieldData("SliderType1", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SliderType1"))    '拉頭種類1
            SetFieldData("SliderType2", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SliderType2"))    '拉頭種類2
            SetFieldData("ManufPlace", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ManufPlace"))      '生產地
            SetFieldData("MakeCAM", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MakeCAM"))    'Make CAM

            SetFieldData("Material", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Material"))              '材質
            DBuyer.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Buyer")             'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SellVendor")   '委託廠商
            DCustomerCode.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("CustomerCode")   'Customer Code

            DQCLT.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("QCLT")               'QC-L/T

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ConfirmFile") <> "" Then       '確認書
                LConfirmFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ConfirmFile")
            Else
                LConfirmFile.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("AuthorizeFile") <> "" Then     '授權書
                LAuthorizeFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("AuthorizeFile")
            Else
                LAuthorizeFile.Visible = False
            End If

            DDevReason.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("DevReason")     '開發理由
            DManufFlow.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ManufFlow")     '

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Sample") = 1 Then              '樣品
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

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SampleFile") <> "" Then        '樣品檔
                LSampleFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SampleFile")
            Else
                LSampleFile.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Price") = 1 Then               '單價
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

            DArMoldFee.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ArMoldFee")     '應收模具費
            DPurMold.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("PurMold")         '模具購入費
            DPullerPrice.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("PullerPrice") '引手購入價
            DSuppiler.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Suppiler")       '外注商
            DMoldQty.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MoldQty")         '模型
            DMoldPoint.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MoldPoint")     '穴取

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Quality1") = 1 Then            '品質1
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
                            SetFieldData("DQAA15", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("EDX"))
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
                            SetFieldData("DQAB15", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("EDX"))
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
                            SetFieldData("DQAC15", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("EDX"))
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
                            SetFieldData("DQAD15", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("EDX"))
                    End Select
                Next
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Quality2") = 1 Then            '品質2
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

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("QAFile") <> "" Then      '測試報告
                LQAFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("QAFile")
            Else
                LQAFile.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ContactFile") <> "" Then      '切結書
                LContactFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ContactFile")
            Else
                LContactFile.Visible = False
            End If

            DFormSno.Text = "單號：" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("FormSno")       '單號
        End If
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowOriFormData()
        Dim Path As String
        Path = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ManufInFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_ManufInSheet "
        SQL = SQL & " Where Sts = 1 "
        SQL = SQL & "   And FormNo =  '" & wOFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ManufInSheet")
        If DBDataSet1.Tables("F_ManufInSheet").Rows.Count > 0 Then
            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MapFile") <> "" Then          '圖樣
                LMapFile.ImageUrl = Path & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MapFile")
            Else
                LMapFile.Visible = False
            End If

            DNo.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("No")                       'No
            SetFieldData("Division", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Division"))  '部門
            SetFieldData("Person", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Person"))      '擔當
            SetFieldData("Cpsc", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Cpsc"))          'Cpsc
            SetFieldData("Logo", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Logo"))          'Logo

            DSliderCode.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SliderCode")       'Slider Code
            DSliderGRCode.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SliderGRCode")   'Slider Group Code
            DSpec.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Spec")                   '規格
            DMapNo.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MapNo")                 '圖號
            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MapNo") <> "" Then                 '圖號
                LMapNo.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MapNo")
                If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("OFormNo") = "000001" Then
                    LMapNo.NavigateUrl = "MapSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("OFormNo") & _
                                                          "&pFormSno=" & CStr(DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("OFormSno"))
                Else
                    LMapNo.NavigateUrl = "MapSheetMod_02.aspx?pFormNo=" & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("OFormNo") & _
                                                          "&pFormSno=" & CStr(DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("OFormSno"))
                End If
            Else
                LMapNo.Visible = False
            End If
            DOFormNo.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("OFormNo")             '原表單
            DOFormSno.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("OFormSno")           '原單號
            DBuyer.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Buyer")             'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SellVendor")   '委託廠商
            DCustomerCode.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("CustomerCode")   'Customer Code

            DQCLT.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("QCLT")               'QC-L/T

            SetFieldData("Level", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Level"))                '難易度
            DAssembler.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Assembler")                 '組立
            SetFieldData("SliderType1", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SliderType1"))    '拉頭種類1
            SetFieldData("SliderType2", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SliderType2"))    '拉頭種類2
            SetFieldData("ManufPlace", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("ManufPlace"))      '生產地
            SetFieldData("Material", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Material"))              '材質
            DSuppiler.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Suppiler")       '外注商

            '----------------------------------------------------------------------------
            'DDate.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Date")                   '日期
            DDate.Text = CStr(DateTime.Now.Today)
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

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("QAttachFile2") <> "" Then       '樣品2-品質分析依賴書
                LQAttachFile2.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("QAttachFile2")
            Else
                LQAttachFile2.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("RefFile") <> "" Then          '其他附件 
                LRefFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("RefFile")
            Else
                LRefFile.Visible = False
            End If

            SetFieldData("MakeCAM", DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MakeCAM"))    'Make CAM

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
            DManufFlow.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("ManufFlow")     '

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Sample") = 1 Then              '樣品
                Dim DBTable1 As DataTable
                SQL = "Select * From F_ManufCOSample "
                SQL = SQL & " Where FormNo =  '" & wOFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
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
                SQL = SQL & " Where FormNo =  '" & wOFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
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
            DMoldQty.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MoldQty")         '模型
            DMoldPoint.Text = DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MoldPoint")     '穴取

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Quality1") = 1 Then            '品質1
                Dim DBTable1 As DataTable
                SQL = "Select * From F_ManufCOQA1 "
                SQL = SQL & " Where FormNo =  '" & wOFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
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
                            SetFieldData("DQAA15", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("EDX"))
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
                            SetFieldData("DQAB15", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("EDX"))
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
                            SetFieldData("DQAC15", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("EDX"))
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
                            SetFieldData("DQAD15", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("EDX"))
                    End Select
                Next
            End If

            If DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Quality2") = 1 Then            '品質2
                Dim DBTable1 As DataTable
                SQL = "Select * From F_ManufCOQA2 "
                SQL = SQL & " Where FormNo =  '" & wOFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "ManufCOQA2")
                DBTable1 = DBDataSet1.Tables("ManufCOQA2")
                For i = 0 To DBTable1.Rows.Count - 1
                    Select Case DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Seqno")
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

            DFormSno.Text = "單號：" & CStr(wOFormSno)       '單號
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
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        'If wFormSno > 0 And wStep > 2 Then      '判斷是否[簽核]
        If wFormSno > 0 Then    '判斷是否有資料
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
                '按鈕位置
                BSAVE.Style.Add("Top", Top)    '儲存按鈕
                BOK.Style.Add("Top", Top)      '閱讀完成
                DFormSno.Style.Add("Top", Top) '單號
            End If
        Else
            'Sheet顯示
            DManuaInSheet1.Visible = True   '表單Sheet-1
            DManuaInSheet2.Visible = True   '表單Sheet-2
            '連結顯示---需再修改
            LMapNo.Visible = True           'MapNo
            LMapFile.Visible = True         '圖樣
            LRefFile.Visible = True         '其他附件
            LConfirmFile.Visible = True     '確認書
            LAuthorizeFile.Visible = True   '授權書
            LSampleFile.Visible = True      '樣品圖
            LQAFile.Visible = True          '測試報告
            LContactFile.Visible = True     '切結書
            '連結隱藏
            'LMapNo.Visible = False           'Map No
            'LMapFile.Visible = False         '圖樣
            'LSAttachFile.Visible = False     '樣品-其他附件
            'LQAttachFile1.Visible = False    '品質1-其他附件
            'LQAttachFile2.Visible = False    '品質2-其他附件
            'LRefFile.Visible = False         '其他附件
            'LConfirmFile.Visible = False     '確認書
            'LAuthorizeFile.Visible = False   '授權書
            'LSampleFile.Visible = False      '樣品圖
            'LQAFile.Visible = False          '測試報告
            'LContactFile.Visible = False     '切結書
            '按鈕位置
            BSAVE.Style.Add("Top", Top)    '儲存按鈕
            BOK.Style.Add("Top", Top)      '閱讀完成
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
                BOK.Visible = False
                BSAVE.Visible = True
                BSAVE.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("SaveDesc")
                '-- 防止按鈕按二次 JOY 16/08/17
                '舊
                'BSAVE.Attributes("onclick") = "Button('SAVE', '" + BSAVE.Value + "');"
                '新
                BSAVE.Attributes("onclick") = "this.disabled = true;" & "Button('SAVE', '" + BSAVE.Value + "');"
            Else
                BOK.Visible = True
                BSAVE.Visible = False
                '-- 防止按鈕按二次 JOY 16/08/17
                '舊
                'BOK.Attributes("onclick") = "Button('OK', '閱讀完成');"
                '新
                BOK.Attributes("onclick") = "this.disabled = true;" & "Button('OK', '閱讀完成');"
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
        BSpec.Attributes("onclick") = "SpecPicker('Form1.DSpec');"      '規格
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
        Top = 1056
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

        'LOGO
        Select Case FindFieldInf("Logo")
            Case 0  '顯示
                DLogo.BackColor = Color.LightGray
                'DLogo.Enabled = False
                DLogo.Visible = True
            Case 1  '修改+檢查
                DLogo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DLogoRqd", "DLogo", "異常：需輸入Logo")
                DLogo.Visible = True
            Case 2  '修改
                DLogo.BackColor = Color.Yellow
                DLogo.Visible = True
            Case Else   '隱藏
                DLogo.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Logo", "ZZZZZZ")

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


        'Customer Code
        Select Case FindFieldInf("CustomerCode")
            Case 0  '顯示
                DCustomerCode.BackColor = Color.LightGray
                DCustomerCode.ReadOnly = True
                DCustomerCode.Visible = True
            Case 1  '修改+檢查
                DCustomerCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerCodeRqd", "DCustomerCode", "異常：需輸入Customer Code")
                DCustomerCode.Visible = True
            Case 2  '修改
                DCustomerCode.BackColor = Color.Yellow
                DCustomerCode.Visible = True
            Case Else   '隱藏
                DCustomerCode.Visible = False
        End Select
        If pPost = "New" Then DCustomerCode.Text = ""

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

        'Buyer
        Select Case FindFieldInf("Buyer")
            Case 0  '顯示
                DBuyer.BackColor = Color.LightGray
                DBuyer.ReadOnly = True
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
        If pPost = "New" Then DBuyer.Text = ""
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
        '開發理由
        Select Case FindFieldInf("ManufFlow")
            Case 0  '顯示
                DManufFlow.BackColor = Color.LightGray
                DManufFlow.ReadOnly = True
                DManufFlow.Visible = True
            Case 1  '修改+檢查
                DManufFlow.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DManufFlowRqd", "DManufFlow", "異常：需輸入開發理由")
                DManufFlow.Visible = True
            Case 2  '修改
                DManufFlow.BackColor = Color.Yellow
                DManufFlow.Visible = True
            Case Else   '隱藏
                DManufFlow.Visible = False
        End Select
        If pPost = "New" Then DManufFlow.Text = ""
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
                DSuppiler.ReadOnly = True
                DSuppiler.Visible = True
            Case 1  '修改+檢查
                DSuppiler.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSuppilerRqd", "DSuppiler", "異常：需輸入外注廠商")
                DSuppiler.Visible = True
            Case 2  '修改
                DSuppiler.BackColor = Color.Yellow
                DSuppiler.Visible = True
            Case Else   '隱藏
                DSuppiler.Visible = False
        End Select
        If pPost = "New" Then DSuppiler.Text = ""

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
                DQAA15.BackColor = Color.LightGray
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
                DQAB15.BackColor = Color.LightGray
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
                DQAC15.BackColor = Color.LightGray
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
                DQAD15.BackColor = Color.LightGray
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
                DQAA15.Visible = True
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
                DQAB15.Visible = True
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
                DQAC15.Visible = True
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
                DQAD15.Visible = True
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
                DQAA15.BackColor = Color.GreenYellow
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
                DQAB15.BackColor = Color.GreenYellow
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
                DQAC15.BackColor = Color.GreenYellow
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
                DQAD15.BackColor = Color.GreenYellow
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
                ShowRequiredFieldValidator("DQAA15Rqd", "DQAA15", "異常：需輸入品質資訊(A15)")
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
                ShowRequiredFieldValidator("DQAB15Rqd", "DQAB15", "異常：需輸入品質資訊(B15)")
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
                ShowRequiredFieldValidator("DQAC15Rqd", "DQAC15", "異常：需輸入品質資訊(C15)")
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
                ShowRequiredFieldValidator("DQAD15Rqd", "DQAD15", "異常：需輸入品質資訊(D15)")
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
                DQAA15.Visible = True
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
                DQAB15.Visible = True
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
                DQAC15.Visible = True
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
                DQAD15.Visible = True
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
                DQAA15.BackColor = Color.Yellow
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
                DQAB15.BackColor = Color.Yellow
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
                DQAC15.BackColor = Color.Yellow
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
                DQAD15.BackColor = Color.Yellow
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
                DQAA15.Visible = True
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
                DQAB15.Visible = True
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
                DQAC15.Visible = True
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
                DQAD15.Visible = True
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
                DQAA15.Visible = False
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
                DQAB15.Visible = False
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
                DQAC15.Visible = False
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
                DQAD15.Visible = False
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


        'LOGO
        If pFieldName = "Logo" Then
            DLogo.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLogo.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='100' and DKey='LOGO' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLogo.Items.Add(ListItem1)
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
                SQL = "Select * From M_Referp Where Cat='007' and (DKey like 'In%' or DKey = '')  Order by DKey, Data "
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
                SQL = "Select * From M_Referp Where Cat='101' Order by DKey, Data "
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
            DQAA3.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA3.Items.Add(ListItem1)
        End If
        'DQAA7
        If pFieldName = "DQAA7" Then
            DQAA7.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA7.Items.Add(ListItem1)
        End If
        'DQAA8
        If pFieldName = "DQAA8" Then
            DQAA8.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA8.Items.Add(ListItem1)
        End If
        'DQAA9
        If pFieldName = "DQAA9" Then
            DQAA9.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA9.Items.Add(ListItem1)
        End If
        'DQAA10
        If pFieldName = "DQAA10" Then
            DQAA10.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA10.Items.Add(ListItem1)
        End If
        'DQAA11
        If pFieldName = "DQAA11" Then
            DQAA11.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA11.Items.Add(ListItem1)
        End If
        'DQAA14
        If pFieldName = "DQAA14" Then
            DQAA14.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA14.Items.Add(ListItem1)
        End If
        'DQAA15
        If pFieldName = "DQAA15" Then
            DQAA15.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA15.Items.Add(ListItem1)
        End If
        'DQAB3
        If pFieldName = "DQAB3" Then
            DQAB3.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB3.Items.Add(ListItem1)
        End If
        'DQAB7
        If pFieldName = "DQAB7" Then
            DQAB7.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB7.Items.Add(ListItem1)
        End If
        'DQAB8
        If pFieldName = "DQAB8" Then
            DQAB8.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB8.Items.Add(ListItem1)
        End If
        'DQAB9
        If pFieldName = "DQAB9" Then
            DQAB9.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB9.Items.Add(ListItem1)
        End If
        'DQAB10
        If pFieldName = "DQAB10" Then
            DQAB10.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB10.Items.Add(ListItem1)
        End If
        'DQAB11
        If pFieldName = "DQAB11" Then
            DQAB11.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB11.Items.Add(ListItem1)
        End If
        'DQAB14
        If pFieldName = "DQAB14" Then
            DQAB14.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB14.Items.Add(ListItem1)
        End If
        'DQAB15
        If pFieldName = "DQAB15" Then
            DQAB15.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB15.Items.Add(ListItem1)
        End If
        'DQAC3
        If pFieldName = "DQAC3" Then
            DQAC3.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC3.Items.Add(ListItem1)
        End If
        'DQAC7
        If pFieldName = "DQAC7" Then
            DQAC7.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC7.Items.Add(ListItem1)
        End If
        'DQAC8
        If pFieldName = "DQAC8" Then
            DQAC8.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC8.Items.Add(ListItem1)
        End If
        'DQAC9
        If pFieldName = "DQAC9" Then
            DQAC9.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC9.Items.Add(ListItem1)
        End If
        'DQAC10
        If pFieldName = "DQAC10" Then
            DQAC10.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC10.Items.Add(ListItem1)
        End If
        'DQAC11
        If pFieldName = "DQAC11" Then
            DQAC11.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC11.Items.Add(ListItem1)
        End If
        'DQAC14
        If pFieldName = "DQAC14" Then
            DQAC14.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC14.Items.Add(ListItem1)
        End If
        'DQAC15
        If pFieldName = "DQAC15" Then
            DQAC15.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC15.Items.Add(ListItem1)
        End If

        'DQAD3
        If pFieldName = "DQAD3" Then
            DQAD3.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD3.Items.Add(ListItem1)
        End If
        'DQAD7
        If pFieldName = "DQAD7" Then
            DQAD7.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD7.Items.Add(ListItem1)
        End If
        'DQAD8
        If pFieldName = "DQAD8" Then
            DQAD8.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD8.Items.Add(ListItem1)
        End If
        'DQAD9
        If pFieldName = "DQAD9" Then
            DQAD9.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD9.Items.Add(ListItem1)
        End If
        'DQAD10
        If pFieldName = "DQAD10" Then
            DQAD10.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD10.Items.Add(ListItem1)
        End If
        'DQAD11
        If pFieldName = "DQAD11" Then
            DQAD11.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD11.Items.Add(ListItem1)
        End If
        'DQAD14
        If pFieldName = "DQAD14" Then
            DQAD14.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD14.Items.Add(ListItem1)
        End If
        'DQAD15
        If pFieldName = "DQAD15" Then
            DQAD15.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD15.Items.Add(ListItem1)
        End If
        '------------------
        'DQAA6
        If pFieldName = "DQAA6" Then
            DQAA6.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA6.Items.Add(ListItem1)
        End If
        'DQAA12
        If pFieldName = "DQAA12" Then
            DQAA12.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA12.Items.Add(ListItem1)
        End If
        'DQAA13
        If pFieldName = "DQAA13" Then
            DQAA13.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA13.Items.Add(ListItem1)
        End If

        'DQAB6
        If pFieldName = "DQAB6" Then
            DQAB6.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB6.Items.Add(ListItem1)
        End If
        'DQAB12
        If pFieldName = "DQAB12" Then
            DQAB12.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB12.Items.Add(ListItem1)
        End If
        'DQAB13
        If pFieldName = "DQAB13" Then
            DQAB13.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB13.Items.Add(ListItem1)
        End If

        'DQAC6
        If pFieldName = "DQAC6" Then
            DQAC6.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC6.Items.Add(ListItem1)
        End If
        'DQAC12
        If pFieldName = "DQAC12" Then
            DQAC12.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC12.Items.Add(ListItem1)
        End If
        'DQAC13
        If pFieldName = "DQAC13" Then
            DQAC13.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC13.Items.Add(ListItem1)
        End If

        'DQAD6
        If pFieldName = "DQAD6" Then
            DQAD6.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD6.Items.Add(ListItem1)
        End If
        'DQAD12
        If pFieldName = "DQAD12" Then
            DQAD12.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD12.Items.Add(ListItem1)
        End If
        'DQAD13
        If pFieldName = "DQAD13" Then
            DQAD13.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD13.Items.Add(ListItem1)
        End If
        '----------------------------------

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
    '**     閱讀完成
    '**
    '*****************************************************************
    Private Sub BOK_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
        Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.close();</script>", "Form1.DReady", "已閱讀"))
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     儲存按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BSAVE_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSAVE.ServerClick
        If Request.Cookies("RunBSAVE").Value = True Then
            Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
            Dim Message As String = ""
            Dim wQCLT As Integer = 0 'QC-L/T

            Dim NewFormSno As Integer = wFormSno    '表單流水號
            Dim RtnCode As Integer = 0

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

            '儲存資料
            If ErrCode = 0 Then
                'If wFormSno > 0 And wStep > 2 Then    '判斷是否有資料
                If wFormSno > 0 Then    '判斷是否有資料
                    ModifyData("MOD", wFormSno)
                Else
                    '取得表單流水號
                    RtnCode = oCommon.GetFormSeqNo(wFormNo, NewFormSno) '表單號碼, 表單流水號
                    If RtnCode <> 0 Then
                        ErrCode = 9110
                    Else
                        AppendData("ADD", NewFormSno)
                    End If
                End If
            Else
                If ErrCode = 9010 Then Message = "上傳檔案格式錯誤,請確認!"
                If ErrCode = 9020 Then Message = "上傳檔案Size超過1024KB,請確認!"
                If ErrCode = 9030 Then Message = "上傳檔案沒指定,請確認!"
                If ErrCode = 9210 Then Message = "延遲理由為其他時需填寫說明,請確認!"
                If ErrCode = 9050 Then Message = "QC L/T需為有效數字,請確認!"
                Response.Write(YKK.ShowMessage(Message))
            End If      '上傳檔案ErrCode=0

            If ErrCode = 0 Then
                'If wFormSno > 0 And wStep > 2 Then      '判斷是否[簽核]
                If wFormSno > 0 Then    '判斷是否有資料
                    If wFormNo = "900001" Then
                        Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.opener.document.{6}.value = '{7:d}'; window.opener.document.{8}.value = '{9:d}'; window.opener.document.{10}.value = '{11:d}'; window.opener.document.{12}.value = '{13:d}'; window.close();</script>", "Form1.DMakeCAM", DMakeCAM.SelectedValue, "Form1.DReady", "已閱讀", "Form1.DLevel", DLevel.SelectedValue, "Form1.DSliderCode", DSliderCode.Text, "Form1.DSuppiler", DSuppiler.Text, "Form1.DMapNo", DMapNo.Text, "Form1.DQCLT", DQCLT.Text))
                    Else
                        Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.opener.document.{6}.value = '{7:d}'; window.close();</script>", "Form1.DReady", "已閱讀", "Form1.DLevel", DLevel.SelectedValue, "Form1.DSliderCode", DSliderCode.Text, "Form1.DMapNo", DMapNo.Text))
                    End If

                Else
                    Dim SQL As String
                    Dim DBDataSet1 As New DataSet
                    'DB連結設定
                    Dim OleDbConnection1 As New OleDbConnection
                    OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
                    OleDbConnection1.Open()

                    SQL = "Select FormNo, FormSno From F_ManufInModSheet "
                    SQL = SQL & " Where Sts = 0 "
                    SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
                    SQL = SQL & "   And FormSno =  '" & CStr(NewFormSno) & "'"
                    SQL = SQL & " Order by Unique_ID Desc "
                    Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter1.Fill(DBDataSet1, "F_ManufInModSheet")
                    If DBDataSet1.Tables("F_ManufInModSheet").Rows.Count > 0 Then
                        rFormNo = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("FormNo")
                        rFormSno = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("FormSno")
                    End If
                    'DB連結關閉
                    OleDbConnection1.Close()

                    If wFormNo = "900001" Then
                        Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.opener.document.{6}.value = '{7:d}'; window.opener.document.{8}.value = '{9:d}'; window.opener.document.{10}.value = '{11:d}'; window.opener.document.{12}.value = '{13:d}'; window.opener.document.{14}.value = '{15:d}'; window.opener.document.{16}.value = '{17:d}'; window.close();</script>", "Form1.DNFormNo", rFormNo, "Form1.DNFormSno", rFormSno, "Form1.DReady", "已閱讀", "Form1.DMakeCAM", DMakeCAM.SelectedValue, "Form1.DLevel", DLevel.SelectedValue, "Form1.DSliderCode", DSliderCode.Text, "Form1.DSuppiler", DSuppiler.Text, "Form1.DMapNo", DMapNo.Text, "Form1.DQCLT", DQCLT.Text))
                    Else
                        Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.opener.document.{6}.value = '{7:d}'; window.opener.document.{8}.value = '{9:d}'; window.opener.document.{10}.value = '{11:d}'; window.close();</script>", "Form1.DNFormNo", rFormNo, "Form1.DNFormSno", rFormSno, "Form1.DReady", "已閱讀", "Form1.DLevel", DLevel.SelectedValue, "Form1.DSliderCode", DSliderCode.Text, "Form1.DMapNo", DMapNo.Text))
                    End If
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("ManufInFilePath"))
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim SQL As String
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_ManufInSheet "
        SQL = SQL & " Where Sts = 1 "
        SQL = SQL & "   And FormNo =  '" & wOFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ManufInSheet")
        If DBDataSet1.Tables("F_ManufInSheet").Rows.Count > 0 Then
            Dim OleDbConnection2 As New OleDbConnection
            OleDbConnection2.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
            Dim OleDBCommand2 As New OleDbCommand

            SQL = "Insert into F_ManufInModSheet "
            SQL = SQL + "( "
            SQL = SQL + "Sts, CompletedTime, FormNo, FormSNo, "                         '1~4
            SQL = SQL + "No, Date, Division, Person, Cpsc, Logo, SliderCode, "                      '5~9
            SQL = SQL + "SliderGRCode, Spec, MapNo, MapFile, RefFile, "               '10~14
            SQL = SQL + "Level, Assembler, SliderType1, SliderType2, ManufPlace, "      '15~19
            SQL = SQL + "Material, MaterialOther, SellVendor, CustomerCode, Buyer, "  '20~24
            SQL = SQL + "ConfirmFile, AuthorizeFile, DevReason, ManufFlow, Sample, Price, "        '25~29
            SQL = SQL + "ArMoldFee, PurMold, PullerPrice, Suppiler, MakeCAM, MoldQty, "          '29~33
            SQL = SQL + "MoldPoint, Quality1, Quality2, QAFile, SampleFile,  "    '34~38
            SQL = SQL + "ContactFile, OFormNo, OFormSno,  QCLT, "    '34~38
            SQL = SQL + "SAttachFile, QAttachFile1, QAttachFile2, "               '39~42
            SQL = SQL + "CreateUser, CreateTime, ModifyUser, ModifyTime "               '39~42
            SQL = SQL + ")  "
            SQL = SQL + "VALUES( "
            SQL = SQL + " '0', "                                '狀態(0:未結,1:已結NG,2:已結OK)
            SQL = SQL + " '" + NowDateTime + "', "              '結案日
            SQL = SQL + " '" + wFormNo + "', "                  '表單代號
            SQL = SQL + " '" + CStr(NewFormSno) + "', "           '表單流水號

            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("No") + "', "        'NO
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Date") + "', "      '日期
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Division") + "', "  '部門
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Person") + "', "    '擔當
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Cpsc") + "', "      'Cpsc
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Logo") + "', "      'Logo

            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SliderCode") + "', "     'Slider Code(顯示用)
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SliderGRCode") + "', "   'Slider Group Code
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Spec") + "', "      '規格
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MapNo") + "', "     '圖號
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MapFile") + "', "   '圖檔1
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("RefFile") + "', "   '圖檔2
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Level") + "', "     '難易度
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Assembler") + "', "    '組立判定
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SliderType1") + "', "  '拉頭種類-1
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SliderType2") + "', "  '拉頭種類-2
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("ManufPlace") + "', "   '生產地
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Material") + "', "     '材質
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MaterialOther") + "', "    '材質其他說明
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SellVendor") + "', "       '委託廠商
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("CustomerCode") + "', "       'Customer Code

            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Buyer") + "', "            'Buyer
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("ConfirmFile") + "', "      '確認書
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("AuthorizeFile") + "', "    '授權書
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("DevReason") + "', "    '開發理由
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("ManufFlow") + "', "    '
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Sample")) + "', "       'Sample
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Price")) + "', "   'Price(顯示用)
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("ArMoldFee") + "', "    '應收模具費
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("PurMold") + "', "        '模具購入費
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("PullerPrice") + "', "    '引手購入單價
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Suppiler") + "', "       '外注商
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MakeCAM") + "', "       'CAM
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MoldQty") + "', "        '模型數
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("MoldPoint") + "', "      '穴取數
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Quality1")) + "', "   '品質分析-1(顯示用)
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("Quality2")) + "', "   '品質分析-2(顯示用)
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("QAFile") + "', "    '品質分析附件
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SampleFile") + "', "    '樣品圖檔
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("ContactFile") + "', "    '切結書
            SQL = SQL + " '" + wOFormNo + "', "           '原表單
            SQL = SQL + " '" + CStr(wOFormSno) + "', "    '原單號
            SQL = SQL + " '" + "" + "', "                 'QC-L/T

            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("SAttachFile") + "', "   '樣品-其他附件
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("QAttachFile1") + "', "  '品質1-其他附件
            SQL = SQL + " '" + DBDataSet1.Tables("F_ManufInSheet").Rows(0).Item("QAttachFile2") + "', "  '品質2-品質分析依賴

            SQL = SQL + " '" + Request.QueryString("pUserID") + "', "     '作成者
            SQL = SQL + " '" + NowDateTime + "', "       '作成時間
            SQL = SQL + " '" + "" + "', "                '修改者
            SQL = SQL + " '" + NowDateTime + "' "        '修改時間
            SQL = SQL + " ) "
            OleDBCommand2.Connection = OleDbConnection2
            OleDBCommand2.CommandText = SQL
            OleDbConnection2.Open()
            OleDBCommand2.ExecuteNonQuery()
            OleDbConnection2.Close()
        End If
        'DB連結關閉
        OleDbConnection1.Close()

        ModifyData(pFun, NewFormSno)
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pFormSno As Integer)
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("ManufInFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQL As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        Dim RtnCode As Integer = 0

        OleDbConnection1.Open()
        SQL = "Update F_ManufInModSheet Set "
        SQL = SQL + " No = '" & YKK.ReplaceString(DNo.Text) & "',"
        SQL = SQL + " Date = '" & DDate.Text & "',"
        SQL = SQL + " Division = '" & DDivision.SelectedValue & "',"
        SQL = SQL + " Person = '" & DPerson.SelectedValue & "',"
        SQL = SQL + " Cpsc = '" & DCpsc.SelectedValue & "',"
        SQL = SQL + " Logo = '" & DLogo.SelectedValue & "',"

        SQL = SQL + " SliderCode = '" & YKK.ReplaceString(DSliderCode.Text) & "',"
        SQL = SQL + " SliderGRCode = '" & YKK.ReplaceString(DSliderGRCode.Text) & "',"
        SQL = SQL + " Spec = '" & DSpec.Text & "',"
        SQL = SQL + " MapNo = '" & YKK.ReplaceString(DMapNo.Text) & "',"
        SQL = SQL + " OFormNo = '" & DOFormNo.Text & "',"
        SQL = SQL + " OFormSno = '" & DOFormSno.Text & "',"

        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(pFormSno) & "-" & "Map" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(pFormSno) & "-" & "Map" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DMapFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQL = SQL + " MapFile = N'" & FileName & "',"
            End If
        End If

        If DRefFile.Visible Then
            If DRefFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(pFormSno) & "-" & "Ref" & "-" & UploadDateTime & "-" & Right(DRefFile.PostedFile.FileName, InStr(StrReverse(DRefFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(pFormSno) & "-" & "Ref" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DRefFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DRefFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQL = SQL + " RefFile = N'" & FileName & "',"
            End If
        End If

        SQL = SQL + " Level = '" & DLevel.SelectedValue & "',"
        SQL = SQL + " Assembler = N'" & YKK.ReplaceString(DAssembler.Text) & "',"
        SQL = SQL + " SliderType1 = '" & DSliderType1.SelectedValue & "',"
        SQL = SQL + " SliderType2 = '" & DSliderType2.SelectedValue & "',"
        SQL = SQL + " ManufPlace = '" & DManufPlace.SelectedValue & "',"
        SQL = SQL + " Material = '" & DMaterial.SelectedValue & "',"
        SQL = SQL + " MaterialOther = '" & "" & "',"
        SQL = SQL + " SellVendor = N'" & YKK.ReplaceString(DSellVendor.Text) & "',"
        SQL = SQL + " CustomerCode = N'" & YKK.ReplaceString(DCustomerCode.Text) & "',"

        SQL = SQL + " Buyer = N'" & YKK.ReplaceString(DBuyer.Text) & "',"
        SQL = SQL + " Suppiler = N'" & YKK.ReplaceString(DSuppiler.Text) & "',"
        SQL = SQL + " MakeCAM = N'" & DMakeCAM.SelectedValue & "',"

        If DConfirmFile.Visible Then
            If DConfirmFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(pFormSno) & "-" & "Confirm" & "-" & UploadDateTime & "-" & Right(DConfirmFile.PostedFile.FileName, InStr(StrReverse(DConfirmFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(pFormSno) & "-" & "Confirm" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DConfirmFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DConfirmFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQL = SQL + " ConfirmFile = N'" & FileName & "',"
            End If
        End If

        If DAuthorizeFile.Visible Then
            If DAuthorizeFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(pFormSno) & "-" & "Authorize" & "-" & UploadDateTime & "-" & Right(DAuthorizeFile.PostedFile.FileName, InStr(StrReverse(DAuthorizeFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(pFormSno) & "-" & "Authorize" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DAuthorizeFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DAuthorizeFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQL = SQL + " AuthorizeFile = N'" & FileName & "',"
            End If
        End If

        If DContactFile.Visible Then
            If DContactFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(pFormSno) & "-" & "Contact" & "-" & UploadDateTime & "-" & Right(DContactFile.PostedFile.FileName, InStr(StrReverse(DContactFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(pFormSno) & "-" & "Contact" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DContactFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DContactFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQL = SQL + " ContactFile = N'" & FileName & "',"
            End If
        End If

        SQL = SQL + " DevReason = N'" & YKK.ReplaceString(DDevReason.Text) & "',"
        SQL = SQL + " ManufFlow = N'" & YKK.ReplaceString(DManufFlow.Text) & "',"
        If DSampleA1.Text <> "" Then
            SQL = SQL + " Sample = '" & "1" & "',"
        Else
            SQL = SQL + " Sample = '" & "0" & "',"
        End If
        If DPriceA1.Text <> "" Then
            SQL = SQL + " Price = '" & "1" & "',"
        Else
            SQL = SQL + " Price = '" & "0" & "',"
        End If
        SQL = SQL + " ArMoldFee = '" & YKK.ReplaceString(DArMoldFee.Text) & "',"
        SQL = SQL + " PurMold = '" & YKK.ReplaceString(DPurMold.Text) & "',"
        SQL = SQL + " PullerPrice = '" & YKK.ReplaceString(DPullerPrice.Text) & "',"
        SQL = SQL + " MoldQty = '" & YKK.ReplaceString(DMoldQty.Text) & "',"
        SQL = SQL + " MoldPoint = '" & YKK.ReplaceString(DMoldPoint.Text) & "',"
        If DQAA1.Text <> "" Then
            SQL = SQL + " Quality1 = '" & "1" & "',"
        Else
            SQL = SQL + " Quality1 = '" & "0" & "',"
        End If
        If DFQAA1.Text <> "" Then
            SQL = SQL + " Quality2 = '" & "1" & "',"
        Else
            SQL = SQL + " Quality2 = '" & "0" & "',"
        End If

        If DQAFile.Visible Then
            If DQAFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(pFormSno) & "-" & "QA" & "-" & UploadDateTime & "-" & Right(DQAFile.PostedFile.FileName, InStr(StrReverse(DQAFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(pFormSno) & "-" & "QA" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQAFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQAFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQL = SQL + " QAFile = N'" & FileName & "',"
            End If
        End If

        If DSampleFile.Visible Then
            If DSampleFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(pFormSno) & "-" & "Sample" & "-" & UploadDateTime & "-" & Right(DSampleFile.PostedFile.FileName, InStr(StrReverse(DSampleFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(pFormSno) & "-" & "Sample" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSampleFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DSampleFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQL = SQL + " SampleFile = N'" & FileName & "',"
            End If
        End If

        If DSAttachFile.Visible Then
            If DSAttachFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(pFormSno) & "-" & "SAttachFile" & "-" & UploadDateTime & "-" & Right(DSAttachFile.PostedFile.FileName, InStr(StrReverse(DSAttachFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(pFormSno) & "-" & "SAttachFile" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSAttachFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DSAttachFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQL = SQL + " SAttachFile = N'" & FileName & "',"
            End If
        End If

        If DQAttachFile1.Visible Then
            If DQAttachFile1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(pFormSno) & "-" & "QAttachFile1" & "-" & UploadDateTime & "-" & Right(DQAttachFile1.PostedFile.FileName, InStr(StrReverse(DQAttachFile1.PostedFile.FileName), "\") - 1)
                FileName = CStr(pFormSno) & "-" & "QAttachFile1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQAttachFile1.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQAttachFile1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQL = SQL + " QAttachFile1 = N'" & FileName & "',"
            End If
        End If

        If DQAttachFile2.Visible Then
            If DQAttachFile2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(pFormSno) & "-" & "QAttachFile2" & "-" & UploadDateTime & "-" & Right(DQAttachFile2.PostedFile.FileName, InStr(StrReverse(DQAttachFile2.PostedFile.FileName), "\") - 1)
                FileName = CStr(pFormSno) & "-" & "QAttachFile2" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQAttachFile2.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQAttachFile2.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQL = SQL + " QAttachFile2 = N'" & FileName & "',"
            End If
        End If
        SQL = SQL + " QCLT = '" & DQCLT.Text & "',"

        SQL = SQL + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
        SQL = SQL + " Where FormNo  =  '" & wFormNo & "'"
        SQL = SQL + "   And FormSno =  '" & CStr(pFormSno) & "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQL
        OleDBCommand1.ExecuteNonQuery()
        '
        'Modify-Start by Joy 09-10-03
        'Sample Table處理
        Dim i As Integer
        For i = 1 To 3
            SQL = "Insert into F_ManufCOSample "
            SQL = SQL + "( "
            SQL = SQL + "Sts, CompletedTime, FormNo, FormSNo, SeqNo, "
            SQL = SQL + "Spec, Color, Qty, "
            SQL = SQL + "CreateUser, CreateTime, ModifyUser, ModifyTime "
            SQL = SQL + ")  "
            SQL = SQL + "VALUES( "
            SQL = SQL + " '0', "
            SQL = SQL + " '" + NowDateTime + "', "
            SQL = SQL + " '" + wFormNo + "', "
            SQL = SQL + " '" + CStr(pFormSno) + "', "
            SQL = SQL + " '" + CStr(i) + "', "
            Select Case i
                Case Is = 1
                    SQL = SQL + " '" + YKK.ReplaceString(DSampleA1.Text) + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DSampleA2.Text) + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DSampleA3.Text) + "', "
                Case Is = 2
                    SQL = SQL + " '" + YKK.ReplaceString(DSampleB1.Text) + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DSampleB2.Text) + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DSampleB3.Text) + "', "
                Case Is = 3
                    SQL = SQL + " '" + YKK.ReplaceString(DSampleC1.Text) + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DSampleC2.Text) + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DSampleC3.Text) + "', "
            End Select
            SQL = SQL + " '" + Request.QueryString("pUserID") + "', "     '作成者
            SQL = SQL + " '" + NowDateTime + "', "       '作成時間
            SQL = SQL + " '" + "" + "', "                       '修改者
            SQL = SQL + " '" + NowDateTime + "' "       '修改時間
            SQL = SQL + " ) "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()
        Next
        'Price Table處理
        For i = 1 To 10
            SQL = "Insert into F_ManufCOPrice "
            SQL = SQL + "( "
            SQL = SQL + "Sts, CompletedTime, FormNo, FormSNo, SeqNo, "
            SQL = SQL + "Spec, Currency, Price, "
            SQL = SQL + "CreateUser, CreateTime, ModifyUser, ModifyTime "
            SQL = SQL + ")  "
            SQL = SQL + "VALUES( "
            SQL = SQL + " '0', "
            SQL = SQL + " '" + NowDateTime + "', "
            SQL = SQL + " '" + wFormNo + "', "
            SQL = SQL + " '" + CStr(pFormSno) + "', "
            SQL = SQL + " '" + CStr(i) + "', "
            Select Case i
                Case Is = 1
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceA1.Text) + "', "
                    SQL = SQL + " '" + DPriceA2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceA3.Text) + "', "
                Case Is = 2
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceB1.Text) + "', "
                    SQL = SQL + " '" + DPriceB2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceB3.Text) + "', "
                Case Is = 3
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceC1.Text) + "', "
                    SQL = SQL + " '" + DPriceC2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceC3.Text) + "', "
                Case Is = 4
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceD1.Text) + "', "
                    SQL = SQL + " '" + DPriceD2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceD3.Text) + "', "
                Case Is = 5
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceE1.Text) + "', "
                    SQL = SQL + " '" + DPriceE2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceE3.Text) + "', "
                Case Is = 6
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceF1.Text) + "', "
                    SQL = SQL + " '" + DPriceF2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceF3.Text) + "', "
                Case Is = 7
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceG1.Text) + "', "
                    SQL = SQL + " '" + DPriceG2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceG3.Text) + "', "
                Case Is = 8
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceH1.Text) + "', "
                    SQL = SQL + " '" + DPriceH2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceH3.Text) + "', "
                Case Is = 9
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceI1.Text) + "', "
                    SQL = SQL + " '" + DPriceI2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceI3.Text) + "', "
                Case Is = 10
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceJ1.Text) + "', "
                    SQL = SQL + " '" + DPriceJ2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceJ3.Text) + "', "
            End Select
            SQL = SQL + " '" + Request.QueryString("pUserID") + "', "     '作成者
            SQL = SQL + " '" + NowDateTime + "', "       '作成時間
            SQL = SQL + " '" + "" + "', "                       '修改者
            SQL = SQL + " '" + NowDateTime + "' "       '修改時間
            SQL = SQL + " ) "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()
        Next
        'QA1 Table處理
        For i = 1 To 4
            SQL = "Insert into F_ManufCOQA1 "
            SQL = SQL + "( "
            SQL = SQL + "Sts, CompletedTime, FormNo, FormSNo, SeqNo, "
            SQL = SQL + "Date, Spec, Assembler, Surface, GenTani, Kyoudo, Nyuryoku, Kensin, Water, Dry, Yellow, Mityaku1,  Mityaku2, CPSC, EDX, "
            SQL = SQL + "CreateUser, CreateTime, ModifyUser, ModifyTime "
            SQL = SQL + ")  "
            SQL = SQL + "VALUES( "
            SQL = SQL + " '0', "
            SQL = SQL + " '" + NowDateTime + "', "
            SQL = SQL + " '" + wFormNo + "', "
            SQL = SQL + " '" + CStr(pFormSno) + "', "
            SQL = SQL + " '" + CStr(i) + "', "
            Select Case i
                Case Is = 1
                    SQL = SQL + " '" + YKK.ReplaceString(DQAA1.Text) + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DQAA2.Text) + "', "
                    SQL = SQL + " '" + DQAA3.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DQAA4.Text) + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DQAA5.Text) + "', "
                    SQL = SQL + " '" + DQAA6.SelectedValue + "', "
                    SQL = SQL + " '" + DQAA7.SelectedValue + "', "
                    SQL = SQL + " '" + DQAA8.SelectedValue + "', "
                    SQL = SQL + " '" + DQAA9.SelectedValue + "', "
                    SQL = SQL + " '" + DQAA10.SelectedValue + "', "
                    SQL = SQL + " '" + DQAA11.SelectedValue + "', "
                    SQL = SQL + " '" + DQAA12.SelectedValue + "', "
                    SQL = SQL + " '" + DQAA13.SelectedValue + "', "
                    SQL = SQL + " '" + DQAA14.SelectedValue + "', "
                    SQL = SQL + " '" + DQAA15.SelectedValue + "', "
                Case Is = 2
                    SQL = SQL + " '" + YKK.ReplaceString(DQAB1.Text) + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DQAB2.Text) + "', "
                    SQL = SQL + " '" + DQAB3.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DQAB4.Text) + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DQAB5.Text) + "', "
                    SQL = SQL + " '" + DQAB6.SelectedValue + "', "
                    SQL = SQL + " '" + DQAB7.SelectedValue + "', "
                    SQL = SQL + " '" + DQAB8.SelectedValue + "', "
                    SQL = SQL + " '" + DQAB9.SelectedValue + "', "
                    SQL = SQL + " '" + DQAB10.SelectedValue + "', "
                    SQL = SQL + " '" + DQAB11.SelectedValue + "', "
                    SQL = SQL + " '" + DQAB12.SelectedValue + "', "
                    SQL = SQL + " '" + DQAB13.SelectedValue + "', "
                    SQL = SQL + " '" + DQAB14.SelectedValue + "', "
                    SQL = SQL + " '" + DQAB15.SelectedValue + "', "
                Case Is = 3
                    SQL = SQL + " '" + YKK.ReplaceString(DQAC1.Text) + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DQAC2.Text) + "', "
                    SQL = SQL + " '" + DQAC3.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DQAC4.Text) + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DQAC5.Text) + "', "
                    SQL = SQL + " '" + DQAC6.SelectedValue + "', "
                    SQL = SQL + " '" + DQAC7.SelectedValue + "', "
                    SQL = SQL + " '" + DQAC8.SelectedValue + "', "
                    SQL = SQL + " '" + DQAC9.SelectedValue + "', "
                    SQL = SQL + " '" + DQAC10.SelectedValue + "', "
                    SQL = SQL + " '" + DQAC11.SelectedValue + "', "
                    SQL = SQL + " '" + DQAC12.SelectedValue + "', "
                    SQL = SQL + " '" + DQAC13.SelectedValue + "', "
                    SQL = SQL + " '" + DQAC14.SelectedValue + "', "
                    SQL = SQL + " '" + DQAC15.SelectedValue + "', "
                Case Is = 4
                    SQL = SQL + " '" + YKK.ReplaceString(DQAD1.Text) + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DQAD2.Text) + "', "
                    SQL = SQL + " '" + DQAD3.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DQAD4.Text) + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DQAD5.Text) + "', "
                    SQL = SQL + " '" + DQAD6.SelectedValue + "', "
                    SQL = SQL + " '" + DQAD7.SelectedValue + "', "
                    SQL = SQL + " '" + DQAD8.SelectedValue + "', "
                    SQL = SQL + " '" + DQAD9.SelectedValue + "', "
                    SQL = SQL + " '" + DQAD10.SelectedValue + "', "
                    SQL = SQL + " '" + DQAD11.SelectedValue + "', "
                    SQL = SQL + " '" + DQAD12.SelectedValue + "', "
                    SQL = SQL + " '" + DQAD13.SelectedValue + "', "
                    SQL = SQL + " '" + DQAD14.SelectedValue + "', "
                    SQL = SQL + " '" + DQAD15.SelectedValue + "', "
            End Select
            SQL = SQL + " '" + Request.QueryString("pUserID") + "', "     '作成者
            SQL = SQL + " '" + NowDateTime + "', "       '作成時間
            SQL = SQL + " '" + "" + "', "                       '修改者
            SQL = SQL + " '" + NowDateTime + "' "       '修改時間
            SQL = SQL + " ) "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()
        Next
        'QA2 Table處理
        For i = 1 To 4
            SQL = "Insert into F_ManufCOQA2 "
            SQL = SQL + "( "
            SQL = SQL + "Sts, CompletedTime, FormNo, FormSNo, SeqNo, "
            SQL = SQL + "Date, QACheck, Remark, "
            SQL = SQL + "CreateUser, CreateTime, ModifyUser, ModifyTime "
            SQL = SQL + ")  "
            SQL = SQL + "VALUES( "
            SQL = SQL + " '0', "
            SQL = SQL + " '" + NowDateTime + "', "
            SQL = SQL + " '" + wFormNo + "', "
            SQL = SQL + " '" + CStr(pFormSno) + "', "
            SQL = SQL + " '" + CStr(i) + "', "
            Select Case i
                Case Is = 1
                    SQL = SQL + " '" + YKK.ReplaceString(DFQAA1.Text) + "', "
                    SQL = SQL + " '" + DFQAA2.SelectedValue + "', "
                    SQL = SQL + " N'" + YKK.ReplaceString(DFQAA3.Text) + "', "
                Case Is = 2
                    SQL = SQL + " '" + YKK.ReplaceString(DFQAB1.Text) + "', "
                    SQL = SQL + " '" + DFQAB2.SelectedValue + "', "
                    SQL = SQL + " N'" + YKK.ReplaceString(DFQAB3.Text) + "', "
                Case Is = 3
                    SQL = SQL + " '" + YKK.ReplaceString(DFQAC1.Text) + "', "
                    SQL = SQL + " '" + DFQAC2.SelectedValue + "', "
                    SQL = SQL + " N'" + YKK.ReplaceString(DFQAC3.Text) + "', "
                Case Is = 4
                    SQL = SQL + " '" + YKK.ReplaceString(DFQAD1.Text) + "', "
                    SQL = SQL + " '" + DFQAD2.SelectedValue + "', "
                    SQL = SQL + " N'" + YKK.ReplaceString(DFQAD3.Text) + "', "
            End Select
            SQL = SQL + " '" + Request.QueryString("pUserID") + "', "     '作成者
            SQL = SQL + " '" + NowDateTime + "', "       '作成時間
            SQL = SQL + " '" + "" + "', "                       '修改者
            SQL = SQL + " '" + NowDateTime + "' "       '修改時間
            SQL = SQL + " ) "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()
        Next
        OleDbConnection1.Close()
        'Modify-end by Joy
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
    Private Sub BPrint_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs)
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

End Class
