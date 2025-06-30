Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class SufaceSheet_02
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
    Protected WithEvents DBFinalDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAllowSample As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQty As System.Web.UI.WebControls.TextBox
    Protected WithEvents DColor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DManufType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSliderSample As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReqQty As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReqDelDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDevReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSufaceSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DSuppiler As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBuyer As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LQCReqFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOPManualFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LEACheckFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFinalFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents LContactFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LForcastFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LManufFlowFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAttachSample As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DORNO As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSufaceSheet2 As System.Web.UI.WebControls.Image
    Protected WithEvents LCustSampleFile As System.Web.UI.WebControls.Image
    Protected WithEvents Panel1 As System.Web.UI.WebControls.Panel
    Protected WithEvents DStatus As System.Web.UI.WebControls.TextBox
    Protected WithEvents LSuface As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSchedule As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCap As System.Web.UI.WebControls.TextBox
    Protected WithEvents LEACheckFile1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DQCCheck14 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck13 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DLOSS As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCCheck15 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DYearSeason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCCheck16 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LPFASFile As System.Web.UI.WebControls.HyperLink

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

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim NowDateTime As String       '現在日期時間
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "SufaceSheet_02.aspx"

        SetParameter()          '設定共用參數

        If Not Me.IsPostBack Then   '不是PostBack
            ShowFormData()      '顯示表單資料
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
    End Sub


    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("SufaceFilePath")
        Dim PathOld1 As String = ConfigurationSettings.AppSettings.Get("HttpOld")
        Dim PathOld2 As String = ConfigurationSettings.AppSettings.Get("SufaceFilePath")
        Dim RtnCode As Integer = 0
        Dim i As Integer
        Dim SQL As String
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
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("UPDSts") = 1 Then          '追加工程狀態
                DStatus.Text = "追加工程進行中"
                DStatus.Visible = True
                Panel1.Visible = True
            Else
                DStatus.Visible = False
                Panel1.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Suface") = 1 Then          '追加表面處理
                LSuface.NavigateUrl = "SufaceList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno & "&pFun=Suface"
                LSuface.Visible = True
            Else
                LSuface.Visible = False
            End If
            '------------------------------------------------------------------------------------------------
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("CustSampleFile") <> "" Then        '客戶樣品圖
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("CustSampleFile"))

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
            DPrice.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Price")                 '售價
            DDevReason.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("DevReason")         '開發理由
            DYearSeason.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("YearSeason")         '年季

            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FinalSampleFile") <> "" Then       '最終樣品圖
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FinalSampleFile"))
                LFinalSampleFile.ImageUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FinalSampleFile")
            Else
                LFinalSampleFile.Visible = False
            End If
            SetFieldData("ManufType", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ManufType"))  '內製/外注
            SetFieldData("Suppiler", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Suppiler"))    '外注商
            DColor.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Color")                   'Color
            DQty.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Qty")                       '數量
            DCap.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Cap")                   '日產能
            DSchedule.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Schedule")                       '基準日程
            DFReason.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FReason")                       '理由
            SetFieldData("AllowSample", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("AllowSample"))    '限度樣品
            DBFinalDate.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("BFinalDate")        '預定完成日
            DCode.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Code")                    'Code
            DEnglishName.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EnglishName")      '英文名稱
            DLOSS.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("LOSS")                    'LOSS
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCReqFile") <> "" Then              '品質依賴書
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCReqFile"))
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
            SetFieldData("QCCheck16", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck16"))   'PFAS


            '  If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck6File") <> "" Then           '電鍍膜厚
            '  LQCCheck6File.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck6File")
            'Else
            'LQCCheck6File.Visible = False
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
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile"))
                LEACheckFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile")
            Else
                LEACheckFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile1") <> "" Then           'A01有害物質報告
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile1"))
                LEACheckFile1.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile1")
            Else
                LEACheckFile1.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCFinalFile") <> "" Then            '測試報告書
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCFinalFile"))
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
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ManufFlowFile"))
                LManufFlowFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ManufFlowFile")
            Else
                LManufFlowFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ForcastFile") <> "" Then             '報價單
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ForcastFile"))
                LForcastFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ForcastFile")
            Else
                LForcastFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OPManualFile") <> "" Then            '作業標準書
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OPManualFile"))
                LOPManualFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OPManualFile")
            Else
                LOPManualFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ContactFile") <> "" Then             '切結書
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ContactFile"))
                LContactFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ContactFile")
            Else
                LContactFile.Visible = False
            End If

            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("PFASFile") <> "" Then             '???
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("PFASFile"))
                LPFASFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("PFASFile")
            Else
                LPFASFile.Visible = False
            End If
            '------------------------------------------------------------------------------------------------
        End If
        DFormSno.Text = "單號：" & CStr(wFormSno)
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    '*****************************************************************
    '**(ShowSheetField)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        '部門
        If pFieldName = "Division" Then
            DDivision.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DDivision.Items.Add(ListItem1)
        End If
        '擔當
        If pFieldName = "Person" Then
            DPerson.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DPerson.Items.Add(ListItem1)
        End If

        'Buyer
        If pFieldName = "Buyer" Then
            DBuyer.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DBuyer.Items.Add(ListItem1)
        End If

        '附樣
        If pFieldName = "AttachSample" Then
            DAttachSample.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAttachSample.Items.Add(ListItem1)
        End If

        '內製/外注
        If pFieldName = "ManufType" Then
            DManufType.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DManufType.Items.Add(ListItem1)
        End If

        'Suppiler
        If pFieldName = "Suppiler" Then
            DSuppiler.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSuppiler.Items.Add(ListItem1)
        End If

        '限度樣品
        If pFieldName = "AllowSample" Then
            DAllowSample.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAllowSample.Items.Add(ListItem1)
        End If

        '口徑寸法
        If pFieldName = "QCCheck1" Then
            DQCCheck1.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck1.Items.Add(ListItem1)
        End If

        '摺動抵抗
        If pFieldName = "QCCheck2" Then
            DQCCheck2.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck2.Items.Add(ListItem1)
        End If

        'LOCK強度
        If pFieldName = "QCCheck3" Then
            DQCCheck3.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck3.Items.Add(ListItem1)
        End If

        '90度強度
        If pFieldName = "QCCheck4" Then
            DQCCheck4.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck4.Items.Add(ListItem1)
        End If

        '扭力
        If pFieldName = "QCCheck5" Then
            DQCCheck5.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck5.Items.Add(ListItem1)
        End If

        'N-ANTI
        If pFieldName = "QCCheck15" Then
            DQCCheck15.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck15.Items.Add(ListItem1)
        End If



        'PFAS
        If pFieldName = "QCCheck16" Then
            DQCCheck16.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck16.Items.Add(ListItem1)
        End If




        '檢針
        If pFieldName = "QCCheck7" Then
            DQCCheck7.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck7.Items.Add(ListItem1)
        End If

        'AATCC
        If pFieldName = "QCCheck8" Then
            DQCCheck8.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck8.Items.Add(ListItem1)
        End If

        '乾洗
        If pFieldName = "QCCheck9" Then
            DQCCheck9.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck9.Items.Add(ListItem1)
        End If

        '鹽水噴霧
        If pFieldName = "QCCheck10" Then
            DQCCheck10.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck10.Items.Add(ListItem1)
        End If

        '一次密著
        If pFieldName = "QCCheck11" Then
            DQCCheck11.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck11.Items.Add(ListItem1)
        End If

        '二次密著
        If pFieldName = "QCCheck12" Then
            '二次密著
            DQCCheck12.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck12.Items.Add(ListItem1)
        End If

        'Oeko-tex
        If pFieldName = "QCCheck13" Then
            DQCCheck13.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck13.Items.Add(ListItem1)
        End If

        'A01
        If pFieldName = "QCCheck14" Then
            DQCCheck14.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck14.Items.Add(ListItem1)
        End If

        '有害物質
        If pFieldName = "EACheck1" Then
            DEACheck1.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DEACheck1.Items.Add(ListItem1)
        End If

        '檢測結果-1
        If pFieldName = "QCResult1" Then
            DQCResult1.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCResult1.Items.Add(ListItem1)
        End If

        '檢測結果-2
        If pFieldName = "QCResult2" Then
            DQCResult2.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCResult2.Items.Add(ListItem1)
        End If

        '檢測結果-3
        If pFieldName = "QCResult3" Then
            DQCResult3.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCResult3.Items.Add(ListItem1)
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
    End Sub
End Class
