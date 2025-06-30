Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class MapSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DLevel As System.Web.UI.WebControls.TextBox
    Protected WithEvents LMapFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LRefMapFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSurface As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBackground As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCramper As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDescription As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMapReqDelDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMapSheet3 As System.Web.UI.WebControls.Image
    Protected WithEvents DMaterial As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMaterialDetail As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFrontBack As System.Web.UI.WebControls.TextBox
    Protected WithEvents DLight As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMapSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DCPSC As System.Web.UI.WebControls.TextBox
    Protected WithEvents DHalfFinish As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMakeMap As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSample As System.Web.UI.WebControls.TextBox
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DManufType As System.Web.UI.WebControls.TextBox
    Protected WithEvents Panel1 As System.Web.UI.WebControls.Panel
    Protected WithEvents DStatus As System.Web.UI.WebControls.TextBox
    Protected WithEvents LMMap As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DSuppiler As System.Web.UI.WebControls.TextBox
    Protected WithEvents LPdfFile As System.Web.UI.WebControls.HyperLink

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
        Response.Cookies("PGM").Value = "MapSheet_02.aspx"

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("MapFilePath")
        Dim PathOld1 As String = ConfigurationSettings.AppSettings.Get("HttpOld")
        Dim PathOld2 As String = ConfigurationSettings.AppSettings.Get("MapFilePath")
        Dim RtnCode As Integer = 0

        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_MapSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_MapSheet")
        If DBDataSet1.Tables("F_MapSheet").Rows.Count > 0 Then
            If DBDataSet1.Tables("F_MapSheet").Rows(0).Item("UPDSts") = 1 Then          '修改工程狀態
                DStatus.Text = "修改工程進行中"
                DStatus.Visible = True
                Panel1.Visible = True
            Else
                DStatus.Visible = False
                Panel1.Visible = False
            End If

            If DBDataSet1.Tables("F_MapSheet").Rows(0).Item("ModMap") = 1 Then          '修改工程狀態
                LMMap.NavigateUrl = "MapList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno & "&pFun=MAP"
                LMMap.Visible = True
            Else
                LMMap.Visible = False
            End If

            DNo.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("No")         'No
            DDate.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Date")                   '日期
            DBuyer.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Buyer")                 'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("SellVendor")       '委託廠商
            DDivision.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Division")           '部門
            DPerson.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Person")               '擔當
            DBackground.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Background")       '開發背景
            DSpec.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Spec")                   'Spec
            DCramper.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Cramper")             'Cramper
            DSurface.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Surface")             '表面處理
            DFrontBack.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("FrontBack")         'Puller--正反面
            DHalfFinish.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("HalfFinish")       '半成品
            DMaterial.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Material")           '材質
            DMaterialDetail.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("MaterialDetail") '材質細項
            If DBDataSet1.Tables("F_MapSheet").Rows(0).Item("RefMapFile") <> "" Then


                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_MapSheet").Rows(0).Item("RefMapFile"))

                LRefMapFile.NavigateUrl = Path & DBDataSet1.Tables("F_MapSheet").Rows(0).Item("RefMapFile")   '原圖號
            Else
                LRefMapFile.Visible = False
            End If
            DDescription.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Description")     '備註
            DManufType.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("ManufType")     '內外製
            DSuppiler.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Suppiler")     '外注商

            DCPSC.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("CPSC")                   'CPSC
            DMapReqDelDate.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("MapReqDelDate") '圖面希望交期
            DLight.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Light")                 '光造型
            DSample.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Sample")               '樣品
            DMapNo.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("MapNo")                 '圖號
            If DBDataSet1.Tables("F_MapSheet").Rows(0).Item("MapFile") <> "" Then

                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_MapSheet").Rows(0).Item("MapFile"))
                LMapFile.NavigateUrl = Path & DBDataSet1.Tables("F_MapSheet").Rows(0).Item("MapFile")  '圖檔
            Else
                LMapFile.Visible = False
            End If

            If DBDataSet1.Tables("F_MapSheet").Rows(0).Item("PdfFile") <> "" Then

                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_MapSheet").Rows(0).Item("pdfFile"))
                LPdfFile.NavigateUrl = Path & DBDataSet1.Tables("F_MapSheet").Rows(0).Item("PdfFile")  '圖檔
            Else
                LPdfFile.Visible = False
            End If


            DLevel.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Level")     '難易度
            DMakeMap.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("MakeMap") '製圖者
        End If

        DFormSno.Text = "單號：" & CStr(wFormSno)
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

End Class
