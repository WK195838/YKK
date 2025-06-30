Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class MapSheetMod_02
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DMapSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DModContent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DModReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBefFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBefFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOriFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBefMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOriMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents LRefAttach As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents LMapFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMapSheet3 As System.Web.UI.WebControls.Image
    Protected WithEvents DOriFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents DModReasonCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMakeMap As System.Web.UI.WebControls.TextBox
    Protected WithEvents DLevel As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSample As System.Web.UI.WebControls.TextBox
    Protected WithEvents LOriMap As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LBefMap As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DManufType As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCpsc As System.Web.UI.WebControls.TextBox
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
        Response.Cookies("PGM").Value = "MapModSheet_02.aspx"

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("MapModFilePath")
        Dim PathOld1 As String = ConfigurationSettings.AppSettings.Get("HttpOld")
        Dim PathOld2 As String = ConfigurationSettings.AppSettings.Get("MapModFilePath")
        Dim RtnCode As Integer = 0


        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_MapModSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_MapModSheet")
        If DBDataSet1.Tables("F_MapModSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("No")         'No
            DDate.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Date")                   '日期
            DBuyer.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Buyer")                 'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("SellVendor")       '委託廠商
            DDivision.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Division")           '部門
            DSample.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Sample")           'sample
            DPerson.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Person")               '擔當
            DCpsc.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Cpsc")               'Cpsc
            DOriMapNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriMapNo")        '原始圖號
            DOriFormNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormNo")      '原始表單號碼
            DOriFormSno.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormSno")    '原始單號
            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriMapNo") <> "" And _
               DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormNo") <> "" And _
               DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormSno") > 0 Then          '原圖號
                LOriMap.NavigateUrl = "MapSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormNo") & _
                                                      "&pFormSno=" & CStr(DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormSno"))
                LOriMap.Visible = True
            Else
                LOriMap.Visible = False
            End If

            DLevel.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Level")              '難易度
            DBefMapNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefMapNo")        '前圖號
            DBefFormNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormNo")      '前表單號碼
            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormSno") > 0 Then
                DBefFormSno.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormSno")    '前單號
            Else
                DBefFormSno.Text = ""
            End If
            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefMapNo") <> "" And _
               DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormNo") <> "" And _
               DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormSno") > 0 Then            '原圖號
                LBefMap.NavigateUrl = "MapModSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormNo") & _
                                                         "&pFormSno=" & CStr(DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormSno"))
            Else
                LBefMap.Visible = False
            End If

            DModReasonCode.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("ModReasonCode")  '原因類別代碼
            DModReasonDesc.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("ModReasonDesc")  '原因說明
            DModContent.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("ModContent")        '修改細節說明

            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("RefAttach") <> "" Then              '參考附件
                LRefAttach.NavigateUrl = Path & DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("RefAttach")
            Else
                LRefAttach.Visible = False
            End If

            DMapNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("MapNo")                     '圖號
            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("MapFile") <> "" Then                   '圖檔
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("MapFile"))

                LMapFile.NavigateUrl = Path & DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("MapFile")
            Else
                LMapFile.Visible = False
            End If

            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("PdfFile") <> "" Then                   '圖檔
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("PdfFile"))

                LPdfFile.NavigateUrl = Path & DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("PdfFile")
            Else
                LPdfFile.Visible = False
            End If

            DMakeMap.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("MakeMap")                '製圖者
            DManufType.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("ManufType")            '內外注
            DSuppiler.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Suppiler")     '外注商
        End If

        DFormSno.Text = "單號：" & CStr(wFormSno)
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

End Class
