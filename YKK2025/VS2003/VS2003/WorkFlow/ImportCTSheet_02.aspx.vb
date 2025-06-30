Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ImportCTSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DOPContactSheet As System.Web.UI.WebControls.Image
    Protected WithEvents BFlow As System.Web.UI.WebControls.ImageButton
    Protected WithEvents BPrint As System.Web.UI.WebControls.ImageButton
    Protected WithEvents LNFormNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOFormNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DNFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents LAttachFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DOFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DContent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTarget As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox

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
    Dim FieldName(50) As String     '各欄位
    Dim Attribute(50) As Integer    '各欄位屬性    
    Dim Top As Integer              '動態元件的Top位置
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wStep As Integer            '工程代碼
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID
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
        Response.Cookies("PGM").Value = "ImportCTSheet_02.aspx"

        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置
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
        wStep = 999                                 '工程代碼
    End Sub

    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ImportCTFilePath")

        Dim PathOld1 As String = ConfigurationSettings.AppSettings.Get("HttpOld")
        Dim PathOld2 As String = ConfigurationSettings.AppSettings.Get("ImportCTFilePath")
        Dim RtnCode As Integer = 0

        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_ImportCTSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ImportCTSheet")
        If DBDataSet1.Tables("F_ImportCTSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("Date")                   '日期
            DDivision.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("Division") '部門
            DPerson.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("Person")   '擔當
            DSliderCode.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("SliderCode")       'Slider Code

            DOFormNo.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("OFormNo")             '圖號
            DOFormSno.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("OFormSno")             '圖號
            If DOFormNo.Text <> "" And DOFormSno.Text <> "" Then
                LOFormNo.NavigateUrl = "ImportSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
            Else
                LOFormNo.Visible = False
            End If

            DNFormNo.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("NFormNo")             '圖號
            If DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("NFormSno") = 0 Then
                DNFormSno.Text = ""
            Else
                DNFormSno.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("NFormSno")             '圖號
            End If
            If DNFormNo.Text <> "" And DNFormSno.Text <> "" Then
                LNFormNo.NavigateUrl = "ImportModSheet_02.aspx?pFormNo=" & DNFormNo.Text & "&pFormSno=" & CInt(DNFormSno.Text) & "&pOFormNo=" & DOFormNo.Text & "&pOFormSno=" & CInt(DOFormSno.Text) & "&pStep=" & wStep
            Else
                LNFormNo.Visible = False
            End If

            DTarget.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("Target")             '圖號
            DSpec.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("Spec")             '圖號

            DContent.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("Content")             '圖號
            DDReason.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("Reason")             '圖號
            If DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("AttachFile") <> "" Then          '圖檔1
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("AttachFile"))

                LAttachFile.NavigateUrl = Path & DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("AttachFile")
            Else
                LAttachFile.Visible = False
            End If
            DFormSno.Text = "單號：" & CStr(wFormSno)    '單號
        End If
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
        Top = 432
    End Sub

End Class
