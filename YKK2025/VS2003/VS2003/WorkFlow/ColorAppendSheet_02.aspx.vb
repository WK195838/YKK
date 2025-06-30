Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ColorAppendSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm
    Protected WithEvents DColorAppendSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DLevel As System.Web.UI.WebControls.TextBox
    Protected WithEvents LForCastFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPullerPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents LOFormNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DOFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DManufFlow As System.Web.UI.WebControls.TextBox
    Protected WithEvents DColorItem As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList

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
        Response.Cookies("PGM").Value = "ColorAppendSheet_02.aspx"

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ColorAppendFilePath")
        Dim PathOld1 As String = ConfigurationSettings.AppSettings.Get("HttpOld")
        Dim PathOld2 As String = ConfigurationSettings.AppSettings.Get("ColorAppendFilePath")
        Dim RtnCode As Integer = 0

        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_ColorAppendSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ColorAppendSheet")
        If DBDataSet1.Tables("F_ColorAppendSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_ColorAppendSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ColorAppendSheet").Rows(0).Item("Date")                   '日期
            SetFieldData("Division", DBDataSet1.Tables("F_ColorAppendSheet").Rows(0).Item("Division"))  '部門
            SetFieldData("Person", DBDataSet1.Tables("F_ColorAppendSheet").Rows(0).Item("Person"))      '擔當
            DSliderCode.Text = DBDataSet1.Tables("F_ColorAppendSheet").Rows(0).Item("SliderCode")       'Slider Code
            DLevel.Text = DBDataSet1.Tables("F_ColorAppendSheet").Rows(0).Item("Level")                 '難易度
            DMapNo.Text = DBDataSet1.Tables("F_ColorAppendSheet").Rows(0).Item("MapNo")                 '圖號
            DOFormNo.Text = DBDataSet1.Tables("F_ColorAppendSheet").Rows(0).Item("OFormNo")             '原表單
            DOFormSno.Text = DBDataSet1.Tables("F_ColorAppendSheet").Rows(0).Item("OFormSno")           '原單號
            If DOFormNo.Text <> "" And DOFormSno.Text <> "" Then
                If DOFormNo.Text = "000011" Then
                    LOFormNo.NavigateUrl = "ImportSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                Else
                    If DOFormNo.Text = "000003" Then
                        LOFormNo.NavigateUrl = "ManufInSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                    Else
                        LOFormNo.NavigateUrl = "ManufOutSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                    End If
                End If
                DOFormNo.Visible = False
                DOFormSno.Visible = False
            Else
                LOFormNo.Visible = False
            End If

            DColorItem.Text = DBDataSet1.Tables("F_ColorAppendSheet").Rows(0).Item("ColorItem")         '色番項目
            DManufFlow.Text = DBDataSet1.Tables("F_ColorAppendSheet").Rows(0).Item("ManufFlow")         '流程說明
            DPullerPrice.Text = DBDataSet1.Tables("F_ColorAppendSheet").Rows(0).Item("PullerPrice")     '購入價
            DPrice.Text = DBDataSet1.Tables("F_ColorAppendSheet").Rows(0).Item("Price")                 '售價
            If DBDataSet1.Tables("F_ColorAppendSheet").Rows(0).Item("ForCastFile") <> "" Then                '報價單
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_ColorAppendSheet").Rows(0).Item("ForCastFile"))

                LForCastFile.NavigateUrl = Path & DBDataSet1.Tables("F_ColorAppendSheet").Rows(0).Item("ForCastFile")
            Else
                LForCastFile.Visible = False
            End If
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
    End Sub
End Class
